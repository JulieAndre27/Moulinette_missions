"""compute CO2e emissions from LMD travel data"""

from __future__ import annotations

import configparser
import logging
import os
from dataclasses import dataclass
from typing import Any

import geopy
import orjson.orjson
import pandas as pd
from geopy import geocoders
from geopy.distance import geodesic

# Setup logging
logging.basicConfig(level=logging.ERROR)
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)


class TravelData:
    """load data from excel/ods according to the conf"""

    def __init__(self, data_path: str, conf_path: str):
        # load conf
        config = configparser.RawConfigParser()
        config.read(conf_path, encoding='utf-8')
        config_dict = {s: dict(config.items(s)) for s in config.sections()}
        assert len(config_dict.keys()) == 1, "Exactly one sheet must be specified in the configuration"
        self.sheet_name = list(config_dict.keys())[0]
        config_dict = config_dict[self.sheet_name]

        # column letter to index
        column_names = ["departure_city", "departure_country", "arrival_city", "arrival_country", "t_type", "round_trip"]
        cols = [ord(config_dict[i].lower()) - 97 for i in column_names]
        # Transportation descriptors
        self.t_type_air, self.t_type_train, self.t_type_car = (config_dict[i].split(', ') for i in ["t_type_air", "t_type_train", "t_type_car"])

        # load sheet
        # noinspection PyTypeChecker
        self.data = pd.read_excel(data_path, sheet_name=self.sheet_name, usecols=cols, names=column_names)

        # determine computed transport type
        self.data["transport_for_emissions"] = self.data.apply(self.get_transport_for_calculator, axis=1)

    def get_transport_for_calculator(self, row: pd.Series) -> str | None:
        """get transport used for computation"""
        if isinstance(row.t_type, str):
            t_type = [i.strip() for i in row.t_type.split(",")]  # used transport modes
            for t_id, t in enumerate([self.t_type_air, self.t_type_train, self.t_type_car]):  # plane > train > car
                for e in t:  # for each known transportation method
                    if e in t_type:  # check if used and return
                        return t_id

        return None


@dataclass
class CustomLocation:
    """Store a location, serializable"""
    address: str
    latitude: float
    longitude: float
    countryCode: str | None

    @classmethod
    def from_location(cls, loc: geopy.location.Location) -> CustomLocation:
        """init from geopy location"""
        return cls(loc.address, loc.latitude, loc.longitude, loc.raw.get("countryCode", None))

    @classmethod
    def from_json(cls, data: dict) -> CustomLocation:
        """init from json dict"""
        return cls(data['address'], data["latitude"], data['longitude'], data["countryCode"])


class GeoDistanceCalculator:
    """compute distance between locations, cached"""

    def __init__(self):
        self.gn = geocoders.GeoNames(username='oaumont')
        self.N_api_calls = 0
        self.N_cache_hits = 0
        self.cache_path = "Data/geocache.json"

        # load cache
        if not os.path.isfile(self.cache_path):  # make sure file exists
            with open(self.cache_path, "w", encoding="utf8") as f:
                f.write("{}")

        with open(self.cache_path, "r", encoding="utf8") as f:
            cache_json = orjson.loads(f.read())
        self.cache = {k: CustomLocation.from_json(v) for k, v in cache_json.items()}

    def print_usage(self) -> None:
        """output usage"""
        logger.info(f"API calls: {self.N_api_calls}, cache hits: {self.N_cache_hits}")

    def save_cache(self) -> None:
        """save cache to disk"""
        with open(self.cache_path, "wb") as f:
            f.write(orjson.dumps(self.cache))

    def find_location(self, loc_str: str) -> CustomLocation:
        """find location of place"""
        if loc_str in self.cache:
            self.N_cache_hits += 1
            return self.cache[loc_str]

        self.N_api_calls += 1
        coord = self.gn.geocode(loc_str, timeout=30)
        if coord is None:
            raise ValueError(f"Error: could not locate <{loc_str}>")

        coord = CustomLocation.from_location(coord)

        self.cache[loc_str] = coord
        self.save_cache()

        return coord

    @staticmethod
    def geodesic_distance(loc1: CustomLocation, loc2: CustomLocation) -> geopy.distance.geodesic:
        """geodesic distance between two locations"""
        return geodesic(
            (loc1.latitude, loc1.longitude), (loc2.latitude, loc2.longitude),
        )

    def get_geodesic_distance_between(self, loc1: str, loc2: str) -> tuple[float | None, CustomLocation, CustomLocation]:
        """get the geo distance (km) between two places from their address

        :returns (distance, geocode loc1, geocode loc2)
        """
        coord1 = self.find_location(loc1)
        coord2 = self.find_location(loc2)

        return self.geodesic_distance(coord1, coord2).km, coord1, coord2


class EmissionCalculator:
    """compute emissions for a trip"""

    def __init__(self):
        self.dist_calculator = GeoDistanceCalculator()

        ## Parameters
        # Determine flight type from distance (short-med-long)
        self.d_short_distance_plane = 1000
        self.d_medium_distance_plane = 3500
        # Emissions factors for plane (kgCO2e / km)
        self.plane_geodesic_correction_offset_km = 95  # an offset to add to the geodesic distance for plane
        self.EF_short_distance_plane = 0.258
        self.EF_medium_distance_plane = 0.187
        self.EF_long_distance_plane = 0.152

        # Emissions factors for train (kgCO2e / km)
        self.train_geodesic_correction = 1.2  # approx. : multiply the geodesic distance to take into account the path is not straight.
        self.d_TER_TGV_km = 200  # threshold in km to chose whether to take TER or TGV emissions
        self.EF_french_TER = 0.018
        self.EF_french_TGV = 0.003
        self.EF_half_french_trains = 0.016  # if the departure is in France but not the destination (vice-versa)
        self.EF_european_trains = 0.037  # for trains between two cities outside France

        # Emission factor for cars (general) (kgCO2e / km)
        self.EF_cars = 0.233
        self.car_geodesic_correction = 1.3  # approx. : multiply the geodesic distance to take into account the path is not straight.

        # Threshold (km) used only if a row is missing the type of transport
        self.threshold_unknown_transportation = 700

    def get_co2e_from_distance(self, dist_km: float, transport_type: int, geo_departure: CustomLocation, geo_arrival: CustomLocation) -> float:
        """get co2e (kg) from data"""
        if transport_type == 0:  # plane
            corrected_distance = dist_km + self.plane_geodesic_correction_offset_km
            if dist_km < self.d_short_distance_plane:
                EF = self.EF_short_distance_plane
            elif dist_km < self.d_medium_distance_plane:
                EF = self.EF_medium_distance_plane
            else:
                EF = self.EF_long_distance_plane

        elif transport_type == 1:  # train
            corrected_distance = self.train_geodesic_correction * dist_km
            # For trains, we distinguish if both cities are in France or not, to use SNCF's emission factors or European ones
            if (geo_departure.countryCode == 'FR') and (geo_arrival.countryCode == 'FR'):  # from AND to france
                EF = self.EF_french_TER if self.train_geodesic_correction * dist_km < self.d_TER_TGV_km else self.EF_french_TGV
            elif (geo_departure.countryCode == 'FR') or (geo_arrival.countryCode == 'FR'):  # from OR to france
                EF = self.EF_half_french_trains
            else:
                EF = self.EF_european_trains

        elif transport_type == 2:  # car
            corrected_distance = self.car_geodesic_correction * dist_km
            EF = self.EF_cars
        else:
            raise ValueError

        return corrected_distance * EF

    def compute_one(self, row: pd.Series) -> pd.Series | None:
        """compute emissions from a single trip"""
        logger.debug(f"computing one trip {row.departure_city} -> {row.arrival_city}")

        departure_str = f"{row.departure_city} ({row.departure_country})"
        arrival_str = f"{row.arrival_city} ({row.arrival_country})"
        try:
            dist_km, geo_departure, geo_arrival = self.dist_calculator.get_geodesic_distance_between(departure_str, arrival_str)
        except ValueError as e:
            logger.error(e)
            return None

        transport_type = row.transport_for_emissions
        # Handle unknown transportation
        try:
            transport_type = int(transport_type)
        except ValueError:
            transport_type = 1 if dist_km < self.threshold_unknown_transportation else 0

        co2e_emissions = self.get_co2e_from_distance(dist_km, transport_type, geo_departure, geo_arrival)

        t_type_str = ["plane", "train", "car"]
        logger.debug(f"One-way emissions from {departure_str} to {arrival_str} ({dist_km:.0f} km) by {t_type_str[transport_type]} is {co2e_emissions:.1f} kg C02e")

        # Round trip
        final_distance_km = dist_km
        final_co2e_emissions = co2e_emissions
        if row.round_trip.lower() in ['oui', "yes", 'o', 'y']:
            final_distance_km *= 2
            final_co2e_emissions *= 2

        return pd.Series(data=[dist_km, final_distance_km, final_co2e_emissions, geo_departure.countryCode, geo_arrival.countryCode, t_type_str[transport_type]], index=["one_way_dist_km", "dist_km", "co2e_emissions_kg", "departure_countrycode", "arrival_countrycode", "transport_for_emissions_str"])

    def compute(self, tv_data: TravelData, out: str):
        """compute emissions for a TravelData and outputs the result as a spreadsheet <out>"""
        df_result = pd.DataFrame(tv_data.data)

        # Compute emissions
        res = df_result.apply(self.compute_one, axis=1)

        self.dist_calculator.print_usage()

        res = pd.concat((df_result, res), axis=1)

        # Format results
        res = res.drop('transport_for_emissions', axis=1)
        res = res.round({"one_way_dist_km": 0, "dist_km": 0, "co2e_emissions_kg": 0})
        res = res.rename(columns={"one_way_dist_km": "Distance (one-way, km)", "dist_km": "Distance (km)", "co2e_emissions_kg": "CO2e emissions (kg)", "departure_countrycode": "CP départ", "arrival_countrycode": "CP arrivée", "transport_for_emissions_str": "Transport utilisé pour calcul"})

        # Save to file
        res.to_excel(out, sheet_name=tv_data.sheet_name, float_format="%.0f", freeze_panes=(0, 1), index=False)


td = TravelData("Data/MIS_2022_extrait.xlsx", "Data/config_file_LMD_CNRS_2022.cfg")
EmissionCalculator().compute(td, "Data/out.ods")
