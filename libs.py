"""Libraries to compute CO2e emissions from LMD travel data"""

from __future__ import annotations

import configparser
import logging
import os
from dataclasses import dataclass
from pathlib import Path

import geopy
import orjson.orjson
import pandas as pd
from geopy import geocoders
from geopy.distance import geodesic

# Setup logging to avoid printing all strings
logging.basicConfig(level=logging.ERROR)
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)


def split_strip_lower(s: str, splitter: str = ",") -> list[str]:
    """lower then split <s> with <splitter>, and strip each member"""
    return [i.strip() for i in s.lower().split(splitter)]


class Enm:
    """Easy Enums"""

    # Values of roundtrip in spreadsheet
    ROUNDTRIP_CORRECTED = "non [corr.]"
    ROUNDTRIP_YES = "oui"
    ROUNDTRIP_NO = "non"

    # Column names from config
    COL_MISSION_ID = "mission_id"
    COL_DEPARTURE_CITY = "departure_city"
    COL_DEPARTURE_COUNTRY = "departure_country"
    COL_ARRIVAL_CITY = "arrival_city"
    COL_ARRIVAL_COUNTRY = "arrival_country"
    COL_TRANSPORT_TYPE = "t_type"
    COL_ROUND_TRIP = "round_trip"
    # Custom column names
    COL_MAIN_TRANSPORT = "main_transport"  # the main transportation method used for computing emissions

    # Main transport types
    MAIN_TRANSPORT_PLANE = "plane"
    MAIN_TRANSPORT_TRAIN = "train"
    MAIN_TRANSPORT_CAR = "car"


class MissionsData:
    """load data from excel/ods according to the conf"""

    def __init__(self, data_path: str | Path, conf_path: str | Path):
        # load conf
        config = configparser.RawConfigParser()
        config.read(conf_path, encoding="utf-8")
        config_dict = {s: dict(config.items(s)) for s in config.sections()}
        assert len(config_dict.keys()) == 1, "Exactly one sheet must be specified in the configuration"
        self.sheet_name = list(config_dict.keys())[0]
        config_dict = config_dict[self.sheet_name]

        # column letter to index
        column_names = [
            Enm.COL_MISSION_ID,
            Enm.COL_DEPARTURE_CITY,
            Enm.COL_DEPARTURE_COUNTRY,
            Enm.COL_ARRIVAL_CITY,
            Enm.COL_ARRIVAL_COUNTRY,
            Enm.COL_TRANSPORT_TYPE,
            Enm.COL_ROUND_TRIP,
        ]
        cols = [ord(config_dict[v].lower()) - 97 for v in column_names]
        # Transportation descriptors
        self.ttypes_air, self.ttypes_train, self.ttypes_car, self.ttypes_ignored = (
            split_strip_lower(config_dict[i]) for i in ["t_type_air", "t_type_train", "t_type_car", "t_type_ignored"]
        )

        # load sheet
        # noinspection PyTypeChecker
        self.data = pd.read_excel(data_path, sheet_name=self.sheet_name, usecols=cols, names=column_names)

        # determine computed transport type
        self.unknown_transport_types = set()
        self.data[Enm.COL_MAIN_TRANSPORT] = self.data.apply(self.get_main_transport, axis=1)
        if len(self.unknown_transport_types):
            logger.warning(f"Unknown transport types: {self.unknown_transport_types}")

        # sanitize data
        self.fix_round_trips()

    def get_main_transport(self, row: pd.Series) -> str | None:
        """get main transport used for computation of emissions"""
        if isinstance(row.t_type, str):
            row_transports = split_strip_lower(row.t_type)  # row's used transport modes
            # Check that we recognize all existing types
            for transport in row_transports:
                if (
                    transport not in self.ttypes_air
                    and transport not in self.ttypes_train
                    and transport not in self.ttypes_car
                    and transport not in self.ttypes_ignored
                ):
                    self.unknown_transport_types.add(transport)

            transport_priority_order = [
                (self.ttypes_air, Enm.MAIN_TRANSPORT_PLANE),
                (self.ttypes_train, Enm.MAIN_TRANSPORT_TRAIN),
                (self.ttypes_car, Enm.MAIN_TRANSPORT_CAR),
            ]
            for ttypes, main_transport in transport_priority_order:  # plane > train > car
                for ttype in ttypes:  # for each known transportation method
                    if ttype in row_transports:  # check if used and return
                        return main_transport

        return None

    def fix_round_trips(self):
        """Fix incorrect round trips
        If a mission ID has several trips, set them all to one-way"""
        self.data[Enm.COL_ROUND_TRIP] = self.data[Enm.COL_ROUND_TRIP].apply(str.lower)
        duplicated_mission_loc = self.data[Enm.COL_MISSION_ID].duplicated(keep=False)
        self.data.loc[duplicated_mission_loc & (self.data[Enm.COL_ROUND_TRIP] == Enm.ROUNDTRIP_YES), Enm.COL_ROUND_TRIP] = Enm.ROUNDTRIP_CORRECTED


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
        return cls(loc.address, loc.latitude, loc.longitude, loc.raw.get("components", {}).get("ISO_3166-1_alpha-2", None))

    @classmethod
    def from_json(cls, data: dict) -> CustomLocation:
        """init from json dict"""
        return cls(data["address"], data["latitude"], data["longitude"], data["countryCode"])


class GeoDistanceCalculator:
    """compute distance between locations, cached"""

    def __init__(self):
        self.locator = geocoders.OpenCage(api_key="f92cf03ca7f24c999891faefd3fc4121", timeout=30)
        self.N_api_calls = 0
        self.N_cache_hits = 0
        self.cache_path = "Data/Config/geocache.json"
        self.save_cache_every_N = 100

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

    def save_cache(self, force=False) -> None:
        """save cache to disk"""
        if force or len(self.cache) % self.save_cache_every_N == 0:
            with open(self.cache_path, "wb") as f:
                f.write(orjson.dumps(self.cache))

    def find_location(self, loc_str: str) -> CustomLocation:
        """find location of place"""
        if loc_str in self.cache:
            self.N_cache_hits += 1
            return self.cache[loc_str]

        self.N_api_calls += 1
        coord = self.locator.geocode(loc_str, language="fr")
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
            (loc1.latitude, loc1.longitude),
            (loc2.latitude, loc2.longitude),
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

    def __init__(self, data_path: str | Path, conf_path: str | Path, out: str | Path):
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
        self.EF_plane_uncertainty = 0.7

        # Emissions factors for train (kgCO2e / km)
        self.train_geodesic_correction = 1.2  # approx. : multiply the geodesic distance to take into account the path is not straight.
        self.d_TER_TGV_km = 200  # threshold in km to chose whether to take TER or TGV emissions
        self.EF_french_TER = 0.018
        self.EF_french_TGV = 0.003
        self.EF_half_french_trains = 0.016  # if the departure is in France but not the destination (vice-versa)
        self.EF_european_trains = 0.037  # for trains between two cities outside France
        self.EF_train_uncertainty = 0.2
        self.EF_train_TER_uncertainty = 0.6

        # Emission factor for cars (general) (kgCO2e / km)
        self.EF_car = 0.233
        self.EF_car_uncertainty = 0.6
        self.car_geodesic_correction = 1.3  # approx. : multiply the geodesic distance to take into account the path is not straight.

        # Threshold (km) used only if a row is missing the type of transport
        self.threshold_unknown_transportation = 700
        self.threshold_force_plane = 4000  # force plane if distance is too big to avoid input errors

        # Apply the compute function, which computes the emissions ect.
        self.compute(MissionsData(data_path, conf_path), out)

    def get_co2e_from_distance(
        self, dist_km: float, transport_type: str, geo_departure: CustomLocation, geo_arrival: CustomLocation, is_round_trip: bool
    ) -> tuple[float, float, str]:
        """get co2e (kg) from data
        :returns (CO2e in kg, uncertainty in kg, transport type used)"""
        if transport_type == Enm.MAIN_TRANSPORT_PLANE:  # plane
            corrected_distance = dist_km + self.plane_geodesic_correction_offset_km
            EF_uncertainty = self.EF_plane_uncertainty
            # determine the associated emission factor
            if dist_km < self.d_short_distance_plane:
                EF = self.EF_short_distance_plane
                used_ttype = "plane (short)"
            elif dist_km < self.d_medium_distance_plane:
                EF = self.EF_medium_distance_plane
                used_ttype = "plane (med)"
            else:
                EF = self.EF_long_distance_plane
                used_ttype = "plane (long)"

        elif transport_type == Enm.MAIN_TRANSPORT_TRAIN:  # train
            corrected_distance = self.train_geodesic_correction * dist_km
            EF_uncertainty = self.EF_train_uncertainty
            # determine the associated emission factor. For trains, we distinguish if both cities are in France or not, to use SNCF's emission factors or European ones
            if (geo_departure.countryCode == "FR") and (geo_arrival.countryCode == "FR"):  # from AND to france
                if corrected_distance < self.d_TER_TGV_km:  # it is a TER not a TGV
                    EF = self.EF_french_TER
                    used_ttype = "train (TER, FR)"
                    EF_uncertainty = self.EF_train_TER_uncertainty
                else:
                    EF = self.EF_french_TGV
                    used_ttype = "train (FR)"

            elif (geo_departure.countryCode == "FR") or (geo_arrival.countryCode == "FR"):  # from OR to france
                EF = self.EF_half_french_trains
                used_ttype = "train (half-FR)"
            else:
                EF = self.EF_european_trains
                used_ttype = "train (EU)"

        elif transport_type == Enm.MAIN_TRANSPORT_CAR:  # car
            corrected_distance = self.car_geodesic_correction * dist_km
            EF_uncertainty = self.EF_car_uncertainty
            # determine the associated emission factor
            EF = self.EF_car
            used_ttype = Enm.MAIN_TRANSPORT_CAR

        else:
            raise ValueError

        if is_round_trip:
            corrected_distance *= 2

        emissions = corrected_distance * EF  # in kg
        uncertainty = emissions * EF_uncertainty  # in kg

        return emissions, uncertainty, used_ttype

    def compute_emissions_one_row(self, row: pd.Series) -> pd.Series | None:
        """compute emissions from a single trip"""
        logger.debug(f"computing one trip {row.departure_city} -> {row.arrival_city}")

        departure_str = f"{row.departure_city} ({row.departure_country})" if isinstance(row.departure_country, str) else row.departure_city
        arrival_str = f"{row.arrival_city} ({row.arrival_country})" if isinstance(row.arrival_country, str) else row.arrival_city

        try:
            dist_km, geo_departure, geo_arrival = self.dist_calculator.get_geodesic_distance_between(departure_str, arrival_str)
        except ValueError as e:
            logger.error(e)
            return None

        transport_type = row[Enm.COL_MAIN_TRANSPORT]
        # Handle unknown transportation
        if not isinstance(transport_type, str):
            transport_type = Enm.MAIN_TRANSPORT_TRAIN if dist_km < self.threshold_unknown_transportation else Enm.MAIN_TRANSPORT_PLANE
        # Handle input errors
        if dist_km > self.threshold_force_plane:
            transport_type = Enm.MAIN_TRANSPORT_PLANE

        is_round_trip = row.round_trip == Enm.ROUNDTRIP_YES

        co2e_emissions, uncertainty, used_transport_type = self.get_co2e_from_distance(dist_km, transport_type, geo_departure, geo_arrival, is_round_trip)

        final_distance_km = dist_km * 2 if is_round_trip else dist_km

        logger.debug(f"One-way emissions from {departure_str} to {arrival_str} ({dist_km:.0f} km) by {used_transport_type} is {co2e_emissions:.1f} kg C02e")

        return pd.Series(
            data=[dist_km, final_distance_km, co2e_emissions, geo_departure.countryCode, geo_arrival.countryCode, used_transport_type, uncertainty],
            index=[
                "one_way_dist_km",
                "final_dist_km",
                "co2e_emissions_kg",
                "departure_countrycode",
                "arrival_countrycode",
                "transport_for_emissions_detailed",
                "uncertainty",
            ],
        )

    def compute(self, tv_data: MissionsData, output_path: str | Path):
        """compute emissions for a TravelData and outputs the result as a spreadsheet <output_path>"""
        # Create a new dataframe (output data) by first copying the input dataframe
        df_data = tv_data.data

        # Compute emissions row by row
        df_emissions = df_data.apply(self.compute_emissions_one_row, axis=1)

        self.dist_calculator.print_usage()  # Print the API and cache usage
        self.dist_calculator.save_cache(force=True)  # Synchronise the cache to the disk

        df_output = pd.concat((df_data, df_emissions), axis=1)

        # Format the results
        df_output = df_output.drop(Enm.COL_MAIN_TRANSPORT, axis=1)  # the information of the main transport type is already in transport_for_emissions_detailed
        df_output = df_output.round({"one_way_dist_km": 0, "final_dist_km": 0, "co2e_emissions_kg": 1})
        df_output = df_output.rename(
            columns={
                Enm.COL_DEPARTURE_CITY: "Départ (ville)",
                Enm.COL_DEPARTURE_COUNTRY: "Départ (pays)",
                Enm.COL_ARRIVAL_CITY: "Arrivée (ville)",
                Enm.COL_ARRIVAL_COUNTRY: "Arrivée (pays)",
                Enm.COL_TRANSPORT_TYPE: "Transport",
                Enm.COL_ROUND_TRIP: "A/R",
                "one_way_dist_km": "Distance (one-way, km)",
                "final_dist_km": "Distance totale (km)",
                "co2e_emissions_kg": "CO2e emissions (kg)",
                "departure_countrycode": "Code pays départ",
                "arrival_countrycode": "Code pays arrivée",
                "transport_for_emissions_detailed": "Transport utilisé pour calcul",
                "uncertainty": "Incertitude (kg)",
            }
        )

        # Save to file
        df_output.to_excel(str(output_path), sheet_name=tv_data.sheet_name, float_format="%.1f", freeze_panes=(0, 1), index=False)
