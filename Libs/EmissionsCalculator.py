"""Compute emissions"""

from __future__ import annotations

import logging
from pathlib import Path

import pandas as pd

from Libs.EasyEnums import Enm
from Libs.Geo import CustomLocation, GeoDistanceCalculator

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)


dist_calculator = GeoDistanceCalculator()


class E_params:
    """emission parameters"""

    # Emission parameters
    # Determine flight type from distance (short-med-long)
    d_short_distance_plane = 1000
    d_medium_distance_plane = 3500
    # Emissions factors for plane (kgCO2e / km)
    plane_geodesic_correction_offset_km = 95  # an offset to add to the geodesic distance for plane
    EF_short_distance_plane = 0.258
    EF_medium_distance_plane = 0.187
    EF_long_distance_plane = 0.152
    EF_plane_uncertainty = 0.7

    # Emissions factors for train (kgCO2e / km)
    train_geodesic_correction = 1.2  # approx. : multiply the geodesic distance to take into account the path is not straight.
    d_TER_TGV_km = 200  # threshold in km to chose whether to take TER or TGV emissions
    EF_french_TER = 0.018
    EF_french_TGV = 0.003
    EF_half_french_trains = 0.016  # if the departure is in France but not the destination (vice-versa)
    EF_european_trains = 0.037  # for trains between two cities outside France
    EF_train_uncertainty = 0.2
    EF_train_TER_uncertainty = 0.6

    # Emission factor for cars (general) (kgCO2e / km)
    EF_car = 0.233
    EF_car_uncertainty = 0.6
    car_geodesic_correction = 1.3  # approx. : multiply the geodesic distance to take into account the path is not straight.

    # Threshold (km) used only if a row is missing the type of transport
    threshold_unknown_transportation = 700
    threshold_force_plane = 4000  # force plane if distance is too big to avoid input errors


def _get_co2e_from_distance(
    dist_km: float, main_ttype: str, loc_departure: CustomLocation, loc_arrival: CustomLocation, is_round_trip: bool
) -> tuple[float, float, str]:
    """get co2e (kg) from data
    :returns (CO2e in kg, uncertainty in kg, transport type used)"""
    if main_ttype == Enm.MAIN_TRANSPORT_PLANE:  # plane
        corrected_distance = dist_km + E_params.plane_geodesic_correction_offset_km
        EF_uncertainty = E_params.EF_plane_uncertainty
        # determine the associated emission factor
        if dist_km < E_params.d_short_distance_plane:
            EF = E_params.EF_short_distance_plane
            emission_ttype = Enm.EF_TRANSPORT_PLANE_SHORT
        elif dist_km < E_params.d_medium_distance_plane:
            EF = E_params.EF_medium_distance_plane
            emission_ttype = Enm.EF_TRANSPORT_PLANE_MED
        else:
            EF = E_params.EF_long_distance_plane
            emission_ttype = Enm.EF_TRANSPORT_PLANE_LONG

    elif main_ttype == Enm.MAIN_TRANSPORT_TRAIN:  # train
        corrected_distance = E_params.train_geodesic_correction * dist_km
        EF_uncertainty = E_params.EF_train_uncertainty
        # determine the associated emission factor. For trains, we distinguish if both cities are in France or not, to use SNCF's emission factors or European ones
        if (loc_departure.countryCode == "FR") and (loc_arrival.countryCode == "FR"):  # from AND to france
            if corrected_distance < E_params.d_TER_TGV_km:  # it is a TER not a TGV
                EF = E_params.EF_french_TER
                emission_ttype = Enm.EF_TRANSPORT_TRAIN_TER
                EF_uncertainty = E_params.EF_train_TER_uncertainty
            else:
                EF = E_params.EF_french_TGV
                emission_ttype = Enm.EF_TRANSPORT_TRAIN_FR

        elif (loc_departure.countryCode == "FR") or (loc_arrival.countryCode == "FR"):  # from OR to france
            EF = E_params.EF_half_french_trains
            emission_ttype = Enm.EF_TRANSPORT_TRAIN_HALF_FR
        else:
            EF = E_params.EF_european_trains
            emission_ttype = Enm.EF_TRANSPORT_TRAIN_EU

    elif main_ttype == Enm.MAIN_TRANSPORT_CAR:  # car
        corrected_distance = E_params.car_geodesic_correction * dist_km
        EF_uncertainty = E_params.EF_car_uncertainty
        # determine the associated emission factor
        EF = E_params.EF_car
        emission_ttype = Enm.EF_TRANSPORT_CAR

    else:
        raise ValueError

    if is_round_trip:
        corrected_distance *= 2

    emissions = corrected_distance * EF  # in kg
    uncertainty = emissions * EF_uncertainty  # in kg

    return emissions, uncertainty, emission_ttype


def compute_emissions_one_row(row: pd.Series) -> pd.Series | None:
    """compute emissions from a single trip (as pd.df row)"""
    logger.debug(f"computing one trip {row.departure_city} -> {row.arrival_city}")

    departure_address = f"{row.departure_city} ({row.departure_country})" if isinstance(row.departure_country, str) else row.departure_city
    arrival_address = f"{row.arrival_city} ({row.arrival_country})" if isinstance(row.arrival_country, str) else row.arrival_city
    is_round_trip = row.round_trip == Enm.ROUNDTRIP_YES
    main_transport = row[Enm.COL_MAIN_TRANSPORT]

    try:
        one_way_dist_km, loc_departure, loc_arrival = dist_calculator.get_geodesic_distance_between(departure_address, arrival_address)
    except ValueError as e:
        logger.error(e)
        return None

    # Handle unknown transportation
    if not isinstance(main_transport, str):
        main_transport = Enm.MAIN_TRANSPORT_TRAIN if one_way_dist_km < E_params.threshold_unknown_transportation else Enm.MAIN_TRANSPORT_PLANE
    # Handle input errors
    if one_way_dist_km > E_params.threshold_force_plane:
        main_transport = Enm.MAIN_TRANSPORT_PLANE

    co2e_emissions, uncertainty, emission_ttype = _get_co2e_from_distance(one_way_dist_km, main_transport, loc_departure, loc_arrival, is_round_trip)
    final_distance_km = one_way_dist_km * 2 if is_round_trip else one_way_dist_km
    logger.debug(
        f"{'Round trip' if is_round_trip else 'One-way'} emissions from {departure_address} to {arrival_address} ({final_distance_km:.0f} km) by {emission_ttype} is {co2e_emissions:.1f} kg C02e"
    )

    return pd.Series(
        data=[one_way_dist_km, final_distance_km, co2e_emissions, loc_departure.countryCode, loc_arrival.countryCode, emission_ttype, uncertainty],
        index=[
            Enm.COL_DIST_ONE_WAY,
            Enm.COL_DIST_TOTAL,
            Enm.COL_EMISSIONS,
            Enm.COL_DEPARTURE_COUNTRYCODE,
            Enm.COL_ARRIVAL_COUNTRYCODE,
            Enm.COL_EMISSION_TRANSPORT,
            Enm.COL_EMISSION_UNCERTAINTY,
        ],
    )


def compute_emissions_df(df_data: pd.DataFrame) -> pd.DataFrame:
    """compute emissions for a pd.df from MissionLoader
    :returns an extended version of <df_data> with the new computed columns"""
    # Compute emissions row by row
    df_emissions = df_data.apply(compute_emissions_one_row, axis=1)

    dist_calculator.print_usage()  # Print the API and cache usage
    dist_calculator.save_cache(force=True)  # Synchronise the cache to the disk

    return pd.concat((df_data, df_emissions), axis=1)


def format_emissions_df(df_emissions: pd.DataFrame) -> pd.DataFrame:
    """format the output of <compute_emissions_df> to expected lab format"""
    # Format the results
    df_output = df_emissions.drop(Enm.COL_MAIN_TRANSPORT, axis=1)  # the information of the main transport type is already in transport_for_emissions_detailed
    df_output = df_output.round({Enm.COL_DIST_ONE_WAY: 0, Enm.COL_DIST_TOTAL: 0, Enm.COL_EMISSIONS: 1})
    df_output = df_output.rename(
        columns={
            Enm.COL_MISSION_ID: "N° mission",
            Enm.COL_DEPARTURE_CITY: "Départ (ville)",
            Enm.COL_DEPARTURE_COUNTRY: "Départ (pays)",
            Enm.COL_ARRIVAL_CITY: "Arrivée (ville)",
            Enm.COL_ARRIVAL_COUNTRY: "Arrivée (pays)",
            Enm.COL_TRANSPORT_TYPE: "Transport",
            Enm.COL_ROUND_TRIP: "A/R",
            Enm.COL_DIST_ONE_WAY: "Distance (one-way, km)",
            Enm.COL_DIST_TOTAL: "Distance totale (km)",
            Enm.COL_EMISSIONS: "CO2e emissions (kg)",
            Enm.COL_DEPARTURE_COUNTRYCODE: "Code pays départ",
            Enm.COL_ARRIVAL_COUNTRYCODE: "Code pays arrivée",
            Enm.COL_EMISSION_TRANSPORT: "Transport utilisé pour calcul",
            Enm.COL_EMISSION_UNCERTAINTY: "Incertitude (kg)",
        }
    )

    return df_output


def save_to_file(df: pd.DataFrame, output_path: str | Path) -> None:
    """Save dataframe to file"""
    assert str(output_path).endswith(".xlsx"), "Output must be in .xlsx format"

    sheet_name = df.attrs["sheet_name"]
    with pd.ExcelWriter(str(output_path), engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name=sheet_name, float_format="%.1f", freeze_panes=(0, 1), index=False)

        # Auto-resize each column
        worksheet = writer.sheets[sheet_name]  # pull worksheet object
        for idx, col in enumerate(df):  # loop through all columns
            series = df[col]
            max_len = (
                max((series.astype(str).map(len).max(), len(str(series.name)))) + 1  # len of largest item  # len of column name/header
            )  # adding a little extra space
            worksheet.set_column(idx, idx, max_len)  # set column width
