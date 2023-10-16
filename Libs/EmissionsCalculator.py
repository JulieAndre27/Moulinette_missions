"""Compute emissions"""

from __future__ import annotations

import logging
from pathlib import Path

import pandas as pd

from Libs.EasyEnums import Enm
from Libs.Excel import auto_resize_columns, format_headers
from Libs.Geo import CustomLocation, GeoDistanceCalculator
from Libs.misc import index_to_column

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

    df_output = pd.concat((df_data, df_emissions), axis=1)
    df_output.insert(1, Enm.COL_CREDITS, df_output.attrs["credits"])

    return df_output


def save_to_file(df_raw: pd.DataFrame, output_path: Path) -> None:
    """Save dataframe to file"""
    assert output_path.suffix == ".xlsx", "Output must be in .xlsx format"
    df = df_raw.copy()

    with pd.ExcelWriter(str(output_path), engine="xlsxwriter") as writer:
        sheet_name = "emissions"
        sheet_name_raw = f"{sheet_name}_raw"
        sheet_name_formatted = f"{sheet_name}_formatted"

        workbook = writer.book
        sheet_formatted = workbook.add_worksheet(sheet_name_formatted)

        # save raw data
        df_raw.to_excel(writer, sheet_name=sheet_name_raw, float_format="%.1f", freeze_panes=(1, 1), index=False)
        sheet_raw = writer.sheets[sheet_name_raw]

        # Create a dynamic link between raw and formatted
        columns_formatted = [
            Enm.COL_CREDITS,
            Enm.COL_MISSION_ID,
            Enm.COL_DEPARTURE_DATE,
            Enm.COL_DEPARTURE_CITY,
            Enm.COL_DEPARTURE_COUNTRYCODE,
            Enm.COL_ARRIVAL_CITY,
            Enm.COL_ARRIVAL_COUNTRYCODE,
            Enm.COL_ROUND_TRIP,
            Enm.COL_TRANSPORT_TYPE,
            Enm.COL_EMISSION_TRANSPORT,
            Enm.COL_DIST_TOTAL,
            Enm.COL_EMISSIONS,
            Enm.COL_EMISSION_UNCERTAINTY,
        ]
        columns_to_fr = {
            Enm.COL_CREDITS: "Crédits",
            Enm.COL_MISSION_ID: f"N°\nmission",
            Enm.COL_DEPARTURE_DATE: f"Départ\n(date)",
            Enm.COL_DEPARTURE_CITY: f"Départ\n(ville)",
            Enm.COL_DEPARTURE_COUNTRYCODE: f"Départ\n(pays)",
            Enm.COL_ARRIVAL_CITY: f"Arrivée\n(ville)",
            Enm.COL_ARRIVAL_COUNTRYCODE: f"Arrivée\n(pays)",
            Enm.COL_TRANSPORT_TYPE: "Transport",
            Enm.COL_ROUND_TRIP: "A/R",
            Enm.COL_DIST_ONE_WAY: f"Distance\n(one-way, km)",
            Enm.COL_DIST_TOTAL: f"Distance\ntotale (km)",
            Enm.COL_EMISSIONS: f"CO2e emissions\n(kg)",
            Enm.COL_EMISSION_TRANSPORT: f"Transport utilisé\npour calcul",
            Enm.COL_EMISSION_UNCERTAINTY: f"Incertitude\n(±kg)",
        }
        df = df.round({Enm.COL_DIST_ONE_WAY: 0, Enm.COL_DIST_TOTAL: 0, Enm.COL_EMISSIONS: 1, Enm.COL_EMISSION_UNCERTAINTY: 1})
        df = df[columns_formatted]
        df = df.rename(columns=columns_to_fr)

        # Add headers
        vertical_offset = 2
        for column_i, column in enumerate(columns_formatted):
            column_i_raw = list(df_raw.columns).index(column)
            column_letter_raw = index_to_column(column_i_raw)
            for row in range(1, len(df) + 1):
                sheet_formatted.write_formula(row + vertical_offset, column_i, f"{sheet_name_raw}!${column_letter_raw}${row + 1}")

        # Resize all
        auto_resize_columns(sheet_formatted, df)  # resize columns and rows
        auto_resize_columns(sheet_raw, df_raw)  # resize columns and rows
        # Just change the transports column which can be quite big
        sheet_formatted.set_column(columns_formatted.index(Enm.COL_TRANSPORT_TYPE), columns_formatted.index(Enm.COL_TRANSPORT_TYPE), 15)

        # Format headers
        header_format = workbook.add_format({"text_wrap": True, "bold": True, "bg_color": "#ffc000", "border": 1})
        format_headers(header_format, sheet_raw, df_raw)
        format_headers(header_format, sheet_formatted, df, vertical_offset)

        # format dates
        date_format = workbook.add_format({"num_format": "yyyy-mm-dd"})
        column_index = columns_formatted.index(Enm.COL_DEPARTURE_DATE)
        sheet_formatted.set_column(column_index, column_index, 15, date_format)

        # Add computation warning
        warning_format = workbook.add_format({"italic": "true", "font_color": "red", "bold": True})
        sheet_formatted.write(
            0, 0, r"/!\ If this page looks empty, you need to recompute the formulas: CTRL+SHIFT+F9 for LibreOffice, F9 (? probably) for Excel", warning_format
        )

        # Add uncertainty text
        uncertainty_format = workbook.add_format({"font_color": "red", "text_wrap": True, "font_size": 10})
        sheet_formatted.merge_range(
            1,
            1,
            1,
            8,
            "Convention pour les incertitudes : Avion avec traînées : ±70% | Train : ±20 % (TER 60%) | Voiture, autocar, ferry : ±60 %",
            uncertainty_format,
        )

        ## Add synthetic table
        synth_column = len(columns_formatted) + 1
        synth_vertical_offset = 5
        # First column
        transport_format = workbook.add_format({"bg_color": "#ffc000", "border": 1, "align": "center", "valign": "vcenter"})
        synth_transports = ["Voiture", "Train", "Vol court", "Vol moyen", "Vol long"]
        sheet_formatted.set_column(synth_column, synth_column, 12)
        for i, s in enumerate(synth_transports):
            sheet_formatted.write(synth_vertical_offset + i + 1, synth_column, s, transport_format)
        total_format = workbook.add_format({"border": 1, "align": "center", "valign": "vcenter"})
        sheet_formatted.write(synth_vertical_offset + 1 + len(synth_transports), synth_column, "Total", total_format)
        # Second column
        transport_id_format = workbook.add_format({"bg_color": "#acb20c", "border": 1, "align": "center", "valign": "vcenter"})
        synth_transports_ids = ["ID", "car", "train *", "plane (short)", "plane (med)", "plane (long)"]
        for i, s in enumerate(synth_transports_ids):
            sheet_formatted.write(synth_vertical_offset + i, synth_column + 1, s, transport_id_format)
        sheet_formatted.set_column(synth_column + 1, synth_column + 1, 15, None, {"hidden": True})  # Hide by default
        # Headers
        synth_headers = ["Nb trajets", "Distance (1000 km)", "Emissions\n(t CO2e)", "Incertitude\n(± t CO2e)"]
        synth_header_format = workbook.add_format({"text_wrap": True, "bg_color": "#ffc000", "border": 1, "align": "center", "valign": "vcenter"})
        for i, s in enumerate(synth_headers):
            sheet_formatted.write(synth_vertical_offset, synth_column + 2 + i, s, synth_header_format)
        sheet_formatted.set_row(synth_vertical_offset, 15 * 2)
        sheet_formatted.set_column(synth_column + 2, synth_column + len(synth_headers) + 1, 12)
        # Cells
        synth_cell_format = workbook.add_format({"border": 1, "align": "center", "valign": "vcenter", "num_format": "0.0"})
        synth_cell_format_int = workbook.add_format({"border": 1, "align": "center", "valign": "vcenter", "num_format": "0"})
        synth_column_transport = index_to_column(columns_formatted.index(Enm.COL_EMISSION_TRANSPORT))
        synth_column_distance = index_to_column(columns_formatted.index(Enm.COL_DIST_TOTAL))
        synth_column_emission = index_to_column(columns_formatted.index(Enm.COL_EMISSIONS))
        synth_column_uncertainty = index_to_column(columns_formatted.index(Enm.COL_EMISSION_UNCERTAINTY))
        column_transport_id = index_to_column(synth_column + 1)
        for i, s in enumerate(synth_transports):  # N_trips
            sheet_formatted.write_formula(
                synth_vertical_offset + i + 1,
                synth_column + 2,
                f"COUNTIF({synth_column_transport}$2:{synth_column_transport}$99999, {column_transport_id}{synth_vertical_offset + 2 + i})",
                synth_cell_format_int,
            )
        for j, col in enumerate([synth_column_distance, synth_column_emission, synth_column_uncertainty]):  # other columns
            for i, transp in enumerate(synth_transports):
                sheet_formatted.write_formula(
                    synth_vertical_offset + i + 1,
                    synth_column + 3 + j,
                    f"SUMIF(${synth_column_transport}$2:${synth_column_transport}$99999, ${column_transport_id}{synth_vertical_offset + 2 + i},{col}$2:{col}$99999)/1000",
                    synth_cell_format,
                )
        synth_total_format = workbook.add_format({"bg_color": "#ffe699", "border": 1, "align": "center", "valign": "vcenter", "num_format": "0.0"})
        for j in range(4):  # totals
            col = synth_column + 2 + j
            sheet_formatted.write_formula(
                synth_vertical_offset + 6,
                synth_column + 2 + j,
                f"SUM({index_to_column(col)}{synth_vertical_offset+2}:{index_to_column(col)}{synth_vertical_offset+6})",
                synth_total_format
                if j > 0
                else workbook.add_format({"bg_color": "#ffe699", "border": 1, "align": "center", "valign": "vcenter", "num_format": "0"}),
            )
        # Trains FR
        synth_train_fr_offset = synth_vertical_offset + 8
        sheet_formatted.write(
            synth_train_fr_offset,
            synth_column,
            "Train FR",
            transport_format,
        )
        sheet_formatted.write(
            synth_train_fr_offset,
            synth_column + 1,
            "train (FR)",
            transport_id_format,
        )
        sheet_formatted.write_formula(
            synth_train_fr_offset,
            synth_column + 2,
            f"COUNTIF({synth_column_transport}$2:{synth_column_transport}$99999, {column_transport_id}{synth_train_fr_offset+1})",
            synth_cell_format_int,
        )
        for j, col in enumerate([synth_column_distance, synth_column_emission, synth_column_uncertainty]):  # other columns
            sheet_formatted.write_formula(
                synth_train_fr_offset,
                synth_column + 3 + j,
                f"SUMIF(${synth_column_transport}$2:${synth_column_transport}$99999, ${column_transport_id}{synth_train_fr_offset+1},{col}$2:{col}$99999)/1000",
                synth_cell_format,
            )
