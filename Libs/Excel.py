"""Excel manipulation"""
from pathlib import Path
from typing import Any

import pandas as pd

from Libs.EasyEnums import Enm
from Libs.misc import index_to_column


def auto_resize_columns(worksheet: Any, df: pd.DataFrame, cell_format=None) -> None:
    """auto-resize rows and columns
    :param worksheet
    :param df: the corresponding pandas dataframe
    :param cell_format
    """
    len_newline = lambda s: max(len(i) for i in s.split("\n"))

    # adjust columns
    for idx, col in enumerate(df):  # loop through all columns
        series = df[col]
        max_len = (
            max(series.astype(str).map(len_newline).max(), len_newline(str(series.name))) + 1  # len of largest item  # len of column name/header
        )  # adding a little extra space
        worksheet.set_column(idx, idx, max_len, cell_format)  # set column width


def format_headers(header_format: dict, worksheet: Any, df: pd.DataFrame, offset: int = 0) -> None:
    """format the headers"""
    worksheet.set_row(offset, 15 * (max(str(col).count("\n") for col in df.columns) + 1))  # resize

    # format each cell
    for i, col in enumerate(df.columns):
        worksheet.write(offset, i, col, header_format)


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

        # Define some common format dicts
        f_border_align_center = {"border": 1, "align": "center", "valign": "vcenter"}

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
        transport_format = workbook.add_format({"bg_color": "#ffc000", **f_border_align_center})
        synth_transports = ["Voiture", "Train", "Vol court", "Vol moyen", "Vol long"]
        sheet_formatted.set_column(synth_column, synth_column, 12)
        for i, s in enumerate(synth_transports):
            sheet_formatted.write(synth_vertical_offset + i + 1, synth_column, s, transport_format)
        total_format = workbook.add_format({"border": 1, "align": "center", "valign": "vcenter"})
        sheet_formatted.write(synth_vertical_offset + 1 + len(synth_transports), synth_column, "Total", total_format)
        # Second column
        transport_id_format = workbook.add_format({"bg_color": "#acb20c", **f_border_align_center})
        synth_transports_ids = ["ID", "car", "train *", "plane (short)", "plane (med)", "plane (long)"]
        for i, s in enumerate(synth_transports_ids):
            sheet_formatted.write(synth_vertical_offset + i, synth_column + 1, s, transport_id_format)
        sheet_formatted.set_column(synth_column + 1, synth_column + 1, 15, None, {"hidden": True})  # Hide by default
        # Headers
        synth_headers = ["Nb trajets", "Distance (1000 km)", "Emissions\n(t CO2e)", "Incertitude\n(± t CO2e)"]
        synth_header_format = workbook.add_format({"text_wrap": True, "bg_color": "#ffc000", **f_border_align_center})
        for i, s in enumerate(synth_headers):
            sheet_formatted.write(synth_vertical_offset, synth_column + 2 + i, s, synth_header_format)
        sheet_formatted.set_row(synth_vertical_offset, 15 * 2)
        sheet_formatted.set_column(synth_column + 2, synth_column + len(synth_headers) + 1, 12)
        # Cells
        synth_cell_format = workbook.add_format({**f_border_align_center, "num_format": "0.0"})
        synth_cell_format_int = workbook.add_format({**f_border_align_center, "num_format": "0"})
        synth_column_transport = index_to_column(columns_formatted.index(Enm.COL_EMISSION_TRANSPORT))
        synth_column_distance = index_to_column(columns_formatted.index(Enm.COL_DIST_TOTAL))
        synth_column_emission = index_to_column(columns_formatted.index(Enm.COL_EMISSIONS))
        synth_column_uncertainty = index_to_column(columns_formatted.index(Enm.COL_EMISSION_UNCERTAINTY))
        column_transport_id = index_to_column(synth_column + 1)
        for i, s in enumerate(synth_transports):  # N_trips
            sheet_formatted.write_formula(
                synth_vertical_offset + i + 1,
                synth_column + 2,
                f"COUNTIF(${synth_column_transport}:${synth_column_transport}, {column_transport_id}{synth_vertical_offset + 2 + i})",
                synth_cell_format_int,
            )
        for j, col in enumerate([synth_column_distance, synth_column_emission, synth_column_uncertainty]):  # other columns
            for i, transp in enumerate(synth_transports):
                sheet_formatted.write_formula(
                    synth_vertical_offset + i + 1,
                    synth_column + 3 + j,
                    f"SUMIF(${synth_column_transport}:${synth_column_transport}, ${column_transport_id}{synth_vertical_offset + 2 + i},{col}:{col})/1000",
                    synth_cell_format,
                )
        synth_total_format = workbook.add_format({"bg_color": "#ffe699", **f_border_align_center, "num_format": "0.0"})
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
            f"COUNTIF(${synth_column_transport}:${synth_column_transport}, {column_transport_id}{synth_train_fr_offset+1})",
            synth_cell_format_int,
        )
        for j, col in enumerate([synth_column_distance, synth_column_emission, synth_column_uncertainty]):  # other columns
            sheet_formatted.write_formula(
                synth_train_fr_offset,
                synth_column + 3 + j,
                f"SUMIF(${synth_column_transport}:${synth_column_transport}, ${column_transport_id}{synth_train_fr_offset+1},{col}:{col})/1000",
                synth_cell_format,
            )

        ## Add trip comparator
        tcomp_column = synth_column
        tcomp_vertical_offset = synth_vertical_offset + 10
        # Input and description
        sheet_formatted.write(tcomp_vertical_offset, tcomp_column, "Pour un trajet précis", workbook.add_format({"bold": True}))
        sheet_formatted.write(
            tcomp_vertical_offset, tcomp_column + 3, 'Utilisez "*" pour indiquer une valeur quelconque', workbook.add_format({"italic": True})
        )
        sheet_formatted.write(tcomp_vertical_offset + 1, tcomp_column + 3, "Ville 1:")
        sheet_formatted.write(tcomp_vertical_offset + 1, tcomp_column + 4, "Paris")
        sheet_formatted.write(tcomp_vertical_offset + 1, tcomp_column + 5, "Pays 1:")
        sheet_formatted.write(tcomp_vertical_offset + 1, tcomp_column + 6, "FR")
        sheet_formatted.write(tcomp_vertical_offset + 2, tcomp_column + 3, "Ville 2:")
        sheet_formatted.write(tcomp_vertical_offset + 2, tcomp_column + 4, "Vienne")
        sheet_formatted.write(tcomp_vertical_offset + 2, tcomp_column + 5, "Pays 2:")
        sheet_formatted.write(tcomp_vertical_offset + 2, tcomp_column + 6, "AT")
        tcomp_vertical_offset += 3
        # First column
        transport_format = workbook.add_format({"bg_color": "#ffc000", **f_border_align_center})
        tcomp_transports = ["Voiture", "Train", "Vol"]
        sheet_formatted.set_column(tcomp_column, tcomp_column, 12)
        for i, s in enumerate(tcomp_transports):
            sheet_formatted.write(tcomp_vertical_offset + i + 1, tcomp_column, s, transport_format)
        total_format = workbook.add_format({"border": 1, "align": "center", "valign": "vcenter"})
        sheet_formatted.write(tcomp_vertical_offset + 1 + len(tcomp_transports), tcomp_column, "Total", total_format)
        # Second column
        transport_id_format = workbook.add_format({"bg_color": "#acb20c", **f_border_align_center})
        tcomp_transports_ids = ["ID", "car", "train *", "plane*"]
        for i, s in enumerate(tcomp_transports_ids):
            sheet_formatted.write(tcomp_vertical_offset + i, tcomp_column + 1, s, transport_id_format)
        sheet_formatted.set_column(tcomp_column + 1, tcomp_column + 1, 15, None, {"hidden": True})  # Hide by default
        # Headers
        synth_headers = ["Nb trajets", "Distance (1000 km)", "Emissions\n(t CO2e)", "Incertitude\n(± t CO2e)"]
        synth_header_format = workbook.add_format({"text_wrap": True, "bg_color": "#ffc000", **f_border_align_center})
        for i, s in enumerate(synth_headers):
            sheet_formatted.write(tcomp_vertical_offset, tcomp_column + 2 + i, s, synth_header_format)
        sheet_formatted.set_row(tcomp_vertical_offset, 15 * 2)
        sheet_formatted.set_column(tcomp_column + 2, tcomp_column + len(synth_headers) + 1, 12)
        # Cells
        synth_cell_format = workbook.add_format({**f_border_align_center, "num_format": "0.0"})
        synth_cell_format_int = workbook.add_format({**f_border_align_center, "num_format": "0"})
        tcomp_column_transport = index_to_column(columns_formatted.index(Enm.COL_EMISSION_TRANSPORT))
        tcomp_column_distance = index_to_column(columns_formatted.index(Enm.COL_DIST_TOTAL))
        tcomp_column_emission = index_to_column(columns_formatted.index(Enm.COL_EMISSIONS))
        tcomp_column_uncertainty = index_to_column(columns_formatted.index(Enm.COL_EMISSION_UNCERTAINTY))
        tcomp_column_departure_city = index_to_column(columns_formatted.index(Enm.COL_DEPARTURE_CITY))
        tcomp_column_departure_country = index_to_column(columns_formatted.index(Enm.COL_DEPARTURE_COUNTRYCODE))
        tcomp_column_arrival_city = index_to_column(columns_formatted.index(Enm.COL_ARRIVAL_CITY))
        tcomp_column_arrival_country = index_to_column(columns_formatted.index(Enm.COL_ARRIVAL_COUNTRYCODE))
        column_transport_id = index_to_column(tcomp_column + 1)
        tcomp_condition_1 = f"${tcomp_column_departure_city}:${tcomp_column_departure_city}, ${index_to_column(tcomp_column+4)}${tcomp_vertical_offset-1}, ${tcomp_column_departure_country}:${tcomp_column_departure_country}, ${index_to_column(tcomp_column+6)}${tcomp_vertical_offset-1}, ${tcomp_column_arrival_city}:${tcomp_column_arrival_city}, ${index_to_column(tcomp_column+4)}${tcomp_vertical_offset}, ${tcomp_column_arrival_country}:${tcomp_column_arrival_country}, ${index_to_column(tcomp_column+6)}${tcomp_vertical_offset}"
        tcomp_condition_2 = f"${tcomp_column_arrival_city}:${tcomp_column_arrival_city}, ${index_to_column(tcomp_column+4)}${tcomp_vertical_offset-1}, ${tcomp_column_arrival_country}:${tcomp_column_arrival_country}, ${index_to_column(tcomp_column+6)}${tcomp_vertical_offset-1}, ${tcomp_column_departure_city}:${tcomp_column_departure_city}, ${index_to_column(tcomp_column+4)}${tcomp_vertical_offset}, ${tcomp_column_departure_country}:${tcomp_column_departure_country}, ${index_to_column(tcomp_column+6)}${tcomp_vertical_offset}"
        for i, s in enumerate(tcomp_transports):  # N_trips
            sheet_formatted.write_formula(
                tcomp_vertical_offset + i + 1,
                tcomp_column + 2,
                f"COUNTIFS(${tcomp_column_transport}:${tcomp_column_transport}, {column_transport_id}{tcomp_vertical_offset + 2 + i}, {tcomp_condition_1})+COUNTIFS(${tcomp_column_transport}:${tcomp_column_transport}, {column_transport_id}{tcomp_vertical_offset + 2 + i}, {tcomp_condition_2})",
                synth_cell_format_int,
            )
        for j, col in enumerate([tcomp_column_distance, tcomp_column_emission, tcomp_column_uncertainty]):  # other columns
            for i, transp in enumerate(tcomp_transports):
                sheet_formatted.write_formula(
                    tcomp_vertical_offset + i + 1,
                    tcomp_column + 3 + j,
                    f"(SUMIFS({col}:{col}, ${tcomp_column_transport}:${tcomp_column_transport}, ${column_transport_id}{tcomp_vertical_offset + 2 + i}, {tcomp_condition_1})+SUMIFS({col}:{col}, ${tcomp_column_transport}:${tcomp_column_transport}, ${column_transport_id}{tcomp_vertical_offset + 2 + i}, {tcomp_condition_2}))/1000",
                    synth_cell_format,
                )
        synth_total_format = workbook.add_format({"bg_color": "#ffe699", **f_border_align_center, "num_format": "0.0"})
        for j in range(4):  # totals
            col = tcomp_column + 2 + j
            sheet_formatted.write_formula(
                tcomp_vertical_offset + 4,
                tcomp_column + 2 + j,
                f"SUM({index_to_column(col)}{tcomp_vertical_offset+2}:{index_to_column(col)}{tcomp_vertical_offset+4})",
                synth_total_format
                if j > 0
                else workbook.add_format({"bg_color": "#ffe699", "border": 1, "align": "center", "valign": "vcenter", "num_format": "0"}),
            )
