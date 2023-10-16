"""Excel manipulation"""
from typing import Any

import pandas as pd


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
