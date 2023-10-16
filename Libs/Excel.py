"""Excel manipulation"""
from typing import Any

import pandas as pd


def auto_resize(worksheet: Any, df: pd.DataFrame) -> None:
    """auto-resize rows and columns
    :param worksheet
    :param df: the corresponding pandas dataframe
    """
    len_newline = lambda s: max(len(i) for i in s.split("\n"))

    # adjust columns
    for idx, col in enumerate(df):  # loop through all columns
        series = df[col]
        max_len = (
            max(series.astype(str).map(len_newline).max(), len_newline(str(series.name))) + 1  # len of largest item  # len of column name/header
        )  # adding a little extra space
        worksheet.set_column(idx, idx, max_len)  # set column width

    # adjust rows
    for idx, row in df.iterrows():
        max_len = row.astype(str).map(lambda s: s.count("\n")).max() + 1
        worksheet.set_row(int(idx) + 1, max(1, max_len) * 15)
    worksheet.set_row(0, 15 * (max(str(col).count("\n") for col in df.columns) + 1))
