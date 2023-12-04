from __future__ import annotations

import configparser
import logging
from pathlib import Path

import pandas as pd

from Libs.EasyEnums import Enm
from Libs.misc import split_strip_lower

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)


def load_data(data_path: Path, conf_path: Path) -> pd.DataFrame:
    """load data from excel/ods according to the conf
    corrects the data to account for wrong round trips
    :param data_path: path to the .xlsx mission data
    :param conf_path: path to the .cfg conf
    :returns a pandas df with the corrected data
    """
    # Check paths
    assert data_path.exists(), f"The raw data file {data_path} doesn't exist"
    assert conf_path.exists(), f"The config file {conf_path} doesn't exist"
    # load conf
    config = configparser.RawConfigParser()
    config.read(conf_path, encoding="utf-8")
    config_dict = {s: dict(config.items(s)) for s in config.sections()}
    assert len(config_dict.keys()) == 1, "Exactly one sheet must be specified in the configuration"
    sheet_name = list(config_dict.keys())[0]
    config_dict = config_dict[sheet_name]

    # column letter to index
    column_names = [
        Enm.COL_MISSION_ID,
        Enm.COL_DEPARTURE_DATE,
        Enm.COL_DEPARTURE_CITY,
        Enm.COL_DEPARTURE_COUNTRY,
        Enm.COL_ARRIVAL_CITY,
        Enm.COL_ARRIVAL_COUNTRY,
        Enm.COL_TRANSPORT_TYPE,
        Enm.COL_ROUND_TRIP,
    ]
    column_ids = [ord(config_dict[v].lower()) - 97 for v in column_names]
    # re-order column_names to match column_ids order
    column_names = [i[0] for i in sorted(list(zip(column_names, column_ids)), key=lambda x: x[1])]

    # load sheet
    # noinspection PyTypeChecker
    df_data = pd.read_excel(data_path, sheet_name=sheet_name, usecols=column_ids, names=column_names)
    df_data.attrs["sheet_name"] = sheet_name
    df_data.attrs["credits"] = config_dict["credits"]

    # determine computed transport type
    unknown_transport_types = set()

    df_data[Enm.COL_MAIN_TRANSPORT] = df_data.apply(_get_main_transport, axis=1, args=(config_dict, unknown_transport_types))
    if len(unknown_transport_types):
        logger.warning(f"Unknown transport types: {unknown_transport_types}")

    # sanitize data
    _fix_round_trips(df_data)

    return df_data


def _fix_round_trips(data: pd.DataFrame) -> None:
    """Fix incorrect round trips
    If a mission ID has several trips, set them all to one-way"""

    data[Enm.COL_ROUND_TRIP] = data[Enm.COL_ROUND_TRIP].apply(str.lower)
    duplicated_mission_loc = data[Enm.COL_MISSION_ID].duplicated(keep=False)
    wrong_roundtrips = data[Enm.COL_MISSION_ID].notna() & duplicated_mission_loc & (data[Enm.COL_ROUND_TRIP] == Enm.ROUNDTRIP_YES)

    N_wrong_roundtrips = wrong_roundtrips.sum()
    if N_wrong_roundtrips > 0:
        logger.info(f"Corrected {N_wrong_roundtrips} round_trips")

    data.loc[wrong_roundtrips, Enm.COL_ROUND_TRIP] = Enm.ROUNDTRIP_NO


def _get_main_transport(row: pd.Series, config_dict: dict, unknown_transport_types: set) -> str | None:
    """get main transport used for computation of emissions"""
    # Transportation descriptors
    ttypes_air, ttypes_train, ttypes_car, ttypes_ignored = (
        split_strip_lower(config_dict[i]) for i in [Enm.TTYPES_PLANE, Enm.TTYPES_TRAIN, Enm.TTYPES_CAR, Enm.TTYPES_IGNORED]
    )

    if isinstance(row.t_type, str):  # if it isn't missing
        row_transports = split_strip_lower(row.t_type)  # row's used transport modes
        # Check that we recognize all existing types
        for transport in row_transports:
            if transport not in ttypes_air and transport not in ttypes_train and transport not in ttypes_car and transport not in ttypes_ignored:
                unknown_transport_types.add(transport)

        transport_priority_order = [
            (ttypes_air, Enm.MAIN_TRANSPORT_PLANE),
            (ttypes_train, Enm.MAIN_TRANSPORT_TRAIN),
            (ttypes_car, Enm.MAIN_TRANSPORT_CAR),
        ]
        for ttypes, main_transport in transport_priority_order:  # plane > train > car
            for ttype in ttypes:  # for each known transportation method
                if ttype in row_transports:  # check if used and return
                    return main_transport

    return None
