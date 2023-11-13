"""compute CO2e emissions from LMD travel data"""
import logging
from pathlib import Path

import pandas as pd

from Libs.EmissionsCalculator import compute_emissions_df
from Libs.Excel import save_to_file
from Libs.MissionsLoader import load_data

# compute emission for a single excel file.
# in_files = ["BilanMissions_test.xlsx"]
# config_files = ["config_file_LMD_X_2022.cfg"]
# out_file = "Missions_2022_test.xlsx"

# you can also put several files, with a unique output file :
in_files = ["MIS_2022_v3_aller_simple.xlsx", "Missions_LMD_ENS_2022.xlsx"]  # Data spreadsheet
config_files = ["config_file_LMD_CNRS_2022.cfg", "config_file_LMD_ENS_2022.cfg"]  # config file
out_file = "Missions_2022_output.xlsx"

# No need to edit below
logging.basicConfig()  # Setup printing of messages

df_data = load_data(Path("Data/Raw") / in_files[0], Path("Data/Config") / config_files[0])  # load data
df_emissions = compute_emissions_df(df_data)  # compute emissions

for i in range(1, len(in_files)):
    df_data = load_data(Path("Data/Raw") / in_files[i], Path("Data/Config") / config_files[i])  # load data
    df_emissions = pd.concat((df_emissions, compute_emissions_df(df_data)))

save_to_file(df_emissions, Path("Data/Generated") / out_file)  # format and save
