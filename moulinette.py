"""compute CO2e emissions from LMD travel data"""
import logging
from pathlib import Path

import pandas as pd

from Libs.EmissionsCalculator import compute_emissions_df
from Libs.Excel import save_to_file
from Libs.MissionsLoader import load_data
from Libs.VisualMap import generate_visual_map

# compute emission for a single excel file.
in_files = ["Missions_ex_file.xlsx"]
config_files = ["config_file_ex.cfg"]
out_file = "Missions_output_ex.xlsx"

# you can also put several files, with a unique output file :
# in_files = ["Missions_ex_file1.xlsx", "Missions_ex_file2.xlsx"]  # Data spreadsheet
# config_files = ["config_file_ex1.cfg", "config_file_ex2.cfg"]  # config file
# out_file = "Missions_output_ex1_2.xlsx"

# No need to edit below
logging.basicConfig()  # Setup printing of messages

df_data = load_data(Path("Data/Raw") / in_files[0], Path("Data/Config") / config_files[0])  # load data
df_emissions = compute_emissions_df(df_data)  # compute emissions

for i in range(1, len(in_files)):
    df_data = load_data(Path("Data/Raw") / in_files[i], Path("Data/Config") / config_files[i])  # load data
    df_emissions = pd.concat((df_emissions, compute_emissions_df(df_data)))

generate_visual_map(df_emissions, Path("Data/Generated") / out_file.replace(".xlsx", ".png"))

save_to_file(df_emissions, Path("Data/Generated") / out_file)  # format and save
