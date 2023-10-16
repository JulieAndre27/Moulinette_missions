"""compute CO2e emissions from LMD travel data"""
import logging
from pathlib import Path

from Libs.EmissionsCalculator import compute_emissions_df, format_emissions_df, save_to_file
from Libs.MissionsLoader import load_data

in_file = "MIS_2022_extrait.xlsx"  # Data spreadsheet
config_file = "config_file_LMD_CNRS_2022.cfg"  # config file
out_file = in_file.replace(".xlsx", "_CO2.xlsx")  # result will be created as <in_file>_CO2.ods

# No need to edit below
logging.basicConfig()  # Setup printing of messages

df_data = load_data(Path("Data/Raw") / in_file, Path("Data/Config") / config_file)  # load data
df_emissions = compute_emissions_df(df_data)  # compute emissions
save_to_file(format_emissions_df(df_emissions), Path("Data/Generated") / out_file)  # format and save
