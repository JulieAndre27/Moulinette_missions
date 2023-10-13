"""compute CO2e emissions from LMD travel data"""
from pathlib import Path

from libs import TravelData, EmissionCalculator

in_file = "MIS_2022_v3_aller_simple.xlsx"  # Data spreadsheet
config_file = "config_file_LMD_CNRS_2022.cfg"  # config file
out_file = in_file.replace(".xlsx", "_CO2.ods")  # result will be created as <in_file>_CO2.ods

# No need to edit below

td = TravelData(Path("Data/Raw") / in_file, Path("Data/Config") / config_file)
EmissionCalculator(td, Path("Data/Generated") / out_file)
