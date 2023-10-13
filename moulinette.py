"""compute CO2e emissions from LMD travel data"""
from pathlib import Path

from libs import TravelData, EmissionCalculator

in_file = "Missions_LMD_ENS_2022.xlsx"  # Data spreadsheet
config_file = "config_file_LMD_ENS_2022.cfg"  # config file

td = TravelData(Path("Data/Raw") / in_file, Path("Data/Config") / config_file)
EmissionCalculator().compute(td, Path("Data/Generated") / in_file.replace(".xlsx", "_CO2.ods"))  # result will be created as <in_file>_CO2.ods
