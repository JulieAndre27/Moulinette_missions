# Moulinette missions

A tool to compute CO2e emisssions from mission transports from a given laboratory, with possibly different tutels (CNRS, universities...), according to Climactions-IPSL 2023 standard. See their documentation (https://climactions.ipsl.fr/) for details about the choices of the hypothesis and emissions factors taken in the code.

## Installation

* Requirements:
    * git
    * python 3.11
    * poetry (`python -m pip install poetry` outside of your virtual environment)
* Recommended:
    * Use a virtual environment. This will be done by default if you set `poetry config virtualenvs.in-project true`.
* Install dependencies
    * Pull this repository `git clone https://github.com/JulieAndre27/Moulinette_missions.git`
    * Navigate inside the directory `cd Moulinette_missions`
    * Install the environment `poetry install`
    * Activate the environment `poetry env use python3.11`
    * Setup pre-commit `pre-commit install`

## Usage

### What this tool does

From raw excel data containing the details of the lab's missions, it generates a formatted excel that shows the emissions for each mission and some aggregated data.

### Input data

Put your input excel file in `Data/Raw`. Your input file should at least contain the following columns (in no particular order): the mission ID, departure city, departure country, arrival city, arrival country, transportation used, whether it's a round trip.

To know how to interpret your data, you must provide a configuration file in `Data/Config`.

At the top of the config file, we specify the name of the sheet where the data is written.
Then, we first define which column corresponds to which data:

```
[SheetName]
# Below we define which column contains which data
mission_id = a
departure_city = c
departure_country = d
arrival_city = e
arrival_country = f
t_type = g
round_trip = i
```

We define how to interpret the transports in the `t_type` column: which correspond to plane/train/car, and which to ignore.

```
# Below we define how to link user inputs and transportation
t_type_air = avion
t_type_train = train
t_type_car = véhicule personnel, location de véhicule, taxi, passager
t_type_ignored = rer, metro, métro, bus, bateau, divers
```

And finally we define which type of credits the mission uses:

```
# Where the credits come from, ex "CNRS" or "ENS"
credits = CNRS
```

### Running the moulinette

* Specify the input excels and their configuration files at the top of `moulinette.py` (`in_files` and `config_files`).
* If you have several excels files corresponding to different credits (CNRS, university, ...), you need to give one config file per excel sheet. Their trajets data will be concatenated in the output (the credits will be indicated in a column).
* Run `python moulinette.py`
* If all goes well, this will generate in `Data/Generated` an output excel file named after the `out_file` variable. This file contains a "well presented" sheet with some aggregated data, and a sheet with all the raw data.

## Contributors
This code was adapted from a Python script called "moulinette" developped by Olivier Aumont (researcher scientist at LOCEAN) and available on demand.
The adaption was done by Julie André and Amaury Barral.
