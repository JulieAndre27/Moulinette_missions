# Moulinette missions

A tool to compute CO2e emisssions from transports in LMD, from ENS and CNRS data.

## How to run

* Requirements:
    * python 3.11
    * poetry (`python -m pip install poetry` outside of your virtual environment)
* Recommended:
    * A virtual environment (`python -m poetry config virtualenvs.in-project true`)

* Install dependencies `poetry install`

* Run `python moulinette.py`

## Structure

Raw speadsheets go in `Data/Raw`, config files in `Data/Config`, and generated data appears in `Data/Generated`.

## Config

Config files follow this format:

```
[SheetName]
# Below we define which column contains which data
departure_city = c
departure_country = d
arrival_city = e
arrival_country = f
t_type = g
round_trip = i

# Below we define how to link user inputs and transportation
t_type_air = avion
t_type_train = train
t_type_car = véhicule personnel, location de véhicule, taxi, passager
t_type_ignored = rer, metro, métro, bus, bateau, divers
```



