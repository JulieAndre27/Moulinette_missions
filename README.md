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

*TODO*


## TODO

* keep a column with the simple (not A/R) distance in the out file
* add a column incertitude in pct CO2 (to be computed later)
* add a plane column to know if it's short medium long
* cache geo results
