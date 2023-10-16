"""Set shared string values as variables"""


class Enm:
    """Easy Enums"""

    # Values of roundtrip in spreadsheet
    ROUNDTRIP_YES = "oui"
    ROUNDTRIP_NO = "non"

    # Column names from config
    COL_MISSION_ID = "mission_id"
    COL_DEPARTURE_DATE = "departure_date"
    COL_DEPARTURE_CITY = "departure_city"
    COL_DEPARTURE_COUNTRY = "departure_country"
    COL_ARRIVAL_CITY = "arrival_city"
    COL_ARRIVAL_COUNTRY = "arrival_country"
    COL_TRANSPORT_TYPE = "t_type"
    COL_ROUND_TRIP = "round_trip"
    # Custom column names
    COL_CREDITS = "credits"
    COL_MAIN_TRANSPORT = "main_transport"  # the main transportation method used for computing emissions
    COL_DIST_ONE_WAY = "one_way_dist_km"
    COL_DIST_TOTAL = "final_dist_km"
    COL_EMISSIONS = "co2e_emissions_kg"
    COL_DEPARTURE_COUNTRYCODE = "departure_countrycode"
    COL_ARRIVAL_COUNTRYCODE = "arrival_countrycode"
    COL_EMISSION_TRANSPORT = "transport_for_emissions_detailed"
    COL_EMISSION_UNCERTAINTY = "emission_uncertainty"

    # Transport types from config
    TTYPES_PLANE = "t_types_plane"
    TTYPES_TRAIN = "t_types_train"
    TTYPES_CAR = "t_types_car"
    TTYPES_IGNORED = "t_types_ignored"

    # Main transport types
    MAIN_TRANSPORT_PLANE = "plane"
    MAIN_TRANSPORT_TRAIN = "train"
    MAIN_TRANSPORT_CAR = "car"
    # Emission transport_types
    EF_TRANSPORT_PLANE_SHORT = "plane (short)"
    EF_TRANSPORT_PLANE_MED = "plane (med)"
    EF_TRANSPORT_PLANE_LONG = "plane (long)"
    EF_TRANSPORT_TRAIN_TER = "train (TER, FR)"
    EF_TRANSPORT_TRAIN_FR = "train (FR)"
    EF_TRANSPORT_TRAIN_HALF_FR = "train (half-FR)"
    EF_TRANSPORT_TRAIN_EU = "train (EU)"
    EF_TRANSPORT_CAR = "car"
