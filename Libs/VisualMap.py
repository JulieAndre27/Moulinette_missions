from pathlib import Path

import numpy as np
import pandas as pd
from matplotlib import pyplot as plt
from matplotlib.colors import ListedColormap, Normalize
from mpl_toolkits.axes_grid1.inset_locator import inset_axes
from mpl_toolkits.basemap import Basemap

from Libs.Geo import CustomLocation, GeoDistanceCalculator


def _prepare_data(df: pd.DataFrame) -> pd.DataFrame:
    """group by destination where we are from/to paris and sum emissions"""

    # Only keep rows from/to paris
    paris_loc = CustomLocation(address="Paris, Île-de-France, France", latitude=48.8588897, longitude=2.320041, countryCode="FR")
    check_distance_to_paris = lambda x: GeoDistanceCalculator.geodesic_distance(x, paris_loc).km < 30
    departure_mask = df["departure_loc"].apply(check_distance_to_paris)
    arrival_mask = df["arrival_loc"].apply(check_distance_to_paris)
    df = df[departure_mask | arrival_mask]
    # Group by the other destination
    df["notparis_address"] = df.apply(
        lambda row: row["arrival_loc"] if check_distance_to_paris(row["departure_loc"]) else row["departure_loc"],
        axis=1,
    )
    df["emissions_oneway"] = df["co2e_emissions_kg"] / (2 * (df["round_trip"] == "oui") + 1)

    new_df = (
        df.groupby("notparis_address")
        .agg(total_emissions_oneway=("emissions_oneway", "sum"), single_emissions_oneway=("emissions_oneway", "mean"))
        .reset_index()
    )

    return new_df


def generate_visual_map(df: pd.DataFrame, output_path: Path) -> None:
    """create an image map summarizing the emissions from Paris"""
    df = _prepare_data(df.copy())
    # kg -> t
    df["single_emissions_oneway"] /= 1e3

    fig, ax = plt.subplots(figsize=(10, 8))
    m = Basemap(projection="mill", llcrnrlat=-90, urcrnrlat=90, llcrnrlon=-180, urcrnrlon=180, resolution="c", ax=ax)
    m.drawcoastlines()
    m.drawcountries()
    cmap = ListedColormap(plt.get_cmap("cool")(np.linspace(0, 1, 5)))
    normalize = Normalize(vmin=df["single_emissions_oneway"].min(), vmax=df["single_emissions_oneway"].max())

    # Plot circles on the map
    df["circle_size"] = df["total_emissions_oneway"] / df["total_emissions_oneway"].max()
    max_radius = 1e2
    for index, row in df.iterrows():
        x, y = m(row["notparis_address"].longitude, row["notparis_address"].latitude)
        m.plot(x, y, "o", markersize=max_radius * row["circle_size"], color=cmap(normalize(row["single_emissions_oneway"])), alpha=0.6)
        m.plot(x, y, "o", markersize=min(2, max_radius * row["circle_size"]), color="black", zorder=999)
    plt.title("Emissions from Paris by destination")

    # Create a colorbar
    sm = plt.cm.ScalarMappable(cmap=cmap, norm=normalize)
    sm.set_array([])  # Dummy array for the colorbar
    cbar = plt.colorbar(sm, ax=ax, orientation="vertical", shrink=0.5)
    cbar.set_label("Emissions per trip (tCO2e)")

    # Add Europe zoom encart
    axins = inset_axes(ax, width="50%", height="45%", loc="lower left", borderpad=0, bbox_to_anchor=(-0.1, 0.01, 1, 1), bbox_transform=ax.transAxes)
    m_europe = Basemap(projection="mill", llcrnrlat=30, urcrnrlat=65, llcrnrlon=-10, urcrnrlon=25, resolution="c", ax=axins)
    m_europe.drawcoastlines()
    m_europe.drawcountries()
    for index, row in df.iterrows():
        x, y = m_europe(row["notparis_address"].longitude, row["notparis_address"].latitude)
        m_europe.plot(x, y, "o", markersize=max_radius * row["circle_size"], color=cmap(normalize(row["single_emissions_oneway"])), alpha=0.6)
        m_europe.plot(x, y, "o", markersize=min(2, max_radius * row["circle_size"]), color="black", zorder=999)

    # Add legend with custom handler
    plt.savefig(output_path, dpi=300, bbox_inches="tight")
    # plt.show()
