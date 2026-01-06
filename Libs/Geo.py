"""Utilities to handle locations and distances"""

from __future__ import annotations

import base64
import logging
import os
from dataclasses import dataclass

import geopy
import orjson
from geopy import geocoders
from geopy.distance import geodesic

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)


@dataclass
class CustomLocation:
    """Store a location, serializable"""

    address: str
    latitude: float
    longitude: float
    countryCode: str | None

    @classmethod
    def from_location(cls, loc: geopy.location.Location) -> CustomLocation:
        """init from geopy location"""
        return cls(loc.address, loc.latitude, loc.longitude, loc.raw.get("components", {}).get("ISO_3166-1_alpha-2", None))

    @classmethod
    def from_json(cls, data: dict) -> CustomLocation:
        """init from json dict"""
        return cls(data["address"], data["latitude"], data["longitude"], data["countryCode"])

    def __hash__(self):
        return hash((self.latitude, self.longitude))

    def __lt__(self, other):  # for ordering
        return self.latitude < other.latitude or (self.latitude == other.latitude and self.longitude < other.longitude)


class GeoDistanceCalculator:
    """compute distance between locations, cached"""

    def __init__(self):
        self.locator = geocoders.OpenCage(api_key=base64.b64decode("MGQzODYzMjJlZmE3NGE3MGE1NDhlNzEzYzI3MDI2ZWM=").decode("utf-8"), timeout=30)  # obviously not meant to be robust obfuscation - only to discourage automated scrapers
        self.N_api_calls = 0
        self.N_cache_hits = 0
        self.cache_path = "Data/Config/geocache.json"
        self.save_cache_every_N = 100

        # load cache
        if not os.path.isfile(self.cache_path):  # make sure file exists
            with open(self.cache_path, "w", encoding="utf8") as f:
                f.write("{}")

        with open(self.cache_path, "r", encoding="utf8") as f:
            cache_json = orjson.loads(f.read())
        self.cache = {k: CustomLocation.from_json(v) for k, v in cache_json.items()}

    def print_usage(self) -> None:
        """output usage"""
        logger.info(f"API calls: {self.N_api_calls}, cache hits: {self.N_cache_hits}")

    def save_cache(self, force=False) -> None:
        """save cache to disk"""
        if force or len(self.cache) % self.save_cache_every_N == 0:
            with open(self.cache_path, "wb") as f:
                f.write(orjson.dumps(self.cache))

    def find_location(self, loc_str: str) -> CustomLocation:
        """find location of place"""
        if loc_str in self.cache:
            self.N_cache_hits += 1
            return self.cache[loc_str]

        self.N_api_calls += 1
        coord = self.locator.geocode(loc_str, language="fr")
        if coord is None:
            raise ValueError(f"Error: could not locate <{loc_str}>")

        coord = CustomLocation.from_location(coord)

        self.cache[loc_str] = coord
        self.save_cache()

        return coord

    @staticmethod
    def geodesic_distance(loc1: CustomLocation, loc2: CustomLocation) -> geopy.distance.geodesic:
        """geodesic distance between two locations"""
        return geodesic(
            (loc1.latitude, loc1.longitude),
            (loc2.latitude, loc2.longitude),
        )

    def get_geodesic_distance_between(self, str1: str, str2: str) -> tuple[float, CustomLocation | None, CustomLocation | None]:
        """get the geo distance (km) between two places from their address

        :returns (distance, geocode loc1, geocode loc2)
        """
        if str1.startswith("nan (") or str2.startswith("nan ("):  # missing city
            return 0, None, None
        coord1 = self.find_location(str1)
        coord2 = self.find_location(str2)

        return self.geodesic_distance(coord1, coord2).km, coord1, coord2
