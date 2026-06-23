#!/usr/bin/env python3
"""Shared geocoding helpers for the supply-chart skill (chart + map).

Street-level geocoding via OpenStreetMap Nominatim (cached + rate-limited), with
an offline ZIP-centroid fallback. Used to compute proximity on the chart and to
place markers on the map.
"""
from __future__ import annotations

import hashlib
import json
import math
import re
import time
import urllib.parse
import urllib.request

import zipcodes

_USER_AGENT = "supply-chart-skill/1.0 (multifamily supply map)"
_last_call = [0.0]


def haversine_miles(a, b):
    (lat1, lon1), (lat2, lon2) = a, b
    r = 3958.8
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dp = math.radians(lat2 - lat1)
    dl = math.radians(lon2 - lon1)
    h = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return 2 * r * math.asin(min(1, math.sqrt(h)))


def geocode_nominatim(query):
    gap = time.time() - _last_call[0]
    if gap < 1.1:                                  # respect Nominatim 1 req/sec
        time.sleep(1.1 - gap)
    url = ("https://nominatim.openstreetmap.org/search?format=json&limit=1&q="
           + urllib.parse.quote(query))
    req = urllib.request.Request(url, headers={"User-Agent": _USER_AGENT})
    try:
        with urllib.request.urlopen(req, timeout=25) as r:
            data = json.load(r)
    finally:
        _last_call[0] = time.time()
    return (float(data[0]["lat"]), float(data[0]["lon"])) if data else None


def zip_centroid(zipcode):
    if not zipcode:
        return None
    z = str(zipcode)[:5]
    recs = zipcodes.matching(z) if zipcodes.is_real(z) else []
    return (float(recs[0]["lat"]), float(recs[0]["long"])) if recs else None


def jitter(seed, scale=0.0055):
    h = int(hashlib.md5(seed.encode()).hexdigest(), 16)
    ang = (h % 360) * math.pi / 180
    rad = ((h // 360) % 100) / 100 * scale
    return rad * math.cos(ang), rad * math.sin(ang)


def subject_zip(address):
    m = re.search(r"\b(\d{5})\b", address or "")
    return m.group(1) if m else None


class Locator:
    """Geocode addresses with a JSON cache and ZIP-centroid fallback."""

    def __init__(self, use_geocoder, cache_path):
        self.use_geocoder = use_geocoder
        self.cache_path = cache_path
        try:
            with open(cache_path) as fh:
                self.cache = json.load(fh)
        except Exception:
            self.cache = {}
        self.dirty = False

    def save(self):
        if self.dirty:
            try:
                with open(self.cache_path, "w") as fh:
                    json.dump(self.cache, fh)
            except Exception:
                pass

    def locate(self, name, address, city, state, zipcode):
        """Return ((lat, lng), approximate_bool) or (None, True)."""
        if self.use_geocoder and address:
            al = address.lower()
            parts = [address]
            for extra in (city, state, str(zipcode) if zipcode else None):
                if extra and extra.lower() not in al:
                    parts.append(extra)
            q = ", ".join(parts)
            if q in self.cache:
                v = self.cache[q]
                if v:
                    return (v[0], v[1]), False
            else:
                try:
                    ll = geocode_nominatim(q)
                except Exception:
                    ll = None
                self.cache[q] = list(ll) if ll else None
                self.dirty = True
                if ll:
                    return ll, False
        c = zip_centroid(zipcode)
        if c:
            dy, dx = jitter(name or address or str(zipcode))
            return (c[0] + dy, c[1] + dx), True
        return None, True
