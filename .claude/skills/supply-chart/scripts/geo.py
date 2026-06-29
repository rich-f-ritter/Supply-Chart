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
    if not data:
        return None
    # include the OSM class: "highway" means the hit is a ROAD centroid (a
    # house-number that couldn't be resolved on a long road), which gives a wrong
    # — often inflated — distance; the Locator treats it as a last-resort fallback.
    d = data[0]
    return (float(d["lat"]), float(d["lon"]), d.get("class"))


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


def parse_city_state(address):
    """Pull (city, state) out of a '..., City, ST ZIP' mailing address.

    The subject is usually given as a full mailing address; without the city/
    state the name-query geocode path is skipped and an apartment on a long road
    can fall to the road centroid. Returns (None, None) if not parseable."""
    if not address:
        return None, None
    m = re.search(r",\s*([A-Za-z .'-]+?),\s*([A-Z]{2})\b", address)
    if m:
        return m.group(1).strip(), m.group(2)
    return None, None


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
        self.anchor = None        # (lat, lng) reference for sanity-checking hits
        self.max_miles = None     # reject geocodes farther than this from anchor

    def set_anchor(self, anchor, max_miles=8.0):
        """Constrain hits to within max_miles of anchor (a 5-mile-radius dataset).

        Relaxed query variants can match a same-named street in the wrong city;
        anything well beyond the radius is a mis-geocode, so we reject it and fall
        through to the next variant / the ZIP centroid."""
        self.anchor = anchor
        self.max_miles = max_miles

    def _plausible(self, ll):
        if not (self.anchor and self.max_miles):
            return True
        return haversine_miles(self.anchor, ll) <= self.max_miles

    def save(self):
        if self.dirty:
            try:
                with open(self.cache_path, "w") as fh:
                    json.dump(self.cache, fh)
            except Exception:
                pass

    def _queries(self, address, city, state, zipcode):
        """Candidate geocode strings, most specific first.

        A wrong city (common in roster exports) makes Nominatim miss, so we also
        try the address without the city — street + state (+ZIP) usually resolves.
        """
        al = address.lower()
        z = str(zipcode) if zipcode else None

        def build(extras):
            parts = [address]
            for e in extras:
                if e and e.lower() not in al and e not in parts:
                    parts.append(e)
            return ", ".join(parts)

        seen, out = set(), []
        for extras in ([city, state, z], [state, z], [state]):
            q = build(extras)
            if q not in seen:
                seen.add(q)
                out.append(q)
        return out

    def locate(self, name, address, city, state, zipcode):
        """Return ((lat, lng), approximate_bool) or (None, True)."""
        if self.use_geocoder:
            queries = self._queries(address, city, state, zipcode) if address else []
            # Apartment complexes geocode reliably by NAME; a suburban house-number
            # on a long road often resolves only to the road centroid, so also try
            # the name and prefer a real point over a road hit.
            if name and city and state:
                queries = list(queries) + [f"{name}, {city}, {state}"]
            road = None                            # plausible road-centroid fallback
            for q in queries:
                if q in self.cache:
                    v = self.cache[q]
                else:
                    try:
                        g = geocode_nominatim(q)
                    except Exception:
                        g = None
                    v = list(g) if g else None
                    self.cache[q] = v
                    self.dirty = True
                if not v:
                    continue
                ll = (v[0], v[1])
                if not self._plausible(ll):
                    continue
                if len(v) > 2 and v[2] == "highway":
                    road = road or ll              # imprecise — keep as last resort
                    continue
                return ll, False
            if road:
                return road, False
        c = zip_centroid(zipcode)
        if c:
            dy, dx = jitter(name or address or str(zipcode))
            return (c[0] + dy, c[1] + dx), True
        return None, True
