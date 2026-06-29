"""Geometry helpers, locale-independent. The analysis AREA comes from the config:
either a radius circle around the subject or a hand-drawn polygon. A local planar CRS
for accurate acreage is chosen automatically (US State Plane is not assumed) - we use
the appropriate UTM zone for the subject longitude unless the config pins crs_local.
"""
import math

import pyproj
from shapely.geometry import Point, Polygon, shape
from shapely.ops import transform


def subject_lonlat(cfg):
    s = cfg["subject"]
    return s["lon"], s["lat"]


def analysis_polygon(cfg):
    """Return the shapely analysis-area polygon (lon, lat)."""
    a = cfg["analysis_area"]
    if a.get("mode") == "polygon":
        return Polygon([(x, y) for x, y in a["polygon"]])
    # radius circle: buffer the subject point in a local azimuthal projection so the
    # circle is a true <radius> miles in every direction, then bring back to lon/lat.
    lon, lat = subject_lonlat(cfg)
    radius_m = float(a.get("radius_mi", 2.0)) * 1609.344
    aeqd = pyproj.CRS.from_proj4(
        f"+proj=aeqd +lat_0={lat} +lon_0={lon} +datum=WGS84 +units=m +no_defs")
    fwd = pyproj.Transformer.from_crs("EPSG:4326", aeqd, always_xy=True).transform
    inv = pyproj.Transformer.from_crs(aeqd, "EPSG:4326", always_xy=True).transform
    circle_local = transform(fwd, Point(lon, lat)).buffer(radius_m, resolution=64)
    return transform(inv, circle_local)


def bbox(cfg):
    return analysis_polygon(cfg).bounds  # (minlon, minlat, maxlon, maxlat)


def rings_4326(cfg):
    """esriGeometryPolygon (WGS84) for ArcGIS spatial queries."""
    poly = analysis_polygon(cfg)
    return {"rings": [[[x, y] for x, y in poly.exterior.coords]],
            "spatialReference": {"wkid": 4326}}


def utm_epsg(lon, lat):
    zone = int((lon + 180) / 6) + 1
    return (32600 if lat >= 0 else 32700) + zone  # WGS84 / UTM north|south


def local_transform(cfg):
    """A 4326 -> local-planar-metre transform for accurate area/length."""
    crs = cfg.get("crs_local")
    if not crs:
        lon, lat = subject_lonlat(cfg)
        crs = f"EPSG:{utm_epsg(lon, lat)}"
    return pyproj.Transformer.from_crs("EPSG:4326", crs, always_xy=True).transform


def acres(geom_lonlat, to_local):
    """Geometric acreage of a lon/lat geometry using a local planar transform."""
    g = geom_lonlat if hasattr(geom_lonlat, "area") else shape(geom_lonlat)
    return transform(to_local, g).area / 4046.8564224  # m^2 -> acres


def haversine_mi(lat1, lon1, lat2, lon2):
    R = 3958.7613
    p1, p2 = math.radians(lat1), math.radians(lat2)
    a = (math.sin(math.radians(lat2 - lat1) / 2) ** 2 +
         math.cos(p1) * math.cos(p2) * math.sin(math.radians(lon2 - lon1) / 2) ** 2)
    return 2 * R * math.asin(math.sqrt(a))
