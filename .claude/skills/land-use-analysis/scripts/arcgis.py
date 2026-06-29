"""Robust ArcGIS REST fetching — the part that kept breaking on real layers.

Why this exists: the naive "GET + returnCountOnly quadtree" pattern fails on common layers:
  * big statewide/county layers DISABLE returnCountOnly, and
  * putting the bbox geometry in a GET querystring blows past the server URL-length limit,
    so the request 400s or silently returns nothing.
This module instead:
  * uses POST for every query (geometry rides in the body — no URL-length limit),
  * pages by OBJECTID (returnIdsOnly -> fetch in id-batches): no dependency on count, works
    on layers that cap transfers, and is naturally resumable,
  * falls back to a POST bbox quadtree (truncation detected via exceededTransferLimit /
    full-page heuristic) when a layer won't return ids,
  * checkpoints every batch to a dir so a huge pull resumes across sessions/interruptions,
  * can thin geometry server-side (maxAllowableOffset) for dense condo metros.

f=geojson is requested (widely supported on ArcGIS 10.4+/hosted). If a layer returns esri
json only, set the source's "format":"esrijson" — handled by a minimal converter here.
"""
import json
import sys
from pathlib import Path

import requests
import urllib3
from shapely.geometry import shape

urllib3.disable_warnings()
H = {"User-Agent": "Mozilla/5.0"}


def _post(url, params, verify, timeout=240):
    r = requests.post(url, data=params, headers=H, timeout=timeout, verify=verify)
    r.raise_for_status()
    return r.json()


def _env(bbox):
    return {"xmin": bbox[0], "ymin": bbox[1], "xmax": bbox[2], "ymax": bbox[3],
            "spatialReference": {"wkid": 4326}}


# ---- esri json -> geojson (minimal; rings/paths/points) --------------------------
def _ring_to_gj(rings):
    # ArcGIS: first ring CW = outer, holes CCW. GeoJSON polygon: [outer, holes...].
    # For analysis we don't need strict winding; emit each ring as its own polygon and
    # let shapely .buffer(0) fix it downstream.
    return {"type": "Polygon", "coordinates": rings} if len(rings) == 1 else \
           {"type": "MultiPolygon", "coordinates": [[r] for r in rings]}


def _esri_to_feature(att, geom):
    g = None
    if geom:
        if "rings" in geom:
            g = _ring_to_gj(geom["rings"])
        elif "paths" in geom:
            paths = geom["paths"]
            g = {"type": "LineString", "coordinates": paths[0]} if len(paths) == 1 else \
                {"type": "MultiLineString", "coordinates": paths}
        elif "x" in geom:
            g = {"type": "Point", "coordinates": [geom["x"], geom["y"]]}
    return {"type": "Feature", "geometry": g, "properties": att}


def _features(j, fmt):
    if fmt == "esrijson":
        return [_esri_to_feature(f.get("attributes", {}), f.get("geometry")) for f in j.get("features", [])]
    return j.get("features", [])


# ---- primary: objectId paging ----------------------------------------------------
def get_ids(url, bbox, where, verify):
    p = {"where": where, "geometry": json.dumps(_env(bbox)),
         "geometryType": "esriGeometryEnvelope", "inSR": "4326",
         "spatialRel": "esriSpatialRelIntersects", "returnIdsOnly": "true", "f": "json"}
    j = _post(url, p, verify)
    return (j.get("objectIdFieldName") or "OBJECTID"), (j.get("objectIds") or [])


def _fetch_ids(url, ids, out_fields, verify, page, simplify, fmt, ck_dir, label):
    parts = []
    nb = (len(ids) + page - 1) // page
    for bi, i in enumerate(range(0, len(ids), page)):
        ckf = ck_dir / f"{label}_b{bi:04d}.json"
        if ckf.exists() and ckf.stat().st_size > 0:
            parts.append(ckf)
            continue
        p = {"objectIds": ",".join(map(str, ids[i:i + page])), "outFields": out_fields or "*",
             "returnGeometry": "true", "outSR": "4326", "f": "json" if fmt == "esrijson" else "geojson"}
        if simplify:
            p["maxAllowableOffset"] = str(simplify)
        feats = _features(_post(url, p, verify), fmt)
        ckf.write_text(json.dumps({"type": "FeatureCollection", "features": feats}))
        parts.append(ckf)
        print(f"    {label}: batch {bi + 1}/{nb}  (+{len(feats)})", file=sys.stderr)
    out = []
    for ckf in parts:
        out.extend(json.loads(ckf.read_text()).get("features", []))
    return out


# ---- fallback: POST bbox quadtree ------------------------------------------------
def _fetch_cell(url, bbox, where, out_fields, verify, simplify, fmt, cap):
    p = {"where": where, "geometry": json.dumps(_env(bbox)),
         "geometryType": "esriGeometryEnvelope", "inSR": "4326",
         "spatialRel": "esriSpatialRelIntersects", "outFields": out_fields or "*",
         "returnGeometry": "true", "outSR": "4326",
         "f": "json" if fmt == "esrijson" else "geojson", "resultRecordCount": cap}
    if simplify:
        p["maxAllowableOffset"] = str(simplify)
    j = _post(url, p, verify)
    feats = _features(j, fmt)
    truncated = bool(j.get("exceededTransferLimit")) or len(feats) >= cap
    return feats, truncated


def _quadtree(url, bbox, where, out_fields, verify, simplify, fmt, cap, depth=0, maxdepth=9):
    feats, truncated = _fetch_cell(url, bbox, where, out_fields, verify, simplify, fmt, cap)
    if not truncated or depth >= maxdepth:
        print(f"    leaf d{depth} got={len(feats)}", file=sys.stderr)
        return feats
    x0, y0, x1, y1 = bbox
    mx, my = (x0 + x1) / 2, (y0 + y1) / 2
    out = []
    for s in ((x0, y0, mx, my), (mx, y0, x1, my), (x0, my, mx, y1), (mx, my, x1, y1)):
        out.extend(_quadtree(url, s, where, out_fields, verify, simplify, fmt, cap, depth + 1, maxdepth))
    return out


# ---- public entry ----------------------------------------------------------------
def fetch_all(url, bbox, *, where="1=1", out_fields="*", verify=True, page=1000,
              simplify=None, fmt="geojson", ck_dir=None, label="layer", cap=1000):
    """Return a list of GeoJSON features for the layer within bbox. Tries objectId paging,
    falls back to a POST quadtree. Checkpoints id-batches under ck_dir for resume."""
    ck_dir = Path(ck_dir) if ck_dir else Path(".")
    ck_dir.mkdir(parents=True, exist_ok=True)
    try:
        oid_field, ids = get_ids(url, bbox, where, verify)
    except Exception as e:  # noqa: BLE001
        print(f"    [{label}] returnIdsOnly failed ({str(e)[:70]}); using quadtree", file=sys.stderr)
        ids = None
    if ids:
        print(f"    [{label}] {len(ids)} object ids -> paging by {page}", file=sys.stderr)
        return _fetch_ids(url, ids, out_fields, verify, page, simplify, fmt, ck_dir, label)
    if ids == []:
        print(f"    [{label}] 0 ids in bbox", file=sys.stderr)
        return []
    return _quadtree(url, bbox, where, out_fields, verify, simplify, fmt, cap)


def keep_in_poly(features, poly):
    out = []
    for ft in features:
        try:
            if ft.get("geometry") and shape(ft["geometry"]).representative_point().within(poly):
                out.append(ft)
        except Exception:
            continue
    return out
