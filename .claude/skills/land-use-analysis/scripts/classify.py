"""Classify every parcel against the per-run config:
  * land-use bucket  (config.landuse_scheme.codes; owner-rule -> Public; degraded source)
  * zoning category + MF-threat  (config.zoning_crosswalk via AREA-WEIGHTED overlap)
  * vacant flag, geometric acreage, distance to subject.
Land use and zoning are treated as INDEPENDENT layers (a commercial-zoned parcel can
legitimately carry MF use = legal nonconforming). Writes parcels_classified.geojson,
the three crosswalk CSV tables, and an unmapped-codes log for the decisions trail.
"""
import csv
import json
import re
import sys
from collections import Counter
from pathlib import Path

from shapely.geometry import shape
from shapely.strtree import STRtree

sys.path.insert(0, str(Path(__file__).resolve().parent))
import config as C  # noqa: E402
import geo  # noqa: E402
import palette as P  # noqa: E402

GAP = P.DATA_GAP_LABEL


def _valid(g):
    if not g.is_valid:
        g = g.buffer(0)
    return g


def main():
    cfg = C.load()
    IN, TBL = C.indir(cfg), C.tables(cfg)
    to_local = geo.local_transform(cfg)
    slon, slat = geo.subject_lonlat(cfg)

    scheme = cfg["landuse_scheme"]
    codes = scheme["codes"]                       # code -> [bucket, desc]
    default_bucket = scheme.get("default_bucket", "Other / Unclassified")
    vacant_buckets = set(scheme.get("vacant_buckets", ["Vacant Land"]))
    public_re = re.compile(scheme["public_owner_regex"], re.I) if scheme.get("public_owner_regex") else None
    xwalk = cfg.get("zoning_crosswalk", {})       # juris -> {code: [category, threat]}
    zplain = cfg.get("zone_plain", {})

    parcels = json.loads((IN / "parcels_all.geojson").read_text())["features"]
    zoning = json.loads((IN / "zoning.geojson").read_text())["features"] if (IN / "zoning.geojson").exists() else []

    zgeoms, zprops = [], []
    for zf in zoning:
        try:
            g = _valid(shape(zf["geometry"]))
        except Exception:
            continue
        if g.is_empty:
            continue
        zgeoms.append(g)
        zprops.append(zf["properties"])
    tree = STRtree(zgeoms) if zgeoms else None

    unmapped_zone, unmapped_use = Counter(), Counter()
    bucket_count, cat_count, threat_count, code_count = Counter(), Counter(), Counter(), Counter()
    out = []
    for pf in parcels:
        p = pf["properties"]
        try:
            geom = _valid(shape(pf["geometry"]))
        except Exception:
            continue
        if geom.is_empty:
            continue
        rep = geom.representative_point()
        clat, clon = rep.y, rep.x
        ac = round(geo.acres(geom, to_local), 4)

        # ---- land-use bucket ----
        code = (p.get("landuse_code") or "").strip()
        code_count[code] += 1
        if p.get("data_confidence") not in (None, "full"):
            bucket = "Use code n/a (degraded source)"
        elif public_re and public_re.search(p.get("owner") or ""):
            bucket = "Public / Airport / Institutional"
        elif code in codes:
            bucket = codes[code][0]
        else:
            bucket = default_bucket
            if code:
                unmapped_use[code] += 1
        is_vacant = bucket in vacant_buckets

        # ---- zoning: containing polygon, else max-overlap >= 50% of the parcel ----
        juris, zcode, category, threat = None, None, GAP, "Unknown"
        if tree is not None:
            for idx in tree.query(rep):
                if zgeoms[idx].contains(rep):
                    juris, zcode = zprops[idx]["jurisdiction"], zprops[idx]["zone_code"]
                    break
            if juris is None:
                best_a, best, parea = 0.0, None, (geom.area or 1e-12)
                for idx in tree.query(geom):
                    inter = geom.intersection(zgeoms[idx]).area
                    if inter > best_a:
                        best_a, best = inter, idx
                if best is not None and best_a / parea >= 0.5:
                    juris, zcode = zprops[best]["jurisdiction"], zprops[best]["zone_code"]
        if juris:
            pair = xwalk.get(juris, {}).get((zcode or "").strip())
            if pair:
                category, threat = pair[0], pair[1]
            else:
                unmapped_zone[(juris, zcode)] += 1
                category, threat = "Planned / Overlay", "Unknown"
        else:
            juris = "(no public zoning)"

        zp = zplain.get((zcode or "").strip(), zcode or "—")
        # gov/airport-owned land: nominal residential zoning is meaningless for MF supply
        if bucket == "Public / Airport / Institutional":
            category, threat, zp = "Public / Airport / Institutional", "Low", "Public / exempt"

        dist = round(geo.haversine_mi(slat, slon, clat, clon), 3)
        bucket_count[bucket] += 1
        cat_count[category] += 1
        if is_vacant:
            threat_count[threat] += 1
        out.append({"type": "Feature", "geometry": pf["geometry"], "properties": {
            "account": p.get("account"), "owner": p.get("owner"), "situs": p.get("situs"),
            "landuse_code": code, "bucket": bucket, "landuse_color": P.landuse_color(bucket),
            "jurisdiction": juris, "zone_code": zcode, "zone_plain": zp,
            "zone_category": category, "zone_color": P.zone_color(category),
            "mf_threat": threat, "threat_color": P.threat_color(threat),
            "is_vacant": is_vacant, "acres": ac, "county": p.get("county"),
            "data_confidence": p.get("data_confidence"), "year_built": p.get("year_built"),
            "dist_mi": dist, "lat": round(clat, 6), "lon": round(clon, 6)}})

    (IN / "parcels_classified.geojson").write_text(json.dumps(
        {"type": "FeatureCollection", "features": out,
         "crs": {"type": "name", "properties": {"name": "EPSG:4326"}}}))

    # ---- land_use_buckets.csv ----
    with open(TBL / "land_use_buckets.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Bucket", "Source", "Code", "Description", "Color", "Parcels_in_area"])
        for cd, (bk, desc) in sorted(codes.items(), key=lambda kv: (kv[1][0], kv[0])):
            w.writerow([bk, scheme.get("name", "land-use code"), cd or "(blank)", desc,
                        P.landuse_color(bk), code_count.get(cd, 0)])
        if public_re:
            w.writerow(["Public / Airport / Institutional", "OWNER_RULE", "(public owner)",
                        "Government / airport / school / church-owned (overrides use code)",
                        P.landuse_color("Public / Airport / Institutional"),
                        bucket_count.get("Public / Airport / Institutional", 0)])

    # ---- zoning_categories.csv ----
    with open(TBL / "zoning_categories.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Category", "Jurisdiction", "Base Zone", "MF Threat (if vacant)", "Color"])
        for juris, table in xwalk.items():
            for cd, pair in sorted(table.items(), key=lambda kv: (kv[1][0], kv[0])):
                w.writerow([pair[0], juris, cd, pair[1], P.zone_color(pair[0])])
        w.writerow([GAP, "(unzoned / no public GIS)", "(none)", "n/a - not assessable",
                    P.zone_color(GAP)])

    # ---- mf_threat_classification.csv ----
    with open(TBL / "mf_threat_classification.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Threat Level", "Jurisdiction", "Base Zone", "Zone Category", "Definition"])
        order = {"High": 0, "Medium": 1, "Low": 2, "Unknown": 3}
        rows = [(order.get(pair[1], 9), juris, pair[1], cd, pair[0])
                for juris, table in xwalk.items() for cd, pair in table.items()]
        for _, juris, thr, cd, cat in sorted(rows):
            w.writerow([thr, juris, cd, cat, P.THREAT_DEF.get(thr, "")])

    # ---- unmapped log (decisions trail) ----
    (TBL / "unmapped_codes.json").write_text(json.dumps({
        "unmapped_zone_codes": [{"jurisdiction": j, "zone_code": z, "parcels": n}
                                for (j, z), n in unmapped_zone.most_common()],
        "unmapped_landuse_codes": [{"code": c, "parcels": n} for c, n in unmapped_use.most_common()],
    }, indent=2))

    print(f"Classified {len(out)} parcels.")
    print("  buckets:", dict(bucket_count))
    print("  vacant threat tiers:", dict(threat_count))
    print("  UNMAPPED zone codes:", dict(unmapped_zone))
    print("  UNMAPPED land-use codes:", dict(unmapped_use))


if __name__ == "__main__":
    main()
