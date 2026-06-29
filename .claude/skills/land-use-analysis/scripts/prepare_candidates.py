"""Prepare enriched vacant-parcel candidates for REASONED threat assessment.

Deterministic ONLY for mechanical work: drop non-developable parcels (roads/slivers by
Polsby-Popper compactness, HOA/POA/condo common areas, churches/schools, gov/airport via
owner rule, data-gap zoning, sub-minimum acreage), cluster adjacent commonly-owned
parcels, and gather facts + owner-type flags. The FINAL threat and ranking are NOT decided
here - they are reasoned over these candidates (see finalize_topN.py + threat-reasoning.md).
The `developable` set written here also drives which vacant parcels the THREAT MAP shows.
"""
import csv
import json
import math
import re
import sys
from collections import defaultdict
from pathlib import Path

from shapely.geometry import shape
from shapely.ops import transform
from shapely.strtree import STRtree

sys.path.insert(0, str(Path(__file__).resolve().parent))
import config as C  # noqa: E402
import geo  # noqa: E402


def main():
    cfg = C.load()
    IN = C.indir(cfg)
    to_local = geo.local_transform(cfg)
    min_acres = float(cfg.get("min_acres", 1.0))
    min_compact = float(cfg.get("min_compactness", 0.16))
    nondev = re.compile(cfg["nondev_owner_regex"], re.I) if cfg.get("nondev_owner_regex") else None
    company = re.compile(cfg["company_owner_regex"], re.I) if cfg.get("company_owner_regex") else None
    gap = cfg.get("data_gap_label", "No public zoning (data gap)")

    def norm_owner(o):
        return re.sub(r"[^A-Z0-9 ]", "", (o or "").upper().strip())[:18]

    def owner_type(o):
        if nondev and nondev.search(o or ""):
            return "non-developer (gov/HOA/inst.)"
        if company and company.search(o or ""):
            return "company/investor"
        return "individual"

    feats = json.loads((IN / "parcels_classified.geojson").read_text())["features"]
    cand, dropped = [], defaultdict(int)
    for f in feats:
        p = f["properties"]
        if not p["is_vacant"]:
            continue
        if p["zone_category"] == gap:
            dropped["no-zoning"] += 1
            continue
        if p["acres"] < min_acres:
            dropped["too-small"] += 1
            continue
        if nondev and nondev.search(p["owner"] or ""):
            dropped["non-developer-owner"] += 1
            continue
        try:
            g = shape(f["geometry"])
            gp = transform(to_local, g)
        except Exception:
            continue
        if not gp.is_valid:
            gp = gp.buffer(0)
        compact = 4 * math.pi * gp.area / (gp.length ** 2) if gp.length else 0
        if compact < min_compact:
            dropped["road/sliver-shape"] += 1
            continue
        p["_compact"] = round(compact, 3)
        cand.append((g, p))

    geoms = [g for g, _ in cand]
    tree = STRtree(geoms) if geoms else None
    parent = list(range(len(cand)))

    def find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    for i, (g, p) in enumerate(cand):
        ok, gb = norm_owner(p["owner"]), g.buffer(0.00025)
        for j in tree.query(gb):
            if j > i and norm_owner(cand[j][1]["owner"]) == ok and gb.intersects(geoms[j]):
                parent[find(i)] = find(j)

    clusters = defaultdict(list)
    for i in range(len(cand)):
        clusters[find(i)].append(i)

    rows = []
    for members in clusters.values():
        props = [cand[i][1] for i in members]
        rep = max(props, key=lambda x: x["acres"])
        rows.append({
            "acres": round(sum(x["acres"] for x in props), 2),
            "parcel_count": len(props),
            "min_dist_mi": round(min(x["dist_mi"] for x in props), 2),
            "owner": rep["owner"], "owner_type": owner_type(rep["owner"]),
            "zoning": ", ".join(sorted({f"{x['jurisdiction']}:{x['zone_code']}" for x in props})),
            "zoning_threat_baseline": rep["mf_threat"], "compactness": rep["_compact"],
            "location": rep["situs"], "lat": rep["lat"], "lon": rep["lon"],
            "parcels": ",".join(str(x["account"]) for x in props if x["account"]),
        })
    tierrank = {"High": 0, "Medium": 1, "Unknown": 2, "Low": 3}
    rows.sort(key=lambda r: (tierrank.get(r["zoning_threat_baseline"], 4),
                             -r["acres"], r["min_dist_mi"]))

    dev_accts = [p["account"] for _, p in cand if p.get("account")]
    (IN / "developable_accounts.json").write_text(json.dumps(dev_accts))
    (IN / "vacant_candidates.json").write_text(json.dumps(rows, indent=2))
    if rows:
        with open(IN / "vacant_candidates.csv", "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
            w.writeheader()
            w.writerows(rows)

    print(f"candidates after filters: {len(cand)} parcels -> {len(rows)} clusters")
    print("  dropped:", dict(dropped))
    print("  -> now REASON over in/vacant_candidates.json; write in/reasoned_ranking.json")
    for r in rows[:30]:
        print(f"  base={r['zoning_threat_baseline']:7} ac={r['acres']:7.2f} n={r['parcel_count']:2} "
              f"d={r['min_dist_mi']:.2f} pp={r['compactness']:.2f} {r['owner_type'][:14]:14} "
              f"| {r['zoning'][:20]:20} | {str(r['owner'])[:26]:26} | {str(r['location'])[:20]}")


if __name__ == "__main__":
    main()
