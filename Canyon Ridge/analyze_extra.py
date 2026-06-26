"""Supplementary analysis for the Canyon Ridge 5-mi land-use study:
- subject's Airport Influence Area (AI-O) zone
- acreage of land within the 5-mi radius by AI-O zone
- tag each reasoned vacant candidate with its AI-O zone
- Micron / Columbia Bench large-parcel summary
- relevant Development Tracker records near the subject
Outputs CSVs into Tables/ and a JSON summary into in/extra/.
"""
import json, csv, math
from pathlib import Path
from shapely.geometry import shape, Point
from shapely.strtree import STRtree
from shapely.prepared import prep
from pyproj import Transformer

ROOT = Path(".")
IN = ROOT / "in"
TBL = ROOT / "Tables"; TBL.mkdir(exist_ok=True)
EX = IN / "extra"

SLAT, SLON = 43.541734, -116.151198
RADIUS_MI = 5.0

# local equal-area-ish transform (UTM 11N, meters) for acreage
to_m = Transformer.from_crs("EPSG:4326", "EPSG:32611", always_xy=True).transform
def to_local(geom):
    from shapely.ops import transform
    return transform(to_m, geom)
def acres(geom):
    return to_local(geom).area / 4046.8564224

def haversine_mi(la1, lo1, la2, lo2):
    R = 3958.7613
    p1, p2 = math.radians(la1), math.radians(la2)
    dp = math.radians(la2 - la1); dl = math.radians(lo2 - lo1)
    a = math.sin(dp/2)**2 + math.cos(p1)*math.cos(p2)*math.sin(dl/2)**2
    return 2*R*math.asin(math.sqrt(a))

# subject point + 5mi circle (buffer in meters)
subj = Point(SLON, SLAT)
subj_m = to_local(subj)
circle_m = subj_m.buffer(RADIUS_MI * 1609.344)
# bring circle back to lonlat for clipping
from shapely.ops import transform as _tf
from pyproj import Transformer as _T
to_ll = _T.from_crs("EPSG:32611", "EPSG:4326", always_xy=True).transform
circle_ll = _tf(to_ll, circle_m)
circle_prep = prep(circle_ll)

# ---------- Airport Influence Areas ----------
aia = json.load(open(EX / "airport_influence.geojson"))["features"]
aia_zone_geoms = {}   # zone -> list of geoms
for f in aia:
    z = (f["properties"].get("ZONE") or "?").strip()
    try:
        g = shape(f["geometry"])
        if not g.is_valid: g = g.buffer(0)
    except Exception:
        continue
    aia_zone_geoms.setdefault(z, []).append(g)

# subject zone
subj_zone = "Outside AIA"
for z, geoms in aia_zone_geoms.items():
    if any(g.contains(subj) for g in geoms):
        subj_zone = z; break

# acreage by zone within 5-mi circle
zone_acres = {}
for z, geoms in aia_zone_geoms.items():
    a = 0.0
    for g in geoms:
        inter = g.intersection(circle_ll)
        if not inter.is_empty:
            a += acres(inter)
    zone_acres[z] = round(a, 1)
circle_acres = round(acres(circle_ll), 1)

# build flat AIA tree for point tagging
aia_flat_geoms, aia_flat_zone = [], []
for z, geoms in aia_zone_geoms.items():
    for g in geoms:
        aia_flat_geoms.append(g); aia_flat_zone.append(z)
aia_tree = STRtree(aia_flat_geoms) if aia_flat_geoms else None
def zone_at(lon, lat):
    if not aia_tree: return "Outside AIA"
    p = Point(lon, lat)
    for idx in aia_tree.query(p):
        if aia_flat_geoms[idx].contains(p):
            return aia_flat_zone[idx]
    return "Outside AIA"

# ---------- tag vacant candidates with AIA zone ----------
cands = json.load(open(IN / "vacant_candidates.json"))
for r in cands:
    r["aia_zone"] = zone_at(r["lon"], r["lat"])
json.dump(cands, open(IN / "vacant_candidates.json", "w"), indent=2)

# ---------- Micron / Columbia Bench large parcels ----------
# large vacant/ag parcels on the Columbia Bench (east of Federal Way, around Columbia Rd)
parcels = json.load(open(IN / "parcels_classified.geojson"))["features"]
# classify residential MF supply: existing apartments can't be told from PROPCODE,
# so we summarize buckets and zoning instead.
bucket_acres = {}
zone_threat_acres = {"High":0.0,"Medium":0.0,"Low":0.0,"Unknown":0.0}
for f in parcels:
    p = f["properties"]
    a = p.get("acres") or 0
    bucket_acres[p["bucket"]] = bucket_acres.get(p["bucket"],0)+a
    if p.get("is_vacant"):
        t = p.get("mf_threat","Unknown")
        zone_threat_acres[t] = zone_threat_acres.get(t,0)+a

# big tracts near Micron campus (within ~2mi of 8000 S Federal Way ~ 43.522,-116.137)
MLAT, MLON = 43.5225, -116.1370
micron_tracts = []
for f in parcels:
    p = f["properties"]
    a = p.get("acres") or 0
    if a < 20: continue
    d = haversine_mi(MLAT, MLON, p["lat"], p["lon"])
    if d <= 2.2 and p["bucket"] in ("Vacant Land","Agricultural / Rural","Commercial"):
        micron_tracts.append((round(a,1), p.get("situs"), p.get("zone_code"), p.get("mf_threat"),
                              zone_at(p["lon"],p["lat"]), round(d,2), p.get("account")))
micron_tracts.sort(reverse=True)

# ---------- Development tracker: relevant nearby ----------
dev = json.load(open(EX / "dev_tracker.geojson"))["features"]
dev_rows = []
for f in dev:
    p = f["properties"]
    try:
        g = shape(f["geometry"]); c = g.representative_point()
    except Exception:
        continue
    d = haversine_mi(SLAT, SLON, c.y, c.x)
    if d > RADIUS_MI: continue
    rt = (p.get("RecordType") or "")
    # keep substantive land-use actions
    if rt in ("Zoning Certificate","Variance","Sign","Certificate of Zoning Compliance"):
        continue
    dev_rows.append((round(d,2), p.get("RecordType"), p.get("Status"), p.get("RecordName"),
                     p.get("ComprehensivePlanningArea"), p.get("ReviewAuthority"),
                     p.get("ZoningCode"), zone_at(c.x,c.y), p.get("RecordID")))
dev_rows.sort()

# ---------- write tables ----------
with open(TBL/"airport_influence_area_acres.csv","w",newline="") as f:
    w=csv.writer(f); w.writerow(["AIA_Zone","Acres_within_5mi","Residential_rule"])
    rule={"A":"New residential allowed w/ 25 dB sound insulation + avigation easement",
          "B":"New residential PROHIBITED (most restrictive; schools/daycare/worship also prohibited)",
          "B-1":"Residential allowed but capped 5 units/acre (CUP for more) + 30 dB insulation + easement",
          "C":"New residential PROHIBITED (existing = legal nonconforming); industrial/commercial w/ insulation"}
    for z in ["A","B","B-1","C"]:
        if z in zone_acres: w.writerow([z, zone_acres[z], rule.get(z,"")])
    w.writerow(["(5-mi circle total land+water)", circle_acres, ""])

with open(TBL/"micron_area_large_tracts.csv","w",newline="") as f:
    w=csv.writer(f); w.writerow(["acres","situs","zone_code","mf_threat","aia_zone","mi_from_Micron","parcel"])
    for r in micron_tracts: w.writerow(r)

with open(TBL/"development_tracker_5mi.csv","w",newline="") as f:
    w=csv.writer(f); w.writerow(["mi_from_subject","record_type","status","name","planning_area","review_authority","zoning","aia_zone","record_id"])
    for r in dev_rows: w.writerow(r)

summary={"subject_aia_zone":subj_zone,"circle_acres":circle_acres,"aia_zone_acres":zone_acres,
         "vacant_threat_acres":{k:round(v,1) for k,v in zone_threat_acres.items()},
         "bucket_acres":{k:round(v,1) for k,v in bucket_acres.items()},
         "micron_area_tract_count":len(micron_tracts),
         "micron_area_tract_acres":round(sum(r[0] for r in micron_tracts),1),
         "dev_tracker_count_5mi":len(dev_rows)}
json.dump(summary, open(EX/"summary.json","w"), indent=2)
print(json.dumps(summary, indent=2))
print("\nMicron-area large tracts (top 12):")
for r in micron_tracts[:12]: print("  ",r)
print(f"\nAIA zone counts among {len(cands)} vacant candidates:")
from collections import Counter
print("  ", dict(Counter(r["aia_zone"] for r in cands)))
