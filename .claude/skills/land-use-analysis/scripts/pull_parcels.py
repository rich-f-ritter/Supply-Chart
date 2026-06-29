"""Pull all parcels in the analysis area from config.parcel_sources and write
in/parcels_all.geojson (normalized schema). Uses the robust POST + objectId-paging fetcher
(scripts/arcgis.py) — NOT the brittle GET+count pattern — so it works on big statewide/county
layers that disable count or reject long GET URLs. Resumable: id-batches are checkpointed
under in/_ck/, so re-running after an interruption continues where it left off.

Per-source config keys (see references/config-schema.md):
  name, county, url(.../query), fields{normalized:source}, dedupe_field, data_confidence,
  verify_ssl, where, and optional: page (id-batch size, default 1000), simplify (server-side
  maxAllowableOffset in degrees for dense metros; omit for exact acreage), format ("geojson"
  default | "esrijson" for old servers that don't emit geojson).

DO NOT hand-build a custom downloader — this handles POST, paging, and resume. If it's slow on
a dense area, just let it run (or re-run to resume); set "simplify": 0.00002 to thin geometry.
If a county truly has no ArcGIS layer, hand-write in/parcels_all.geojson in the normalized shape.
"""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import arcgis  # noqa: E402
import config as C  # noqa: E402
import geo  # noqa: E402


def _norm(ft, src):
    p = ft.get("properties", {}) or {}
    fm = src["fields"]

    def g(key):
        f = fm.get(key)
        v = p.get(f) if f else None
        return v.strip() if isinstance(v, str) else v

    return {"type": "Feature", "geometry": ft.get("geometry"), "properties": {
        "account": g("account"), "owner": g("owner"), "situs": g("situs"),
        "landuse_code": (g("landuse_code") or ""), "legal": g("legal"),
        "county": src.get("county"), "impr_value": g("impr_value"),
        "total_value": g("total_value"), "year_built": g("year_built"),
        "data_confidence": src.get("data_confidence", "full"), "source": src.get("name")}}


def main():
    cfg = C.load()
    IN = C.indir(cfg)
    ck = IN / "_ck"
    poly = geo.analysis_polygon(cfg)
    bbox = geo.bbox(cfg)

    feats, seen = [], set()
    for si, src in enumerate(cfg["parcel_sources"]):
        name = src.get("name", src["url"])
        label = "".join(ch if ch.isalnum() else "_" for ch in name)[:24] or f"parcels{si}"
        out_fields = ",".join(sorted({v for v in src["fields"].values() if v})) or "*"
        print(f"{name} ...", file=sys.stderr)
        try:
            raw = arcgis.fetch_all(
                src["url"], bbox, where=src.get("where", "1=1"), out_fields=out_fields,
                verify=src.get("verify_ssl", True), page=int(src.get("page", 1000)),
                simplify=src.get("simplify"), fmt=src.get("format", "geojson"),
                ck_dir=ck, label=label)
        except Exception as e:  # noqa: BLE001
            print(f"  [warn] {name} pull failed: {e}", file=sys.stderr)
            continue
        inpoly = arcgis.keep_in_poly(raw, poly)
        dd = src.get("dedupe_field")
        kept = 0
        for ft in inpoly:
            key = (name, (ft.get("properties", {}) or {}).get(dd)) if dd else id(ft)
            if dd and key in seen:
                continue
            if dd:
                seen.add(key)
            feats.append(_norm(ft, src))
            kept += 1
        print(f"  {name}: fetched {len(raw)} -> in-area {len(inpoly)} -> kept {kept}", file=sys.stderr)

    out = IN / "parcels_all.geojson"
    out.write_text(json.dumps({"type": "FeatureCollection", "features": feats,
                               "crs": {"type": "name", "properties": {"name": "EPSG:4326"}}}))
    print(f"Wrote {len(feats)} parcels -> {out}")
    if len(feats) > 40000:
        print("  [note] large parcel set — classify/build may be slow; consider a smaller radius "
              "or a 'simplify' offset on the source.")


if __name__ == "__main__":
    main()
