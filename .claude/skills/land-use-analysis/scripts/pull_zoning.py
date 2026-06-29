"""Pull zoning polygons for every jurisdiction in config.zoning_sources and normalize to
{jurisdiction, zone_code, zone_desc}. Uses the robust POST + objectId-paging fetcher
(scripts/arcgis.py) so it won't silently truncate at the server's record cap. Clipped to the
analysis bbox. Areas with no source are handled downstream as a labeled data-gap class — never
faked. Output: in/zoning.geojson.

Per-source config keys: jurisdiction, url(.../query), code_field, desc_field (optional),
verify_ssl, where, and optional page / simplify / format (as in pull_parcels).
"""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import arcgis  # noqa: E402
import config as C  # noqa: E402
import geo  # noqa: E402


def main():
    cfg = C.load()
    IN = C.indir(cfg)
    ck = IN / "_ck"
    bbox = geo.bbox(cfg)

    feats = []
    for si, src in enumerate(cfg.get("zoning_sources", [])):
        juris = src["jurisdiction"]
        code_f = src["code_field"]
        desc_f = src.get("desc_field") or code_f
        label = "zon_" + "".join(ch if ch.isalnum() else "_" for ch in juris)[:20]
        print(f"{juris} zoning ...", file=sys.stderr)
        try:
            raw = arcgis.fetch_all(
                src["url"], bbox, where=src.get("where", "1=1"),
                out_fields=",".join(sorted({code_f, desc_f})),
                verify=src.get("verify_ssl", True), page=int(src.get("page", 1000)),
                simplify=src.get("simplify"), fmt=src.get("format", "geojson"),
                ck_dir=ck, label=label)
        except Exception as e:  # noqa: BLE001
            print(f"  [warn] {juris} zoning failed: {e}", file=sys.stderr)
            continue
        for ft in raw:
            p = ft.get("properties", {}) or {}
            feats.append({"type": "Feature", "geometry": ft.get("geometry"),
                          "properties": {"jurisdiction": juris,
                                         "zone_code": str(p.get(code_f) or "").strip(),
                                         "zone_desc": str(p.get(desc_f) or "").strip()}})
        print(f"  {juris}: {len(raw)} polys", file=sys.stderr)

    out = IN / "zoning.geojson"
    out.write_text(json.dumps({"type": "FeatureCollection", "features": feats,
                               "crs": {"type": "name", "properties": {"name": "EPSG:4326"}}}))
    print(f"Wrote {len(feats)} zoning polygons "
          f"({len(cfg.get('zoning_sources', []))} sources) -> {out}")


if __name__ == "__main__":
    main()
