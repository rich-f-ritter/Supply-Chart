"""Turn Claude's REASONED ranking (in/reasoned_ranking.json) into the final
Tables/top10_concerning_vacant.csv. This step is deliberately dumb: it does not decide
threat or rank - it joins each reasoned entry back to its prepared candidate cluster and
emits the table. The judgment lives in reasoned_ranking.json (see threat-reasoning.md).

in/reasoned_ranking.json schema:
{
  "headline": "one-paragraph supply-threat conclusion",
  "ranking": [
    { "match": {"account": "12345"}  OR  {"owner_prefix": "...", "location_prefix": "..."},
      "reasoned_threat": "High|Medium|Low|Low-Watch|Unknown",
      "rationale": "why this parcel is (or isn't) a concerning MF-supply site" },
    ...
  ]
}
Threat strings are free text (so "Low-Watch" etc. are allowed); the workbook color-keys
on the leading word (High/Medium/Low/Unknown).
"""
import csv
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import config as C  # noqa: E402


def zoning_plain(zstr, zplain):
    out, seen = [], set()
    for tok in (zstr or "").split(","):
        tok = tok.strip()
        if ":" in tok:
            juris, code = tok.split(":", 1)
            label = f"{zplain.get(code.strip(), code.strip())} ({juris})"
            if label not in seen:
                seen.add(label)
                out.append(label)
    return "; ".join(out)


def main():
    cfg = C.load()
    IN, TBL = C.indir(cfg), C.tables(cfg)
    zplain = cfg.get("zone_plain", {})
    cands = json.loads((IN / "vacant_candidates.json").read_text())
    rk_path = IN / "reasoned_ranking.json"
    if not rk_path.exists():
        sys.exit(f"[finalize] missing {rk_path}\nReason over in/vacant_candidates.json and "
                 "write in/reasoned_ranking.json first (see references/threat-reasoning.md).")
    rk = json.loads(rk_path.read_text())

    def match(spec):
        acct = spec.get("account")
        if acct:
            for c in cands:
                if str(acct) in (c.get("parcels") or "").split(","):
                    return c
        op = (spec.get("owner_prefix") or "").upper()
        lp = (spec.get("location_prefix") or "").upper()
        for c in cands:
            if op and not (c.get("owner") or "").upper().startswith(op):
                continue
            if lp and lp not in (c.get("location") or "").upper():
                continue
            if op or lp:
                return c
        return None

    rows = []
    for rank, item in enumerate(rk.get("ranking", []), 1):
        c = match(item.get("match", {}))
        if not c:
            print(f"  [warn] no candidate matched ranking #{rank}: {item.get('match')}", file=sys.stderr)
            continue
        rows.append([rank, item.get("reasoned_threat", ""), zoning_plain(c["zoning"], zplain),
                     c["zoning"], c["parcel_count"], c["acres"], c["min_dist_mi"], c["lat"],
                     c["lon"], c["location"], c["owner"], c["owner_type"], c["parcels"],
                     item.get("rationale", "")])

    with open(TBL / "top10_concerning_vacant.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["rank", "reasoned_threat", "zoning_plain", "zoning_code", "parcel_count",
                    "acres", "dist_mi", "lat", "lon", "LOCATION", "OWNER1", "owner_type",
                    "PARCEL", "rationale"])
        w.writerows(rows)

    if rk.get("headline"):
        (TBL / "headline.txt").write_text(rk["headline"], encoding="utf-8")
    print(f"wrote reasoned top-{len(rows)} -> Tables/top10_concerning_vacant.csv")
    for r in rows:
        print(f"  {r[0]:2} {str(r[1]):11} ac={r[5]:7.2f} d={r[6]:.2f} | {r[2][:24]:24} | {str(r[9])[:26]}")


if __name__ == "__main__":
    main()
