# Supply-Chart

A Claude Code **skill** that automates the multifamily 5-mile new-construction
**Supply Chart** used in underwriting. It reconciles CoStar and RealPage radius
exports into one competitive-supply roster, buckets each property by lifecycle
stage (stabilized / leasing-up / under-construction / proposed), and writes a
formatted workbook with a **Supply & Absorption** forecast tab — a relative-year
(TTM) projection of new supply, absorption, and overall occupancy with editable
demand scenarios, subject rent/occupancy rows, and a 3rd-party-forecast block.
**Every run also exports a companion map** (interactive HTML + static PNG).

This is the **rent-analysis** layer: it shows when market occupancy re-stabilizes
enough to push rents. Output is built to drop straight into the underwriting model.

## Quick start
```bash
pip install -r requirements.txt   # openpyxl (chart) + folium/zipcodes/matplotlib/contextily (map)

python scripts/build_supply_chart.py \
  --subject-name    "Canyon Ridge" \
  --subject-address "2552 E Gowen Rd, Boise, ID 83716" \
  --costar-roster    examples/canyon_ridge/CoStar_5mi_50unit_properties.xlsx \
  --costar-analytics examples/canyon_ridge/CoStar_5mi_Data_Analytics.xlsx \
  --realpage         examples/canyon_ridge/Realpage_5mi.xlsx \
  --out              output/Canyon_Ridge__Supply_Chart.xlsx
# writes Canyon_Ridge__Supply_Chart.xlsx AND Canyon_Ridge__Map.html + .png

# Full (adds subject rows from the RR-T12 intake + per-property rents):
#   --intake examples/bella_mirage/Bella_Underwriting_Intake.xlsx \
#   --costar-subject-rents examples/bella_mirage/CoStar_5mi_property_rents.xlsx
```
The companion map needs `folium`/`matplotlib`/`contextily` and network for
geocoding; without them the workbook still writes and the map is skipped with a
note. Add `--no-map` to skip it, or `--no-model-link` for a chart with no model
references.

## Carrying it into the underwriting model
The subject rows are pre-linked to the model by default. Either **drag** the
`Competitive Analysis` + `Supply & Absorption` tabs into the model (the links
resolve to its own `Cash Flow (Annual)` / `Rent & Occ Data`), **or** embed them:
```bash
python scripts/embed_into_model.py \
  --chart output/<deal>__Supply_Chart.xlsx \
  --model "TMG Acquisition Model.xlsm" \
  --out   "TMG Acquisition Model__with_Supply_Chart.xlsm"
```
`embed_into_model.py` grafts the tabs at the package level so VBA, charts, and
formatting are preserved byte-for-byte (it does not round-trip the `.xlsm`).

## Installing as a team skill
This folder *is* the skill (SKILL.md carries the `name`/`description` frontmatter).
To make it available in Claude Code, drop the folder at
`~/.claude/skills/supply-chart/` (personal) or `.claude/skills/supply-chart/` in a
shared repo, then `pip install -r requirements.txt`. Claude invokes it when asked
to build/update a supply chart or given CoStar + RealPage radius exports.

## Layout
```
SKILL.md                         Methodology + when/how Claude uses this skill
scripts/build_supply_chart.py    The generator (parsing, reconciliation, writer) — also exports the map
scripts/build_map.py             Companion map (folium HTML + matplotlib PNG)
scripts/geo.py                   Shared geocoding (Nominatim, cached, ZIP fallback)
scripts/embed_into_model.py      Graft the tabs into a model .xlsm, preserving everything
reference/                       The template this output mirrors
examples/aura_beacon_island/     Worked example — Houston (full intake + diligence)
examples/bella_mirage/           Worked example — large supply-heavy market (Phoenix)
examples/canyon_ridge/           Worked example — small market (Boise, ID)
output/                          Sample generated charts + maps
```

See [SKILL.md](SKILL.md) for the full methodology, reconciliation rules, the
model-incorporation mapping, and the analyst follow-ups the script leaves open.
