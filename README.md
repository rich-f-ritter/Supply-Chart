# Supply-Chart

A Claude Code **skill** that automates TMG's multifamily 5-mile new-construction
**Supply Chart**. It reconciles CoStar + RealPage radius exports into one
competitive-supply roster, buckets each property by lifecycle stage, builds a
relative-year (TTM) **Supply & Absorption** forecast (with a 3rd-party-forecast
block and model-linked subject rows), and is built to drop straight into the
underwriting model. **Every run also exports a companion map** (HTML + PNG).

## Use it in Claude Code on the web

**Option A — automatically, when you work on this repo (no setup).**
The skill lives at `.claude/skills/supply-chart/`, so any Claude Code **web
session started on the `Supply-Chart` repo loads it automatically**. Open a
session on this repo, upload a deal's CoStar + RealPage exports, and ask Claude to
"build a supply chart" — it runs the skill and returns the workbook **and** the map.

**Option B — account-wide, across all your sessions.**
Enable it as a Skill on claude.ai (**Customize → Skills/Capabilities**) by
uploading the skill folder. Per the docs, *"Skills you enable on claude.ai are
loaded into cloud sessions automatically."* (See the team note below for the
upload package.)

> The web doesn't support the `/plugin` command — that's CLI/desktop only — so the
> skill is distributed as a repo `.claude/skills/` folder, not a plugin.

## Layout
```
.claude/skills/supply-chart/
    SKILL.md                       Methodology + how Claude uses the skill
    scripts/build_supply_chart.py  Generator — also exports the map
    scripts/build_map.py           Companion map (folium HTML + matplotlib PNG)
    scripts/geo.py                 Shared geocoding (Nominatim, cached)
    scripts/embed_into_model.py    Graft the tabs into a model .xlsm
    reference/                     The template the output mirrors
    requirements.txt
examples/                          Worked-example exports (Aura, Bella, Canyon)
output/                            Sample generated charts + maps
```

## Dependencies
```bash
pip install -r .claude/skills/supply-chart/requirements.txt
```
The chart needs only `openpyxl`; the map adds `folium`/`matplotlib`/`contextily`
+ network for geocoding (it degrades gracefully without them).

## Run the scripts directly
```bash
python .claude/skills/supply-chart/scripts/build_supply_chart.py \
  --subject-name    "Canyon Ridge" \
  --subject-address "2552 E Gowen Rd, Boise, ID 83716" \
  --costar-roster    examples/canyon_ridge/CoStar_5mi_50unit_properties.xlsx \
  --costar-analytics examples/canyon_ridge/CoStar_5mi_Data_Analytics.xlsx \
  --realpage         examples/canyon_ridge/Realpage_5mi.xlsx \
  --out              output/Canyon_Ridge__Supply_Chart.xlsx
# writes the workbook AND Canyon_Ridge__Map.html + .png
```

## Carry it into the underwriting model
Subject rows are pre-linked by default — **drag** the `Competitive Analysis` +
`Supply & Absorption` tabs into the model, or **embed** them losslessly:
```bash
python .claude/skills/supply-chart/scripts/embed_into_model.py \
  --chart output/<deal>__Supply_Chart.xlsx \
  --model "TMG Acquisition Model.xlsm" \
  --out   "TMG Acquisition Model__with_Supply_Chart.xlsm"
```

See [SKILL.md](.claude/skills/supply-chart/SKILL.md) for the full methodology.
