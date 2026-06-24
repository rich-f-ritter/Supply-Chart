# Supply-Chart

A Claude Code **plugin** that automates TMG's multifamily 5-mile new-construction
**Supply Chart**. It reconciles CoStar + RealPage radius exports into one
competitive-supply roster, buckets each property by lifecycle stage, builds a
relative-year (TTM) **Supply & Absorption** forecast (with a 3rd-party-forecast
block and model-linked subject rows), and is built to drop straight into the
underwriting model. **Every run also exports a companion map** (HTML + PNG).

## Install it on your Claude account (web or desktop)
This repo is a **plugin marketplace**. In Claude Code, run:
```
/plugin marketplace add rich-f-ritter/supply-chart
/plugin install supply-chart@tmg
```
Then install the Python dependencies the scripts use (one-time, in the
environment Claude Code runs in):
```
pip install openpyxl folium zipcodes matplotlib contextily
```
The skill is then available across your sessions — Claude invokes it when you ask
to build a supply chart or hand it CoStar + RealPage radius exports. Skills from a
plugin are namespaced, so you can also call it explicitly with
`/supply-chart:supply-chart`. Get updates with `/plugin marketplace update`.

> The chart needs only `openpyxl`; the map adds `folium`/`matplotlib`/`contextily`
> + network for geocoding. Without them the chart still builds and the map is
> skipped with a note.

## Repo layout
```
.claude-plugin/marketplace.json                 Marketplace catalog ("tmg")
plugins/supply-chart/.claude-plugin/plugin.json Plugin manifest
plugins/supply-chart/skills/supply-chart/
    SKILL.md                                    Methodology + how Claude uses the skill
    scripts/build_supply_chart.py               Generator — also exports the map
    scripts/build_map.py                        Companion map (folium HTML + matplotlib PNG)
    scripts/geo.py                              Shared geocoding (Nominatim, cached)
    scripts/embed_into_model.py                 Graft the tabs into a model .xlsm
    reference/                                  The template the output mirrors
    requirements.txt
examples/                                       Worked-example exports (not bundled in the plugin)
output/                                         Sample generated charts + maps
```

## Run the scripts directly (without the plugin)
```bash
pip install -r plugins/supply-chart/skills/supply-chart/requirements.txt
python plugins/supply-chart/skills/supply-chart/scripts/build_supply_chart.py \
  --subject-name    "Canyon Ridge" \
  --subject-address "2552 E Gowen Rd, Boise, ID 83716" \
  --costar-roster    examples/canyon_ridge/CoStar_5mi_50unit_properties.xlsx \
  --costar-analytics examples/canyon_ridge/CoStar_5mi_Data_Analytics.xlsx \
  --realpage         examples/canyon_ridge/Realpage_5mi.xlsx \
  --out              output/Canyon_Ridge__Supply_Chart.xlsx
# writes the workbook AND Canyon_Ridge__Map.html + .png
```

## Carrying it into the underwriting model
Subject rows are pre-linked by default — **drag** the `Competitive Analysis` +
`Supply & Absorption` tabs into the model (links resolve to `Cash Flow (Annual)` /
`Rent & Occ Data`), or **embed** them losslessly:
```bash
python plugins/supply-chart/skills/supply-chart/scripts/embed_into_model.py \
  --chart output/<deal>__Supply_Chart.xlsx \
  --model "TMG Acquisition Model.xlsm" \
  --out   "TMG Acquisition Model__with_Supply_Chart.xlsm"
```

See [SKILL.md](plugins/supply-chart/skills/supply-chart/SKILL.md) for the full
methodology, reconciliation rules, and the model-incorporation mapping.
