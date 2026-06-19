# Supply-Chart

A Claude Code **skill** that automates the multifamily 5-mile new-construction
**Supply Chart** used in underwriting. It reconciles CoStar and RealPage
radius exports into one competitive-supply roster, pins delivery quarters against
the CoStar Data Analytics deliveries series, buckets each property by lifecycle
stage, and writes a formatted workbook with a quarterly absorption summary.

The supply layer feeds the **rent-analysis** step of the model (layering supply
on top of historical occupancy / absorption / deliveries to forecast demand and
identify when occupancy is strong enough to push rents).

## Quick start
```bash
pip install openpyxl

python scripts/build_supply_chart.py \
  --subject-name    "Canyon Ridge" \
  --subject-address "2552 E Gowen Rd" \
  --costar-roster    examples/canyon_ridge/CoStar_5mi_50unit_properties.xlsx \
  --costar-analytics examples/canyon_ridge/CoStar_5mi_Data_Analytics.xlsx \
  --realpage         examples/canyon_ridge/Realpage_5mi.xlsx \
  --out              output/Canyon_Ridge__Supply_Chart.xlsx
```

## Layout
```
SKILL.md                         Methodology + when/how Claude uses this skill
scripts/build_supply_chart.py    The generator (parsing, reconciliation, writer)
reference/                       The template this output mirrors
examples/canyon_ridge/           Real CoStar + RealPage exports (worked example)
output/                          Generated supply charts
```

See [SKILL.md](SKILL.md) for the full methodology, reconciliation rules, and the
analyst follow-ups (proximity, HelloData lease-up verification, pipeline timing)
the script intentionally leaves open.
