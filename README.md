# Supply-Chart

A Claude Code **skill** that automates the multifamily 5-mile new-construction
**Supply Chart** used in underwriting. It reconciles CoStar and RealPage radius
exports into one competitive-supply roster, buckets each property by lifecycle
stage (stabilized / leasing-up / under-construction / proposed), and writes a
formatted workbook with a **Supply & Absorption** forecast tab — a relative-year
(TTM) projection of new supply, absorption, and overall occupancy with editable
demand scenarios, plus subject rent/occupancy rows.

This is the **rent-analysis** layer: it shows when market occupancy re-stabilizes
enough to push rents. Output feeds the underwriting model.

## Quick start
```bash
pip install -r requirements.txt   # openpyxl

# Minimal (3 market exports):
python scripts/build_supply_chart.py \
  --subject-name    "Canyon Ridge" \
  --subject-address "2552 E Gowen Rd" \
  --costar-roster    examples/canyon_ridge/CoStar_5mi_50unit_properties.xlsx \
  --costar-analytics examples/canyon_ridge/CoStar_5mi_Data_Analytics.xlsx \
  --realpage         examples/canyon_ridge/Realpage_5mi.xlsx \
  --out              output/Canyon_Ridge__Supply_Chart.xlsx

# Full (adds subject rows from the RR-T12 intake + CoStar per-property rents):
#   --intake examples/bella_mirage/Bella_Underwriting_Intake.xlsx \
#   --costar-subject-rents examples/bella_mirage/CoStar_5mi_property_rents.xlsx
```

## Layout
```
SKILL.md                         Methodology + when/how Claude uses this skill
scripts/build_supply_chart.py    The generator (parsing, reconciliation, writer)
reference/                       The template this output mirrors
examples/canyon_ridge/           Worked example — small market (Boise, ID)
examples/bella_mirage/           Worked example — large supply-heavy market (Phoenix)
output/                          Generated supply charts
```

See [SKILL.md](SKILL.md) for the full methodology, reconciliation rules, and the
analyst follow-ups (proximity, HelloData lease-up verification, pipeline timing)
the script intentionally leaves open.
