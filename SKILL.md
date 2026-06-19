---
name: supply-chart
description: >-
  Build a 5-mile new-construction "Supply Chart" for a multifamily underwriting
  deal. Reconciles CoStar and RealPage 5-mile-radius exports (plus the CoStar
  Data Analytics time series) into one competitive-supply roster, pins delivery
  quarters, buckets each property by lifecycle stage, and writes a formatted
  workbook with a quarterly absorption summary. Use when the user asks to build,
  update, or automate a supply chart / competitive supply analysis, or provides
  CoStar + RealPage radius exports for a deal. Output feeds the rent-analysis tab
  of the underwriting model.
---

# Supply Chart

## What this produces
A workbook (mirroring `reference/EXAMPLE_Seasons_Supply_Chart_v2.xlsx`) with:

1. **Competitive Analysis** — the new-construction roster grouped into four
   lifecycle buckets, then a per-quarter absorption summary
   (# props, total units, weighted-avg occupancy, currently occupied,
   target @ goal, units needed to reach 95% / 92.5% / 90%), a TOTAL row, the
   total current 5-mile inventory, and **% to be absorbed**.
2. **Settings** — the stabilization occupancy target (default 95%) that the
   absorption formulas reference.
3. **Reconciliation Log** — every property with its merged values, sources, and
   any conflict notes, so the analyst can audit the automated decisions.

This supply layer is the input to the **rent-analysis** step: it sits on top of
the historical overall occupancy / absorption / deliveries series (CoStar Data
Analytics) so demand can be forecast and the "when is occupancy strong enough to
push rents" question answered. (Rent analysis is a *separate, later* step — this
skill only builds the supply chart.)

## Inputs (three exports, all 5-mile radius around the subject)
| Input | Provides | Notes |
|---|---|---|
| **CoStar 50-unit property roster** (`Export…xlsx`) | Property-level roster of existing assets: name, address, year built, **construction begin**, units, bed mix, stories, owner. | Used as the **primary roster + unit counts + delivery timing basis**. Misses sub-50-unit and not-yet-tracked pipeline deals. |
| **CoStar Data Analytics** (`MultifamilyDataGrid`) | Quarterly submarket time series: inventory, rent, occupancy, absorption, under-construction, **deliveries by quarter**. | Source of **Total Current Inventory** and of the **delivery-quarter pinning** (match a property's unit count to the quarter it was delivered). |
| **RealPage 5-mile** | Competitor roster with **per-property occupancy & effective rent**, plus the **forward pipeline** (`Property Status` = Pre-Planned / Under Construction). | Fills occupancy/rent and catches deals CoStar hasn't picked up (sub-50u, pipeline). |

> The current CoStar property export does **not** include per-property rent or
> occupancy. When the user re-pulls CoStar with those fields, the script will
> cross-check them against RealPage automatically (it already backfills occ/rent
> only when missing, and notes unit/year conflicts).

## How to run
```bash
python scripts/build_supply_chart.py \
  --subject-name    "Canyon Ridge" \
  --subject-address "2552 E Gowen Rd" \
  --costar-roster   path/to/CoStar_5mi_50unit_properties.xlsx \
  --costar-analytics path/to/CoStar_5mi_Data_Analytics.xlsx \
  --realpage        path/to/Realpage_5mi.xlsx \
  --out             output/<Deal>__Supply_Chart.xlsx
```
Optional: `--as-of "2026 Q2"` (defaults to the latest quarter in the analytics
file) and `--target 0.95` (stabilization occupancy goal).

The script prints a reconciliation report to the console and writes the workbook.
Always read the console report and the **Reconciliation Log** sheet, then review
the workbook before handing it off.

## Methodology

### 1. Reconcile the two rosters
- Match properties across CoStar and RealPage by **street address** (leading
  number + first significant street word; handles ranges and directionals), with
  a **name fallback** for missing/typo'd addresses.
- **Unit counts:** keep CoStar's; flag a note when the two sources differ by ≥3
  units (±2 is treated as noise).
- **Occupancy & effective rent:** taken from RealPage today (CoStar lacks them);
  if a future CoStar pull has them, only missing values are backfilled and
  conflicts are noted.
- **Status:** a forward-looking status (Pre-Planned / Under Construction) from
  either source wins, so pipeline deals are never mislabeled as existing.
- The **subject property** is dropped from the competitive roster.

### 2. Pin the delivery quarter (the key cross-comparison)
CoStar's roster only gives a *year built*. To get the quarter, match each
delivered property's **unit count to the CoStar Data Analytics quarterly
deliveries series** within ±1 year of its year built — an exact unit match wins,
else the closest. (e.g. Timbers 274u → Q2 2023; The Mill 125u → Q2 2024;
The Betty 46u → Q4 2023.) If no match, fall back to `Q? <year>`. Pipeline deals
with no real date get `TBD` and are listed but excluded from the dated absorption
table.

### 3. Bucket each property
- **PROPOSED** — RealPage Pre-Planned / proposed (no real delivery date, no rent/occ).
- **UNDER CONSTRUCTION** — construction begun, not yet delivered (no rent/occ).
- **LEASING UP** — delivered within the last ~6 quarters **and** occupancy below
  the stabilized threshold (90%). **These are flagged "Verify rent/occ w/
  HelloData"** — lease-ups are where accurate rent & occupancy matter most and
  where a HelloData pull is worth it.
- **STABILIZED / STABILIZING** — delivered and either ≥90% occupied or open >6
  quarters.

Only genuine new supply is charted: deliveries within the last
`NEW_CONSTRUCTION_LOOKBACK_YEARS` (default 4) plus the entire pipeline.

### 4. Absorption summary
For every quarter spanned by the **dated** deliveries, the workbook computes (via
live Excel formulas, so they stay correct if the analyst edits a cell):
`# props`, `total units`, `weighted-avg occupancy`, `currently occupied`,
`target @ goal`, and `units needed to reach 95% / 92.5% / 90%`; then a TOTAL and
**% to be absorbed** against the total current inventory. A near-zero or negative
"% to be absorbed" means little competitive lease-up overhang — supportive of
pushing rents.

## Analyst follow-ups the script intentionally leaves open
- **Proximity (miles)** — left blank; neither export carries distance-from-subject. Fill manually (or paste CoStar's "Distance" column if a future pull includes it).
- **Lease-up rent & occupancy** — verify the flagged lease-ups with a HelloData pull.
- **Pipeline timing** — Pre-Planned deals are undated (`TBD`); add an estimated delivery quarter if CoStar/RealPage provides one so they flow into the absorption table.

## Tunable thresholds
Top of `scripts/build_supply_chart.py`: `DEFAULT_STABILIZATION_TARGET`,
`STABILIZED_OCC`, `LEASEUP_WINDOW_QTRS`, `NEW_CONSTRUCTION_LOOKBACK_YEARS`.

## Worked example
`examples/canyon_ridge/` holds the three real exports for Canyon Ridge (Boise,
ID). Running the command above reproduces
`output/Canyon_Ridge__Supply_Chart.xlsx`: 3 stabilized comps (445 u, ≈fully
absorbed) and a 1,763-unit proposed pipeline — i.e. minimal near-term
competitive supply.
