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
| **RealPage 5-mile** | Competitor roster with **per-property occupancy & effective rent**, plus the **forward pipeline** (`Property Status` = Pre-Planned / Under Construction). | Catches deals CoStar hasn't picked up (sub-50u, pipeline). |

> **CoStar rent/occupancy:** if the CoStar export includes `Avg Asking/Unit` and
> `Vacancy %` (the "v2" pull), the script reads them and **cross-checks against
> RealPage** — occupancy gaps ≥2 pts and rent gaps ≥5% are flagged in Notes, and
> both sources sit side-by-side in the Reconciliation Log. Note CoStar rent is
> *asking* and RealPage rent is *effective*, so a gap is expected. Choose which
> source drives the chart with `--occ-source` / `--rent-source` (default CoStar).
> If the CoStar export lacks those columns ("v1"), occupancy/rent fall back to
> RealPage automatically.

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
Optional flags:
- `--as-of "2026 Q2"` — analysis quarter (defaults to latest in the analytics file)
- `--target 0.95` — stabilization occupancy goal
- `--occ-source costar|realpage`, `--rent-source costar|realpage` — which source drives the chart (default `costar`)
- `--pipeline-dates path.csv` — analyst-supplied delivery quarters for pipeline deals (see below)

The script prints a reconciliation report to the console and writes the workbook.
Always read the console report and the **Reconciliation Log** sheet, then review
the workbook before handing it off.

### Dating the pipeline (so it enters the absorption forecast)
Pre-Planned / under-construction deals usually have no delivery date, so they are
listed in the roster but **excluded from the absorption math** until dated. Every
run writes `<out>__pipeline_dates_TEMPLATE.csv` listing the undated pipeline
deals. Fill the `est_delivery` column (`Q2 2028`, units optional to override),
then re-run with `--pipeline-dates that.csv`. Dated pipeline deals enter the
quarterly table at 0% occupancy (lease-up not started), so their full unit count
flows into "units to reach target" and lifts **% to be absorbed** — i.e. the
forward supply the deal must compete with.

## Methodology

### 1. Reconcile the two rosters
- Match properties across CoStar and RealPage by **street address** (leading
  number + first significant street word; handles ranges and directionals), with
  a **name fallback** for missing/typo'd addresses.
- **Unit counts:** keep CoStar's; flag a note when the two sources differ by ≥3
  units (±2 is treated as noise).
- **Occupancy & rent:** both sources' values are kept side-by-side; the displayed
  value follows `--occ-source` / `--rent-source` (default CoStar), and material
  divergence (≥2 pts occ, ≥5% rent) is flagged in Notes.
- **Status:** a forward-looking status (Pre-Planned / Under Construction) from
  either source wins, so pipeline deals are never mislabeled as existing.
- The **subject property** is dropped from the competitive roster.

### 2. Pin the delivery quarter (the key cross-comparison)
Neither roster gives a delivery *quarter* (only a year). For **delivered**
properties only, the quarter is inferred from the CoStar Data Analytics quarterly
deliveries series: (1) an exact unit-count match within ±1 year pins the quarter
precisely (e.g. Timbers 274u → Q2 2023); (2) otherwise the same-year quarter with
the closest delivered count is used and flagged *"Delivery quarter estimated"*
(keeps the absorption table populated in large markets); (3) if the year has no
recorded deliveries, only the year is kept (`Q? <year>`). **Under-construction**
deals are auto-estimated from CoStar's `Construction Begin` + ~24 months
(reconciled against CoStar's completion year) and pushed into the future so they
enter the forecast; **proposed** deals with no date get `TBD` for the analyst.

### 3. Bucket each property (status-driven)
Buckets are driven by the lifecycle status from **both** sources, with precedence
*delivered → under-construction → proposed* (an "it exists" signal from either
source beats a stale "pre-planned"):
- **STABILIZED / STABILIZING** — delivered and ≥90% occupied (or an older,
  within-lookback asset that simply underperforms).
- **LEASING UP** — delivered within ~2 years and below 90% occupied. **Flagged
  "Verify rent/occ w/ HelloData"** — lease-ups are where accurate rent/occ matter
  most.
- **UNDER CONSTRUCTION** — CoStar `Under Construction` or RealPage `Under
  Construction[/Lease-Up]` (no rent/occ shown).
- **PROPOSED** — CoStar `Proposed` or RealPage `Pre-Planned`/`Planned` (no rent/occ).

Only genuine new supply is charted: deliveries within the last
`NEW_CONSTRUCTION_LOOKBACK_YEARS` (default 4) plus the entire pipeline. Each
section is colour-coded, carried over from the reference template: **blue**
(stabilized), **orange** (leasing up), **green** (under construction), **red**
(proposed), each with a light matching row tint.

### 4. Supply & absorption forecast (the "Supply & Absorption" tab)
The rent-analysis core: a market-level view of **new supply, absorption, and
overall occupancy** in **relative-year (trailing-12-month) columns**, built from
the CoStar Data Analytics series + the supply pipeline, on **live Excel formulas**.

- **Columns** are TTM windows anchored to the as-of quarter: **Y0** = the T12
  ending at as-of, **-Y1…-Y6** step back a year each (toward 2020); **Y1…Y6** are
  the hold years anchored to an editable **Close Quarter** (default as-of + 2
  quarters), so the analysis→close gap is one cell.
- **Historical years** = CoStar 5-mi actuals (Σ deliveries, Σ absorption,
  end-of-window occupancy).
- **Forecast years**: `Inventory += scheduled pipeline` (pipeline deals are
  assigned a hold-year by their delivery quarter vs. the close quarter);
  `Occupied += selected annual demand` (capped at target × inventory);
  `Occupancy = Occupied / Inventory`. Also a **cumulative-unabsorbed** row since -Y3.
- **Demand** is a Bear/Base/Bull annual-absorption assumption (Base = trailing
  CoStar average), with an **Implied 5-mi Occupancy** block showing all three
  paths, plus editable **market-rent-growth / concession** blocks and a derived
  **effective-rent-growth** block (the bridge to rents).
- **Editable (yellow) cells**: close quarter, demand scenario & the three
  absorption levels, each pipeline deal's delivery quarter & Include? (Y/N), and
  the rent-growth/concession assumptions.
- A **reconciliation note** compares the scheduled pipeline (Include=Y) against
  CoStar's current Under-Construction unit count so the pipeline ties out.

Reading the occupancy row shows **when the market re-stabilizes** to target — i.e.
when occupancy is strong enough to push rents. (Subject rent rows from RR-T12 are
a planned next layer.)

## Analyst follow-ups the script intentionally leaves open
- **Proximity (miles)** — left blank; neither export carries distance-from-subject. Fill manually (or paste CoStar's "Distance" column if a future pull includes it).
- **Lease-up rent & occupancy** — verify the flagged lease-ups with a HelloData pull.
- **Pipeline timing** — undated Pre-Planned deals (`TBD`) are listed but not in the forecast; fill the emitted `…__pipeline_dates_TEMPLATE.csv` and re-run with `--pipeline-dates` to fold them in.

## Tunable thresholds
Top of `scripts/build_supply_chart.py`: `DEFAULT_STABILIZATION_TARGET`,
`STABILIZED_OCC`, `LEASEUP_WINDOW_QTRS`, `NEW_CONSTRUCTION_LOOKBACK_YEARS`.

## Worked examples
`examples/bella_mirage/` is a large, supply-heavy Phoenix market (104 comps,
26.6k-unit inventory) that exercises all four buckets — ~8.8% of inventory still
absorbing from current lease-ups alone, before a large proposed pipeline.

`examples/canyon_ridge/` holds the real exports for Canyon Ridge (Boise, ID),
including the CoStar v2 pull (with rent/occ) and a sample `pipeline_dates.csv`.
Running the command above reproduces `output/Canyon_Ridge__Supply_Chart.xlsx`:
3 stabilized comps (445 u, ≈fully absorbed near-term) plus a 1,763-unit proposed
pipeline. With the pipeline dated via `--pipeline-dates`, the deliveries spread
across 2026–2028 and "% to be absorbed" rises to ≈28% — the forward supply the
deal will compete with.
