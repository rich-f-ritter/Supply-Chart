---
name: supply-chart
description: >-
  Build a 5-mile new-construction "Supply Chart" for a multifamily underwriting
  deal. Reconciles CoStar and RealPage 5-mile-radius exports (plus the CoStar
  Data Analytics time series) into one competitive-supply roster, buckets each
  property by lifecycle stage, and writes a formatted workbook with a relative-
  year (TTM) Supply & Absorption forecast — new supply, absorption, and overall
  occupancy projected forward with editable demand scenarios and pipeline
  toggles, plus subject rent/occupancy rows from the RR-T12 intake. Use when the
  user asks to build, update, or automate a supply chart / competitive supply or
  rent analysis, or provides CoStar + RealPage radius exports for a deal.
---

# Supply Chart

## What this produces
A workbook (mirroring `reference/EXAMPLE_Seasons_Supply_Chart_v2.xlsx`) with:

1. **Competitive Analysis** — the new-construction roster grouped into four
   lifecycle (colour-coded) buckets.
2. **Supply & Absorption** — the relative-year (TTM) supply / absorption /
   occupancy forecast, subject rent & occupancy rows, demand & rent-growth
   scenarios, and the editable pipeline blocks (see §4). All editable inputs live
   here (yellow cells), including the stabilization target.
3. **Reconciliation Log** — every property with its merged values, sources, and
   any conflict notes, so the analyst can audit the automated decisions.

This supply layer is the input to the **rent-analysis** step: it sits on top of
the historical overall occupancy / absorption / deliveries series (CoStar Data
Analytics) so demand can be forecast and the "when is occupancy strong enough to
push rents" question answered. (Rent analysis is a *separate, later* step — this
skill only builds the supply chart.)

## Operating procedure (the order to work in)
1. **Build the supply chart standalone first** — reconcile the exports, bucket the
   roster, build the Supply & Absorption forecast. Hand it over as its own workbook.
2. **Let the user refine** — delivery quarters, demand scenarios, source selection,
   diligence corrections. Re-run with their edits.
3. **Make it jive with every input you were given** before calling it done:
   - the **RR-T12 intake** (from the RR-T12-Processor skill) drives the subject
     market/effective-rent and occupancy rows — confirm those reconcile;
   - the delivered roster ties to CoStar's deliveries series (the completeness
     check, §3a);
   - diligence corrections (if run) are reflected in the buckets and pipeline.
4. **Only then ask whether to incorporate into the underwriting model.** If yes,
   regenerate with `--model-link` and follow **Incorporating into the underwriting
   model** below. Don't link to the model until the chart itself is settled.

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
- `--intake path.xlsx` — RR-T12 underwriting intake; adds **subject rows** (market/effective rent from HelloData mix-weighted, occupancy from the T12 financials) to the Supply & Absorption tab
- `--costar-subject-rents path.xlsx` and/or `--realpage-subject-rents path.xlsx` — per-property rent histories used to extend the subject rents before HelloData coverage; whichever tracks HelloData more closely in the overlap is auto-selected (varies by market)
- `--no-geocode` — skip Nominatim and fill the **Proximity (mi)** column from offline ZIP centroids only (no network)
- `--model-link` — wire the subject rows to the underwriting model (see **Incorporating into the underwriting model** below). Off by default; only used in the incorporation step, not the standalone deliverable. `--model-sheet "…"` overrides the target tab (default `Cash Flow (Annual)`)

The script prints a reconciliation report to the console and writes the workbook.
Always read the console report and the **Reconciliation Log** sheet, then review
the workbook before handing it off.

### Dating the pipeline
Under-construction deals are auto-dated (from CoStar `Construction Begin` + ~24
mo) and flow into the forecast automatically. Proposed deals without a date are
listed but undated; every run writes `<out>__pipeline_dates_TEMPLATE.csv` of the
undated deals. Fill `est_delivery` (`Q2 2028`) and re-run with `--pipeline-dates
that.csv`, or just edit the delivery quarter directly in the **Proposed Pipeline**
block on the Supply & Absorption tab. Whether a proposed deal adds supply is
controlled by its **Built In** toggle (Bear only / Bear+Base / All / None).

### Companion map (optional)
`scripts/build_map.py` plots the subject + competitive roster on a **satellite**
basemap (default; `--base terrain|streets` to switch), colour-coded by the same
lifecycle buckets, with a yellow 5-mile ring and per-property popups. Markers use
saturated colours and white halos so they read over imagery (Under Construction =
spring green). It writes an interactive **HTML** map and a static **PNG**.
```bash
python scripts/build_map.py \
  --subject-name "Aura Beacon Island" \
  --subject-address "2200 Beacon Cir, League City, TX 77573" \
  --costar-roster examples/aura_beacon_island/CoStar_5mi_50unit_properties.xlsx \
  --costar-analytics examples/aura_beacon_island/CoStar_5mi_Data_Analytics.xlsx \
  --realpage examples/aura_beacon_island/Realpage_5mi.xlsx \
  --out output/Aura_Beacon_Island__Map.html
```
Geocoding: street-level by default (Nominatim, cached, needs network); pass
`--no-geocode` for offline **ZIP-centroid** placement (approximate — same-ZIP
properties are jittered apart). Pass `--subject-latlng "lat,lng"` to pin the
subject exactly. Needs `folium`, `zipcodes`, `matplotlib`. The chart and the map
share one geocode cache (`output/.geocode_cache.json`) via `scripts/geo.py`.

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
- **UNDER CONSTRUCTION** — ground has broken: a UC status from either source, **or
  `Construction Begin` ≤ the as-of quarter** (even if a source still lags as
  "Proposed").
- **PROPOSED** — not yet started: `Construction Begin` in the future (or absent)
  and a proposed/pre-planned status. A future begin still yields an *estimated
  completion* (begin + ~24 mo) for the pipeline, but the deal stays Proposed
  until ground breaks.

Only genuine new supply is charted: deliveries within the last
`NEW_CONSTRUCTION_LOOKBACK_YEARS` (default 4) plus the entire pipeline. Each
section is colour-coded, carried over from the reference template: **blue**
(stabilized), **orange** (leasing up), **green** (under construction), **red**
(proposed), each with a light matching row tint.

### 3a. Completeness check — tie the roster to CoStar's deliveries series
Before trusting the stabilized/lease-up section, **reconcile the named comps
against CoStar's own delivered-units series** per TTM window. For each relative
year, sum the analytics `deliveries` and compare to the roster units delivered in
that window; they should jive (often to the unit). When they don't, the gap is
one of:
- **The subject itself.** CoStar's *submarket* deliveries include the subject,
  but the subject is dropped from the *competitive* roster — so always **add the
  subject's units back** before comparing. (Worked example: Aura -Y3 = 1,221 =
  Reserve 291 + Livano 325 + Lenox 315 **+ subject Aura 290** — exact once the
  subject is added back; the apparent 290 "gap" was the subject, not a missing
  comp.)
- **Sub-50-unit product** CoStar's submarket series counts but the 50-unit
  property pull never lists (small, can't be itemized — note it and move on).
- **A genuinely missing or mis-bucketed 50+ comp** — the case worth catching.
  Look for a property whose CoStar `year_built` falls just outside the lookback
  (so it was filtered out) but which actually *delivered* inside the window, or
  one pinned to the wrong quarter. A delivery quarter can be off by one (the
  pin estimates from year/unit-count) without changing a TTM total, but a comp
  dropped for an off-by-one **year** will silently vanish — check it.

This is a fast, high-value audit: if the three historical TTM windows tie to
CoStar's deliveries (modulo the subject and sub-50 product), the stabilized
section is complete.

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
- **Forecast years**: `Inventory += scheduled pipeline`; `Occupied += selected
  annual demand` (capped at target × inventory); `Occupancy = Occupied /
  Inventory`. Plus a **cumulative-unabsorbed** row since -Y3.
- **Pipeline = two blocks** (right side): **UNDER CONSTRUCTION** is linked from
  the Competitive Analysis roster and counts in every scenario; **PROPOSED** is
  speculative with a per-deal **"Built In"** dropdown (*Bear only / Bear+Base /
  All / None*) so a deal layers into the downside (more-supply) case only. Each
  hold-year's new supply = UC + proposed built under the selected scenario.
- **Subject rows**: market & effective rent, occupancy, and **concession %**, plus
  **YoY %** rows. A collapsed ("+") detail group lists the source used per year
  (HelloData vs CoStar) and the UC-vs-CoStar reconciliation. Forward concession %
  is an editable assumption block.
- **Demand** is a Bear/Base/Bull annual-absorption assumption (Base = trailing
  CoStar average), with an **Implied 5-mi Occupancy** block (all three paths),
  editable **market-rent-growth / concession** blocks, and a derived
  **effective-rent-growth** block.
- **Editable (yellow) cells**: stabilization target, close quarter, demand
  scenario & the three absorption levels, each proposed deal's delivery quarter &
  Built-In, and the rent-growth/concession assumptions.
- A **reconciliation note** compares the scheduled pipeline (Include=Y) against
  CoStar's current Under-Construction unit count so the pipeline ties out.

Reading the occupancy row shows **when the market re-stabilizes** to target — i.e.
when occupancy is strong enough to push rents.

**Subject rows** (with `--intake`): the tab also shows the subject's own
**market rent, effective rent, and occupancy** in the same relative-year columns.
Source hierarchy: **HelloData mix-weighted** rents where available (≈2023→);
before that, **CoStar or RealPage per-property** rents — whichever tracks
HelloData more closely in the overlap is auto-selected (CoStar tends to win in
some markets, RealPage in others), level-aligned to HelloData via the overlap
ratio. The chosen source is shown in the row label (`…→HD`) and per-year in the
collapsed detail group. Subject occupancy comes from the **T12 financial
statements** (filled from CoStar where the financials have gaps). Forecast-year
subject rents are deferred to the (later) market-rent-growth → effective step —
unless the chart is linked to the model (next section), in which case Y0→Y6 read
the model's own projection.

## Incorporating into the underwriting model
After the chart is settled and the user asks to incorporate it (step 4 above),
run with `--model-link`. This rewrites the subject **Market Rent**, **Effective
Rent**, and **Occupancy** rows (cols Y0→Y6) as **internal references** to the
TMG model's `Cash Flow (Annual)` tab — so the chart shows the deal's own
projected path next to the 5-mile market frame, as a **sanity overlay** (the
chart reads the model; it does *not* drive it). The fixed mapping:

| Subject row | Model row (`Cash Flow (Annual)`) | Y0 = col I | Y1–Y6 = cols J:O |
|---|---|---|---|
| Market Rent | row 4 (Market Rent /U) | ← `F` | ← `K,L,M,N,O,P` |
| Effective Rent | row 5 (Eff Mkt Rent /U) | ← `F` | ← `K:P` |
| Occupancy | row 14 (Physical Occupancy) | ← `F` | ← `K:P` |

Historical columns (−Y6…−Y1) stay as the static RealPage/HelloData actuals. The
YoY% and Concession% rows are live formulas, so once linked they automatically
show the **model's** market-rent growth, effective-rent growth, and concession
trend in the same frame as the market assumption blocks — the comparison is
implicit, no separate delta block. Refs are written **internal** (`='Cash Flow
(Annual)'!…`), so they resolve the moment the tab lives in the model.

**Two delivery modes** (both produce the same linked tabs):
- **User drags the tabs in** — they Move/Copy the `Supply & Absorption` (and its
  `Competitive Analysis` dependency) sheets into the model; the internal refs
  resolve to the model's own `Cash Flow (Annual)`. In the standalone file those
  Y0→Y6 cells read `#REF!` until the tab is in the model (everything else shows).
- **You embed it for them** — given the model, insert the linked tabs **without
  round-tripping the .xlsm through openpyxl** (it silently drops charts,
  conditional formatting, slicers — it warns on this very model). Graft the
  sheet XML into the package directly so the live macro workbook is preserved.

Keep the distinction the model already encodes: the chart's **market-rent
growth** is a pure asking-rent trend; the model's **AGPR growth** bakes in
loss-to-lease/gain-to-lease burn-off and lease-term roll. They are *not* the same
line — compare market-rent growth to the model's `RRA` GPR-growth row, not AGPR.

## Research & diligence (opt-in)
The mechanical build trusts whatever CoStar/RealPage report. When the user asks to
**research the pipeline / verify the supply / do diligence** (or runs with
`--diligence`), perform this phase — it is not run by default.

Every build emits `…__diligence_TEMPLATE.csv` (cols: `type,property,units,
est_delivery,status,leasing_pace,notes,source`) pre-listing the lease-up /
under-construction / proposed deals. The research workflow:

1. **Per-project diligence (Part A).** For each lease-up / UC / proposed deal,
   web-research and confirm: actual construction status (proposed / permitted /
   under construction / leasing / delivered / **stalled or cancelled**), expected
   delivery quarter, developer/owner, unit count, any **affordable or
   age-restricted** component, and leasing pace. **Cite a source URL.** Correct
   the `status`, `est_delivery`, and `units` — these fold back into the buckets
   and the absorption forecast (`type=pipeline`).
2. **Shadow-supply scan (Part B).** Research the 5-mile radius for **latent supply
   not in the vendor data**: multifamily **rezonings**, **entitled / under-contract**
   apartment land, large vacant MF-zoned tracts, master-planned MF phases, and
   newly announced developments. Add each as a `type=shadow` row with a source.
3. Be skeptical — if a project can't be verified, say so in `notes` rather than
   guessing. Then re-run with `--diligence filled.csv`: the workbook gains a
   **Diligence** sheet (pipeline diligence + shadow-supply watch list) and the
   researched delivery dates flow into the forecast. Pass the same CSV's
   `type=shadow` rows to `build_map.py --shadow-supply` to plot them.

## Proximity
The **Proximity (mi)** column is filled automatically: the script geocodes the
subject and each comp (OpenStreetMap Nominatim, cached in
`output/.geocode_cache.json`) and writes the straight-line distance. Comps that
only resolve to a ZIP centroid — or that sit just past the 5-mile road radius —
can read slightly over 5 mi; the export defines membership, the column is an
approximate as-the-crow-flies distance. Geocoding tries the full address, then
relaxes (drops the city, which roster exports often get wrong) and **rejects any
hit more than 8 mi from the subject** before falling back to the ZIP centroid, so
a mis-geocoded street name can't produce a wild distance. Pass `--no-geocode` to
skip the network call and use offline ZIP centroids only.

## Analyst follow-ups the script intentionally leaves open
- **Lease-up rent & occupancy** — verify the flagged lease-ups with a HelloData pull.
- **Pipeline timing** — undated Pre-Planned deals (`TBD`) are listed but not in the forecast; fill the emitted `…__pipeline_dates_TEMPLATE.csv` and re-run with `--pipeline-dates`, or research them via the diligence phase above.

## Tunable thresholds
Top of `scripts/build_supply_chart.py`: `DEFAULT_STABILIZATION_TARGET`,
`STABILIZED_OCC`, `NEW_CONSTRUCTION_LOOKBACK_YEARS`, `CONSTRUCTION_MONTHS`.

## Worked examples
`examples/bella_mirage/` — a large, supply-heavy Phoenix market (104 comps,
26.6k-unit inventory) that exercises all four buckets and the subject rows. Run:
```bash
python scripts/build_supply_chart.py \
  --subject-name "Bella Mirage" --subject-address "3800 N El Mirage Rd" \
  --costar-roster    examples/bella_mirage/CoStar_5mi_50unit_properties.xlsx \
  --costar-analytics examples/bella_mirage/CoStar_5mi_Data_Analytics.xlsx \
  --realpage         examples/bella_mirage/Realpage_5mi.xlsx \
  --intake           examples/bella_mirage/Bella_Underwriting_Intake.xlsx \
  --costar-subject-rents examples/bella_mirage/CoStar_5mi_property_rents.xlsx \
  --out output/Bella_Mirage__Supply_Chart.xlsx
```

`examples/canyon_ridge/` holds the real exports for Canyon Ridge (Boise, ID),
including the CoStar v2 pull (with rent/occ) and a sample `pipeline_dates.csv`.
Running the command above reproduces `output/Canyon_Ridge__Supply_Chart.xlsx`:
3 stabilized comps (445 u, ≈fully absorbed near-term) plus a 1,763-unit proposed
pipeline. With the pipeline dated via `--pipeline-dates`, the deliveries spread
across 2026–2028 and "% to be absorbed" rises to ≈28% — the forward supply the
deal will compete with.
