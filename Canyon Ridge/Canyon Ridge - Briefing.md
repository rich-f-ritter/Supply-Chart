# Canyon Ridge — 5-Mile Land-Use & Supply Briefing
**Subject:** Canyon Ridge apartments, 2552 E Gowen Rd, Boise, ID 83716 (SE Boise / Ada County)
**Analysis area:** 5.0-mile radius (~50,185 acres) · **Run date:** 2026-06-26
**Data:** Ada County Assessor parcels (28,021 in-area) + Ada County/Boise zoning + Boise Airport Influence Areas + avigation easements + the City Development Tracker, reconciled with sourced web research.

> **One-line takeaway:** The Boise Airport is the defining feature of this submarket. ~37% of the land in the ring is under the Airport Influence Area overlay, and **~9,400 acres (about 19% of the radius) cannot add new housing** because of the flight path. That makes the airport an effective *supply moat* on the subject's airport-facing sides — the land nearest Canyon Ridge is building out as Micron/industrial, not competing apartments — while the genuine apartment competition sits farther out, away from the cone.

---

## 1) Who owns what (and a hard data limit)
**Idaho's Public Records Act bars Ada County from publishing owner names in its free GIS** — and the statewide cadastral redacts owner too. So there is **no bulk owner roll** to map. Ownership below is established from research + parcel location/zoning, not the assessor roll. (For named owners on specific parcels, the Ada County Assessor parcel viewer at adacountyassessor.org gives owner per-parcel.)

The dominant landowners in the ring:
- **City of Boise / Boise Airport** — the single largest holder, **>5,000 acres** around the airfield, ground-leased for industrial development. (iflyboise.com Airport Master Plan)
- **Micron Technology** — HQ campus + expansion, **~900 acres** and growing (see §3).
- **Idaho Air National Guard / Gowen Field** — ~570 ac exclusive military lease + ~1,500 ac joint-use with the City; 124th Fighter Wing. (gowenstrong.com)
- **Industrial developers on airport/City land** — Adler Industrial (Boise Airport Industrial Holdings), Boyer Co. (Boise Gateway, wraps the WinCo DC), Flint Development (Red River Logistics, ~120 ac/1.3M sf on Gowen), Sudler, Ball Ventures Ahlquist (Pleasant Valley, 520 ac). (CCDC, BoiseDev, Adler, Boyer)
- **Harris family / Barber Valley Development** — ~1,300-ac Harris Ranch master-planned community on the NE edge of the ring.
- **State of Idaho (Dept. of Lands)** — scattered endowment parcels near the airport (e.g., a Gowen/Federal Way ground lease).

Land-use mix in the ring (by acreage): Agricultural/Rural ~20,300 ac and Vacant ~15,300 ac (mostly foothills "Rural Preservation," desert, and airport buffer — **not** developable), Single-Family Residential ~5,000 ac, Commercial ~4,600 ac. See the **Parcels** sheet for all 28,021 classified parcels.

## 2) The flight-path / "undevelopable" question — confirmed, and quantified
Your recollection is **substantially correct, but tier-dependent.** Boise (and Ada County) enforce an **Airport Influence Area Overlay (AI-O)** with four sub-zones:

| Zone | New residential? | Acres in 5-mi ring |
|---|---|---|
| **A** | Allowed, with 25 dB sound insulation + avigation easement | ~9,219 |
| **B-1** | Allowed but **capped at 5 units/acre** (+30 dB, easement) | ~63 |
| **B** | **PROHIBITED** (also bars schools/daycare/worship) | ~4,460 |
| **C** | **PROHIBITED** (existing housing = legal nonconforming) | ~4,961 |

- **~18,700 ac (37% of the ring) is inside the overlay; ~9,421 ac (Zones B+C, ~19% of the ring) is off-limits to new housing.**
- **Canyon Ridge itself plots inside Zone B** — yet it was built as a 287-unit complex in 2024 (via prior entitlement / development agreement / legal-nonconforming status). A notable tension worth knowing, not a contradiction in the data.
- Enforcement is real and recent: **Elder Street Apartments (84 units) was DENIED in 2020** over airport noise (P&Z + unanimous Council); **1770 S Maple Grove (174 units) was P&Z-recommended-denial in 2024** for exceeding the 5 du/ac airport cap, then narrowly approved on appeal over the Airport Director's objection — with Council ordering a re-look at the rules. (BoiseDev 2020-09-02; 2024-09-11)
- **Headwind ahead:** the 124th Fighter Wing transitions **A-10 → F-16 around 2027** (Air Force EIS underway); louder jets could *expand* the restricted noise footprint. (BoiseDev 2023-06-27; KTVB)

See **Tables/airport_influence_area_acres.csv** and the **Assumptions & Decisions** sheet.

## 3) What Micron owns / what's been happening
- **HQ + fab campus at 8000 S Federal Way**, on the "Columbia Bench" (south of ID-21, north of I-84, east of Federal Way). Campus pushing toward **~900 acres** with **two new fabs**, a **~$50B** Boise expansion (largest private investment in Idaho history); first DRAM wafers targeted **2H 2027**; ~17,000 projected Idaho jobs. (BoiseDev 2022–2025; Micron; Idaho Commerce)
- **Land assembled by successive annexations:** ~358 ac (Dec 2022) + ~596 ac (Dec 2023, including **79 ac rezoned "planned community" and explicitly held undeveloped**) + **238+ ac (Apr 2026)** for support/power — a 40-ac electrical yard and mile-long, 40-ft earth berms abutting the Painted Ridge/Sunny Ridge neighborhoods. Roughly **7,470 acres of large (>20-ac) tracts** sit within 2.2 mi of the campus (Columbia Bench: RUT/RP/I-3/M3 land) — the frontier the City is now master-planning. (BoiseDev 2022-12-08; 2023-12-08; 2026-04-06 — see **Tables/micron_area_large_tracts.csv**)
- **Dec 2025:** Micron filed for a new 51-ft corporate jet hangar at 1250 W Gowen Rd on the airport's south side. (BoiseDev 2025-12-05)
- **Jan 2026:** the City launched a ~12-month planning effort for three undeveloped SE Boise zones — **South Airport, East Columbia (south of Gowen, east of Micron), and the "Third Bench."** Unlocking residential there hinges on the **Lake Hazel Rd extension** and a secondary access road (wildfire egress). (BoiseDev 2026-01-28)

## 4) What's proposed now (live Development Tracker, within 5 mi)
259 active planning records sit in the ring (**Tables/development_tracker_5mi.csv**). Highlights:
- **At the subject:** 2774/2350 E Gowen Rd — entitled project + conditional-use and design-review applications in the Columbia town-center node (Zone B).
- **Micron:** 8000 S Federal Way — applications in review / neighborhood meetings.
- **Columbia Bench frontier:** an **Annexation-Rezone + Comprehensive Plan Amendment in review** at S1605336000 (~1.1 mi, Zone A) — the kind of master-plan/annexation that precedes large residential.
- **Airport/Eisenman corridor:** subdivisions and projects at 6490/6651/6880 S Eisenman, 2500 E Freight (Boise Gateway), 601 W Gowen — all **industrial** (Zone B/C).
- **Barber Valley (outside the cone):** residential projects at Hopes Well Way, Brightside St, Millbrook Way.
- Larger sourced pipelines nearby (mostly **for-sale single-family**, not apartment competition): **East Hollow ~7,000 homes** near Micron, CBH **"Locale" ~2,000**, **Murio Farms ~3,000** — all gated by infrastructure and, in Locale's case, held up Feb 2026 over fire-service questions. (BoiseDev 2025-11-26; 2026-02-03)

## 5) Competitive apartment-supply read (the ranked threats)
Only **~248 acres of vacant land in the entire ring carries by-right multifamily zoning**, and much of it is airport-restricted or small. After layering the airport overlay onto base zoning, the genuine apartment threats sit **outside the cone**, generally 2–4.5 mi away (full reasoning in the **Top Vacant Threats** sheet / **Tables/top10_concerning_vacant.csv**):
1. **62 ac vacant R-3, W Victory Rd (~3.5 mi, outside AIA)** — by far the largest by-right MF land basis.
2. **~22.8 ac SP-01 + MX-3/MX-1 tracts around Parkcenter / the Boise River (~2–4 mi)** — mixed-use with residential.
3. Scattered **R-2/R-3 infill** (Oakland, River Run, Hillcrest, Fairbrook).
- **Demoted to "Low (airport-restricted)":** the MX-2 parcels closest to Canyon Ridge (7983 S Federal Way 0.84 mi, 6650 S Eisenman 0.61 mi, 2149 E Hospitality 0.34 mi) — MF-zoned on paper but in Zone B, so **no new apartments**; they'll be airport/industrial.

**Underwriting implication:** the airport shields Canyon Ridge from new apartment supply on its airport-facing sides; model competition from the non-airport SE/Parkcenter submarkets, and watch the F-16 noise EIS and the East Columbia/South Airport planning effort.

---
*Deliverables: `Canyon Ridge - Land Use Analysis.xlsx` (10 sheets), `Canyon Ridge - Land Use Viewer.html` (interactive map: land use / zoning / vacant threats), and `Tables/*.csv`. Caveats: no public owner roll (Idaho law); PROPCODE land use is coarse (no single-family/apartment split — zoning used for the MF read); AI-O zone lines are the City's mapped polygons. Full source list + per-decision rationale in the Assumptions & Decisions sheet / `Tables/decisions_log.md`.*
