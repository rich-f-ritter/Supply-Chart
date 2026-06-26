# Land Use Analysis — Canyon Ridge (2552 E Gowen Rd, Boise, ID 83716) — Decisions Log

Run date: 2026-06-26. Analyst: automated land-use-analysis skill + reasoned review.

## 1. Subject location (verified point)
- Address: 2552 E Gowen Rd, Boise, ID 83716 (Canyon Ridge — a 287-unit, 4-story apartment community completed 2024).
- Point used: lat 43.541734, lon -116.151198 (ArcGIS World GeocodeServer, PointAddress, score 100).
- Verification note: the geocoded point lands on the **Columbia Town Center** MX-2 node (parcel R1525500185, "2774 E Gowen Rd", $53M improved town-center) at Gowen Rd & Federal Way. The Canyon Ridge apartment footprint is within ~0.1 mi of this point in the same Columbia Village / Columbia Town Center MX-2 node; the ~0.1-mi offset is immaterial to a 5-mile-radius study, and the airport-overlay classification (AIA Zone B) is identical across the node. Distances are anchored to the geocoded point.
- The Census onelineaddress geocoder returned no match (rural address ranges); ArcGIS World was used instead.

## 2. Analysis area
- 5.0-mile geodesic radius (user request) around the subject point. Circle area ≈ **50,185 acres**.
- Spans: Boise Airport (BOI) / Gowen Field; Micron HQ + expansion (Columbia Bench); the Gowen Rd / Federal Way / S. Orchard / Eisenman industrial corridor; Columbia Village; the Parkcenter / Boise River / Barber Valley edge; and unincorporated Ada County rural-urban frontier to the south and east.

## 3. Parcel data source + coverage
- **Ada County Assessor Parcels (AdaCountyGIS)**, ArcGIS FeatureServer layer 5 (services2.arcgis.com/dgGjZc6xAH5m5JyP). Pulled via POST + objectId paging; 28,021 parcels kept in-area.
- Fields mapped: account=PARCEL, situs=ADDRESS, landuse_code=PROPCODE, legal=LEGAL1, total_value=TOTALVALUE; dedupe on PARCEL.
- **Owner names are NOT available.** Idaho's Public Records Act bars Ada County from publishing owner names/mailing addresses in the free GIS, and the IDWR statewide cadastral also returns OWNER = null for Ada County. Consequence: the "who owns what" question is answered from research + location/zoning inference, NOT a bulk owner roll. The pipeline's owner-based filters (HOA/gov screen, owner clustering) are therefore inert; HOA commons and public land were instead screened by acreage/compactness and reasoned out by hand.

## 4. Land-use source & granularity
- Ada County **PROPCODE** single-letter property class (confirmed via Ada County Assessor GIS User Guide): **R = Residential, C = Commercial, F = Farm, L = Land, M = Manufactured.**
- Crosswalk to canonical buckets: R, M → Single Family Residential; C → Commercial; F → Agricultural / Rural; L → Vacant Land; blank → Other/Unclassified.
- This is coarse: PROPCODE does not separate single-family from apartments, so existing multifamily cannot be isolated from the land-use layer (zoning is used for the MF-supply read instead). 'L' (Land) mixes true vacant lots with $0-value HOA/condo common areas; 'F' (Farm) is ag acreage including Boise River floodplain.
- **"Vacant" definition for the supply screen:** buckets L (Vacant Land) **and** F (Agricultural / Rural) — i.e., undeveloped land — then mechanically filtered and hand-reasoned to remove non-developable parcels.

## 5. Zoning source + honest gaps
- **Ada County / Boise Zoning** (AdaCountyGIS Zoning FeatureServer layer 22), 492 polygons in-area. code_field = BASEZONE (suffix-stripped base zone), desc_field = ZONING (carries /DA development-agreement and /AI-O airport-overlay suffixes). One source labeled jurisdiction "Ada County/Boise"; the layer's CITY field distinguishes City of Boise vs unincorporated Ada County districts (codes are distinct enough that a merged crosswalk is unambiguous except A-1/A-2, both mapped Low).
- Zoning coverage is effectively complete in-area (the parcel layer's own ZONING field corroborates); only 1 parcel was unmapped on first pass (MX-5) and was added. No material data-gap zone.

## 6. Zoning crosswalk corrections / calls
- High (MF by-right): R-3, MX-2, MX-3, MX-U, MX-5, county R12. Medium (MF conditional/low-density): R-2, county R8, MX-1. Low: R-1A/B/C, county R1/R6, C2, I-1/2/3, county M1/2/3, A-1/A-2 (open/ag), RUT, RR, RP (rural preservation/foothills). Unknown: SP-01/SP-02 (Specific Plan — negotiated).
- **Airport overlay is decisive and is layered ON TOP of base zoning.** A parcel zoned MX-2/MX-3/R-3 (MF by-right) inside Airport Influence Area Zone B or C **cannot add new residential** under City of Boise §11-02-07 (AI-O) / Ada County Title 8 Art. A. Such parcels were down-ranked to "Low (airport-restricted)" regardless of base zone. This is the single most important correction in the analysis.

## 7. Airport Influence Area (AI-O) — the "flight path" question
- Source: City of Boise "Airport Influence Areas" GIS layer (services1.arcgis.com/WHM6qC35aMtyAAlN), 7 polygons in the radius; corroborated by avigation-easement layer (567 recorded easements in-area).
- Zone meanings (City of Boise AI-O overlay page): **A** = new residential allowed with 25 dB sound insulation + avigation easement; **B-1** = residential allowed but capped at 5 units/acre (CUP for more) + 30 dB insulation; **B** = new residential PROHIBITED (most restrictive; schools/daycare/worship also barred); **C** = new residential PROHIBITED (existing = legal nonconforming).
- **Acreage within the 5-mi radius by zone: A ≈ 9,219 ac; B ≈ 4,460 ac; C ≈ 4,961 ac; B-1 ≈ 63 ac.** Total AIA ≈ 18,700 ac (~37% of the circle). **Land where NEW residential is prohibited (B + C) ≈ 9,421 ac (~19% of the radius).**
- **The subject (Canyon Ridge) plots inside Zone B.** It was nonetheless built as a 287-unit complex in 2024 — likely via prior entitlement / development agreement / legal-nonconforming status; flagged as a notable tension, not independently adjudicated here.
- The user's recollection ("some land is undevelopable because it's in a flight path") is **substantially correct but tier-dependent**: true for Zones B and C (no new housing) and for land inside the DNL 65 noise contour / runway protection zones off the runway ends; not a blanket rule (Zones A and B-1 allow residential with insulation, density caps and an avigation easement). The pending A-10 → F-16 transition (124th Fighter Wing, ~2027; Air Force EIS underway) may EXPAND the noise/restriction footprint.

## 8. Vacant-threat assessment — mechanical pre-filters + reasoning
- Mechanical (prepare_candidates): vacant bucket (L or F), zoning not a data gap, **min 2.0 acres**, Polsby-Popper compactness ≥ 0.16 (drops road slivers/detention shapes), adjacency clustering. 707 vacant parcels ≥ filters → 93 clusters. Dropped: 2,544 too-small, 66 sliver-shaped.
- Reasoning layered the **AIA overlay** and research on top: MF-zoned vacant parcels in Zone B/C were demoted to "Low (airport-restricted)"; huge A-1/RP clusters (foothills preservation, airport buffer, desert — over-merged because blank owner defeats owner-clustering) were treated as non-developable; river/gravel-pit "Loggers Pond/River Run" R-2/R-3 parcels were flagged Verify (possible water/HOA open space).
- Only **~248 acres of vacant land carries High (MF by-right) zoning** in the entire radius, and much of that is airport-restricted or small. The genuine competitive-apartment threat is concentrated OUTSIDE the airport cone (W Victory R-3; Parkcenter MX-3/MX-1/SP-01; scattered R-2/R-3), generally 2–4.5 mi from the subject. See Top-N table.

## 9. Micron & major landowners (research-sourced; owner roll unavailable)
- **Micron Technology** — HQ at 8000 S Federal Way; campus pushing toward ~900 acres with two new fabs (~$50B, first DRAM 2H 2027; ~17,000 projected ID jobs). Land assembled via successive annexations: ~358 ac (Dec 2022), ~596 ac (Dec 2023, incl. 79 ac rezoned "planned community" and held undeveloped), 238+ ac (Apr 2026, support/power, 40-ac electrical yard, mile-long berms by Painted Ridge). ~7,472 ac of large (>20-ac) vacant/ag/commercial tracts sit within 2.2 mi of the campus (Columbia Bench: RUT/RP/I-3/M3), much of it the development frontier the City is studying (Jan 2026 "South Airport / East Columbia / Third Bench" planning effort). New corporate jet hangar filed Dec 2025 at 1250 W Gowen Rd.
- **Boise Airport / City of Boise** — dominant landowner (>5,000 ac around the airfield), ground-leasing for industrial (Adler/Boise Airport Industrial Holdings, Boyer/Boise Gateway, Sudler, Red River/Flint, Pleasant Valley/Ball Ventures-Ahlquist).
- **Idaho Air National Guard / Gowen Field** — ~570 ac exclusive military lease + ~1,500 ac joint-use; 124th Fighter Wing (A-10 → F-16 ~2027).
- **Harris family / Barber Valley Development** — ~1,300-ac Harris Ranch master-planned community (NE edge of radius).
- Owner attribution for individual parcels is inference-based (location + zoning + research), NOT the assessor owner roll, which is not public.

## 10. Proposals turned down / restricted (research-sourced)
- **Elder Street Apartments (84 units, 2020)** — DENIED; Boise P&Z recommended denial and City Council unanimously upheld, citing airport noise/flight-path proximity and a formal Airport objection; density-capped near the airport.
- **1770 S Maple Grove Rd (174 units, ~28 du/ac, 2024)** — P&Z recommended DENIAL over the AIA 5-units/acre cap; City Council approved on appeal over the Airport Director's and staff's objection, with direction to re-evaluate the 1990s-era AIA density rules.
- **2025** policy re-evaluation of near-airport density / outdated noise-contour maps deferred pending the airport master plan update.
- These confirm the flight-path restriction is actively enforced and litigated in SE Boise.

## 11. Sources & vintage
- Parcels & zoning: Ada County GIS / AdaCountyGIS ArcGIS services (live, June 2026).
- Airport Influence Areas, Avigation Easements, Development Tracker: City of Boise GIS (Boise_GIS) ArcGIS services (live, June 2026; Development Tracker = active planning permits).
- PROPCODE definitions: Ada County Assessor GIS User Guide.
- AI-O zone rules: City of Boise §11-02-07 / AI-O overlay page; Ada County Title 8 Art. A.
- Narrative (Micron, landowners, denials, F-16 transition): BoiseDev, Idaho Statesman, City of Boise / Ada County planning, iflyboise.com, Gowen Field, FAA Part 150 — see the accompanying briefing for per-claim URLs and dates.
- Limitations: owner names unavailable (Public Records Act); PROPCODE land use is coarse (no SF/MF split); AIA zone boundaries are the City's mapped polygons (parcel-edge precision not independently surveyed).
