# Building the crosswalks (land-use buckets + zoning→MF-threat)

Two crosswalks turn raw local codes into the canonical vocabulary the deliverables render.
Both are **judgment calls** — document them in the decisions log.

## Canonical targets (do not invent new ones lightly)
From `scripts/palette.py`:
- **Land-use buckets**: Apartments / Multifamily · Single Family Residential · Vacant Land ·
  Commercial · Industrial / Service / Auto · Public / Airport / Institutional ·
  Agricultural / Rural · Utilities / Infrastructure · Other / Unclassified ·
  Use code n/a (degraded source).
- **Zoning categories**: Single-Family · Two-Family / Low-Density · Townhouse / Med-Density ·
  Multifamily Residential · Mixed Use · Manufactured Home · Commercial · Office ·
  Industrial / Business Park · Institutional / Governmental · Park / Recreation ·
  Agricultural / Rural · Planned / Overlay · Public / Airport / Institutional ·
  No public zoning (data gap).
- **Threat tiers**: High / Medium / Low / Unknown.

## Land-use crosswalk (`landuse_scheme.codes`)
Map every use code in the area to `[bucket, description]`. Rules of thumb:
- Multifamily/apartment classes → *Apartments / Multifamily*.
- Vacant platted lots/tracts + builder inventory → *Vacant Land* (this defines "vacant" —
  state it in `vacant_note`).
- Retail/office/hotel (often lumped) → *Commercial*; manufacturing → *Industrial*.
- Exempt/gov classes → *Public / Airport / Institutional* (and back this up with the
  `public_owner_regex`, since ownership is more reliable than the code for gov land).
- Codes you can't confidently map → leave to `default_bucket`; they're logged.

## Zoning → category + MF-threat (`zoning_crosswalk`)
For each district in each city, decide the category **and** the multifamily-supply threat
*if the land were vacant/redeveloped*:
- **High** — apartments buildable **BY-RIGHT** (e.g. R-MF, RM-30, R-5, MF-1/2; some
  mixed-use/MXU that permit MF by-right). Verify the key MF districts against the ordinance.
- **Medium** — MF only **conditionally**: townhouse, 3-4-family, mixed-use / TOD / corridor
  districts that allow residential at limited density or via special permitting.
- **Low** — base zone does **not** permit MF: single-family, duplex, commercial, office,
  industrial, civic, agricultural, manufactured-home.
- **Unknown** — planned-development / PUD / overlay / older "unit development" (CUD) — the
  allowed use depends on the negotiated plan; you must read it to know. Default unmapped
  codes here.

Watch-outs that bit the reference runs:
- An older "**community/planned unit development**" code (e.g. Euless `CUD`) reads like
  Commercial but is really Planned/Overlay = Unknown — miscategorizing it manufactures fake
  "commercial-zoned but MF" contradictions.
- **State-highway multi-use corridor** districts (e.g. `TX-10`, `TX-121`) allow residential
  → Mixed Use = Medium, not Unknown.
- Treat **land use and zoning as independent** — a Commercial-zoned parcel can legitimately
  carry MF use (legal nonconforming); don't "correct" one from the other.
- **Gov/airport-owned** land geometrically inside a residential zone is meaningless for MF
  supply — the pipeline already forces such parcels to Public/exempt, Low.

## Plain names (`zone_plain`)
Give every code a human label so tables and hover never show a bare code. Reuse the broad
set in the reference example as a starting point and add the local ones.
