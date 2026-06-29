# Finding parcel + zoning data for ANY US property

There is no single national parcel/zoning API. You discover the right county assessor
(CAD) parcel layer and the municipal zoning layer(s) per run, confirm their fields, and
record them in `config.json`. Open internet + web search/fetch are assumed.

## 1. Locate the subject + jurisdictions
- Geocode the address (Census `geocoding.geo.census.gov/geocoder/locations/onelineaddress`,
  free; or ArcGIS World GeocodeServer). Confirm the point against satellite imagery and the
  parcel footprint — override mis-geocodes (see config-schema `subject`).
- Identify the **county** (for parcels) and the **city/municipality** (for zoning). Within a
  radius you often span several cities + unincorporated county — each city is its own zoning
  source, and a parcel can have no city zoning (data gap).

## Prefer reachable, broadly-hosted layers (allowlisted ≠ reachable)
Sandboxed runs have hit hosts that pass the allowlist but are **unreachable** from the sandbox's
egress IPs (HTTP 503 "upstream connect error … connection timeout"). The repeat offender is a
**single city's self-hosted ArcGIS server** (e.g. `maps.<city>.com`). So, when you have a choice:
- Prefer a **county or statewide** layer hosted on **`*.arcgis.com` (ArcGIS Online)** or the
  county GIS server over a city's own box. Many states publish a **statewide cadastral / parcel**
  layer (e.g. a state DOR/GIS clearinghouse) — one reachable source can cover the whole radius.
- Pick endpoints so the fewest distinct hosts are needed, and favor hosts you've seen respond.
- The `preflight` step probes every chosen endpoint before pulling; if one is unreachable,
  substitute and **log the swap** in the decisions log. Don't ship a config that depends on a
  host you haven't confirmed responds.

## 2. Find the county parcel layer (ArcGIS REST)
Most US county assessors publish an ArcGIS FeatureServer/MapServer parcel layer.
- Search: `"<county> county" parcels arcgis rest services` or `<county> CAD parcel
  FeatureServer`, and check the county GIS / appraisal-district site, plus
  `<org>.maps.arcgis.com`. ArcGIS Hub and `services*.arcgis.com` host many.
- Verify the layer:
  - Open the layer URL (ends with a number, e.g. `/FeatureServer/0`); read its **Fields**.
  - Test a query: append `/query?where=1=1&outFields=*&resultRecordCount=3&f=json`.
  - Confirm it has **owner**, a **land-use / property-class code**, geometry, and ideally a
    unique parcel id (for `dedupe_field`). Map these into `parcel_sources[].fields`.
- If the county only offers a download (Shapefile/GeoJSON) or a non-ArcGIS API: fetch it,
  clip to the area, and write `in/parcels_all.geojson` in the normalized shape yourself
  (skip `pull_parcels`).
- **Cross-county areas:** add one source per county. If a neighboring county's layer is
  missing or lacks a use code, include it with `data_confidence:"degraded"` (geometry+owner
  only) and document the gap — never fabricate land use.

## 3. Identify the land-use code scheme (it varies by state/county)
The use-code field and its meaning differ everywhere:
- Texas CADs: `State_Use_` (SPTB letter codes A/B/C1/F1/X...).
- Many counties: a numeric property-class / `IMPLOCALCODE` / `LANDCODE` / `PropertyUseCode`.
Get the official code list (county data dictionary / appraisal manual) and map each code to a
canonical bucket in `landuse_scheme.codes`. When in doubt about whether a code is "vacant,"
prefer the county's own definition; log ambiguous codes rather than silently bucketing.

## 4. Find municipal zoning layers
Zoning is even more fragmented (per-city, and many places publish none):
- Search `"<city>" zoning arcgis rest FeatureServer` and the city GIS / open-data portal.
- **Verify it's the right city** — generic AGOL names collide (e.g. a "Bedford" layer that's
  Bedford, VA, or a regional "Zoning" layer that's actually another city). Check a couple of
  polygons fall on the city and the code values match the city's ordinance.
- Read the field holding the zone code (`code_field`) and a description field if present.
- Get each city's **zoning ordinance** to build the zone→category→threat crosswalk
  (see crosswalks.md). Confirm the key MF districts against the actual ordinance text.
- Cities/areas with no public zoning GIS → leave them out of `zoning_sources`; they render as
  the data-gap class. Prioritize full coverage **near the subject** (aim for 0% gap within
  the inner part of the area) and be explicit about gaps farther out.

## 5. Record everything
Put every endpoint in `config.json` (it flows into the workbook's Data Sources sheet) and
note every judgment (which sources, which were rejected and why, coverage gaps) in the
decisions log. Reproducibility + honesty about gaps is a hard requirement.
