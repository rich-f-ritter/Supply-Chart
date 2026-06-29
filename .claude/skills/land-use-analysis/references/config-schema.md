# config.json — the per-run brain of the analysis

Everything locale-specific lives in one `config.json` in the run directory. The Python
scripts are generic; they read this file. You (Claude) author it during the discovery +
reasoning phases. Start from `templates/config.example.json`.

## Top-level keys

### `subject` (required)
- `name` — display name (drives titles + output filenames).
- `address` — mailing address as given.
- `lat`, `lon` — the **verified** point (WGS84). Geocode the address, then sanity-check
  against the satellite/parcel footprint. If the mailing address mis-geocodes (very common
  for apartments on ground-leased / master-planned land), override to the true footprint and
  explain in `location_note`. This point anchors the map pin and **every distance**.
- `location_note` — one sentence on how the point was verified / why it was overridden.
- `mailing_geocode` — optional raw geocoder result, kept for the audit trail.

### `analysis_area` (required)
- `mode`: `"radius"` or `"polygon"`.
- radius: `radius_mi` (number). A true geodesic circle is built around the subject.
- polygon: `polygon` = `[[lon,lat], ...]` (a hand-drawn area; first≈last point). Use this
  when the user wants a submarket shape rather than a circle.
- `note` — describe the area for the workbook.

### `parcel_sources` (required, 1+)
County assessor / CAD parcel layers (ArcGIS REST `.../query`). Each:
- `name`, `county`, `url` (must end in `/query`).
- `fields` — map **normalized → source** field names. Normalized keys the pipeline uses:
  `account, owner, situs, landuse_code, legal, impr_value, total_value, year_built`.
  Only `owner` + geometry are strictly required; omit keys the source lacks.
- `dedupe_field` — a source field that uniquely IDs a parcel (de-dups quadtree overlap).
- `data_confidence` — `"full"` normally; `"degraded"` if the source lacks a use code
  (those parcels render as a muted "use code n/a" bucket and are excluded from threats).
- `verify_ssl` — set `false` only for a known-good source with a bad cert chain.
- `where` — optional SQL filter (default `"1=1"`; e.g. `"CO_NO=60"` to pin a county on a
  statewide layer).
- `page` — objectId batch size (default 1000; lower to ~500 for very heavy geometry).
- `simplify` — server-side `maxAllowableOffset` in **degrees** (e.g. `0.00002` ≈ 2 m) to thin
  geometry on dense condo metros. Omit for exact acreage; set it only if a pull is huge/slow.
- `format` — `"geojson"` (default) or `"esrijson"` for old ArcGIS Server layers that don't emit
  GeoJSON (the fetcher converts esri JSON itself).

The pull uses **POST + objectId paging** (`scripts/arcgis.py`) — it does NOT depend on
`returnCountOnly` and is immune to GET URL-length limits, so you generally just wire the
endpoint in and run; no pre-probing needed. It **checkpoints** id-batches under `in/_ck/`, so a
big/interrupted pull **resumes** on re-run. Do not hand-build a downloader.

If a county has **no ArcGIS layer**, skip `pull_parcels` and hand-write
`in/parcels_all.geojson` yourself in the normalized shape (FeatureCollection; each
feature's `properties` = the normalized keys + `data_confidence`, `source`).

### `zoning_sources` (optional, 0+)
Municipal/county zoning layers (ArcGIS REST). Each: `jurisdiction`, `url`, `code_field`,
`desc_field` (optional), `verify_ssl`, `where`. Areas with no source become the labeled
**data-gap** class — never fake a district. Note gaps in the decisions log.

### `landuse_scheme` (required)
- `name` — e.g. "Texas State Use (SPTB)" or "El Paso County IMPLOCALCODE/LANDCODE".
- `codes` — map **raw use code → `[bucket, description]`**. `bucket` MUST be one of the
  canonical buckets in `scripts/palette.py` (`LANDUSE_COLOR`).
- `vacant_buckets` — which buckets count as vacant (usually `["Vacant Land"]`).
- `vacant_note` — plain-English definition for the workbook.
- `default_bucket` — bucket for unmapped codes (`"Other / Unclassified"`); unmapped codes
  are logged to `Tables/unmapped_codes.json`.
- `public_owner_regex` — owner-name regex that routes gov/airport/school/church land to the
  **Public / Airport / Institutional** bucket regardless of use code.

### `zoning_crosswalk` (required if any zoning_sources)
`{ jurisdiction: { zone_code: [category, threat] } }`. `category` ∈ canonical
`ZONE_CATEGORY_COLOR`; `threat` ∈ `High|Medium|Low|Unknown`. Convenience pairs in
`palette.THREAT_SHORTCUTS`. Unmapped codes default to `Planned / Overlay` + `Unknown`
and are logged.

### `zone_plain` (recommended)
`{ zone_code: "Plain English name" }` so tables/hover never show bare codes.

### Tuning + owner heuristics
- `nondev_owner_regex` — owners that are NOT private competitive MF developers
  (gov/HOA/POA/condo/church/school/utility). Make it **truncation-tolerant** if the source
  truncates owner names (e.g. TAD's 26 chars: `HOMEOWN`, not `HOMEOWNERS`).
- `company_owner_regex` — LLC/LP/etc., used only to label owner_type.
- `min_acres` (default 1.0), `min_compactness` (Polsby-Popper, default 0.16).
- `data_gap_label` (default "No public zoning (data gap)").
- `crs_local` — leave `null` to auto-pick the subject's UTM zone (good nationwide); pin an
  EPSG (e.g. a State Plane) only if you want exact local acreage.
- `branding` — `primary` + `accent` hex for the viewer/workbook.

## Validate before running
- Every `bucket` in `landuse_scheme.codes` is a key in `palette.LANDUSE_COLOR`.
- Every `category` in `zoning_crosswalk` is a key in `palette.ZONE_CATEGORY_COLOR`.
- `subject.lat/lon` actually sits inside `analysis_area`.
