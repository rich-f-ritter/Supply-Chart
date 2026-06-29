---
name: land-use-analysis
description: >-
  Produce a competitive land-use / zoning / multifamily-supply-threat analysis for ANY US
  property (apartment subject). Pulls county parcels + municipal zoning, classifies land use
  and zoning, reasons over genuinely-developable vacant land to rank the most concerning
  potential apartment-supply sites, and ships a self-documenting Excel workbook + a
  self-contained interactive HTML map viewer. Use when given a subject property/address and
  asked for a land use analysis, supply-threat study, vacant-land/competitive analysis, or
  "what could be built near here." Works in Claude Code and claude.ai (open internet assumed).
license: Proprietary
---

# Land Use Analysis

Generalizes a battle-tested DFW pipeline to **any US property**. The scripts are generic and
deterministic; the **locale-specific knowledge and the threat ranking are yours to discover
and reason** per run. Two canonical deliverables: a comprehensive **Excel workbook** (a reader
should understand the entire analysis from it alone) and a **self-contained interactive HTML
viewer** (Land Use / Zoning / Vacant Threats, satellite+streets, hover, PNG export).

## Setup
Open internet required. Install deps once:
`pip install requests shapely pyproj openpyxl urllib3`
Make a **run directory** for this property (e.g. `./<Subject Name>/`); all inputs/outputs live
there. Point scripts at it with `--root <dir>` (or `LUA_ROOT`); the config is `<root>/config.json`.

## The pipeline (one model-driven gap, the rest deterministic)

**Phase A — Resolve the subject (reason).** Geocode the address; verify the point against
satellite + parcel footprint and override mis-geocodes (common for apartments on ground-leased
/ master-planned land). Decide the analysis area with the user's intent: a radius (default 2 mi)
or a hand-drawn polygon.

**Phase B — Discover data sources (reason + web).** Find the county assessor/CAD **parcel**
layer and each municipal **zoning** layer; confirm fields; identify the **land-use code scheme**
(varies by state/county) and get each zoning ordinance. See `references/finding-sources.md`.
Be honest about coverage gaps — never fake a zone.

**Phase C — Author `config.json` (reason).** Encode subject, area, sources+field maps,
land-use crosswalk, and zoning→category→MF-threat crosswalk. Schema:
`references/config-schema.md`; crosswalk how-to: `references/crosswalks.md`; start from
`templates/config.example.json`. Canonical buckets/categories/colors are fixed in
`scripts/palette.py` — map local codes *into* them.

**Phase C.5 — Endpoint discovery → allowlist (REQUIRED before any pull).** Sandboxed
environments block network egress to domains that aren't allowlisted — the usual reason a run
silently fails. The county/city GIS hosts differ per property, so surface them and have the
user allowlist them first:
```
python <skill>/scripts/run_all.py endpoints --root .
```
This prints every host the run will contact (parcel CAD, city zoning, geocoders, viewer
basemap/CDN) and writes `in/required_domains.txt` plus `in/required_domains_wildcard.txt`.
**Present the list and ask the user to add it to the allowlist.** In **claude.ai**: *Settings →
Capabilities → "Code execution and file creation"* → turn **Allow network egress** ON → under
**Additional allowed domains**, Add each entry (the "Package managers only" mode still honors
these). Wildcards work — **prefer the wildcard set** (e.g. `*.arcgis.com`, `*.arcgisonline.com`,
`*.census.gov`): many cities' zoning layers live on rotating `servicesN.arcgis.com` subdomains,
so one `*.arcgis.com` entry saves re-editing the allowlist per property. (Package installs —
pypi/`pip` — are already covered by "Package managers only.") In **Claude Code**: the sandbox
network allowlist in settings.

**Allowlist changes only take effect in a NEW session** — the sandbox reads the network policy
once, at session start. So `endpoints` also writes `in/next_session_prompt.txt`: a ready-to-paste
prompt that rebuilds this exact run. **Give the user both** (a) the domain list to add, and (b)
that next-session prompt to paste into a fresh chat. If egress is already live in the current
session, you can instead record consent and continue here:
```
echo ok > in/allowlist_ok        # the flag the pull stage checks for
```
Do not run `pull` until `in/allowlist_ok` exists. If a pull fails with `host_not_allowed`, the
domain wasn't allowlisted (or you're still in the pre-allowlist session) — re-list, get it (or
its `*.` wildcard) added, and start a fresh session.

**Reachability pre-flight (allowlisted ≠ reachable).** A host can pass the allowlist yet be
unreachable from the sandbox's egress IPs — you'll see HTTP 503 "upstream connect error …
connection timeout" (or a hang), **not** `host_not_allowed`. `pull` runs `preflight` first and
**halts** if any source is `blocked` (allowlist) or `unreachable` (origin down). A city's
self-hosted ArcGIS box (e.g. `maps.mycity.com`) is the most common offender. When a host is
unreachable, **substitute a reachable equivalent** — prefer a county/state layer on
`*.arcgis.com` or the county GIS server over a single city's server, and many states publish a
statewide cadastral. Re-author `config.json`, re-run `endpoints` if the host changed, log the
swap in the decisions log, then retry.

**Phase D — Pull + classify (deterministic).** From the run dir, after the allowlist is in place:
```
python <skill>/scripts/run_all.py pull --root .          # preflight + parcels + zoning (refuses without in/allowlist_ok)
python <skill>/scripts/run_all.py classify --root .       # land-use + zoning + threat join
python <skill>/scripts/run_all.py prep --root .           # vacant candidates + developable set
```
Inspect the printed bucket/threat/UNMAPPED counts and `Tables/unmapped_codes.json`; fix the
crosswalk and re-run `classify` if codes went unmapped.

*Pulling is solved — don't hand-roll it.* `pull_parcels`/`pull_zoning` use POST + objectId
paging (`scripts/arcgis.py`): no dependency on `returnCountOnly`, immune to GET URL-length
limits (the two things that break naive pullers on big statewide/county layers), and they
**checkpoint to `in/_ck/` so a large or interrupted pull resumes on re-run**. So: wire the
discovered endpoint into `config.json` and run `pull` — do **not** pre-probe count or write a
custom downloader (that's where past runs burned their whole budget). Dense metro (thousands of
condo parcels)? It still completes via paging+resume; if it's slow, set `"simplify": 0.00002` on
the source and/or trim the radius. If a pull dies mid-way, just run `pull` again — it continues.

**Phase E — Reason the vacant threat (reason — the core output).** This is a judgment call,
**not a score** (a formula surfaces HOA commons, slivers, floodplain, gov land as false
"threats"). Read `in/vacant_candidates.json`, reason over each genuinely-developable candidate
(developable? real MF basis / entitlement? developer vs HOA owner? right submarket?), do
targeted research where it changes the call, and write `in/reasoned_ranking.json` +
`Tables/decisions_log.md`. Full guidance + schemas: `references/threat-reasoning.md`.

**Phase F — Finalize + deliver (deterministic).**
```
python <skill>/scripts/run_all.py deliver --root .       # = finalize_topN, build_viewer, build_workbook
```
Outputs in the run dir: `<Subject> - Land Use Analysis.xlsx`, `<Subject> - Land Use Viewer.html`,
and `Tables/*.csv`. The workbook now includes the **actual data** — a **Parcels** sheet (every
classified parcel with land use + zoning + threat + vacant flags) and a **Vacant - Developable**
sheet (the flagged threat set) — alongside the grouping/methodology sheets and the reasoned
Top-N. The viewer **bundles Leaflet** (from `assets/vendor/`, no CDN) and **dedupes coincident
condo footprints** so each physical footprint draws once (avoids the stacked-translucency
"two shades of one color" artifact). Basemap raster tiles still need `*.arcgisonline.com`
reachable at view time; if blocked/offline the map degrades gracefully (neutral canvas +
"offline basemap" badge, a "None" basemap button) and parcels still render. (Large/dense metros:
`python <skill>/scripts/build_viewer.py --served --root .` for sidecar data + a local server.)

`run_all.py all --root .` walks the whole pipeline but **stops at two gates**: the allowlist
gate (after `endpoints`, until `in/allowlist_ok` exists) and the reason gate (after `prep`,
until `in/reasoned_ranking.json` exists). Resume past each with `pull` / `deliver`.

## Non-negotiables (learned from prior runs)
- **Document every subjective decision** — verified point, area choice, sources used/rejected,
  crosswalk calls, the vacant definition, pre-filters, ranking rationale — in
  `Tables/decisions_log.md` (→ the workbook's Assumptions & Decisions sheet). Log ambiguous/
  unmapped codes; don't silently bucket them.
- **Threat ranking is reasoned**, with a rationale per parcel; scripts only prepare candidates.
- **Honest gaps** — areas with no public parcel/zoning data are labeled, never fabricated;
  prioritize full coverage near the subject.
- **Allowlisted ≠ reachable** — verify each endpoint actually responds (`preflight`) before
  pulling; substitute any allowed-but-unreachable host (prefer county/state `*.arcgis.com` over
  a city's self-hosted server) and log the substitution in the decisions log.
- **Land use and zoning are independent layers** — don't "correct" one from the other.

## Files
- `scripts/` — `config.py`, `geo.py`, `palette.py` (canonical vocab+colors), `arcgis.py`
  (robust POST/paging/resume fetcher), `list_endpoints.py`, `preflight.py`, `pull_parcels.py`,
  `pull_zoning.py`, `classify.py`, `prepare_candidates.py`, `finalize_topN.py`,
  `build_viewer.py`, `build_workbook.py`, `run_all.py`.
- `references/` — `config-schema.md`, `finding-sources.md`, `crosswalks.md`, `threat-reasoning.md`.
- `templates/` — `config.example.json`, `reasoned_ranking.example.json`.
