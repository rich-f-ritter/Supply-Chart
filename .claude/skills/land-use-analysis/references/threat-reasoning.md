# The reasoned vacant-threat assessment (the core analytical output)

The "most concerning vacant parcels" ranking is a **reasoned judgment, not a deterministic
score.** `prepare_candidates.py` does only mechanical prep; you reason over its output and
write `in/reasoned_ranking.json`, which `finalize_topN.py` renders.

Why not a formula: a mechanical score (threat tier × acres × proximity) surfaces false
positives — HOA common areas / detention ponds / greenbelts zoned PD, remnant slivers,
floodplain, non-developable tracts, gov/airport land. Judgment is required: is the parcel
genuinely developable, is the PD actually entitled for MF, is the owner a developer vs an
HOA, is it the relevant submarket?

## Inputs
- `in/vacant_candidates.json` — clusters after mechanical filters (compactness, owner rules,
  data-gap, min acreage), each with: acres, parcel_count, min_dist_mi, owner, owner_type,
  zoning (`juris:code`), `zoning_threat_baseline`, compactness, location, lat/lon, parcels.
- The classified parcels + the zoning crosswalk + your knowledge of the submarket. Do
  **targeted research** where it changes the call (read a PD/PUD concept plan, check whether a
  corridor is actively adding apartments, confirm an owner is a developer vs an HOA/SPE).

## How to reason each candidate
For each plausible candidate ask:
1. **Developable?** Real buildable site, or a remnant/detention/greenbelt/floodplain/access
   sliver that slipped the filters? Drop the latter.
2. **MF basis?** By-right MF zoning (High) is the strongest threat. For PD/overlay (Unknown),
   what does the negotiated plan actually entitle? For corridor/mixed (Medium), how real is
   residential? For Low base zoning, is there genuine **rezoning optionality** given momentum
   (large single-owner tract on an active corridor) → call it "Low-Watch", not High.
3. **Owner.** Developer / homebuilder / investor SPE raises concern; HOA / church / school /
   gov / utility removes it.
4. **Submarket + distance.** Far parcels in a different submarket (e.g. >3 mi across a barrier)
   are de-prioritized even if large — say so.

## Output: `in/reasoned_ranking.json`
```json
{
  "headline": "one-paragraph supply-threat conclusion for the workbook + viewer",
  "ranking": [
    { "match": {"account": "12345"},                      // OR owner_prefix + location_prefix
      "reasoned_threat": "High|Medium|Low|Low-Watch|Unknown",
      "rationale": "why it is (or isn't) a concerning MF-supply site; cite the plan/corridor/owner" }
  ]
}
```
- Rank in concern order (most concerning first). 10 is typical; fewer is fine if that's the
  honest set. `reasoned_threat` is free text (so "Low-Watch" is allowed); the workbook color-
  keys on the leading word.
- Prefer `match.account` (exact) when you can; otherwise `owner_prefix` + `location_prefix`.

## Always write the decisions log
Write `Tables/decisions_log.md` capturing every judgment call: subject point/override,
analysis-area choice, data sources used + rejected + coverage gaps, land-use source &
granularity, zoning crosswalk corrections, the vacant definition, the mechanical pre-filters,
and the reasoning behind the ranking + headline. This feeds the workbook's "Assumptions &
Decisions" sheet and is a hard requirement — these analyses must be auditable and reproducible.
```
# Land Use Analysis — <Subject> — Decisions Log
## 1. Subject location (verified point)
## 2. Analysis area
## 3. Parcel data source(s) + any degraded/missing coverage
## 4. Land-use source & granularity
## 5. Zoning sources + honest gaps
## 6. Zoning crosswalk corrections
## 7. "Vacant" definition
## 8. Vacant-threat assessment — reasoned, with the mechanical pre-filters listed
## 9. Sources & vintage
```
