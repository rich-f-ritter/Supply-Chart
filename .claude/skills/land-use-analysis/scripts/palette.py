"""Canonical, locale-INDEPENDENT vocabulary for the land-use analysis: the fixed set
of land-use buckets, zoning categories, MF-threat tiers, their approved colors, and
the threat definitions. Every locale's raw codes are mapped INTO these canonical names
by the per-run config (config.json -> landuse_scheme / zoning_crosswalk); this module is
what keeps a Texas run and a Colorado run looking and reading the same.

Do not add locale-specific codes here. Add new *canonical categories* only if a place
genuinely has a land-use/zoning concept none of these cover, and give it a color.
"""

# ── canonical land-use buckets -> color ───────────────────────────────────────
LANDUSE_COLOR = {
    "Apartments / Multifamily":          "#D62728",
    "Single Family Residential":         "#FFF6B3",
    "Vacant Land":                       "#2E8B57",
    "Commercial":                        "#1F77B4",
    "Industrial / Service / Auto":       "#7F7F7F",
    "Public / Airport / Institutional":  "#6BAED6",
    "Agricultural / Rural":              "#9C8A57",
    "Utilities / Infrastructure":        "#525252",
    "Other / Unclassified":              "#D9D9D9",
    "Use code n/a (degraded source)":    "#CDB7D6",
}
LANDUSE_ORDER = list(LANDUSE_COLOR)

# ── canonical zoning categories -> color ──────────────────────────────────────
ZONE_CATEGORY_COLOR = {
    "Single-Family Residential":             "#FFF6B3",
    "Two-Family / Low-Density Residential":  "#FCDE9C",
    "Townhouse / Med-Density Residential":   "#FDBF6F",
    "Multifamily Residential":               "#D62728",
    "Mixed Use":                             "#FF8C42",
    "Manufactured Home":                     "#E5C8A0",
    "Commercial":                            "#1F77B4",
    "Office":                                "#9467BD",
    "Industrial / Business Park":            "#525252",
    "Institutional / Governmental":          "#6BAED6",
    "Park / Recreation":                     "#2CA02C",
    "Agricultural / Rural":                  "#9C8A57",
    "Planned / Overlay":                     "#B399D4",
    "Public / Airport / Institutional":      "#7FA8C9",
    "No public zoning (data gap)":           "#E3E3E3",
}
ZONE_CATEGORY_ORDER = list(ZONE_CATEGORY_COLOR)

# ── MF-supply threat tiers ────────────────────────────────────────────────────
# threat = multifamily SUPPLY threat if the parcel were vacant / redeveloped.
THREAT_COLOR = {"High": "#D62728", "Medium": "#FF8C42", "Low": "#FFF2B3", "Unknown": "#B399D4"}
THREAT_ORDER = ["High", "Medium", "Low", "Unknown"]
THREAT_DEF = {
    "High":    "Multifamily can be built BY-RIGHT - no special approval needed.",
    "Medium":  "Multifamily possible but CONDITIONAL - townhouse / 3-4-family / mixed-use / "
               "TOD / corridor; limited density or special permitting.",
    "Low":     "Multifamily NOT permitted by base zoning (single-family, duplex, commercial, "
               "industrial, civic, or agricultural).",
    "Unknown": "Planned-development / overlay / older unit-development district - allowed use "
               "depends on the negotiated plan; must read it to determine MF potential.",
}
# Short canonical->(category, threat) shortcuts authors can reuse when writing a
# config's zoning_crosswalk by hand (purely a convenience; the config stores the
# resolved [category, threat] pairs, not these names).
THREAT_SHORTCUTS = {
    "SF":   ("Single-Family Residential", "Low"),
    "TWO":  ("Two-Family / Low-Density Residential", "Low"),
    "TH":   ("Townhouse / Med-Density Residential", "Medium"),
    "MF":   ("Multifamily Residential", "High"),
    "MIX":  ("Mixed Use", "Medium"),
    "MIXH": ("Mixed Use", "High"),
    "COM":  ("Commercial", "Low"),
    "OFF":  ("Office", "Low"),
    "IND":  ("Industrial / Business Park", "Low"),
    "INST": ("Institutional / Governmental", "Low"),
    "PARK": ("Park / Recreation", "Low"),
    "AG":   ("Agricultural / Rural", "Low"),
    "MH":   ("Manufactured Home", "Low"),
    "PLN":  ("Planned / Overlay", "Unknown"),
}

DATA_GAP_LABEL = "No public zoning (data gap)"


def landuse_color(bucket):
    return LANDUSE_COLOR.get(bucket, "#D9D9D9")


def zone_color(category):
    return ZONE_CATEGORY_COLOR.get(category, "#E3E3E3")


def threat_color(threat):
    return THREAT_COLOR.get(threat, "#B399D4")
