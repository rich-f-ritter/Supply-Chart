#!/usr/bin/env python3
"""
build_supply_chart.py — Automate the 5-mile new-construction Supply Chart.

Reads the three market exports (CoStar property roster, CoStar Data Analytics
time series, RealPage roster) for a subject property's 5-mile radius, reconciles
them into a single competitive-supply roster, buckets each property by lifecycle
stage, pins delivery quarters against the CoStar quarterly deliveries series, and
writes a formatted workbook that mirrors the reference "Supply Chart" template.

See SKILL.md for the full methodology. This script handles the *mechanical* parts
(parsing, address matching, quarter pinning, formatting). Judgment fields that the
script cannot know — Proximity (miles) and verified lease-up rent/occupancy from
HelloData — are intentionally left blank / flagged for the analyst.

Usage:
    python build_supply_chart.py \
        --subject-name "Canyon Ridge" \
        --subject-address "2552 E Gowen Rd" \
        --costar-roster   examples/canyon_ridge/CoStar_5mi_50unit_properties.xlsx \
        --costar-analytics examples/canyon_ridge/CoStar_5mi_Data_Analytics.xlsx \
        --realpage        examples/canyon_ridge/Realpage_5mi.xlsx \
        --out             output/Canyon_Ridge__Supply_Chart.xlsx
"""
from __future__ import annotations

import argparse
import datetime as dt
import re
import sys
from dataclasses import dataclass, field
from typing import Optional

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.formula import ArrayFormula

# --------------------------------------------------------------------------- #
# Tunable thresholds (documented in SKILL.md)
# --------------------------------------------------------------------------- #
DEFAULT_STABILIZATION_TARGET = 0.95
# A delivered property is "stabilized" once occupancy >= this OR it has been open
# longer than LEASEUP_WINDOW_QTRS quarters. Otherwise it is still "leasing up".
STABILIZED_OCC = 0.90
LEASEUP_WINDOW_QTRS = 6          # ~18 months to lease up a new building
# How far back a delivery counts as "new construction" worth charting.
NEW_CONSTRUCTION_LOOKBACK_YEARS = 4

# --------------------------------------------------------------------------- #
# Styling (sampled from the reference template)
# --------------------------------------------------------------------------- #
NAVY = "FF1F3864"
BLUE = "FF2E75B6"
GRAY = "FFF2F2F2"
WHITE = "FFFFFFFF"
YELLOW = "FFFFFF00"

THIN = Side(style="thin", color="FFBFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def fill(color: str) -> PatternFill:
    return PatternFill("solid", fgColor=color)


def font(bold=False, size=9, color="FF000000"):
    return Font(name="Calibri", bold=bold, size=size, color=color)


CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center")
RIGHT = Alignment(horizontal="right", vertical="center")


# --------------------------------------------------------------------------- #
# Data model
# --------------------------------------------------------------------------- #
QUARTERS = ["Q1", "Q2", "Q3", "Q4"]


def quarter_index(year: int, q: int) -> int:
    """Absolute quarter index for ordering / arithmetic."""
    return year * 4 + (q - 1)


def fmt_quarter(year: int, q: int) -> str:
    return f"Q{q} {year}"


@dataclass
class Prop:
    name: str
    address: str
    units: Optional[int] = None
    year_built: Optional[int] = None
    construction_begin: Optional[str] = None
    status_raw: str = ""               # e.g. Existing / Stabilized / Pre-Planned
    occupancy: Optional[float] = None
    eff_rent: Optional[float] = None
    owner: Optional[str] = None
    stories: Optional[int] = None
    style: Optional[str] = None        # RealPage property style
    sources: set = field(default_factory=set)
    # derived
    est_delivery: Optional[str] = None     # "Q2 2023" or None (undated pipeline)
    deliv_year: Optional[int] = None
    deliv_q: Optional[int] = None
    bucket: Optional[str] = None
    prop_type: Optional[str] = None
    notes: list = field(default_factory=list)

    def note(self, txt: str):
        if txt and txt not in self.notes:
            self.notes.append(txt)


# --------------------------------------------------------------------------- #
# Address normalization & matching
# --------------------------------------------------------------------------- #
_STREET_NOISE = re.compile(
    r"\b(apartments?|apts?|townhomes?|the|at|on|of)\b", re.I)


def addr_key(address: str) -> Optional[str]:
    """Build a match key from a street address: leading number + first word.

    Handles ranges ("2410-2490 W Canal St" -> 2410) and directionals.
    """
    if not address:
        return None
    a = str(address).strip().lower()
    m = re.match(r"(\d+)", a)
    if not m:
        return None
    number = m.group(1)
    rest = a[m.end():].strip(" -")
    # drop a trailing range number like "-2490"
    rest = re.sub(r"^\d+\s*", "", rest)
    toks = [t for t in re.split(r"[\s,]+", rest) if t]
    # skip leading directional (n/s/e/w)
    dirs = {"n", "s", "e", "w", "ne", "nw", "se", "sw",
            "north", "south", "east", "west"}
    sig = next((t for t in toks if t not in dirs), toks[0] if toks else "")
    return f"{number} {sig}"


def name_key(name: str) -> str:
    n = _STREET_NOISE.sub(" ", str(name or "").lower())
    return re.sub(r"\s+", " ", n).strip()


# --------------------------------------------------------------------------- #
# Parsers
# --------------------------------------------------------------------------- #
def _hdr_map(ws):
    headers = {}
    for j, c in enumerate(next(ws.iter_rows(min_row=1, max_row=1)), 1):
        if c.value is not None:
            headers[str(c.value).strip()] = j
    return headers


def _int(v):
    try:
        if v in (None, "", "-", "—"):
            return None
        return int(round(float(v)))
    except (TypeError, ValueError):
        return None


def _float(v):
    try:
        if v in (None, "", "-", "—"):
            return None
        return float(v)
    except (TypeError, ValueError):
        return None


def parse_costar_roster(path) -> list[Prop]:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    h = _hdr_map(ws)

    def col(row, key):
        j = h.get(key)
        return row[j - 1] if j else None

    out = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        name = col(row, "Property Name")
        if not name:
            continue
        p = Prop(
            name=str(name).strip(),
            address=str(col(row, "Property Address") or "").strip(),
            units=_int(col(row, "Number of Units")),
            year_built=_int(col(row, "Year Built")),
            construction_begin=(str(col(row, "Construction Begin")).strip()
                                if col(row, "Construction Begin") else None),
            status_raw=str(col(row, "Building Status") or "").strip(),
            owner=(str(col(row, "Owner Name")).strip()
                   if col(row, "Owner Name") else None),
            stories=_int(col(row, "Number of Stories")),
        )
        p.sources.add("CoStar")
        out.append(p)
    return out


def parse_realpage(path) -> list[Prop]:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    h = _hdr_map(ws)

    def col(row, key):
        j = h.get(key)
        return row[j - 1] if j else None

    out = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        name = col(row, "Name")
        if not name:
            continue
        p = Prop(
            name=str(name).strip(),
            address=str(col(row, "Address") or "").strip(),
            units=_int(col(row, "Total Units")),
            year_built=_int(col(row, "Year Built")),
            status_raw=str(col(row, "Property Status") or "").strip(),
            occupancy=_float(col(row, "Occupancy")),
            eff_rent=_float(col(row, "Effective Rent")),
            owner=(str(col(row, "Property Owner")).strip()
                   if col(row, "Property Owner") else None),
            stories=_int(col(row, "Stories")),
            style=(str(col(row, "Property Style")).strip()
                   if col(row, "Property Style") else None),
        )
        p.sources.add("RealPage")
        out.append(p)
    return out


_QRE = re.compile(r"(\d{4})\s*Q([1-4])")


def parse_costar_analytics(path):
    """Return (latest_inventory_units, {(year,q): delivered_units}, latest_label)."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    h = _hdr_map(ws)
    pcol = h.get("Period")
    inv_col = h.get("Inventory Units")
    del_col = h.get("Deliveries Units")
    deliveries = {}
    latest_inv = None
    latest_label = None
    for row in ws.iter_rows(min_row=2, values_only=True):
        period = row[pcol - 1] if pcol else None
        if not period:
            continue
        m = _QRE.search(str(period))
        if not m:
            continue
        year, q = int(m.group(1)), int(m.group(2))
        d = _int(row[del_col - 1]) if del_col else None
        if d:
            deliveries[(year, q)] = d
        if latest_inv is None:  # rows are newest-first
            latest_inv = _int(row[inv_col - 1]) if inv_col else None
            latest_label = str(period).strip()
    return latest_inv, deliveries, latest_label


# --------------------------------------------------------------------------- #
# Reconciliation
# --------------------------------------------------------------------------- #
def reconcile(costar: list[Prop], realpage: list[Prop]) -> list[Prop]:
    """Merge the two rosters by address (primary) / name (fallback)."""
    merged: dict[str, Prop] = {}
    by_name: dict[str, str] = {}

    def register(p: Prop):
        k = addr_key(p.address) or ("name:" + name_key(p.name))
        if k in merged:
            return merge_into(merged[k], p)
        # name fallback for missing/typo'd addresses
        nk = name_key(p.name)
        if nk in by_name and by_name[nk] in merged:
            return merge_into(merged[by_name[nk]], p)
        merged[k] = p
        by_name[nk] = k
        return p

    def merge_into(base: Prop, other: Prop):
        base.sources |= other.sources
        # Units: keep CoStar's; flag if they differ materially
        if other.units and base.units and other.units != base.units:
            lo, hi = sorted((base.units, other.units))
            if hi - lo >= 3:  # ignore +/-2 unit noise
                base.note(f"Units: CoStar {base.units if 'CoStar' in base.sources else hi} "
                          f"vs RealPage {other.units if 'RealPage' in other.sources else lo}")
        # Occupancy / rent: take from whichever has them (RealPage today)
        if base.occupancy is None and other.occupancy is not None:
            base.occupancy = other.occupancy
        if base.eff_rent is None and other.eff_rent is not None:
            base.eff_rent = other.eff_rent
        # Style / stories / owner backfill
        base.style = base.style or other.style
        base.stories = base.stories or other.stories
        base.owner = base.owner or other.owner
        # Year built conflict
        if other.year_built and base.year_built and other.year_built != base.year_built:
            base.note(f"Year built: {min(base.year_built, other.year_built)}/"
                      f"{max(base.year_built, other.year_built)}")
        base.year_built = base.year_built or other.year_built
        base.construction_begin = base.construction_begin or other.construction_begin
        # Prefer a forward-looking status (Pre-Planned/Under Construction)
        if _is_pipeline(other.status_raw) and not _is_pipeline(base.status_raw):
            base.status_raw = other.status_raw
        return base

    for p in costar:
        register(p)
    for p in realpage:
        register(p)
    return list(merged.values())


def _is_pipeline(status: str) -> bool:
    s = (status or "").lower()
    return any(t in s for t in ("pre-plan", "preplan", "proposed",
                                "under construction", "planned", "lease-up",
                                "lease up"))


# --------------------------------------------------------------------------- #
# Delivery-quarter pinning + bucketing
# --------------------------------------------------------------------------- #
def pin_delivery(p: Prop, deliveries: dict, as_of: tuple[int, int]):
    """Assign an est. delivery quarter using the CoStar quarterly deliveries series.

    Strategy: among quarters within +/-1 year of the property's year-built, pick
    the one whose delivered-unit count exactly matches the property's units; if no
    exact match, pick the closest by unit count; else fall back to "Q? <year>".
    """
    if not p.year_built or not p.units:
        return
    candidates = [(y, q, u) for (y, q), u in deliveries.items()
                  if abs(y - p.year_built) <= 1]
    if candidates:
        exact = [(y, q) for (y, q, u) in candidates if u == p.units]
        if exact:
            y, q = sorted(exact)[0]
        else:
            y, q, _ = min(candidates, key=lambda t: abs(t[2] - p.units))
        p.deliv_year, p.deliv_q = y, q
        p.est_delivery = fmt_quarter(y, q)
    else:
        p.deliv_year = p.year_built
        p.est_delivery = f"Q? {p.year_built}"


def classify(p: Prop, as_of: tuple[int, int], target: float):
    """Assign one of the four lifecycle buckets."""
    as_of_idx = quarter_index(*as_of)
    delivered = (p.deliv_year is not None and p.deliv_q is not None
                 and quarter_index(p.deliv_year, p.deliv_q) <= as_of_idx)

    if _is_pipeline(p.status_raw) and "under construction" in p.status_raw.lower():
        p.bucket = "UNDER CONSTRUCTION"
    elif _is_pipeline(p.status_raw):
        # Pre-Planned / Proposed (no real delivery date)
        p.bucket = "PROPOSED"
    elif not delivered and p.year_built and p.year_built >= as_of[0]:
        # has a construction begin / near-term year but not yet delivered
        p.bucket = "UNDER CONSTRUCTION"
    elif delivered:
        age_q = as_of_idx - quarter_index(p.deliv_year, p.deliv_q)
        occ = p.occupancy if p.occupancy is not None else 0
        if occ >= STABILIZED_OCC or age_q >= LEASEUP_WINDOW_QTRS:
            p.bucket = "STABILIZED / STABILIZING"
        else:
            p.bucket = "LEASING UP"
    else:
        p.bucket = "PROPOSED"

    # Lease-ups are where rent/occ accuracy matters most -> flag for HelloData
    if p.bucket == "LEASING UP":
        p.note("Verify rent/occ w/ HelloData")
    # Pipeline assets carry no real rent/occ yet
    if p.bucket in ("UNDER CONSTRUCTION", "PROPOSED"):
        p.occupancy = None
        p.eff_rent = None
    if p.bucket == "PROPOSED" and "RealPage" in p.sources and "CoStar" not in p.sources:
        p.note("RealPage pipeline (not yet in CoStar)")
    if p.units and p.units < 50 and "CoStar" not in p.sources:
        p.note("RealPage only (sub-50 unit)")


def derive_type(p: Prop) -> str:
    style = (p.style or "").lower()
    if "town" in style or "town" in p.name.lower():
        return "Townhome"
    if "single" in style:
        return "Single Family"
    if style in ("podium", "wrap"):
        return "Mid-Rise"
    if style == "garden":
        # refine by stories
        if p.stories and p.stories >= 4:
            return "Mid-Rise"
        return "Garden"
    s = p.stories or 0
    if s >= 8:
        return "High-Rise"
    if s >= 4:
        return "Mid-Rise"
    if s == 3:
        return "Low-Rise"
    return "Garden"


# --------------------------------------------------------------------------- #
# Workbook writer
# --------------------------------------------------------------------------- #
BUCKET_ORDER = [
    ("STABILIZED / STABILIZING", "Stabilized / leasing-complete deliveries"),
    ("LEASING UP", "Recently delivered, still leasing"),
    ("UNDER CONSTRUCTION", "Under construction, not yet delivered"),
    ("PROPOSED", "Proposed / pre-planned pipeline"),
]

COLS = {  # column letter -> (header, width)
    "B": ("#", 3.5),
    "C": ("Property", 30),
    "D": ("Units", 8),
    "E": ("Est. Delivery", 11.5),
    "F": ("Occupancy", 10),
    "G": ("Avg Mkt Rent", 13),
    "H": ("Owner", 22),
    "I": ("Proximity", 11),
    "J": ("Type", 14),
    "K": ("Notes", 30),
}


def write_workbook(props, subject_name, latest_inv, latest_label,
                   as_of, target, out_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Competitive Analysis"
    ws.sheet_view.showGridLines = False

    for col, (_, width) in COLS.items():
        ws.column_dimensions[col].width = width
    ws.column_dimensions["A"].width = 3

    # ---- Title / subtitle ----
    ws.merge_cells("C2:G2")
    t = ws["C2"]
    t.value = f"{subject_name.upper()} — 5-Mile Competitive Supply"
    t.fill = fill(NAVY); t.font = font(bold=True, size=13, color=WHITE)
    t.alignment = CENTER
    ws.row_dimensions[2].height = 22

    delivered_years = sorted({p.deliv_year for p in props if p.deliv_year})
    yr_lo = delivered_years[0] if delivered_years else as_of[0]
    ws.merge_cells("B3:F3")
    s = ws["B3"]
    s.value = (f"{yr_lo}–{as_of[0]} New Construction  |  5-mile radius  |  "
               f"As of {latest_label}  |  Stabilization Target:")
    s.fill = fill(BLUE); s.font = font(size=9, color=WHITE); s.alignment = LEFT
    g = ws["G3"]
    g.value = '=TEXT(Settings!$B$2,"0%")&" target"'
    g.fill = fill(BLUE); g.font = font(bold=True, size=9, color=YELLOW)
    g.alignment = CENTER

    # ---- Header row 4 ----
    hr = 4
    for col, (label, _) in COLS.items():
        c = ws[f"{col}{hr}"]
        c.value = label
        c.fill = fill(NAVY); c.font = font(bold=True, size=9, color=WHITE)
        c.alignment = CENTER; c.border = BORDER
    ws.row_dimensions[hr].height = 18

    # ---- Buckets ----
    r = hr + 1
    roster_first = None
    roster_last = None
    idx = 0
    for bucket, _desc in BUCKET_ORDER:
        members = [p for p in props if p.bucket == bucket]
        if not members:
            continue
        members.sort(key=lambda p: (-(quarter_index(p.deliv_year, p.deliv_q)
                                       if p.deliv_year else -1), -(p.units or 0)))
        tot_u = sum(p.units or 0 for p in members)
        # section header
        ws.merge_cells(f"C{r}:K{r}")
        sc = ws[f"C{r}"]
        sc.value = f"{bucket} ({len(members)} properties, {tot_u:,} units)"
        sc.fill = fill(BLUE); sc.font = font(bold=True, size=9, color=WHITE)
        sc.alignment = LEFT
        ws[f"B{r}"].fill = fill(BLUE)
        r += 1
        for p in members:
            idx += 1
            roster_first = roster_first or r
            roster_last = r
            row_vals = {
                "B": idx,
                "C": p.name,
                "D": p.units,
                "E": p.est_delivery or "TBD",
                "F": p.occupancy,
                "G": p.eff_rent if p.eff_rent else "—",
                "H": p.owner or "—",
                "I": None,                       # Proximity — manual fill
                "J": p.prop_type,
                "K": "; ".join(p.notes) if p.notes else None,
            }
            for col, val in row_vals.items():
                c = ws[f"{col}{r}"]
                c.value = val
                c.fill = fill(GRAY); c.font = font(size=9)
                c.border = BORDER
                c.alignment = LEFT if col in ("C", "H", "K") else CENTER
            ws[f"D{r}"].number_format = "#,##0"
            ws[f"F{r}"].number_format = "0.0%"
            ws[f"G{r}"].number_format = '#,##0;;"—"'
            r += 1
        r += 1  # spacer

    # ---- Quarterly absorption summary ----
    summ_start = r + 1
    headers = ["Est. Delivery", "", "# Props", "Total Units",
               "Wtd Avg\nOccupancy", "Currently\nOccupied", "Target\n@ Goal",
               "Units to\n95%", "Units to\n92.5%", "Units to\n90%"]
    cols = ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K"]
    ws.merge_cells(f"B{summ_start}:C{summ_start}")
    for col, label in zip(cols, headers):
        c = ws[f"{col}{summ_start}"]
        if col != "C":            # C is inside the B:C merge (read-only)
            c.value = label
        c.fill = fill(BLUE); c.font = font(bold=True, size=8, color=WHITE)
        c.alignment = CENTER; c.border = BORDER
    ws.row_dimensions[summ_start].height = 24

    # Quarter span: newest delivery quarter down to oldest (dated rows only)
    dated = [(p.deliv_year, p.deliv_q) for p in props if p.deliv_year and p.deliv_q]
    rng = f"$E${roster_first}:$E${roster_last}"
    drng = f"$D${roster_first}:$D${roster_last}"
    frng = f"$F${roster_first}:$F${roster_last}"
    sr = summ_start + 1
    first_q = sr
    if dated:
        hi = max(quarter_index(y, q) for y, q in dated)
        lo = min(quarter_index(y, q) for y, q in dated)
        for qi in range(hi, lo - 1, -1):
            y, q = qi // 4, (qi % 4) + 1
            label = fmt_quarter(y, q)
            ws[f"B{sr}"] = label
            ws.merge_cells(f"B{sr}:C{sr}")
            ws[f"D{sr}"] = f'=COUNTIF({rng},B{sr})'
            ws[f"E{sr}"] = f'=SUMIF({rng},B{sr},{drng})'
            ws[f"F{sr}"] = ArrayFormula(
                f"F{sr}",
                f"=IFERROR(SUMPRODUCT(({rng}=B{sr})*IFERROR({drng}*{frng},0))"
                f"/SUMPRODUCT(({rng}=B{sr})*ISNUMBER({frng})*{drng}),0)")
            ws[f"G{sr}"] = f"=F{sr}*E{sr}"
            ws[f"H{sr}"] = f"=E{sr}*Settings!$B$2"
            ws[f"I{sr}"] = f"=(E{sr}*0.95)-G{sr}"
            ws[f"J{sr}"] = f"=(E{sr}*0.925)-G{sr}"
            ws[f"K{sr}"] = f"=(E{sr}*0.90)-G{sr}"
            for col in cols:
                c = ws[f"{col}{sr}"]
                c.fill = fill(GRAY); c.font = font(size=9)
                c.border = BORDER; c.alignment = CENTER
            ws[f"E{sr}"].number_format = "#,##0"
            ws[f"F{sr}"].number_format = "0.0%"
            for col in ("G", "H", "I", "J", "K"):
                ws[f"{col}{sr}"].number_format = "#,##0"
            sr += 1
    last_q = sr - 1

    # ---- TOTAL row ----
    tr = sr
    ws.merge_cells(f"B{tr}:C{tr}")
    ws[f"B{tr}"] = "TOTAL"
    ws[f"D{tr}"] = f"=SUM(D{first_q}:D{last_q})"
    ws[f"E{tr}"] = f"=SUM(E{first_q}:E{last_q})"
    ws[f"F{tr}"] = f"=IFERROR(SUMPRODUCT($E${first_q}:$E${last_q},F{first_q}:F{last_q})/$E${tr},0)"
    for col in ("G", "H", "I", "J", "K"):
        ws[f"{col}{tr}"] = f"=SUM({col}{first_q}:{col}{last_q})"
    for col in cols:
        c = ws[f"{col}{tr}"]
        c.fill = fill(NAVY); c.font = font(bold=True, size=9, color=WHITE)
        c.border = BORDER; c.alignment = CENTER
    ws[f"B{tr}"].alignment = LEFT
    ws[f"E{tr}"].number_format = "#,##0"
    ws[f"F{tr}"].number_format = "0.0%"
    for col in ("G", "H", "I", "J", "K"):
        ws[f"{col}{tr}"].number_format = "#,##0"

    # ---- Inventory + % to be absorbed ----
    ir = tr + 2
    ws.merge_cells(f"G{ir}:H{ir}")
    ws[f"G{ir}"] = "Total Current Inventory"
    ws[f"G{ir}"].fill = fill(NAVY); ws[f"G{ir}"].font = font(bold=True, color=WHITE)
    ws[f"G{ir}"].alignment = CENTER
    ws[f"I{ir}"] = latest_inv
    ws[f"I{ir}"].number_format = "#,##0"; ws[f"I{ir}"].alignment = CENTER
    ws[f"J{ir}"] = f"=I{ir}"; ws[f"K{ir}"] = f"=I{ir}"
    ws[f"J{ir}"].number_format = "#,##0"; ws[f"K{ir}"].number_format = "#,##0"

    ar = ir + 1
    ws.merge_cells(f"G{ar}:H{ar}")
    ws[f"G{ar}"] = "% to be Absorbed"
    ws[f"G{ar}"].fill = fill(NAVY); ws[f"G{ar}"].font = font(bold=True, color=WHITE)
    ws[f"G{ar}"].alignment = CENTER
    ws[f"I{ar}"] = f"=IFERROR(I{tr}/I{ir},0)"
    ws[f"J{ar}"] = f"=IFERROR(J{tr}/J{ir},0)"
    ws[f"K{ar}"] = f"=IFERROR(K{tr}/K{ir},0)"
    for col in ("I", "J", "K"):
        ws[f"{col}{ar}"].number_format = "0.0%"; ws[f"{col}{ar}"].alignment = CENTER

    note_r = ar + 2
    ws.merge_cells(f"B{note_r}:K{note_r}")
    ws[f"B{note_r}"] = ("* 5-Mile radius. Total inventory from CoStar Data Analytics. "
                        "Proximity (mi) to be filled manually. Lease-up rent/occupancy "
                        "should be verified with a HelloData pull.")
    ws[f"B{note_r}"].font = font(size=8, color="FF808080")

    # ---- Settings sheet ----
    sset = wb.create_sheet("Settings")
    sset["A1"] = "ANALYSIS SETTINGS"; sset["A1"].font = font(bold=True, size=11)
    sset["A2"] = "Stabilized Occupancy Target"
    sset["B2"] = target; sset["B2"].number_format = "0%"
    sset["C2"] = "Target occupancy for absorption calculation"
    sset.column_dimensions["A"].width = 28
    sset.column_dimensions["C"].width = 42

    # ---- Reconciliation log sheet ----
    rlog = wb.create_sheet("Reconciliation Log")
    rlog.append(["Property", "Address", "Units", "Year Built", "Est. Delivery",
                 "Bucket", "Occupancy", "Eff Rent", "Sources", "Notes"])
    for c in rlog[1]:
        c.font = font(bold=True, color=WHITE); c.fill = fill(NAVY)
    for p in sorted(props, key=lambda x: (BUCKET_ORDER.index(
            next(b for b in BUCKET_ORDER if b[0] == x.bucket)),
            -(p_qi(x)))):
        rlog.append([p.name, p.address, p.units, p.year_built,
                     p.est_delivery, p.bucket, p.occupancy, p.eff_rent,
                     "+".join(sorted(p.sources)), "; ".join(p.notes)])
    for col in "ABCDEFGHIJ":
        rlog.column_dimensions[col].width = 18
    rlog.column_dimensions["A"].width = 30
    rlog.column_dimensions["J"].width = 40

    wb.save(out_path)


def p_qi(p):
    return quarter_index(p.deliv_year, p.deliv_q) if p.deliv_year and p.deliv_q else -1


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def parse_as_of(label: str) -> tuple[int, int]:
    m = _QRE.search(label or "")
    if m:
        return int(m.group(1)), int(m.group(2))
    today = dt.date.today()
    return today.year, (today.month - 1) // 3 + 1


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--subject-name", required=True)
    ap.add_argument("--subject-address", default="")
    ap.add_argument("--costar-roster", required=True)
    ap.add_argument("--costar-analytics", required=True)
    ap.add_argument("--realpage", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--as-of", default=None,
                    help="Override analysis quarter, e.g. '2026 Q2'")
    ap.add_argument("--target", type=float, default=DEFAULT_STABILIZATION_TARGET)
    args = ap.parse_args(argv)

    latest_inv, deliveries, latest_label = parse_costar_analytics(args.costar_analytics)
    as_of = parse_as_of(args.as_of or latest_label)

    costar = parse_costar_roster(args.costar_roster)
    realpage = parse_realpage(args.realpage)

    props = reconcile(costar, realpage)

    # Drop the subject property from the *competitive* roster.
    subj_addr = addr_key(args.subject_address)
    subj_name = name_key(args.subject_name)
    props = [p for p in props
             if not ((subj_addr and addr_key(p.address) == subj_addr)
                     or name_key(p.name) == subj_name)]

    # Keep only genuine new-construction supply: recent deliveries + pipeline.
    keep = []
    for p in props:
        pin_delivery(p, deliveries, as_of)
        recent = (p.year_built and
                  p.year_built >= as_of[0] - NEW_CONSTRUCTION_LOOKBACK_YEARS)
        if recent or _is_pipeline(p.status_raw):
            classify(p, as_of, args.target)
            p.prop_type = derive_type(p)
            keep.append(p)
    props = keep

    write_workbook(props, args.subject_name, latest_inv, latest_label,
                   as_of, args.target, args.out)

    # Console reconciliation report
    print(f"\nSupply chart written: {args.out}")
    print(f"As-of quarter: {fmt_quarter(*as_of)}  |  "
          f"Total current inventory (CoStar): {latest_inv:,}")
    print(f"Competitive new-construction properties: {len(props)}\n")
    for bucket, _ in BUCKET_ORDER:
        members = [p for p in props if p.bucket == bucket]
        if not members:
            continue
        u = sum(p.units or 0 for p in members)
        print(f"  {bucket} — {len(members)} props, {u:,} units")
        for p in sorted(members, key=lambda x: -p_qi(x)):
            occ = f"{p.occupancy:.1%}" if p.occupancy is not None else "  —  "
            print(f"      {p.name[:34]:34s} {str(p.units or '—'):>5} u  "
                  f"{p.est_delivery or 'TBD':>8}  occ {occ:>6}  "
                  f"[{'+'.join(sorted(p.sources))}]"
                  f"{('  | ' + '; '.join(p.notes)) if p.notes else ''}")
    print()


if __name__ == "__main__":
    sys.exit(main())
