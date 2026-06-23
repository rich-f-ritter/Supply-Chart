#!/usr/bin/env python3
"""
build_map.py — Companion map for the 5-mile competitive Supply Chart.

Plots the subject property and the competitive new-construction roster (the same
buckets the chart uses) on an interactive HTML map, colour-coded by lifecycle
stage, with a 5-mile radius ring and per-property popups.

Geocoding: street-level via a geocoder when reachable (--geocode), otherwise an
offline ZIP-code centroid (with a small deterministic jitter so same-ZIP
properties don't overlap). ZIP-level placement is approximate — clearly labelled.

Usage:
    python build_map.py \
        --subject-name "Aura Beacon Island" \
        --subject-address "2200 Beacon Cir, League City, TX 77573" \
        --costar-roster   examples/aura_beacon_island/CoStar_5mi_50unit_properties.xlsx \
        --costar-analytics examples/aura_beacon_island/CoStar_5mi_Data_Analytics.xlsx \
        --realpage        examples/aura_beacon_island/Realpage_5mi.xlsx \
        --out             output/Aura_Beacon_Island__Map.html
"""
from __future__ import annotations

import argparse
import hashlib
import math
import re
import sys

import folium
import zipcodes

import build_supply_chart as bc

MILE_M = 1609.34

# Bucket -> (hex colour, short label), matching the chart's section colours.
BUCKET_STYLE = {
    "STABILIZED / STABILIZING": ("#2E75B6", "Stabilized"),
    "LEASING UP":               ("#ED7D31", "Leasing Up"),
    "UNDER CONSTRUCTION":       ("#375623", "Under Construction"),
    "PROPOSED":                 ("#C00000", "Proposed"),
}


def zip_centroid(zipcode):
    if not zipcode:
        return None
    recs = zipcodes.matching(str(zipcode)[:5]) if zipcodes.is_real(str(zipcode)[:5]) else []
    if not recs:
        return None
    return float(recs[0]["lat"]), float(recs[0]["long"])


def jitter(seed, scale=0.0055):
    """Deterministic ~0.3-mi offset so same-ZIP markers separate."""
    h = int(hashlib.md5(seed.encode()).hexdigest(), 16)
    ang = (h % 360) * math.pi / 180
    rad = ((h // 360) % 100) / 100 * scale
    return rad * math.cos(ang), rad * math.sin(ang)


def geocode_nominatim(query):
    """Optional street-level geocode (only if the host is reachable)."""
    import json
    import urllib.parse
    import urllib.request
    url = ("https://nominatim.openstreetmap.org/search?format=json&limit=1&q="
           + urllib.parse.quote(query))
    req = urllib.request.Request(url, headers={"User-Agent": "supply-chart-skill/1.0"})
    with urllib.request.urlopen(req, timeout=20) as r:
        data = json.load(r)
    return (float(data[0]["lat"]), float(data[0]["lon"])) if data else None


def locate(name, address, city, zipcode, use_geocoder):
    """Return ((lat, lng), approximate_bool)."""
    if use_geocoder:
        try:
            q = ", ".join(x for x in (address, city, zipcode) if x)
            ll = geocode_nominatim(q)
            if ll:
                return ll, False
        except Exception:
            pass
    c = zip_centroid(zipcode)
    if c:
        dy, dx = jitter(name or address or zipcode)
        return (c[0] + dy, c[1] + dx), True
    return None, True


def _subject_zip(address):
    m = re.search(r"\b(\d{5})\b", address or "")
    return m.group(1) if m else None


def build_map(props, subject_name, subject_address, subject_latlng,
              use_geocoder, target, out_path):
    # ---- subject location ----
    subj_approx = False
    if subject_latlng:
        slat, slng = [float(x) for x in subject_latlng.split(",")]
    else:
        loc, subj_approx = locate(subject_name, subject_address, None,
                                  _subject_zip(subject_address), use_geocoder)
        if not loc:
            raise SystemExit("Could not locate subject — pass --subject-latlng "
                             "'lat,lng' or a subject address with a ZIP.")
        slat, slng = loc

    # ---- locate every comp once (shared by HTML + PNG) ----
    located = []   # (prop, lat, lng)
    approx_any = subj_approx
    for p in sorted(props, key=lambda x: list(BUCKET_STYLE).index(x.bucket)):
        loc, approx = locate(p.name, p.address, p.city, p.zipcode, use_geocoder)
        if not loc:
            continue
        approx_any = approx_any or approx
        located.append((p, loc[0], loc[1]))

    m = folium.Map(location=[slat, slng], zoom_start=12, tiles="cartodbpositron",
                   control_scale=True)

    # 5-mile ring
    folium.Circle([slat, slng], radius=5 * MILE_M, color="#1F3864", weight=2,
                  fill=True, fill_opacity=0.04,
                  tooltip="5-mile radius").add_to(m)

    placed = len(located)
    for p, plat, plng in located:
        loc = (plat, plng)
        color, blabel = BUCKET_STYLE.get(p.bucket, ("#888888", p.bucket))
        units = p.units or 0
        radius = max(5, min(20, 4 + math.sqrt(units) / 2))
        occ = (f"{p.occupancy:.0%}" if p.occupancy is not None else "—")
        rent = (f"${p.eff_rent:,.0f}" if p.eff_rent else "—")
        popup = folium.Popup(html=(
            f"<b>{p.name}</b><br>{blabel}<br>"
            f"{units:,} units &middot; {p.est_delivery or 'TBD'}<br>"
            f"Occ {occ} &middot; Rent {rent}<br>"
            f"<span style='color:#888'>{p.address or ''} {p.zipcode or ''}</span>"),
            max_width=260)
        folium.CircleMarker(
            loc, radius=radius, color=color, weight=1, fill=True,
            fill_color=color, fill_opacity=0.75, popup=popup,
            tooltip=f"{p.name} ({blabel})").add_to(m)

    # subject marker (distinct)
    folium.Marker(
        [slat, slng], tooltip=f"SUBJECT: {subject_name}",
        popup=folium.Popup(f"<b>SUBJECT</b><br>{subject_name}<br>{subject_address}",
                           max_width=260),
        icon=folium.Icon(color="black", icon="star", prefix="fa")).add_to(m)

    # ---- title + legend ----
    counts = {}
    for p in props:
        counts[p.bucket] = counts.get(p.bucket, 0) + 1
    legend_rows = "".join(
        f"<div><span style='display:inline-block;width:11px;height:11px;"
        f"background:{c};border-radius:50%;margin-right:6px'></span>"
        f"{lab} ({counts.get(b,0)})</div>"
        for b, (c, lab) in BUCKET_STYLE.items())
    note = ("&#9888; Marker positions are approximate (ZIP-centroid) — street "
            "geocoding unavailable; re-run with --geocode for exact pins."
            if approx_any else "")
    title = (f"<b>{subject_name} — 5-Mile Competitive Supply</b><br>"
             f"<span style='font-size:11px'>{placed} competitive properties &middot; "
             f"5-mile radius</span>")
    html = f"""
    <div style="position:fixed;top:10px;left:50px;z-index:9999;background:white;
         padding:8px 12px;border:1px solid #999;border-radius:4px;
         font-family:Arial;font-size:13px">{title}</div>
    <div style="position:fixed;bottom:24px;left:12px;z-index:9999;background:white;
         padding:8px 12px;border:1px solid #999;border-radius:4px;
         font-family:Arial;font-size:12px">
         <b>Lifecycle</b>{legend_rows}
         <div style="margin-top:4px">&#9733; Subject</div>
         <div style="color:#b00;margin-top:4px;max-width:240px">{note}</div></div>
    """
    m.get_root().html.add_child(folium.Element(html))
    m.save(out_path)

    # ---- static PNG quick-look (axes in miles from subject, no basemap) ----
    if out_path.lower().endswith(".html"):
        render_png(located, slat, slng, subject_name, counts,
                   approx_any, out_path[:-5] + ".png")
    return placed, approx_any


def render_png(located, slat, slng, subject_name, counts, approx, png_path):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.patches import Circle

    def to_miles(lat, lng):
        return ((lng - slng) * 69.0 * math.cos(math.radians(slat)),
                (lat - slat) * 69.0)

    fig, ax = plt.subplots(figsize=(9, 9))
    ax.add_patch(Circle((0, 0), 5, fill=True, color="#1F3864", alpha=0.05))
    ax.add_patch(Circle((0, 0), 5, fill=False, color="#1F3864", lw=1.5, ls="--"))
    for p, plat, plng in located:
        x, y = to_miles(plat, plng)
        color, _ = BUCKET_STYLE.get(p.bucket, ("#888888", ""))
        ax.scatter(x, y, s=max(40, min(400, (p.units or 0) * 1.1)),
                   c=color, edgecolors="white", linewidths=0.6, alpha=0.85, zorder=3)
        ax.annotate(p.name[:22], (x, y), fontsize=6, xytext=(0, 6),
                    textcoords="offset points", ha="center", color="#333")
    ax.scatter(0, 0, marker="*", s=520, c="black", edgecolors="white",
               linewidths=0.8, zorder=4)
    ax.annotate(f"SUBJECT: {subject_name}", (0, 0), fontsize=8, fontweight="bold",
                xytext=(0, -14), textcoords="offset points", ha="center")
    ax.set_xlim(-6, 6); ax.set_ylim(-6, 6); ax.set_aspect("equal")
    ax.set_xlabel("miles east →"); ax.set_ylabel("miles north →")
    ax.grid(True, color="#eee")
    handles = [plt.Line2D([], [], marker="o", ls="", color=c,
                          label=f"{lab} ({counts.get(b,0)})")
               for b, (c, lab) in BUCKET_STYLE.items()]
    handles.append(plt.Line2D([], [], marker="*", ls="", color="black", label="Subject"))
    ax.legend(handles=handles, loc="upper left", fontsize=8, framealpha=0.9)
    title = f"{subject_name} — 5-Mile Competitive Supply"
    if approx:
        title += "  (positions approximate · ZIP-level)"
    ax.set_title(title, fontsize=11)
    fig.tight_layout()
    fig.savefig(png_path, dpi=130)
    plt.close(fig)
    return png_path


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--subject-name", required=True)
    ap.add_argument("--subject-address", default="")
    ap.add_argument("--subject-latlng", default=None,
                    help="Exact subject 'lat,lng' to override geocoding.")
    ap.add_argument("--costar-roster", required=True)
    ap.add_argument("--costar-analytics", required=True)
    ap.add_argument("--realpage", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--as-of", default=None)
    ap.add_argument("--target", type=float, default=bc.DEFAULT_STABILIZATION_TARGET)
    ap.add_argument("--geocode", action="store_true",
                    help="Attempt street-level geocoding (needs network access).")
    args = ap.parse_args(argv)

    _, deliveries, latest_label, _, _ = bc.parse_costar_analytics(args.costar_analytics)
    as_of = bc.parse_as_of(args.as_of or latest_label)
    props = bc.build_competitive_roster(
        args.costar_roster, args.realpage, deliveries, as_of,
        args.subject_name, args.subject_address, args.target)

    placed, approx = build_map(props, args.subject_name, args.subject_address,
                               args.subject_latlng, args.geocode, args.target,
                               args.out)
    print(f"Map written: {args.out}  ({placed} properties plotted"
          f"{', ZIP-approximate' if approx else ''})")


if __name__ == "__main__":
    sys.exit(main())
