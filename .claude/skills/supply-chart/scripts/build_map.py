#!/usr/bin/env python3
"""
build_map.py — Companion map for the 5-mile competitive Supply Chart.

Plots the subject property and the competitive new-construction roster (the same
buckets the chart uses) on a map, colour-coded by lifecycle stage, with a 5-mile
radius ring and per-property popups. Writes an interactive HTML map (with
Satellite / Terrain / Streets base layers; satellite default) and a matching PNG.

Geocoding is street-level via OpenStreetMap Nominatim (cached + rate-limited),
with an offline ZIP-centroid fallback per property. Use --no-geocode to force
ZIP-level placement; --subject-latlng "lat,lng" pins the subject exactly.

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
import math
import os
import sys

import folium

import build_supply_chart as bc
import geo

MILE_M = 1609.34

# Bucket -> (hex colour, short label), matching the chart's section colours.
BUCKET_STYLE = {
    "STABILIZED / STABILIZING": ("#1E90FF", "Stabilized"),
    "LEASING UP":               ("#FF8C00", "Leasing Up"),
    "UNDER CONSTRUCTION":       ("#00E676", "Under Construction"),
    "PROPOSED":                 ("#FF1744", "Proposed"),
}

TILES = {
    "terrain": ("https://{s}.tile.opentopomap.org/{z}/{x}/{y}.png", "Terrain",
                "Map data © OpenStreetMap contributors, SRTM | © OpenTopoMap (CC-BY-SA)"),
    "streets": ("OpenStreetMap", "Streets", None),
    "satellite": ("https://server.arcgisonline.com/ArcGIS/rest/services/"
                  "World_Imagery/MapServer/tile/{z}/{y}/{x}", "Satellite", "Esri"),
}


# --------------------------------------------------------------------------- #
# Map rendering
# --------------------------------------------------------------------------- #
def build_map(props, subject_name, subject_address, subject_latlng,
              use_geocoder, base, out_path, shadow_rows=None):
    loc = geo.Locator(use_geocoder,
                      os.path.join(os.path.dirname(out_path) or ".", ".geocode_cache.json"))

    # ---- subject ----
    subj_approx = False
    if subject_latlng:
        slat, slng = [float(x) for x in subject_latlng.split(",")]
    else:
        ll, subj_approx = loc.locate(subject_name, subject_address, None, None,
                                     geo.subject_zip(subject_address))
        if not ll:
            raise SystemExit("Could not locate subject — pass --subject-latlng "
                             "'lat,lng' or an address with a ZIP.")
        slat, slng = ll

    loc.set_anchor((slat, slng))       # reject comps mis-geocoded outside the radius

    # ---- locate comps once (shared by HTML + PNG) ----
    located, approx_any = [], subj_approx
    for p in sorted(props, key=lambda x: list(BUCKET_STYLE).index(x.bucket)):
        ll, approx = loc.locate(p.name, p.address, p.city, p.state, p.zipcode)
        if ll:
            located.append((p, ll[0], ll[1]))
            approx_any = approx_any or approx

    # ---- shadow-supply sites (latent / untracked) ----
    shadow_located = []
    for r in (shadow_rows or []):
        q = r.get("property") or r.get("notes")
        ll, _ = loc.locate(q, q, None, "TX", None)
        if ll:
            shadow_located.append((r, ll[0], ll[1]))
    loc.save()

    # ---- interactive HTML ----
    m = folium.Map(location=[slat, slng], zoom_start=12, tiles=None,
                   control_scale=True)
    for key in ([base] + [k for k in TILES if k != base]):
        url, nm, attr = TILES[key]
        folium.TileLayer(url, name=nm, attr=attr).add_to(m)
    folium.Circle([slat, slng], radius=5 * MILE_M, color="#FFD400", weight=3,
                  fill=True, fill_color="#FFD400", fill_opacity=0.03,
                  tooltip="5-mile radius").add_to(m)
    for p, plat, plng in located:
        color, blabel = BUCKET_STYLE.get(p.bucket, ("#888888", p.bucket))
        occ = f"{p.occupancy:.0%}" if p.occupancy is not None else "—"
        rent = f"${p.asking_rent:,.0f}" if p.asking_rent else "—"
        popup = folium.Popup(html=(
            f"<b>{p.name}</b><br>{blabel}<br>{(p.units or 0):,} units &middot; "
            f"{p.est_delivery or 'TBD'}<br>Occ {occ} &middot; Rent {rent}<br>"
            f"<span style='color:#888'>{p.address or ''} {p.zipcode or ''}</span>"),
            max_width=260)
        folium.CircleMarker(
            [plat, plng], radius=max(6, min(21, 5 + math.sqrt(p.units or 0) / 2)),
            color="#ffffff", weight=2, fill=True, fill_color=color, fill_opacity=1.0,
            opacity=1.0, popup=popup, tooltip=f"{p.name} ({blabel})").add_to(m)
    for r, rlat, rlng in shadow_located:
        folium.CircleMarker(
            [rlat, rlng], radius=10, color="#ffffff", weight=2, fill=True,
            fill_color="#B388FF", fill_opacity=0.55, dash_array="4",
            tooltip=f"SHADOW: {r.get('property')}",
            popup=folium.Popup(html=(
                f"<b>Shadow supply</b><br>{r.get('property')}<br>"
                f"{r.get('units') or '?'} units &middot; {r.get('status') or ''}<br>"
                f"{r.get('notes') or ''}"), max_width=260)).add_to(m)
    folium.Marker([slat, slng], tooltip=f"SUBJECT: {subject_name}",
                  popup=folium.Popup(f"<b>SUBJECT</b><br>{subject_name}<br>"
                                     f"{subject_address}", max_width=260),
                  icon=folium.Icon(color="black", icon="star", prefix="fa")).add_to(m)
    folium.LayerControl(collapsed=False).add_to(m)

    counts = {}
    for p in props:
        counts[p.bucket] = counts.get(p.bucket, 0) + 1
    legend = "".join(
        f"<div><span style='display:inline-block;width:11px;height:11px;"
        f"background:{c};border-radius:50%;margin-right:6px'></span>{lab} "
        f"({counts.get(b,0)})</div>" for b, (c, lab) in BUCKET_STYLE.items())
    if shadow_located:
        legend += ("<div><span style='display:inline-block;width:11px;height:11px;"
                   "background:#B388FF;border-radius:50%;margin-right:6px'></span>"
                   f"Shadow supply ({len(shadow_located)})</div>")
    note = ("&#9888; Some pins ZIP-approximate (geocode failed)."
            if approx_any else "")
    m.get_root().html.add_child(folium.Element(f"""
      <div style="position:fixed;top:10px;left:50px;z-index:9999;background:white;
        padding:8px 12px;border:1px solid #999;border-radius:4px;font-family:Arial;
        font-size:13px"><b>{subject_name} — 5-Mile Competitive Supply</b><br>
        <span style="font-size:11px">{len(located)} competitive properties &middot;
        5-mile radius</span></div>
      <div style="position:fixed;bottom:24px;left:12px;z-index:9999;background:white;
        padding:8px 12px;border:1px solid #999;border-radius:4px;font-family:Arial;
        font-size:12px"><b>Lifecycle</b>{legend}
        <div style="margin-top:4px">&#9733; Subject</div>
        <div style="color:#b00;margin-top:4px">{note}</div></div>"""))
    m.save(out_path)

    if out_path.lower().endswith(".html"):
        render_png(located, slat, slng, subject_name, counts, approx_any, base,
                   out_path[:-5] + ".png", shadow_located)
    return len(located), approx_any


def render_png(located, slat, slng, subject_name, counts, approx, base, png_path,
               shadow_located=None):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import matplotlib.patheffects as pe
    import contextily as cx

    # dark outline so white labels stay legible over satellite imagery
    _OUTLINE = [pe.withStroke(linewidth=2, foreground="#000000")]

    coslat = math.cos(math.radians(slat))
    span_lat = 6.0 / 69.0
    span_lng = 6.0 / (69.0 * coslat)

    fig, ax = plt.subplots(figsize=(10, 10))
    # 5-mile ring as a lat/lon polygon
    ring = [(slng + (5 / (69.0 * coslat)) * math.sin(t),
             slat + (5 / 69.0) * math.cos(t))
            for t in [i * math.pi / 45 for i in range(91)]]
    ax.plot([x for x, _ in ring], [y for _, y in ring], color="#FFD400",
            lw=2.2, ls="--", zorder=4)
    for p, plat, plng in located:
        color, _ = BUCKET_STYLE.get(p.bucket, ("#888888", ""))
        ax.scatter(plng, plat, s=max(60, min(460, (p.units or 0) * 1.2)), c=color,
                   edgecolors="white", linewidths=1.4, alpha=1.0, zorder=5)
        ax.annotate(p.name[:22], (plng, plat), fontsize=6, xytext=(0, 7),
                    textcoords="offset points", ha="center", color="white", zorder=6,
                    path_effects=_OUTLINE)
    for r, rlat, rlng in (shadow_located or []):
        ax.scatter(rlng, rlat, marker="P", s=180, c="#B388FF", edgecolors="white",
                   linewidths=1.2, alpha=1.0, zorder=5)
        ax.annotate((r.get("property") or "")[:20], (rlng, rlat), fontsize=6,
                    xytext=(0, 7), textcoords="offset points", ha="center",
                    color="white", zorder=6, path_effects=_OUTLINE)
    ax.scatter(slng, slat, marker="*", s=640, c="#FFD400", edgecolors="black",
               linewidths=1.2, zorder=7)
    ax.annotate(f"SUBJECT: {subject_name}", (slng, slat), fontsize=8.5,
                fontweight="bold", xytext=(0, -16), textcoords="offset points",
                ha="center", color="white", zorder=7, path_effects=_OUTLINE)
    ax.set_xlim(slng - span_lng, slng + span_lng)
    ax.set_ylim(slat - span_lat, slat + span_lat)
    ax.set_aspect(1.0 / coslat)
    src = (cx.providers.OpenTopoMap if base == "terrain"
           else cx.providers.Esri.WorldImagery if base == "satellite"
           else cx.providers.OpenStreetMap.Mapnik)
    try:
        cx.add_basemap(ax, crs="EPSG:4326", source=src, attribution_size=5)
    except Exception as e:
        print(f"  (basemap fetch failed: {e}; PNG drawn without tiles)")
    ax.set_xticks([]); ax.set_yticks([])
    handles = [plt.Line2D([], [], marker="o", ls="", color=c,
               label=f"{lab} ({counts.get(b,0)})")
               for b, (c, lab) in BUCKET_STYLE.items()]
    if shadow_located:
        handles.append(plt.Line2D([], [], marker="P", ls="", color="#B388FF",
                       label=f"Shadow supply ({len(shadow_located)})"))
    handles.append(plt.Line2D([], [], marker="*", ls="", color="black", label="Subject"))
    ax.legend(handles=handles, loc="upper left", fontsize=8, framealpha=0.92)
    title = f"{subject_name} — 5-Mile Competitive Supply"
    if approx:
        title += "  (some positions ZIP-approximate)"
    ax.set_title(title, fontsize=12)
    fig.tight_layout()
    fig.savefig(png_path, dpi=140)
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
    ap.add_argument("--base", choices=list(TILES), default="satellite",
                    help="Default base layer (satellite/terrain/streets).")
    ap.add_argument("--no-geocode", action="store_true",
                    help="Skip street geocoding; use ZIP centroids only.")
    ap.add_argument("--diligence", default=None,
                    help="Filled diligence CSV — applies researched bucket/date "
                         "corrections and plots its type=shadow rows as latent "
                         "supply (keeps the map consistent with the chart).")
    args = ap.parse_args(argv)

    _, deliveries, latest_label, _, _ = bc.parse_costar_analytics(args.costar_analytics)
    as_of = bc.parse_as_of(args.as_of or latest_label)
    props = bc.build_competitive_roster(
        args.costar_roster, args.realpage, deliveries, as_of,
        args.subject_name, args.subject_address, args.target)

    shadow_rows = None
    if args.diligence:
        rows = bc.load_diligence(args.diligence)
        bc.apply_diligence(props, rows)          # match the chart's corrected buckets
        shadow_rows = [r for r in rows if r.get("type") == "shadow"]

    placed, approx = build_map(props, args.subject_name, args.subject_address,
                               args.subject_latlng, not args.no_geocode,
                               args.base, args.out, shadow_rows)
    print(f"Map written: {args.out}  ({placed} properties"
          f"{', some ZIP-approximate' if approx else ', street-geocoded'})")


if __name__ == "__main__":
    sys.exit(main())
