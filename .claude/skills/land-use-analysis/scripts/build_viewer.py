"""Build the self-contained interactive HTML viewer (single file, data + Leaflet inlined ->
opens by double-click, emails, works offline / in the claude.ai preview with no JS CDN).
Three views (Land Use / Zoning / Vacant Threats), satellite/streets basemaps, hover read-outs,
a subject pin, present-only legends, a 'Download PNG' button.

Two correctness fixes baked in:
  * COINCIDENT-FOOTPRINT DEDUPE — condos store one building footprint shared by every unit
    parcel (hundreds of identical polygons). Drawn translucent they compound into a darker
    shade than a single parcel of the SAME category, so one red category reads as two colors
    and bloats the file. We draw each distinct footprint ONCE (carrying a unit count `n`).
  * Fills are drawn near-opaque so on-map colors match the solid legend swatches.

Leaflet is inlined from assets/vendor/ (bundled with the skill) so there's NO CDN dependency;
only the basemap RASTER TILES need internet at view time (they can't be bundled). If tiles are
blocked/offline the map degrades gracefully (neutral canvas + an 'offline basemap' badge) and
the parcels still render. Pass --served for sidecar data files instead of one inlined HTML.
"""
import hashlib
import json
import sys
from collections import Counter
from pathlib import Path

from shapely.geometry import mapping, shape

sys.path.insert(0, str(Path(__file__).resolve().parent))
import config as C  # noqa: E402
import geo  # noqa: E402
import palette as P  # noqa: E402

SIMPLIFY_TOL = 0.00004  # ~4 m
VENDOR = Path(__file__).resolve().parents[1] / "assets" / "vendor"


def _round(c, nd=5):
    if isinstance(c[0], (int, float)):
        return [round(c[0], nd), round(c[1], nd)]
    return [_round(x, nd) for x in c]


def simp(geom):
    g = shape(geom)
    if not g.is_valid:
        g = g.buffer(0)
    m = mapping(g.simplify(SIMPLIFY_TOL, preserve_topology=True))
    return {"type": m["type"], "coordinates": _round(m["coordinates"])}


def build_payload(cfg):
    IN = C.indir(cfg)
    feats = json.loads((IN / "parcels_classified.geojson").read_text())["features"]
    dev = set(json.loads((IN / "developable_accounts.json").read_text())) \
        if (IN / "developable_accounts.json").exists() else set()

    # dedupe by simplified-geometry signature: draw each physical footprint once.
    out, sig_ix = [], {}
    dupes = 0
    for f in feats:
        p = f["properties"]
        g = simp(f["geometry"])
        sig = hashlib.md5(json.dumps(g["coordinates"], separators=(",", ":")).encode()).hexdigest()
        dv = bool(p["is_vacant"] and p.get("account") in dev)
        if sig in sig_ix:
            kept = out[sig_ix[sig]]["properties"]
            kept["n"] += 1
            if dv:
                kept["dv"] = True
            dupes += 1
            continue
        sig_ix[sig] = len(out)
        out.append({"type": "Feature", "geometry": g, "properties": {
            "o": p["owner"], "s": p["situs"], "u": p["landuse_code"], "b": p["bucket"],
            "j": p["jurisdiction"], "z": p["zone_plain"], "zc": p["zone_category"],
            "t": p["mf_threat"], "ac": p["acres"], "d": p["dist_mi"],
            "dc": p.get("data_confidence"), "lc": p["landuse_color"],
            "zcol": p["zone_color"], "tc": p["threat_color"], "dv": dv, "n": 1}})

    def leg(order, cmap, key):
        present = Counter(f["properties"][key] for f in feats)
        return [{"label": k, "color": cmap[k]} for k in order if present.get(k)]

    threat_present = lambda t: any(  # noqa: E731
        f["properties"]["mf_threat"] == t and f["properties"]["is_vacant"]
        and f["properties"].get("account") in dev for f in feats)
    legends = {
        "landuse": {"title": "LAND USE", "items": leg(P.LANDUSE_ORDER, P.LANDUSE_COLOR, "bucket")},
        "zoning": {"title": "ZONING (BASE DISTRICT)",
                   "items": leg(P.ZONE_CATEGORY_ORDER, P.ZONE_CATEGORY_COLOR, "zone_category")},
        "threat": {"title": "VACANT LAND - MF SUPPLY THREAT", "items": [
            {"label": lab, "color": P.THREAT_COLOR[t]} for t, lab in [
                ("High", "High - MF by-right"),
                ("Medium", "Medium - townhouse/mixed (conditional)"),
                ("Low", "Low - MF not permitted by base zone"),
                ("Unknown", "Unknown - planned-dev / corridor (verify plan)")]
            if threat_present(t)]},
    }
    slon, slat = geo.subject_lonlat(cfg)
    b = geo.bbox(cfg)
    cfgout = {"subject": {"lat": slat, "lon": slon, "name": cfg["subject"]["name"]},
              "bounds_all": [[b[1], b[0]], [b[3], b[2]]],
              "bounds_threat": [[slat - 0.033, slon - 0.045], [slat + 0.033, slon + 0.045]]}
    print(f"  viewer: {len(feats)} parcel records -> {len(out)} distinct footprints "
          f"({dupes} coincident merged)", file=sys.stderr)
    return out, legends, cfgout


def head_libs():
    """Inline Leaflet from bundled vendor files (no CDN); fall back to CDN if missing."""
    css, js, li = VENDOR / "leaflet.css", VENDOR / "leaflet.js", VENDOR / "leaflet-image.js"
    if css.exists() and js.exists():
        parts = [f"<style>{css.read_text(encoding='utf-8')}</style>",
                 f"<script>{js.read_text(encoding='utf-8')}</script>"]
        if li.exists():
            parts.append(f"<script>{li.read_text(encoding='utf-8')}</script>")
        return "\n".join(parts)
    return ('<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"/>\n'
            '<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>\n'
            '<script src="https://unpkg.com/leaflet-image@0.4.0/leaflet-image.js"></script>')


def html(cfg, mode, parcels=None, legends=None, viewcfg=None):
    name = cfg["subject"]["name"]
    brand = cfg.get("branding", {})
    navy, accent = brand.get("primary", "#074070"), brand.get("accent", "#B49955")
    libs = head_libs()
    if mode == "inline":
        data_block = (
            f'<script type="application/json" id="d-parcels">{json.dumps({"type":"FeatureCollection","features":parcels})}</script>'
            f'<script type="application/json" id="d-legends">{json.dumps(legends)}</script>'
            f'<script type="application/json" id="d-config">{json.dumps(viewcfg)}</script>')
        loader = """
  const J=id=>JSON.parse(document.getElementById(id).textContent);
  start(J('d-parcels'),J('d-legends'),J('d-config'));"""
    else:
        data_block = ""
        loader = """
  Promise.all([fetch('viewer_parcels.json').then(r=>r.json()),
    fetch('viewer_legends.json').then(r=>r.json()),
    fetch('viewer_config.json').then(r=>r.json())]).then(([p,l,c])=>start(p,l,c));"""

    return f"""<!DOCTYPE html><html><head><meta charset="utf-8"/>
<title>{name} - Land Use Analysis</title>
{libs}
<style>
 html,body{{height:100%;margin:0;font-family:Calibri,Arial,sans-serif}}
 #map{{position:absolute;top:54px;left:0;right:0;bottom:0;background:#e8eaed}}
 #bar{{position:absolute;top:0;left:0;right:0;height:54px;z-index:1100;background:{navy};color:#fff;
   display:flex;align-items:center;gap:16px;padding:0 16px;box-shadow:0 1px 6px rgba(0,0,0,.3)}}
 #bar .brand{{font-weight:700;font-size:15px;margin-right:6px}}
 #bar .brand small{{font-weight:400;opacity:.8}}
 .grp{{display:flex;border:1px solid rgba(255,255,255,.4);border-radius:5px;overflow:hidden}}
 .grp button{{background:transparent;color:#fff;border:0;padding:7px 13px;font:600 13px Calibri,Arial;
   cursor:pointer;border-right:1px solid rgba(255,255,255,.25)}}
 .grp button:last-child{{border-right:0}}
 .grp button.on{{background:{accent};color:#1a1a1a}}
 .grp button:hover:not(.on){{background:rgba(255,255,255,.15)}}
 .lbl{{font-size:11px;text-transform:uppercase;letter-spacing:.08em;opacity:.7;margin-right:-8px}}
 #dl{{margin-left:auto;background:{accent};color:#1a1a1a;border:0;border-radius:5px;padding:8px 15px;
   font:700 13px Calibri,Arial;cursor:pointer}}
 .legend{{position:absolute;right:14px;bottom:18px;z-index:1000;background:rgba(255,255,255,.95);
   border:1px solid #bbb;border-radius:4px;padding:9px 12px;font:12px/1.35 Calibri,Arial;color:#222;
   box-shadow:0 1px 4px rgba(0,0,0,.25);max-width:300px}}
 .legend h4{{margin:0 0 6px;font:700 11px Calibri,Arial;letter-spacing:.06em;color:{navy}}}
 .legend .row{{display:flex;align-items:center;margin:2px 0}}
 .legend .sw{{width:14px;height:14px;margin-right:7px;border:1px solid #555;flex:none}}
 .legend .pin{{width:11px;height:11px;margin-right:7px;border-radius:50% 50% 50% 0;transform:rotate(-45deg);
   background:#F39200;border:1px solid #8a5200;flex:none}}
 #info{{position:absolute;left:14px;bottom:18px;z-index:1000;background:rgba(255,255,255,.96);
   border:1px solid #bbb;border-radius:4px;padding:8px 11px;font:12px/1.4 Calibri,Arial;color:#222;
   box-shadow:0 1px 4px rgba(0,0,0,.25);max-width:320px;display:none}}
 #info b{{color:{navy}}}
 #status,#tilewarn{{position:absolute;left:50%;transform:translateX(-50%);z-index:1200;color:#fff;
   padding:6px 14px;border-radius:4px;font:13px Calibri,Arial}}
 #status{{top:64px;background:{navy}}}
 #tilewarn{{bottom:14px;background:#8a5a00;display:none;font-size:12px;opacity:.95}}
</style></head><body>
<div id="bar">
 <div class="brand">{name} <small>Land Use Analysis</small></div>
 <span class="lbl">View</span>
 <div class="grp" id="mapgrp">
  <button data-map="landuse" class="on">Land Use</button>
  <button data-map="zoning">Zoning</button>
  <button data-map="threat">Vacant Threats</button></div>
 <span class="lbl">Basemap</span>
 <div class="grp" id="basegrp">
  <button data-base="satellite" class="on">Satellite</button>
  <button data-base="streets">Streets</button>
  <button data-base="none">None</button></div>
 <button id="dl">&#x2193; Download PNG</button>
</div>
<div id="map"></div><div id="status">Loading parcels…</div>
<div id="tilewarn">Basemap tiles unavailable (offline) — parcels shown on a plain canvas.</div>
<div class="legend" id="legend"></div><div id="info"></div>
{data_block}
<script>
let curMap='landuse', curBase='satellite', CFG, LEG, subj, allLayer, threatLayer, baseLayer=null, tileErrs=0;
const map=L.map('map',{{zoomControl:true,attributionControl:true,preferCanvas:true}});
const TILES={{satellite:['https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{{z}}/{{y}}/{{x}}','Esri, Maxar, Earthstar Geographics'],
 streets:['https://server.arcgisonline.com/ArcGIS/rest/services/World_Street_Map/MapServer/tile/{{z}}/{{y}}/{{x}}','Esri, HERE, Garmin']}};
function setBase(b){{ if(baseLayer){{map.removeLayer(baseLayer);baseLayer=null;}}
 document.getElementById('tilewarn').style.display='none';
 if(b==='none'||!TILES[b]) return;
 tileErrs=0;
 baseLayer=L.tileLayer(TILES[b][0],{{maxZoom:19,crossOrigin:true,attribution:TILES[b][1]}});
 baseLayer.on('tileerror',()=>{{ if(++tileErrs>=4) document.getElementById('tilewarn').style.display='block'; }});
 baseLayer.addTo(map); baseLayer.bringToBack(); }}
const info=document.getElementById('info');
function fmtAc(a){{return (a==null?'?':a.toLocaleString(undefined,{{maximumFractionDigits:2}}))+' ac';}}
function show(p){{ let h='<b>'+(p.o||'(owner n/a)')+'</b><br>'+(p.s||'')+'<br>';
 if(curMap==='landuse') h+='Land use: <b>'+p.b+'</b>'+(p.u?' ('+p.u+')':'')+'<br>'+fmtAc(p.ac)+' · '+p.d+' mi';
 else if(curMap==='zoning') h+='Zoning: <b>'+(p.z||'—')+'</b> ('+p.j+')<br>'+p.zc+'<br>MF threat: <b>'+p.t+'</b> · '+fmtAc(p.ac);
 else h+='Zoning: <b>'+(p.z||'—')+'</b> ('+p.j+')<br>MF threat: <b>'+p.t+'</b><br>'+fmtAc(p.ac)+' · '+p.d+' mi · developable vacant';
 if(p.n>1) h+='<br><i>'+p.n+' unit/parcel records share this footprint (e.g. condo)</i>';
 if(p.dc&&p.dc!=='full') h+='<br><i>'+p.dc+' source</i>';
 info.innerHTML=h; info.style.display='block'; }}
function colorFor(p){{ return curMap==='landuse'?p.lc:curMap==='zoning'?p.zcol:p.tc; }}
function styleFor(f){{ const p=f.properties, threat=curMap==='threat';
 return {{color:threat?'#222':'#555', weight:threat?0.9:0.3, fillColor:colorFor(p),
  fillOpacity:(curMap==='zoning'&&p.zc==='No public zoning (data gap)')?0.3:(threat?0.85:0.8)}}; }}
function hookHover(layer){{ layer.on('mouseover',e=>{{ if(e.layer&&e.layer.feature) show(e.layer.feature.properties);
   e.layer.setStyle({{weight:2,color:'#fff'}}); }});
 layer.on('mouseout',e=>{{ info.style.display='none'; if(e.layer) layer.resetStyle(e.layer); }}); }}
function setLegend(){{ const leg=LEG[curMap], S=CFG.subject;
 let h='<h4>'+leg.title+'</h4><div class="row"><span class="pin"></span>Subject &mdash; '+S.name+'</div>';
 for(const it of leg.items) h+='<div class="row"><span class="sw" style="background:'+it.color+'"></span>'+it.label+'</div>';
 document.getElementById('legend').innerHTML=h; }}
function applyMap(){{ const threat=curMap==='threat';
 if(allLayer) map.removeLayer(allLayer); if(threatLayer) map.removeLayer(threatLayer);
 if(threat){{ threatLayer.setStyle(styleFor); threatLayer.addTo(map); }}
 else {{ allLayer.setStyle(styleFor); allLayer.addTo(map); }}
 subj.bringToFront(); setLegend(); map.fitBounds(threat?CFG.bounds_threat:CFG.bounds_all); }}
function start(parcels,legends,cfg){{ LEG=legends; CFG=cfg; setBase(curBase);
 subj=L.circleMarker([cfg.subject.lat,cfg.subject.lon],{{radius:8,color:'#7a4a00',weight:2,fillColor:'#F39200',fillOpacity:1}});
 allLayer=L.geoJSON(parcels,{{style:styleFor}}); hookHover(allLayer);
 const th={{type:'FeatureCollection',features:parcels.features.filter(f=>f.properties.dv)}};
 threatLayer=L.geoJSON(th,{{style:styleFor}}); hookHover(threatLayer);
 subj.addTo(map); document.getElementById('status').remove(); applyMap(); }}
document.querySelectorAll('#mapgrp button').forEach(b=>b.onclick=()=>{{
 document.querySelectorAll('#mapgrp button').forEach(x=>x.classList.remove('on'));
 b.classList.add('on'); curMap=b.dataset.map; applyMap(); }});
document.querySelectorAll('#basegrp button').forEach(b=>b.onclick=()=>{{
 document.querySelectorAll('#basegrp button').forEach(x=>x.classList.remove('on'));
 b.classList.add('on'); curBase=b.dataset.base; setBase(curBase); }});
document.getElementById('dl').onclick=()=>{{ if(typeof leafletImage==='undefined'){{alert('PNG export unavailable');return;}}
 leafletImage(map,(err,canvas)=>{{ if(err)return;
 const W=canvas.width,H2=canvas.height,c=document.createElement('canvas'); c.width=W;c.height=H2;
 const x=c.getContext('2d'); x.drawImage(canvas,0,0);
 x.fillStyle='{navy}'; x.fillRect(0,0,W,40); x.fillStyle='#fff'; x.font='bold 20px Calibri';
 x.fillText('{name} — '+LEG[curMap].title,14,26);
 const a=document.createElement('a'); a.download='{name}'.replace(/[^a-z0-9]+/gi,'-').toLowerCase()+'_'+curMap+'_'+curBase+'.png';
 a.href=c.toDataURL('image/png'); a.click(); }}); }};
{loader}
</script></body></html>"""


def main():
    cfg = C.load()
    root = C.root(cfg)
    served = "--served" in sys.argv
    parcels, legends, viewcfg = build_payload(cfg)
    safe = "".join(ch if ch.isalnum() or ch in " -_" else "" for ch in cfg["subject"]["name"]).strip()
    out = root / f"{safe} - Land Use Viewer.html"

    if served:
        (root / "viewer_parcels.json").write_text(json.dumps({"type": "FeatureCollection", "features": parcels}))
        (root / "viewer_legends.json").write_text(json.dumps(legends))
        (root / "viewer_config.json").write_text(json.dumps(viewcfg))
        out.write_text(html(cfg, "served"), encoding="utf-8")
        print(f"wrote served viewer (3 sidecar files) -> {out.name}  (serve folder with python -m http.server)")
        return

    page = html(cfg, "inline", parcels, legends, viewcfg)
    out.write_text(page, encoding="utf-8")
    mb = len(page.encode("utf-8")) / 1e6
    print(f"wrote self-contained viewer -> {out.name}  ({mb:.1f} MB, {len(parcels)} footprints)")
    if mb > 35:
        print("  [note] large file - if sluggish, rebuild with --served for sidecar data + a local server.")


if __name__ == "__main__":
    main()
