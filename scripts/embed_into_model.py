#!/usr/bin/env python3
"""embed_into_model.py — graft the supply-chart tabs into an underwriting model.

Inserts the sheets from a generated supply-chart workbook (build with
`--model-link` so the subject rows reference the model's `Cash Flow (Annual)`)
into an existing `.xlsm`/`.xlsx` model **at the package/XML level**, so every
existing part — VBA, charts, conditional formatting, pivot caches, external
links — is preserved byte-for-byte. openpyxl is deliberately NOT used to write
the model: it silently drops features it doesn't understand.

What it edits, minimally:
  - styles.xml         (appends the chart's fonts/fills/borders/numFmts/cellXfs,
                        remapping the chart sheets' style indices to match)
  - workbook.xml       (adds the new <sheet> entries; sets fullCalcOnLoad)
  - workbook.xml.rels  (adds worksheet relationships; drops the calcChain rel)
  - [Content_Types].xml(adds the new sheet overrides; drops the calcChain one)
  - xl/calcChain.xml   (removed — Excel rebuilds it on open)

The chart workbook must store strings inline and carry no drawings/hyperlinks
(openpyxl's default output for this skill does). Run the validation at the end
and always open the result in Excel once before relying on it.

Usage:
    python embed_into_model.py \
        --chart  output/Aura__Supply_Chart_linked.xlsx \
        --model  "TMG Acquisition Model.xlsm" \
        --out    "TMG Acquisition Model__with_Supply_Chart.xlsm" \
        [--sheets "Competitive Analysis,Supply & Absorption,Reconciliation Log,Diligence"]
"""
import argparse
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
ET.register_namespace("", MAIN)
WS_CT = ("application/vnd.openxmlformats-officedocument."
         "spreadsheetml.worksheet+xml")
WS_REL = ("http://schemas.openxmlformats.org/officeDocument/"
          "2006/relationships/worksheet")


def _inner(xml, tag):
    m = re.search(rf"<{tag}(?:\s[^>]*)?>(.*?)</{tag}>", xml, re.S)
    return m.group(1) if m else ""


def _count(xml, tag):
    m = re.search(rf"<{tag}\s+count=\"(\d+)\"", xml)
    return int(m.group(1)) if m else 0


def _attr(s, name):
    m = re.search(rf'{name}="([^"]+)"', s)
    return m.group(1) if m else None


def merge_styles(model_styles, chart_styles):
    """Append the chart's style records to the model's; return (new_xml, info).

    info = (xf_offset, fmt_map) used to rewrite the chart sheets' style indices.
    """
    m_fonts = _count(model_styles, "fonts")
    m_fills = _count(model_styles, "fills")
    m_borders = _count(model_styles, "borders")
    m_cellxfs = _count(model_styles, "cellXfs")
    model_fmt_ids = [int(x) for x in re.findall(r'<numFmt numFmtId="(\d+)"', model_styles)]
    m_maxfmt = max(model_fmt_ids) if model_fmt_ids else 163

    c_numfmts = _inner(chart_styles, "numFmts")
    c_fonts = _inner(chart_styles, "fonts")
    c_fills = _inner(chart_styles, "fills")
    c_borders = _inner(chart_styles, "borders")
    c_cellxfs = _inner(chart_styles, "cellXfs")
    n_fonts = _count(chart_styles, "fonts")
    n_fills = _count(chart_styles, "fills")
    n_borders = _count(chart_styles, "borders")
    n_cellxfs = _count(chart_styles, "cellXfs")

    # remap chart custom numFmt ids (>=164) to fresh ids above the model's max
    fmt_map, nxt = {}, m_maxfmt + 1
    for fid in (int(x) for x in re.findall(r'numFmtId="(\d+)"', c_numfmts)):
        if fid >= 164 and fid not in fmt_map:
            fmt_map[fid] = nxt
            nxt += 1
    for old, new in fmt_map.items():
        c_numfmts = c_numfmts.replace(f'numFmtId="{old}"', f'numFmtId="{new}"')

    def remap_xf(xf):
        xf = re.sub(r'fontId="(\d+)"', lambda m: f'fontId="{int(m.group(1))+m_fonts}"', xf)
        xf = re.sub(r'fillId="(\d+)"', lambda m: f'fillId="{int(m.group(1))+m_fills}"', xf)
        xf = re.sub(r'borderId="(\d+)"', lambda m: f'borderId="{int(m.group(1))+m_borders}"', xf)
        xf = re.sub(r'numFmtId="(\d+)"',
                    lambda m: f'numFmtId="{fmt_map.get(int(m.group(1)), int(m.group(1)))}"', xf)
        return xf

    xf_items = re.findall(r"<xf\b[^>]*?/>|<xf\b[^>]*?>.*?</xf>", c_cellxfs, re.S)
    if len(xf_items) != n_cellxfs:
        raise SystemExit(f"cellXfs split mismatch: {len(xf_items)} vs {n_cellxfs}")
    remapped = "".join(remap_xf(x) for x in xf_items)

    def bump(xml, tag, delta):
        return re.sub(rf'(<{tag}\s+count=")(\d+)"',
                      lambda m: f'{m.group(1)}{int(m.group(2))+delta}"', xml, count=1)

    s = model_styles
    s = s.replace("</numFmts>", c_numfmts + "</numFmts>", 1)
    s = bump(s, "numFmts", len(fmt_map))
    s = s.replace("</fonts>", c_fonts + "</fonts>", 1)
    s = bump(s, "fonts", n_fonts)
    s = s.replace("</fills>", c_fills + "</fills>", 1)
    s = bump(s, "fills", n_fills)
    s = s.replace("</borders>", c_borders + "</borders>", 1)
    s = bump(s, "borders", n_borders)
    s = s.replace("</cellXfs>", remapped + "</cellXfs>", 1)
    s = bump(s, "cellXfs", n_cellxfs)
    return s, (m_cellxfs, fmt_map, (n_fonts, n_fills, n_borders, n_cellxfs))


def offset_sheet(xmlbytes, xf_offset):
    root = ET.fromstring(xmlbytes)
    for el in root.iter():
        tag = el.tag.split("}")[-1]
        if tag in ("c", "row") and "s" in el.attrib:
            el.set("s", str(int(el.attrib["s"]) + xf_offset))
        if tag == "col" and "style" in el.attrib:
            el.set("style", str(int(el.attrib["style"]) + xf_offset))
    return (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            + ET.tostring(root, encoding="utf-8"))


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--chart", required=True)
    ap.add_argument("--model", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--sheets", default=None,
                    help="Comma-separated sheet names to graft (default: all "
                         "sheets in the chart workbook).")
    args = ap.parse_args(argv)

    zc = zipfile.ZipFile(args.chart)
    zm = zipfile.ZipFile(args.model)

    # map chart sheet name -> source part
    cwb = zc.read("xl/workbook.xml").decode("utf-8")
    crels = zc.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rel = {}
    for r in re.findall(r"<Relationship\b[^>]*/>", crels):
        rel[_attr(r, "Id")] = _attr(r, "Target")
    name2part, order = {}, []
    for sh in re.findall(r"<sheet\b[^>]*/>", cwb):
        nm = _attr(sh, "name").replace("&amp;", "&")
        tgt = rel[_attr(sh, "r:id")]
        name2part[nm] = tgt.lstrip("/") if tgt.startswith("/") else "xl/" + tgt
        order.append(nm)
    graft = [s.strip() for s in args.sheets.split(",")] if args.sheets else order
    for nm in graft:
        if nm not in name2part:
            raise SystemExit(f"sheet not in chart workbook: {nm!r}")

    styles_new, (xf_offset, fmt_map, added) = merge_styles(
        zm.read("xl/styles.xml").decode("utf-8"),
        zc.read("xl/styles.xml").decode("utf-8"))

    new_parts = {}
    for i, nm in enumerate(graft, start=1):
        new_parts[f"xl/worksheets/sheetSC{i}.xml"] = offset_sheet(
            zc.read(name2part[nm]), xf_offset)

    mwb = zm.read("xl/workbook.xml").decode("utf-8")
    mrels = zm.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    mct = zm.read("[Content_Types].xml").decode("utf-8")
    max_sid = max(int(x) for x in re.findall(r'sheetId="(\d+)"', mwb))
    max_rid = max(int(x) for x in re.findall(r'Id="rId(\d+)"', mrels))

    sheets_xml = rels_xml = ct_xml = ""
    for i, nm in enumerate(graft, start=1):
        sid, rid = max_sid + i, f"rId{max_rid + i}"
        sheets_xml += f'<sheet name="{nm.replace("&", "&amp;")}" sheetId="{sid}" r:id="{rid}"/>'
        rels_xml += f'<Relationship Id="{rid}" Type="{WS_REL}" Target="worksheets/sheetSC{i}.xml"/>'
        ct_xml += f'<Override PartName="/xl/worksheets/sheetSC{i}.xml" ContentType="{WS_CT}"/>'

    mwb = mwb.replace("</sheets>", sheets_xml + "</sheets>", 1)
    if "fullCalcOnLoad" not in mwb:
        if "<calcPr" in mwb:
            mwb = re.sub(r"<calcPr\b", '<calcPr fullCalcOnLoad="1" ', mwb, count=1)
        else:
            mwb = mwb.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>', 1)
    mrels = mrels.replace("</Relationships>", rels_xml + "</Relationships>", 1)
    mrels = re.sub(r"<Relationship[^>]*calcChain[^>]*/>", "", mrels)
    mct = mct.replace("</Types>", ct_xml + "</Types>", 1)
    mct = re.sub(r"<Override[^>]*calcChain[^>]*/>", "", mct)

    drop = {"xl/calcChain.xml"}
    repl = {"xl/styles.xml": styles_new.encode("utf-8"),
            "xl/workbook.xml": mwb.encode("utf-8"),
            "xl/_rels/workbook.xml.rels": mrels.encode("utf-8"),
            "[Content_Types].xml": mct.encode("utf-8")}
    with zipfile.ZipFile(args.out, "w", zipfile.ZIP_DEFLATED) as zo:
        for item in zm.infolist():
            if item.filename in drop:
                continue
            zo.writestr(item, repl.get(item.filename, zm.read(item.filename)))
        for part, data in new_parts.items():
            zo.writestr(part, data)

    # validation: well-formedness + all retained model parts byte-identical
    # (fresh handles — re-reading a heavily-iterated ZipFile can misfire)
    zmv = zipfile.ZipFile(args.model)
    zo = zipfile.ZipFile(args.out)
    bad = zo.testzip()
    for n in ["xl/styles.xml", "xl/workbook.xml"] + list(new_parts):
        ET.fromstring(zo.read(n))
    changed = sorted(n for n in (set(zmv.namelist()) & set(zo.namelist()))
                     if zmv.read(n) != zo.read(n))
    print(f"Embedded {len(graft)} sheet(s) into {args.out}")
    print(f"  grafted: {graft}")
    print(f"  styles appended: +{added[0]} fonts, +{added[1]} fills, "
          f"+{added[2]} borders, +{added[3]} cellXfs; numFmt remap {fmt_map}")
    print(f"  testzip: {bad or 'OK'}  |  original parts changed: {changed}  "
          f"|  calcChain dropped (Excel rebuilds)")
    print("  Open once in Excel to confirm; links recalc on load.")


if __name__ == "__main__":
    sys.exit(main())
