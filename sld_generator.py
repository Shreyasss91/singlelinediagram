#!/usr/bin/env python3
"""
MESCOM SLD generator
====================
Reads the DTC master workbook (sheets "ALL" + "HT") and writes an editable
diagrams.net / draw.io single line diagram for one feeder.

    python sld_generator.py --list
    python sld_generator.py --feeder "Mandli F13" --preview
    python sld_generator.py --template

The sheet is a schematic: a distributor down the left and every DTC and HT
service on equally spaced laterals, so the drawing stays legible regardless of
where the transformers sit on the ground.

Output: "<name> SLD.drawio". Open it in draw.io / diagrams.net, or in VS Code
with the "Draw.io Integration" extension. Drag a DTC and its line follows.

The blank A3 sheet and symbol library live in "template/".
"""

from __future__ import annotations

import argparse
import math
import os
import re
import sys
from collections import defaultdict
from datetime import date
from xml.sax.saxutils import quoteattr

import openpyxl

# ----------------------------------------------------------------------------
# defaults - edit once if the master workbook moves
# ----------------------------------------------------------------------------
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
MASTER_DEFAULT = os.path.join(SCRIPT_DIR, "DTC DETAILS U4.xlsx")
TEMPLATE_DIR = os.path.join(SCRIPT_DIR, "template")
TEMPLATE_DRAWIO = os.path.join(TEMPLATE_DIR, "MESCOM_SLD_TEMPLATE.drawio")
SHEET_ALL = "ALL"
SHEET_HT = "HT"

# A3 landscape in draw.io units (px at 96 dpi)
PAGE_W, PAGE_H = 1654, 1169

MARGIN_L = 40
DIST_X = 110            # x of the vertical distributor
LATERAL_L = 150
LATERAL_R = 1610
FIRST_ROW_Y = 230
ROW_H_MAX = 230
TABLE_TOP = 1030

SLOT_MIN = 92
DTC_W, DTC_H = 44, 22
LBL_W, LBL_H = 100, 26
MUSS_W, MUSS_H = 170, 56

# ---- draw.io styles --------------------------------------------------------
ST_DTC = ("rounded=0;whiteSpace=wrap;html=1;fillColor=#dae8fc;strokeColor=#6c8ebf;"
          "fontSize=9;fontStyle=1;fontColor=#1a1a1a;")
ST_HT = ("rounded=0;whiteSpace=wrap;html=1;fillColor=#ffffff;strokeColor=#6c8ebf;"
         "fontSize=9;fontStyle=1;")
ST_LBL = ("text;html=1;align=center;verticalAlign=top;spacingTop=0;spacingLeft=0;"
          "spacingRight=0;fontSize=7;whiteSpace=wrap;fontColor=#1a1a1a;")
ST_MUSS = ("rounded=1;whiteSpace=wrap;html=1;fillColor=#ffffff;strokeColor=#6c8ebf;"
           "strokeWidth=1.5;fontSize=11;fontStyle=1;arcSize=18;")
ST_RMU = ("rounded=1;whiteSpace=wrap;html=1;fillColor=#ffffff;strokeColor=#6c8ebf;"
          "strokeWidth=1.2;fontSize=9;fontStyle=1;arcSize=10;")
ST_LINE = "endArrow=none;html=1;rounded=0;strokeColor=#6c8ebf;strokeWidth=1;"
ST_TRUNK = "endArrow=none;html=1;rounded=0;strokeColor=#6c8ebf;strokeWidth=1.5;"
ST_GEO = "endArrow=none;html=1;rounded=0;strokeColor=#6c8ebf;strokeWidth=1.3;edgeStyle=none;"
ST_STUB = "endArrow=none;html=1;rounded=0;strokeColor=#6c8ebf;strokeWidth=0.8;edgeStyle=none;"
ST_DASH = ("endArrow=classic;html=1;rounded=0;dashed=1;strokeColor=#6c8ebf;"
           "strokeWidth=1;edgeStyle=none;")
ST_JUNC = "ellipse;whiteSpace=wrap;html=1;fillColor=#6c8ebf;strokeColor=#6c8ebf;"
ST_HIDDEN = "ellipse;whiteSpace=wrap;html=1;fillColor=none;strokeColor=none;"
ST_SWITCH = "rounded=0;whiteSpace=wrap;html=1;fillColor=#6c8ebf;strokeColor=none;rotation=45;"
ST_BUS = "rounded=0;whiteSpace=wrap;html=1;fillColor=#6c8ebf;strokeColor=none;"
ST_ARROW = ("triangle;whiteSpace=wrap;html=1;fillColor=#6c8ebf;strokeColor=none;"
            "direction=north;")
ST_BOX = ("rounded=0;whiteSpace=wrap;html=1;fillColor=none;strokeColor=#999999;"
          "dashed=1;fontSize=8;align=left;verticalAlign=top;spacingLeft=6;spacingTop=4;")
ST_CELL = ("rounded=0;whiteSpace=wrap;html=1;fillColor=none;strokeColor=#000000;"
           "fontSize=8;align=center;verticalAlign=middle;")
ST_CELL_H = ("rounded=0;whiteSpace=wrap;html=1;fillColor=#f2f2f2;strokeColor=#000000;"
             "fontSize=8;fontStyle=1;align=center;verticalAlign=middle;")
ST_TITLE = ("text;html=1;align=center;verticalAlign=middle;fontSize=14;fontStyle=1;"
            "whiteSpace=wrap;")
ST_SUB = "text;html=1;align=center;verticalAlign=middle;fontSize=9;whiteSpace=wrap;"
ST_NOTE = "text;html=1;align=left;verticalAlign=top;fontSize=9;whiteSpace=wrap;"


# ----------------------------------------------------------------------------
# data helpers
# ----------------------------------------------------------------------------
def _txt(v) -> str:
    return "" if v is None else str(v).strip()


def _num(v):
    try:
        return float(str(v).replace(",", "").strip())
    except (TypeError, ValueError):
        return None


def _kva(v):
    s = _txt(v).upper().replace("KVA", "").strip()
    if not s:
        return None
    try:
        return int(float(s))
    except ValueError:
        return None


def crop_name(name: str, limit: int = 22) -> str:
    name = re.sub(r"\s+", " ", name).strip()
    return name if len(name) <= limit else name[: limit - 1].rstrip() + "."


def read_master(path: str):
    """Return (dtcs, hts). Codes normalised to SCD-01 / HT-87 style."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    if SHEET_ALL not in wb.sheetnames:
        raise SystemExit(f"Sheet '{SHEET_ALL}' not found in {path}. Found: {wb.sheetnames}")

    dtcs, seen = [], set()
    for row in wb[SHEET_ALL].iter_rows(min_row=2, values_only=True):
        row = list(row) + [None] * 14
        prefix = _txt(row[1]).upper() or "SCD"
        num = _txt(row[2])
        name = _txt(row[4]) or _txt(row[3])
        if not (num or name):
            continue
        code = f"{prefix}-{int(num):02d}" if num.isdigit() else f"{prefix}-{num}"
        if code.upper() in seen:
            continue
        seen.add(code.upper())
        dtcs.append(dict(code=code, name=name or "(unnamed)", kva=_kva(row[8]),
                         feeder=_txt(row[5]), station=_txt(row[9]), kind="DTC",
                         lat=_num(row[29]), lon=_num(row[30])))

    hts = []
    if SHEET_HT in wb.sheetnames:
        for row in wb[SHEET_HT].iter_rows(min_row=1, values_only=True):
            row = list(row) + [None] * 12
            code = _txt(row[4])
            if not re.fullmatch(r"HT-\d+", code, re.I):
                continue
            hts.append(dict(code=code.upper(), name=_txt(row[5]), kva=_kva(row[7]),
                            feeder=_txt(row[11]), kind="HT", lat=None, lon=None))
    return dtcs, hts


def _norm(s: str) -> str:
    return re.sub(r"[\s\-]+", " ", s.strip().lower())


def pick(items, feeder: str):
    want = _norm(feeder)
    return [it for it in items if _norm(it["feeder"]) == want]


# ----------------------------------------------------------------------------
# diagram model
# ----------------------------------------------------------------------------
class Page:
    def __init__(self, name: str):
        self.name = name
        self.cells: list[str] = []
        self._n = 0

    def _id(self, prefix="c"):
        self._n += 1
        return f"{prefix}{self._n}"

    def vertex(self, value, x, y, w, h, style, cid=None):
        cid = cid or self._id("n")
        self.cells.append(
            f'<mxCell id="{cid}" value={quoteattr(str(value))} style="{style}" vertex="1" parent="1">'
            f'<mxGeometry x="{x:.1f}" y="{y:.1f}" width="{w:.1f}" height="{h:.1f}" as="geometry"/></mxCell>'
        )
        return cid

    def edge(self, src, tgt, style=ST_LINE, points=None):
        cid = self._id("e")
        geom = '<mxGeometry relative="1" as="geometry">'
        if points:
            pts = "".join(f'<mxPoint x="{px:.1f}" y="{py:.1f}"/>' for px, py in points)
            geom += f'<Array as="points">{pts}</Array>'
        geom += "</mxGeometry>"
        self.cells.append(
            f'<mxCell id="{cid}" style="{style}" edge="1" parent="1" source="{src}" target="{tgt}">{geom}</mxCell>'
        )
        return cid


class Workbook:
    def __init__(self):
        self.pages: list[Page] = []

    def page(self, name: str) -> Page:
        p = Page(name)
        self.pages.append(p)
        return p

    def to_xml(self) -> str:
        out = ['<mxfile host="app.diagrams.net" agent="MESCOM SLD generator" version="24.7.17">']
        for i, p in enumerate(self.pages):
            body = "\n        ".join(p.cells)
            out.append(f'  <diagram id="p{i}" name={quoteattr(p.name[:60])}>')
            out.append(
                '    <mxGraphModel dx="1600" dy="1100" grid="0" gridSize="10" guides="1" tooltips="1" '
                'connect="1" arrows="1" fold="1" page="1" pageScale="1" '
                f'pageWidth="{PAGE_W}" pageHeight="{PAGE_H}" math="0" shadow="0">'
            )
            out.append("      <root>")
            out.append('        <mxCell id="0"/>')
            out.append('        <mxCell id="1" parent="0"/>')
            out.append(f"        {body}")
            out.append("      </root>")
            out.append("    </mxGraphModel>")
            out.append("  </diagram>")
        out.append("</mxfile>")
        return "\n".join(out) + "\n"


# ----------------------------------------------------------------------------
# shared sheet furniture
# ----------------------------------------------------------------------------
def sheet_header(d: Page, title: str, subtitle: str):
    d.vertex(title, MARGIN_L, 34, PAGE_W - 2 * MARGIN_L, 30, ST_TITLE)
    d.vertex(subtitle, MARGIN_L, 64, PAGE_W - 2 * MARGIN_L, 20, ST_SUB)


def summary_table(d: Page, dtcs, y0=1060):
    counts = {}
    for it in dtcs:
        counts[it["kva"]] = counts.get(it["kva"], 0) + 1
    cols = [25, 63, 100, 200, 250]
    cols += sorted(k for k in counts if k not in cols)
    tw, th = 62, 20
    x0 = PAGE_W - MARGIN_L - tw * (len(cols) + 2)
    d.vertex("kVA rating  →  number of DTCs", x0, y0 - 16, tw * 3, 16, ST_SUB)
    d.vertex("TOTAL", x0, y0 + th, tw, th, ST_CELL_H)
    for i, k in enumerate(cols):
        d.vertex(str(k), x0 + tw * (i + 1), y0, tw, th, ST_CELL_H)
        d.vertex(str(counts.get(k, 0)), x0 + tw * (i + 1), y0 + th, tw, th, ST_CELL)
    xt = x0 + tw * (len(cols) + 1)
    d.vertex("ALL", xt, y0, tw, th, ST_CELL_H)
    d.vertex(str(len(dtcs)), xt, y0 + th, tw, th, ST_CELL_H)


# ----------------------------------------------------------------------------
# layout A: schematic (distributor + equally spaced laterals)
# ----------------------------------------------------------------------------
def choose_grid(n: int):
    usable = LATERAL_R - LATERAL_L
    per_row = max(4, min(14, math.ceil(n / 3))) if n else 1
    while per_row > 1 and usable / per_row < SLOT_MIN:
        per_row -= 1
    rows = max(1, math.ceil(n / per_row))
    row_h = ROW_H_MAX if rows == 1 else max(
        120, min(ROW_H_MAX, (TABLE_TOP - 100 - FIRST_ROW_Y) / (rows - 1)))
    return per_row, rows, row_h


def build_schematic(dtcs, hts, label, station, title_text=None, subtitle=None):
    wb = Workbook()
    d = wb.page(label or "SLD")
    n = len(dtcs) + len(hts)
    per_row, rows, row_h = choose_grid(n)
    slot = (LATERAL_R - LATERAL_L) / per_row

    title = title_text or f"SINGLE LINE DIAGRAM OF {label.upper()} in Unit-4, CSD-2, SMG"
    sheet_header(d, title, subtitle or (
        f"{station}  MUSS   |   generated from DTC DETAILS U4.xlsx   |   {date.today():%d-%m-%Y}"))

    muss_x = 1370
    muss = d.vertex(f"110/11 KV<br>{station.replace('110/11KV ', '') or 'MUSS'}",
                    muss_x, 100, MUSS_W, MUSS_H, ST_MUSS)

    junctions = []
    first = d.vertex("", DIST_X, FIRST_ROW_Y, 4, 4, ST_JUNC)
    junctions.append(first)
    d.edge(muss, first, ST_TRUNK, points=[(muss_x + MUSS_W / 2, 176), (DIST_X + 2, 176)])
    for r in range(1, rows):
        j = d.vertex("", DIST_X, FIRST_ROW_Y + r * row_h, 4, 4, ST_JUNC)
        d.edge(junctions[-1], j, ST_TRUNK)
        junctions.append(j)

    for idx, item in enumerate(list(dtcs) + list(hts)):
        r, c = divmod(idx, per_row)
        y = FIRST_ROW_Y + r * row_h
        x = LATERAL_L + c * slot + SLOT_MIN / 2
        j = d.vertex("", x - 2, y, 4, 4, ST_JUNC)
        d.edge(junctions[r] if c == 0 else prev_j, j, ST_LINE)
        prev_j = j
        kind = item.get("kind", "DTC")
        name = crop_name(item["name"])
        if kind == "HT":
            tail = f" ({item['kva']} kVA)" if item.get("kva") else ""
            label_txt, face, w_item, style = f"{name}<br><b>{item['code']}</b>{tail}", "HT", DTC_W, ST_HT
        elif kind == "RMU":
            label_txt, face, w_item, style = f"<b>{crop_name(item['name'], 24)}</b>", "RMU", 90, ST_RMU
        else:
            label_txt = f"{name}<br><b>{item['code']}</b>"
            face = str(item["kva"]) if item.get("kva") else "?"
            w_item, style = DTC_W, ST_DTC
        sq = d.vertex(face, x - w_item / 2, y + 30, w_item, DTC_H, style)
        d.edge(j, sq, ST_LINE)
        d.vertex(label_txt, x - LBL_W / 2, y + 54, LBL_W, LBL_H, ST_LBL)

    summary_table(d, dtcs)
    return wb, dict(per_row=per_row, rows=rows, row_h=row_h, scale=None)


# ----------------------------------------------------------------------------
# template: blank A3 sheet + symbol library
# ----------------------------------------------------------------------------
def build_template() -> Workbook:
    wb = Workbook()
    d = wb.page("A3 SHEET")
    sheet_header(d, "SINGLE LINE DIAGRAM OF <FEEDER NAME> in Unit-4, CSD-2, SMG",
                 "110/11 KV  MUSS   |   <date>")
    muss = d.vertex("110/11 KV<br>MUSS", 1370, 100, MUSS_W, MUSS_H, ST_MUSS)

    ys = [FIRST_ROW_Y + i * 170 for i in range(5)]
    js = []
    for i, y in enumerate(ys):
        j = d.vertex("", DIST_X, y, 4, 4, ST_JUNC)
        if i == 0:
            d.edge(muss, j, ST_TRUNK, points=[(1370 + MUSS_W / 2, 176), (DIST_X + 2, 176)])
        else:
            d.edge(js[-1], j, ST_TRUNK)
        js.append(j)
        end = d.vertex("", LATERAL_R, y, 4, 4, ST_JUNC)
        d.edge(j, end, ST_LINE)
    summary_table(d, [])

    lib = wb.page("LIBRARY")
    lib.vertex("MESCOM SYMBOL LIBRARY  -  copy a symbol, paste it on the A3 sheet",
               MARGIN_L, 30, 900, 24, ST_TITLE)

    def caption(text, x, y):
        lib.vertex(text, x - 60, y, 200, 26, ST_NOTE)

    lib.vertex("110/11 KV<br>MANDLI", 90, 90, MUSS_W, MUSS_H, ST_MUSS)
    caption("MUSS / source", 90, 150)
    for i, kva in enumerate([25, 63, 100, 200, 250]):
        x = 90 + i * 80
        lib.vertex(str(kva), x, 220, DTC_W, DTC_H, ST_DTC)
        lib.vertex(f"<b>DTC {kva} kVA</b>", x - 28, 248, 100, 16, ST_LBL)
    caption("Distribution transformer - rating inside the square", 90, 275)
    lib.vertex("HT", 90, 330, DTC_W, DTC_H, ST_HT)
    lib.vertex("<b>HT service</b>", 62, 358, 100, 16, ST_LBL)
    caption("HT consumer (kVA shown in the label)", 90, 380)
    lib.vertex("", 100, 440, 30, 3, ST_SWITCH)
    lib.vertex("<b>GOS</b>", 60, 455, 100, 16, ST_LBL)
    caption("Gang operated switch", 90, 478)
    lib.vertex("RMU", 90, 530, 90, 24, ST_RMU)
    caption("Ring main unit", 90, 560)
    lib.vertex("", 96, 620, 4, 4, ST_JUNC)
    lib.vertex("<b>tee / junction</b>", 40, 630, 120, 16, ST_LBL)
    lib.vertex("", 60, 690, 160, 2, ST_BUS)
    lib.vertex("<b>lateral / bus</b>", 40, 700, 120, 16, ST_LBL)
    lib.vertex("How to use:  open this file in draw.io (or VS Code + Draw.io Integration).<br>"
               "1. Go to the A3 SHEET tab.<br>"
               "2. Copy a symbol from this LIBRARY tab and paste it on the sheet.<br>"
               "3. Join symbols with connectors (drag from the blue arrow until it turns green) - "
               "the line then follows the symbol when you move it.<br>"
               "4. Run  python sld_generator.py --feeder \"Mandli F13\" --geo  for a sheet placed by GPS.<br>"
               "5. File > Export as > PDF (A3 landscape) for the signed copy.",
               MARGIN_L, 780, 760, 140, ST_NOTE)
    return wb


# ----------------------------------------------------------------------------
# preview (optional) - draws the same coordinates with PyMuPDF
# ----------------------------------------------------------------------------
def preview(xml: str, out_png: str, page_idx: int = 0):
    try:
        import fitz
    except ImportError:
        print("  (preview skipped: pip install pymupdf)", file=sys.stderr)
        return
    import xml.etree.ElementTree as ET

    root = ET.fromstring(xml)
    model = root.findall(".//mxGraphModel")[page_idx]
    scale = 1190.0 / PAGE_W
    doc = fitz.open()
    page = doc.new_page(width=1190, height=842)
    shapes, edges = {}, []

    for c in model.iter("mxCell"):
        g = c.find("mxGeometry")
        if g is None:
            continue
        if c.get("vertex") == "1":
            shapes[c.get("id")] = (float(g.get("x")), float(g.get("y")),
                                   float(g.get("width")), float(g.get("height")),
                                   c.get("value") or "", c.get("style") or "")
        elif c.get("edge") == "1":
            pts = [(float(p.get("x")), float(p.get("y")))
                   for p in g.findall("./Array/mxPoint")]
            edges.append((c.get("source"), c.get("target"), pts))

    for s, t, pts in edges:
        if not (s in shapes and t in shapes):
            continue
        a, b = shapes[s], shapes[t]
        chain = [(a[0] + a[2] / 2, a[1] + a[3] / 2)] + pts + [(b[0] + b[2] / 2, b[1] + b[3] / 2)]
        for p, q in zip(chain, chain[1:]):
            page.draw_line(fitz.Point(p[0] * scale, p[1] * scale),
                           fitz.Point(q[0] * scale, q[1] * scale),
                           color=(0.42, 0.56, 0.75), width=0.6)

    for x, y, w, h, val, style in shapes.values():
        r = fitz.Rect(x * scale, y * scale, (x + w) * scale, (y + h) * scale)
        lines = [re.sub(r"<[^>]+>", "", ln) for ln in re.split(r"<br\s*/?>", val)]
        fs = re.search(r"fontSize=(\d+)", style)
        size = (int(fs.group(1)) if fs else 8) * scale
        rot = re.search(r"rotation=(-?\d+(?:\.\d+)?)", style)
        if rot:
            ang = math.radians(float(rot.group(1)))
            cx, cy = r.x0 + r.width / 2, r.y0 + r.height / 2
            dx, dy = math.cos(ang) * r.width / 2, -math.sin(ang) * r.width / 2
            page.draw_line((cx - dx, cy - dy), (cx + dx, cy + dy),
                           color=(0.42, 0.56, 0.75), width=1.2)
            continue
        if "text;" not in style:
            page.draw_rect(r, color=(0.42, 0.56, 0.75),
                           fill=(0.85, 0.91, 0.99) if "dae8fc" in style else None, width=0.7)
        for i, ln in enumerate(lines):
            if not ln.strip():
                continue
            tw = fitz.get_text_length(ln, fontsize=size)
            page.insert_text((r.x0 + (r.width - tw) / 2, r.y0 + size * (1.35 + 1.25 * i)),
                             ln, fontsize=size, color=(0.1, 0.1, 0.1))
    page.get_pixmap(dpi=120).save(out_png)
    print(f"  preview     : {out_png}")


# ----------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="Generate a MESCOM feeder SLD as an editable .drawio file.")
    ap.add_argument("--master", default=MASTER_DEFAULT, help="DTC master workbook")
    ap.add_argument("--feeder", help='feeder name as in the master, e.g. "Mandli F13"')
    ap.add_argument("--name", help="label printed on the sheet (default: feeder)")
    ap.add_argument("--title", help="override the title line")
    ap.add_argument("--out", help="output .drawio path")
    ap.add_argument("--preview", action="store_true", help="also write a PNG preview")
    ap.add_argument("--template", action="store_true",
                    help="write the blank A3 sheet + symbol library to template/")
    ap.add_argument("--list", action="store_true", help="list feeders found in the master")
    args = ap.parse_args()

    if args.template:
        out = args.out or TEMPLATE_DRAWIO
        os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
        xml = build_template().to_xml()
        with open(out, "w", encoding="utf-8") as fh:
            fh.write(xml)
        print(f"  template    : {out}")
        if args.preview:
            preview(xml, os.path.splitext(out)[0] + "_library.png", page_idx=1)
        return

    if not os.path.exists(args.master):
        raise SystemExit(
            f"Master workbook not found: {args.master}\n"
            f"Put 'DTC DETAILS U4.xlsx' next to this script ({SCRIPT_DIR}),\n"
            f"or pass the path with --master \"D:\\path\\to\\DTC DETAILS U4.xlsx\"")
    dtcs, hts = read_master(args.master)

    if args.list or not args.feeder:
        from collections import Counter
        cnt = Counter(it["feeder"] for it in dtcs)
        print("Feeders in the master:")
        for f, c in cnt.items():
            with_gps = sum(1 for it in dtcs if it["feeder"] == f and it["lat"] is not None)
            print(f"  {f:<14} {c:>3} DTCs  ({with_gps} with GPS)")
        if not args.feeder:
            print('\nNow run e.g.:  python sld_generator.py --feeder "Mandli F13"')
        return

    sel = pick(dtcs, args.feeder)
    sel_ht = pick(hts, args.feeder)
    if not sel:
        raise SystemExit(f"No DTCs for feeder {args.feeder!r}. Use --list to see the names.")
    sel.sort(key=lambda x: int(re.sub(r"\D", "", x["code"]) or 0))

    station = next((x["station"] for x in sel if x["station"]), "110/11KV MUSS")
    label = args.name or args.feeder

    wb, info = build_schematic(sel, sel_ht, label, station, title_text=args.title)

    xml = wb.to_xml()
    out = args.out or f"{label} SLD.drawio"
    with open(out, "w", encoding="utf-8") as fh:
        fh.write(xml)

    print(f"  feeder      : {label}  ({len(sel)} DTCs, {len(sel_ht)} HT services)")
    print(f"  layout      : schematic  |  {info['per_row']} per lateral x {info['rows']} laterals")
    print(f"  written     : {out}")
    if args.preview:
        preview(xml, os.path.splitext(out)[0] + ".png")


if __name__ == "__main__":
    main()
