#!/usr/bin/env python3
"""
MESCOM SLD generator
====================
Reads the DTC master workbook (sheets "ALL" + "HT") and writes an editable
diagrams.net / draw.io single line diagram for one feeder.

    python sld_generator.py --list
    python sld_generator.py --feeder "Mandli F13" --preview
    python sld_generator.py --feeder "MRS F1" --geo --preview
    python sld_generator.py --feeder "MRS F1" --geo --muss "13.9305,75.5842"
    python sld_generator.py --feeder "MRS F1" --geo --route poles.csv
    python sld_generator.py --template
    python sld_generator.py --topology-template

Two layouts:
  default  - distributor with equally spaced laterals (clean schematic)
  --geo    - every DTC at its surveyed GPS position, trunk inferred by a
             minimum spanning tree (or taken from --route pole coordinates)

Output: "outputs/<YYYY-MM-DD>/<name> SLD.drawio" (or "<name> - GEO SLD.drawio").
Open it in draw.io / diagrams.net, or in VS Code with the "Draw.io Integration"
extension. Drag a DTC and its line follows. Pass --outdir or --out to override.

Templates live in "template/": the blank A3 sheet + symbol library, the
surveyed-route sample, and the topology workbook to fill in.
"""

from __future__ import annotations

import argparse
import csv
import math
import os
import re
import sys
from collections import Counter, defaultdict
from datetime import date
from xml.sax.saxutils import quoteattr

import openpyxl
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

# ----------------------------------------------------------------------------
# defaults - edit once if the master workbook moves
# ----------------------------------------------------------------------------
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(SCRIPT_DIR, "field_data")
MASTER_DEFAULT = os.path.join(DATA_DIR, "DTC DETAILS U4.xlsx")
TEMPLATE_DIR = os.path.join(SCRIPT_DIR, "template")
OUTPUT_ROOT = os.path.join(SCRIPT_DIR, "outputs")
TEMPLATE_DRAWIO = os.path.join(TEMPLATE_DIR, "MESCOM_SLD_TEMPLATE.drawio")
TOPOLOGY_TEMPLATE = os.path.join(TEMPLATE_DIR, "TOPOLOGY TEMPLATE.xlsx")
SHEET_ALL = "ALL"
SHEET_HT = "HT"
SHEET_NODES = "NON DTC NODES"

# real substation coordinates: fill once, or pass --muss "lat,lon"
STATIONS = {
    "110/11KV MRS": None,
    "110/11KV MANDLI": None,
}

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

# drawing area for the geographical layout
GEO_X, GEO_Y, GEO_W, GEO_H = 120, 205, 1460, 735
GEO_MAX_SCALE = 3.2     # px per metre, so a tiny feeder does not fill the sheet

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


def _display(path: str) -> str:
    """Short path for console output: relative when inside the tool folder."""
    try:
        rel = os.path.relpath(os.path.abspath(path), SCRIPT_DIR)
    except ValueError:
        return path
    return path if rel.startswith("..") else rel


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


def read_nodes(wb):
    """Sheet 'NON DTC NODES': name, GPS LAT, GPS LONG - MUSS and RMU positions."""
    nodes = []
    if SHEET_NODES not in wb.sheetnames:
        return nodes
    for row in wb[SHEET_NODES].iter_rows(min_row=2, values_only=True):
        row = list(row) + [None] * 4
        name = _txt(row[0])
        lat, lon = _num(row[1]), _num(row[2])
        if not name or lat is None or lon is None:
            continue
        if abs(lat) > 90 or abs(lon) > 180:
            continue
        up = name.upper()
        kind = "MUSS" if "MUSS" in up else ("RMU" if "RMU" in up else "NODE")
        nodes.append(dict(name=name, kind=kind, lat=lat, lon=lon))
    return nodes


def find_station(nodes, station: str):
    """Match the sheet's station text (e.g. '110/11KV MRS') to a MUSS node name."""
    s = _norm(station)
    for n in nodes:
        if n["kind"] != "MUSS":
            continue
        nn = _norm(n["name"])
        if s in nn or nn in s:
            return (n["lat"], n["lon"])
    return None


def find_station_node(nodes, station: str):
    """Same match as find_station, but returns the whole node dict."""
    s = _norm(station)
    for n in nodes:
        if n["kind"] != "MUSS":
            continue
        nn = _norm(n["name"])
        if s in nn or nn in s:
            return n
    return None


def assign_nodes_to_feeders(nodes, dtcs):
    """An RMU belongs to the feeder whose nearest DTC is closest to it."""
    located = [d for d in dtcs if d.get("lat") is not None and d.get("lon") is not None]
    if not located:
        return nodes
    for n in nodes:
        if n["kind"] != "RMU":
            continue
        near = min(located, key=lambda d: (d["lat"] - n["lat"]) ** 2 +
                   ((d["lon"] - n["lon"]) * math.cos(math.radians(n["lat"]))) ** 2)
        n["feeder"] = near["feeder"]
    return nodes


def read_master(path: str):
    """Return (dtcs, hts, nodes). Codes normalised to SCD-01 / HT-87 style."""
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
    return dtcs, hts, read_nodes(wb)


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


def unpositioned_block(d: Page, bad):
    if not bad:
        return
    codes = ", ".join(f"{it['code']} {crop_name(it['name'], 16)}" for it in bad)
    x0, y0, w = 60, 1012, 1080
    lines = (w - 12) / 4.6
    h = 26 + 11 * max(1, math.ceil(len(codes) / lines))
    d.vertex(f"<b>UNPOSITIONED  -  no usable GPS in the master</b><br>{codes}",
             x0, y0, w, h, ST_BOX)


def scale_bar(d: Page, px_per_m, x=150, y=1120):
    target = None
    for cand in (25, 50, 100, 200, 250, 500, 1000):
        if 70 <= cand * px_per_m <= 240:
            target = cand
            break
    if target is None:
        target = 100
    w = target * px_per_m
    d.vertex("", x, y, w, 2, ST_BUS)
    d.vertex("", x, y - 9, 1, 9, ST_BUS)
    d.vertex("", x + w - 1, y - 9, 1, 9, ST_BUS)
    d.vertex(f"{target} m", x, y + 4, w, 14, ST_SUB)
    d.vertex("", x - 34, y - 22, 12, 16, ST_ARROW)
    d.vertex("N", x - 38, y - 6, 20, 14, ST_SUB)


def _overlap_area(a, b):
    w = min(a[2], b[2]) - max(a[0], b[0])
    h = min(a[3], b[3]) - max(a[1], b[1])
    return max(0.0, w) * max(0.0, h)


def _overlaps(a, b, pad=2):
    return not (a[2] + pad <= b[0] or b[2] + pad <= a[0] or
                a[3] + pad <= b[1] or b[3] + pad <= a[1])


def _clamp_label(x, y):
    x = min(max(x, 20), PAGE_W - 20 - LBL_W)
    y = min(max(y, 96), TABLE_TOP - 24 - LBL_H)
    return x, y


def place_label(d: Page, cx, cy, value, taken, bw=DTC_W, bh=DTC_H):
    """Put a label on the freest spot around the symbol at (cx, cy)."""
    candidates = []
    for dist in (0, 22, 44):
        candidates += [
            (cx - LBL_W / 2, cy + bh / 2 + 3 + dist),             # below
            (cx - LBL_W / 2, cy - bh / 2 - LBL_H - 3 - dist),     # above
            (cx + bw / 2 + 3 + dist, cy - LBL_H / 2),             # right
            (cx - bw / 2 - LBL_W - 3 - dist, cy - LBL_H / 2),     # left
        ]
    best, best_cost = None, None
    for cx0, cy0 in candidates:
        x, y = _clamp_label(cx0, cy0)
        r = (x, y, x + LBL_W, y + LBL_H)
        cost = sum(_overlap_area(r, t) for t in taken)
        if cost == 0:
            taken.append(r)
            return d.vertex(value, x, y, LBL_W, LBL_H, ST_LBL)
        if best_cost is None or cost < best_cost:
            best, best_cost = r, cost
    taken.append(best)
    return d.vertex(value, best[0], best[1], LBL_W, LBL_H, ST_LBL)


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


def build_schematic(dtcs, hts, label, station, rmus=(), title_text=None, subtitle=None):
    wb = Workbook()
    d = wb.page(label or "SLD")
    n = len(dtcs) + len(hts) + len(rmus)
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

    for idx, item in enumerate(list(dtcs) + list(hts) + list(rmus)):
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
            label_txt = (f"<b>{crop_name(item['name'], 24)}</b>"
                         "<br><font style='font-size:7px'>5-way</font>")
            face, w_item, style = "RMU", 90, ST_RMU
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
# layout B: geographical (GPS positions, inferred or surveyed trunk)
# ----------------------------------------------------------------------------
def clean_points(items):
    """Split into usable / unusable coordinates, flagging far outliers."""
    good, bad, warn = [], [], []
    for it in items:
        if it.get("lat") is None or it.get("lon") is None:
            bad.append(it)
        else:
            good.append(it)
    if len(good) >= 3:
        mlat = sorted(p["lat"] for p in good)[len(good) // 2]
        mlon = sorted(p["lon"] for p in good)[len(good) // 2]
        keep = []
        for p in good:
            d = math.hypot((p["lat"] - mlat) * 111.0,
                           (p["lon"] - mlon) * 111.0 * math.cos(math.radians(mlat)))
            if d > 5.0:
                warn.append(f"{p['code']}: {p['lat']}, {p['lon']} is {d:.0f} km from the "
                            f"rest of the feeder - ignored (check the master)")
                bad.append(p)
            else:
                keep.append(p)
        good = keep
    return good, bad, warn


def _metres(lat, lon, lat0, lon0):
    return ((lon - lon0) * 111320.0 * math.cos(math.radians(lat0)),
            (lat - lat0) * 110540.0)


def prim(parent_of, pts, root=0):
    """Minimum spanning tree over the DTC points; returns parent indices."""
    n = len(pts)
    in_tree = [False] * n
    parent = [-1] * n
    best = [float("inf")] * n
    best[root] = 0.0
    for _ in range(n):
        u = min((i for i in range(n) if not in_tree[i]), key=lambda i: best[i])
        in_tree[u] = True
        ux, uy = pts[u]
        for v in range(n):
            if in_tree[v]:
                continue
            d = math.hypot(pts[v][0] - ux, pts[v][1] - uy)
            if d < best[v]:
                best[v], parent[v] = d, u
    return parent


def tree_ends(parent):
    """Two ends of the tree diameter - used when no MUSS position is known."""
    adj = defaultdict(list)
    for i, p in enumerate(parent):
        if p >= 0:
            adj[i].append(p)
            adj[p].append(i)

    def far(start):
        dist = {start: 0}
        stack = [start]
        while stack:
            u = stack.pop()
            for v in adj[u]:
                if v not in dist:
                    dist[v] = dist[u] + 1
                    stack.append(v)
        return max(dist, key=dist.get)

    a = far(0)
    return a, far(a)


def read_route(path):
    """Read surveyed pole coordinates: NODE_ID, LAT, LONG (csv or xlsx).

    An optional 4th column, WAY, is ignored here: it names the RMU way a point
    passes through, which only the topology workbook uses. Repeating an RMU row
    (once as ODin, once as ODout) shows the run entering and leaving the unit.
    """
    pts = []
    if path.lower().endswith((".xlsx", ".xlsm")):
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        for row in wb.worksheets[0].iter_rows(values_only=True):
            row = list(row) + [None] * 4
            lat, lon = _num(row[1]), _num(row[2])
            if lat is not None and lon is not None and abs(lat) <= 90 and abs(lon) <= 180:
                pts.append((_txt(row[0]) or f"P{len(pts)}", lat, lon))
    else:
        with open(path, newline="", encoding="utf-8-sig") as fh:
            for row in csv.reader(fh):
                row = list(row) + [""] * 4
                lat, lon = _num(row[1]), _num(row[2])
                if lat is not None and lon is not None and abs(lat) <= 90 and abs(lon) <= 180:
                    pts.append((row[0].strip() or f"P{len(pts)}", lat, lon))
    return pts


def _project(px, py, ax, ay, bx, by):
    vx, vy = bx - ax, by - ay
    L = vx * vx + vy * vy
    t = 0.0 if L == 0 else max(0.0, min(1.0, ((px - ax) * vx + (py - ay) * vy) / L))
    x, y = ax + t * vx, ay + t * vy
    return math.hypot(px - x, py - y), x, y


def build_geo(dtcs, hts, label, station, muss=None, route=None, rmus=(),
              extent="cluster", title_text=None, subtitle=None):
    wb = Workbook()
    d = wb.page(label or "SLD")

    good, bad, warn = clean_points(list(dtcs) + list(hts))
    for w in warn:
        print("  ! " + w)
    if len(good) < 2:
        raise SystemExit("Not enough DTCs with GPS coordinates for a geographical layout.")

    lat0 = sorted(p["lat"] for p in good)[len(good) // 2]
    lon0 = sorted(p["lon"] for p in good)[len(good) // 2]
    local = {p["code"]: _metres(p["lat"], p["lon"], lat0, lon0) for p in good}
    order = [p["code"] for p in good]

    by_code = {p["code"]: p for p in good}

    # ---- trunk geometry ----------------------------------------------------
    trunk_pts = []       # surveyed polyline in metres, in drawing order
    parent = [-1] * len(order)
    root = None
    if route:
        trunk_pts = [_metres(lat, lon, lat0, lon0) for _, lat, lon in route]
        print(f"  route       : {len(trunk_pts)} surveyed points")
    else:
        pts = [local[c] for c in order]
        if muss:
            mx, my = _metres(muss[0], muss[1], lat0, lon0)
            root = min(range(len(pts)),
                       key=lambda i: math.hypot(pts[i][0] - mx, pts[i][1] - my))
        else:
            root = tree_ends(prim([], pts))[0]
        parent = prim([], pts, root)

    # ---- page transform ----------------------------------------------------
    def fit(xs, ys):
        mnx, mxx = min(xs), max(xs)
        mny, mxy = min(ys), max(ys)
        sx, sy = max(mxx - mnx, 1e-6), max(mxy - mny, 1e-6)
        return mnx, mxx, mny, mxy, min(GEO_W / sx, GEO_H / sy, GEO_MAX_SCALE)

    net_x = [local[c][0] for c in order]
    net_y = [local[c][1] for c in order]
    if trunk_pts:
        net_x += [p[0] for p in trunk_pts]
        net_y += [p[1] for p in trunk_pts]

    box_net = fit(net_x, net_y)
    minx, maxx, miny, maxy, scale = box_net
    off_map_muss = False
    if muss:
        mxm, mym = _metres(muss[0], muss[1], lat0, lon0)
        box_all = fit(net_x + [mxm], net_y + [mym])
        if extent == "all" or box_all[4] >= 0.7 * box_net[4]:
            minx, maxx, miny, maxy, scale = box_all
        else:
            # the substation is far outside the network: keep the network readable
            off_map_muss = True

    spanx, spany = max(maxx - minx, 1e-6), max(maxy - miny, 1e-6)
    ox = GEO_X + (GEO_W - spanx * scale) / 2
    oy = GEO_Y + (GEO_H - spany * scale) / 2

    def P(xy):
        x, y = xy
        return ox + (x - minx) * scale, oy + (maxy - y) * scale

    cx_pg = sum(P(local[c])[0] for c in order) / len(order)
    cy_pg = sum(P(local[c])[1] for c in order) / len(order)
    taken = []

    # ---- trunk -------------------------------------------------------------
    junctions = {}
    route_ids = []
    if trunk_pts:
        for pt in trunk_pts:
            jx, jy = P(pt)
            route_ids.append(d.vertex("", jx - 1.5, jy - 1.5, 3, 3, ST_HIDDEN))
        for a, b in zip(route_ids, route_ids[1:]):
            d.edge(a, b, ST_GEO)
    else:
        for code in order:
            jx, jy = P(local[code])
            junctions[code] = d.vertex("", jx - 1.5, jy - 1.5, 3, 3, ST_HIDDEN)
        for i, code in enumerate(order):
            if parent[i] >= 0:
                d.edge(junctions[order[parent[i]]], junctions[code], ST_GEO)

    # ---- MUSS --------------------------------------------------------------
    if trunk_pts:
        anchor_id = route_ids[0] if route_ids else None
        ref = trunk_pts[0]
    else:
        anchor_id = junctions.get(order[root]) if root is not None else None
        ref = local[order[root]] if root is not None else local[order[0]]

    # the MUSS box is drawn after the DTCs so it can be kept clear of them

    # ---- DTC symbols, spurs and labels -------------------------------------
    trunk_pg = [P(p) for p in trunk_pts]

    for i, code in enumerate(order):
        item = by_code[code]
        px, py = P(local[code])

        if trunk_pg:                       # surveyed route: use the local segment
            best = None
            for k in range(len(trunk_pg) - 1):
                dist, ax, ay = _project(px, py, *trunk_pg[k], *trunk_pg[k + 1])
                if best is None or dist < best[0]:
                    best = (dist, ax, ay, k)
            _, ax, ay, k = best
            dx = trunk_pg[k + 1][0] - trunk_pg[k][0]
            dy = trunk_pg[k + 1][1] - trunk_pg[k][1]
            anchor_pt = (ax, ay)
        else:                              # inferred tree: use the parent edge
            p = parent[i]
            if p >= 0:
                ax, ay = P(local[order[p]])
            else:
                ax, ay = cx_pg, cy_pg
            dx, dy = px - ax, py - ay
            anchor_pt = (ax, ay)

        n = math.hypot(dx, dy) or 1.0
        perp = (-dy / n, dx / n)
        if perp[0] * (px - cx_pg) + perp[1] * (py - cy_pg) < 0:
            perp = (-perp[0], -perp[1])
        cx, cy = px + perp[0] * 30, py + perp[1] * 30
        for _ in range(6):                 # separate DTCs sitting on the same spot
            rect = (cx - DTC_W / 2, cy - DTC_H / 2, cx + DTC_W / 2, cy + DTC_H / 2)
            if not any(_overlaps(rect, t) for t in taken):
                break
            cx += perp[0] * 26
            cy += perp[1] * 26

        is_ht = code.upper().startswith("HT-")
        face = "HT" if is_ht else (str(item["kva"]) if item.get("kva") else "?")
        sq = d.vertex(face, cx - DTC_W / 2, cy - DTC_H / 2, DTC_W, DTC_H,
                      ST_HT if is_ht else ST_DTC)

        anchor_id = junctions.get(code)
        if anchor_id is None:
            anchor_id = d.vertex("", anchor_pt[0] - 1.5, anchor_pt[1] - 1.5, 3, 3, ST_HIDDEN)
        d.edge(anchor_id, sq, ST_STUB)

        name = crop_name(item["name"])
        if is_ht:
            tail = f" ({item['kva']} kVA)" if item.get("kva") else ""
            txt = f"{name}<br><b>{code}</b>{tail}"
        else:
            txt = f"{name}<br><b>{code}</b>"
        taken.append((cx - DTC_W / 2, cy - DTC_H / 2, cx + DTC_W / 2, cy + DTC_H / 2))
        place_label(d, cx, cy, txt, taken)

    # ---- RMUs from the NON DTC NODES sheet --------------------------------
    for nd in rmus:
        rx, ry = P(_metres(nd["lat"], nd["lon"], lat0, lon0))
        rw, rh = 90, 24
        bx, by = rx, ry
        search = [(0.0, 0.0)]
        for dist in (26, 52, 78, 110):
            for step in range(8):
                ang = step * math.pi / 4
                search.append((math.cos(ang) * dist, math.sin(ang) * dist))
        for ddx, ddy in search:
            rect = (rx + ddx - rw / 2, ry + ddy - rh / 2,
                    rx + ddx + rw / 2, ry + ddy + rh / 2)
            if not any(_overlaps(rect, t) for t in taken):
                bx, by = rx + ddx, ry + ddy
                break
        node = d.vertex("RMU", bx - rw / 2, by - rh / 2, rw, rh, ST_RMU)
        if trunk_pg:
            best = None
            for k in range(len(trunk_pg) - 1):
                dist, ax, ay = _project(rx, ry, *trunk_pg[k], *trunk_pg[k + 1])
                if best is None or dist < best[0]:
                    best = (dist, ax, ay)
            _, ax, ay = best
        else:
            ax, ay = min((P(local[c]) for c in order),
                         key=lambda q: math.hypot(q[0] - rx, q[1] - ry))
        a_id = d.vertex("", ax - 1.5, ay - 1.5, 3, 3, ST_HIDDEN)
        d.edge(a_id, node, ST_STUB)
        taken.append((bx - rw / 2, by - rh / 2, bx + rw / 2, by + rh / 2))
        place_label(d, bx, by, f"<b>{nd['name']}</b><br>5-way", taken, bw=rw, bh=rh)

    # ---- MUSS --------------------------------------------------------------
    if anchor_id:
        base = (f"110/11 KV<br>{station.replace('110/11KV ', '') or 'MUSS'}")
        if muss and off_map_muss:
            bx, by = P(_metres(muss[0], muss[1], lat0, lon0))
            bx = min(max(bx, GEO_X + MUSS_W / 2 + 6), GEO_X + GEO_W - MUSS_W / 2 - 6)
            by = min(max(by, GEO_Y + MUSS_H / 2 + 6), GEO_Y + GEO_H - MUSS_H / 2 - 6)
            mxm, mym = _metres(muss[0], muss[1], lat0, lon0)
            km = math.hypot(mxm - sum(net_x) / len(net_x),
                            mym - sum(net_y) / len(net_y)) / 1000
            base += (f"<br><font style='font-size:7px'>({km:.1f} km beyond the edge, "
                     f"not to scale)</font>")
            box = (bx - MUSS_W / 2, by - MUSS_H / 2, bx + MUSS_W / 2, by + MUSS_H / 2 + 8)
        elif muss:
            mx, my = P(_metres(muss[0], muss[1], lat0, lon0))
            box = (mx - MUSS_W / 2, my - MUSS_H / 2, mx + MUSS_W / 2, my + MUSS_H / 2)
            cand = [(0.0, 0.0)]
            for dist in (36, 72, 110, 150, 200):
                for step in range(8):
                    ang = step * math.pi / 4
                    cand.append((math.cos(ang) * dist, math.sin(ang) * dist))
            for dx, dy in cand:
                cx0 = min(max(mx + dx, MUSS_W / 2 + 20), PAGE_W - MUSS_W / 2 - 20)
                cy0 = min(max(my + dy, 130 + MUSS_H / 2), TABLE_TOP - 96)
                r = (cx0 - MUSS_W / 2, cy0 - MUSS_H / 2,
                     cx0 + MUSS_W / 2, cy0 + MUSS_H / 2 + 8)
                if not any(_overlaps(r, t) for t in taken):
                    box = r
                    break
        else:
            ax, ay = P(ref)
            vx, vy = ax - cx_pg, ay - cy_pg
            nv = math.hypot(vx, vy) or 1.0
            base_ang = math.atan2(vy / nv, vx / nv)
            box = None
            for dist in (74, 112, 152, 200, 258, 320):
                for step in (0, 1, -1, 2, -2, 3, -3, 4):
                    ang = base_ang + step * math.pi / 8
                    mx = min(max(ax + math.cos(ang) * dist, MUSS_W / 2 + 20),
                             PAGE_W - MUSS_W / 2 - 20)
                    my = min(max(ay + math.sin(ang) * dist, 130 + MUSS_H / 2 + 8),
                             TABLE_TOP - 96)
                    r = (mx - MUSS_W / 2, my - MUSS_H / 2,
                         mx + MUSS_W / 2, my + MUSS_H / 2 + 8)
                    if box is None:
                        box = r
                    if not any(_overlaps(r, t) for t in taken):
                        box = r
                        break
                if not any(_overlaps(box, t) for t in taken):
                    break
            base += "<br><font style='font-size:7px'>(position approx.)</font>"
        m_id = d.vertex(base, box[0], box[1], MUSS_W, MUSS_H + 8, ST_MUSS)
        d.edge(m_id, anchor_id, ST_DASH if off_map_muss else ST_TRUNK)
        taken.append(box)

    # ---- furniture ---------------------------------------------------------
    unpositioned_block(d, bad)
    scale_bar(d, scale)
    summary_table(d, dtcs)
    sheet_header(d, title_text or
                 f"SINGLE LINE DIAGRAM OF {label.upper()} in Unit-4, CSD-2, SMG",
                 subtitle or (f"{station}  MUSS   |   geographical layout from GPS in the master"
                              f"   |   {'surveyed route' if route else 'inferred trunk'}"
                              f"   |   {date.today():%d-%m-%Y}"))
    return wb, dict(scale=scale, per_row=None, rows=None,
                    positioned=len(good), unpositioned=len(bad))


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

    def caption(text, x, y, h=26):
        lib.vertex(text, x - 60, y, 200, h, ST_NOTE)

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
    caption("Ring main unit, 5-way:<br>ODin, ODout, VL1, VL2, VL3", 90, 560, h=34)
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
# topology data template - the workbook the user fills in
# ----------------------------------------------------------------------------
NODE_HEADERS = [
    ("NODE_ID", 15), ("TYPE", 9), ("NAME", 34), ("FEEDER", 12),
    ("LAT", 12), ("LONG", 12), ("PARENT_ID", 15), ("SPUR_ID", 8), ("WAY", 9),
    ("KVA", 7), ("PHASE", 7), ("NORMALLY", 9), ("CONDUCTOR", 14),
    ("BASIS", 10), ("REMARKS", 52),
]

# An RMU is a five-way switchgear: one incoming way and four ways out.
# ODin is the way the RMU itself is fed through, so it is recorded by the RMU's
# own PARENT_ID; every child of an RMU names the way it leaves by in NODES!WAY.
# A way is a terminal, not a connection: only ODin must be cabled, and any of
# the four outgoing ways - ODout included - may be empty.
RMU_WAYS = [("ODin", "IN"), ("ODout", "OUT"), ("VL1", "OUT"),
            ("VL2", "OUT"), ("VL3", "OUT")]
RMU_WAY_HEADERS = [("RMU_ID", 15), ("WAY", 9), ("DIRECTION", 10), ("CONNECTED", 11),
                   ("NORMALLY", 10), ("LABEL", 30), ("NOTES", 56)]


def _way_cabled(way: str) -> bool:
    """Only the incoming way is always cabled; the other four may be empty."""
    return way == "ODin"


def _node_id(kind: str, name: str) -> str:
    """Stable short id: '110/11KV MRS MUSS' -> 'MUSS-MRS', 'KOTE ROAD RMU' -> 'RMU-KOTE-ROAD'."""
    core = re.sub(r"^110\s*/\s*11\s*kv\s*", "", name, flags=re.I)
    core = re.sub(rf"\b{re.escape(kind)}\b", "", core, flags=re.I)
    core = re.sub(r"[^A-Za-z0-9]+", "-", core).strip("-").upper()
    return f"{kind}-{core}" if core else kind


def _infer_topology(sel, muss_node, muss_id):
    """Draft PARENT_ID / SPUR_ID for one feeder from a minimum spanning tree.

    This is deliberately marked INFERRED: it groups DTCs onto plausible
    laterals, but only a survey or the hand-drawn SLD says what the line
    really does.
    """
    parents, spurs, spur_from, notes = {}, {}, {}, {}
    good, bad, _ = clean_points(sel)
    for p in bad:
        notes[p["code"]] = "no usable GPS - set PARENT_ID by hand"
    if not muss_node:
        for p in good:
            notes[p["code"]] = "no MUSS row - set PARENT_ID by hand"
        return parents, spurs, spur_from, notes
    if not good:
        return parents, spurs, spur_from, notes
    if len(good) == 1:
        parents[good[0]["code"]] = muss_id
        spurs[good[0]["code"]] = "T"
        return parents, spurs, spur_from, notes

    lat0 = sorted(p["lat"] for p in good)[len(good) // 2]
    lon0 = sorted(p["lon"] for p in good)[len(good) // 2]
    pts = [_metres(muss_node["lat"], muss_node["lon"], lat0, lon0)]
    pts += [_metres(p["lat"], p["lon"], lat0, lon0) for p in good]
    parent = prim([], pts, root=0)          # index 0 is the MUSS
    keys = [muss_id] + [p["code"] for p in good]

    children = defaultdict(list)
    for i, p in enumerate(parent):
        if p >= 0:
            children[p].append(i)

    def span_len(i):
        p = parent[i]
        return 0.0 if p < 0 else math.hypot(pts[i][0] - pts[p][0], pts[i][1] - pts[p][1])

    # trunk = the longest path measured from the MUSS
    best_d, best_path = -1.0, [0]
    stack = [(0, 0.0, [0])]
    while stack:
        u, d, path = stack.pop()
        if d > best_d:
            best_d, best_path = d, path
        for v in children[u]:
            stack.append((v, d + span_len(v), path + [v]))
    trunk = set(best_path)

    for i in range(1, len(pts)):
        p = parent[i]
        parents[keys[i]] = keys[p] if p >= 0 else ""
        spurs[keys[i]] = "T" if i in trunk else ""

    counter = 0
    for u in best_path:
        for v in children[u]:
            if v in trunk:
                continue
            counter += 1
            sid = f"S{counter}"
            spur_from[sid] = keys[u]
            todo = [v]
            while todo:
                w = todo.pop()
                spurs[keys[w]] = sid
                todo.extend(children[w])
    return parents, spurs, spur_from, notes


def build_topology_template(master_path: str, out_path: str) -> str:
    """Write template/TOPOLOGY TEMPLATE.xlsx prefilled from the DTC master."""
    dtcs, hts, nodes = read_master(master_path)
    nodes = assign_nodes_to_feeders(nodes, dtcs)

    feeder_names = [f for f, _ in Counter(d["feeder"] for d in dtcs).most_common()]
    station_of = {f: next((d["station"] for d in dtcs if d["feeder"] == f), "")
                  for f in feeder_names}
    muss_of = {f: find_station_node(nodes, station_of[f]) for f in feeder_names}
    topo = {f: _infer_topology([d for d in dtcs if d["feeder"] == f],
                               muss_of[f],
                               _node_id("MUSS", muss_of[f]["name"]) if muss_of[f] else "")
            for f in feeder_names}

    dupes = defaultdict(list)
    for d in dtcs:
        if d["lat"] is not None and d["lon"] is not None:
            dupes[(round(d["lat"], 6), round(d["lon"], 6))].append(d["code"])

    rows = []
    for n in [x for x in nodes if x["kind"] == "MUSS"]:
        feeds = [f for f in feeder_names if find_station_node([n], station_of[f])]
        rows.append([_node_id("MUSS", n["name"]), "MUSS", n["name"], "", n["lat"], n["lon"],
                     "", "T", "", "", "", "", "", "SOURCE",
                     "Root - no parent. Feeds: " + (", ".join(feeds) or "?")])
    for n in [x for x in nodes if x["kind"] == "RMU"]:
        rows.append([_node_id("RMU", n["name"]), "RMU", n["name"], n.get("feeder", ""),
                     n["lat"], n["lon"], "", "", "", "", "", "NC", "", "UNKNOWN",
                     "5-way RMU: ODin, ODout, VL1, VL2, VL3 - not all need be used. Set PARENT_ID = "
                     "the node it is fed from (that link is way ODin); each child names its way in "
                     "WAY. See the RMU WAYS sheet"])
    for f in feeder_names:
        parents, spurs, _, notes = topo[f]
        sel = sorted([d for d in dtcs if d["feeder"] == f],
                     key=lambda x: int(re.sub(r"\D", "", x["code"]) or 0))
        for d in sel:
            parent = parents.get(d["code"], "")
            remark = notes.get(d["code"], "")
            key = (round(d["lat"], 6), round(d["lon"], 6)) if d["lat"] is not None else None
            if key and len(dupes.get(key, [])) > 1:
                others = [c for c in dupes[key] if c != d["code"]]
                remark = (remark + "; " if remark else "") + "shared coordinate with " + ", ".join(others)
            rows.append([d["code"], "DTC", d["name"], f, d["lat"], d["lon"], parent,
                         spurs.get(d["code"], ""), "", d["kva"], "", "", "",
                         "INFERRED" if parent else "UNKNOWN", remark])
        for h in sorted([x for x in hts if x["feeder"] == f],
                        key=lambda x: int(re.sub(r"\D", "", x["code"]) or 0)):
            rows.append([h["code"], "HT", h["name"], f, None, None, "", "", "", h["kva"],
                         "", "", "", "UNKNOWN", "No GPS - set PARENT_ID to the tap pole"])

    # ---- workbook shell ----------------------------------------------------
    wb = openpyxl.Workbook()
    rm = wb.active
    rm.title = "READ ME"
    ws = wb.create_sheet("NODES")
    sp = wb.create_sheet("SPURS")
    rw = wb.create_sheet("RMU WAYS")
    bd = wb.create_sheet("BENDS")
    ex = wb.create_sheet("EXAMPLE")
    ls = wb.create_sheet("LISTS")

    head_fill = PatternFill("solid", fgColor="1F4E79")
    head_font = Font(bold=True, color="FFFFFF", size=10)
    title_font = Font(bold=True, size=15, color="1F4E79")
    sub_font = Font(size=10, color="595959")
    band_fill = PatternFill("solid", fgColor="DEEAF6")
    thin = Side(style="thin", color="BFBFBF")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)

    def header(sheet, row, cols):
        for i, (name, width) in enumerate(cols, start=1):
            c = sheet.cell(row=row, column=i, value=name)
            c.fill, c.font = head_fill, head_font
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = box
            sheet.column_dimensions[c.column_letter].width = width
        sheet.row_dimensions[row].height = 30

    def body(sheet, row, values, ncols, wrap_cols=(3, 15)):
        for i in range(1, ncols + 1):
            v = values[i - 1] if i - 1 < len(values) else None
            c = sheet.cell(row=row, column=i, value=v)
            c.font = Font(size=9)
            c.border = box
            c.alignment = Alignment(vertical="top",
                                    wrap_text=i in wrap_cols,
                                    horizontal="center" if i in (2, 8, 9, 10, 11, 12, 14) else "left")
            if i in (5, 6):
                c.number_format = "0.0000000"
        if row % 2 == 0:
            for i in range(1, ncols + 1):
                sheet.cell(row=row, column=i).fill = PatternFill("solid", fgColor="F2F7FC")

    def fit_rows(sheet, first_row, data, name_idx, note_idx, name_w, note_w):
        """Give wrapped rows enough height that Excel does not clip them."""
        for n, values in enumerate(data):
            nm = str(values[name_idx] or "") if name_idx < len(values) else ""
            nt = str(values[note_idx] or "") if note_idx < len(values) else ""
            lines = max(math.ceil(len(nm) / name_w), math.ceil(len(nt) / note_w), 1)
            sheet.row_dimensions[first_row + n].height = max(15, 12 * lines + 3)

    # ---- READ ME -----------------------------------------------------------
    rm.column_dimensions["A"].width = 3
    rm.column_dimensions["B"].width = 118
    r = 1

    def para(text, bold=False, size=10, color="1F1F1F", height=None):
        nonlocal r
        c = rm.cell(row=r, column=2, value=text)
        c.font = Font(bold=bold, size=size, color=color)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        rm.row_dimensions[r].height = height or max(15, 14 * math.ceil(len(text) / 105))
        r += 1

    para("MESCOM SLD TOOL  -  TOPOLOGY DATA TEMPLATE", bold=True, size=16, color="1F4E79", height=24)
    para("Fill in PARENT_ID, SPUR_ID and WAY. Everything else is already populated from "
         "'DTC DETAILS U4.xlsx'.", size=10, color="595959")
    r += 1
    para("1.  WHAT THIS FILE IS FOR", bold=True, size=11, color="1F4E79")
    para("The GPS rows in the master tell the tool where every DTC, RMU and MUSS sits. They do not "
         "say what is connected to what. With points only, the geographical sheet joins them with a "
         "minimum spanning tree - a plausible guess that can put a DTC on the wrong lateral, draw a "
         "line across a block, or make a spur run through a building.")
    para("One connection per node fixes that. Each row in NODES says 'this node is fed from that "
         "node'. Read together, the rows are the whole feeder tree: the trunk, every spur, and the "
         "order of DTCs along each one.")
    r += 1
    para("2.  THE SIX RULES", bold=True, size=11, color="1F4E79")
    para("1.  One row per physical node. NODE_ID must be unique.")
    para("2.  Exactly one PARENT_ID per row - the node immediately upstream, towards the MUSS. "
         "Only a MUSS row has a blank parent.")
    para("3.  The parent must already exist in the NODE_ID column. No cycles: following parent "
         "links must always end at a MUSS.")
    para("4.  Power flows FROM the parent TO the node. Do not write the child in the parent's row.")
    para("5.  SPUR_ID is T for the main run from the MUSS, S1, S2, ... for laterals. Every node on "
         "the same lateral carries the same SPUR_ID.")
    para("6.  WAY is only used when the parent is an RMU. An RMU is a five-way unit - ODin, ODout, "
         "VL1, VL2, VL3 - so a child of an RMU must say which way it leaves by. Leave WAY blank "
         "everywhere else. A unit need not use all five ways: only ODin is always cabled, and any "
         "of the four outgoing ways, ODout included, may be empty.")
    r += 1
    para("3.  SHEETS", bold=True, size=11, color="1F4E79")
    para("NODES    The only mandatory sheet. One row per MUSS, RMU, GOS, pole, DTC and HT service. "
         "PARENT_ID is the connection; everything else describes the node.")
    para("SPURS    Auto-generated register of the laterals found in NODES. Rename them here "
         "(e.g. 'Kote Road lateral'); the names are for the drawing, not for the tree.")
    para("RMU WAYS One row per way of every RMU, five per unit: ODin, ODout, VL1, VL2, VL3. "
         "CONNECTED says whether a cable is actually landed; the four outgoing ways start empty "
         "(N), which is normal. Records the normal state and the label of each way. Prefilled from "
         "the RMU rows in NODES; it describes the switchgear, it does not carry the connections.")
    para("BENDS    Optional. Add rows only where a span has to bend round a block instead of running "
         "straight from parent to child. SEQ orders the bend points along that one span.")
    para("EXAMPLE  A small, complete feeder showing exactly how the rows should read. Copy its "
         "pattern; do not edit it.")
    para("LISTS    Dropdown values. Leave hidden.")
    r += 1
    para("4.  WHAT IS ALREADY FILLED IN", bold=True, size=11, color="1F4E79")
    para(f"•  {len([x for x in rows if x[1] == 'DTC'])} DTCs, "
         f"{len([x for x in rows if x[1] == 'HT'])} HT services, "
         f"{len([x for x in rows if x[1] == 'MUSS'])} MUSS and "
         f"{len([x for x in rows if x[1] == 'RMU'])} RMU rows, with the coordinates the master holds.")
    para("•  PARENT_ID is pre-filled from a minimum spanning tree and marked BASIS = INFERRED. "
         "Treat it as a first draft to correct, never as surveyed fact.")
    para("•  Rows with BASIS = UNKNOWN have no usable GPS or are HT/RMU nodes whose tap is not in "
         "the workbook. They need a manual parent.")
    para("•  SPUR_ID is also inferred: the longest path from the MUSS is called T, every branch "
         "off it is numbered S1, S2, ...")
    para(f"•  RMU WAYS already lists all {len([x for x in nodes if x['kind'] == 'RMU']) * 5} ways "
         f"of the {len([x for x in nodes if x['kind'] == 'RMU'])} RMUs: ODin cabled, the four "
         "outgoing ways marked empty. Fill in a way only where a cable really lands.")
    r += 1
    para("5.  COLOURS", bold=True, size=11, color="1F4E79")
    para("Amber BASIS      the parent is inferred - check it")
    para("Orange PARENT_ID  no parent yet - the node will be left out of the tree")
    para("Green BASIS      surveyed or read off the existing SLD - trusted")
    para("Orange WAY       a way is filled in but the node has no parent - it cannot be placed")
    para("Grey RMU WAYS    an empty way - no cable, nothing to draw")
    para("Amber NORMALLY   CONNECTED = Y but the switch state is missing")
    r += 1
    para("6.  WHAT TO DO FIRST", bold=True, size=11, color="1F4E79")
    para("Start with MRS F1 (the largest feeder). Open the hand-drawn SLD beside this file and, for "
         "each DTC, correct PARENT_ID to the pole or DTC it actually hangs from. Add the poles you "
         "need as new rows (TYPE = POLE, NODE_ID = P-001 ...). Leave pole coordinates blank if you "
         "do not have them - the drawing will place the pole on a straight line between its "
         "neighbours, which is good enough for the topology.")
    para("When NODES is complete, the generator can be extended to draw the real trunk and spurs "
         "instead of inferring them. This workbook is the collection format; the reader is the next "
         "code step. See docs/TOPOLOGY_DATA_REQUIREMENTS.md for the reasoning and the optional extras.")
    r += 1
    para("7.  REBUILDING THIS TEMPLATE", bold=True, size=11, color="1F4E79")
    para('python sld_generator.py --topology-template          # refuses to overwrite a filled file\n'
         'python sld_generator.py --topology-template --force  # discards your edits and rebuilds',
         size=9, color="404040")
    r += 1
    para("Coordinates are decimal degrees. Valid values for this area: lat about 13.9, "
         "long about 75.5. A latitude of 3.9 means a dropped leading digit.", size=9, color="C00000")

    # ---- NODES -------------------------------------------------------------
    header(ws, 1, NODE_HEADERS)
    for i, values in enumerate(rows, start=2):
        body(ws, i, values, len(NODE_HEADERS))
    fit_rows(ws, 2, rows, 2, 14, 34, 50)
    last = len(rows) + 1
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = f"A1:O{last}"
    ws.sheet_view.zoomScale = 90
    ws.print_title_rows = "1:1"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # ---- SPURS -------------------------------------------------------------
    header(sp, 1, [("FEEDER", 14), ("SPUR_ID", 9), ("LABEL", 26), ("FROM_NODE", 16),
                   ("IS_TRUNK", 9), ("DTC_COUNT", 10), ("KVA_TOTAL", 10), ("NOTES", 46)])
    spur_rows = []
    for f in feeder_names:
        _, spurs, spur_from, _ = topo[f]
        sel = [d for d in dtcs if d["feeder"] == f]
        muss_id = _node_id("MUSS", muss_of[f]["name"]) if muss_of[f] else ""
        sids = [s for s in spurs.values() if s]
        order = ["T"] + sorted({s for s in sids if s != "T"},
                               key=lambda s: int(re.sub(r"\D", "", s) or 0))
        for sid in order:
            members = [d for d in sel if spurs.get(d["code"]) == sid]
            if sid == "T":
                spur_rows.append([f, "T", "Main line from MUSS", muss_id, "Y", len(members),
                                  sum(d["kva"] or 0 for d in members),
                                  "Longest path from the MUSS - rename if the real trunk is another run"])
            elif members:
                spur_rows.append([f, sid, "", spur_from.get(sid, ""), "N", len(members),
                                  sum(d["kva"] or 0 for d in members), ""])
    for i, values in enumerate(spur_rows, start=2):
        body(sp, i, values, 8, wrap_cols=(3, 8))
    fit_rows(sp, 2, spur_rows, 2, 7, 24, 44)
    sp.freeze_panes = "A2"
    sp.auto_filter.ref = f"A1:H{len(spur_rows) + 1}"

    # ---- RMU WAYS ----------------------------------------------------------
    header(rw, 1, RMU_WAY_HEADERS)
    way_rows = []
    for n in [x for x in nodes if x["kind"] == "RMU"]:
        rid = _node_id("RMU", n["name"])
        for way, direction in RMU_WAYS:
            cabled = _way_cabled(way)
            note = ("Incoming way, always cabled - the RMU's own PARENT_ID is this link"
                    if way == "ODin" else
                    "Empty is normal: set CONNECTED = Y only when a cable is landed and a "
                    "child row names this way in NODES!WAY")
            way_rows.append([rid, way, direction, "Y" if cabled else "N",
                             "NC" if cabled else "", "", note])
    for i, values in enumerate(way_rows, start=2):
        body(rw, i, values, len(RMU_WAY_HEADERS), wrap_cols=(6, 7))
    fit_rows(rw, 2, way_rows, 5, 6, 28, 54)
    rw.freeze_panes = "A2"
    rw.auto_filter.ref = f"A1:G{len(way_rows) + 1}"

    # ---- BENDS -------------------------------------------------------------
    header(bd, 1, [("FROM_ID", 15), ("TO_ID", 15), ("SEQ", 7), ("LAT", 12), ("LONG", 12),
                   ("NOTES", 46)])
    bd.freeze_panes = "A2"
    bd["A2"] = "Add one row per bend on a span, SEQ = 1, 2, 3 ... along the line."
    bd["A2"].font = Font(italic=True, size=9, color="808080")
    bd.merge_cells("A2:F2")

    # ---- EXAMPLE -----------------------------------------------------------
    ex.column_dimensions["A"].width = 3
    ex.column_dimensions["B"].width = 118
    ex["B1"] = "EXAMPLE  -  a small feeder, filled in the required format"
    ex["B1"].font = Font(bold=True, size=13, color="1F4E79")
    ex["B2"] = ("The trunk runs MUSS -> P-001 -> GOS-01 -> P-002 -> RMU (way ODin) -> P-003 "
                "(way VL2). SCD-08 leaves the RMU by way VL1; ODout and VL3 are empty. "
                "SCD-01 taps P-001. Spur S1 leaves the trunk at P-002 and feeds SCD-07. Every row's "
                "PARENT_ID points upstream; SPUR_ID groups the branch and WAY names the RMU way a "
                "child leaves by.")
    ex["B2"].alignment = Alignment(wrap_text=True, vertical="top")
    ex["B2"].font = Font(size=10)
    ex.merge_cells("B2:O2")
    ex.row_dimensions[2].height = 32
    ex_head = 4
    header(ex, ex_head, NODE_HEADERS)
    example_rows = [
        ["MUSS-DEMO", "MUSS", "110/11KV DEMO MUSS", "DEMO", 13.930000, 75.585000, "", "T", "",
         None, "", "", "", "SURVEYED", "Root - no parent"],
        ["P-001", "POLE", "Pole 1", "DEMO", 13.929800, 75.584800, "MUSS-DEMO", "T", "",
         None, "", "", "3x95 ABC", "SURVEYED", ""],
        ["GOS-01", "GOS", "DEMO GOS", "DEMO", 13.929600, 75.584600, "P-001", "T", "",
         None, "", "NC", "3x95 ABC", "SURVEYED", "Normally closed"],
        ["SCD-01", "DTC", "GANESHA TC", "DEMO", 13.929550, 75.584550, "P-001", "T", "",
         250, "3PH", "", "3x25 ABC", "SURVEYED", "Tap off P-001, 30 m service"],
        ["P-002", "POLE", "Pole 2", "DEMO", 13.929300, 75.584300, "GOS-01", "T", "",
         None, "", "", "3x95 ABC", "SURVEYED", "Branch point for S1"],
        ["P-010", "POLE", "Branch pole", "DEMO", 13.929250, 75.584250, "P-002", "S1", "",
         None, "", "", "3x95 ABC", "SURVEYED", "First pole of spur S1"],
        ["SCD-07", "DTC", "TEMPLE STREET", "DEMO", 13.929200, 75.584200, "P-010", "S1", "",
         63, "3PH", "", "3x25 ABC", "SURVEYED", ""],
        ["RMU-DEMO", "RMU", "DEMO RMU", "DEMO", 13.929350, 75.584350, "P-002", "T", "",
         None, "", "NC", "3x95 ABC", "SURVEYED",
         "5-way unit. Fed from P-002 - that link is way ODin"],
        ["P-003", "POLE", "Pole 3", "DEMO", 13.929000, 75.584000, "RMU-DEMO", "T", "VL2",
         None, "", "", "3x95 ABC", "SURVEYED",
         "Continues beyond the RMU by way VL2 - the run out is not always ODout"],
        ["SCD-08", "DTC", "SCHOOL", "DEMO", 13.929380, 75.584320, "RMU-DEMO", "T", "VL1",
         100, "3PH", "", "3x25 ABC", "SURVEYED", "Tapped off RMU way VL1"],
    ]
    for i, values in enumerate(example_rows, start=ex_head + 1):
        body(ex, i, values, len(NODE_HEADERS))
    fit_rows(ex, ex_head + 1, example_rows, 2, 14, 34, 50)
    bend_head = ex_head + len(example_rows) + 3
    ex.cell(row=bend_head - 1, column=2, value="BENDS  -  optional shape points for one span").font = \
        Font(bold=True, size=11, color="1F4E79")
    header(ex, bend_head, [("FROM_ID", 15), ("TO_ID", 15), ("SEQ", 7), ("LAT", 12),
                           ("LONG", 12), ("NOTES", 46)])
    bend_rows = [
        ["P-001", "P-002", 1, 13.929700, 75.584900, "Line turns at the school corner"],
        ["P-001", "P-002", 2, 13.929500, 75.584700, "Then runs along the compound wall"],
    ]
    for i, values in enumerate(bend_rows, start=bend_head + 1):
        body(ex, i, values, 6, wrap_cols=(6,))
    ways_head = bend_head + len(bend_rows) + 3
    ex.cell(row=ways_head - 1, column=2,
            value="RMU WAYS  -  the five ways of DEMO RMU, as listed on the RMU WAYS sheet").font = \
        Font(bold=True, size=11, color="1F4E79")
    header(ex, ways_head, RMU_WAY_HEADERS)
    example_ways = [
        ["RMU-DEMO", "ODin", "IN", "Y", "NC", "In from P-002",
         "Fed from P-002; no child row names this way"],
        ["RMU-DEMO", "ODout", "OUT", "N", "", "",
         "Empty - no cable landed on this way"],
        ["RMU-DEMO", "VL1", "OUT", "Y", "NC", "", "SCD-08 carries WAY = VL1"],
        ["RMU-DEMO", "VL2", "OUT", "Y", "NC", "Out to P-003",
         "P-003 carries WAY = VL2; the continuing run is not always ODout"],
        ["RMU-DEMO", "VL3", "OUT", "N", "", "Spare",
         "Empty spare way, ready for a new cable"],
    ]
    for i, values in enumerate(example_ways, start=ways_head + 1):
        body(ex, i, values, len(RMU_WAY_HEADERS), wrap_cols=(6, 7))
    fit_rows(ex, ways_head + 1, example_ways, 5, 6, 28, 52)
    # One width set for the whole sheet: the three tables share these columns.
    # F is wide because both BENDS and RMU WAYS keep their free text there.
    for col, width in zip("ABCDEFGHIJKLMNO",
                          (15, 15, 34, 12, 12, 46, 15, 8, 9, 7, 7, 9, 14, 10, 52)):
        ex.column_dimensions[col].width = width

    # ---- LISTS -------------------------------------------------------------
    lists = {
        "A": ("TYPE", ["MUSS", "RMU", "GOS", "POLE", "DTC", "HT", "FUSE", "LA", "JUNCTION"]),
        "B": ("PHASE", ["3PH", "1PH"]),
        "C": ("NORMALLY", ["NC", "NO"]),
        "D": ("BASIS", ["SOURCE", "INFERRED", "SURVEYED", "FROM-SLD", "UNKNOWN"]),
        "E": ("WAY_ALL", [w for w, _ in RMU_WAYS]),
        "F": ("WAY_CHILD", [w for w, d in RMU_WAYS if d == "OUT"]),
        "G": ("DIRECTION", ["IN", "OUT"]),
        "H": ("YES_NO", ["Y", "N"]),
    }
    for col, (name, values) in lists.items():
        ls[f"{col}1"] = name
        ls[f"{col}1"].font = Font(bold=True, size=9)
        for i, v in enumerate(values, start=2):
            ls[f"{col}{i}"] = v
        ls.column_dimensions[col].width = 14
    ls.sheet_state = "hidden"

    # ---- validation and colour rules --------------------------------------
    def dv(formula, cell_range, sheet=ws):
        d = DataValidation(type="list", formula1=formula, allow_blank=True)
        d.error = "Pick a value from the list."
        d.errorTitle = "Invalid entry"
        sheet.add_data_validation(d)
        d.add(cell_range)

    dv("=LISTS!$A$2:$A$10", f"B2:B{last}")
    dv("=LISTS!$B$2:$B$3", f"K2:K{last}")
    dv("=LISTS!$C$2:$C$3", f"L2:L{last}")
    dv("=LISTS!$D$2:$D$6", f"N2:N{last}")
    dv("=LISTS!$F$2:$F$5", f"I2:I{last}")
    way_last = len(way_rows) + 1
    dv("=LISTS!$E$2:$E$6", f"B2:B{way_last}", rw)
    dv("=LISTS!$G$2:$G$3", f"C2:C{way_last}", rw)
    dv("=LISTS!$H$2:$H$3", f"D2:D{way_last}", rw)
    dv("=LISTS!$C$2:$C$3", f"E2:E{way_last}", rw)

    amber = PatternFill("solid", fgColor="FFF2CC")
    orange = PatternFill("solid", fgColor="FCE4D6")
    green = PatternFill("solid", fgColor="E2EFDA")
    grey = PatternFill("solid", fgColor="F2F2F2")
    ws.conditional_formatting.add(f"N2:N{last}", FormulaRule(formula=['$N2="INFERRED"'], fill=amber))
    ws.conditional_formatting.add(f"N2:N{last}", FormulaRule(formula=['$N2="UNKNOWN"'], fill=orange))
    ws.conditional_formatting.add(
        f"N2:N{last}", FormulaRule(formula=['OR($N2="SURVEYED",$N2="FROM-SLD")'], fill=green))
    ws.conditional_formatting.add(
        f"G2:G{last}", FormulaRule(formula=['AND($B2<>"MUSS",$G2="")'], fill=orange))
    # A way is only meaningful on a child of an RMU: flag one that has no parent.
    ws.conditional_formatting.add(
        f"I2:I{last}", FormulaRule(formula=['AND($I2<>"",$G2="")'], fill=orange))
    # Empty ways are normal; grey them so the cabled ones stand out.
    rw.conditional_formatting.add(
        f"A2:G{way_last}", FormulaRule(formula=['$D2="N"'], fill=grey))
    rw.conditional_formatting.add(
        f"E2:E{way_last}", FormulaRule(formula=['AND($D2="Y",$E2="")'], fill=amber))

    wb.save(out_path)
    return out_path


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
    print(f"  preview     : {_display(out_png)}")


# ----------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="Generate a MESCOM feeder SLD as an editable .drawio file.")
    ap.add_argument("--master", default=MASTER_DEFAULT, help="DTC master workbook")
    ap.add_argument("--feeder", help='feeder name as in the master, e.g. "Mandli F13"')
    ap.add_argument("--name", help="label printed on the sheet (default: feeder)")
    ap.add_argument("--title", help="override the title line")
    ap.add_argument("--out", help="output .drawio path (overrides --outdir)")
    ap.add_argument("--outdir", help="folder for generated .drawio files (default: outputs/<today>)")
    ap.add_argument("--preview", action="store_true", help="also write a PNG preview")
    ap.add_argument("--geo", action="store_true", help="place DTCs at their GPS positions")
    ap.add_argument("--muss", help='substation coordinates "lat,lon" for the geo layout')
    ap.add_argument("--extent", choices=["cluster", "all"], default="cluster",
                    help="geo layout extent: fit the DTC network (default) or include the MUSS")
    ap.add_argument("--route", help="surveyed pole coordinates (csv/xlsx: NODE_ID,LAT,LONG)")
    ap.add_argument("--template", action="store_true",
                    help="write the blank A3 sheet + symbol library to template/")
    ap.add_argument("--topology-template", action="store_true",
                    help="write template/TOPOLOGY TEMPLATE.xlsx from the master")
    ap.add_argument("--force", action="store_true",
                    help="overwrite an existing topology template (discards edits)")
    ap.add_argument("--list", action="store_true", help="list feeders found in the master")
    args = ap.parse_args()

    if args.topology_template:
        if not os.path.exists(args.master):
            raise SystemExit(f"Master workbook not found: {args.master}")
        out = args.out or TOPOLOGY_TEMPLATE
        if os.path.exists(out) and not args.force:
            raise SystemExit(
                f"Topology template already exists: {_display(out)}\n"
                "It may contain your edits. Rebuild with --force, or pass --out with a new file name.")
        os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
        build_topology_template(args.master, out)
        print(f"  topology    : {_display(out)}")
        if not args.template:
            return

    if args.template:
        out = args.out if not args.topology_template else TEMPLATE_DRAWIO
        out = out or TEMPLATE_DRAWIO
        os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
        xml = build_template().to_xml()
        with open(out, "w", encoding="utf-8") as fh:
            fh.write(xml)
        print(f"  template    : {_display(out)}")
        if args.preview:
            preview(xml, os.path.splitext(out)[0] + "_library.png", page_idx=1)
        return

    if not os.path.exists(args.master):
        raise SystemExit(
            f"Master workbook not found: {args.master}\n"
            f"Put 'DTC DETAILS U4.xlsx' in {DATA_DIR},\n"
            f"or pass the path with --master \"D:\\path\\to\\DTC DETAILS U4.xlsx\"")
    dtcs, hts, nodes = read_master(args.master)
    nodes = assign_nodes_to_feeders(nodes, dtcs)
    rmus = [n for n in nodes
            if n["kind"] == "RMU" and _norm(n.get("feeder", "")) == _norm(args.feeder or "")]

    if args.list or not args.feeder:
        from collections import Counter
        cnt = Counter(it["feeder"] for it in dtcs)
        print("Feeders in the master:")
        for f, c in cnt.items():
            with_gps = sum(1 for it in dtcs if it["feeder"] == f and it["lat"] is not None)
            print(f"  {f:<14} {c:>3} DTCs  ({with_gps} with GPS)")
        if nodes:
            print("\nNon-DTC nodes (sheet '%s'):" % SHEET_NODES)
            for n in nodes:
                print(f"  {n['kind']:<5} {n['name']:<28} {n['lat']:.6f}, {n['lon']:.6f}")
        if not args.feeder:
            print('\nNow run e.g.:  python sld_generator.py --feeder "Mandli F13" --geo')
        return

    sel = pick(dtcs, args.feeder)
    sel_ht = pick(hts, args.feeder)
    if not sel:
        raise SystemExit(f"No DTCs for feeder {args.feeder!r}. Use --list to see the names.")
    sel.sort(key=lambda x: int(re.sub(r"\D", "", x["code"]) or 0))

    station = next((x["station"] for x in sel if x["station"]), "110/11KV MUSS")
    label = args.name or args.feeder

    muss = None
    if args.muss:
        parts = re.split(r"[,\s]+", args.muss.strip())
        if len(parts) != 2:
            raise SystemExit('--muss needs "lat,lon", e.g. --muss "13.9305,75.5842"')
        muss = (float(parts[0]), float(parts[1]))
    elif args.geo:
        muss = find_station(nodes, station) or STATIONS.get(station)
        if muss is None:
            print("  note        : no MUSS coordinates found - trunk starts at one end of the"
                  "\n                network; add the MUSS to sheet 'NON DTC NODES' or pass --muss \"lat,lon\"")
        else:
            print(f"  muss        : {muss[0]:.6f}, {muss[1]:.6f}  (from '{SHEET_NODES}')")

    route = read_route(args.route) if args.route else None
    if args.route and not route:
        raise SystemExit(f"No usable LAT/LONG rows found in {args.route}")

    if args.geo:
        wb, info = build_geo(sel, sel_ht, label, station, muss=muss, route=route,
                             rmus=rmus, extent=args.extent, title_text=args.title)
    else:
        wb, info = build_schematic(sel, sel_ht, label, station, rmus=rmus,
                                   title_text=args.title)

    xml = wb.to_xml()
    if args.out:
        out = args.out
    else:
        outdir = args.outdir or os.path.join(OUTPUT_ROOT, date.today().strftime("%Y-%m-%d"))
        out = os.path.join(outdir, f"{label}{' - GEO' if args.geo else ''} SLD.drawio")
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    with open(out, "w", encoding="utf-8") as fh:
        fh.write(xml)

    print(f"  feeder      : {label}  ({len(sel)} DTCs, {len(sel_ht)} HT services)")
    if args.geo:
        print(f"  layout      : geo  |  {info['positioned']} positioned, "
              f"{info['unpositioned']} unpositioned  |  1 m = {info['scale']:.2f} px")
    else:
        print(f"  layout      : schematic  |  {info['per_row']} per lateral x {info['rows']} laterals")
    print(f"  written     : {_display(out)}")
    if args.preview:
        preview(xml, os.path.splitext(out)[0] + ".png")


if __name__ == "__main__":
    main()
