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

Two layouts:
  default  - distributor with equally spaced laterals (clean schematic)
  --geo    - every DTC at its surveyed GPS position, trunk inferred by a
             minimum spanning tree (or taken from --route pole coordinates)

Output: "<name> SLD.drawio" (or "<name> - GEO SLD.drawio"). Open it in
draw.io / diagrams.net, or in VS Code with the "Draw.io Integration" extension.
Drag a DTC and its line follows.

Templates live in "template/": the blank A3 sheet + symbol library and the
surveyed-route sample.
"""

from __future__ import annotations

import argparse
import csv
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
    """Read surveyed pole coordinates: NODE_ID, LAT, LONG (csv or xlsx)."""
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
        place_label(d, bx, by, f"<b>{nd['name']}</b>", taken, bw=rw, bh=rh)

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
    ap.add_argument("--geo", action="store_true", help="place DTCs at their GPS positions")
    ap.add_argument("--muss", help='substation coordinates "lat,lon" for the geo layout')
    ap.add_argument("--extent", choices=["cluster", "all"], default="cluster",
                    help="geo layout extent: fit the DTC network (default) or include the MUSS")
    ap.add_argument("--route", help="surveyed pole coordinates (csv/xlsx: NODE_ID,LAT,LONG)")
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
    out = args.out or f"{label}{' - GEO' if args.geo else ''} SLD.drawio"
    with open(out, "w", encoding="utf-8") as fh:
        fh.write(xml)

    print(f"  feeder      : {label}  ({len(sel)} DTCs, {len(sel_ht)} HT services)")
    if args.geo:
        print(f"  layout      : geo  |  {info['positioned']} positioned, "
              f"{info['unpositioned']} unpositioned  |  1 m = {info['scale']:.2f} px")
    else:
        print(f"  layout      : schematic  |  {info['per_row']} per lateral x {info['rows']} laterals")
    print(f"  written     : {out}")
    if args.preview:
        preview(xml, os.path.splitext(out)[0] + ".png")


if __name__ == "__main__":
    main()
