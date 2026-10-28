#!/usr/bin/env python3
"""
MESCOM SLD generator
====================
Reads the DTC master workbook (sheets "ALL", "HT", "NON DTC NODES") and the
connection workbook "field_data/TOPOLOGY.xlsx", and writes an editable
diagrams.net / draw.io single line diagram for one feeder.

    python sld_generator.py --list
    python sld_generator.py --feeder "Mandli F13" --preview
    python sld_generator.py --feeder "MRS F1" --geo --preview
    python sld_generator.py --feeder "MRS F1" --geo --no-topology
    python sld_generator.py --feeder "MRS F1" --geo --muss "13.9305,75.5842"
    python sld_generator.py --feeder "MRS F1" --geo --route poles.csv
    python sld_generator.py --template
    python sld_generator.py --topology-template

Two layouts:
  default  - distributor with equally spaced laterals (clean schematic)
  --geo    - every node at its surveyed GPS position. What connects to what
             comes from the connection workbook: PARENT_ID per node, RMU ways,
             BENDS to shape a span, SPURS to name a lateral. Without that
             workbook (or with --no-topology) the trunk is inferred by a
             minimum spanning tree, and --route overrides it with a surveyed
             pole list.

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
TOPOLOGY_DEFAULT = os.path.join(DATA_DIR, "TOPOLOGY.xlsx")
SHEET_ALL = "ALL"
SHEET_HT = "HT"
SHEET_NODES = "NON DTC NODES"
SHEET_TOPO_NODES = "NODES"
SHEET_TOPO_WAYS = "RMU WAYS"
SHEET_TOPO_BENDS = "BENDS"
SHEET_TOPO_SPURS = "SPURS"

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

# a five-way RMU is drawn as its box plus one terminal per way: ODin on the side
# it is fed from, the four outgoing ways in register order down the other side
RMU_W, RMU_H = 90, 24
RMU_TERM = 7            # the terminal square
RMU_LEAD = 15           # box edge to terminal
RMU_ROW = 18            # between two terminals on one side
RMU_LBL = 74            # room reserved for a way's label
RMU_CLUSTER = (RMU_W / 2 + RMU_LEAD + RMU_TERM / 2 + 5 + RMU_LBL,
               RMU_ROW * 1.5 + RMU_TERM / 2 + 6)

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
ST_POLE = "ellipse;whiteSpace=wrap;html=1;fillColor=#6c8ebf;strokeColor=#6c8ebf;"
ST_GOS = "rounded=0;whiteSpace=wrap;html=1;fillColor=#6c8ebf;strokeColor=none;"
# an open point: the blade is lifted off the line and left hollow, so the state
# survives a black-and-white print of the sheet
ST_GOS_OPEN = ("rounded=0;whiteSpace=wrap;html=1;fillColor=#ffffff;strokeColor=#6c8ebf;"
               "strokeWidth=1.6;")
ST_TINY = ("text;html=1;align=center;verticalAlign=middle;fontSize=6;whiteSpace=wrap;"
           "fontColor=#1a1a1a;")
ST_TINY_L = ("text;html=1;align=left;verticalAlign=middle;fontSize=6;whiteSpace=wrap;"
             "fontColor=#1a1a1a;")
ST_TINY_R = ST_TINY_L.replace("align=left", "align=right")
# one terminal of a five-way unit: filled when a cable lands on it, hollow when the
# register says the way is spare
ST_TERM = "rounded=0;whiteSpace=wrap;html=1;fillColor=#6c8ebf;strokeColor=#6c8ebf;"
ST_TERM_EMPTY = ("rounded=0;whiteSpace=wrap;html=1;fillColor=#ffffff;strokeColor=#6c8ebf;"
                 "strokeWidth=1.2;")
ST_LINE = "endArrow=none;html=1;rounded=0;strokeColor=#6c8ebf;strokeWidth=1;"
ST_TRUNK = "endArrow=none;html=1;rounded=0;strokeColor=#6c8ebf;strokeWidth=1.5;"
ST_GEO = "endArrow=none;html=1;rounded=0;strokeColor=#6c8ebf;strokeWidth=1.3;edgeStyle=none;"
# a guess, not a record: dashed so nobody reads it as surveyed
ST_GUESS = ("endArrow=none;html=1;rounded=0;dashed=1;strokeColor=#6c8ebf;"
            "strokeWidth=1.1;edgeStyle=none;")
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


def _station_matches(name: str, station: str) -> bool:
    """'110/11KV MANDLI' matches the node name '110/11KV MANDLI MUSS'."""
    a, b = _norm(name), _norm(station)
    return bool(a) and bool(b) and (a in b or b in a)


def pick(items, feeder: str):
    want = _norm(feeder)
    return [it for it in items if _norm(it["feeder"]) == want]


# ----------------------------------------------------------------------------
# topology workbook - the connection record the user fills in
# ----------------------------------------------------------------------------
# NODES is the tree: one row per physical node, one PARENT_ID each. RMU WAYS is a
# register of terminals, BENDS shapes a span around a corner, SPURS names the
# laterals. Read here, drawn by build_geo().
ENVELOPE_LAT = (13.5, 14.3)      # Shivamogga, generously drawn
ENVELOPE_LON = (75.1, 76.0)
VALID_CHILD_WAYS = ("ODout", "VL1", "VL2", "VL3")


def read_topology(path: str):
    """Read the connection workbook. Returns None when there is nothing to read."""
    if not path or not os.path.exists(path):
        return None
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    nodes = []
    if SHEET_TOPO_NODES in wb.sheetnames:
        for i, row in enumerate(wb[SHEET_TOPO_NODES].iter_rows(min_row=2, values_only=True),
                               start=2):
            row = list(row) + [None] * 15
            nid = _txt(row[0])
            if not nid:
                continue
            nodes.append(dict(id=nid, row=i, type=_txt(row[1]).upper(), name=_txt(row[2]),
                              feeder=_txt(row[3]), lat=_num(row[4]), lon=_num(row[5]),
                              parent=_txt(row[6]), spur=_txt(row[7]).upper(),
                              way=_txt(row[8]), kva=_kva(row[9]),
                              normally=_txt(row[11]).upper(), basis=_txt(row[13]).upper()))
    ways = defaultdict(dict)
    if SHEET_TOPO_WAYS in wb.sheetnames:
        for i, row in enumerate(wb[SHEET_TOPO_WAYS].iter_rows(min_row=2, values_only=True),
                                start=2):
            row = list(row) + [None] * 7
            rid, way = _txt(row[0]), _txt(row[1])
            if rid and way:
                ways[rid][way] = dict(connected=_txt(row[3]).upper() == "Y",
                                      normally=_txt(row[4]).upper(), label=_txt(row[5]),
                                      row=i)
    bends = defaultdict(list)
    if SHEET_TOPO_BENDS in wb.sheetnames:
        for row in wb[SHEET_TOPO_BENDS].iter_rows(min_row=2, values_only=True):
            row = list(row) + [None] * 6
            frm, to = _txt(row[0]), _txt(row[1])
            lat, lon = _num(row[3]), _num(row[4])
            if not (frm and to) or lat is None or lon is None:
                continue
            seq = _num(row[2]) or len(bends[(frm, to)]) + 1
            bends[(frm, to)].append((seq, lat, lon))
    for pts in bends.values():
        pts.sort()
    spurs = {}
    if SHEET_TOPO_SPURS in wb.sheetnames:
        for row in wb[SHEET_TOPO_SPURS].iter_rows(min_row=2, values_only=True):
            row = list(row) + [None] * 8
            if _txt(row[0]) and _txt(row[1]):
                spurs[(_norm(row[0]), _txt(row[1]).upper())] = _txt(row[2])
    dups = sorted(nid for nid, c in Counter(n["id"] for n in nodes).items() if c > 1)
    wb.close()
    return dict(path=path, nodes=nodes, ways=ways, bends=bends, spurs=spurs, dups=dups)


def topology_root(sel, station: str):
    """The source row this feeder hangs off: the MUSS the master names for it."""
    muss = [n for n in sel if n["type"] == "MUSS"]
    for n in muss:
        if _station_matches(n["name"], station):
            return n
    return muss[0] if muss else None


def topology_for_feeder(topo, feeder: str, station: str):
    """The rows one feeder draws: its own rows, the MUSS above them and any row
    whose FEEDER cell was left blank but hangs under one of them.

    Taking the subtree rather than "every row whose FEEDER matches" means a
    feeder cannot pick up another one's network, and a hand-added row under a
    feeder node is drawn even before its FEEDER cell is filled in.
    """
    if not topo:
        return []
    nodes = topo["nodes"]
    by_id = {n["id"]: n for n in nodes}
    want = _norm(feeder)
    chosen = {n["id"]: n for n in nodes if _norm(n["feeder"]) == want}
    root = topology_root([n for n in nodes if n["type"] == "MUSS"], station)
    if root:
        chosen[root["id"]] = root
    if not chosen:
        return []
    for nid in list(chosen):                      # ancestors - the source above it
        cur, guard = by_id.get(nid), 0
        while cur and guard < 500:
            p = by_id.get(cur["parent"])
            if not p or p["id"] in chosen:
                break
            chosen[p["id"]] = p
            cur, guard = p, guard + 1
    kids = defaultdict(list)
    for n in nodes:
        if n["parent"]:
            kids[n["parent"]].append(n)
    stack = [n["id"] for n in chosen.values() if _norm(n["feeder"]) == want]
    while stack:                                  # descendants - blank FEEDER cells
        for k in kids.get(stack.pop(), []):
            if k["id"] not in chosen:
                chosen[k["id"]] = k
                stack.append(k["id"])
    return [n for n in nodes if n["id"] in chosen]


def validate_topology(sel, topo, master_codes):
    """Structural checks on the rows this feeder will draw.

    An error means the drawing would state something untrue - a broken parent
    link, a cable on a way that cannot carry one - and stops the run. A warning
    is printed and the sheet is still drawn.
    """
    errors, warns = [], []
    by_id = {n["id"]: n for n in sel}
    for nid in topo.get("dups", []):
        if nid in by_id:
            rows = ", ".join(str(n["row"]) for n in sel if n["id"] == nid)
            errors.append(f"NODES!NODE_ID {nid} appears on rows {rows} - one row per node, "
                          "rename one of them")
    for n in sel:
        if n["parent"] == n["id"]:
            errors.append(f"NODES!{n['row']} {n['id']}: PARENT_ID points at itself")
        elif n["parent"] and n["parent"] not in by_id:
            errors.append(f"NODES!{n['row']} {n['id']}: PARENT_ID {n['parent']!r} is not a "
                          "node in this workbook")
    for n in sel:                                 # a loop would draw forever upstream
        seen, cur = set(), n
        while cur and cur["parent"]:
            if cur["id"] in seen:
                errors.append(f"NODES!{n['row']} {n['id']}: the PARENT_ID chain loops back "
                              f"to {cur['id']}")
                break
            seen.add(cur["id"])
            cur = by_id.get(cur["parent"])
    for n in sel:
        if not n["parent"] and n["type"] != "MUSS" and n["lat"] is not None:
            warns.append(f"NODES!{n['row']} {n['id']}: no PARENT_ID - its parent is guessed "
                         "from the nearest node")
        if (n["lat"] is None) != (n["lon"] is None):
            warns.append(f"NODES!{n['row']} {n['id']}: only one of LAT/LONG is filled - "
                         "treated as having no coordinates")
        elif n["lat"] is not None and not (ENVELOPE_LAT[0] <= n["lat"] <= ENVELOPE_LAT[1]
                                           and ENVELOPE_LON[0] <= n["lon"] <= ENVELOPE_LON[1]):
            warns.append(f"NODES!{n['row']} {n['id']}: {n['lat']}, {n['lon']} is outside the "
                         "Shivamogga envelope - check the row")
        if n["type"] in ("DTC", "HT") and n["id"].upper() not in master_codes:
            warns.append(f"NODES!{n['row']} {n['id']}: no {n['type']} row with that code in the "
                         "master - the register and the topology disagree")
        if n["way"]:
            parent = by_id.get(n["parent"])
            if not parent or parent["type"] != "RMU":
                errors.append(f"NODES!{n['row']} {n['id']}: WAY {n['way']!r} but its parent "
                              f"{n['parent'] or '(none)'!r} is not an RMU")
            elif n["way"] not in VALID_CHILD_WAYS:
                errors.append(f"NODES!{n['row']} {n['id']}: WAY {n['way']!r} is not a way out "
                              f"of an RMU - use {'/'.join(VALID_CHILD_WAYS)}")

    kids = defaultdict(list)
    for n in sel:
        kids[n["parent"]].append(n)

    def load_below(code):
        """DTCs and kVA under a node - what a switch or a way is holding."""
        out, stack = [], list(kids.get(code, []))
        while stack:
            u = stack.pop()
            out.append(u)
            stack.extend(kids.get(u["id"], []))
        dtcs = [x for x in out if x["type"] == "DTC"]
        return len(dtcs), sum(x["kva"] or 0 for x in dtcs)

    for n in sel:                                 # the state of a switch is not optional
        if n["type"] != "GOS":
            continue
        state = (n["normally"] or "").upper()
        count, kva = load_below(n["id"])
        tail = (f" - {count} DTC(s) / {kva} kVA recorded downstream" if count
                else " - nothing is recorded downstream of it")
        if not state:
            warns.append(f"NODES!{n['row']} {n['id']}: NORMALLY is blank - drawn closed; "
                         f"record NC or NO{tail}")
        elif state == "NO":
            warns.append(f"NODES!{n['row']} {n['id']}: NORMALLY = NO - drawn as an open point"
                         f"{tail}; while it stands open that load is fed from elsewhere")

    used = defaultdict(list)                      # one cable per terminal
    for n in sel:
        if n["way"] and by_id.get(n["parent"], {}).get("type") == "RMU":
            used[(n["parent"], n["way"])].append(n)
    for (rid, way), kids in sorted(used.items()):
        if len(kids) > 1:
            who = ", ".join(k["id"] for k in kids)
            errors.append(f"RMU {rid} way {way} is claimed by {who} - one cable per terminal")
    for n in sel:
        if n["type"] != "RMU":
            continue
        reg = topo["ways"].get(n["id"], {})
        if not reg:
            warns.append(f"{n['id']}: no rows on sheet '{SHEET_TOPO_WAYS}' - the five ways "
                         "are not recorded")
        for way, rec in sorted(reg.items()):
            if way == "ODin":
                continue
            kids = [k["id"] for k in used.get((n["id"], way), [])]
            if rec["connected"] and not kids:
                warns.append(f"{SHEET_TOPO_WAYS}!{rec['row']} {n['id']} {way}: CONNECTED = Y "
                             "but no child names that way - is a node missing?")
            if kids and not rec["connected"]:
                # NODES is the tree and RMU WAYS is a register: the cable is drawn,
                # because a child is parented here, and the register is the stale one
                warns.append(f"{SHEET_TOPO_WAYS}!{rec['row']} {n['id']} {way}: CONNECTED = N "
                             f"but {', '.join(kids)} leaves by that way - set it to Y")
            if not rec["connected"] and rec["normally"]:
                warns.append(f"{SHEET_TOPO_WAYS}!{rec['row']} {n['id']} {way}: empty way with "
                             "a NORMALLY value - an empty way has no switch state")
            if kids and rec["normally"] == "NO":
                warns.append(f"{n['id']} {way}: normally open, with {', '.join(kids)} recorded "
                             "downstream - the load is off")
    for (frm, to) in sorted(topo["bends"]):       # a bend must belong to a real span
        if frm in by_id and to in by_id and by_id[to]["parent"] != frm:
            warns.append(f"BENDS row for {frm} -> {to}, but {to}'s PARENT_ID is "
                         f"{by_id[to]['parent'] or '(blank)'} - the bend is ignored")
    return errors, warns


def topology_items(dtcs, hts, sel, extra_nodes=()):
    """Merge the topology rows with the register in the master.

    The master owns a DTC's name, rating and GPS - it is the register the
    estimates are made from - while the topology provides the non-DTC equipment
    (poles, the GOS, the RMU) and every parent link. Equipment that only the
    master knows about is still drawn, so nothing quietly disappears.
    """
    master = {it["code"].upper(): it for it in list(dtcs) + list(hts)}
    items, missing, moved = [], [], []
    listed = {n["id"].upper() for n in sel}
    for n in sel:
        it = dict(n)
        it["code"] = n["id"]
        m = master.get(n["id"].upper())
        if m:
            if m.get("lat") is not None and n["lat"] is not None:
                gap = _distance_m((m["lat"], m["lon"]), (n["lat"], n["lon"]))
                if gap > 100:
                    moved.append(f"{n['id']}: master and topology are {gap:.0f} m apart - "
                                 "the master is used")
            it.update(name=m["name"], kva=m["kva"], kind=m["kind"], type=m["kind"])
            if m.get("lat") is not None:
                it["lat"], it["lon"] = m["lat"], m["lon"]
        else:
            it["kind"] = n["type"] or "NODE"
        items.append(it)
    for code, m in master.items():                # a DTC added to the master only
        if code in listed:
            continue
        it = dict(m)
        it.update(code=m["code"], type=m["kind"], parent="", spur="", way="", row=None,
                  normally="", basis="MASTER-ONLY")
        items.append(it)
        missing.append(m["code"])
    for nd in extra_nodes:                        # an RMU recorded in the master only
        gap = [_distance_m((it["lat"], it["lon"]), (nd["lat"], nd["lon"]))
               for it in items if it.get("lat") is not None]
        if gap and min(gap) < 30:
            continue                              # the workbook has it, under its own id
        it = dict(nd)
        it.update(code=nd["name"], type="RMU", kind="RMU", parent="", spur="", way="",
                  row=None, normally="", basis="MASTER-ONLY")
        items.append(it)
        missing.append(nd["name"])
    return items, missing, moved


def resolve_links(items, root_id, drawn=None):
    """The link to draw for every node, and why when it is a guess.

    A parent with no coordinates cannot be drawn to, so the line goes to the
    nearest ancestor that has them; a node with no usable ancestor is hung off
    the nearest node that is not downstream of it. Both are guesses - the
    drawing dashes them and the console says so. `drawn` is the set that will
    actually appear, so a rejected coordinate cannot pull a line across the
    sheet to nowhere.
    """
    by_id = {it["code"]: it for it in items}
    placed = set(drawn) if drawn else {c for c, it in by_id.items()
                                       if it.get("lat") is not None}
    kids = defaultdict(list)
    for it in items:
        if it["parent"] in by_id and it["parent"] != it["code"]:
            kids[it["parent"]].append(it["code"])

    def downstream(code):
        out, stack = set(), [code]
        while stack:
            u = stack.pop()
            out.add(u)
            stack.extend(k for k in kids[u] if k not in out)
        return out

    links = {}
    for it in items:
        code = it["code"]
        if code == root_id:
            links[code] = (None, "")
            continue
        cur, skipped = it["parent"], []
        while cur and cur in by_id and cur not in placed:
            skipped.append(cur)
            cur = by_id[cur]["parent"]
        if cur and cur in placed and cur not in downstream(code):
            why = ""
            if it["parent"] and it["parent"] not in by_id:
                why = f"parent {it['parent']} is not in the workbook - attached to {cur}, guessed"
            elif skipped:
                why = (f"line drawn to {cur}, past the unpositioned "
                       f"{', '.join(skipped)}")
            links[code] = (cur, why)
            continue
        if code not in placed:
            links[code] = (None, "")
            continue
        here = (it["lat"], it["lon"])
        blocked = downstream(code)
        cand = [c for c in placed if c not in blocked and c != code]
        if not cand:
            links[code] = (None, "")
            continue
        near = min(cand, key=lambda c: _distance_m(here, (by_id[c]["lat"], by_id[c]["lon"])))
        links[code] = (near, f"no usable parent - attached to the nearest node {near}, guessed")
    return links


def draw_rmu(d: Page, reg: dict, used: set, cx: float, cy: float):
    """Draw a five-way unit as its box plus one terminal per way.

    `reg` is the unit's rows from `RMU WAYS`, `used` the ways its children name in
    `NODES!WAY`. ODin is the way the unit is fed through, so it is drawn on the
    incoming side; the four outgoing ways sit in register order down the other
    side. A terminal a cable lands on is filled and named with its way; one that
    is left spare is hollow and named `empty`, so a way with nothing on it cannot
    be read as a record missing from the workbook.

    Returns the cluster rectangle to keep clear, and {way: (cell id, centre)}.
    """
    box_id = d.vertex("RMU", cx - RMU_W / 2, cy - RMU_H / 2, RMU_W, RMU_H, ST_RMU)
    terms = {}

    def terminal(way, side, y):
        cabled = side < 0 or way in used or (reg.get(way, {}).get("connected") or "").upper() == "Y"
        lx = cx + side * RMU_W / 2
        tx = lx + side * RMU_LEAD
        d.vertex("", min(lx, tx), y - 1, RMU_LEAD, 2, ST_BUS)      # box to terminal
        cell = d.vertex("", tx - RMU_TERM / 2, y - RMU_TERM / 2, RMU_TERM, RMU_TERM,
                        ST_TERM if cabled else ST_TERM_EMPTY)
        txt = way if cabled else f"{way} empty"
        state = (reg.get(way, {}).get("normally") or "").upper()
        if cabled and state:
            txt += f" {state}"
        lx0 = tx + 5 if side > 0 else tx - 5 - RMU_LBL
        d.vertex(txt, lx0, y - 6, RMU_LBL, 12, ST_TINY_L if side > 0 else ST_TINY_R)
        terms[way] = (cell, (tx, y))

    terminal("ODin", -1, cy)
    for i, way in enumerate(VALID_CHILD_WAYS):
        terminal(way, 1, cy + (i - 1.5) * RMU_ROW)
    hw, hh = RMU_CLUSTER
    return (cx - hw, cy - hh, cx + hw, cy + hh), terms, box_id


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
    parts = []
    for it in bad:
        parent = it.get("parent")
        parts.append(f"{it['code']} {crop_name(it['name'], 16)}"
                     + (f" (from {parent})" if parent else ""))
    codes = ", ".join(parts)
    x0, y0, w = 60, 1012, 1080
    lines = (w - 12) / 4.6
    h = 26 + 11 * max(1, math.ceil(len(codes) / lines))
    d.vertex(f"<b>UNPOSITIONED  -  no usable GPS</b><br>{codes}",
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


def _clamp_label(x, y, w=LBL_W, h=LBL_H):
    x = min(max(x, 20), PAGE_W - 20 - w)
    y = min(max(y, 96), TABLE_TOP - 24 - h)
    return x, y


def place_label(d: Page, cx, cy, value, taken, bw=DTC_W, bh=DTC_H,
                w=LBL_W, h=LBL_H, style=ST_LBL):
    """Put a label on the freest spot around the symbol at (cx, cy)."""
    candidates = []
    for dist in (0, 22, 44):
        candidates += [
            (cx - w / 2, cy + bh / 2 + 3 + dist),                 # below
            (cx - w / 2, cy - bh / 2 - h - 3 - dist),             # above
            (cx + bw / 2 + 3 + dist, cy - h / 2),                 # right
            (cx - bw / 2 - w - 3 - dist, cy - h / 2),             # left
        ]
    best, best_cost = None, None
    for cx0, cy0 in candidates:
        x, y = _clamp_label(cx0, cy0)
        r = (x, y, x + w, y + h)
        cost = sum(_overlap_area(r, t) for t in taken)
        if cost == 0:
            taken.append(r)
            return d.vertex(value, x, y, w, h, style)
        if best_cost is None or cost < best_cost:
            best, best_cost = r, cost
    taken.append(best)
    return d.vertex(value, best[0], best[1], w, h, style)


def free_slot(px, py, bw, bh, taken, first=None, dists=(0, 26, 52, 84, 120, 165),
              bounds=None):
    """Nearest free spot for a box that marks its true position.

    Wanted where a dense cluster would put two symbols on top of each other: the
    box walks outwards and the caller draws a stub back to the real spot. With
    `first` the walk starts on the side a DTC should hang off, which keeps the
    sheet reading as taps on a line rather than a scatter of boxes.
    """
    angles = [0.0, -math.pi / 2, math.pi / 2, math.pi,
              -math.pi / 4, math.pi / 4, -3 * math.pi / 4, 3 * math.pi / 4]
    base = 0.0
    if first:
        n = math.hypot(first[0], first[1]) or 1.0
        base = math.atan2(first[1] / n, first[0] / n)
    ring = [(0.0, 0.0)]
    for dist in dists[1:]:
        for ang in angles:
            ring.append((math.cos(base + ang) * dist, math.sin(base + ang) * dist))
    fallback = None
    for dx, dy in ring:
        x, y = px + dx, py + dy
        if bounds:
            x = min(max(x, bounds[0] + bw / 2), bounds[2] - bw / 2)
            y = min(max(y, bounds[1] + bh / 2), bounds[3] - bh / 2)
        rect = (x - bw / 2, y - bh / 2, x + bw / 2, y + bh / 2)
        if fallback is None:
            fallback = rect
        if not any(_overlaps(rect, t) for t in taken):
            return x, y, rect
    return fallback[0] + bw / 2, fallback[1] + bh / 2, fallback


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


def tree_order(items, root_id):
    """The workbook's tree in drawing order: the root first, then each branch.

    Depth first in workbook row order, so a parent is always drawn before its
    children and a spur's nodes stay together in one run of the grid. That is
    what lets the same lateral be followed on both sheets of a feeder.
    Returns the ordered items and the codes with no parent row to hang under.
    """
    by_id = {it["code"]: it for it in items}
    kids = defaultdict(list)
    for it in items:
        kids[it["parent"]].append(it["code"])
    out, seen, lost = [], set(), []

    def walk(code):
        if code in seen or code not in by_id:
            return
        seen.add(code)
        out.append(by_id[code])
        for kid in kids[code]:
            walk(kid)

    walk(root_id)
    for it in items:                    # nothing to hang under: drawn after the trunk
        if it["code"] in seen:
            continue
        if it["code"] != root_id and it["parent"] not in by_id:
            lost.append(it)
        walk(it["code"])
    return out, lost


def rmu_way_text(used, known):
    """Which ways cables arrive on, for the line under a unit's name.

    The geographical sheet draws these as five terminals; a grid row has no room
    for a cluster, so the same fact is said in words - and both sheets name the
    same ways.
    """
    if not known:
        return "5-way"
    out = [w for w in VALID_CHILD_WAYS if w in used]
    return "5-way" + (f" - out {', '.join(out)}" if out else " - no way out recorded")


def build_schematic(dtcs, hts, label, station, rmus=(), title_text=None, subtitle=None,
                    topo=None, root_id=None, spur_labels=None):
    """One A3 sheet as a grid of laterals hanging off a distributor.

    Without a workbook this is the register's rows in code order, which is all a
    schematic can be. With one, the nodes come out in the order the record
    connects them - depth first from the MUSS - each is labelled with the node
    that feeds it, and the node a spur starts at is tagged with the spur's name.
    The row line is a bus and not a drawing of the parents: there is no
    geography here, so the parent is written under the box instead.
    """
    wb = Workbook()
    d = wb.page(label or "SLD")
    if topo:
        order, lost = tree_order(list(topo), root_id)
        by_why = defaultdict(list)         # one line per reason, not one per node
        for it in lost:
            why = ("not in the workbook" if it.get("basis") == "MASTER-ONLY" else
                   "no PARENT_ID recorded" if not it["parent"] else
                   f"PARENT_ID {it['parent']!r} is not a row in the workbook")
            by_why[why].append(it["code"])
        for why, codes in by_why.items():
            print(f"  ! {len(codes)} node(s) drawn after the trunk, no parent named - {why}: "
                  f"{', '.join(codes)}")
        items = [it for it in order if (it.get("type") or "").upper() != "MUSS"]
        codes = {it["code"] for it in items}
        used_ways = defaultdict(set)
        for it in items:
            if it.get("way"):
                used_ways[it["parent"]].add(it["way"])
        # the tag marks the branch point, not every box on the spur: the first
        # node of a spur names where it leaves from, and depth first order puts
        # that node at the head of the spur's run
        start = {}
        for it in order:
            sp = (it.get("spur") or "").upper()
            if sp and sp != "T" and sp not in start and it["parent"] in codes:
                start[sp] = it["parent"]
        spurs_at = defaultdict(list)
        for sp, at in start.items():
            spurs_at[at].append(sp)
    else:
        items = [dict(it, type=it.get("kind", "DTC"), parent="", spur="", way="")
                 for it in list(dtcs) + list(hts) + list(rmus)]
        used_ways, spurs_at = defaultdict(set), {}
    per_row, rows, row_h = choose_grid(len(items))
    slot = (LATERAL_R - LATERAL_L) / per_row

    title = title_text or f"SINGLE LINE DIAGRAM OF {label.upper()} in Unit-4, CSD-2, SMG"
    sheet_header(d, title, subtitle or (
        f"{station}  MUSS   |   connection workbook   |   {len(items)} nodes   |   "
        f"{date.today():%d-%m-%Y}" if topo else
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

    for idx, item in enumerate(items):
        r, c = divmod(idx, per_row)
        y = FIRST_ROW_Y + r * row_h
        x = LATERAL_L + c * slot + SLOT_MIN / 2
        j = d.vertex("", x - 2, y, 4, 4, ST_JUNC)
        d.edge(junctions[r] if c == 0 else prev_j, j, ST_LINE)
        prev_j = j
        kind = (item.get("type") or item.get("kind") or "DTC").upper()
        name = crop_name(item["name"])
        code = item["code"]
        parent = item.get("parent") or ""
        # the node that feeds this one, so a box can be checked against the
        # geographical sheet, which draws that same link as a line
        from_txt = (f"<br><font style='font-size:7px;color=#6c8ebf'>from {parent}</font>"
                    if topo and parent else "")
        if kind == "POLE":
            d.vertex("", x - 3, y - 3, 6, 6, ST_POLE)        # a pole is in the run
            d.vertex(f"<b>{code}</b>{from_txt}", x - 54, y + 12, 108, 24, ST_TINY)
            continue
        if kind == "GOS":
            is_open = (item.get("normally") or "").upper() == "NO"
            d.vertex("", x - 15, y - 2, 30, 4, ST_GOS_OPEN if is_open else ST_GOS)
            d.vertex(f"<b>{code}</b>{' (open)' if is_open else ''}<br>{crop_name(name, 18)}"
                     f"{from_txt}", x - 54, y + 12, 108, 34, ST_TINY)
            continue
        if kind == "HT":
            tail = f" ({item['kva']} kVA)" if item.get("kva") else ""
            label_txt, face, w_item, style = (f"{name}<br><b>{code}</b>{tail}{from_txt}",
                                             "HT", DTC_W, ST_HT)
        elif kind == "RMU":
            label_txt = (f"<b>{crop_name(item['name'], 24)}</b>"
                         f"<br><font style='font-size:7px'>"
                         f"{rmu_way_text(used_ways.get(code, set()), bool(topo))}</font>"
                         f"{from_txt}")
            face, w_item, style = "RMU", 90, ST_RMU
        else:
            label_txt = f"{name}<br><b>{code}</b>{from_txt}"
            face = str(item["kva"]) if item.get("kva") else "?"
            w_item, style = DTC_W, ST_DTC
        sq = d.vertex(face, x - w_item / 2, y + 30, w_item, DTC_H, style)
        d.edge(j, sq, ST_LINE)
        d.vertex(label_txt, x - LBL_W / 2, y + 54, LBL_W, LBL_H + (10 if from_txt else 0),
                 ST_LBL)
        for k, spur in enumerate(spurs_at.get(code, ())):    # the tag marks the branch, not
            tag = (spur_labels or {}).get(spur, "")          # the first box of the spur
            d.vertex(f"<b>{spur}</b>{(' ' + tag) if tag else ''}", x - 60, y + 92 + k * 14,
                     120, 14, ST_TINY)

    summary_table(d, dtcs)
    return wb, dict(per_row=per_row, rows=rows, row_h=row_h, scale=None,
                    source="connection workbook" if topo else "DTC DETAILS U4.xlsx")


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


def _distance_m(a, b):
    """Metres between two (lat, lon) pairs - short distances, flat earth."""
    return math.hypot((a[0] - b[0]) * 110540.0,
                      (a[1] - b[1]) * 111320.0 * math.cos(math.radians(a[0])))


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


def ortho_route(a, b, min_run=70.0):
    """Waypoints for a run that goes across and down, turning square.

    A long diagonal reads as a survey line; a hand-drawn SLD turns at a right
    angle. The run leaves on the axis with the larger gap, turns at the
    midpoint and arrives square, and the corner is a single point so the two
    segments meet at exactly 90°. A short hop stays a straight line, because a
    jog of a few pixels looks like a mistake, not like a route.
    """
    ax, ay = a
    bx, by = b
    dx, dy = bx - ax, by - ay
    if math.hypot(dx, dy) < min_run:
        return []
    if abs(dx) < 6 or abs(dy) < 6:              # already square
        return []
    if abs(dx) >= abs(dy):
        corners = [(ax + dx / 2, ay), (ax + dx / 2, by)]
    else:
        corners = [(ax, ay + dy / 2), (bx, ay + dy / 2)]
    return corners


def _project(px, py, ax, ay, bx, by):
    vx, vy = bx - ax, by - ay
    L = vx * vx + vy * vy
    t = 0.0 if L == 0 else max(0.0, min(1.0, ((px - ax) * vx + (py - ay) * vy) / L))
    x, y = ax + t * vx, ay + t * vy
    return math.hypot(px - x, py - y), x, y


def build_geo(dtcs, hts, label, station, muss=None, route=None, rmus=(),
              extent="cluster", title_text=None, subtitle=None, topo=None,
              root_id=None, bends=None, spur_labels=None, ways=None):
    """One A3 sheet with every node at its surveyed position.

    `topo` is the feeder's rows from the connection workbook, already merged
    with the master (`topology_items`): that gives the parent links, the poles,
    the GOS and the RMU. Without it the sheet falls back to a minimum spanning
    tree over the DTC coordinates, which is a likeness, not a record.
    """
    wb = Workbook()
    d = wb.page(label or "SLD")

    if topo:
        items = list(topo)
    else:
        rmus = [dict(nd, code=nd["name"], type="RMU", kind="RMU", parent="", spur="",
                     way="") for nd in rmus]
        items = [dict(it, type=it.get("kind", "DTC"), parent="", spur="", way="")
                 for it in list(dtcs) + list(hts) + rmus]

    good, bad, warn = clean_points(items)
    for w in warn:
        print("  ! " + w)
    if len(good) < 2:
        raise SystemExit("Not enough nodes with GPS coordinates for a geographical layout.")

    lat0 = sorted(p["lat"] for p in good)[len(good) // 2]
    lon0 = sorted(p["lon"] for p in good)[len(good) // 2]
    local = {p["code"]: _metres(p["lat"], p["lon"], lat0, lon0) for p in good}
    order = [p["code"] for p in good]

    by_code = {p["code"]: p for p in good}

    # ---- who feeds whom ----------------------------------------------------
    trunk_pts = []       # surveyed polyline in metres, in drawing order
    links = {}           # code -> (parent code, why the link is a guess)
    guessed = []
    if route:
        trunk_pts = [_metres(lat, lon, lat0, lon0) for _, lat, lon in route]
        print(f"  route       : {len(trunk_pts)} surveyed points")
        if topo:
            print(f"  note        : --route shapes the run; the workbook's parent "
                  f"links and {len([i for i in items if i['type'] not in ('DTC', 'HT')])}"
                  " non-DTC node(s) are drawn on it")
    elif topo:
        links = resolve_links(items, root_id, drawn=set(local))
        guessed = sorted(c for c, (_, why) in links.items() if why)
        if guessed:
            print(f"  links       : {len(guessed)} guessed, drawn dashed - {', '.join(guessed)}")
        for c in guessed:
            print(f"  ! {c}: {links[c][1]}")
    else:
        pts = [local[c] for c in order]
        if muss:
            mx, my = _metres(muss[0], muss[1], lat0, lon0)
            root = min(range(len(pts)),
                       key=lambda i: math.hypot(pts[i][0] - mx, pts[i][1] - my))
        else:
            root = tree_ends(prim([], pts))[0]
        root_id = order[root]
        parent = prim([], pts, root)
        for i, code in enumerate(order):
            links[code] = (order[parent[i]] if parent[i] >= 0 else None, "")

    # ---- page transform ----------------------------------------------------
    # The source stays out of the fit: on two of these feeders the MUSS is more
    # than a kilometre from its own load, and pulling it in would shrink the
    # network to nothing. It is placed on its own and labelled when it is out.
    if muss is None and topo and root_id and root_id in by_code:
        root = by_code[root_id]
        if root.get("lat") is not None:
            muss = (root["lat"], root["lon"])
    net_codes = [c for c in order if (by_code[c].get("type") or "").upper() != "MUSS"]
    net_codes = net_codes or order

    def fit(xs, ys):
        mnx, mxx = min(xs), max(xs)
        mny, mxy = min(ys), max(ys)
        sx, sy = max(mxx - mnx, 1e-6), max(mxy - mny, 1e-6)
        return mnx, mxx, mny, mxy, min(GEO_W / sx, GEO_H / sy, GEO_MAX_SCALE)

    net_x = [local[c][0] for c in net_codes]
    net_y = [local[c][1] for c in net_codes]
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

    cx_pg = sum(P(local[c])[0] for c in net_codes) / len(net_codes)
    cy_pg = sum(P(local[c])[1] for c in net_codes) / len(net_codes)
    taken = []
    rmu_terms = {}       # (unit, way) -> (terminal cell, centre) for the cable ends

    # ---- where each node meets the line ------------------------------------
    # The line runs between these points and follows the recorded bends. A symbol
    # that had to move off a crowded spot is joined back to its own point by a
    # short service tap - a DTC taps the line, it is not plastered onto it.
    #
    # A topology drawing's source is the MUSS row itself. When that box is drawn at
    # the page edge for readability, its children hang off the drawn box and not off
    # a point a kilometre past the paper - lines run off a printed page.
    root_is_muss = bool(topo and root_id
                        and (by_code.get(root_id, {}).get("type") or "").upper() == "MUSS")
    tap, tap_id = {}, {}
    for code in order:
        px, py = P(local[code])
        tap[code] = (px, py)
        if code == root_id and root_is_muss and off_map_muss:
            continue                 # placed with the MUSS box, below
        tap_id[code] = d.vertex("", px - 1.5, py - 1.5, 3, 3, ST_HIDDEN)

    def along(code):
        """Unit vector from the parent towards the node - the line's direction."""
        p = links.get(code, (None, ""))[0]
        if p and p in tap:
            dx, dy = tap[code][0] - tap[p][0], tap[code][1] - tap[p][1]
        else:
            dx, dy = tap[code][0] - cx_pg, tap[code][1] - cy_pg
        n = math.hypot(dx, dy) or 1.0
        return dx / n, dy / n

    # ---- equipment symbols -------------------------------------------------
    symbols, moved = {}, set()
    for code in order:
        item = by_code[code]
        px, py = tap[code]
        kind = (item.get("type") or item.get("kind") or "NODE").upper()
        if kind == "MUSS":
            continue                  # the source is placed after the equipment
        dx, dy = along(code)
        perp = (-dy, dx)              # a DTC hangs off the side of the line
        if perp[0] * (px - cx_pg) + perp[1] * (py - cy_pg) < 0:
            perp = (-perp[0], -perp[1])
        if kind in ("POLE", "JUNCTION", "FUSE", "LA", "NODE"):
            symbols[code] = d.vertex("", px - 3, py - 3, 6, 6, ST_POLE)
            taken.append((px - 3, py - 3, px + 3, py + 3))
            place_label(d, px, py, code, taken, bw=6, bh=6, w=54, h=11, style=ST_TINY)
            continue
        if kind == "GOS":
            # drawn on the line, so the sheet shows where the break really is
            is_open = (item.get("normally") or "").upper() == "NO"
            if is_open:
                # a blade hinged on the line and lifted: the open point of the ring
                ang, blade = math.radians(42.0), 30.0
                tip = (px + math.cos(ang) * blade, py - math.sin(ang) * blade)
                cx0, cy0 = (px + tip[0]) / 2, (py + tip[1]) / 2
                # mxGraph rotates about the shape's centre, so the rect is centred
                # on the blade's midpoint and the hinge lands back on the line
                symbols[code] = d.vertex("", cx0 - blade / 2, cy0 - 2, blade, 4,
                                         ST_GOS_OPEN + f"rotation={math.degrees(ang):.0f};")
                taken.append((min(px, tip[0]) - 6, min(py, tip[1]) - 6,
                              max(px, tip[0]) + 6, max(py, tip[1]) + 6))
                d.vertex("", px - 3, py - 3, 6, 6, ST_POLE)   # the hinge
            else:
                symbols[code] = d.vertex("", px - 15, py - 2, 30, 4, ST_GOS)
                taken.append((px - 15, py - 6, px + 15, py + 6))
            txt = (f"<b>{code}</b> (open)<br>{crop_name(item['name'], 18)}" if is_open
                   else f"<b>{code}</b><br>{crop_name(item['name'], 18)}")
            place_label(d, px, py, txt, taken, bw=30, bh=4, w=104, h=24)
            continue
        if kind == "RMU" and topo:
            # the five ways are recorded, so they are drawn: a terminal each, and
            # the cable that lands on one starts there instead of at the box
            used = {o["way"] for o in items if o.get("parent") == code and o.get("way")}
            cw, ch = RMU_CLUSTER[0] * 2, RMU_CLUSTER[1] * 2
            bx, by, rect = free_slot(px, py, cw, ch, taken)
            taken.append(rect)
            rect, terms, box_id = draw_rmu(d, (ways or {}).get(code, {}), used, bx, by)
            taken.append(rect)
            symbols[code] = box_id
            for way, cell_pt in terms.items():
                rmu_terms[(code, way)] = cell_pt
            if abs(bx - px) > 2 or abs(by - py) > 2:
                moved.add(code)
            place_label(d, bx, by, f"<b>{crop_name(item['name'], 26)}</b>", taken,
                        bw=cw, bh=ch, w=150, h=12)
            continue
        if kind == "RMU":
            bw, bh, style, face = 90, 24, ST_RMU, "RMU"
            txt = (f"<b>{crop_name(item['name'], 24)}</b><br>"
                   f"<font style='font-size:7px'>5-way</font>")
        elif kind == "HT" or code.upper().startswith("HT-"):
            bw, bh, style, face = DTC_W, DTC_H, ST_HT, "HT"
            tail = f" ({item['kva']} kVA)" if item.get("kva") else ""
            txt = f"{crop_name(item['name'])}<br><b>{code}</b>{tail}"
        else:
            bw, bh, style = DTC_W, DTC_H, ST_DTC
            face = str(item["kva"]) if item.get("kva") else "?"
            txt = f"{crop_name(item['name'])}<br><b>{code}</b>"
        # a transformer taps the line: it hangs beside it on a service tap, so the
        # run itself never crosses a symbol
        off = (-perp[0] * 26, -perp[1] * 26)
        bx, by, rect = free_slot(px + off[0], py + off[1], bw, bh, taken, first=perp)
        symbols[code] = d.vertex(face, bx - bw / 2, by - bh / 2, bw, bh, style)
        taken.append(rect)
        if abs(bx - px) > 2 or abs(by - py) > 2:
            moved.add(code)
        place_label(d, bx, by, txt, taken, bw=bw, bh=bh)

    trunk_pg = [P(p) for p in trunk_pts]

    # ---- the source ---------------------------------------------------------
    # Placed after the equipment so it can be kept clear of it.
    base = f"110/11 KV<br>{station.replace('110/11KV ', '') or 'MUSS'}"
    if trunk_pts:                    # a surveyed route starts at the MUSS
        where = trunk_pg[0]
        anchor_id = d.vertex("", where[0] - 1.5, where[1] - 1.5, 3, 3, ST_HIDDEN)
    else:
        anchor_id = tap_id.get(root_id) if root_id else None
        where = P(_metres(muss[0], muss[1], lat0, lon0)) if muss else None
    box = None
    if where is None:
        # nothing records where the source is: hang it off the network, and say so
        ax, ay = tap.get(root_id, (cx_pg, cy_pg))
        vx, vy = ax - cx_pg, ay - cy_pg
        nv = math.hypot(vx, vy) or 1.0
        base_ang = math.atan2(vy / nv, vx / nv)
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
    else:
        mx, my = where
        if off_map_muss:
            # the substation is kilometres from the load: keep the network readable and
            # say so, rather than shrinking the drawing to fit one point off the page
            mx = min(max(mx, GEO_X + MUSS_W / 2 + 6), GEO_X + GEO_W - MUSS_W / 2 - 6)
            my = min(max(my, GEO_Y + MUSS_H / 2 + 6), GEO_Y + GEO_H - MUSS_H / 2 - 6)
            mxm, mym = _metres(muss[0], muss[1], lat0, lon0)
            km = math.hypot(mxm - sum(net_x) / len(net_x),
                            mym - sum(net_y) / len(net_y)) / 1000
            base += (f"<br><font style='font-size:7px'>({km:.1f} km beyond the edge, "
                     f"not to scale)</font>")
        _, _, box = free_slot(mx, my, MUSS_W, MUSS_H + 8, taken,
                              dists=(0, 36, 72, 110, 150, 200),
                              bounds=(20, 130, PAGE_W - 20, TABLE_TOP - 96))
    if root_is_muss and off_map_muss:
        tap[root_id] = (box[0] + MUSS_W / 2, box[1] + (MUSS_H + 8) / 2)
        tap_id[root_id] = d.vertex("", tap[root_id][0] - 1.5, tap[root_id][1] - 1.5,
                                  3, 3, ST_HIDDEN)
    m_id = d.vertex(base, box[0], box[1], MUSS_W, MUSS_H + 8, ST_MUSS)
    taken.append(box)
    if anchor_id and not (root_is_muss and off_map_muss):
        box_mid = (box[0] + MUSS_W / 2, box[1] + (MUSS_H + 8) / 2)
        d.edge(m_id, anchor_id, ST_DASH if off_map_muss else ST_TRUNK,
               points=ortho_route(box_mid, tap.get(root_id, box_mid)) or None)

    # ---- the network --------------------------------------------------------
    # Drawn last, so a tap lands on its symbol: the trunk heavy, laterals lighter,
    # anything guessed dashed.
    if trunk_pts:
        route_ids = []
        for pt in trunk_pg:
            route_ids.append(d.vertex("", pt[0] - 1.5, pt[1] - 1.5, 3, 3, ST_HIDDEN))
        for a, b in zip(route_ids, route_ids[1:]):
            d.edge(a, b, ST_GEO)
    for code in order:
        item = by_code[code]
        parent_code, why = links.get(code, (None, ""))
        if trunk_pts:                # surveyed route: tap the nearest point on it
            px, py = tap[code]
            best = None
            for k in range(len(trunk_pg) - 1):
                dist, ax, ay = _project(px, py, *trunk_pg[k], *trunk_pg[k + 1])
                if best is None or dist < best[0]:
                    best = (dist, ax, ay)
            _, ax, ay = best
            a_id = d.vertex("", ax - 1.5, ay - 1.5, 3, 3, ST_HIDDEN)
            if code in symbols:
                d.edge(a_id, symbols[code], ST_STUB)
            continue
        if not parent_code or parent_code not in tap_id or code not in symbols:
            continue
        # a cable off an RMU leaves by the way the workbook names, so the line
        # starts at that terminal and not at the unit's box
        src_id, src_pt = tap_id[parent_code], tap[parent_code]
        if (parent_code, (item.get("way") or "")) in rmu_terms:
            src_id, src_pt = rmu_terms[(parent_code, item["way"])]
        spur = (item.get("spur") or "T").upper()
        style = ST_GUESS if why else (ST_TRUNK if spur == "T" else ST_GEO)
        # a bend off the edge would drag the line past the paper: the span from an
        # off-map substation is not to scale, so its shape is not used
        if parent_code == root_id and root_is_muss and off_map_muss:
            bends_here = []
        else:
            bends_here = list((bends or {}).get((parent_code, code), []))
        surveyed = [P(_metres(lat, lon, lat0, lon0)) for _, lat, lon in bends_here]
        if surveyed:                  # a recorded shape: its corners, as surveyed
            route = surveyed
        else:                         # nothing recorded: right angles, not a diagonal
            route = ortho_route(src_pt, tap[code])
        d.edge(src_id, tap_id[code], style, points=route or None)
        if code in moved:
            d.edge(tap_id[code], symbols[code], ST_STUB)
        # a lateral is named at the point it leaves the trunk, once SPURS names it
        spur_tag = (spur_labels or {}).get(spur)
        if (spur_tag and spur and spur != "T"
                and (by_code.get(parent_code, {}).get("spur") or "T").upper() != spur):
            place_label(d, (tap[parent_code][0] + tap[code][0]) / 2,
                        (tap[parent_code][1] + tap[code][1]) / 2,
                        f"<b>{spur}</b> {spur_tag}", taken, bw=8, bh=8, w=110, h=11,
                        style=ST_TINY)

    # ---- furniture ---------------------------------------------------------
    source = ("surveyed route" if trunk_pts else
              ("connection workbook" if topo else "inferred trunk"))
    if guessed:
        source += f", {len(guessed)} link(s) guessed"
    unpositioned_block(d, bad)
    scale_bar(d, scale)
    summary_table(d, dtcs)
    sheet_header(d, title_text or
                 f"SINGLE LINE DIAGRAM OF {label.upper()} in Unit-4, CSD-2, SMG",
                 subtitle or (f"{station}  MUSS   |   geographical layout from GPS"
                              f"   |   {source}"
                              f"   |   {len(order)} of {len(items)} nodes placed"
                              f"   |   {date.today():%d-%m-%Y}"))
    return wb, dict(scale=scale, per_row=None, rows=None, source=source,
                    positioned=len(order), unpositioned=len(bad), guessed=len(guessed))


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
    ap.add_argument("--geo", action="store_true",
                    help="place every node at its GPS position, connected as the workbook records")
    ap.add_argument("--muss", help='substation coordinates "lat,lon" for the geo layout')
    ap.add_argument("--extent", choices=["cluster", "all"], default="cluster",
                    help="geo layout extent: fit the DTC network (default) or include the MUSS")
    ap.add_argument("--route", help="surveyed pole coordinates (csv/xlsx: NODE_ID,LAT,LONG)")
    ap.add_argument("--topology", help=f"connection workbook (default: {_display(TOPOLOGY_DEFAULT)})")
    ap.add_argument("--no-topology", action="store_true",
                    help="ignore the connection workbook and infer the trunk again")
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

    topo = None if args.no_topology else read_topology(args.topology or TOPOLOGY_DEFAULT)
    if args.topology and topo is None:
        raise SystemExit(f"Topology workbook not found: {args.topology}")

    if args.list or not args.feeder:
        cnt = Counter(it["feeder"] for it in dtcs)
        print("Feeders in the master:")
        for f, c in cnt.items():
            with_gps = sum(1 for it in dtcs if it["feeder"] == f and it["lat"] is not None)
            print(f"  {f:<14} {c:>3} DTCs  ({with_gps} with GPS)")
        if nodes:
            print("\nNon-DTC nodes (sheet '%s'):" % SHEET_NODES)
            for n in nodes:
                print(f"  {n['kind']:<5} {n['name']:<28} {n['lat']:.6f}, {n['lon']:.6f}")
        if topo:
            print(f"\nConnections (writer '{_display(topo['path'])}'):")
            for f in cnt:
                rows = [n for n in topo["nodes"] if _norm(n["feeder"]) == _norm(f)]
                basis = Counter(n["basis"] or "UNKNOWN" for n in rows)
                seen = ", ".join(f"{k} {v}" for k, v in basis.most_common())
                print(f"  {f:<14} {len(rows):>3} rows  ({seen})")
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

    # ---- the connection workbook -------------------------------------------
    items, root_id, ways, bends, spur_labels = None, None, None, None, None
    if topo:            # both layouts are drawn from the record when it is there
        sel_topo = topology_for_feeder(topo, args.feeder, station)
        if sel_topo:
            errs, warns = validate_topology(sel_topo, topo,
                                            {it["code"].upper() for it in dtcs + hts})
            for w in warns:
                print(f"  ! {w}")
            if errs:
                raise SystemExit(
                    f"{_display(topo['path'])} disagrees with itself - nothing was drawn:\n  "
                    + "\n  ".join(errs)
                    + "\nFix those rows; a drawing must not contradict the record.")
            items, only_master, moved = topology_items(sel, sel_ht, sel_topo, rmus)
            for m in moved:
                print(f"  ! {m}")
            if only_master:
                print(f"  ! {', '.join(only_master)}: not in the workbook - drawn from the "
                      "master with a guessed parent")
            root = topology_root(sel_topo, station)
            root_id = root["id"] if root else None
            bends = topo["bends"]
            ways = topo["ways"]
            spur_labels = {sid: lab for (f, sid), lab in topo["spurs"].items()
                           if _norm(f) == _norm(args.feeder)}
            kind_count = Counter(i["type"] for i in items)
            print(f"  topology    : {_display(topo['path'])}  |  "
                  + ", ".join(f"{v} {k}" for k, v in sorted(kind_count.items())))
            if muss is None and root and root["lat"] is not None:
                muss = (root["lat"], root["lon"])
                print(f"  muss        : {muss[0]:.6f}, {muss[1]:.6f}  (row {root['id']})")

    if args.geo:
        wb, info = build_geo(sel, sel_ht, label, station, muss=muss, route=route,
                             rmus=rmus, extent=args.extent, title_text=args.title,
                             topo=items, root_id=root_id, bends=bends,
                             spur_labels=spur_labels, ways=ways)
    else:
        wb, info = build_schematic(sel, sel_ht, label, station, rmus=rmus,
                                   title_text=args.title, topo=items, root_id=root_id,
                                   spur_labels=spur_labels)

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
        print(f"  layout      : geo  |  {info['positioned']} nodes placed, "
              f"{info['unpositioned']} unpositioned  |  1 m = {info['scale']:.2f} px")
        print(f"  source      : {info['source']}")
    else:
        print(f"  layout      : schematic  |  {info['per_row']} per lateral x {info['rows']} laterals")
        print(f"  source      : {info['source']}")
    print(f"  written     : {_display(out)}")
    if args.preview:
        preview(xml, os.path.splitext(out)[0] + ".png")


if __name__ == "__main__":
    main()
