# Session handoff — MESCOM SLD tool

**Date:** 26 October 2026 · **Machine:** Windows, Python 3.13.5 · **Shell:** bash (Git Bash)
**Folder:** `D:\Office PC\D DRIVE\ESTIMATES\SLD TOOL\`

> **Update, later the same day:** all templates now live in `template\` and generated sheets in
> `outputs\<date>\`. Two files are new — `template\TOPOLOGY TEMPLATE.xlsx` (the connection data to
> fill in) and `TOPOLOGY_DATA_REQUIREMENTS.md` (why it is needed and how to collect it). The newest
> CHANGELOG entry covers the move. File locations in sections 2–4 below describe the earlier state.
>
> **Update 3:** an RMU is now modelled as a **five-way unit** — `ODin`, `ODout`, `VL1`, `VL2`, `VL3`.
> `NODES` gained a `WAY` column (which terminal a child leaves by) and the workbook gained an
> `RMU WAYS` sheet (the terminals themselves, with `CONNECTED` to mark empty ways). Only `ODin` is
> always cabled; any outgoing way — `ODout` included — may be empty. RMU symbols are labelled
> `5-way`. Sections 2–5 below predate that change; see the two newest CHANGELOG entries.
>
> **Update 4:** the two workbooks moved into `field_data\` — `DTC DETAILS U4.xlsx` (the master the
> tool reads by default) and `TOPOLOGY.xlsx` (the copy to fill in). `template\` now holds only the
> blanks, so a template rebuild cannot reach the field work. Section 2 below lists the earlier,
> flat layout.
>
> **Update 5:** `README.md`, `CHANGELOG.md` and this file now live in `docs\`. Paths in sections
> 2–4 below are the earlier root-level ones; the content is history, not current layout.
>
> **Update 6:** `TOPOLOGY_DATA_REQUIREMENTS.md` has joined them in `docs\`, so all four documents
> now sit together. The workbook READ ME points at the new path.
>
> **Update 7 (27 October 2026):** the `--geo` layout is now **drawn from the connection workbook**
> (`field_data/TOPOLOGY.xlsx`): parent links, poles, the GOS, the RMU and its cabled ways, `BENDS`,
> spur labels, plus a validation pass. §2–§5 below were rewritten for it; §7 is this session's
> verification. The reason was a reported defect: on `Mandli F13` the RMU, `GOS-01` and `P-001` were
> absent from the sheet and `SCD-10` appeared fed from two DTCs, because the layout could only read
> the master and had to infer the tree.
>
> **Update 8 (28 October 2026):** four drawing steps landed, one commit each — runs are routed at
> right angles with square corners; a `GOS` shows `NORMALLY` with the load behind an open point
> named; an RMU is drawn as **a panel divided into its ways** (five, or three for a three-way unit);
> and **the schematic is
> drawn from the same tree** (depth-first order, `from <parent>` under each box, equipment symbols,
> spur tags at the branch point). §5 and §7 below describe the current behaviour; §9 is the updated
> list of what is left.
>
> **Update 9 (29 October 2026):** connectors now meet a symbol at **the midpoint of one of its four
> edges**, join symbol to symbol, and are routed around the other symbols by a new
> `ConnectorRouter` (short right-angled shapes scored against every symbol and label, with an 8 px
> grid search when a box blocks them all). The surveyed-point dot is now visible, its leader is
> square and picks a clean edge, and labels sit on the far side of a tapped symbol and are routing
> obstacles. §5.6/§5.8 and §7 are current; the data issues in §8/§9 are unchanged.
>
> **Update 10 (29 October 2026):** `docs/FIELD_CAPTURE.md` is new — a proposal for collecting the
> tree in the field and deriving the rest. Nothing in it is implemented and no code changed. Its
> conclusions: the gap is connectivity, not position (114 of 129 rows are `INFERRED` while 113 of
> 119 DTCs already have GPS); a feeder walked out and back is an Euler tour, so **the order nodes
> are tapped rebuilds every parent** without any distance test; a GPS fix is only needed to confirm
> which DTC is in front of you, and order plus pole counts plus pole numbers do the rest; the
> pipeline's output is the existing `field_data/TOPOLOGY.xlsx`, so the generator needs no change.
> Open decisions are listed in its §11.

---

## 1. Snapshot

The user is a MESCOM (Shivamogga) distribution engineer who was drawing feeder single line diagrams
by hand in Word — hundreds of loose shapes per sheet. This session replaced that with a data-driven
generator: **two spreadsheets in, an editable `.drawio` single line diagram out**, in two layouts
(schematic and geographical).

The tool works, and **both** sheets are now drawings of the connection record rather than likenesses
of it: the geographical sheet draws it as lines, the schematic as order. All eight sheets were
regenerated on 29 October 2026 and pass structural validation; every geographical connector now
leaves and enters at an edge midpoint of its symbol. MUSS and RMU positions come from the
master's `NON DTC NODES` sheet; parents, poles, the GOS, the RMU's terminals, the spur tags and the
bend points come from `field_data/TOPOLOGY.xlsx`. The open items are data clean-up in that workbook
(three RMUs without a tap pole, one stale `RMU WAYS` row, a blank `NORMALLY` on `GOS-01`), the
remaining missing GPS, and pole-level detail for the feeders that still have none.

---

## 2. What was delivered

| File | Size | Role |
|---|---|---|
| `sld_generator.py` | ~138 KB / 2980 lines | The tool. Single file, no local packages. |
| `field_data/DTC DETAILS U4.xlsx` | 51 KB | Master workbook — the register of record, read by default. |
| `field_data/TOPOLOGY.xlsx` | 37 KB | The connection workbook the user fills in. **Read by both layouts; never written by the tool.** |
| `template/MESCOM_SLD_TEMPLATE.drawio` | 17 KB | Blank A3 sheet + symbol library, two tabs. |
| `template/TOPOLOGY TEMPLATE.xlsx` | 28 KB | The blank connection workbook (`--topology-template`). |
| `template/route_template.csv` | <1 KB | Format sample for `--route` (values are invented). |
| `docs/README.md` | 41 KB | User-facing manual. |
| `docs/TOPOLOGY_DATA_REQUIREMENTS.md` | 27 KB | The field specification for the connection data. |
| `docs/FIELD_CAPTURE.md` | 18 KB | Proposal: field capture → workbook → drawings. **Nothing implemented.** |
| `docs/CHANGELOG.md` | 45 KB | Every change, newest first. |
| `docs/SESSION_HANDOFF.md` | 24 KB | This document. |
| `CLAUDE.md` | 15 KB | Working guide for an agent picking the repository up. |
| `outputs/<date>/*.drawio` + `.png` | — | Generated sheets, one folder per run day, git-ignored. |

Latest output folder `outputs\2026-10-27\`: four geographical sheets (`MRS F1`, `Mandli F13`,
`Mandli F22`, `Mandli F3`) and four schematic ones.

---

## 3. How to run it

```bash
cd "D:\Office PC\D DRIVE\ESTIMATES\SLD TOOL"

python sld_generator.py --list                          # feeders, GPS and connection coverage
python sld_generator.py --feeder "Mandli F13"           # schematic
python sld_generator.py --feeder "MRS F1" --geo         # geographical, drawn from the workbook
python sld_generator.py --feeder "MRS F1" --geo --preview
python sld_generator.py --feeder "MRS F1" --geo --no-topology   # back to the inferred trunk
python sld_generator.py --template                      # rebuild template + library
```

Full regeneration of everything currently delivered (both layouts, all four feeders):

```bash
python sld_generator.py --template --preview
for f in "MRS F1" "Mandli F13" "Mandli F22" "Mandli F3"; do
  python sld_generator.py --feeder "$f" --name "$f" --geo
  python sld_generator.py --feeder "$f" --name "$f"
done
```

Dependencies: `openpyxl` (required), `pymupdf` (only for `--preview`). Both installed here.

---

## 4. Architecture

One file, top-to-bottom:

```
constants & styles      L52–165   page size, geometry, draw.io styles, EDGE_SIDES
data helpers            L167–310  _txt, _num, _kva, crop_name, read_nodes, find_station,
                                  assign_nodes_to_feeders, read_master, _norm, pick
connection workbook     L311–702  read_topology, topology_for_feeder, validate_topology,
                                  topology_items, resolve_links, draw_rmu
diagram model           L704–773  Page (vertex/edge emitters, edge ports), Workbook (→ XML)
sheet furniture         L774–921  sheet_header, summary_table, unpositioned_block, scale_bar,
                                  _overlaps, _clamp_label, place_label, free_slot
schematic layout        L922–1101 choose_grid, tree_order, rmu_way_text, build_schematic
geo helpers             L1102–1230 clean_points, _metres, prim, tree_ends, read_route,
                                  ortho_route, _project
connector router        L1232–1650 _port, _seg_crosses, _shapes, _search, ConnectorRouter
geo layout              L1651–2119 build_geo
templates               L2120–2715 build_template, build_topology_template, _infer_topology
preview                 L2716–2797 preview (ports, border clipping, dots)
CLI                     L2798+    main
```

(Line numbers are as of the handoff; they shift with edits.)

### Data flow

```
field_data/DTC DETAILS U4.xlsx
        ──read_master──► DTC/HT dicts {code, name, kva, feeder, station, kind, lat, lon}
        ──read_nodes──► MUSS/RMU dicts {name, kind, lat, lon}          (HT rows: no lat/lon)
                                            │
field_data/TOPOLOGY.xlsx                    │ pick(feeder)
        ──read_topology──► nodes {id, type, name, feeder, lat, lon, parent, spur, way, kva,
        │                         normally, basis, row}, ways{rmu: {way: {...}}},
        │                         bends{(from,to): [(seq, lat, lon)]}, spurs{(feeder, id): label}
        ──topology_for_feeder──► the subtree this feeder draws + the MUSS above it
        ──validate_topology──► errors (stop) / warnings (print)
        ──topology_items──► merged with the register: master owns a DTC's name/rating/GPS
        ──resolve_links──► {code: (parent, why guessed)}                      │
                                            └───────────────┬───────────────┘
                    ┌───────────────────────────────────────┴───────────────┐
                    ▼                                                       ▼
        build_schematic(...)                                     build_geo(...)
        grid of laterals                          project + workbook links, or MST, or --route
                    └───────────────────────┬───────────────────────────────┘
                                            ▼
                                   Workbook.to_xml()  →  outputs/<date>/<label> SLD.drawio
                                            ▼
                              preview(xml, ...)  →  PNG (optional)
```

### Diagram model

`Page.vertex(value, x, y, w, h, style)` and `Page.edge(src, tgt, style, points)` append raw
`<mxCell>` XML; `Workbook.to_xml()` wraps pages into an `<mxfile>` with A3 landscape geometry
(1654 × 1169 draw.io units = 420 × 297 mm at 96 dpi). Every label is stored as escaped HTML
(`&lt;br&gt;`, `&lt;b&gt;`), which is what draw.io expects for `html=1` labels.

Connectors carry `source`/`target` ids, so in draw.io a DTC can be dragged and its line follows —
the reason the tool beats the Word file.

---

## 5. The two layouts

### Schematic (`build_schematic`, default)

Vertical distributor at `x = 110`, horizontal laterals from `x = 150` to `1610`, MUSS box at
`(1370, 100)`, summary table bottom right, `UNPOSITIONED`-free (every node is drawn).

**Order.** `tree_order()` walks `NODES` depth first from the MUSS in workbook row order, so a parent
is drawn before its children and a spur's nodes stay together in one run of the grid. Without a
workbook (`--no-topology`) the register's code order is used. `rmu_way_text()` turns a unit's ways
into words for the grid (`5-way - out VL1`), and `spurs_at` records the node each spur leaves from.

Grid: `choose_grid(n)` → `per_row = clamp(ceil(n/3), 4, 14)` reduced until the slot width is at
least 92 px; `rows = ceil(n / per_row)`; row height up to 230 px. It counts **nodes**, so 22 nodes on
`Mandli F13` give 8 × 3 and 70 on `MRS F1` give 14 × 5.

Each node: junction dot on the lateral → 30 px stub → the symbol for its `TYPE` (44 × 22 square with
the kVA, a white HT square, a `90 × 24` RMU box, a 6 px pole dot, a 30 × 4 GOS bar) → label 22 px
below (name over code, then `from <parent>` in 7 px blue) → the spur tag under the node it leaves
from. A node with no parent is drawn after the trunk and reported once per reason, with its code
listed. **The row bus is a convention, not wiring:** only the `from` line states a parent.

### Geographical (`build_geo`, `--geo`)

1. `clean_points` splits usable / unusable coordinates and rejects anything >5 km from the median.
2. `_metres` performs a local equirectangular projection (111.32 km/° lon, 110.54 km/° lat).
3. **Links**, in priority order:
   * `--route` — the supplied polyline is drawn and every symbol attaches by `_project` onto its
     nearest segment (surveyed data beats a recorded tree; the run says so);
   * else the connection workbook — `resolve_links()` gives each node the parent named in
     `PARENT_ID`, or, when a parent has no coordinates, the nearest ancestor that has them, or, for a
     node with no usable parent, the nearest node that is not downstream of it. Both fallbacks are
     dashed (`ST_GUESS`) and reported;
   * else `prim` builds a minimum spanning tree over the metres, rooted at the DTC nearest the MUSS
     (`tree_ends` when there is no MUSS position).
4. **Page transform**: the fit excludes `TYPE = MUSS` nodes and the MUSS coordinate, so a substation
   1.7–2.6 km away cannot shrink the drawing; scale = `min(1460/spanx, 735/spany, 3.2)`, north up,
   centred.
5. **Tap points.** One 3 × 3 dot per node at its own surveyed page position. It is drawn (`ST_TAP`),
   not hidden, and a symbol that had to move off that spot is tied back to it by a thin leader; the
   dot gives the leader an end. `BENDS` rows become edge waypoints.
6. **Routing.** Every connector is asked of `ConnectorRouter`, which picks a port on each end and a
   square path between them. A **port is the midpoint of a symbol edge** (`EDGE_SIDES`), written into
   the edge style as `exitX/exitY` and `entryX/entryY`, so draw.io keeps the attachment when the
   symbol is dragged. The router scores every port pair on the straight, one-bend and two-bend Z
   shapes; a segment through a symbol costs more than any detour, a label is a smaller penalty. A
   14 px square lead off the port (shortened when a neighbour is closer) takes the line clear before
   its first turn. When every simple shape is blocked it runs an 8 px grid A* with symbol cells made
   ruinously expensive; the result is folded into long runs, and if no way round exists the least-bad
   simple shape is kept and counted (`connectors : 64 routed, 1 round a box, 0 still crossing one`).
   A `BENDS` span is still drawn through its recorded points unchanged; the KOTE ROAD panel case that
   forced the search is in §7.
7. **Symbols**, per `TYPE`: `DTC` = 44 × 22 square with the kVA, `HT` = white square, `POLE`/`JUNCTION`
   = 6 × 6 dot with its code in 6 px text, `GOS` = 30 × 4 bar sitting on the line (a lifted hollow
   blade when `NORMALLY = NO`), `RMU` = `draw_rmu()`'s panel divided into one `34 × 26` portion per
   way — `ODin` first, `ODout` last, the VLs between, filled where a cable lands and white `empty`
   where not — `MUSS` = source box, `NODE` = DTC-shaped fallback. A boxed symbol that has to move off
   a crowded spot is placed by `free_slot()` (perpendicular-first for a DTC) and joined back to its
   tap by a square leader that enters the nearest edge midpoint not crossed by another symbol; a
   cable off an RMU starts at the portion, on the edge facing its node, and the anchor for that edge
   is created by the link that uses it.
8. `place_label` tries 12 positions per label, starting on the far side of a moved symbol's tap, and
   scores overlap area against everything already placed (symbols and labels), clamping to the
   drawing area. The rectangle it takes joins `label_boxes`, so the connector router treats text as
   an obstacle.
9. The MUSS box is drawn after the equipment. If including the substation would cost more than 30 %
   of the scale, the network keeps its own scale and the box is drawn at the page edge in its true
   direction, labelled with the real distance — and, in a workbook run, the root's children hang off
   that drawn box so no line runs past the paper (`--extent all` forces the strict fit instead).
10. `scale_bar`, north arrow, `unpositioned_block` (naming each node's parent), `summary_table`,
    `sheet_header` finish the sheet. The subtitle names the source of the shape and, when links were
    guessed, how many.

Key constants: `GEO_X, GEO_Y, GEO_W, GEO_H = 120, 205, 1460, 735`; `GEO_MAX_SCALE = 3.2 px/m`;
`DTC_W, DTC_H = 44, 22`; `LBL_W, LBL_H = 100, 26`; `VALID_CHILD_WAYS = (ODout, VL1, VL2, VL3)`;
`ENVELOPE_LAT = (13.5, 14.3)`, `ENVELOPE_LON = (75.1, 76.0)`.

---

## 6. The data

Master: `DTC DETAILS U4.xlsx`, sheets `ALL` (119 DTC rows) and `HT` (5 rows).

| Column | Header | Index | Used for |
|---|---|---|---|
| B | SCD | 1 | code prefix (`SCD`/`SCF`/`SCE`) |
| C | DTC_CODE | 2 | number, zero-padded to two digits |
| E | DTC_Name | 4 | label, cropped to 22 chars |
| F | Feeder | 5 | feeder matching |
| I | KVA | 8 | rating inside the square |
| J | Station | 9 | MUSS box text (`110/11KV MRS` / `MANDLI`) |
| AD | GPS LAT | 29 | decimal degrees |
| AE | GPS LONG | 30 | decimal degrees |

HT sheet: `RR No.` (4) = code, `NAME` (5), `CD(KVA)` (7), `Feeder` (11).

Sheet `NON DTC NODES`: `NODES` (0), `GPS LAT` (1), `GPS LONG` (2), `GPS HYPERLINK` (3, ignored).
The type is inferred from the name — `MUSS` or `RMU` — and a MUSS is matched to the station text in
column `J` of sheet `ALL` (`110/11KV MRS` → `110/11KV MRS MUSS`). An RMU is assigned to the feeder
whose nearest DTC is closest to it. Current rows: two MUSS (MRS 13.910279, 75.595896; Mandli
13.917882, 75.562761) and two RMU (Kote Road, Gangothti College).

Feeder population: **MRS F1 63**, **Mandli F22 32**, **Mandli F13 19**, **Mandli F3 5**.
GPS coverage: MRS F1 63/63, F22 29/32, F13 17/19, F3 4/5 — 113 of 119 overall, after the user
filled in nine missing coordinates mid-session.

Matching rule: feeder names compare case-insensitively with dashes and repeated spaces collapsed
(`_norm`), corrected from an earlier substring test that could have matched `Mandli F13` to
`Mandli F1`. A MUSS row is matched to the master's station text by name containment
(`_station_matches`), which is how `110/11KV MANDLI` finds `110/11KV MANDLI MUSS`.

Connection workbook `field_data/TOPOLOGY.xlsx`, read by both layouts. Columns are addressed by
position:
`NODES` A–O = `NODE_ID` 0, `TYPE` 1, `NAME` 2, `FEEDER` 3, `LAT` 4, `LONG` 5, `PARENT_ID` 6,
`SPUR_ID` 7, `WAY` 8, `KVA` 9, `PHASE` 10, `NORMALLY` 11, `CONDUCTOR` 12, `BASIS` 13, `REMARKS` 14;
`RMU WAYS` A–G = `RMU_ID` 0, `WAY` 1, `DIRECTION` 2, `CONNECTED` 3, `NORMALLY` 4, `LABEL` 5,
`NOTES` 6; `BENDS` A–F = `FROM_ID` 0, `TO_ID` 1, `SEQ` 2, `LAT` 3, `LONG` 4, `NOTES` 5;
`SPURS` A–H = `FEEDER` 0, `SPUR_ID` 1, `LABEL` 2, `FROM_NODE` 3, `IS_TRUNK` 4, `DTC_COUNT` 5,
`KVA_TOTAL` 6, `NOTES` 7. Row numbers are kept on every node so a message can name the exact row.
`EXAMPLE` and `LISTS` are ignored by the reader; `READ ME` is documentation.

Coverage after the user's first pass: MRS F1 70 rows, Mandli F22 32, `Mandli F13` 22 (21 + one row
spelled `MANDLI F13`), Mandli F3 5, two rows with no feeder at all — 131 rows, of which `SOURCE` 2,
`INFERRED` 114, `UNKNOWN` 15. Hand-added equipment: three `RMU` rows (`KOTE-ROAD` and
`GANGOTHTI-COLLEGE` on MRS F1, `RAMANNA-SHETTY-PARK` on F13), plus `GOS-01` and `P-001` on F13.
Three `RMU WAYS` units and 23 `SPURS` rows are recorded; `BENDS` is still empty.

---

## 7. Verification performed (27 October 2026)

* **Structure** of all eight `outputs\2026-10-27\*.drawio` files plus the template — nine files,
  ten diagram pages, since the template holds two tabs — checked with Python + ElementTree:
  **0 dangling connectors, 0 shapes off the A3 page** in every one.
* **The two reported defects**, checked against the user's own command
  (`--feeder "Mandli F13" --name "Mandli F13" --geo --preview`): the RMU, `GOS-01` and `P-001` are
  now drawn (the run reports `19 DTC, 1 GOS, 1 MUSS, 1 POLE, 1 RMU`), and `SCD-10` is joined to
  `SCD-36` only — `SCD-36` is its `PARENT_ID`.
* **Visual**: PNG of every feeder's geographical sheet inspected (F1, F13, F22, F3) plus the F13
  schematic. Three defects were found this way and fixed: the spur label was printing as the sheet
  title, the off-map MUSS ran two trunk lines off the corner of the paper, and `SCD-68` — a rejected
  coordinate — was still being attached to `MUSS-MRS` in the link list.
* **Workbook features** exercised on a throwaway copy of `field_data/TOPOLOGY.xlsx` (the user's file
  was never edited): `BENDS` produced 2 edge waypoints, a `SPURS!LABEL` produced the
  `S1 Chor Bazar lateral` tag, and `RMU WAYS`/`NODES` disagreement produced the warning.
* **Other code paths**: `--no-topology` (MST, 17 nodes placed), `--route template/route_template.csv`
  (surveyed route, 0.76 px/m), `--route` + workbook, `--extent all` (0.85 px/m), `--list`,
  `--template`, `--topology-template --out <temp>` (129 NODES rows rebuilt, template untouched).
* **Console**: warnings name sheet, row and column (`NODES!6 …`, `RMU WAYS!14 …`).
* **Not verified**: opening the `.drawio` files in draw.io itself — still no installation here. The
  XML is the same schema the earlier sheets used and the earlier ones opened; the explicit
  multi-point polylines the router writes are the newest shape not yet seen in the desktop app.

### Verification performed (28 October 2026)

* **Structure**: all eight `outputs\2026-10-27\*.drawio` files regenerated from
  `field_data/TOPOLOGY.xlsx` as it stands, then parsed with ElementTree: **0 dangling connectors,
  0 shapes off the A3 page** in each. Vertices/edges after the RMU panel: MRS F1 239/121 geo and
  244/145 schematic, F13 97/37 and 86/45, F22 112/58 and 122/67, F3 37/8 and 34/12.
* **Routing**: every edge's waypoints measured against its own endpoints on the regenerated geo
  sheets — **58 routed runs, every corner exactly 90°, no diagonal segment on any routed run**. The
  other 166 edges carry no waypoints: service taps, the MUSS feeder, short hops under 70 px and runs
  with under 6 px of off-axis drift, which the router leaves straight on purpose.
* **GOS**: `GOS-01` is blank in the workbook, so every F13 run prints `NORMALLY is blank - drawn
  closed; record NC or NO - 14 DTC(s) / 1951 kVA recorded downstream`. On a throwaway copy with
  `NO`, the blade lifts off the line and the label carries `(open)`.
* **RMU panels**: F13's `VL1` portion is filled and `P-001`'s cable leaves from its top edge, the
  other three are white and `empty`; a `--no-topology` run keeps the plain `5-way` box. A throwaway
  copy of the workbook with `VL2`/`VL3` deleted from `RMU-KOTE-ROAD` draws `ODin | VL1 | ODout` and
  the schematic says `3-way`, which is the whole of the three-way support checked end to end.
* **Schematic**: F13 comes out 8 nodes × 3 laterals (22 nodes), MRS F1 14 × 5 (70 nodes); every box
  carries its `from` line, the RMU its ways, the five MRS F1 HT services their kVA, and the spur tags
  sit under the nodes the spurs leave from. The no-parent report is one line per reason.

### Verification performed (29 October 2026)

* **Structure**: all eight `outputs\2026-10-27\*.drawio` files regenerated from the workbook as it
  stands, then parsed with ElementTree: **0 dangling connectors, 0 shapes off the A3 page** in each.
  Vertices/edges: MRS F1 239/120 geo and 244/145 schematic, F13 97/37 and 86/45, F22 112/55 and
  122/67, F3 37/8 and 34/12.
* **Ports and squareness**: every edge's chain was rebuilt from its `exitX/exitY`/`entryX/entryY`
  fractions (or clipped at the symbol border when it has none) and measured. **0 diagonal segments on
  all eight sheets**; every connector ending on a symbol carries an edge-midpoint port (177 of 240
  ends on MRS F1 are ports; the rest are RMU-portion anchors, position leaders and the surveyed-route
  taps). The four schematic sheets have no explicit ports and none need them — their runs are
  straight bus-to-box stubs that cross the border at the midpoint anyway.
* **Symbols**: **0 connector-through-a-symbol crossings** on F13, F22, F3 and the schematic sheets.
  MRS F1 has **4**, all of them 0.8 pt position leaders in the tightest cluster, not network links.
  One feed (the `25` DTC below KOTE ROAD RMU) is routed round the 170 × 26 panel: the console says
  `64 routed, 1 round a box, 0 still crossing one`, and the panel crossing that existed in the first
  cut is what drove the grid search and the shorter-stub rule.
* **Labels**: 0 crossings on all four schematic sheets, 3 on F13, 2 on F22, 1 on F3 and 20 on MRS F1
  (10 network links, 10 position leaders). Labels are obstacles in the router; the remaining crossings
  are the label density of a 65-node sheet, not a missing check. The four `README` §11 bullets record
  this as a limit.
* **Visual**: PNGs of F13, F22, F3 and MRS F1 inspected, plus the F13 schematic and a
  `--no-topology` run. Defects found this way and fixed were the panel feed that ran across the
  compartments and a VL cable that cut through the panel; the leader diagonals; and hidden RMU
  anchors showing as blue squares in the preview.
* **Other paths**: `--no-topology` (49 of 66 ends are ports; 0 diagonal, 1 symbol crossing — a leader),
  `--route template/route_template.csv` (the surveyed polyline itself is diagonal by design; its taps
  are square), `--extent all`, `--list`, `--template`, and the four schematics all run and parse.
* **Not verified**: opening the files in draw.io itself — no installation here. The new construct is
  the standard `exitX/entryX` pair plus a visible 3 × 3 dot, but these are the first sheets to use
  them.

---

## 8. Open data issues (user's workbook)

These are reported by the tool but must be fixed at source:

| Issue | Detail | Effect |
|---|---|---|
| Out-of-range coordinate | `SCD-68 OLD TALLUK OFFICE` lat `3.926426` (should be `13.926426`) | rejected, listed in `UNPOSITIONED` |
| Duplicate coordinates | `SCD-35`/`SCD-72` both `13.92277, 75.57755`; `SCD-48`/`SCD-106` both `13.92685, 75.57756` | one of each pair is likely a copy-paste; symbols are nudged apart |
| Missing GPS | `SCF-3`, `SCF-5` (F13); `SCF-94`, `SCF-112`, `SCF-140` (F22); `SCE-3` (F3) | listed in `UNPOSITIONED` |
| No HT coordinates | all five MRS F1 HT services | always `UNPOSITIONED` in `--geo` |
| Stale RMU register row | `RMU WAYS` row 14: `RMU-RAMANNA-SHETTY-PARK` / `VL1` says `CONNECTED = N`, but `P-001` leaves by `VL1` | the drawing follows `NODES` and warns; the cell needs `Y` |
| RMU parents missing | `RMU-KOTE-ROAD`, `RMU-GANGOTHTI-COLLEGE`, `RMU-RAMANNA-SHETTY-PARK` have no `PARENT_ID` | each RMU is attached to the nearest node and drawn dashed |
| Connections still drafts | 114 of 131 `NODES` rows are `BASIS = INFERRED` | the sheet is only as true as the parents; `--list` reports the split per feeder |
| Duplicate table | `MF13 Station Fesibility.docx` holds two copies that disagree on the DTC row | outside this tool's scope, flagged to the user |

---

## 9. Pending / recommended next steps

1. **Fix `RMU WAYS!D14`** (`VL1` of `RMU-RAMANNA-SHETTY-PARK` → `Y`) and **give the three RMUs a
   `PARENT_ID`**. Both are single cells in `field_data/TOPOLOGY.xlsx`; the dashed links and the
   warning disappear on the next run. This is the cheapest accuracy gain available.
2. **Correct `PARENT_ID` on the DTC rows.** 114 of 131 rows are still `BASIS = INFERRED`, so most of
   what the sheet now draws faithfully is a first draft. Work feeder by feeder against the hand SLD;
   `--list` shows how much of each is still inferred.
3. **Poles.** Only `P-001` exists. A pole row per branch point is what makes a spur leave the trunk
   at the right place — and it is the input the tool cannot infer.
4. **`BENDS`** for the runs that visibly cut across a block: two or three bend points per lateral.
5. **Record `NORMALLY` on the GOS rows.** `GOS-01` is blank, so every F13 run says it was drawn
   closed as an assumption. One cell (`NODES!L131`, `NC` or `NO`) removes the warning and, if `NO`,
   makes the drawn break real.
6. **Fix `SCD-68`** (`3.926426` → `13.926426`) and the duplicate coordinate pairs at source; MRS F1
   gains one positioned DTC and loses one `UNPOSITIONED` line.
7. **HT coordinates** if HT services should appear on the geographical sheet; without a `LAT`/`LONG`
   they are drawn on the schematic only, after the trunk.
8. **A per-spur count in the summary table**, from the same rows as the spur lines — the last place
   where the sheet still says "feeder" where the reader wants "lateral".
9. **Nudge the last few labels on MRS F1** if the sheet is going to print: 20 connectors still touch
   text (10 links, 10 position leaders). The router avoids labels where a corridor exists; the rest
   is the density of that cluster, and moving the label in draw.io is cheaper than another search.

---

## 10. Decisions and why

* **draw.io over Word/Visio/AutoCAD** — real connectors that stay attached, free, A3 PDF export,
  shape library, and it opens the same files in VS Code (already installed). AutoCAD only wins if
  MESCOM ever demands DWG; QGIS only if the network is already traced there.
* **Excel as the single source of truth** — the master already held code/name/kVA/feeder; deriving
  every label and the summary table from it removes the divergence between drawing and register.
* **Draw the workbook, not a guess from the coordinates** — the reported defect was exactly the
  failure mode the requirements document predicted: an inferred tree put a DTC on the wrong parent
  and left recorded equipment off the sheet entirely. When rows exist, they win.
* **`NODES` outranks the `RMU WAYS` register.** A child parented to an RMU is drawn on the way it
  names and the register row is reported, because the connections are what a sheet is signed against
  and the register is a separate record.
* **Errors stop, warnings do not.** A drawing that contradicts itself is not produced at all
  (duplicate `NODE_ID`, missing parent row, a loop, a bad `WAY`, two cables on one terminal);
  anything unfinished still draws, with the guesses dashed and listed. An unfinished workbook must
  remain usable.
* **The master keeps owning the register facts.** A DTC's name, rating and GPS come from the master
  even when the workbook has a row; the workbook owns the equipment and the parents. That is why a
  mismatch over 100 m between the two is reported and the master's value is used.
* **Inferred trunk rather than no geo layout** — with no workbook rows for a feeder,
  `--no-topology`, or no workbook at all, an MST over real coordinates is a defensible "somewhere
  near the actual route" and needs zero fieldwork.
* **Unpositioned block instead of interpolation** — the user explicitly chose to keep unplaced DTCs
  visible in a separate box rather than guess their position.
* **Script-relative master path** — the user keeps a working copy in the tool folder; hard-coding
  the network path would have split the source of truth.

---

## 11. Gotchas for the next session

* **Regenerating overwrites the `.drawio`.** Hand edits must be saved under a new name first.
* **Excel lock file** `~$DTC DETAILS U4.xlsx` appears whenever the workbook is open; harmless, but
  if the user is mid-edit the script reads the last *saved* state.
* **`--preview` is a re-draw, not a renderer.** It shares the generator's coordinates and now honours
  the edge ports (a floating connector is clipped at the border), hides the anchors and fills small
  dots; fonts and HTML wrapping still differ from draw.io. It is a placement check, not a proof of
  what draw.io will show.
* **Label density is a real limit.** MRS F1 at ~46 px between symbols with 100 px labels cannot be
  collision-free; the router leaves 20 crossings there (10 links, 10 leaders) and avoids labels where
  a clean corridor exists. Do not "fix" it by shrinking the page or symbols without checking all
  four PNGs, and do not move a label in code for one sheet — another feeder's run may use that spot.
* **A port is a style fraction, not a point.** `Page.edge(..., src_side=...)` writes
  `exitX/exitY`; the first waypoint must be offset along the same normal or the line kinks at the
  border. `ConnectorRouter` guarantees that with its stub — a hand-written edge must do the same.
* **A tap dot is visible on purpose.** `ST_TAP` marks the surveyed position; `ST_HIDDEN` is an anchor
  (RMU portion, route point, leader start). Swapping them prints a dot for every surveyed bend, or
  leaves a leader ending in mid-air.
* **Do not reroute a feed to dodge text.** Labels are a smaller penalty than symbols on purpose; a
  line through a label is a readability problem, a line through a transformer terminal is a wrong
  drawing. The console counts both separately.
* **Coordinate sanity** — valid values for this area are lat ≈ 13.9, lon ≈ 75.5. A lat of 3.9 means
  a dropped leading digit.
* **Never edit `field_data/TOPOLOGY.xlsx` to "fix" it.** It is field work; the tool reads it and
  reports. `--topology <copy>` is the way to try a change without touching it.
* **`--topology-template --force` only ever rebuilds `template/TOPOLOGY TEMPLATE.xlsx`.** Rebuild it
  to a temp path unless the shipped blank really needs regenerating; the workbook's READ ME names
  document paths, so a doc move means a rebuild.
* **A `--geo` sheet is only as good as `PARENT_ID`.** A wrong parent now draws a wrong line but no
  warning — `BASIS` is what tells you a row is still a draft, and nothing on the sheet shows it yet.
  Check `--list` before trusting a drawing.
* **Both layouts now read the workbook**, so `--no-topology` changes two things at once: the geo
  links become an MST *and* the schematic falls back to code order. Compare two runs with the same
  switches.
* **A node with no `PARENT_ID` is reported differently by each sheet.** The geo sheet attaches it to
  the nearest node and dashes the link; the schematic has no line to dash, so it draws the node after
  the trunk with no parent named. Neither is a fix: the cell is the fix.
* **The row bus on the schematic is not wiring.** Moving a box in draw.io does not change what feeds
  it; the `from` line under it does, and that comes from `PARENT_ID`.
* **The MUSS is excluded from the fit** (`net_codes`), on purpose. If a future change puts a `MUSS`
  row inside the load centre, the drawing will still reserve the off-map treatment for it; the
  extent rule is what decides.

---

## 12. Environment

| Item | Value |
|---|---|
| OS | Windows (win32), bash shell (Git Bash) |
| Python | 3.13.5 |
| openpyxl | 3.1.5 |
| PyMuPDF | 1.28.0 |
| Project folder | `D:\Office PC\D DRIVE\ESTIMATES\SLD TOOL\` |
| Upstream data | `D:\Office PC\D DRIVE\DTC LIST\DTC DETAILS U4.xlsx` (original, now superseded by the copy here) |
