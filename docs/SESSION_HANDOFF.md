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

---

## 1. Snapshot

The user is a MESCOM (Shivamogga) distribution engineer who was drawing feeder single line diagrams
by hand in Word — hundreds of loose shapes per sheet. This session replaced that with a data-driven
generator: **one spreadsheet in, an editable `.drawio` single line diagram out**, in two layouts
(schematic and geographical).

The tool works. Seven sheets were regenerated from the user's updated master at the end of the
session and all pass structural validation. MUSS and RMU positions are now read from sheet
`NON DTC NODES`, so no configuration is needed for a substation that is already recorded. The open
items are data clean-up, GOS positions, and pole-level route data.

---

## 2. What was delivered

| File | Size | Role |
|---|---|---|
| `sld_generator.py` | ~37 KB | The tool. Single file, no local packages. |
| `DTC DETAILS U4.xlsx` | 49 KB | Master workbook (copied here by the user; the script reads this copy). |
| `MESCOM_SLD_TEMPLATE.drawio` | 17 KB | Blank A3 sheet + symbol library, two tabs. |
| `route_template.csv` | <1 KB | Format sample for `--route` (values are invented). |
| `MRS F1 OLDTOWN SLD.drawio` | 89 KB | Schematic, 63 DTCs + 5 HT. |
| `MRS F1 OLDTOWN - GEO SLD.drawio` | ~86 KB | Geographical, 62 positioned + 6 unpositioned, 2 RMUs. |
| `MANDLI F13 BYEPASS SLD.drawio` | 29 KB | Schematic, 19 DTCs. |
| `MANDLI F13 BYEPASS - GEO SLD.drawio` | 27 KB | Geographical, 17 positioned + 2 unpositioned. |
| `MANDLI F22 GANDHIBAZAAR SLD.drawio` | 45 KB | Schematic, 32 DTCs. |
| `MANDLI F22 GANDHIBAZAAR - GEO SLD.drawio` | 42 KB | Geographical, 29 positioned + 3 unpositioned. |
| `*.png` | — | Previews written by `--preview` (one per sheet). |
| `README.md` | 14 KB | User-facing manual. |
| `CHANGELOG.md` | — | Every change made in the session, newest first. |
| `SESSION_HANDOFF.md` | — | This document. |

---

## 3. How to run it

```bash
cd "D:\Office PC\D DRIVE\ESTIMATES\SLD TOOL"

python sld_generator.py --list                          # feeders + GPS coverage
python sld_generator.py --feeder "Mandli F13"           # schematic
python sld_generator.py --feeder "MRS F1" --geo         # geographical
python sld_generator.py --feeder "MRS F1" --geo --preview
python sld_generator.py --template                      # rebuild template + library
```

Full regeneration of everything currently delivered:

```bash
python sld_generator.py --template --preview
for f in "MRS F1|MRS F1 OLDTOWN" "Mandli F13|MANDLI F13 BYEPASS" "Mandli F22|MANDLI F22 GANDHIBAZAAR"; do
  key="${f%%|*}"; label="${f##*|}"
  python sld_generator.py --feeder "$key" --name "$label" --preview
  python sld_generator.py --feeder "$key" --name "$label" --geo --preview
done
```

Dependencies: `openpyxl` (required), `pymupdf` (only for `--preview`). Both installed here.

---

## 4. Architecture

One file, top-to-bottom:

```
constants & styles      L1–100    page size, geometry, draw.io style strings
data helpers            L108–200  _txt, _num, _kva, crop_name, read_nodes, find_station,
                                  assign_nodes_to_feeders, read_master, pick
diagram model           L180–242  Page (vertex/edge emitters), Workbook (multi-page → XML)
sheet furniture         L244–336  sheet_header, summary_table, unpositioned_block,
                                  scale_bar, _overlaps, _clamp_label, place_label
schematic layout        L339–397  choose_grid, build_schematic
geo helpers             L400–498  clean_points, _metres, prim, tree_ends, read_route, _project
geo layout              L501–706  build_geo
template                L709–766  build_template
preview                 L769–831  preview (PyMuPDF re-draw of the same coordinates)
CLI                     L834+      main
```

(Line numbers are as of the handoff; they shift with edits.)

### Data flow

```
DTC DETAILS U4.xlsx ──read_master──► DTC dicts   {code, name, kva, feeder, station, kind, lat, lon}
                       ──read_nodes───► node dicts  {name, kind: MUSS|RMU, lat, lon}
                                       (HT rows: no lat/lon)
                                            │
                              pick(feeder)  │
                    ┌───────────────────────┴───────────────────────┐
                    ▼                                               ▼
        build_schematic(...)                                 build_geo(...)
        grid of laterals                                     project + MST/route
                    └───────────────────────┬───────────────────────┘
                                            ▼
                                   Workbook.to_xml()  →  <label> SLD.drawio
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
`(1370, 100)`, summary table bottom right, `UNPOSITIONED`-free (every DTC is drawn).

Grid: `choose_grid(n)` → `per_row = clamp(ceil(n/3), 4, 14)` reduced until the slot width is at
least 92 px; `rows = ceil(n / per_row)`; row height up to 230 px. 19 DTCs → 7 × 3; 63 → 14 × 5.

Each DTC: junction dot on the lateral → 30 px stub → 44 × 22 square with the kVA inside → label
(name over code) 22 px below.

### Geographical (`build_geo`, `--geo`)

1. `clean_points` splits usable / unusable coordinates and rejects anything >5 km from the median.
2. `_metres` performs a local equirectangular projection (111.32 km/° lon, 110.54 km/° lat).
3. Trunk: `prim` builds a minimum spanning tree over the metres; root is the DTC nearest the MUSS,
   else an end of the tree diameter (`tree_ends`). With `--route`, the supplied polyline is drawn
   instead and every DTC attaches by `_project` onto its nearest segment.
4. Page transform: one uniform scale = `min(1460/spanx, 735/spany, 3.2)`; north up; centred.
5. Symbols sit 30 px perpendicular to the feeding segment, away from the centroid; stubs join them.
6. `place_label` tries 12 positions per label and scores overlap area against everything already
   placed (symbols and labels), clamping to the drawing area.
7. `scale_bar`, north arrow, `unpositioned_block`, `summary_table` finish the sheet.
8. The MUSS box is drawn after the DTCs and searches 5 distances × 8 directions for clear space.
   If including the substation would cost more than 30 % of the scale, the network keeps its own
   scale and the box is drawn at the page edge in its true direction, joined by a dashed connector
   and labelled with the real distance (`--extent all` forces the strict fit instead).
9. Each RMU belonging to the feeder is drawn as a `90 × 24` box at its surveyed position, stubbed
   to the nearest point on the trunk, searching 8 directions × 4 distances for a clear spot.

Key constants: `GEO_X, GEO_Y, GEO_W, GEO_H = 120, 205, 1460, 735`; `GEO_MAX_SCALE = 3.2 px/m`;
`DTC_W, DTC_H = 44, 22`; `LBL_W, LBL_H = 100, 26`.

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
`Mandli F1`.

---

## 7. Verification performed

* **Structure** of all seven `.drawio` files — eight diagram pages, since the template holds two
  tabs — checked with Python + ElementTree: no dangling connectors, no duplicate cell ids, every
  shape inside the A3 page.
  Latest run: 7 files / 8 pages, 0 dangling, 0 off-page, largest 369 cells / 141 edges
  (MRS F1 schematic).
* **Visual**: every sheet rendered to PNG and inspected — schematic (F1, F13, F22), geographical
  (F1, F13, F22), template library page, plus a `--muss` test and a `--route` test. Several layout
  defects were found this way and fixed (see CHANGELOG): MUSS box covering DTCs, stacked labels,
  unpositioned box colliding with the bottom row, truncated long names.
* **Portability**: `--list` and a generation run executed from `ESTIMATES\` and from `D:\` to prove
  the script-relative master path.
* **Not verified**: opening the files in draw.io itself — no draw.io installation was available.
  The XML follows the documented mxGraph schema and labels are correctly escaped, but the first
  human open is still the real acceptance test.

---

## 8. Open data issues (user's workbook)

These are reported by the tool but must be fixed at source:

| Issue | Detail | Effect |
|---|---|---|
| Out-of-range coordinate | `SCD-68 OLD TALLUK OFFICE` lat `3.926426` (should be `13.926426`) | rejected, listed in `UNPOSITIONED` |
| Duplicate coordinates | `SCD-35`/`SCD-72` both `13.92277, 75.57755`; `SCD-48`/`SCD-106` both `13.92685, 75.57756` | one of each pair is likely a copy-paste; symbols are nudged apart |
| Missing GPS | `SCF-3`, `SCF-5` (F13); `SCF-94`, `SCF-112`, `SCF-140` (F22); `SCE-3` (F3) | listed in `UNPOSITIONED` |
| No HT coordinates | all five MRS F1 HT services | always `UNPOSITIONED` in `--geo` |
| GOS not recorded | gang operated switches appear in the hand drawings but not in the workbook | manual only; no automatic placement |
| Duplicate table | `MF13 Station Fesibility.docx` holds two copies that disagree on the DTC row | outside this tool's scope, flagged to the user |

---

## 9. Pending / recommended next steps

1. **Verify the MUSS/RMU coordinates on a map once.** They are trusted as entered and are not
   outlier-checked. MRS MUSS sits 2.6 km from its load centre and Mandli MUSS 1.7 km, which is why
   both sheets use the off-map substation treatment — worth confirming those two points are right.
2. **Pole-level route data** — export pole coordinates (GIS / field app / KMZ→CSV) per feeder and
   pass `--route`. The code path is built and tested with a synthetic file; it needs real data.
3. **Topology beyond a chain** — a `PARENT` column in the route file would let real laterals branch
   instead of relying on the nearest-segment attachment.
4. **Mandli F3 sheet** — only 5 DTCs, not yet delivered as a sheet; one command away.
5. **GOS / RMU placement** — symbols exist in the library only. If their positions are ever
   recorded, add them as rows in the route file.
6. **Fix SCD-68 and the duplicate pairs**, then regenerate: MRS F1 would gain one positioned DTC.

---

## 10. Decisions and why

* **draw.io over Word/Visio/AutoCAD** — real connectors that stay attached, free, A3 PDF export,
  shape library, and it opens the same files in VS Code (already installed). AutoCAD only wins if
  MESCOM ever demands DWG; QGIS only if the network is already traced there.
* **Excel as the single source of truth** — the master already held code/name/kVA/feeder; deriving
  every label and the summary table from it removes the divergence between drawing and register.
* **Inferred trunk rather than no geo layout** — with no topology data, an MST over real
  coordinates is a defensible "somewhere near the actual route" (the user's words) and needs zero
  fieldwork. `--route` upgrades it later without a rewrite.
* **Unpositioned block instead of interpolation** — the user explicitly chose to keep unplaced DTCs
  visible in a separate box rather than guess their position.
* **Script-relative master path** — the user keeps a working copy in the tool folder; hard-coding
  the network path would have split the source of truth.

---

## 11. Gotchas for the next session

* **Regenerating overwrites the `.drawio`.** Hand edits must be saved under a new name first.
* **Excel lock file** `~$DTC DETAILS U4.xlsx` appears whenever the workbook is open; harmless, but
  if the user is mid-edit the script reads the last *saved* state.
* **`--preview` is a re-draw, not a renderer.** It shares the generator's coordinates, so it is
  faithful for layout, but fonts and HTML wrapping differ slightly from draw.io.
* **Label density is a real limit.** MRS F1 at ~46 px between symbols with 100 px labels cannot be
  collision-free; do not "fix" this by shrinking the page or the symbols without checking the PNG.
* **Coordinate sanity** — valid values for this area are lat ≈ 13.9, lon ≈ 75.5. A lat of 3.9 means
  a dropped leading digit.

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
