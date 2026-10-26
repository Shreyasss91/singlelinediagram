# Changelog — MESCOM SLD tool

Every change made in the session of **26 October 2026**, newest entry on top.
This covers the whole arc: reading the two Word documents, choosing the tool, and building the
generator up to its current state.

---

## 2026-10-26 · Newest — Template folder, dated output folders, topology data model

### Added
- **`template/` folder** — all reusable templates in one place: `MESCOM_SLD_TEMPLATE.drawio`,
  `route_template.csv` and the new topology workbook. `--template` writes there by default.
- **`template/TOPOLOGY TEMPLATE.xlsx`** — the connection data the drawing has been missing, in the
  required format:
  * `NODES` — 128 rows: 119 DTCs, 5 HT services, 2 MUSS and 2 RMU, each with the coordinates the
    master holds, plus `PARENT_ID` (the node it is fed from) and `SPUR_ID` (`T` for the trunk,
    `S1`, `S2`, … for laterals).
  * `SPURS` — auto-computed lateral register with DTC count and connected kVA.
  * `BENDS` — optional shape points for spans that must follow a road.
  * `EXAMPLE` — a small complete feeder showing the required format.
  * `READ ME` and a hidden `LISTS` sheet for the dropdowns.
  `PARENT_ID` is pre-filled from a minimum spanning tree and marked `BASIS = INFERRED`: a first draft
  to correct, not surveyed fact. 112 of 119 DTCs receive an inferred parent; 14 rows are `UNKNOWN`
  (7 DTCs without usable GPS, 5 HT services, 2 RMUs).
- **`--topology-template`** — builds or refreshes that workbook from the current master.
- **`--force`** — allows `--topology-template` to overwrite a filled file. Without it the command
  refuses, so a session of data entry cannot be wiped by accident.
- **`TOPOLOGY_DATA_REQUIREMENTS.md`** — the analysis behind the model: why GPS alone is not enough,
  the failure modes of an inferred trunk, the field specification, how trunk and spur grouping are
  derived, collection tiers with effort, data sources, validation rules and the decisions still open.

### Changed
- **Generated sheets now go to `outputs\<YYYY-MM-DD>\`** instead of the tool folder, so a run today
  cannot overwrite yesterday's drawing. `--outdir` replaces the date folder; `--out` still names one
  exact file. Console output prints paths relative to the tool folder.
- `--template` writes to `template/MESCOM_SLD_TEMPLATE.drawio`; `--preview` puts the library PNG
  beside it.
- Existing generated files were moved into `outputs\2026-10-26\`, and `route_template.csv` into
  `template\`.
- README: new folder tree, dated outputs, `--outdir` / `--topology-template` / `--force`, a new §8
  on topology, renumbered later sections, and troubleshooting rows for the locked template and the
  new output location.

### Regenerated
- `template/MESCOM_SLD_TEMPLATE.drawio` (+ library PNG).
- `outputs/2026-10-26/`: all four feeders in both layouts — MRS F1, Mandli F13, Mandli F22 and, for
  the first time, Mandli F3 — plus a PNG for each. Eight output files (eight pages) plus the
  two-page template: zero dangling connectors, zero duplicate ids, zero off-page shapes.
- Mandli F3 is new: 5 DTCs, 4 positioned, `1 m = 3.20 px`. It was a pending item in the handoff.

### Notes
- The topology workbook is a **data-collection format, not yet a drawing input**. `build_geo()` still
  infers the trunk by MST and reads only `--route` for a supplied line. Wiring `NODES` into the
  layout is the next build step, once the connections are confirmed.
- The template is locked while open in Excel (`~$TOPOLOGY TEMPLATE.xlsx`); a rebuild reports
  `Permission denied` rather than writing into an open workbook.

---

## 2026-10-26 — MUSS and RMU coordinates from sheet `NON DTC NODES`

### Added
- **Reads the new `NON DTC NODES` sheet** (`NODES`, `GPS LAT`, `GPS LONG`, `GPS HYPERLINK`). Each row
  is classified by name into a **MUSS** or an **RMU** node and picked up automatically — no more
  editing `STATIONS` in the script.
- **MUSS position is now taken from the sheet.** The station text in column `J` of sheet `ALL` is
  matched to the MUSS node name (`110/11KV MRS` → `110/11KV MRS MUSS`), and that coordinate both
  roots the inferred trunk and places the substation box. `--muss "lat,lon"` still overrides it.
- **RMU symbols are drawn.** In the geographical layout each RMU sits at its surveyed position,
  attached by a stub to the nearest point on the trunk, labelled with its sheet name. In the
  schematic layout the RMU is appended after the HT services as a `90 × 22` box marked `RMU`.
- **Automatic feeder assignment for non-DTC nodes** — `assign_nodes_to_feeders()` gives each RMU to
  the feeder whose nearest DTC is closest, so Kote Road and Gangothti College RMU appear on the
  MRS F1 sheet and on no other.
- **`--extent cluster|all`** — `cluster` (default) fits the sheet to the DTC network; `all` includes
  the MUSS in the fitted area even when it drags the scale down.
- **Off-map MUSS handling** — when the substation is far outside the network (MRS MUSS is 2.6 km
  from its load centre, Mandli MUSS 1.7 km), the network keeps its own scale and the MUSS is drawn
  at the page edge in its true direction, joined by a **dashed** connector and labelled with the real
  distance and `not to scale`. Previously the whole sheet shrank to 0.50 px/m to include it.
- `--list` now prints the non-DTC nodes with their coordinates.

### Changed
- `place_label()` takes the symbol width and height, so labels position correctly around the wider
  RMU boxes instead of assuming a DTC footprint.
- `read_master()` returns `(dtcs, hts, nodes)`; `build_geo()` and `build_schematic()` accept an
  `rmus` argument.

### Fixed
- RMUs were being added to every feeder's sheet; they are now assigned per feeder.
- RMU boxes landed on top of DTC symbols in the dense Bazaar area; they now search 8 directions ×
  4 distances for a clear spot and keep a stub to their true position.
- Duplicate `rmus` console line printed once by the CLI and once by the layout.
- Schematic RMU labels repeated the word `RMU` under a box that already said it; now just the name.

### Regenerated
All seven sheets from the workbook, now carrying the real substation position:
MRS F1 `1 m = 0.89 px` with both RMUs placed; F13 `2.34 px/m`; F22 `1.28 px/m`.

---

## 2026-10-26 — Local workbook, output naming, full documentation

### Changed
- **The master workbook is now read from this folder.** `MASTER_DEFAULT` changed from the hard-coded
  `D:\Office PC\D DRIVE\DTC LIST\DTC DETAILS U4.xlsx` to
  `os.path.join(SCRIPT_DIR, "DTC DETAILS U4.xlsx")`, so the script works from any working
  directory and always uses the copy you maintain here. Verified by running from `ESTIMATES\`
  and from `D:\`.
- `--list` now reports GPS coverage per feeder, e.g. `MRS F1  63 DTCs  (63 with GPS)`.
- The "master not found" error now names the expected location and shows the `--master` syntax.
- Geographical drawing area narrowed from `GEO_H = 790` to `735` so the `UNPOSITIONED` box always
  has clear space beneath the last row of DTCs.

### Fixed
- **The geographical sheet was overwriting the schematic sheet.** Both used
  `<label> SLD.drawio`, so running `--geo` after the default replaced the tidy sheet. Output names
  are now `<label> SLD.drawio` and `<label> - GEO SLD.drawio`.
- DTC labels and symbols in the bottom row no longer sit inside the `UNPOSITIONED` box.

### Added
- `README.md` rewritten as a full manual: 13 sections, column-by-column mapping of the workbook,
  step-by-step explanation of both layouts, all ten CLI options, five day-to-day workflows,
  limitations, data-quality warnings, troubleshooting table and a glossary.
- `CHANGELOG.md` (this file) and `SESSION_HANDOFF.md`.

### Regenerated
All seven sheets rebuilt from the updated workbook (113 of 119 DTCs now carry GPS):
`MRS F1 OLDTOWN`, `MANDLI F13 BYEPASS`, `MANDLI F22 GANDHIBAZAAR` in both layouts, plus the
template.

---

## 2026-10-26 — Geographical layout (`--geo`)

### Added
- **`--geo`** — places every DTC at its surveyed GPS position, north up, uniform scale, with a
  metre scale bar and a north arrow. MRS F1 projects at `1 m = 0.96 px`, F13 at `2.52 px`.
- **`--muss "lat,lon"`** and the `STATIONS` dictionary — substation position; without it the trunk
  starts at one end of the network and the MUSS box is marked *(position approx.)*.
- **`--route file.csv|xlsx`** — surveyed pole coordinates (`NODE_ID, LAT, LONG`, in order along the
  line). The trunk follows the real path and each DTC attaches to the nearest point on it. This is
  the "level 3" hook: full topology can be dropped in later without changing the code.
- **Minimum spanning tree trunk** (Prim) over the DTC coordinates when no route is supplied,
  rooted at the nearest DTC to the MUSS, or at the end of the tree diameter when the MUSS is unknown.
- **Perpendicular spurs** — each symbol sits ~30 px off the line, on the side away from the network
  centre, joined by a thin stub, so the trunk never runs through a transformer symbol.
- **Outlier rejection** — any coordinate more than 5 km from the feeder median is dropped and
  reported on the console. This caught SCD-68 (`3.926426` instead of `13.926426`), which alone
  stretched MRS F1 across 1,110 km.
- **`UNPOSITIONED` block** — DTCs with no usable GPS are listed by code and name in a dashed box at
  the bottom of the sheet instead of silently vanishing.
- **Label collision avoidance** — twelve candidate positions per label (four sides × three
  distances), each scored by overlap area; the freest one wins, clamped inside the drawing area.
- **Symbol de-overlap** — DTCs sharing the same coordinate are nudged apart (up to six 26 px steps)
  so both ratings remain readable.
- **MUSS placement search** — the substation box is drawn after the DTCs and searches 5 distances ×
  8 directions for a clear spot, so it never covers a symbol.
- `route_template.csv` — format sample for `--route` (denotes example values, not real ones).

### Fixed
- `TypeError: expected str … not list` — `os.path.basename()` was called on the parsed route list
  instead of the path.
- `AttributeError: 'Rect' object has no attribute 'center'` — PyMuPDF rectangles use
  `x0 + width/2`; preview renderer corrected.
- Preview crashed on rotated styles (the GOS switch); rotation is now drawn as an angled line.
- MUSS box overlapping DTCs in three of four test sheets (initial fixed-offset attempt; replaced by
  the directional search).
- Labels stacked on top of each other and on top of symbols in the dense MRS F1 centre; replaced
  single-pass placement with free-spot scoring.
- `prim(None, …)` was called with a `None` placeholder argument; now `prim([], …)`.

---

## 2026-10-26 — Guidance on opening `.drawio` files (no file changes)

- Explained the three ways to open the output: **draw.io desktop** (offline installer),
  **VS Code + Draw.io Integration** (`hediet.vscode-drawio`, already installed on this machine), and
  **app.diagrams.net** in a browser.
- Recommended VS Code and the habit of `Save As` under a new name before hand-editing, because
  regenerating overwrites the file.

---

## 2026-10-26 — First generator release (schematic layout + template)

### Added
- **`SLD TOOL/sld_generator.py`** — reads the DTC master (sheet `ALL` for DTCs, sheet `HT` for HT
  services) and writes an editable `.drawio` single line diagram for one feeder.
- **Schematic layout** — vertical distributor, equally spaced horizontal laterals, DTCs hung below
  in code order, MUSS box top right, summary table bottom right. Grid adapts to the DTC count
  (19 DTCs → 7 × 3, 63 DTCs → 14 × 5).
- **`--template`** — writes `MESCOM_SLD_TEMPLATE.drawio` with two tabs: a blank A3 sheet (title,
  MUSS, distributor, five laterals, summary table) and a `LIBRARY` page (MUSS, DTC 25/63/100/200/250,
  HT service, GOS switch, RMU, junction, bus, usage notes).
- **`--preview`** — renders the sheet to PNG with PyMuPDF from the same coordinates, for checking
  without opening draw.io.
- **`--list`**, `--feeder`, `--name`, `--title`, `--out`, `--master`.
- **Summary table** computed from the data (`25 / 63 / 100 / 200 / 250 / ALL` + total), so the sheet
  and the register cannot disagree.
- Generated `MRS F1 OLDTOWN`, `MANDLI F13 BYEPASS` and `MANDLI F22 GANDHIBAZAAR` sheets.

### Fixed
- Summary table header showed a DTC count in a column styled like a kVA rating; the last column is
  now labelled `ALL`.
- Long DTC names (`MANDLINGESHAMMA APARTMENTS`, `TRF_KOTE_WETWELL_BACKSIDE`) overflowed their label
  boxes; names are cropped to 22 characters on the sheet, full text stays in the workbook.
- Preview ignored edge waypoints, drawing diagonal trunks instead of orthogonal ones; waypoints are
  now honoured.
- Labels were drawn with a text-box call that silently drops text when it does not fit; replaced by
  per-line drawing.

---

## 2026-10-26 — Tooling decision and data groundwork (analysis only)

- Rendered `MRS F1 SLD.pdf` and `MANDLI F13 SLD.pdf` to inspect what the drawings actually are:
  radial distribution SLDs with a MUSS, a trunk, laterals, DTC squares carrying the kVA rating,
  name and `SCD-xx` code, inline GOS/RMU switches, HT service boxes and a capacity summary table.
- Measured the incumbent approach: `MRS F1 SLD.docx` holds **333 separate Word shapes** (every DTC
  is a square + text box + unattached line), `MRSF1_MANDLIF22_SLD.docx` 338, `MANDLI F13 SLD.docx` 72.
- Compared options — draw.io (chosen), Word/PowerPoint shapes, Visio, AutoCAD/AutoCAD Electrical,
  QGIS print layout, script-generated SVG — on cost, reuse, connector behaviour and A3 export.
- Located the data: `DTC DETAILS U4.xlsx` sheet `ALL` (119 DTCs: code, name, feeder, kVA, station,
  GPS lat/long) and sheet `HT` (5 HT services with contracted demand and feeder).
- Found and reported the pre-existing data defects: SCD-68 latitude typo, duplicate coordinates on
  SCD-35/SCD-72, and missing GPS on nine DTCs.
- Examined `DTC TAGGING.xlsx` (13,862 consumer rows) and `MASTER_FILE_RR_NUMBERS_DTCwise.xlsx` as
  candidate masters; both are consumer-level, so the DTC-level `DTC DETAILS U4.xlsx` was chosen.

---

## 2026-10-26 — Session start: document inspection (no file changes)

- Extracted the text of `MF13 Station Fesibility.docx` and reported its 11-row feasibility table
  (MUSS MANDLI, 1×10 + 1×20 MVA, MF-13 Mandli Bypass, 200 A, RABBIT, CT 200/1, 85 A peak,
  HTVR 3.63 %, DTC 63→100 kVA, 5349 kVA connected load).
  Flagged that the file contains **two copies of the table** which disagree on the DTC row
  (*"Enhancement from 63KVA to 100KVA"* vs *"Proposed 1X100 KVA"*).
- Parsed `MRS F1 SLD.docx`: single line diagram of MRS F1 OLDTOWN (Unit-4, CSD-2, SMG), 110/11 kV
  MRS MUSS, feeders MF-13 Bypass and MF-1 Oldtown, 77 unique DTC codes, GOS/RMU nodes, five HT
  services. Noted the hundreds of shapes, the one embedded 5.6 KB PNG (logo, not the diagram) and
  the doubled text layer.
