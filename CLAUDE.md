# CLAUDE.md

Guidance for Claude when working in this repository. Read this before changing code or data.

## What this is

A single-file Python tool that turns the DTC master workbook into editable `.drawio` single line
diagrams for MESCOM Unit-4 / CSD-2 feeders. Data lives in the spreadsheet; the drawing is derived
from it, so the sheet and the register cannot disagree.

The domain matters more than the code here: a wrong parent link or a wrongly-closed switch is a
safety-relevant error on a distribution drawing, not a cosmetic one.

## Commands

```powershell
python sld_generator.py --list                                          # feeders, DTC/GPS/connection coverage
python sld_generator.py --feeder "MRS F1" --name "MRS F1"               # schematic sheet
python sld_generator.py --feeder "MRS F1" --name "MRS F1" --geo         # geographical sheet, drawn from the workbook
python sld_generator.py --feeder "MRS F1" --name "MRS F1" --geo --preview
python sld_generator.py --feeder "MRS F1" --geo --no-topology           # back to the inferred trunk
python sld_generator.py --feeder "MRS F1" --geo --topology other.xlsx
python sld_generator.py --feeder "MRS F1" --route template\route_template.csv
python sld_generator.py --template                                      # rewrite the blank drawio template
python sld_generator.py --topology-template --out "%TEMP%\tt.xlsx"      # rebuild the topology workbook, not in place
python sld_generator.py --topology-template --force                     # ... over the shipped blank
```

Output goes to `outputs\<YYYY-MM-DD>\`, one folder per run day, `--outdir` / `--out` to override.
`outputs/` is git-ignored — never commit generated sheets.

Dependencies: `openpyxl` (required), `pymupdf` (`--preview` only). No pinning, no build step, no
test suite.

## Layout

```
sld_generator.py                 the whole tool (~1600 lines, no packages)
field_data/DTC DETAILS U4.xlsx   the master workbook - the register of record
field_data/TOPOLOGY.xlsx         your filled-in connection workbook - read by --geo, never written by the tool
docs/README.md                   the user manual
docs/CHANGELOG.md                newest entry on top, dated
docs/SESSION_HANDOFF.md          state of play, updated at the end of a work session
docs/TOPOLOGY_DATA_REQUIREMENTS.md   field specification for the connection data
CLAUDE.md                        this file
template/MESCOM_SLD_TEMPLATE.drawio   blank A3 sheet + symbol library (2 pages)
template/TOPOLOGY TEMPLATE.xlsx       the blank collection workbook - never fill this one
template/route_template.csv           format sample for --route
outputs/<date>/                  generated .drawio + PNG (ignored)
```

## How the code is organised

`sld_generator.py` is deliberately one file, top to bottom:

1. **Constants** — column letters for the master sheets (`SHEET_ALL`, `SHEET_HT`, `SHEET_NODES`),
   page geometry (`PAGE_W, PAGE_H = 1654, 1169`, A3 landscape in draw.io units), layout constants,
   and one `ST_*` string per draw.io style. Add a style here, never inline one.
2. **Readers** — `read_master()` pulls DTCs/HT/non-DTC nodes from the master by hard-coded column
   letter; `read_route()` reads the optional pole CSV/XLSX; `assign_nodes_to_feeders()` gives each
   RMU/MUSS its feeder from the nearest DTC. The connection workbook has its own reader set:
   `read_topology()` (NODES / RMU WAYS / BENDS / SPURS, keeping the row number for messages),
   `topology_for_feeder()` (the subtree this feeder draws), `validate_topology()` (structural checks),
   `topology_items()` (merge with the register) and `resolve_links()` (link per node, with the reason
   when it is a guess).
3. **Model** — `Page` and `Workbook` emit the `mxCell` XML. `Workbook.to_xml()` is the only writer.
4. **Furniture** — `sheet_header`, `summary_table`, `unpositioned_block`, `scale_bar`, label
   placement helpers. Shared by both layouts.
5. **Layouts** — `build_schematic()` (a grid, but ordered by `tree_order()` when the workbook is
   there) and `build_geo()` (projection, then either the workbook's parent links or an MST or
   `--route`). Both return `(Workbook, info)`; `info` feeds the console summary. `build_geo()` order
   matters: page transform, then one tap point per node, then the symbols (which may move off a
   crowded spot — `draw_rmu()` for a five-way unit and its terminals), then the MUSS box, then the
   lines — so a line meets the position the survey recorded, not the box that had to be moved.
   `build_schematic()` draws the node's parent under its box (`from SCD-01`) and the spur tag at the
   branch point, because a grid has no line to say either; the row bus is a convention, not a claim.
6. **Template builders** — `build_template()` writes the drawio template; `build_topology_template()`
   writes the workbook, with `_infer_topology()` (Prim's MST rooted at the MUSS, then the longest path
   from the MUSS becomes `SPUR_ID = T` and each branch off it becomes `S1`, `S2`, …) and `_node_id()`.
7. **`preview()`** — re-draws the same XML with PyMuPDF into a PNG. It is an approximation of the
   sheet, not a render of draw.io: use it to check placement and density, not fonts or styling.
8. **`main()`** — argparse dispatch. `--topology-template` and `--template` return early; everything
   below them is the normal generate path.

## Invariants — do not break these

**Master column mapping.** `read_master()` indexes by letter: `ALL` uses B/C/E/F/I/J/AD/AE, `HT`
uses E/F/H/L, `NON DTC NODES` uses A-D. If the master moves a column, update the reader *and* the
tables in `docs/README.md` §3 and the requirements doc. MUSS/RMU type is detected from the name containing
`MUSS`/`RMU`; a MUSS is matched to `ALL!J` station text.

**One parent per node.** Every topology row has exactly one `PARENT_ID`, pointing upstream towards
the MUSS; only a MUSS row is parentless. Connections live in `NODES` and nowhere else. `RMU WAYS` is
a register of terminals, not a second copy of the tree.

**RMUs are five-way units, and not all five ways are cabled.** `RMU_WAYS = ODin (IN), ODout, VL1,
VL2, VL3 (OUT)`. The RMU is one `NODES` row whose own `PARENT_ID` *is* way `ODin`; children name
their exit way in `NODES!WAY` (`ODout`/`VL1`–`VL3`, never `ODin`). Only `ODin` is always connected;
any outgoing way, `ODout` included, may be empty — `CONNECTED = N` with a blank `NORMALLY`.
`_way_cabled()` is the template default for a freshly built workbook (only `ODin` starts `Y`); the
drawing decides separately in `draw_rmu()` — filled when a child names the way or the register says
`CONNECTED = Y`, hollow and labelled `empty` otherwise. Both rules exist because the template holds
the register and the drawing holds the connections.

**NODES is the tree, RMU WAYS is a register.** When the two disagree the drawing follows `NODES` —
a child parented to an RMU is drawn on the way it names — and the register row is reported. Do not
invert this: the connections are the record a sheet is signed against.

**Errors stop, warnings do not.** `validate_topology()` returns both: a duplicate `NODE_ID`, a
`PARENT_ID` that is not a row, a loop, a bad `WAY` and two cables on one terminal are errors and
raise `SystemExit` from `main()`; everything else is printed and the sheet is still produced. Keep
that split — an unfinished workbook must still draw, and a contradiction must not.

**One source per fact.** Positions and ratings come from the master for a DTC/HT (the register);
everything else — poles, GOS, RMU, parents, ways, bends — comes from the connection workbook. A DTC
that only the master holds is still drawn, with a guessed parent, and listed.

**Workbook sheets and ranges.** `NODES` is A-O; `RMU WAYS` is A-G with `CONNECTED` in `D`; `LISTS`
feeds the dropdowns (`WAY_ALL` E, `WAY_CHILD` F, `DIRECTION` G, `YES_NO` H). Changing a column means
updating `*_HEADERS`, the row builder, the validation ranges, the auto-filter, the READ ME sheet and
the docs together.

**Coordinates** are decimal degrees, never DMS. Geo layout drops points >5 km from the feeder median
and lists them in `UNPOSITIONED` — that check is why a typo cannot wreck a sheet.

## Conventions

- Keep the single-file design and the existing section-comment banners. Match the surrounding style:
  4-space indent, `f"..."` messages, short helper functions, no classes beyond `Page`/`Workbook`.
- **Lines are routed in the tool, not by draw.io.** A run turns at right angles (`ortho_route`),
  and the corner is one point: the two segments meet at exactly 90°, with no rounding and no
  chamfer. That keeps the `--preview` PNG honest — the two renderers can only agree while the shape
  is the tool's own invention. Equipment that sits in the run (pole, GOS, RMU) is drawn on the line;
  a DTC/HT hangs 26 px beside it on a service tap, which is what keeps a routed line out of a symbol.
- Console output is aligned `key : value` lines — keep that shape, users read it.
- The script is deliberately forgiving with data: skip and report, never crash on one bad row. It
  stops only for a configuration problem — missing master, unknown feeder, unusable route, or an
  existing topology template without `--force`.
- Comments explain *why* (a rule of the network or of draw.io), not what the line does. The existing
  density is the target.
- Drawings are schematic, not surveys; do not add precision the data does not have.

## Verifying a change

There is no test suite, so verify with the tool itself and say what you ran.

```powershell
python sld_generator.py --list                                     # data still parses
python sld_generator.py --feeder "MRS F1" --geo --preview --outdir "$env:TEMP\sldcheck"
python sld_generator.py --feeder "Mandli F13" --geo --no-topology  # the MST path still works
python sld_generator.py --feeder "Mandli F13"                      # schematic path too
python sld_generator.py --feeder "MRS F1" --geo --route template\route_template.csv
```

Then:

- open the PNG (or read it) and check placement, labels, the summary table and the `UNPOSITIONED` box;
- check the drawing against the workbook: every row of the feeder that has coordinates should have a
  symbol, and each line should join a node to its `PARENT_ID`;
- confirm the `.drawio` parses:
  `python -c "import xml.etree.ElementTree as ET; ET.parse(r'...\file.drawio'); print('ok')"`
- for workbook changes, re-read it with openpyxl and assert the shape you intended: headers,
  row/column counts, validation ranges, conditional formatting, filter range;
- for topology changes, also check: no duplicate `NODE_ID`, no dangling `PARENT_ID`, no cycles,
  every RMU has exactly five `RMU WAYS` rows, and `CONNECTED` agrees with `NODES!WAY`.

Verification scripts are throwaway — put them in the temp folder, not the repo.

## Gotchas

- **Excel locks its files.** `--topology-template --force` fails with `PermissionError: [Errno 13]`
  (or a `~$` lock file) while the workbook is open. Ask the user to close it rather than working
  around it. Same for the master if it is open with unsaved edits — changes are then ignored.
- **`--topology-template` refuses to overwrite** without `--force`, deliberately. The blank template
  belongs in `template/`; the file you fill in is `field_data/TOPOLOGY.xlsx`, so `--force` should
  never touch field work. If someone has been editing the template in place, say so before rebuilding.
- **The connection workbook is read, never written.** `--topology-template --force` rewrites only
  `template/TOPOLOGY TEMPLATE.xlsx`; `field_data/TOPOLOGY.xlsx` is field work. If a stray cell there
  is wrong, say so — do not correct it for the user.
- **`field_data/TOPOLOGY.xlsx` has a stale register row.** `RMU WAYS!D14`
  (`RMU-RAMANNA-SHETTY-PARK` / `VL1`) says `N` while `P-001` leaves by `VL1`, so every F13 run prints
  the warning. That is the workbook, not the code. Compare with the file's own `git diff` before
  blaming the reader.
- **An RMU is drawn as its five ways, and only the terminals know.** `draw_rmu()` labels each
  terminal from `NODES!WAY` and `RMU WAYS!CONNECTED`; nothing else names the ways. A `--no-topology`
  run has neither, so it draws the plain `5-way` box on purpose — do not "fix" that into five hollow
  terminals, which would state that a unit nobody has surveyed is unused.
- **Line endings.** Git reports `LF will be replaced by CRLF` on this Windows checkout. Expected;
  leave it alone.
- **This repository is public.** The master already contains real DTC names, meter numbers, GPS and
  consumer references; that was the user's explicit decision. Do not add new personal data (consumer
  names, phone numbers, photos) or new workbooks without asking.
- **Do not "fix" data quirks** unless asked: `SCD-68`'s latitude is `3.926…` instead of `13.926…`,
  several DTC pairs share coordinates, and a handful have no GPS. They are documented in `docs/README.md` §12
  as things to watch, not bugs to correct silently.

## Docs and commits

A change is not finished until the documents agree with the code. Depending on the change:

| Change | Also update |
|---|---|
| CLI option, layout, output name | `docs/README.md` (options table, relevant section) |
| Master column / sheet | `docs/README.md` §3, `docs/TOPOLOGY_DATA_REQUIREMENTS.md`, workbook READ ME |
| Topology schema, validation | `docs/TOPOLOGY_DATA_REQUIREMENTS.md` (§ field spec, § validation rules, worked example), `docs/README.md` §8, regenerated `template/TOPOLOGY TEMPLATE.xlsx` |
| Any user-visible change | `docs/CHANGELOG.md` — new entry at the top, `## <date> · Newest — <title>`, plain-language why first |
| End of session | `docs/SESSION_HANDOFF.md` update |

Commit messages use seven fixed sections, in this order, filled with substance — no one-line
bodies:

```
WHAT        files and behaviours changed, specific
WHY         the modelling or engineering reason it had to change
HOW         how it is represented / how the pieces fit
NEED        why the format or code needs it now
USAGE       commands and how to fill things in
BENEFITS    what becomes checkable or impossible to get wrong
REQUIREMENTS new constraints, dependencies, limits that remain
```

Write them to a file and use `git commit -F` (PowerShell has no heredoc). Commit each self-contained
step separately. **Do not push unless the user asks.**

## Current state

Done: five-way RMUs (`NODES!WAY`, `RMU WAYS`, empty ways via `CONNECTED`) drawn as five terminals,
two layouts, `--route`, both templates, docs, and **both layouts drawn from the connection
workbook** — `--geo` with parent links, poles, the GOS and its state, the RMU's terminals,
orthogonal runs with square corners, `BENDS` shaping a span and spur tags from `SPURS!LABEL`; the
schematic in depth-first order with `from <parent>` under each box, equipment symbols and a spur tag
at the branch point. Validation separates errors from warnings. Known limitation list is
`docs/README.md` §11 — keep it honest as work lands. Next candidates, in the order the data will pay
for itself: a `PARENT_ID` on each RMU row (three still guessed), a recorded `NORMALLY` on the GOS
rows, poles for the remaining feeders, and a per-spur count in the summary table.
