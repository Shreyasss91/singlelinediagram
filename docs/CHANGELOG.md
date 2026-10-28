# Changelog — MESCOM SLD tool

Every change made in the session of **26–27 October 2026**, newest entry on top.
This covers the whole arc: reading the two Word documents, choosing the tool, and building the
generator up to its current state.

---

## 2026-10-28 · Newest — the schematic is the same tree as the geographical sheet

The schematic still has no geography, but it no longer has a different tree: it is now drawn in the
order the connection workbook records, each box says what feeds it, and a spur is tagged where it
leaves the trunk. The two sheets of a feeder can be read against each other.

### Added
- **`tree_order()`** — depth first from the MUSS in workbook row order. A parent is drawn before its
  children and a spur's nodes stay in one run of the grid, which is what the geographical sheet shows
  with lines.
- **The feed is named under each box** — `from SCD-01` in 7 px blue. It is the fact the schematic
  otherwise lacks: the bus under a row says *which lateral*, never *which node*.
- **`rmu_way_text()`** — a grid row has no room for a terminal cluster, so the unit's ways are said in
  words (`5-way - out VL1`, or `5-way - no way out recorded`), the same ways the geographical sheet
  draws as terminals.
- **A `POLE` and a `GOS` are drawn** — a dot with its code, a bar with its state, an `(open)` after
  the code. Before, only DTC, HT and RMU had symbols.
- **A spur tag at its branch point** — `S1` plus the `SPURS!LABEL` name under the node the spur leaves
  from. A tag at the first box of the spur would say the spur starts there, which is not what the
  register means.

### Changed
- **The order is the workbook's, not the register's code order.** `--no-topology` (or no topology
  file) keeps code order, because no parent is recorded there.
- **Nodes with nothing to hang under are reported once per reason**, not once per node — a 63-DTC
  feeder with eight such rows printed eight lines for one fact:
  ```
  ! 8 node(s) drawn after the trunk, no parent named - no PARENT_ID recorded:
    RMU-KOTE-ROAD, RMU-GANGOTHTI-COLLEGE, SCD-68, HT-28, HT-87, HT-92, HT-188, HT-193
  ```
- The subtitle says `connection workbook` instead of `generated from DTC DETAILS U4.xlsx`, and the
  console gains a `source` line on the schematic path as well.

### Regenerated
- Regenerated after this change: `outputs\2026-10-27\` (the four schematic sheets and their PNGs).

### Notes
- The row bus is still a bus: a drawing convention standing in for "these hang off the lateral",
  not a claim about which node feeds which. That claim is on the box.
- The summary table is unchanged — it is still the kVA histogram of the master's DTCs.

---

## 2026-10-27 · an RMU is drawn as its five ways

A five-way unit was one box with its cabled ways named under it. It is now a box with **a terminal
per way**, because "which way does this cable leave by" is the question an RMU exists to answer, and
it could not be read off the sheet.

### Added
- **`draw_rmu()`** — the unit as a `90 × 24` box plus five terminals: `ODin` on the side the feed
  arrives from, the four outgoing ways (`ODout`, `VL1`–`VL3`) in register order down the other side,
  each a 7 × 7 square on a short leader out of the box and labelled with its way.
- **The state of a way is visible.** A terminal is **filled** when a cable lands on it (a child names
  the way in `NODES!WAY`, or the register says `CONNECTED = Y`) and the label carries the switch
  state where one is recorded (`VL1 NC`). A way with nothing on it is drawn **hollow** and labelled
  `empty`, so a spare way cannot be read as a row missing from the workbook.
- **`ST_TERM`, `ST_TERM_EMPTY`, `ST_TINY_L`, `ST_TINY_R`** — styles for the terminals and their
  left/right-aligned 6 px labels, and `RMU_W, RMU_H, RMU_TERM, RMU_LEAD, RMU_ROW, RMU_LBL,
  RMU_CLUSTER` for the cluster's geometry.

### Changed
- **A cable off an RMU starts at its terminal**, not at the box: `build_geo()` looks the way up in
  `rmu_terms` and runs the line from that terminal's centre, so the sheet shows the way a run leaves
  by. The unit's own feed still arrives at its surveyed point and taps into the box.
- **The whole cluster is reserved** in the label-collision map, so a DTC label cannot land on a
  terminal and the unit's own name is placed outside it.
- **`--no-topology` keeps the plain box** marked `5-way`: with no register and no children, five
  hollow terminals would claim the unit is unused, which is not known.
- The unit's label no longer carries `5-way - out VL1`; the terminals say it. README §6 and §11 follow.

### Regenerated
- Regenerated after this change: `outputs\2026-10-27\` (all four geographical sheets).

### Notes
- On `MRS F1` neither RMU has a recorded child, so all four outgoing terminals are hollow — that is
  the data, not a defect: the register has `CONNECTED = N` on every way of both units.
- On `Mandli F13` the `VL1` terminal is filled and `P-001`'s cable leaves from it, while the register
  row for that way still says `N` — the warning stands.

---

## 2026-10-27 — a GOS shows whether it is open

`NODES!NORMALLY` now decides how the switch is drawn, and the run says what stands behind an open
point.

### Added
- **`ST_GOS_OPEN`** — a hollow blade, so the state survives a black-and-white print where colour
  would not.
- **An open point is drawn as one:** the blade is hinged on the line at the switch's own coordinate
  and lifted at 42°, with a small dot at the hinge so the break is obvious. A normally closed switch
  keeps the solid bar it had.
- **`load_below(code)` in `validate_topology()`** — walks the selection's children and reports the DTC
  count and connected kVA under a node, so an open point comes with its consequences:
  ```
  ! NODES!131 GOS-01: NORMALLY = NO - drawn as an open point - 5 DTC(s) / 800 kVA recorded
    downstream; while it stands open that load is fed from elsewhere
  ```

### Changed
- **A blank `NORMALLY` on a `GOS` is now reported** (`drawn closed; record NC or NO`) instead of being
  passed over. `GOS-01` in the current workbook is blank, so every Mandli F13 run says so.
- **The label carries the state** — `GOS-01 (open)` — because a rotated blade is easy to miss on a
  dense sheet.
- `docs/TOPOLOGY_DATA_REQUIREMENTS.md` §12 moves *GOS open/closed* from the open list to the built
  table, and §13 gains the two switch-state rows. README §6 and the glossary follow.

### Regenerated
- Regenerated after this change: `outputs\2026-10-27\` (all four geographical sheets).

### Notes
- The state is read from `NODES!NORMALLY`, never guessed: a switch with no recorded state is drawn
  closed and reported, because closing is the assumption a radial feeder is operated under.
- `RMU WAYS!NORMALLY` still only warns (the load on a way marked `NO` is named); a way's state is not
  drawn yet — it belongs with the five terminals.

---

## 2026-10-27 — runs routed at right angles, corners rounded

The geographical sheet stopped drawing diagonals between taps. Requested after the first workbook-
driven sheets: *"it need not always be a direct straight line, you can make use of horizontal,
vertical lines, curve at edge so that the diagram looks pretty."*

### Added
- **`ortho_route(a, b, radius, min_run)`** — waypoints for a run that leaves on the axis with the
  larger gap, turns at the midpoint and arrives square. Under `min_run` (70 px) it returns nothing, so
  a short tap stays a straight line instead of growing a meaningless little jog.
- **`_chamfer(pts, radius)`** — cuts the corner of a right-angle polyline back into a short diagonal.
  draw.io does not round the corners of an explicit point list, so the curve has to be in the points:
  it shows in draw.io *and* in the `--preview` PNG, which the two can only agree on if the shape is
  the tool's rather than the renderer's.

### Changed
- **Every node-to-node run, and the edge from the MUSS box**, is now routed by `ortho_route`. A span
  with `BENDS` rows keeps its recorded shape, with its corners chamfered at 10 px.
- **A DTC or HT service hangs beside the line, not on it.** Symbols are placed 26 px off the run
  along the perpendicular from the parent and joined by a service tap; poles, the GOS and the RMU
  stay on the line, because they sit in the run. This is what stops an orthogonal route from crossing
  a transformer box — the failure `docs/TOPOLOGY_DATA_REQUIREMENTS.md` §2 lists as "spur through a
  symbol".

### Fixed
- **A `BENDS` row on a span fed from an off-map MUSS dragged the line past the paper.** Bends between
  the substation and the network are now ignored: that span is labelled `not to scale`, so its shape
  is not drawn from coordinates 1.7 km outside the fit. Found by testing bends on exactly that span.

### Regenerated
- Regenerated after this change: `outputs\2026-10-27\` (all four geographical sheets).

### Notes
- Routing is a drawing convention, not a survey: the run turns where it is convenient to read, which
  is why the sheet still carries a scale bar and the `BENDS` sheet remains the way to record the real
  road. A link shorter than 70 px is left straight on purpose.
- The schematic layout is untouched — it is orthogonal by construction.

---

## 2026-10-27 — the geographical sheet is drawn from the connection workbook

The `--geo` layout stops inferring the tree. It now reads `field_data/TOPOLOGY.xlsx` and draws what
the record says: parent links, poles, the GOS, the RMU and its ways.

This came from a report on `--feeder "Mandli F13" --geo`: **the RMU, `GOS-01` and `P-001` were not
on the sheet at all**, and `SCD-10` — whose `PARENT_ID` is `SCD-36` — appeared to be fed from both
`SCD-36` and `SCD-11`. Both symptoms had the same cause: `build_geo()` did not read the workbook, so
it could only draw the 19 DTCs the master holds and had to guess the links between them.

### Added
- **`read_topology(path)`** — reads `NODES`, `RMU WAYS`, `BENDS` and `SPURS` from the connection
  workbook, with the row number kept so every message can name it.
- **`topology_for_feeder(topo, feeder, station)`** — the rows one feeder draws: its own rows, the
  MUSS above them, and any row whose `FEEDER` cell is blank but hangs under one of them. Taking the
  subtree is what stops a feeder from picking up another one's network.
- **`validate_topology(sel, topo, master_codes)`** — the checks from
  `docs/TOPOLOGY_DATA_REQUIREMENTS.md` §13. Structural errors (duplicate `NODE_ID`, a `PARENT_ID`
  that is not a row, a loop, a bad `WAY`, two cables on one RMU terminal) stop the run and name the
  row; everything else is printed and the sheet is still drawn.
- **`topology_items(...)`** — merges the workbook with the register: the master owns a DTC's name,
  rating and GPS, the workbook owns the equipment and the parents. A DTC or RMU that only the master
  knows about is still drawn, with its parent guessed and listed.
- **`resolve_links(items, root_id, drawn)`** — the link to draw for every node, plus the reason when
  it is a guess: a parent with no coordinates sends the line to the nearest ancestor that has them,
  and a node with no usable ancestor is hung off the nearest node that is not downstream of it.
- **Symbols per `TYPE`** — `POLE`/`JUNCTION` as a dot carrying its code, `GOS` as a bar sitting *on*
  the line, `RMU` as a box naming the ways its cables arrive on, `MUSS`/`DTC`/`HT` as before,
  `NODE` as a fallback square.
- **`BENDS` are drawn** — a span with bend rows is routed through them (`<Array as="points">`), so a
  run can follow a road corner instead of cutting across a block.
- **Spur tags** — when `SPURS!LABEL` names a lateral, `<b>S1</b> Chor Bazar lateral` is printed where
  that lateral leaves the trunk.
- **`--topology <file>`** and **`--no-topology`** — point at another copy of the workbook, or go back
  to the inferred trunk for one run.
- **`--list` reports connection coverage** — rows per feeder, split by `BASIS` (`INFERRED` /
  `UNKNOWN` / `SOURCE`), so it is obvious which feeder is still a draft.
- **`free_slot(...)`** — one placement helper for every boxed symbol, walking outwards from the
true position, perpendicular-first for a DTC so the sheet still reads as taps on a line. Replaces
  the per-symbol search loops.

### Changed
- **The trunk is no longer an MST when the workbook has rows for the feeder.** Each node is joined to
  its `PARENT_ID`, the trunk (`SPUR_ID = T`) is drawn heavier than a lateral, and a guessed link is
  drawn **dashed** with the reason printed.
- **`build_geo()` draws symbols first and lines last**, between the nodes' own positions: the line
  marks where equipment is, and a symbol that had to move off a crowded spot is joined back to it by
  a service tap.
- **The MUSS is a node, not a special case.** In a topology run the root row is the source: its
  children hang off its box, and when it is drawn off the page edge for readability those lines no
  longer run past the paper (they start at the box).
- **`unpositioned_block()`** now names the parent each unplaced node hangs off.
- **`--route` still overrides the trunk shape** — a surveyed pole list beats a recorded tree — and
  the run says so; the workbook's non-DTC equipment is still drawn on it.

### Fixed
- **Nodes missing from the sheet** — the RMU, the GOS and the poles are drawn because they are now
  read from the workbook instead of from the master's two-row `NON DTC NODES` sheet.
- **A line drawn from the wrong parent** — `SCD-10` is drawn to `SCD-36` because that is its
  `PARENT_ID`; the tool no longer connects it to whatever was geometrically nearest.
- **`label` clobbered by a spur label** in `build_geo()`, which printed the spur name as the sheet
  title (`SINGLE LINE DIAGRAM OF MAIN LINE FROM MUSS`).
- **`Counter` used before its local import** in `main()` when `--list` was not the path taken.
- **Every DTC in the master reported as "not in the workbook"** — the merge compared against all
  feeders instead of the one being drawn.
- **A rejected coordinate could still pull a link**: `SCD-68` was being attached to `MUSS-MRS` in the
  link list although the outlier check had already dropped it.

### Regenerated
- All eight sheets in `outputs\2026-10-27\` — four geographical (now from the workbook) and four
  schematic (unchanged layout), plus `template/MESCOM_SLD_TEMPLATE.drawio`. All parse clean:
  0 dangling connectors, 0 shapes off the A3 page.

### Notes
- **The workbook was not edited.** `field_data/TOPOLOGY.xlsx` needs one cell changed by hand:
  `RMU WAYS!D14` for `RMU-RAMANNA-SHETTY-PARK`/`VL1` says `N` while `P-001` leaves by that way. The
  drawing follows `NODES` and prints the mismatch; the register row is the stale one.
- Three RMUs still have no `PARENT_ID` (`RMU-KOTE-ROAD`, `RMU-GANGOTHTI-COLLEGE`,
  `RMU-RAMANNA-SHETTY-PARK`), so each is attached to the nearest node and drawn dashed.
- A five-way RMU is still one symbol with the cabled ways named under it; five drawn terminals are
  not part of this change.
- An unfinished workbook remains usable: unfinished rows are guessed, dashed and listed.

---

## 2026-10-26 — `TOPOLOGY_DATA_REQUIREMENTS.md` joins `docs\`

The last document at the root moved into `docs\`. The manual, the history, the handoff and the field
specification now sit in one place.

### Changed
- **`docs/TOPOLOGY_DATA_REQUIREMENTS.md`** — moved with `git mv`. Its header now says that paths are
  written from the repository root, and it names the master as `field_data/DTC DETAILS U4.xlsx`.
- **`docs/README.md`** — folder tree and path table list the document under `docs\`; the two inline
  references (§1 template row, §8 “full field-by-field specification”) carry the new prefix.
- **`CLAUDE.md`** — layout tree and the doc/commit table point at `docs/TOPOLOGY_DATA_REQUIREMENTS.md`.
- **`sld_generator.py`** — one string in the workbook READ ME now says
  `docs/TOPOLOGY_DATA_REQUIREMENTS.md`. No behaviour change.
- **`template/TOPOLOGY TEMPLATE.xlsx`** — regenerated so the shipped workbook carries that path.
- **`field_data/TOPOLOGY.xlsx`** — refreshed from the regenerated template. This was done only
  because the two files were byte-identical beforehand (same SHA-256), so it held no edits; a working
  copy with changes would have been left alone.
- **`docs/SESSION_HANDOFF.md`** — Update 6 records the move.

### Why
The previous entry left this document at the root on the argument that a field team is pointed at it
directly. That still holds, but pointing them at `docs\` costs nothing, and the alternative — four
documents, three in one folder and one on its own — recreates the ambiguity the move was meant to
remove.

### Notes
- The workbook binary is part of this commit for a single sentence in its READ ME; the sheets, rows
  and validations are unchanged.
- No drawing or template layout change, and nothing in `outputs\` needed regenerating.
- The root now holds code, `CLAUDE.md`, `field_data\`, `docs\`, `template\` and `outputs\` — one
  folder per role.

---

## 2026-10-26 — Documentation moves into `docs\`

The three running documents moved out of the root: `README.md`, `CHANGELOG.md` and
`SESSION_HANDOFF.md` are now in `docs\`. Code and data paths are untouched.

### Changed
- **`docs/README.md`** — the manual. §1 retitled “What is in this repository” and rebuilt to show
  `docs\`, `field_data\`, `template\` and `outputs\`, with a line stating that paths in the manual are
  written from the repository root.
- **`docs/CHANGELOG.md`** — this file, newest entry still on top.
- **`docs/SESSION_HANDOFF.md`** — the handoff, with Update 5 noting the move.
- **`CLAUDE.md`** — the layout tree and the doc/commit table now point at `docs/…`. The file itself
  stays at the root, where an agent looks for it.
- **`TOPOLOGY_DATA_REQUIREMENTS.md`** — the two references to the manual and the changelog now carry
  their `docs/` paths.

### Why
The root had grown to five markdown files with three different jobs: the manual, an append-only
history, and a working note. Only one of them is the entry point, and it was competing for attention
with two files a reader almost never opens. Grouping the three keeps the root to things that are
either code, data, templates, or the agent guide.

### Notes
- GitHub renders `docs/README.md` on the repository page when there is no root README, so the
  project still opens on the manual.
- `TOPOLOGY_DATA_REQUIREMENTS.md` deliberately stayed at the root for now: it is the one document a
  field team is pointed at directly, and it sits beside `field_data\`. Say the word and it can join
  the others.
- No code, template or workbook change. Verified by running `--list` and a geographical generation,
  which do not read the docs.

---

## 2026-10-26 — Workbooks move into `field_data\`

Data and templates now live apart. The two workbooks moved into a new `field_data\` folder, leaving
`template\` with nothing but blanks that a rebuild may safely overwrite.

### Changed
- **`field_data/DTC DETAILS U4.xlsx`** — the master, moved from the repository root. `MASTER_DEFAULT`
  now resolves to it, so the default commands are unchanged: `python sld_generator.py --feeder
  "MRS F1"` still reads the same workbook.
- **`field_data/TOPOLOGY.xlsx`** — the connection workbook you fill in, moved from the root. It is
  currently still an unedited copy of the template; the difference is that this one is *meant* to
  diverge.
- **`template\` is blanks only** — `TOPOLOGY TEMPLATE.xlsx` stays the distributable that
  `--topology-template --force` rewrites. Since field work now lives elsewhere, `--force` can no
  longer destroy it.
- **`sld_generator.py`** — new `DATA_DIR` constant; `MASTER_DEFAULT` points inside it. The
  "Master workbook not found" message now names `field_data\` instead of "next to this script".
- **`README.md`** — the folder tree and path table show `field_data\`; §3 says where the master is
  read from; §8 says to copy the template to `field_data\TOPOLOGY.xlsx` and fill *that*; the
  Troubleshooting row points at the new folder.
- **`TOPOLOGY_DATA_REQUIREMENTS.md`** — the collection checklist now copies the template to
  `field_data\TOPOLOGY.xlsx`.
- **`CLAUDE.md`** and **`SESSION_HANDOFF.md`** — layout and the `--force` caveat updated.

### Notes
- No change to the workbook contents or the drawings; this is a move. Verified with `--list` and a
  geographical run, which read the master from its new path and wrote the same 62 positioned DTCs.
- The generated sheets in `outputs\` are unaffected: they name the source workbook, not its path.
- A lock file (`~$…`) will now appear inside `field_data\` while a workbook is open in Excel, not at
  the root.

---

## 2026-10-26 — `CLAUDE.md`: a working guide to the repo

Added a short orientation file for AI-assisted work in this folder. The manual explains the tool to
a person drawing sheets; this explains the repository to whoever — or whatever — is about to change
it, in the order the questions actually come up: what the thing is, how to run it, where the code
lives, what must not break, how to check a change, and what is deceptive about this environment.

### Added
- **`CLAUDE.md`** — commands (including which output folder a run writes), a map of
  `sld_generator.py` top to bottom, the invariants worth losing sleep over (one parent per node;
  five ways with any outgoing one empty; master column letters; workbook sheet ranges), the
  conventions the code follows, and the gotchas that cost time: the Excel lock, `--topology-template`
  refusing to overwrite field edits, `build_geo()` not reading the topology workbook yet, the
  `LF → CRLF` warnings, and the fact that this is a **public** repository holding real asset data.
- **A verification recipe** — since there is no test suite: `--list`, a geo run with `--preview` into
  a scratch folder, a schematic run, an XML parse, and, for workbook changes, an openpyxl read-back of
  headers, ranges, validation and conditional formatting. Plus the four topology checks (duplicate id,
  dangling parent, cycle, five ways per RMU with `CONNECTED` agreeing with `NODES!WAY`).
- **The doc/commit contract** — which document to update for which kind of change, and the seven
  section commit format (`WHAT / WHY / HOW / NEED / USAGE / BENEFITS / REQUIREMENTS`), with the
  reminder that PowerShell has no heredoc so the message goes through `git commit -F`.

### Notes
- Documentation only; no change to the generator, the templates or the drawings.
- The file states the current state of play and the next step (read the topology workbook in
  `build_geo()`), so it will need a line or two there as that work lands — the same discipline
  `SESSION_HANDOFF.md` already asks for.

---

## 2026-10-26 — Empty RMU ways: `CONNECTED` column, empty-by-default register

An RMU has five terminals but not five cables. `ODin` is always fed; the four outgoing ways —
`ODout` included — may be empty, and most units in service run with two or three cables. The workbook
now says so explicitly instead of leaving an empty way to be inferred from a missing row.

### Added
- **`RMU WAYS!CONNECTED`** — `Y` when a cable is landed on the way, `N` when it is empty. It is the
  one column that separates “this way is spare” from “nobody has checked this way yet”.
- **Dropdown for it** — new `LISTS!YES_NO` (`Y`/`N`), validated on `D2:D11`. `NORMALLY` moved to
  column E and is now optional, because an empty way has no switch state worth recording.
- **Two consistency rules in the sheet** — an empty way is greyed out across the row so the cabled
  ways stand out, and `CONNECTED = Y` with `NORMALLY` blank is amber.

### Changed
- **The register starts empty, honestly.** `ODin` is generated `CONNECTED = Y`, `NORMALLY = NC`; the
  four outgoing ways are `N`, with no normal state. Previously all ten rows were `NC`, which claimed
  more than anyone knows.
- **`EXAMPLE` now shows an empty `ODout`.** The demo run leaves the RMU by **`VL2`** (`P-003`), `VL1`
  carries `SCD-08`, and `ODout` plus `VL3` are empty — a unit in service with three cables, which is
  the normal case this change is about.
- **`TOPOLOGY_DATA_REQUIREMENTS.md`** — §3.4 gains a “A five-way unit is not five connections”
  subsection (only `ODin` is always cabled; `ODout` may be empty; blank `NORMALLY`), §5 documents
  `CONNECTED` with the cross-check against `NODES!WAY`, §9 and §14 match the regenerated workbook,
  §13 gains rules 16–17 (register vs connections, empty way with a state) and states that an empty
  way is not an error, §15 and §16 ask for the empty ways per unit.
- **`README.md`** — the five-way section says only `ODin` is always cabled and that any outgoing way
  may be empty, the `RMU WAYS` table row mentions `CONNECTED`, and the glossary defines an empty way.
- **`READ ME` in the workbook** — rule 6 says a unit need not use all five ways; the `RMU WAYS`
  description explains `CONNECTED`; the colours list gains grey (empty way) and amber (`Y` with no
  state); the RMU row remark says not all ways need be used.

### Regenerated
- `template/TOPOLOGY TEMPLATE.xlsx` — `RMU WAYS` is now A–G (7 columns, 10 rows), `CONNECTED` filled
  `Y` for the two `ODin` rows and `N` for the other eight. `NODES` and the other sheets are unchanged
  apart from the RMU remark.

### Notes
- No drawing change: the sheet still labels the unit `5-way` rather than drawing five terminals, and
  `build_geo()` still does not read the workbook. When the reader is written, `CONNECTED = N` is what
  will draw a way as empty instead of guessing a cable.
- The route sample still shows a run entering `ODin` and leaving `ODout`. That is one valid case, not
  a rule — the example workbook deliberately shows the opposite (`VL2` carrying the run on).

---

## 2026-10-26 — Five-way RMUs: `WAY` column and the `RMU WAYS` sheet

An RMU in this network is a five-way switchgear unit — `ODin`, `ODout`, `VL1`, `VL2`, `VL3` — not a
two-terminal link in a chain. The topology format now records that: one `WAY` column on `NODES` for
the terminal a child leaves by, and an `RMU WAYS` register for the terminals themselves.

### Added
- **`NODES!WAY` (column I)** — the RMU way a node leaves by: `ODout`, `VL1`, `VL2` or `VL3`. Blank on
  every row whose parent is not an RMU, and blank on the RMU row itself: the link *into* an RMU is
  way `ODin`, and the RMU's own `PARENT_ID` already names the node feeding it. So the model stays
  one-parent-per-child, and a way is never written twice.
- **Sheet `RMU WAYS`** — five rows per unit (`RMU_ID`, `WAY`, `DIRECTION`, `NORMALLY`, `LABEL`,
  `NOTES`), 10 rows generated from the two RMUs in the master. `ODin` is `IN`, the other four `OUT`,
  all pre-filled `NC`. A spare way has a row here even though it has no child in `NODES`.
- **Dropdown lists for ways** — `LISTS` grew `WAY_ALL` (5 values, used on `RMU WAYS`), `WAY_CHILD`
  (4 values, used on `NODES!WAY` — `ODin` is deliberately not offered to a child) and `DIRECTION`
  (`IN`/`OUT`), with data validation on both sheets.
- **A colour rule for a half-written connection** — `WAY` filled in but `PARENT_ID` blank turns
  orange, because a way with no parent cannot be placed.
- **Five-way examples** — the `EXAMPLE` sheet gains `SCD-08` hanging off `RMU-DEMO` way `VL1`, a
  `WAY = ODout` on `P-003`, and a second block listing the five `RMU WAYS` rows for that unit.
- **`WAY` in the route sample** — `template/route_template.csv` is now 4 columns, with
  `RMU-KOTE-ROAD` written twice (once `ODin`, once `ODout` at the same coordinates) to show a run
  entering and leaving the unit. `read_route()` ignores the column; the docstring says so.

### Changed
- **`NODES` is now A–O, not A–N.** Inserting `WAY` after `SPUR_ID` shifted `KVA`, `PHASE`,
  `NORMALLY`, `CONDUCTOR`, `BASIS` and `REMARKS` one column right. Column widths, wrap columns,
  centred columns, row-height fitting, the autofilter range, the validation ranges and the amber /
  orange / green conditional formatting on `BASIS` all follow the new positions.
- **RMU boxes are labelled `5-way`** — under the name on the schematic and in the geographical
  layout, in small type so the box keeps its size. The library caption now reads
  “Ring main unit, 5-way: ODin, ODout, VL1, VL2, VL3”.
- **`READ ME` in the workbook** — six rules instead of five (the new one is `WAY`), an `RMU WAYS`
  entry under sheets, the pre-filled way count, and the orange `WAY` colour explained.
- **`TOPOLOGY_DATA_REQUIREMENTS.md`** — §3.4 rewritten as the five-way model with the terminal table
  and both link directions, §4 re-specced A–O with `WAY` at I, a new §5 for `RMU WAYS` (which
  renumbers §5–§15 to §6–§16), §10 tier 4 now includes `WAY`, §13 gains five validation rules
  (`WAY` on a non-RMU child, `ODin` on a child, duplicated way, two children on one way, a `NC` node
  downstream of a `NO` way), and §14 has a two-table worked example with the ways.
- **`README.md`** — §5 and §6 describe the `5-way` label, §7 documents the optional `WAY` column in a
  route file, §8 gains a “RMUs: five ways, one row” section plus the `RMU WAYS` row and the new
  counts, §10 says to label the ways when a unit is commissioned, §11 records that the ways are not
  drawn yet, and the glossary defines a way.

### Regenerated
- `template/TOPOLOGY TEMPLATE.xlsx` — 128 `NODES` rows in A–O, 10 `RMU WAYS` rows, 6 visible sheets
  plus hidden `LISTS`. No `WAY` value is pre-filled in `NODES`: neither RMU has a recorded child, and
  the spanning-tree guess never puts a DTC under one.
- `template/MESCOM_SLD_TEMPLATE.drawio` (+ library PNG) with the five-way caption.
- `outputs/2026-10-26/`: all four feeders in both layouts, with the `5-way` RMU labels.

### Notes
- `TOPOLOGY TEMPLATE.xlsx` is locked while Excel has it open; the rebuild reports `Permission denied`
  until it is closed. The new version was generated to a temporary path first and checked before
  replacing the tracked file.
- Still a **collection format, not a drawing input**: `build_geo()` continues to infer the trunk by
  MST and to read only `--route`. `WAY` is recorded now so the reader, when it is written, has the
  data it needs; drawing five terminals per RMU is part of that same step.

---

## 2026-10-26 — Template folder, dated output folders, topology data model

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
