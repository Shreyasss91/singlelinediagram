# MESCOM SLD Tool

Generates distribution **single line diagrams** for the Unit-4 / CSD-2 feeders directly from the
DTC master workbook, as editable `.drawio` files.

Instead of drawing 60–140 DTC symbols by hand in Word, you keep the data in two spreadsheets and
the diagram is derived from them. Every DTC code, name, kVA rating, count and — in geographical
mode — position comes from the master, and the tree both layouts are drawn from (who feeds whom,
which spur, which RMU terminal) comes from the connection workbook, so the sheet and the record can
never disagree.

---

## 1. What is in this repository

```
SLD TOOL\
├── sld_generator.py               the tool
├── docx_to_drawio.py              companion: legacy Word drawing -> .drawio
├── field_data\
│   ├── DTC DETAILS U4.xlsx        the DTC master it reads
│   └── TOPOLOGY.xlsx              the connection workbook you fill in
├── docs\
│   ├── README.md                  this manual
│   ├── CHANGELOG.md               change history
│   ├── SESSION_HANDOFF.md         state of play, updated at the end of a session
│   ├── TOPOLOGY_DATA_REQUIREMENTS.md  what to record to get a real trunk and spurs
│   └── FIELD_CAPTURE.md           proposal: collecting the tree in the field
├── CLAUDE.md                      working guide for changing the tool
├── template\
│   ├── MESCOM_SLD_TEMPLATE.drawio blank A3 sheet + symbol library
│   ├── TOPOLOGY TEMPLATE.xlsx     blank connection workbook, copied from here
│   └── route_template.csv         format sample for --route
├── outputs\
│   └── <YYYY-MM-DD>\              generated sheets, one folder per run day
│       ├── MRS F1 SLD.drawio
│       ├── MRS F1 - GEO SLD.drawio
│       └── … plus a PNG for each sheet when --preview is used
```

Paths in this manual are written from the repository root: `docs/…`, `field_data/…`, `template/…`.
Commands are run from the root, too.

| Path | Purpose |
|---|---|
| `sld_generator.py` | The whole tool. Reads the workbook, writes `.drawio` files, optionally renders a PNG preview. |
| `docx_to_drawio.py` | Companion importer: a legacy Word drawing (`.docx` with Word shapes) becomes an editable `.drawio`, one page per Word page. Reads the DOCX, never a PNG. See §9. |
| `field_data/DTC DETAILS U4.xlsx` | The DTC master. **The script reads this file from here by default.** |
| `field_data/TOPOLOGY.xlsx` | Your copy of the connection workbook, and the one you fill in. Kept apart from `template/` so a template rebuild can never overwrite field work. **Both layouts are drawn from this file**; see §8. |
| `docs/README.md` | This manual. |
| `docs/CHANGELOG.md` | Every change, newest first. |
| `docs/SESSION_HANDOFF.md` | Where the work stands, and what is next. |
| `docs/TOPOLOGY_DATA_REQUIREMENTS.md` | The data-collection guide: what is missing, the field specification, how spurs are derived, where to get each item. |
| `docs/FIELD_CAPTURE.md` | Proposal, nothing implemented: capturing the tree by the order nodes are visited, positioning without GPS, snapping track to nodes, and what that unlocks (voltage drop, loss, evolution). |
| `template/MESCOM_SLD_TEMPLATE.drawio` | Two-tab template: a blank A3 sheet and a symbol library (MUSS, DTC, HT, GOS, RMU, junction, bus). |
| `template/TOPOLOGY TEMPLATE.xlsx` | One row per node with its parent. Pre-filled from the master with a first-draft connection; see §8 and `docs/TOPOLOGY_DATA_REQUIREMENTS.md`. This is the blank, not your working copy. |
| `template/route_template.csv` | Format sample for surveyed pole coordinates (`--route`). Values are examples only; the optional `WAY` column shows a run entering and leaving an RMU. |
| `outputs/<date>/` | Every generated sheet goes here, named after the feeder and layout. Today's runs do not overwrite yesterday's. |
| `CLAUDE.md` | Notes for whoever changes the tool: commands, code map, invariants, how to verify a change, and the traps in this environment. |

`field_data/` holds the two workbooks that are data rather than templates: the master it reads and
the connection workbook you fill in. `template/` holds things that ship and get regenerated.

`outputs/` is git-ignored: every sheet in it is generated and can be rebuilt from the master.

`~$DTC DETAILS U4.xlsx` is an Excel lock file. It appears inside `field_data\` whenever a workbook is
open in Excel and can be ignored (see Troubleshooting). The same applies to a `~$` file beside the
topology template or `TOPOLOGY.xlsx`.

---

## 2. Install

Python 3.9 or newer, then:

```bash
pip install openpyxl pymupdf
```

* `openpyxl` — required, reads the workbook.
* `pymupdf` — only for `--preview`. Without it everything else still works.

Nothing else is needed. The script does not need an internet connection.

---

## 3. The master workbook

`sld_generator.py` looks for `field_data\DTC DETAILS U4.xlsx` **beside itself**. Point it elsewhere
with `--master "D:\path\to\file.xlsx"` if you ever move it.

### Sheet `ALL` — one row per DTC

The script uses these columns (A is column 1):

| Column | Header | Used for |
|---|---|---|
| B | `SCD` | code prefix (`SCD`, `SCF`, `SCE`) |
| C | `DTC_CODE` | code number — rendered zero-padded, so `1` becomes `SCD-01` |
| E | `DTC_Name` | the name printed under the symbol, cropped to 22 characters |
| F | `Feeder` | feeder this DTC belongs to; must match `--feeder` |
| I | `KVA` | rating shown inside the square (25 / 63 / 100 / 200 / 250) |
| J | `Station` | `110/11KV MRS` or `110/11KV MANDLI`; printed on the MUSS box |
| AD | `GPS LAT` | decimal degrees, e.g. `13.9262197` |
| AE | `GPS LONG` | decimal degrees, e.g. `75.5799872` |

All other columns (meter number, make, CT ratio, photos …) are ignored but harmless.

Current coverage at the time of writing: **119 DTCs, 113 with GPS** — MRS F1 63, Mandli F22 32,
Mandli F13 19, Mandli F3 5.

### Sheet `HT` — one row per HT service

| Column | Header | Used for |
|---|---|---|
| E | `RR No.` | code, must look like `HT-87` |
| F | `NAME` | consumer name |
| H | `CD(KVA)` | contracted demand shown in the label |
| L | `Feeder` | feeder it belongs to |

HT rows only appear if their `Feeder` matches the `ALL` sheet's spelling. Today only MRS F1 has
HT entries (HT-28, HT-87, HT-92, HT-188, HT-193). They have no GPS, so in geographical mode they
are listed in the `UNPOSITIONED` box.

### Sheet `NON DTC NODES` — MUSS and RMU positions

Everything on the feeder that is **not** a transformer lives here. Four columns:

| Column | Header | Used for |
|---|---|---|
| A | `NODES` | the node name; it must contain `MUSS` or `RMU` — that is how the type is detected |
| B | `GPS LAT` | decimal degrees |
| C | `GPS LONG` | decimal degrees |
| D | `GPS HYPERLINK` | ignored (kept for the Google Maps link) |

Current rows:

| Node | Type | Latitude | Longitude |
|---|---|---|---|
| `110/11KV MRS MUSS` | MUSS | 13.910279 | 75.595896 |
| `110/11KV MANDLI MUSS` | MUSS | 13.917882 | 75.562761 |
| `KOTE ROAD RMU` | RMU | 13.930495 | 75.585759 |
| `GANGOTHTI COLLEGE RMU` | RMU | 13.931183 | 75.582701 |

**How the MUSS is matched.** Column `J` of sheet `ALL` says `110/11KV MRS` or `110/11KV MANDLI`.
The tool looks for a MUSS node whose name contains that text, so `110/11KV MRS` finds
`110/11KV MRS MUSS`. If you add a new substation, keep the same pattern (`110/11KV <NAME> MUSS`)
and the sheets will pick it up with no code change.

**How an RMU is assigned to a feeder.** From its coordinates: whichever feeder's nearest DTC is
closest to the RMU owns it. That is why both RMUs appear on the MRS F1 sheet and on none of the
Mandli sheets. Move an RMU far enough and it will switch feeder by itself.

### Rules the data must follow

* Codes must be unique per feeder. Duplicates are skipped (first one wins).
* Feeder names are compared ignoring case, dashes and double spaces, but `Mandli F13` ≠ `Mandli F3`.
* GPS must be **decimal degrees**, not degrees-minutes-seconds. A value like `3.92` instead of
  `13.92` will be rejected as an outlier (see §6).

---

## 4. Quick start

```bash
cd "D:\Office PC\D DRIVE\ESTIMATES\SLD TOOL"

python sld_generator.py --list                        # feeders and GPS coverage
python sld_generator.py --feeder "Mandli F13" --name "Mandli F13"   # schematic sheet
python sld_generator.py --feeder "MRS F1" --name "MRS F1" --geo     # geographical sheet
```

Every run prints what it did and where it wrote it:

```
  topology    : field_data\TOPOLOGY.xlsx  |  63 DTC, 5 HT, 1 MUSS, 2 RMU
  feeder      : MRS F1  (63 DTCs, 5 HT services)
  layout      : geo  |  65 nodes placed, 6 unpositioned  |  1 m = 0.89 px
  source      : connection workbook, 2 link(s) guessed
  written     : outputs\2026-10-27\MRS F1 - GEO SLD.drawio
```

**Output goes to `outputs\<today's date>\`.** A run on 26 October 2026 writes to
`outputs\2026-10-26\`; tomorrow's run makes `outputs\2026-10-27\` and leaves today's files alone.
Use `--outdir` to choose another folder, or `--out` for one specific file.

### All options

| Option | Meaning |
|---|---|
| `--feeder "Name"` | which feeder to draw (must match sheet `ALL` column F) |
| `--name "Text"` | label printed on the sheet and used in the filename |
| `--title "Text"` | full override of the title line |
| `--out "file.drawio"` | write this exact file instead of the dated folder |
| `--outdir "folder"` | date-stamped folder is replaced by this one |
| `--geo` | geographical layout instead of the schematic |
| `--muss "lat,lon"` | override the substation position, e.g. `--muss "13.9305,75.58584"` |
| `--extent cluster\|all` | geo only: fit the node network (default) or include the MUSS in the fitted area |
| `--route file.csv` | surveyed pole coordinates for the trunk (see §7); the run follows these points exactly |
| `--topology file.xlsx` | use a different connection workbook (default `field_data/TOPOLOGY.xlsx`) |
| `--no-topology` | ignore the connection workbook: infer the trunk (see §6) and draw the schematic in code order |
| `--master path.xlsx` | use a different workbook |
| `--preview` | also write a PNG next to the `.drawio` |
| `--template` | rewrite `template/MESCOM_SLD_TEMPLATE.drawio` |
| `--topology-template` | write `template/TOPOLOGY TEMPLATE.xlsx` from the master (see §8) |
| `--force` | allow `--topology-template` to overwrite a file you have edited |
| `--list` | show the feeders found in the workbook |

Output names are predictable, so one layout never overwrites the other:

* schematic → `outputs\<date>\<name> SLD.drawio`
* geographical → `outputs\<date>\<name> - GEO SLD.drawio`

With `--name "MRS F1"` on 26 October 2026 that is `outputs\2026-10-26\MRS F1 SLD.drawio` and
`outputs\2026-10-26\MRS F1 - GEO SLD.drawio`.

---

## 5. The schematic layout (default)

A clean, conventional sheet: a vertical distributor on the left, horizontal laterals across the
page, the nodes of the feeder hung beneath in **the order the connection workbook connects them**, a
MUSS box at the top right, and the 25/63/100/200/250 summary table bottom right.

The order is depth first from the MUSS in workbook row order, so a parent is always drawn before its
children and a spur's nodes stay together in one run of the grid — the same grouping the geographical
sheet shows with lines. It is the only order the sheet uses; the register's code order is not mixed
in, because `SCD-49` sorted after `SCD-48` is not the same statement as `SCD-49` fed from `SCD-48`.

Each box says what feeds it, in small type under its code:

```
CHOR BAZAR
SCD-19
from SCD-01
            S1
```

A `POLE` is a dot with its code, the GOS a bar (open when `NORMALLY = NO`), and the RMU a box whose
second line names its ways in words (`5-way - out VL1`; a three-way unit says `3-way`), so an
equipment row the master does not hold still has its symbol even where the grid has no room for a
panel. A spur tag (`S1`, plus the name from `SPURS!LABEL`) is printed under
the node the spur leaves from, not under its first box. A node with no `PARENT_ID` — or one the
workbook has no row for — is drawn after the trunk with no parent named, and the run says how many
and which:

```
! 2 node(s) drawn after the trunk, no parent named - no PARENT_ID recorded:
  RMU-RAMANNA-SHETTY-PARK, SCF-03
```

The column count adapts to how many nodes there are — 22 nodes get 8 per lateral and 3 laterals,
70 nodes get 14 and 5. Slots never drop below 92 px, so labels cannot collide horizontally. HT
services and RMUs come out of the same tree, so they are no longer appended at the end.

Without a workbook (`--no-topology`, or no topology file) the rows are drawn in the register's code
order and nothing is said about parents, because none are recorded. The summary table stays the kVA
histogram of the master's DTCs either way.

Use it when you want a **neat drawing** for submission. Use `--geo` when you want the drawing to
**look like the network on the ground**.

---

## 6. The geographical layout (`--geo`)

Every node is placed at its surveyed position, north up, at a uniform scale, and joined the way the
connection workbook (§8) records it. What the generator does:

1. **Reads and cleans the coordinates.** Anything more than 5 km from the median of the feeder is
   dropped and reported, so a single typo cannot stretch the sheet. Example from the current data:
   ```
   ! SCD-68: 3.926426, 75.5808048 is 1110 km from the rest of the feeder - ignored (check the master)
   ```
2. **Projects to the sheet.** A local equirectangular projection (≈111.32 km per degree of
   longitude, 110.54 km per degree of latitude) converts degrees to metres; metres then map to the
   A3 drawing area with one uniform scale for both axes, so distances stay proportional.
3. **Joins each node to its parent.** `PARENT_ID` in the connection workbook is the tree: a DTC fed
   from `SCD-36` is drawn to `SCD-36` and to nothing else. The trunk (`SPUR_ID = T`) is drawn heavier
   than a lateral (`S1`, `S2`, …), a `BENDS` row bends that one span around the corner it records,
   and a link the tool had to guess is drawn **dashed**. With no workbook — or with `--no-topology` —
   a **minimum spanning tree** over the coordinates is used instead: a likeness, not a record.
4. **Routes each run orthogonally, between edge midpoints.** A connector leaves and enters at the
   midpoint of one of the symbol's four edges, never at an arbitrary point on its border. The tool
   tries the straight, one-bend and two-bend Z shapes for every pair of ports, keeps the shortest one
   that does not pass through another symbol, and when every simple shape is blocked it searches an
   8 px grid for a way round. Every turn is a right angle, with a small square lead off each port. A
   span with `BENDS` rows keeps its recorded shape. Equipment that sits *in* the run — a pole, the
   GOS — is drawn on the line; a DTC or HT service hangs *beside* it on a short tap, and an RMU is a
   panel the runs go round and land on by their portion, so a run does not sit on a transformer box
   or across a compartment.
5. **Draws each node as the symbol for its `TYPE`.** A DTC is a square with its kVA inside, an HT
   service a white square, a pole a dot carrying its code, the GOS a bar sitting on the line, the
   MUSS the source box, and the RMU its panel **divided into one portion per way**. A symbol that has
   to move off a crowded spot is joined back to its own point by a thin square leader ending in a
   visible dot, so the surveyed position is still marked even though the cable meets the symbol.

   **An RMU shows its ways inside the symbol.** `ODin` is the left end of the panel, `ODout` the
   right, and the VLs sit between them in register order — a five-way unit is five portions, a
   three-way one is three. A way a cable lands on is **filled** and named with its way and switch
   state (`ODin NC`); a way with nothing on it is white and named `empty`, so a spare way cannot be
   read as a row missing from the workbook. A cable off an RMU starts at the portion the workbook
   names in `NODES!WAY`, on the edge facing the node it feeds, so the sheet shows which way a run
   leaves by. With no workbook the unit is drawn as a plain `5-way` box instead — nothing is known
   about its ways.

   **A GOS shows its state.** `NODES!NORMALLY` decides: `NC` (or blank) is a solid bar across the
   line, `NO` is a hollow blade hinged on the line and lifted off it, labelled `GOS-01 (open)`. The
   run reports what stands behind an open point:
   ```
   ! NODES!131 GOS-01: NORMALLY = NO - drawn as an open point - 5 DTC(s) / 800 kVA recorded
     downstream; while it stands open that load is fed from elsewhere
   ```
6. **Places the labels.** Each label tries twelve positions (below, above, right, left, at three
   distances), starting on the side away from a moved symbol's tap, and keeps the freest one. The
   rectangle it took is then an obstacle for the connector router, so a line goes round a label when
   the gaps allow it. Where two nodes sit on the same spot, the symbols are nudged apart so both
   ratings stay visible.
7. **Adds the drawing aids** — a scale bar in metres and a north arrow, so the sheet cannot be
   mistaken for a to-scale survey.
8. **Lists what it could not place** in a dashed `UNPOSITIONED` box along the bottom, with code,
   name and the parent it hangs off, so nothing silently disappears and a missing coordinate can be
   filled in from the register.

Because the runs are routed at right angles and a transformer hangs off the line rather than on it,
the sheet shows *what is connected to what*, not the road the cable follows. The true shape comes
from `BENDS` (§8): a couple of points per lateral is enough to make a run follow its road.

### Where a line meets a symbol

A port is the midpoint of one of the symbol's edges, and the side is chosen for the path, not fixed
by type: on a DTC it is usually the edge facing the parent, on an RMU the portion edge facing the
node, on the MUSS the side facing the network. draw.io is told the fraction (`exitX/exitY`,
`entryX/entryY`), so dragging a symbol in the editor keeps the line on that midpoint instead of
letting it slide to a corner.

The line is scored against every symbol already placed: a crossing costs more than any detour, a
label is a smaller penalty. A 14 px square lead takes the line off the port before its first turn; if
a neighbouring box is closer than that the lead is shortened, so the stub cannot reach into it. On a
dense sheet a run can still be forced through a label or behind a box. The console reports it:

```
connectors : 64 routed, 1 round a box, 0 still crossing one
```

`routed` counts the connectors the router chose a path for, `round a box` those it had to detour for
a symbol, and `still crossing one` the ones no clean path existed for. A line that has to cross a
box is drawn anyway, because the connection is the fact the sheet is signed against — but it is
counted, never hidden.

### Where the shape came from

The subtitle names the source, so a printed sheet says how much to trust it:

| Subtitle says | Meaning |
|---|---|
| `connection workbook` | parents, spurs, ways and bends were read from `field_data/TOPOLOGY.xlsx` |
| `connection workbook, 2 link(s) guessed` | two nodes had no usable parent and were attached to the nearest node; those links are drawn dashed |
| `inferred trunk` | no workbook rows for this feeder, or `--no-topology`: minimum spanning tree |
| `surveyed route` | `--route` was given; the run follows that pole list |

A parent that has no coordinates cannot be drawn to, so the link goes to the nearest ancestor that
has them and the console says which node was skipped. Nothing is drawn from a coordinate the tool
rejected: `SCD-68`'s `3.926426` lands in `UNPOSITIONED` and pulls no line with it.

### Substation position

The trunk needs a starting point and the MUSS box needs a place. Both come from sheet
`NON DTC NODES` (§3) — no configuration needed:

```
muss : 13.910279, 75.595896  (from 'NON DTC NODES')
```

`--muss "lat,lon"` overrides it for one run. With no MUSS row at all, the trunk starts at one end of
the network and the box is marked `(position approx.)`.

### When the substation is far away (off-map MUSS)

A substation can sit well outside its load centre — MRS MUSS is **2.6 km** from the middle of its
network, Mandli MUSS **1.7 km**. Fitting both on one A3 sheet would shrink the DTC cluster to
0.50 px/m, where no label can be read.

So the tool measures the cost first. If including the MUSS would shrink the drawing by more than
30 %, it keeps the network at full scale and draws the substation at the **page edge, in its true
direction**, joined by a **dashed** line and labelled:

```
110/11 KV
MRS
(2.6 km beyond the edge, not to scale)
```

The dashed line means “the line continues”; the distance is stated so the sheet cannot be misread
as a to-scale drawing. Use `--extent all` when you would rather have the true geometry and a small,
dense cluster.

### RMUs

Every RMU that belongs to this feeder is drawn **split into its ways**: one panel on the surveyed
position, divided inside into one portion per way, in the order a unit is read — `ODin` at the left
end, `ODout` at the right, the VLs between:

```
┌──────┬──────┬───────┬───────┬────────┐
│ ODin │ VL1  │ VL2   │ VL3   │ ODout  │
│  NC  │      │ empty │ empty │ empty  │
└──────┴──────┴───────┴───────┴────────┘
     RAMANNASHETTY PARK RMU
```

Each portion is `34 × 26` px, so a five-way unit is a 170 px panel and a three-way one is 102 px. The
number of portions follows the `RMU WAYS` register, not a constant: three rows give `ODin | VL1 |
ODout`, five give the panel above. A way a cable lands on is **filled** and named with its way, plus
its switch state where the register records one (`ODin NC`); a way with nothing on it is white and
named `empty`, so a spare way cannot be read as a row missing from the workbook. `NODES!WAY` says
which cables land where (§8): a child whose `PARENT_ID` is this RMU names the way it leaves by, and
its line starts at that portion — not at the panel — so the sheet shows which way a run goes.

A cable leaves or enters by the portion's edge that faces the other end of the run: the ring ways at
their own end of the panel, a VL at its top or bottom, whichever side the node is on. A middle
portion's sides are dividers, so a VL has only those two free edges, and a feed into `ODin` comes in
square to the end or the top or bottom, never across the other portions.

With no workbook (`--no-topology`, or no topology file), the unit is drawn as a plain box marked
`5-way` under its name: nothing is known about its ways, and five empty portions would claim it.

The schematic layout has no room for a panel in its grid, so there the ways are said in words under
the box — `5-way - out VL1`, or `3-way - no way out recorded` for a three-way unit.

When the `RMU WAYS` register disagrees with the connections — a way marked `CONNECTED = N` that a
child leaves by, or `CONNECTED = Y` on a way no child names — the run prints it and draws what
`NODES` says, because the connections are the tree and the register is a separate record that needs
the update.

### Reading the scale line

```
layout : geo | 62 positioned, 6 unpositioned | 1 m = 0.96 px
```

0.96 px per metre means the whole 1.3 km feeder fits across the A3 sheet. Small feeders get a larger
scale but are capped at 3.2 px/m so a 300 m network does not fill the page.

---

## 7. Surveyed routes (`--route`)

If you have pole coordinates — from GIS, a field survey app, or a Google Earth KMZ converted to
CSV — the trunk can follow the real line instead of an inferred one.

Format: one row per pole, **in order along the line**, `.csv` or `.xlsx`:

```
NODE_ID,LAT,LONG,WAY
MUSS,13.93050,75.58580,
P-001,13.92990,75.58460,
P-002,13.92900,75.58330,
RMU-KOTE-ROAD,13.92850,75.58265,ODin
RMU-KOTE-ROAD,13.92850,75.58265,ODout
P-003,13.92800,75.58200,
```

* NODE_ID is free text (a pole number is ideal).
* LAT/LONG decimal degrees. A header row is skipped automatically because it is not numeric.
* Rows without two valid numbers are ignored.
* For `.xlsx`, the first sheet is read and the same column order applies.
* `WAY` is **optional and ignored** here. It exists so the same file can show which terminal a run
  enters and leaves an RMU by: the RMU row is written twice, once as `ODin` and once as `ODout`, at
  the same coordinates. Only the topology workbook (§8) acts on the value.

Each DTC is then attached to the **nearest point on the drawn line** by perpendicular projection,
and its spur is drawn from there. `template/route_template.csv` shows the format — its numbers are
invented, replace them before use.

One `--route` file is one chain. It can shape the trunk, but it cannot express branches or say which
DTC belongs to which lateral — for that, use the topology workbook (§8).

---

## 8. Topology — trunk, spurs and connections

GPS says *where* every DTC is; it does not say *what it is connected to*. Until that is recorded,
the geographical sheet joins the points with a minimum spanning tree — a geometric guess. It is
usually close, but it can put a DTC on the wrong lateral, cut a line across a block, or draw a spur
through a transformer symbol.

The fix is one extra column per node: **`PARENT_ID`**, the node immediately upstream towards the
MUSS. Read together, the parent links are the whole feeder tree — trunk, spurs and the order of
equipment along each one.

The one addition is at an RMU: because the unit has five ways, a child of an RMU also names the way
it leaves by in **`WAY`**. The link into the unit is `ODin`, so it needs nothing extra — see
“RMUs: five ways, one row” below.

### The template

`template/TOPOLOGY TEMPLATE.xlsx` is pre-filled from the master:

| Sheet | What it holds |
|---|---|
| `NODES` | one row per DTC, HT, MUSS, RMU — 128 rows today. `PARENT_ID`, `SPUR_ID` and `WAY` are the columns to correct. |
| `SPURS` | the laterals found in `NODES`, with DTC count and connected kVA. Add a `LABEL` for each one. |
| `RMU WAYS` | the five ways of every RMU — `ODin`, `ODout`, `VL1`, `VL2`, `VL3` — with direction, `CONNECTED`, normal state and label. 10 rows today: `ODin` cabled, the four outgoing ways empty. |
| `BENDS` | optional shape points, for a span that must follow a road instead of a straight line. |
| `EXAMPLE` | a small complete feeder showing exactly how the rows read, including a five-way RMU. |

`PARENT_ID` is pre-filled from a minimum spanning tree and marked `BASIS = INFERRED` — amber, and
to be treated as a first draft. `BASIS = UNKNOWN` (14 rows) means there was no usable coordinate,
or the node is an HT/RMU whose tap is not recorded. Those need a manual parent.

**Work on a copy in `field_data\TOPOLOGY.xlsx`, not on the template.** Copy
`template/TOPOLOGY TEMPLATE.xlsx` there once, then fill that file in. `template/` is the blank that
`--topology-template --force` rewrites, so anything edited in place there is destroyed on the next
rebuild. Keep the sheet names and column order as they are — the reader will key on them.

**The generator reads this workbook** — it is what shapes a `--geo` sheet. For the feeder being
drawn it takes the rows whose `FEEDER` cell is that feeder, the MUSS above them, and any row whose
`FEEDER` cell was left blank but hangs under one of them, then:

* joins every node to its `PARENT_ID` — so a DTC fed from `SCD-36` is drawn to `SCD-36` and to
  nothing else;
* draws the equipment the master does not hold — poles, the GOS and the RMU;
* divides an RMU into one portion per recorded way and lands each cable on the way `NODES!WAY` names;
* bends a span through its `BENDS` rows instead of running it straight;
* prints `<b>S1</b> Chor Bazar lateral` at the branch point when that spur has a `LABEL` on sheet
  `SPURS`, so the sheet and the spur register read together;
* checks the rows against each other before drawing, and says what it had to guess.

When `NODES` and the `RMU WAYS` register disagree, the connection tree wins: a child parented to an
RMU is drawn on its recorded way, and the register row is reported as needing an update.

```
  topology    : field_data\TOPOLOGY.xlsx  |  19 DTC, 1 GOS, 1 MUSS, 1 POLE, 1 RMU
  ! NODES!6 RMU-RAMANNA-SHETTY-PARK: no PARENT_ID - its parent is guessed from the nearest node
  ! RMU WAYS!14 RMU-RAMANNA-SHETTY-PARK VL1: CONNECTED = N but P-001 leaves by that way - set it to Y
  links       : 1 guessed, drawn dashed - RMU-RAMANNA-SHETTY-PARK
```

### What the reader checks

| Check | Reported as |
|---|---|
| duplicate `NODE_ID` | **error** — nothing is drawn until it is fixed |
| `PARENT_ID` that is not a row in the workbook | **error** |
| a node that is its own ancestor | **error** |
| `WAY` naming `ODin`, or on a node whose parent is not an RMU | **error** |
| two children of one RMU claiming the same way | **error** |
| no `PARENT_ID` on a node that is not a MUSS, with coordinates | warning — parent guessed from the nearest node, drawn dashed |
| `CONNECTED = N` on a way that a child leaves by | warning — the register needs updating |
| `CONNECTED = Y` on a way that no child names | warning — is a node missing? |
| an empty way carrying a `NORMALLY` value | warning — an empty way has no switch state |
| a way marked `NO` with load recorded downstream | warning — the load is off while it stands open |
| a `GOS` with a blank `NORMALLY` | warning — drawn closed; record `NC` or `NO` |
| a `GOS` marked `NO` | warning — drawn as an open point, with the DTCs and kVA behind it named |
| `LAT`/`LONG` outside lat 13.5–14.3, lon 75.1–76.0, or only one of the pair filled | warning — treated as unpositioned |
| a `DTC`/`HT` row whose code is not in the master | warning |
| a `BENDS` row for a span that is not a parent link | warning — the bend is ignored |

Errors stop the run and name the sheet, row and column to change. Warnings are printed and the sheet
is still drawn — an unfinished workbook still produces a usable sheet, with the guesses dashed.

`--topology <file>` points at another copy (useful for trying a change without touching the field
file); `--no-topology` goes back to the inferred trunk for one run.

Rebuild it from the current master when the workbook changes:

```bash
python sld_generator.py --topology-template          # refuses to overwrite a filled file
python sld_generator.py --topology-template --force  # discards your edits
```

### RMUs: five ways, one row

An RMU here is **one node with five terminals**, not a link in a chain:

| Way | Direction | Meaning |
|---|---|---|
| `ODin` | in | the incoming cable — the RMU is fed through this way, always cabled |
| `ODout` | out | one of the four outgoing ways; often the continuing run, but it may be empty |
| `VL1`, `VL2`, `VL3` | out | the other outgoing ways — any of them may be empty |

Two places, two jobs:

* **`NODES!WAY`** answers *which way did this child leave by*. Any row whose `PARENT_ID` is an RMU
  must name `ODout`, `VL1`, `VL2` or `VL3`. The link **into** the RMU is way `ODin`, so it needs no
  `WAY` value — the RMU's own `PARENT_ID` already says where it is fed from. Every other row leaves
  `WAY` blank.
* **`RMU WAYS`** answers *what is on the switchgear*. Five rows per unit: direction, whether a cable
  is landed (`CONNECTED`), the normal state (`NC`/`NO`) and a label per way.

```
                  ┌── ODout ── (empty)
P-002 ── ODin ──[RMU]── VL1 ── SCD-08
                  ├── VL2 ── P-003 ── ...
                  └── VL3 ── (empty)
```

**A five-way unit is not five connections.** Only `ODin` is always cabled; the four outgoing ways —
`ODout` included — are commonly empty, and a unit sitting in service with two cables is perfectly
normal. An empty way has no child row in `NODES` and `CONNECTED = N` on `RMU WAYS`, with `NORMALLY`
left blank because there is no switch state to record.

Keeping the connections in `NODES` alone means a child is still parented exactly once, so the tree
cannot disagree with itself. `RMU WAYS` is a register, not a second copy of the tree.

### How a spur is defined

The **trunk** is the main run from the MUSS. Every branch off it is a **spur**. In the template:

* `SPUR_ID = T` — on the trunk
* `SPUR_ID = S1`, `S2`, … — on that lateral
* the spur begins at the first node whose parent is on the trunk

No separate grouping table is needed: `P-002` having two children (`P-010` and an RMU) is what makes
the `P-010` branch a spur. DTC count and kVA per spur are computed from the same `NODES` rows, so
they cannot drift from the node list.

### What to fill in first

Start with `PARENT_ID` and `SPUR_ID` for the DTCs on one feeder, reading the hand-drawn SLD. Poles
are optional — parent each DTC to the upstream DTC if that is what the drawing shows. Pole rows give
better branch points and can be added later without redoing the DTCs.

The full field-by-field specification, the collection order, the sources for each item and a
worked example are in **`docs/TOPOLOGY_DATA_REQUIREMENTS.md`**.

---

## 9. Opening, editing, exporting

**draw.io desktop** (free, offline): install from the drawio-desktop releases page and open the
`.drawio` file. Export with `File ▸ Export as ▸ PDF`, page size **A3 landscape**.

**VS Code** (already installed): add the *Draw.io Integration* extension
(`hediet.vscode-drawio`), then click any `.drawio` file in the Explorer. It edits and saves in
place.

Once open:

* Every DTC is a real shape joined by a real connector. **Drag a DTC and its line follows** — no
  redrawing branches, which is what made the Word version painful.
* Symbols can be copied from the `LIBRARY` tab of `template/MESCOM_SLD_TEMPLATE.drawio` (GOS, RMU,
  HT, junction, bus).
* To tidy a dense area, move the junction dots rather than the symbols.
* **Save your edited copy under a new name.** Regenerating writes into today's folder and overwrites
  the generated file of the same name there — and nothing older.

### Importing a legacy Word drawing

`docx_to_drawio.py` turns a sheet made in Word — the drawings this tool replaced — into an editable
`.drawio`, one page per Word page:

```bash
python docx_to_drawio.py "field_data/test/SLD_MRSF1_Feeder_and_BHRoad_Feeder.docx" \
    --out "outputs/2026-10-29/legacy/SLD_MRSF1_and_BHRoad_legacy.drawio"
```

**It reads the DOCX, not the PNG.** A page of that document is ~400 vector shapes; a 600 dpi render
of it is pixels, and no tracer can turn pixels back into named DTC symbols. The DOCX still has the
shapes, so that is what the converter reads. Rectangles, the DTC `frame` (written as two concentric
squares), lines, rounded MUSS boxes, labels and kVA numbers are mapped 1:1 at 1 in = 100 px, the same
A3 units the generator uses.

It is a redraw, not a facsimile. Word's frame inset and text-box insets are approximated, the page is
re-centred rather than placed at Word's own offset, and a label may sit a few pixels from where Word
put it. **Compare the result with the PNG export before submitting it.** Anchors Word placed outside
the page (one stray shape per page in the sample, at x ≈ 18.8 in) are dropped and counted.

The import is one-way and offline: it does not touch the master, the topology workbook or the
generator. Once a page is in draw.io it is edited like any other sheet, and the data-driven path
(`--feeder … --geo`) stays the way to regenerate it.

---

## 10. Everyday workflows

**A feeder gets new DTCs**
1. Add rows to sheet `ALL` (code, name, feeder, kVA, station, GPS).
2. `python sld_generator.py --feeder "<feeder>" --geo --preview`
3. Open the new file in `outputs\<today>\`, reposition anything the inference got wrong, export the PDF.

**A DTC is shifted on the ground**
Update `GPS LAT` / `GPS LONG`, regenerate, compare the PNG preview — no drawing work at all.

**A DTC changes rating (63 → 100 kVA)**
Update column `I` and regenerate. The summary table and the symbol both follow.

**A new RMU or substation is commissioned**
Add a row to sheet `NON DTC NODES` — the name must contain `RMU` or `MUSS` — then regenerate. An RMU
attaches itself to whichever feeder's DTCs are nearest; a MUSS is matched to the station text in
column `J` of sheet `ALL`. No code change, no configuration. Then rebuild the topology template (or
add the five `RMU WAYS` rows by hand), put a `WAY` on every child of the unit, and set `CONNECTED = Y`
only on the ways that really have a cable.

**You need a sheet for a feeder that has never been drawn**
Add its DTCs to `ALL` with the feeder name, run `--list` to confirm it appears, then generate.
Nothing in the code is feeder-specific.

**You want the real trunk and spurs instead of a guess**
Fill `field_data/TOPOLOGY.xlsx` (§8): copy the template there once, then correct `PARENT_ID`,
`SPUR_ID` and `WAY` per feeder. Both sheets are drawn from it on the next run, and every row
you leave `INFERRED` is still an MST guess worth correcting.

**You want to prove the drawing matches the register**
The summary table on every sheet is computed from the same rows that produced the symbols, so a
mismatch is impossible by construction. Once the topology workbook is in use, the per-spur counts
come from the same rows as the spur lines.

---

## 11. Known limitations

* **A tree the workbook does not record is inferred.** For a feeder with no rows, or with
  `--no-topology`, the master has no parent/child data at all, so the trunk is a minimum spanning
  tree: it can take a shortcut the real line does not take and a lateral can be guessed wrongly.
  Fill `field_data/TOPOLOGY.xlsx` (§8), or use `--route` for a surveyed trunk. The four current
  feeders have their tree recorded, so this is the `--no-topology` path and the path a new feeder
  takes until its rows exist.
* **A guess is drawn dashed, but it is still a guess.** A node with no `PARENT_ID`, a DTC the
  workbook has no row for: each is attached to the nearest node that is not downstream of it, and
  the console lists it. Read the console before signing the sheet.
* **Dense areas need a hand.** On MRS F1 the DTCs are 48 m apart, which is ~46 px on A3, while a
  label is 100 px wide. The label placement minimises overlap but cannot eliminate it; the tightest
  corner will still need a nudge in draw.io. F13 and F22 have plenty of room.
* **A connector can still cross a label.** Labels are obstacles in the router, and the port pair is
  chosen to avoid them where a clean route exists, but on MRS F1 twenty crossings remain — ten on
  network links, ten on the thin position leaders — because the labels are packed into the same
  46 px gaps as the symbols. The other three geographical sheets have six between them. Moving a
  label in draw.io is the fix for the few that matter; the tool will not silently reroute a feed to
  dodge text.
* **A position leader can pass a neighbour.** The leader that ties a moved symbol back to its
  surveyed dot picks the nearest clean edge, but in the tightest part of MRS F1 four of them still
  touch a neighbouring box. They are 0.8 pt lines, not feeds, and are reported by the same
  `connectors` count as a crossing.
* **A switch state is drawn, not solved.** A GOS with `NORMALLY = NO` shows an open blade and its
  downstream DTCs and kVA, and a `RMU WAYS` way marked `NO` carries `NO` on its portion — but
  neither de-energises anything or re-routes a feed. The drawing is the record, not a load flow.
* **A blank `NORMALLY` is drawn closed.** Closing is the assumption a radial feeder is operated
  under, and every run says so, but it is an assumption rather than a reading.
* **RMU attachment is approximate.** No RMU has a `PARENT_ID` yet, so each is drawn at its surveyed
  position and joined to the nearest defensible node with a dashed link. The run names each one
  (`links : 2 guessed, drawn dashed - RMU-GANGOTHTI-COLLEGE, RMU-KOTE-ROAD`); record the tap pole and
  the guess disappears.
* **Five ways is a fixed list.** `ODin`, `ODout`, `VL1`–`VL3` are what `VALID_CHILD_WAYS` holds; a
  two- or four-way unit would need the constant, the template ranges and §8 changed.
* **Long names are cropped** to 22 characters on the sheet. The full name stays in the workbook.
* **kVA categories in the summary table** are fixed at 25/63/100/200/250 plus any other value found
  in the data.
* **The drawings are schematic, not survey documents.** The scale bar makes the projection honest,
  but conductor lengths, angles and offsets between the line and the transformer are not surveyed.

---

## 12. Data quality to watch

The generator is deliberately forgiving, but four classes of problem are worth fixing at source.
At the time of writing:

* **Out-of-range coordinate** — `SCD-68 OLD TALLUK OFFICE` has latitude `3.926426` instead of
  `13.926426`; it is rejected and appears in `UNPOSITIONED`.
* **Shared coordinates** — `SCD-35`/`SCD-72` and `SCD-48`/`SCD-106` have identical GPS values; one
  of each pair is probably a copy-paste error. The drawing nudges them apart so both are visible.
* **Missing GPS** — `SCF-3`, `SCF-5`, `SCF-94`, `SCF-112`, `SCF-140`, `SCE-3`. They appear in the
  `UNPOSITIONED` box until filled in.
* **Node coordinates** — MUSS and RMU positions in `NON DTC NODES` are trusted as given and are not
  outlier-checked against the DTC cluster. A wrong value puts the substation box or an RMU in the
  wrong part of the city; check them once on a map.
* **Inferred parents** — in `field_data/TOPOLOGY.xlsx`, `PARENT_ID` is a minimum-spanning-tree guess
  until it is confirmed against the hand SLD. `BASIS = INFERRED` rows are the ones to check; the
  guess is a starting point, not a survey. `--list` reports the split per feeder.

Two connection-workbook rows are stale at the time of writing, and the drawings follow `NODES`
while saying so:

* **`RMU WAYS` disagrees with `NODES` for `RMU-RAMANNA-SHETTY-PARK`.** `P-001` leaves by way `VL1`,
  but the register says `VL1` is `CONNECTED = N`. Set that cell to `Y` (one of the two rows is
  wrong, and the child is the one with the cable).
* **No RMU has a `PARENT_ID` yet** — `RMU-KOTE-ROAD`, `RMU-GANGOTHTI-COLLEGE` and
  `RMU-RAMANNA-SHETTY-PARK` are each attached to the nearest node that is not downstream of them and
  drawn dashed. Record the pole each unit is fed from and the guess goes away.

---

## 13. Troubleshooting

| Symptom | Cause and fix |
|---|---|
| `Master workbook not found: …` | The workbook is not in `field_data\`. Put `DTC DETAILS U4.xlsx` there, or pass `--master` with its path. |
| `No DTCs for feeder 'X'` | The spelling does not match column F. Run `--list` and copy the name exactly. |
| Diagram is empty / one giant line | Coordinates are wrong units or swapped. Check `GPS LAT` ≈ 13.9, `GPS LONG` ≈ 75.5. |
| Changes to the workbook are ignored | The file is open in Excel with unsaved edits; save it, then rerun. |
| `preview skipped: pip install pymupdf` | PyMuPDF missing — install it, or ignore it, the `.drawio` is still written. |
| Labels overlap badly | Expected in the densest clusters; move them in draw.io. |
| MUSS box says `(position approx.)` | No MUSS row in `NON DTC NODES`, or its name does not contain the station text from column `J`. |
| MUSS box says `x km beyond the edge` | Normal when the substation is far from the load centre; use `--extent all` for a strict true-scale sheet. |
| An RMU is missing from a sheet | In a `--geo` run it comes from the workbook: check its `FEEDER` cell and that its `PARENT_ID` chain reaches this feeder. Without a workbook row it is assigned by nearest DTC, so it can land on another feeder. |
| A pole, GOS or `P-001` is missing from a sheet | Its row is not in the chain the feeder draws: set `FEEDER` on the row, or parent it under a node of this feeder. |
| A line comes off two DTCs at once | Look at `NODES!PARENT_ID` for the DTC that should hang off one of them; the drawing is only following the workbook. |
| `X disagrees with itself - nothing was drawn` | A structural error: duplicate `NODE_ID`, a `PARENT_ID` that is not a row, a loop, a bad `WAY`. The message names sheet, row and column. |
| `no PARENT_ID - its parent is guessed` | Fill in `PARENT_ID` for that row (§8). Until then the geo sheet draws a dashed link to the nearest node and the run lists it. |
| A `.drawio` will not open | Check it was not renamed to `.xml`; draw.io needs the `.drawio` extension. |
| `--topology-template` says the file already exists | It refuses to overwrite your edits. Close Excel, then use `--force` to rebuild or `--out` for a new file. |
| `Permission denied` saving `TOPOLOGY TEMPLATE.xlsx` | The workbook is open in Excel. Close it and rerun; the `~$` file beside it is the lock. |
| A generated file is not where it used to be | It is in `outputs\<today's date>\`. Use `--outdir` or `--out` to place it elsewhere. |

---

## 14. Glossary

| Term | Meaning |
|---|---|
| **DTC** | Distribution Transformer Centre — the pole-mounted transformer feeding a group of consumers; a small square with the rating inside. |
| **SCD / SCF / SCE** | DTC code prefixes, numbered per division. |
| **HT service** | High-tension consumer billed on contracted demand, drawn as a box marked `HT`. |
| **MUSS** | Main Urban Sub Station — the 110/11 kV source where the feeder starts. |
| **GOS** | Gang Operated Switch — a hand-operated isolator. Drawn as a solid bar on the line when normally closed, and as a hollow blade lifted off the line when normally open (`NORMALLY = NO`). |
| **RMU** | Ring Main Unit — a switchgear enclosure with five ways: `ODin` in, `ODout` plus `VL1`–`VL3` out. |
| **Feeder** | The 11 kV line leaving the MUSS; the subject of one sheet. |
| **Trunk** | The main run of the feeder from the MUSS; laterals branch off it. |
| **Spur / lateral** | A branch leaving the trunk. Drawn lighter than the trunk and labelled `S1`, `S2`… |
| **Parent** | The node immediately upstream of another, in the direction of the MUSS. The `PARENT_ID` column is the whole connection record. |
| **Way** | One terminal of an RMU. `ODin` is the incoming way, `ODout` and `VL1`–`VL3` are the outgoing ways. `NODES!WAY` names the way a child leaves by; a way with no child is simply empty. |
| **`.drawio`** | Plain-XML drawing file for draw.io / diagrams.net. |
