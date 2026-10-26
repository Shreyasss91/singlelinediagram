# Topology data requirements — trunk, spurs and connections

**MESCOM SLD tool** · 26 October 2026 · companion to `README.md` and `template/TOPOLOGY TEMPLATE.xlsx`

---

## 1. Why this document exists

The tool can already draw a feeder from the GPS coordinates in `DTC DETAILS U4.xlsx`. Every DTC,
RMU and MUSS appears at its surveyed position, north up, at a uniform scale.

What the workbook does **not** contain is **connectivity**: which node feeds which. Two DTCs 40 m
apart may be on the same lateral, on two different laterals, or on different feeders. GPS cannot
tell the difference. This document defines exactly what to record, in what format, so the drawing
stops guessing.

The format is implemented in **`template/TOPOLOGY TEMPLATE.xlsx`** — already populated with every
DTC, HT service, MUSS and RMU from the master, with a first-draft connection for each.

---

## 2. What is missing today

With points only, the geographical layout builds a **minimum spanning tree** (MST): each point is
joined to its nearest already-connected neighbour, rooted at the MUSS. On a radial feeder this
usually looks right, but it is a geometric guess and fails in four predictable ways:

| Failure | What you see on the sheet | Why it happens |
|---|---|---|
| Wrong lateral | A DTC is drawn on a branch it is not on | Nearest neighbour is not the electrical parent |
| Short cut | The line crosses a block, tank or building | The MST takes the shortest geometric path, not the road |
| Spur through a symbol | A line runs through a DTC square | The trunk has no knowledge of tap points |
| Wrong order | A DTC appears upstream of a GOS or RMU that actually protects it | Devices are points, not elements in a chain |

The same missing data also blocks anything that needs the tree: downstream DTC count and kVA per
lateral, which DTCs an RMU isolates, which section a GOS splits, feeder loading by section.

**One extra column per node removes almost all of it.**

---

## 3. The data model

### 3.1 One row per node, one parent each

```
MUSS ── P-001 ── GOS-01 ── P-002 ──┬── RMU ── P-003 ── ...
         │                         │
         └── SCD-01                └── P-010 ── SCD-07    (spur S1)
```

Every physical point is a **node**. Every node except the source has exactly one **`PARENT_ID`** —
the node immediately upstream, towards the MUSS. Reading the parent column together gives the whole
tree: the trunk, every spur, and the order of equipment along each one.

This is deliberately simpler than the `NODES` + `SPANS` sketch discussed earlier. A separate span
list has to be kept in step with the node list; a parent column is one cell per DTC and cannot
disagree with itself. A tree has exactly one incoming span per node, so span attributes (conductor,
bend points) can attach to the downstream node or to the optional `BENDS` sheet.

### 3.2 Poles are nodes — and optionally may be omitted

A DTC is not usually *in* the line; it hangs off a pole by a short service. Modelling the pole gives
the correct picture:

* `P-002` has two children, `P-010` and `RMU` → `P-002` is a **branch point**
* `SCD-01` has parent `P-001` → its span is the **service tap**, not the 11 kV line

If you do not want to record poles, parent each DTC to the nearest **DTC or RMU** upstream instead.
The tree is still correct; only the branch points become transformers rather than poles. The drawing
will still separate the symbol from the line. Poles are better, but not compulsory.

### 3.3 Direction

`PARENT_ID` always points **towards the source**. Power flows parent → child. Only a `MUSS` row has
a blank parent.

### 3.4 Rings and RMUs

An RMU in a ring has two ways in service: one in, one out. Record the **incoming** way as
`PARENT_ID` and the **outgoing** way as a child of the RMU. That draws the ring as a chain through
the RMU, which is how a radial SLD is normally read.

If the ring is **closed** (both ends energised), a tree cannot represent the parallel path. Record the
tree through the RMU and add the tie as a note in `REMARKS`; the drawing will show the ring as one
path. Closed-ring operation is out of scope until the drawing needs it — tell me if it does.

---

## 4. Field specification — sheet `NODES`

One row per MUSS, RMU, GOS, pole, DTC and HT service. Columns A–N:

| Col | Field | Required | Meaning / accepted values |
|---|---|---|---|
| A | `NODE_ID` | **yes** | Unique. DTC and HT rows must use the master code (`SCD-01`, `HT-87`) so the tool can join to `ALL` / `HT`. Poles: `P-001`. MUSS / RMU: `MUSS-MRS`, `RMU-KOTE-ROAD`. |
| B | `TYPE` | **yes** | `MUSS`, `RMU`, `GOS`, `POLE`, `DTC`, `HT`, `FUSE`, `LA`, `JUNCTION` |
| C | `NAME` | yes | Text printed on the sheet. For DTC/HT, keep the master name so labels stay consistent. |
| D | `FEEDER` | for DTC/HT | `MRS F1`, `Mandli F13`, `Mandli F22`, `Mandli F3`. Blank for a shared MUSS. |
| E | `LAT` | if known | Decimal degrees. Blank is allowed for poles (see §4.1). |
| F | `LONG` | if known | Decimal degrees. Blank is allowed for poles. |
| G | `PARENT_ID` | **all except MUSS** | The node immediately upstream. Must exist in column A. |
| H | `SPUR_ID` | recommended | `T` = main run from the MUSS. `S1`, `S2`, … = laterals. Constant along a lateral. |
| I | `KVA` | DTC/HT | Rating or contracted demand. Already filled from the master. |
| J | `PHASE` | optional | `3PH` or `1PH`. Not in the current master — leave blank if unknown. |
| K | `NORMALLY` | switches | `NC` (normally closed) or `NO` (normally open). Applies to `GOS` and `RMU`. |
| L | `CONDUCTOR` | optional | Conductor of the span **from the parent to this node**, e.g. `3x95 ABC`, `3x25 ABC`, `RABBIT`. Leave blank if not recorded. |
| M | `BASIS` | yes | Where `PARENT_ID` came from: `SOURCE`, `INFERRED`, `SURVEYED`, `FROM-SLD`, `UNKNOWN`. |
| N | `REMARKS` | optional | Pole number, tap length, “ring open here”, anything the next reader needs. |

### 4.1 Blank coordinates on a pole

A pole with no latitude/longitude is still useful. The reader will place it on the straight line
between its parent and its children. Topology stays exact; only the bend of that one span is
approximate. This is the cheapest way to get correct spur grouping before a field survey.

### 4.2 Minimum viable submission

If you fill only four columns — `NODE_ID`, `TYPE`, `PARENT_ID`, `SPUR_ID` — the tree is complete and
the drawing improves immediately. Everything else is refinement.

---

## 5. Sheet `SPURS`

An auto-generated register of the laterals found in `NODES`.

| Field | Meaning |
|---|---|
| `FEEDER` | which feeder the lateral belongs to |
| `SPUR_ID` | `T`, `S1`, … as used in `NODES` |
| `LABEL` | **fill this in** — the name to print on the sheet, e.g. “Kote Road lateral” |
| `FROM_NODE` | the trunk node where the lateral leaves the main run |
| `IS_TRUNK` | `Y` for the main run, `N` for a lateral |
| `DTC_COUNT`, `KVA_TOTAL` | computed; do not edit |
| `NOTES` | optional |

The counts and kVA totals are the reason spur grouping is worth doing: they give the connected load
per lateral straight from the register, with no separate calculation.

---

## 6. Sheet `BENDS`

Optional. One row per bend on a span:

| Field | Meaning |
|---|---|
| `FROM_ID`, `TO_ID` | the span being shaped (must match `PARENT_ID`) |
| `SEQ` | 1, 2, 3 … in order from `FROM_ID` to `TO_ID` |
| `LAT`, `LONG` | the bend point |
| `NOTES` | e.g. “turns at the school corner” |

Without bends, every span is a straight line between its two nodes. With bends, the line follows the
road. A handful of bends per lateral is enough; there is no need to record every pole.

---

## 7. How the trunk and spurs are derived

Given complete parents, the grouping is mechanical:

1. **Root** the tree at the MUSS.
2. **Trunk** = the path the engineer designates as the main run. In the template it is guessed as the
   longest path from the MUSS, measured in metres.
3. **Branch point** = any node on the trunk with a child that is not the next trunk node.
4. Each off-trunk child starts a spur: `S1`, `S2`, … in order along the trunk.
5. Every node in that child's subtree carries the same `SPUR_ID`.
6. `FROM_NODE` for the spur is the branch point on the trunk.

A spur that itself branches keeps the same `SPUR_ID` in the template. If you want sub-laterals
(`S1a`, `S1b`), say so — the drawing can support them, but most MESCOM sheets do not need them.

**The guess is only a starting point.** If the real trunk follows a different road, change the
`SPUR_ID` values; the reader will use your labels, not the guess.

---

## 8. What is already in the template

`template/TOPOLOGY TEMPLATE.xlsx` was built from the current master on 26 October 2026:

| Item | Count |
|---|---|
| DTC rows | 119 |
| HT service rows | 5 |
| MUSS rows | 2 |
| RMU rows | 2 |
| **Total node rows** | **128** |
| With an inferred parent | 112 |
| Needing a manual parent (`BASIS = UNKNOWN`) | 14 |

The 14 are: 7 DTCs without usable GPS (6 missing, plus `SCD-68` rejected as an outlier), 5 HT services
(no coordinates in the master), and the 2 RMUs (their way position is not recorded anywhere).

Inferred per feeder:

| Feeder | DTCs positioned | Inferred spurs |
|---|---|---|
| MRS F1 | 62 / 63 | T + 12 laterals (`S1`–`S12`) |
| Mandli F22 | 29 / 32 | T + 6 laterals (`S1`–`S6`) |
| Mandli F13 | 17 / 19 | T + 3 laterals (`S1`–`S3`) |
| Mandli F3 | 4 / 5 | T only |

Amber `BASIS` cells are inferred and need checking. Orange `PARENT_ID` cells have no parent yet.
Green is trusted (`SURVEYED` or `FROM-SLD`).

---

## 9. Collection order and effort

| Tier | What to collect | Effort | What it unlocks |
|---|---|---|---|
| **0 — done** | GPS for DTC, RMU, MUSS | — | symbols at true positions |
| **1 — do this first** | `PARENT_ID` + `SPUR_ID` for every DTC | 2–4 h per feeder, reading the existing SLD | correct trunk, correct laterals, no short cuts, downstream kVA per spur |
| **2** | pole rows and pole coordinates | 1 field day per feeder | branch points in the right place, true span lengths |
| **3** | `BENDS` for spans that do not run straight | 1–2 h per feeder from Google Earth | line follows roads instead of crossing blocks |
| **4** | `CONDUCTOR` per span, `NORMALLY` on GOS/RMU | from the register | conductor register on the drawing, open points, later load flow |
| **5** | `PHASE` per DTC | from the register | 1φ/3φ symbols and phase-wise totals |

Tier 1 is the whole ask. Tiers 2–5 can follow over time; the template is designed so you never have
to redo tier 1.

### How to do tier 1 quickly

1. Open the hand-drawn SLD (`MRS F1 SLD.docx` / PDF) beside the template.
2. Filter `NODES` to one feeder.
3. For each DTC, read which lateral and which upstream node it sits on. Type that node into
   `PARENT_ID`.
4. Set `SPUR_ID` to `T` or `S1`… as on the drawing.
5. Do not worry about poles yet: parent each DTC to the upstream DTC if that is what the drawing
   shows. Poles can be inserted later as rows between them.
6. Change `BASIS` from `INFERRED` to `FROM-SLD` as you go, so progress is visible.

The template refuses to be overwritten by accident; when you want to rebuild from a newer master,
run `--topology-template --force` and your edits will be discarded.

---

## 10. Where to get the data

| Source | Covers | Notes |
|---|---|---|
| **Existing hand-drawn SLD** (`.docx` / PDF) | tiers 1, 3, 4 | Already the authority for submission. The topology is visibly drawn; it only needs transcribing. Start here. |
| **Section office 11 kV feeder diagram / feeder register** | tiers 1, 4 | Usually maintained for outage planning; often has pole numbers and GOS. Ask for the current signed copy. |
| **MESCOM GIS** | tiers 1, 2, 3 | If the DTCs were GPS-tagged, a line layer or DTC→pole mapping may already exist. A shapefile / GeoJSON / KML of the 11 kV line answers tiers 1–3 at once. |
| **Google Earth** | tier 3 | Trace each lateral as a path, export KMZ. Good enough for bends; no electrical meaning. |
| **Field walk / field app** | tiers 1–5 | Highest accuracy, highest cost. Record pole number, parent pole, GOS state, conductor. One feeder per day. |

Do **not** send consumer-level data (`DTC TAGGING.xlsx`, `MASTER_FILE_RR_NUMBERS_DTCwise.xlsx`).
Service connections and meter data do not affect the feeder diagram.

---

## 11. What changes in the output once this is supplied

When the tool reads a completed topology workbook:

* trunk drawn as one continuous heavy line, following the real run and its bends
* each lateral drawn as its own lighter line, branching from a marked junction
* DTC symbol at its surveyed position, joined by a short service tap — not projected onto the line
* spur labels (`S1 Kote Road`, …) and per-spur DTC count / kVA in the summary table
* GOS drawn inline on the correct span, open or closed as recorded
* RMU drawn with its incoming and outgoing ways, so the ring reads correctly
* `UNPOSITIONED` box still holds anything without coordinates — nothing disappears
* downstream count and kVA per section available for the feasibility reports

None of this changes the schematic layout, which already ignores geography.

---

## 12. Validation the tool will run

Before drawing, the topology reader should reject or flag:

1. duplicate `NODE_ID`
2. `PARENT_ID` that does not exist
3. a node that is its own ancestor (cycle)
4. more than one `MUSS` root per feeder (unless a genuine second source)
5. a node with no parent that is not a MUSS
6. two parents for the same node (impossible in this format — one cell only — but a duplicated row
   with the same `NODE_ID` catches it)
7. coordinates outside the Shivamogga envelope (lat ≈ 13.9, lon ≈ 75.5)
8. a DTC/HT code that does not exist in the master
9. `SPUR_ID` used on two different feeders without a `FEEDER` value
10. a `BENDS` row whose span does not exist in `NODES`

Errors stop the run with the row number; warnings are printed and the drawing continues.

---

## 13. Worked example

The `EXAMPLE` sheet is a small complete feeder. The essential rows:

| NODE_ID | TYPE | PARENT_ID | SPUR_ID | Reading |
|---|---|---|---|---|
| `MUSS-DEMO` | MUSS | *(blank)* | `T` | source |
| `P-001` | POLE | `MUSS-DEMO` | `T` | trunk |
| `GOS-01` | GOS | `P-001` | `T` | switch in the trunk |
| `SCD-01` | DTC | `P-001` | `T` | taps the pole, not the line |
| `P-002` | POLE | `GOS-01` | `T` | branch point |
| `P-010` | POLE | `P-002` | `S1` | first pole of the lateral |
| `SCD-07` | DTC | `P-010` | `S1` | on the lateral |
| `RMU-DEMO` | RMU | `P-002` | `T` | way in from `P-002`, way out to `P-003` |
| `P-003` | POLE | `RMU-DEMO` | `T` | continues beyond the RMU |

`P-002` has two children (`P-010` and `RMU-DEMO`) — that single fact is what makes `S1` a lateral.
No separate “spur” table is needed to express it.

---

## 14. Decisions needed from you

1. **Which run is the trunk** on each feeder? The template assumes the longest path from the MUSS.
   If the main road is a different run, say which, or set `SPUR_ID` by hand.
2. **Do you want poles as rows**, or is DTC-to-DTC parenting enough for now? Poles give correct
   branch points; DTC-only is faster.
3. **Are any GOS normally open?** If yes, which, and does the drawing need to show the open point
   differently?
4. **Are the RMUs in a closed ring or a radial chain?** If closed, the parallel path needs a
   separate tie representation.
5. **Is `PHASE` recorded anywhere** (1φ/3φ per DTC)? It is not in the current master.

---

## 15. Quick checklist

```
[ ] Copy template/TOPOLOGY TEMPLATE.xlsx to a working name (keep the blank one)
[ ] Filter NODES to one feeder
[ ] Correct PARENT_ID for every DTC   (the one thing that matters)
[ ] Set SPUR_ID = T / S1 / S2 ...
[ ] Change BASIS to FROM-SLD as each row is confirmed
[ ] Add pole rows where the branch point matters
[ ] Add BENDS only where a line must follow a road
[ ] Fill SPURS.LABEL with the names used on the hand drawing
[ ] Check the §12 rules, then send the file back
```

Rebuild the blank template from the current master at any time:

```bash
python sld_generator.py --topology-template          # refuses to overwrite a filled file
python sld_generator.py --topology-template --force  # rebuild, discards edits
```

---

*Companion files: `template/TOPOLOGY TEMPLATE.xlsx` (the data), `template/route_template.csv`
(single-chain `--route` fallback), `README.md` (tool manual), `CHANGELOG.md` (change history).*
