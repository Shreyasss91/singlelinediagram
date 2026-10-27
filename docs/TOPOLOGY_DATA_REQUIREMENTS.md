# Topology data requirements — trunk, spurs and connections

**MESCOM SLD tool** · 26 October 2026 · companion to `docs/README.md` and `template/TOPOLOGY TEMPLATE.xlsx`

Paths here are written from the repository root.

---

## 1. Why this document exists

The tool can already draw a feeder from the GPS coordinates in `field_data/DTC DETAILS U4.xlsx`. Every DTC,
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
MUSS ── P-001 ── GOS-01 ── P-002 ──┬── RMU ──ODout── P-003 ── ...
         │                         │
         └── SCD-01                └── P-010 ── SCD-07    (spur S1)
```

The link `P-002 → RMU` is way `ODin`; the link `RMU → P-003` is way `ODout`. See §3.4.

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

### 3.4 RMUs are five-way units

Every RMU in this network is a **five-way switchgear**: one way in and four ways out.

| Way | Direction | Use |
|---|---|---|
| `ODin` | in | the incoming cable — the RMU is fed through this way, always cabled |
| `ODout` | out | one of the four outgoing ways; often the continuing run, but it may be empty |
| `VL1` | out | outgoing way 1 — may be empty |
| `VL2` | out | outgoing way 2 — may be empty |
| `VL3` | out | outgoing way 3 — may be empty |

So an RMU is **not** a single two-terminal device in a chain: it is one node with five terminals, and
anything downstream of it leaves by a named way. Two things are recorded, in two places:

1. The RMU is **one row** in `NODES`. Its own `PARENT_ID` is the node that feeds it — that link *is*
   way `ODin`, so no `WAY` value is written on the RMU row.
2. Every **child** of the RMU is a normal row whose `PARENT_ID` is the RMU, plus a `WAY` value naming
   the way it leaves by: `ODout`, `VL1`, `VL2` or `VL3`.

That keeps exactly one parent per node — the rule the whole format rests on — while still saying which
terminal the cable lands on.

```
                  ┌── ODout ── (empty)
P-002 ── ODin ──[RMU]── VL1 ── SCD-08
                  ├── VL2 ── P-003 ── ...
                  └── VL3 ── (empty)
```

#### A five-way unit is not five connections

The five ways are **terminals, not cables**. A unit can sit in service with only two of them used:
`ODin` in and a single outgoing way out. The rest are simply empty — no cable, no child row, nothing
downstream, and nothing missing from the model.

* Only `ODin` is always cabled. Every other way, **`ODout` included, may be empty**.
* `ODout` is often the continuing run to the next unit or pole, but there is no rule that says so.
  Plenty of units continue from `VL1`, `VL2` or `VL3`; on some, `ODout` sits empty for years.
* An empty way is recorded by its absence: no child row names it in `NODES!WAY`, and it carries
  `CONNECTED = N` on the `RMU WAYS` sheet (§5).
* A way with no cable has no switch state worth recording, so `NORMALLY` stays blank on it. `NC`/`NO`
  is only meaningful where something is connected.

`NODES!WAY` is the connection; the `RMU WAYS` sheet (§5) is the register of the five ways themselves —
direction, whether a cable is landed, normal state, label. A way can carry a child without appearing
on that sheet and vice versa, but the two are meant to agree, and the validator (§13) checks it.

When an RMU is used as a through unit in a ring, two ways are in service: `ODin` in and one outgoing
way out (commonly `ODout`, but any of the four). Recorded this way, the ring reads as a chain through
the RMU, which is how a radial SLD is normally drawn. If the ring is **closed** (both ends energised),
a tree cannot represent the parallel path: record the tree through the RMU and note the tie in
`REMARKS`. Closed-ring operation is out of scope until the drawing needs it — tell me if it does.

Ways that are **not** in use still exist physically, so they keep a row. They are marked
`CONNECTED = N` and listed as spare; mark one `NO` only if it is a real open point with a cable on it.

---

## 4. Field specification — sheet `NODES`

One row per MUSS, RMU, GOS, pole, DTC and HT service. Columns A–O:

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
| I | `WAY` | child of an RMU | Which RMU way this node leaves by: `ODout`, `VL1`, `VL2`, `VL3`. Blank on every other row — including the RMU row itself, whose incoming way is `ODin` and is expressed by its `PARENT_ID`. A way with no child is simply empty; see §3.4. |
| J | `KVA` | DTC/HT | Rating or contracted demand. Already filled from the master. |
| K | `PHASE` | optional | `3PH` or `1PH`. Not in the current master — leave blank if unknown. |
| L | `NORMALLY` | switches | `NC` (normally closed) or `NO` (normally open). Applies to `GOS` and `RMU`. |
| M | `CONDUCTOR` | optional | Conductor of the span **from the parent to this node**, e.g. `3x95 ABC`, `3x25 ABC`, `RABBIT`. Leave blank if not recorded. |
| N | `BASIS` | yes | Where `PARENT_ID` came from: `SOURCE`, `INFERRED`, `SURVEYED`, `FROM-SLD`, `UNKNOWN`. |
| O | `REMARKS` | optional | Pole number, tap length, “ring open here”, anything the next reader needs. |

### 4.1 Blank coordinates on a pole

A pole with no latitude/longitude is still useful. The reader will place it on the straight line
between its parent and its children. Topology stays exact; only the bend of that one span is
approximate. This is the cheapest way to get correct spur grouping before a field survey.

### 4.2 Minimum viable submission

If you fill only four columns — `NODE_ID`, `TYPE`, `PARENT_ID`, `SPUR_ID` — the tree is complete and
the drawing improves immediately. Everything else is refinement; `WAY` becomes necessary only on the
children of an RMU, because otherwise the tool cannot tell which of the five terminals a cable lands
on.

---

## 5. Sheet `RMU WAYS`

The register of the switchgear itself: **one row per way, five rows per RMU**. It is pre-filled for
every RMU in the master.

| Field | Meaning |
|---|---|
| `RMU_ID` | the `NODE_ID` of the RMU row in `NODES` |
| `WAY` | `ODin`, `ODout`, `VL1`, `VL2`, `VL3` |
| `DIRECTION` | `IN` for `ODin`, `OUT` for the other four |
| `CONNECTED` | `Y` if a cable is landed on the way, `N` if the way is empty. `ODin` is always `Y`; the other four start `N` |
| `NORMALLY` | `NC` (closed) or `NO` (open) where a cable exists; blank on an empty way, which has no state worth recording |
| `LABEL` | **fill this in if the unit is labelled** — e.g. “Kote Road out”, “spare” |
| `NOTES` | cable size, fuse rating, which feeder it ties to |

`CONNECTED` is the column that makes “not all ways are used” explicit. The default is honest rather
than optimistic: one cable in, four ways empty, because nothing in the master records a way yet. Set a
way to `Y` only when a cable is really landed on it — and then a child row in `NODES` should name that
way in `WAY`. The two are cross-checked by the validator (§13):

```
CONNECTED = Y  <->  exactly one child row in NODES carries WAY = that way
CONNECTED = N  <->  no child row carries it
```

Why a separate sheet instead of five columns on the RMU row:

* the **connections** stay in `NODES`, one parent cell per child — they cannot disagree with
  themselves, and a way can be described without being wired to anything yet;
* a way carries attributes that belong to the switchgear, not to the node downstream of it (whether a
  cable exists, normal state, label, fuse);
* an **empty way** has no child at all, yet still exists — it has nowhere to live in `NODES`, and a
  register that simply omitted it would be indistinguishable from a way nobody has checked yet.

The sheet is a description, not a second copy of the tree: a child named on this sheet is still
parented only once, in `NODES`.

---

## 6. Sheet `SPURS`

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

## 7. Sheet `BENDS`

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

## 8. How the trunk and spurs are derived

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

## 9. What is already in the template

`template/TOPOLOGY TEMPLATE.xlsx` was built from the current master on 26 October 2026:

| Item | Count |
|---|---|
| DTC rows | 119 |
| HT service rows | 5 |
| MUSS rows | 2 |
| RMU rows | 2 |
| **Total node rows** | **128** |
| RMU way rows (`RMU WAYS`) | 10 |
| With an inferred parent | 112 |
| Needing a manual parent (`BASIS = UNKNOWN`) | 14 |

The 14 are: 7 DTCs without usable GPS (6 missing, plus `SCD-68` rejected as an outlier), 5 HT services
(no coordinates in the master), and the 2 RMUs (their way position is not recorded anywhere).

The 10 RMU-way rows are the five ways of each of the two units: `ODin` cabled, the four outgoing
ways marked empty (`CONNECTED = N`, no `NORMALLY`), because no cable is recorded yet. No `WAY` value is
pre-filled in `NODES`: neither RMU has a recorded child, and the minimum-spanning-tree guess never puts
a DTC under an RMU.

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

## 10. Collection order and effort

| Tier | What to collect | Effort | What it unlocks |
|---|---|---|---|
| **0 — done** | GPS for DTC, RMU, MUSS | — | symbols at true positions |
| **1 — do this first** | `PARENT_ID` + `SPUR_ID` for every DTC | 2–4 h per feeder, reading the existing SLD | correct trunk, correct laterals, no short cuts, downstream kVA per spur |
| **2** | pole rows and pole coordinates | 1 field day per feeder | branch points in the right place, true span lengths |
| **3** | `BENDS` for spans that do not run straight | 1–2 h per feeder from Google Earth | line follows roads instead of crossing blocks |
| **4** | `CONDUCTOR` per span, `NORMALLY` on GOS/RMU, `WAY` on every RMU child | from the register | conductor register on the drawing, open points, which terminal each cable lands on, later load flow |
| **5** | `PHASE` per DTC | from the register | 1φ/3φ symbols and phase-wise totals |

Tier 1 is the whole ask. Tiers 2–5 can follow over time; the template is designed so you never have
to redo tier 1. The two units only need tier 4's `WAY` work if you care which terminal each cable
lands on — the drawing already labels them as five-way units.

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

## 11. Where to get the data

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

## 12. What changes in the output once this is supplied

**This is now built.** With rows in the workbook, a `--geo` run is drawn from them:

| Supplied | What the sheet does |
|---|---|
| `PARENT_ID` | each node is joined to its parent and to nothing else; a node parented to `SCD-36` no longer appears fed from a neighbour that happened to be nearer |
| `SPUR_ID` | the trunk (`T`) is drawn heavier than a lateral (`S1`, `S2`…) |
| `TYPE` | the symbol follows it: dot for a `POLE`, bar on the line for a `GOS`, box for an `RMU`, square with the kVA for a `DTC`, white square for an `HT`, source box for the `MUSS` |
| `NODES!WAY` | the RMU box names the ways its cables arrive on (`5-way - out VL1`), so the ring can be followed off the sheet |
| `BENDS` | that span is routed through the corners instead of run straight |
| `SPURS!LABEL` | printed at the branch point (`S1 Chor Bazar lateral`) |
| `LAT`/`LONG` on a non-DTC row | the pole, GOS or RMU appears, which the master alone cannot show |
| `NORMALLY` on a `GOS` | the switch is drawn open or closed as recorded, and the run names the DTCs and kVA standing behind an open point |
| a gap | `UNPOSITIONED` still holds anything without coordinates, now naming the parent it hangs off — nothing disappears |

Still open, in the order the data pays for it:

* **per-spur DTC count / kVA in the summary table** — today the table is the kVA histogram of the
  whole feeder; a per-lateral breakdown is a drawing change, not a data one.
* **five drawn RMU terminals** — the unit is one symbol with its cabled ways named, so an empty way is
  simply a way nobody names rather than a visible spare terminal.
* **`RMU WAYS!NORMALLY` is not drawn** — a way marked `NO` is warned about and the load behind it is
  named, but the terminal itself does not yet show the state. Part of the same step as the terminals.
* **downstream count and kVA per section**, for the feasibility reports — a `GOS` reports its own load
  now (`load_below()`), but there is no per-section table on the sheet.
* **the schematic layout** still draws DTC/HT/RMU in a fixed grid and ignores all of this.

---## 13. Validation the tool runs

`validate_topology()` runs these on the rows of the feeder being drawn, before anything is drawn.

**Errors — the run stops, naming sheet, row and column:**

1. duplicate `NODE_ID` (the same node defined twice — the format itself makes two parents impossible)
2. `PARENT_ID` that is not a row in the workbook
3. a node that is its own ancestor (a loop in the chain)
4. a `WAY` value on a node whose parent is not an `RMU`
5. a `WAY` value that is not one of `ODout`, `VL1`, `VL2`, `VL3` — `ODin` belongs to the RMU row,
   not to its children
6. two children of one RMU claiming the same way (physically impossible — one cable per terminal)

**Warnings — printed, drawing continues:**

7. a node with no `PARENT_ID` that is not a MUSS: the parent is guessed from the nearest node that is
   not downstream of it, and that link is drawn dashed
8. a parent with no coordinates: the line is drawn to the nearest ancestor that has them, and the
   skipped node is named
9. `CONNECTED = N` on a way that a child does name — the register is stale, the cable is drawn
   (see §8: `NODES` is the tree, `RMU WAYS` is a register)
10. `CONNECTED = Y` on a way with no child naming it — is a node missing?
11. `CONNECTED = N` with `NORMALLY` filled in — an empty way has no switch state
12. a way marked `NORMALLY = NO` with a child recorded downstream — the load is off that way
13. no `RMU WAYS` rows at all for a unit that has children in `NODES`
14. coordinates outside the Shivamogga envelope (lat 13.5–14.3, lon 75.1–76.0), or only one of
    `LAT`/`LONG` filled: treated as unpositioned and listed in `UNPOSITIONED`
15. a DTC/HT code that does not exist in the master — the register and the topology disagree
16. a `BENDS` row whose span does not exist in `NODES` (the bend is ignored)
17. a DTC or RMU that only the master holds: drawn, parent guessed, and listed

An empty way is **not** an error. Only `ODin` must be `CONNECTED = Y`; the other four are commonly
empty, `ODout` included.

Still not checked, because nothing in the format carries the fact: a genuine second source on one
feeder (item 4 above is moot while there is one `MUSS` per chain), and `SPUR_ID` reused across
feeders (each run only ever sees one feeder's subtree).

---

## 14. Worked example

The `EXAMPLE` sheet is a small complete feeder. The essential rows:

| NODE_ID | TYPE | PARENT_ID | SPUR_ID | WAY | Reading |
|---|---|---|---|---|---|
| `MUSS-DEMO` | MUSS | *(blank)* | `T` | | source |
| `P-001` | POLE | `MUSS-DEMO` | `T` | | trunk |
| `GOS-01` | GOS | `P-001` | `T` | | switch in the trunk |
| `SCD-01` | DTC | `P-001` | `T` | | taps the pole, not the line |
| `P-002` | POLE | `GOS-01` | `T` | | branch point |
| `P-010` | POLE | `P-002` | `S1` | | first pole of the lateral |
| `SCD-07` | DTC | `P-010` | `S1` | | on the lateral |
| `RMU-DEMO` | RMU | `P-002` | `T` | | fed from `P-002` — that link is way `ODin` |
| `P-003` | POLE | `RMU-DEMO` | `T` | `VL2` | the run out uses `VL2`, not `ODout` |
| `SCD-08` | DTC | `RMU-DEMO` | `T` | `VL1` | leaves the RMU by way `VL1` |

And the matching `RMU WAYS` rows for that unit — five terminals, three cables:

| RMU_ID | WAY | DIRECTION | CONNECTED | NORMALLY | Reading |
|---|---|---|---|---|---|
| `RMU-DEMO` | `ODin` | IN | Y | `NC` | fed from `P-002`; no child names this way |
| `RMU-DEMO` | `ODout` | OUT | N | *(blank)* | empty — no cable on this way |
| `RMU-DEMO` | `VL1` | OUT | Y | `NC` | carries `SCD-08` |
| `RMU-DEMO` | `VL2` | OUT | Y | `NC` | carries `P-003` |
| `RMU-DEMO` | `VL3` | OUT | N | *(blank)* | empty spare way |

`P-002` has two children (`P-010` and `RMU-DEMO`) — that single fact is what makes `S1` a lateral.
No separate “spur” table is needed to express it. The RMU's five rows describe terminals, not
connections: `P-003` appears once, parented to `RMU-DEMO` with `WAY = VL2`. Two of the five rows carry
no cable at all, and that is a valid, complete record — the unit is in service with three cables.

---

## 15. Decisions needed from you

1. **Which run is the trunk** on each feeder? The template assumes the longest path from the MUSS.
   If the main road is a different run, say which, or set `SPUR_ID` by hand.
2. **Do you want poles as rows**, or is DTC-to-DTC parenting enough for now? Poles give correct
   branch points; DTC-only is faster.
3. **Are any GOS normally open?** If yes, which, and does the drawing need to show the open point
   differently?
4. **Are the RMUs in a closed ring or a radial chain?** If closed, the parallel path needs a
   separate tie representation.
5. **Is `PHASE` recorded anywhere** (1φ/3φ per DTC)? It is not in the current master.
6. **Which way does each cable use** on the two RMUs — `ODout`, `VL1`, `VL2` or `VL3`? Nothing in the
   master records it; `NODES!WAY` is where it goes once someone reads it off the unit.
7. **Which ways are empty** on each unit, and are any cabled ways normally open? The template assumes
   one cable in and four empty ways, because nothing is recorded yet; correct it from the unit door.

---

## 16. Quick checklist

```
[ ] Copy template/TOPOLOGY TEMPLATE.xlsx to `field_data\TOPOLOGY.xlsx` (keep the blank one)
[ ] Filter NODES to one feeder
[ ] Correct PARENT_ID for every DTC   (the one thing that matters)
[ ] Set SPUR_ID = T / S1 / S2 ...
[ ] Change BASIS to FROM-SLD as each row is confirmed
[ ] Add pole rows where the branch point matters
[ ] Add BENDS only where a line must follow a road
[ ] Fill SPURS.LABEL with the names used on the hand drawing
[ ] Put a WAY on every child of an RMU, and set RMU WAYS: CONNECTED = Y only where a cable is landed
[ ] Leave NORMALLY blank on the empty ways - most units use two or three of the five
[ ] Check the §13 rules, then send the file back
```

Rebuild the blank template from the current master at any time:

```bash
python sld_generator.py --topology-template          # refuses to overwrite a filled file
python sld_generator.py --topology-template --force  # rebuild, discards edits
```

---

*Companion files: `template/TOPOLOGY TEMPLATE.xlsx` (the data), `template/route_template.csv`
(single-chain `--route` fallback), `docs/README.md` (tool manual), `docs/CHANGELOG.md` (change history).*
