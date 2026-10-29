# Field capture — one walk, then the software takes over

**MESCOM SLD tool** · 29 October 2026 · **proposal: nothing in this document is implemented yet** ·
companion to `docs/TOPOLOGY_DATA_REQUIREMENTS.md`

Paths here are written from the repository root.

---

## 1. The gap is connectivity, not position

`--list` reports **129 node rows** in `field_data/TOPOLOGY.xlsx` across the four feeders: **114 are
`BASIS = INFERRED`** (a geometric guess) and **15 have no parent at all**. The master already holds
GPS for 113 of the 119 DTCs.

More coordinates cannot fix a single one of the four failures `TOPOLOGY_DATA_REQUIREMENTS.md` §2
lists, because **GPS locates a node; it cannot say who feeds it**. A field programme should be aimed
at the `PARENT_ID` column and treat coordinates as the optional layer.

That is the whole argument of this document, and it changes what a field app has to do: it is a
**checklist that records an order**, not a blank map that records points.

---

## 2. Three layers that get conflated

| Layer | What it is | Where it comes from |
|---|---|---|
| **Connectivity** | who feeds whom — the tree, the RMU way, the open point | the **order** nodes are met in, plus what is written on the unit door |
| **Route** | the physical path in space — spans, bends, service taps | GPS track, pole taps, or imagery |
| **Switch state** | which of five ways is cabled, `NC`/`NO`, conductor, phase | only eyes at the equipment |

A KMZ walk supplies the middle row. It supplies neither of the others, and connectivity is the row
the drawings are missing.

### 2.1 The schematic needs no geometry at all

`--geo` uses position; the schematic layout groups each subtree and prints `from SCD-36` under the
box. A field walk that records **order, branch side and pole count** is enough for the schematic —
which is the sheet submitted for approval. Geometry is what makes the geographical sheet
trustworthy, and it is a refinement on top of a correct tree, never a substitute for one.

---

## 3. Order of visit is the topology

A radial feeder walked once from the MUSS outward and back, tapping each node as you reach it,
produces an **Euler tour of the tree**. Rebuilding the parents is then a stack operation:

```python
stack = []
for n in taps:                      # node ids in the order they were tapped
    if n in stack:                  # we walked back out of a spur
        while stack[-1] != n:
            stack.pop()
    else:                           # first visit: we came from the top of the stack
        parent[n] = stack[-1] if stack else None
        stack.append(n)
```

One pass, no distance threshold, no nearest-neighbour guess. A node met twice is a backtrack, and
that repeat is exactly the information a naive "snap the nearest DTC to the track" loses.

### 3.1 Worked example

Taps in this order (each pole/DTC tapped as it is reached):

```
MUSS, P1, P2, SCD-42, P2, P3, P4, P3, P2, P1, SCD-51
```

Rebuilt parents:

| Node | Parent | Why |
|---|---|---|
| `P1` | `MUSS` | first visit, stack was `[MUSS]` |
| `P2` | `P1` | first visit |
| `SCD-42` | `P2` | first visit — the DTC taps the pole it was met at |
| `P2` (repeat) | — | backtrack: `SCD-42` was a leaf, pop to `P2` |
| `P3` | `P2` | first visit |
| `P4` | `P3` | first visit |
| `P3`, `P2`, `P1` (repeats) | — | backtracks out of the `P4` run |
| `SCD-51` | `P1` | first visit, stack had fallen back to `P1` |

Branch points and spurs then fall out with no extra data: `P1` has two children (`P2`, `SCD-51`), so
`S1` leaves `P1`; `P2` has two children (`SCD-42`, `P3`), so `S2` leaves `P2`. The trunk is the
engineer's designation, or the longest path from the MUSS as today.

### 3.2 What breaks the method, and how it shows up

| Mistake | What the rebuilt tree looks like | Guard |
|---|---|---|
| a spur walked but not tapped | its DTCs hang off the next node instead | app shows the nearest untapped DTC and asks before leaving a pole |
| two spurs walked in one pass without returning | the second spur's first node becomes the first spur's last child | backtrack tap at every branch point |
| feeder not identified at the tap | an `MRS F1` DTC appears under an `F13` node | every tap carries its feeder; the run refuses a mixed tree |
| pole numbers not recorded | branch point lands on a transformer instead of a pole | pole rows are part of the survey, not extra (`§7`) |

The tree also fails loudly rather than quietly: `validate_topology()` already errors on a duplicate
`NODE_ID`, a missing parent and a loop, and warns on a `PARENT_ID` that crosses a feeder's subtree.

---

## 4. Positioning without GPS

A phone does not need a position fix to give most of the value. In decreasing order of usefulness:

| Method | What it gives | Accuracy | Cost | Verdict |
|---|---|---|---|---|
| **Order of visit only** | the whole tree, branch points, spurs, RMU ways | exact | free | **the baseline** — no positioning at all |
| **Pole count between taps** | span length ≈ count × 40–50 m; enough for voltage drop and loss | ±10–15 % | free | **do it** — one number per DTC, how a lineman already thinks |
| **Pole number read off the pole / photo (OCR)** | exact asset identity; decides “two DTCs on one pole” vs “two poles 10 m apart” | exact | free | **do it** — settles arguments for years |
| **Imagery digitising** (offline satellite tiles, 30–50 cm) | the line traced on the image; bends for the `--geo` sheet | 1–3 m | free | **best geometry in town** — a 4 m lane defeats a 5 m GPS fix, not a 50 cm pixel |
| **QR / asset plate on poles** | identity without OCR or spelling | exact | one tagging pass, printed plates | good if poles get tagged anyway |
| **Compass + step counting (dead reckoning)** | local sketch: bearing and distance between taps | 5–20 % drift per 100 m | free | only under canopy where GPS dies; the tree does not need it |
| **AR offset measuring** | DTC plinth offset from the pole → service-tap length | 0.1–0.5 m at 10 m | free | narrow use; the tap is drawn at a standard offset anyway |
| **Landmark notes** | “30 m past the temple gate, on the left” | — | free | resolves ambiguity when two visits disagree |
| **Bluetooth RTK GNSS** | survey-grade fixes | 1–2 cm | receiver + subscription | hire for a benchmark survey; overkill for 119 DTCs |
| **Phone GPS** | confirms you are at the right DTC | 5–15 m, worse in lanes and under the line | free | confirms; does not decide |

The order of that table is the recommendation: **order + pole counts + pole numbers, imagery for the
rural runs, GPS only to confirm the DTC in front of you.**

There is also a no-phone method that beats all of them on effort for tier 1: transcribing the
hand-drawn SLD into the template, 2–4 h per feeder (`TOPOLOGY_DATA_REQUIREMENTS.md` §10). A walk is
the right tool for switch state, conductor, phase and anything the drawing gets wrong — not for
re-typing a tree that already exists on paper.

---

## 5. Geometry when there is a track

The tree comes from §3; the track only shapes spans. The rules that keep the drawing honest:

1. **Taps are control points; the track is filler between them.** A 1 Hz track does not place a line
   as well as one tap per pole, 20–30 m apart.
2. **Never bend the line to a DTC.** A transformer sits on a plinth 10–40 m off the road; the 11 kV
   runs pole to pole along the road. Project the DTC onto the route and draw the service tap — the
   model already has this (`SCD-01` parented to `P-001`). Snapping the route to the DTC is the
   “spur through a symbol” failure `TOPOLOGY_DATA_REQUIREMENTS.md` §2 warns about.
3. **A DTC 60–150 m from the route is flagged, not attached.** It is probably on another lateral, or
   another feeder.
4. **Taps win over the track.** If a tap and the track disagree by more than ~15 m, keep the surveyed
   position and print the gap — never move a node to make a line look tidy.
5. **Filter before shaping.** Drop fixes worse than ~25 m, simplify to 2–5 m, and split the track at
   every repeat visit so each leg is one chain of spans.
6. **No track is a valid submission.** Spans run straight parent-to-child, as they do today; `BENDS`
   refine them later.

---

## 6. What no sensor records

Recorded at the equipment, while someone is standing in front of it:

* **RMU ways** — which of `ODin`/`ODout`/`VL1`–`VL3` has a cable, and `NC`/`NO`. This is the only way
  `NODES!WAY` and `RMU WAYS!CONNECTED` ever agree.
* **GOS blade position** — drawn as a solid bar or a lifted blade.
* **Conductor per span** — visible from the ground: `3x95 ABC`, `3x25 ABC`, bare `Rabbit`, etc. The
  one field value that makes voltage drop computable (§9).
* **Phase** — two or three wires on the crossarm; also on the DTC nameplate.
* **Pole number** — painted on the pole.
* **A photo per node** — pole number, DTC nameplate, unit door, switch. One tap, and it settles the
  record a decade later.
* **Ties** — an RMU-to-RMU ring or tie, recorded even though a tree cannot draw it yet (§7).

These are tier 4 in `TOPOLOGY_DATA_REQUIREMENTS.md`, and they are nearly free while the walk is
happening. Postponing them means a second visit.

---

## 7. Output contract

The end of the pipeline is **`field_data/TOPOLOGY.xlsx`**, not a new format. The generator already
reads `NODES`, `RMU WAYS`, `BENDS` and `SPURS` and runs 17 structural checks over them, so a capture
pipeline that writes that workbook needs **no change to the generator** and keeps one source of
truth. Everything below is additive and backward compatible.

### 7.1 `TIES` — rings the tree cannot hold

One row per tie: `FROM_RMU`, `FROM_WAY`, `TO_RMU`, `TO_WAY`, `NORMALLY`, `NOTES`. Today a closed ring
is out of scope and survives as a `REMARKS` note; a sheet keeps the option open without touching the
one-parent rule.

### 7.2 `STATUS` and `AS_OF` — how the network evolves

Append, never edit in place:

| Column | Meaning |
|---|---|
| `STATUS` | `IN` (in service), `RETIRED`, `PLANNED` |
| `AS_OF` | the date the row took effect |
| `NODE_ID` | never reused, never deleted — a retired DTC keeps its row |

A new DTC is a new row; a re-routed span is a new `PARENT_ID` with a later `AS_OF`; a re-conductored
span is a new `CONDUCTOR` row. `--as-of 2027-03-31` then regenerates the sheet as it stood on that
date, and “what changed since last March” is a diff of two runs. One column and one flag buy all of
that.

### 7.3 Poles become part of the survey

`POLE` rows already work (`P-001` in the workbook) and are optional in the specification. In a survey
they should be **compulsory on the trunk and at every branch**, because the pole is both the tap
point and the span boundary; without it the branch point lands on a transformer.

### 7.4 `BASIS`

A walked row is `BASIS = SURVEYED`; a row read off the hand drawing stays `FROM-SLD`. The split is
the progress bar, and `--list` already reports it.

---

## 8. Tooling ladder

Do not start with the web app. Each step is useful on its own and the later ones reuse the earlier
one's output.

| Step | What | Effort | Why first |
|---|---|---|---|
| **0** | Ask the section office / MESCOM GIS for the 11 kV line layer (shapefile, GeoJSON or KML) | one phone call | if it exists, most geometry is already done and the walk is only for switch state and conductor |
| **1** | Tablet + `TOPOLOGY TEMPLATE.xlsx`, one feeder with the lineman, `PARENT_ID` filled in visit order | one day per feeder | proves the protocol before any software exists |
| **2** | `field_import.py`: GPX/KML + tap order → tree by §3, snapped geometry by §5 → `TOPOLOGY.xlsx` → `validate_topology()` → gap report | 1–2 days | this is the piece that saves the time, and it is the backend any app would call |
| **3** | The app, only if this repeats yearly across many feeders | weeks | 119 DTCs over three days does not justify it yet |

For step 1, two existing phones apps cover most of the need with no code: **QField** (QGIS on Android:
offline basemap, custom forms, photos, GNSS tracking, GeoPackage export) and **OsmAnd** (track
recording plus waypoints, GPX/KML export). Either can carry the DTC list as a layer and produce the
track and taps that step 2 consumes.

If step 3 is ever built, it needs no server for this scale: an offline PWA that preloads the feeder's
nodes and tiles, records taps, photos and a track, and exports one ZIP. A backend is only for
multi-team sync and long-term storage — which the workbook on the office PC already provides.

### 8.1 Sketch of `field_import.py` (not written)

```
read GPX/KML track + tap list (CSV or workbook sheet)
  -> project to metres, drop bad fixes, simplify
  -> rebuild tree from tap order (§3)
  -> attach DTCs to their tap pole, build BENDS from the track (§5)
  -> write TOPOLOGY.xlsx (NODES, BENDS, SPURS labels, TIES)
  -> validate_topology(); print gap report (untapped nodes, long spans, far DTCs)
```

The gap report is the deliverable that matters: every place the walk and the geometry disagree, named
by row, for a human to resolve — never silently smoothed over.

---

## 9. What this unlocks

Once parents, spans and conductor exist — `_metres()` is already in the tool — the rest is traversal,
not new field work.

### 9.1 Span lengths and per-section load

Straight from the geometry and the tree. Per-section downstream DTC count and kVA is the same walk as
`load_below()`, printed as a table; `TOPOLOGY_DATA_REQUIREMENTS.md` §12 already lists it as the next
drawing change.

### 9.2 Voltage regulation

For a span carrying `S` kVA downstream, with conductor `R`, `X` in Ω and `cos φ` the power factor:

```
three-phase, 11 kV:   ΔV_LL = (S_kVA / 11) · (R cos φ + X sin φ)  volts
                      ΔV_%  =  S_kVA · (R cos φ + X sin φ) / 1210
single-phase, 230 V:  ΔV    = 2 · (S_kVA / 0.23) · (R cos φ + X sin φ)  volts
```

Add the spans from the MUSS to each node, report the worst node and the path to it, and cap the sheet
at the limit being applied (typically 3 % for a lateral, 6 % overall — confirm the figure the section
uses). Inputs: length (geometry), conductor (field), downstream kVA (master), `cos φ` (assumption,
0.85 typical), phase (field).

Indicative 20 °C values, to be replaced by the conductor register before any number is submitted:

| Conductor | R Ω/km | X Ω/km |
|---|---|---|
| `Rabbit` (ACSR, 6/1.0) | 0.54 | 0.35 |
| `Weasel` (ACSR, 20 mm²) | 0.91 | 0.35 |
| `3x25 ABC` | 1.20 | 0.09 |
| `3x95 ABC` | 0.32 | 0.08 |

### 9.3 Loss and energy saving — an honest boundary

Technical loss per span is `3 I² R`; annual energy is that loss × `LLF` × 8760, with
`LLF ≈ 0.3 LF + 0.7 LF²`. The field gives `R` and the tree. **It does not give the load shape** —
annual energy per DTC and the feeder peak come from billing / MIS, not from a walk.

So the honest model is downstream kVA × diversity × `cos φ`, scaled to the measured feeder peak, and
then scenarios: reconductor a lateral, convert to HVDS, replace bare wire with ABC, split an
overloaded DTC. Each is the same traversal with a changed `R` or a changed tree, and the output is a
before/after loss and voltage figure per section. That scenario ability — not the first drawing — is
the real payoff of a correct tree.

### 9.4 Evolution

With `§7.2` in place, a new DTC or a re-routed span is a row, not a re-survey; every past sheet is
regenerable; and the “what changed” table is a diff between two `--as-of` runs.

---

## 10. Traps specific to this network

* **Town feeders share poles.** `MRS F1` and `Mandli F13` run on the same streets; every tap must
  carry its feeder or a DTC silently joins the wrong sheet and the walk still looks correct.
* **Rings exist physically whether or not the drawing holds them.** Capture ties into `TIES`.
* **Service tap vs line node.** Two DTCs on one pole and two poles 10 m apart are identical to GPS;
  the pole number is what decides it.
* **GPS drift in lanes and under the line** is 5–15 m, and worse between buildings. Order-of-visit
  survives it; snapping does not. This is the reason §3 exists.
* **The lineman's knowledge outranks the app.** One walk per feeder with the person who operates it
  is worth more than any positioning method in §4.
* **A wrong parent is a safety-relevant error**, not a cosmetic one (`CLAUDE.md`). The pipeline must
  refuse, not repair: a tree that fails `validate_topology()` is not written.

---

## 11. Decisions needed

1. **Is there an existing GIS line layer?** Step 0 changes the whole plan.
2. **Poles as compulsory rows** on the trunk and at branch points — yes or DTC-to-DTC only?
3. **Which feeders get walked first**, and is there a lineman available per feeder?
4. **`TIES` and `STATUS`/`AS_OF`** — adopt both, or is a tree-only workbook enough for now?
5. **Voltage limit** the section applies for a lateral and for the feeder (3 % / 6 % or otherwise).
6. **Where the load data comes from** for loss work — DTC readings, feeder energy, or both.

---

## 12. What this document does not claim

No code reads GPX or KML; there is no `TIES`, `STATUS` or `AS_OF` column; nothing calculates voltage
drop or loss; there is no field app. Today the tool draws a tree from a workbook and validates it.
Everything above is a plan for the next layer, written so that each step is useful even if the next
one never happens.

---

*Companion files: `docs/TOPOLOGY_DATA_REQUIREMENTS.md` (the field specification as it stands),
`field_data/TOPOLOGY.xlsx` (the workbook this pipeline would produce),
`template/TOPOLOGY TEMPLATE.xlsx` (the blank), `docs/README.md` (the tool manual).*
