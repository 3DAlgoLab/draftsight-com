---
name: draftsight-com
description: Draw, dimension, hatch, and export in DraftSight CAD by driving the running app over COM from the Python kernel (win32com). Use when asked to automate DraftSight, script or generate a DWG from measurements or a sketch, draw a tapped hole from a DP callout, dimension a drawing, save DWG, produce a PNG/SVG proof, or when a DraftSight problem involves dimension text size, linetype dash length, entity layers/linestyles, document save/close, or COM automation speed. Windows only.
compatibility: Windows with DraftSight installed (DraftSight.Application ProgID registered) and a licensed user session; pywin32 (win32com) in the Python environment.
metadata:
  app: DraftSight
  api: COM IDispatch through pywin32 (typelib-bound gen_py wrappers)
---

# DraftSight over COM

One persistent COM connection to the running DraftSight app. Geometry is pushed
in-process, so a full drawing (layers, polylines, hatches, dimensions, leaders,
DWG save, image export) takes a couple of seconds instead of one process launch
per call, which is what a PowerShell-per-call driver pays.

Everything below is verified against a live DraftSight session, including the
places where the API returns success while doing nothing.

## Setup

`import draftsight_com` must succeed **in the interpreter that runs this session's
code**, which is usually not the first `python` on `PATH`. Check it before drawing:

```python
import draftsight_com, sys; print(draftsight_com.__file__, sys.executable)
```

If that raises, install into that exact interpreter — `pip` may be absent from it (a uv-created venv
has none), so prefer `uv pip --python`. When the code runs inside an agent kernel, ask that kernel
which interpreter it is instead of guessing from `PATH`:

```python
import sys, os
print(sys.executable)
print(os.environ.get("PI_PYTHON", ""))
```

```
uv pip install --python "<that path>" --editable "<this skill directory>"
```

Install into the kernel's interpreter and nothing else. A copy sitting in some *other* Python makes
the check above pass while the cell still fails, and that is the one failure mode worth refusing to
tolerate. This package ships no agent-side installer: `SKILL.md` is instructions, the code ships as
`draftsight_com`.
DraftSight must be reachable: either already running, or startable by `connect`.
Windows only: every call here is COM through `pywin32`.

## Quickstart

```python
import draftsight_com as ds

app = ds.connect(timeout=90)      # attach; start the app and wait if not running
ds.close_document(app, OUT_DWG)   # only needed on a rerun - see save rules below
sh = ds.new_document(app)         # from standardiso.dwt, ScaleUnit = mm

sh.drafting_style(text_height=14.0, arrow_size=6.0)   # BEFORE inserting anything
sh.layer("OUTLINE", 7); sh.layer("DIM", 3); sh.layer("CENTER", 1, "CENTER")

sh.polyline([(0, 0), (40, 0), (40, 120), (0, 120)], closed=True, layer="OUTLINE")
sh.hatch([[(0, 0), (40, 0), (40, 120), (0, 120)]], scale=2.0, layer="OUTLINE")
sh.line((0, -10), (0, 130), layer="CENTER", linetype="CENTER")
sh.dimension((0, 120), (40, 120), (0, 140), 0.0, ds.DIAMETER + "40", layer="DIM")
sh.leader([(10, 0), (50, -30), (60, -30)], 160.0, "M20 TAP,DP35", layer="DIM")

sh.zoom_fit()
sh.save(OUT_DWG)
sh.export_png(OUT_PNG)            # proof image; use it to verify, do not guess
```

Always finish a drawing by looking at the export with `attach_image`. Rendering
is the only honest check: COM returns success for operations that change nothing.

## Document maker

`build(sheet, spec)` draws a declarative spec. `make(spec, out, png=...)` runs the whole
pipeline: connect, new document, release a stale target, build, zoom-fit, save, export.
`run(spec, out, png, template)` is the same from a spec file, and it is also the
`draftsight_com` shell command:

```
draftsight_com --spec examples\simple_section_spec.py --out C:\temp\bracket.dwg --png C:\temp\bracket.png
```

```python
import draftsight_com as ds
sh = ds.make(SPEC, OUT_DWG, png=OUT_PNG)   # returns the Sheet, ready for more edits
```

A spec is a plain dict whose keys map onto the client methods:

```python
SPEC = {
  "style":   {"text_height": 5.0, "arrow_size": 2.5},      # drafting_style kwargs
  "layers":  [("OUTLINE", 7), ("CENTER", 1, "CENTER")],     # (name, color[, linetype])
  "entities": [
    {"polyline": pts, "closed": True, "layer": "OUTLINE"},
    {"line": [a, b], "layer": "CENTER", "linetype": "CENTER"},
    {"hatch": [loop], "pattern": "ANSI31", "scale": 1.5, "layer": "HATCH"},
    {"note": [x, y, height, "SECTION B-B"], "layer": "TEXT"},
    {"leader": {"points": [...], "width": 70.0, "text": "BREAK EDGE"}, "layer": "TEXT"},
    {"dimension": {"p1": p, "p2": p, "at": p, "angle": 0.0, "text": "\u00d823"}, "layer": "DIM"},  # literal diameter sign - see rule 5
  ],
}
```

Every entity needs `layer`, and that layer must appear in `layers` too; an unknown entity
key or a missing layer raises naming the entity index. `examples\simple_section_spec.py` is
a complete worked spec (90 x 90 x 12 L-bracket: outline, hatch, two centerlines, two
dimensions, leader, note) that renders as described.

## A tapped hole in one call

`tapped_hole(sheet, "M30", 55)` draws the DP-family representation - section plus top view -
from a thread size and a thread depth, and returns the numbers it used. `tapped_hole_spec(size,
depth)` is the same output as a plain spec dict, computed with no COM, so it feeds `make`/`build`,
prints for review, or gets asserted in a test:

```python
import draftsight_com as ds
ds.make(ds.tapped_hole_spec("M30", 55), OUT_DWG, png=OUT_PNG)
n = ds.tapped_hole(sh, "M30", 55)   # -> tap_drill 26.5, point_depth 8, drilled_depth 63
```

The layout is measured off the drawn family (`M30DP45`, `M30DP55`): top view at `at`, section axis
75 mm to its left, both views centred vertically on the drilled depth, tap drill = major - pitch,
118-degree point rounded to the millimetre, ANSI31 hatch stopping at thread depth, major-diameter
arc drawn 90 to 360 degrees in the top view. `references/how-to-draw-hole.md` carries the number
table and the two places where the family departs from ISO.

## Six rules that cost the most to learn

1. **Sizing comes from drafting variables, not dimension styles.**
   `IDimensionStyle` text/arrow/extension option groups accept writes and read
   back the new numbers but never affect rendered entities; per-entity
   `LineScale` does nothing. `Sheet.drafting_style()` writes the document
   variables that actually work (verified 1:1 mm). Set them before inserting.
2. **`SaveAs2` into a path another open document owns fails silently and still
   returns success.** `Sheet.save()` proves the write with `GetPathName()` plus
   mtime. On a rerun, release the old document with `close_document()` first.
3. **Never close the ACTIVE document.** `CloseDocument` on the active document
   takes the COM server down (the next call faults, the process dies). Create
   the new document first, then close the stale one; `close_document()` refuses
   when the target is active, and `activate(app, path)` moves another document to
   the foreground when you need to release the one that is open.
4. **Style entities through the entity helper.** `IEntityHelper.SetLayer` and
   `SetLineStyle` (exposed as `Sheet._style`) work on every entity type; the
   `Layer`/`LineStyle` properties on returned dimension, hatch, and leader
   objects are not writable under late binding.
5. **Text arguments are literal: `%%c` never becomes a diameter sign.** The dimension text
   override and the `InsertLeader` / `InsertSimpleNote` text reach the entity unchanged, so
   `"%%c26.5"` draws as `%c26.5`. Pass the character (`ds.DIAMETER`, U+00D8). A leader shelf
   narrower than its text *wraps* the note onto two lines, so size it from the text:
   `NOTE_WIDTH_FACTOR` (0.83 x note height per character, measured).
6. **Read a drawing back through layers and exports, not selection.** `Sheet.layers()`
   (`GetLayerManager().GetLayers()`; `IColor` answers `GetColorIndex()`, it has no `ColorValue`)
   and `Sheet.house_style()` return a drawing's layers and drafting sizes exactly. The point-taking
   `SelectByWindow` / `SelectByPoint` members exist but every call shape tried raises
   `E_INVALIDARG`, so plan read-back around layers, SVG and PNG.

## Measurement and diagnosis

`references/draftsight-api.md` holds the verified API map, the drafting-variable
table with numeric proof, the COM quirks list, the crash modes, and the
SVG/pixel rigs that measure rendered text height and dash length exactly. Read
it before debugging anything that "should" work through the style objects. Selection sets and per-entity length
(`GetSelectionManager` plus `GetLength()`) are documented there too, including the
selector argument that returns nothing unless it is `1`, and why selection cannot enumerate
an existing drawing.

To match a new drawing to an existing one, read the old one: `sh = ds.open_document(app, ref)`
then `sh.house_style()` and `sh.layers()` are exactly what `drafting_style(**kwargs)` and
`layer(name, color, linetype)` take back. In an exported SVG, arcs and arrowheads arrive as
`<path>` elements while hatch lines arrive as `<polyline>` groups stroked in the layer colour -
scan both before concluding an entity is missing.

`references/how-to-draw-hole.md` is the other note in that folder, and it answers a different
question: not how the API behaves but what the geometry should be. Tapped-hole section and top
view conventions (major diameter thick, minor/tap drill thin, drill point), the callout style
machinists expect, and the ISO 261/724 tap drill and depth tables. Read it when a drawing asks
for a hole, before deciding which circles and lines to insert.

## Official reference: a search target, never a reading target

The installed CHMs unpack into `docs/` - **6,457 pages / 63 MB, 12 s with 7-Zip** (`draftsight_com docs`
falls back to `hh.exe` when 7-Zip is absent). The tree is gitignored and not shipped, so generate it once
per machine. Pages are 1-4 KB, which sounds readable and is not: nothing tells you which of 6,457 pages
you want, so opening pages is how context gets destroyed. Every access is a search, and `ds-docs lookup`
is that search: it resolves the page **by filename** first (the API tree names pages `Interface~Member`),
then prints prose of at most `-n` pages truncated to `-l` lines. One lookup costs ~25 dense lines.

```
ds-docs decompile                               # all default CHMs; -a for every CHM found, -f to redo
ds-docs status                                  # sources, extractor, pages per set
ds-docs lookup IDocumentExporter~ExportToPng     # signature, what each parameter means, vendor examples
ds-docs lookup -t ZoomFit                       # command reference only: is this RunCommand name real?
ds-docs lookup -c InsertHatchByBoundary -n 3     # full-text search instead of filename search
ds-docs lookup dsLnWt -n 3 -l 20                # an enum family, shorter output
ds-docs html2txt docs\DraftSight\cmdref\command_reference_chart.htm | grep -i zoom
```

| Set | Pages | Contents |
|---|---|---|
| `docs/draftsightapi/` | 4,353 | COM API reference, **named by interface**: `Interop.dsAutomation~Interop.dsAutomation.IPrintManager~PrintOut.html`. Search `IPrintManager`, not `dsPrintManager` |
| `docs/DraftSight/` | 1,318 | application guide + command reference: `cmdref/command_reference_chart.htm`, `cmdvar/` |
| `docs/DraftSight_Lisp_Reference/` | 257 | AutoLISP |
| `docs/DraftSightMech/`, `docs/DraftSightSW/` | 222 / 256 | mechanical, sheet metal |
| `docs/DraftSightConnected/` | 49 | Connected |
| `docs/DSAPIHelp/`, `docs/NestingManager/` | 1 each | stubs, low value |

Same surface from Python (`from draftsight_com import docs`): `docs.lookup(pattern, pages=1, max_lines=70,
cmdref=False, content=False)`, `docs.decompile(names=None, all_chms=False, force=False)`,
`docs.find_pages`, `docs.html2txt`, `docs.status`. `lookup` is the one to call from a script: it returns a
string, and a miss says so rather than dumping the tree.

**Look things up in this order**, cheapest first, stopping as soon as one answers:

1. **`references/draftsight-api.md`** in this skill - the verified COM map, drafting-variable table, and
   the quirk list, including which members return success while doing nothing. For what the geometry
   should *be* (tapped hole representation, tap drill sizes, hole depth conventions) the sibling
   **`references/how-to-draw-hole.md`** is the same kind of verified note.
2. **the help pages** - `ds-docs lookup`, for the prose no map carries: what a parameter actually means,
   which `RunCommand` names exist (`ZoomExtents` does not; `ZoomFit` does), the vendor's worked examples.
3. **a live probe** - the only source for what no document states: how this build marshals a given call
   and which members silently do nothing.

## Recovery

If a COM call raises `pywintypes.com_error` with "The remote procedure call
failed", the app has died. Check the process, relaunch with
`ds.connect(timeout=120)`, and rerun the drawing script; scripts are cheap and
idempotent when they follow rules 2 and 3.
