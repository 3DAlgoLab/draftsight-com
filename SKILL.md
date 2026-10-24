---
name: draftsight-com
description: Draw, dimension, hatch, and export in DraftSight CAD by driving the running app over COM from the Python kernel (win32com). Use when asked to automate DraftSight, script or generate a DWG from measurements or a sketch, dimension a drawing, save DWG, produce a PNG/SVG proof, or when a DraftSight problem involves dimension text size, linetype dash length, entity layers/linestyles, document save/close, or COM automation speed. Windows only.
compatibility: Windows with DraftSight installed (DraftSight.Application ProgID registered) and a licensed user session; pywin32 (win32com) in the Python environment.
metadata:
  app: DraftSight
  api: COM IDispatch, late binding
---

# DraftSight over COM

One persistent COM connection to the running DraftSight app. Geometry is pushed
in-process, so a full drawing (layers, polylines, hatches, dimensions, leaders,
DWG save, image export) takes a couple of seconds instead of one process launch
per call, which is what a PowerShell-per-call driver pays.

Everything below is verified against a live DraftSight session, including the
places where the API returns success while doing nothing.

## Setup

The skill package `draftsight_com` installs editable into the agent kernel venv
at session start, so it is importable directly. For a project virtualenv, install
it once from its directory:

```
uv pip install -e "C:\Users\three\.local\share\prime-agent\skills\draftsight-com"
```

DraftSight must be reachable: either already running, or startable by `connect`.

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
sh.dimension((0, 120), (40, 120), (0, 140), 0.0, "%%c40", layer="DIM")
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
    {"dimension": {"p1": p, "p2": p, "at": p, "angle": 0.0, "text": "%%c23"}, "layer": "DIM"},
  ],
}
```

Every entity needs `layer`, and that layer must appear in `layers` too; an unknown entity
key or a missing layer raises naming the entity index. `examples\simple_section_spec.py` is
a complete worked spec (90 x 90 x 12 L-bracket: outline, hatch, two centerlines, two
dimensions, leader, note) that renders as described.

## Four rules that cost the most to learn

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

## Measurement and diagnosis

`references/draftsight-api.md` holds the verified API map, the drafting-variable
table with numeric proof, the COM quirks list, the crash modes, and the
SVG/pixel rigs that measure rendered text height and dash length exactly. Read
it before debugging anything that "should" work through the style objects.

## Recovery

If a COM call raises `pywintypes.com_error` with "The remote procedure call
failed", the app has died. Check the process, relaunch with
`ds.connect(timeout=120)`, and rerun the drawing script; scripts are cheap and
idempotent when they follow rules 2 and 3.
