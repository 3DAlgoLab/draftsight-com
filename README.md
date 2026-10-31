# draftsight-com

Drive a **running DraftSight** application over COM from Python: draw, dimension, hatch,
save DWG, export a PNG proof, and search the vendor's 6,457 pages of help without reading them.

One persistent `win32com` connection means geometry is pushed **in-process**. A full
drawing — layers, polylines, hatches, dimensions, leaders, DWG save, PNG proof export —
takes a couple of seconds instead of one process launch per COM call, which is the tax a
PowerShell-per-call driver pays.

Everything in the references was verified against a live DraftSight session, including
the members that return success while doing nothing.

```
draftsight-com/
├── src/draftsight_com/   the package: the COM client, plus `ds-docs`, the CHM help extractor/searcher
├── SKILL.md              Agent Skills file: how to drive the package, and the six rules that cost most
├── references/           verified API map, hard-won rules, drafting conventions
├── examples/             a complete worked drawing spec, a help-lookup script
├── tests/                the tapped-hole arithmetic — no DraftSight needed
└── pyproject.toml        the Python distribution (`draftsight_com` and `ds-docs` console scripts)


## Prerequisites

| Requirement | Notes |
|---|---|
| Windows | COM automation via `pywin32`; nothing cross-platform here |
| DraftSight installed and licensed | ProgID `DraftSight.Application`; default install path `C:\Program Files\Dassault Systemes\DraftSight\` |
| Python ≥ 3.11 with `pywin32` | install it into the interpreter that will run the code — see [Install](#install) |


## Install

### Install

```powershell
git clone https://github.com/3DAlgoLab/draftsight-com
cd draftsight-com
uv pip install --python <the-python-that-runs-your-code> --editable .
```

Use `python -m pip install -e .` if you do not have [uv](https://docs.astral.sh/uv/). Add `.[docs]`
if you want the pixel-measurement rigs described in the references (they need `numpy` and
`pillow`).

**The target interpreter is the whole problem.** In the last `python` you ran, does
`import draftsight_com` succeed? This matters more here than in most packages because the code
usually runs in an interpreter that is not the first `python` on `PATH`: a REPL kernel, a job
runner, a notebook server. A wrong target is silent — the install succeeds, the import still
fails. Ask the process that will run the code:

```python
import sys, os
print(sys.executable)
print(os.environ.get("PI_PYTHON", ""))
```

Then install into exactly that path, and verify from the same place rather than from your shell:

```
uv pip install --python "$PI_PYTHON" --editable .
```

```python
import draftsight_com
print(draftsight_com.__version__, draftsight_com.__file__)
```

### As a skill for a coding agent

`SKILL.md` is an [Agent Skills](https://agentskills.io/specification) file, so any harness that
follows the spec can pick it up. For pi, point the skills directory at the checkout so there stays
one copy of the files:

```powershell
cmd /c mklink /J "%USERPROFILE%\.pi\agent\skills\draftsight-com" "D:\dev\draftsight-com"
```

A junction rather than a copy: `SKILL.md` reaches `references/` by relative path, so copying the
file alone breaks the map, and copying the tree leaves two versions to keep in sync. The Python
package still has to be installed into the interpreter the agent runs cells in — a skill tells the
model what to call, it does not carry Python code.

## Quickstart

```python
import draftsight_com as ds

OUT_DWG = r"C:\temp\bracket.dwg"
OUT_PNG = r"C:\temp\bracket.png"

app = ds.connect(timeout=90)      # attach; start DraftSight and wait if not running
ds.close_document(app, OUT_DWG)   # reruns only: release a stale target before saving
sh = ds.new_document(app)         # from standardiso.dwt, ScaleUnit = mm

sh.drafting_style(text_height=14.0, arrow_size=6.0)   # BEFORE inserting anything
sh.layer("OUTLINE", 7); sh.layer("DIM", 3); sh.layer("CENTER", 1, "CENTER")

sh.polyline([(0, 0), (40, 0), (40, 120), (0, 120)], closed=True, layer="OUTLINE")
sh.hatch([[(0, 0), (40, 0), (40, 120), (0, 120)]], scale=2.0, layer="OUTLINE")
sh.circle((20, 60), 5.0, layer="OUTLINE")
sh.arc((20, 60), 8.0, 0.0, 3.14159, layer="OUTLINE")          # radians, CCW from +x
sh.line((0, -10), (0, 130), layer="CENTER", linetype="CENTER")
sh.dimension((0, 120), (40, 120), (0, 140), 0.0, ds.DIAMETER + "40", layer="DIM")  # literal Ø, not "%%c"
sh.leader([(10, 0), (50, -30), (60, -30)], 160.0, "M20 TAP,DP35", layer="DIM")

sh.zoom_fit()
sh.save(OUT_DWG)
sh.export_png(OUT_PNG)
```

**Then look at the PNG.** Rendering is the only honest check: COM reports success for
operations that change nothing, and several style objects accept writes, read back the new
numbers, and never affect a rendered entity.

### Declarative specs

`build(sheet, spec)` draws a dict; `make(spec, out, png=...)` runs the whole pipeline
(connect → new document → release stale target → build → zoom-fit → save → export).
`draftsight_com` is also a shell command:

```
draftsight_com --spec examples\simple_section_spec.py --out C:\temp\bracket.dwg --png C:\temp\bracket.png
```

```python
import draftsight_com as ds
sh = ds.make(SPEC, OUT_DWG, png=OUT_PNG)   # Sheet returned, ready for more edits
```

```python
SPEC = {
  "style":    {"text_height": 5.0, "arrow_size": 2.5},
  "layers":   [("OUTLINE", 7), ("CENTER", 1, "CENTER")],
  "entities": [
    {"polyline": pts, "closed": True, "layer": "OUTLINE"},
    {"circle": {"center": (20, 60), "radius": 5.0}, "layer": "OUTLINE"},
    {"arc":    {"center": (20, 60), "radius": 8.0, "start": 0.0, "end": 3.14159}, "layer": "OUTLINE"},
    {"hatch":  [loop], "pattern": "ANSI31", "scale": 1.5, "layer": "HATCH"},
    {"note":   [x, y, height, "SECTION B-B"], "layer": "TEXT"},
    {"leader": {"points": [...], "width": 70.0, "text": "BREAK EDGE"}, "layer": "TEXT"},
    {"dimension": {"p1": p, "p2": p, "at": p, "angle": 0.0, "text": "\u00d823"}, "layer": "DIM"},
  ],
}
```

Every entity needs a `layer`, and that layer must also appear in `layers`; an unknown key or
a missing layer raises naming the entity index. See `examples/simple_section_spec.py` for a
complete 90 × 90 × 12 mm L-bracket.

### A tapped hole in one call

`ds.tapped_hole(sheet, "M30", 55)` draws the section-and-top-view representation of an M30x3.5
tapped hole 55 deep and returns its numbers (tap drill 26.5, point cone 8, drilled depth 63).
`ds.tapped_hole_spec(size, depth)` is the same thing as a plain spec dict computed without COM, so:

```python
ds.make(ds.tapped_hole_spec("M30", 55), OUT_DWG, png=OUT_PNG)
```

The layout, thread arithmetic and layer set are measured off the drawings in the user's own set;
`references/how-to-draw-hole.md` has the table, and `tests/test_tapped_hole.py` pins the numbers.

### Searching the official help

DraftSight ships 6,457 pages of CHM help. Unpacking is cheap (**63 MB in 12 s with 7-Zip**);
*reading* it is the expensive part, so `ds-docs` is a search tool, never a reader:

```
ds-docs decompile                                # once per machine; gitignored, not shipped
ds-docs status
ds-docs lookup IDocumentExporter~ExportToPng     # signature + what each parameter means
ds-docs lookup -t ZoomFit                        # command reference: is this RunCommand name real?
ds-docs lookup -c InsertHatchByBoundary -n 3     # full-text instead of filename search
ds-docs lookup dsLnWt -n 3 -l 20                 # an enum family, shorter output
```

Same from Python: `from draftsight_com import docs` → `docs.lookup(...)`, `docs.decompile(...)`,
`docs.find_pages`, `docs.html2txt`, `docs.status`.

## Which interpreter, in one place

There is exactly one install target: **the interpreter that will execute the code**. Everything
else here is detail.

| You run the code from | Find the interpreter by | Then |
|---|---|---|
| a shell or a scheduled script | `Get-Command python` | install into it |
| an agent kernel (a REPL tool, Jupyter) | a cell: `import sys; print(sys.executable)`, or the pin that kernel exports (`PI_PYTHON` for pytool) | `uv pip install --python "$PI_PYTHON" --editable .` |
| a uv-created venv | it usually has **no `pip` at all** | `uv pip install --python <that venv python>` |

Two traps that cost hours:

- **32/64-bit.** A 32-bit Python cannot attach to a 64-bit DraftSight. Check once:
  `python -c "import sys; print(sys.maxsize > 2**32 and '64-bit' or '32-bit')"`.
- **Stale editable installs.** An editable `.pth` is a promise about a *directory*, not a copy of
  the code. If it points into `AppData\Local\Temp`, the import works today and breaks the day TEMP
  is cleaned. `uv pip list --python <exe> | grep -i draftsight`, then uninstall and reinstall.
## References

| Document | Contents |
|---|---|
| `SKILL.md` | the skill body: quickstart, the tapped-hole maker, the six rules that cost the most, lookup order |
| `references/draftsight-api.md` | verified COM API map, drafting-variable table with numeric proof, COM quirks, crash modes, the SVG/pixel measurement rigs |
| `references/how-to-draw-hole.md` | what the geometry *should* be: tapped-hole section and top view, callout style, ISO 261/724 tap drill and depth tables |
| `docs/` | generated locally by `ds-docs decompile`; vendor copyrighted, so not shipped |

Six rules that save the most time, in one place:

1. **Sizing comes from drafting variables**, not dimension styles — `Sheet.drafting_style()`
   before inserting anything.
2. **`SaveAs2` into a path another open document owns fails silently** and still returns
   success; `Sheet.save()` proves the write with `GetPathName()` plus mtime.
3. **Never close the active document** — `CloseDocument` on it takes the COM server down.
4. **Style entities through the entity helper** (`IEntityHelper.SetLayer` / `SetLineStyle`);
   `Layer`/`LineStyle` on returned dimension, hatch, and leader objects are not writable under
   late binding.
5. **`text=` arguments are literal** - `"%%c26.5"` draws as `%c26.5`; pass `ds.DIAMETER` (Ø), and
   keep a leader shelf wider than its text or the note wraps.
6. **Read back through layers and exports** - `Sheet.layers()` and `Sheet.house_style()`; the
   point-taking selection members raise `E_INVALIDARG` from Python.

## Troubleshooting

| Symptom | Cause and fix |
|---|---|
| `ModuleNotFoundError: draftsight_com` in the agent but fine in your shell | installed into a different interpreter; ask that kernel `print(sys.executable)` and install into the path it names |
| `ImportError: DLL load failed while importing win32api` | pywin32 is absent from that interpreter, or it is 32-bit against a 64-bit app |

| `No module named pip` from the venv's python | expected: that venv was created by uv, so install with `uv pip install --python <that python>` |
| `draftsight_com.__file__` points into `AppData\Local\Temp` | a stale editable install. An editable `.pth` is a promise about a directory, not a copy of the code, so the import works today on borrowed code and breaks the moment TEMP is cleaned. Uninstall from that interpreter and reinstall against the checkout |
| `pywintypes.com_error: The remote procedure call failed` | DraftSight died; `ds.connect(timeout=120)` and rerun. Scripts are idempotent when they follow rules 2 and 3 |
| `%c26.5` drawn where `Ø26.5` was wanted | `%%c` is a command-line code and is not expanded inside COM `text=` arguments; pass `ds.DIAMETER` |
| Leader note split over two lines | the shelf (`width`) is narrower than the text; use `NOTE_WIDTH_FACTOR` |
| Dimensions/notes render at the wrong size | drafting variables were set after inserting, or a dimension angle was passed in degrees — angles are radians |
| `UnicodeEncodeError` printing help pages | a cp949 console; `ds-docs` already reconfigures stdout to UTF-8, so update the package |

## Uninstall

```sh
uv pip uninstall --python <that python> draftsight-com
```


The generated `docs/` tree is yours to keep or delete; nothing references it at runtime until
you run `ds-docs decompile` again.

## Status

Verified against a live DraftSight session on Windows. Windows-only by construction, no
published license yet, and the API surface covered is the drafting subset used by real
drawings (geometry, hatching, dimensioning, export) rather than the whole `dsAutomation`
tree.
