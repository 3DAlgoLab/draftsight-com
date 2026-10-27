# draftsight-com

Drive a **running DraftSight** application over COM from Python, and hand the whole
workflow to a coding agent as a [pi](https://pi.dev) skill.

One persistent `win32com` connection means geometry is pushed **in-process**. A full
drawing — layers, polylines, hatches, dimensions, leaders, DWG save, PNG proof export —
takes a couple of seconds instead of one process launch per COM call, which is the tax a
PowerShell-per-call driver pays.

Everything in the references was verified against a live DraftSight session, including
the members that return success while doing nothing.

```
repo root = pi package
├── package.json          pi manifest: skill "." + bootstrap extension
├── SKILL.md              the skill the model loads (Agent Skills format)
├── extensions/index.ts   installs the Python half into the agent kernel venv
├── src/draftsight_com/   the Python package (`draftsight_com`, `ds-docs`)
├── references/           verified API map, hard-won rules, drafting conventions
├── examples/             a complete worked spec
└── pyproject.toml        the Python distribution
```

## Prerequisites

| Requirement | Notes |
|---|---|
| Windows | COM automation via `pywin32`; nothing cross-platform here |
| DraftSight installed and licensed | ProgID `DraftSight.Application`; default install path `C:\Program Files\Dassault Systemes\DraftSight\` |
| Python ≥ 3.11 with `pywin32` | the interpreter your agent actually runs code in, see [How the install works](#how-the-install-works) |
| pi (optional) | only for the skill side |

## Install

### As a pi package (recommended)

```sh
pi install git:github.com/3DAlgoLab/draftsight-com
```

That gives you the skill plus an extension that installs the Python package for you. Restart
pi, then check both halves:

```
/draftsight-com            # action: ready   python: ...\.pi\.venv\Scripts\python.exe
```

`/draftsight-com install` forces a reinstall (`uv pip install --editable`).

### Python only, any agent or plain scripts

```powershell
git clone https://github.com/3DAlgoLab/draftsight-com
uv pip install --python <your-python> --editable .\draftsight-com
```

Use `python -m pip install -e .` if you do not have [uv](https://docs.astral.sh/uv/).
Install with `.[docs]` if you want the pixel-measurement rigs from the references.

**Point it at the interpreter that will run the code.** In the last `python` you ran, does
`import draftsight_com` succeed? That question matters more here than in most packages,
because agent harnesses often have three interpreters in play:

```powershell
python -c "import draftsight_com, sys; print(draftsight_com.__file__); print(sys.executable)"
```

### From a local checkout, as a pi package

```sh
pi install ./draftsight-com        # local sources load in place, no copy
pi remove ./draftsight-com         # later
```

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
sh.dimension((0, 120), (40, 120), (0, 140), 0.0, "%%c40", layer="DIM")
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
    {"dimension": {"p1": p, "p2": p, "at": p, "angle": 0.0, "text": "%%c23"}, "layer": "DIM"},
  ],
}
```

Every entity needs a `layer`, and that layer must also appear in `layers`; an unknown key or
a missing layer raises naming the entity index. See `examples/simple_section_spec.py` for a
complete 90 × 90 × 12 mm L-bracket.

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

## How the install works

A repo that is both a pi package and a Python package has two install paths, and pi only runs
npm installs. The Python half therefore needs the *right* interpreter, which for an agent
harness is rarely the first `python` on `PATH` — on a typical machine `python` is a uv-managed
3.12, `py -3` a 3.14, and the pi kernel is a uv-created venv at `~/.pi/.venv` that has **no
`pip` at all**. Installing into the wrong one is silent: the command succeeds, the import still
fails.

`extensions/index.ts` avoids that by resolving candidates in order, probing each, and installing
with `uv pip install --python <exe> --editable .` only into a candidate that is a real, installable
interpreter:

1. `$DRAFTSIGHT_COM_PYTHON`, then `$PI_PYTHON`
2. the pi kernel venv (`$PI_CODING_AGENT_DIR/../.venv`, then `~/.pi/.venv`)
3. `$VIRTUAL_ENV`

Probes use `importlib.util.find_spec`, so nothing in your environment is executed. Nothing is
reinstalled when a candidate already satisfies the import. Set `DRAFTSIGHT_COM_PYTHON` if your
kernel lives somewhere unusual — that is the one switch worth knowing.

`/draftsight-com` prints the resolved interpreter, the action taken, and a copy-pasteable manual
install command. In an interactive session that arrives as a notification; in `--print` or RPC
mode, where there is no dialog surface, the same report is written to stderr.

Without pi, you are the extension: run the `uv pip install` line above against your own
interpreter.

## References

| Document | Contents |
|---|---|
| `SKILL.md` | the skill body: quickstart, the four rules that cost the most, lookup order |
| `references/draftsight-api.md` | verified COM API map, drafting-variable table with numeric proof, COM quirks, crash modes, the SVG/pixel measurement rigs |
| `references/how-to-draw-hole.md` | what the geometry *should* be: tapped-hole section and top view, callout style, ISO 261/724 tap drill and depth tables |
| `docs/` | generated locally by `ds-docs decompile`; vendor copyrighted, so not shipped |

Four rules that save the most time, in one place:

1. **Sizing comes from drafting variables**, not dimension styles — `Sheet.drafting_style()`
   before inserting anything.
2. **`SaveAs2` into a path another open document owns fails silently** and still returns
   success; `Sheet.save()` proves the write with `GetPathName()` plus mtime.
3. **Never close the active document** — `CloseDocument` on it takes the COM server down.
4. **Style entities through the entity helper** (`IEntityHelper.SetLayer` / `SetLineStyle`);
   `Layer`/`LineStyle` on returned dimension, hatch, and leader objects are not writable under
   late binding.

## Troubleshooting

| Symptom | Cause and fix |
|---|---|
| `ModuleNotFoundError: draftsight_com` in the agent, fine in your shell | installed into a different interpreter; check `sys.executable`, then `/draftsight-com` or `DRAFTSIGHT_COM_PYTHON` |
| bootstrap says `no-python` | no candidate interpreter exists at all; set `DRAFTSIGHT_COM_PYTHON` to the one your agent actually runs |
| `uv: command not found` during bootstrap | use `python -m pip install -e .` manually |
| `No module named pip` from the venv's python | expected: that venv was created by uv, so install with `uv pip install --python <that python>` |
| `pywintypes.com_error: The remote procedure call failed` | DraftSight died; `ds.connect(timeout=120)` and rerun. Scripts are idempotent when they follow rules 2 and 3 |
| Dimensions/notes render at the wrong size | drafting variables were set after inserting, or a dimension angle was passed in degrees — angles are radians |
| `UnicodeEncodeError` printing help pages | a cp949 console; `ds-docs` already reconfigures stdout to UTF-8, so update the package |

## Uninstall

```sh
pi remove git:github.com/3DAlgoLab/draftsight-com
uv pip uninstall --python <that python> prime-agent-skill-draftsight-com
```

The generated `docs/` tree is yours to keep or delete; nothing references it at runtime until
you run `ds-docs decompile` again.

## Status

Verified against a live DraftSight session on Windows. Windows-only by construction, no
published license yet, and the API surface covered is the drafting subset used by real
drawings (geometry, hatching, dimensioning, export) rather than the whole `dsAutomation`
tree.
