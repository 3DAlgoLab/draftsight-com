# Verified DraftSight COM reference

Everything here was measured against a live DraftSight session through
`win32com.client.GetActiveObject("DraftSight.Application")` with **late binding**
(no `EnsureDispatch`, no gencache). Every DraftSight interface answers IDispatch,
so early binding buys nothing and risks clashing with the user’s instance.

## Object map

```
app  = GetActiveObject("DraftSight.Application")
doc  = app.GetActiveDocument()          app.GetDocuments() -> Python tuple of IDocument
doc.NewDocument / OpenDocument / SaveAs2 / GetPathName / GetModified
model = doc.GetModel(); sm = model.GetSketchManager()     # all entity creation
lm   = doc.GetLayerManager()            CreateLayer(n) -> (Layer, Result)  GetLayer(n) -> Layer|None
lsm  = doc.GetLineStyleManager()        GetLineStyle(n)  LoadLineStyle(n, file)
dsm  = doc.GetDimensionStyleManager()   GetActiveDimensionStyle()  (see sizing rules)
tsm  = doc.GetTextStyleManager()
eh   = app.GetEntityHelper()            SetLayer / SetLineStyle / GetLayer / SetErased
exp  = doc.GetDocumentExporter()        ExportToPng(path) ExportToSvg(path)
app.Zoom(0, None, None)                 zoom-fit before export
app.GetColorByIndex(i)                  layer colours
```

`IDocument` has no `Name`; use `GetPathName()`. Untitled documents report
`NONAME_<n>.dwg`.

## Point marshalling

| Target | Array shape |
|---|---|
| `InsertPolyline2D`, `InsertHatchByBoundary` | flat `x,y` pairs: `VARIANT(VT_ARRAY\|VT_R8, (x1,y1,x2,y2,...))` |
| `InsertLeader` | flat `x,y,z` triples |
| `InsertRotatedDimension` | three separate point variants (ext p1, ext p2, dimension line point) |
| `InsertHatchByBoundary` also needs | `VARIANT(VT_ARRAY\|VT_I4, (n_pts_per_loop, ...))` |
| string arrays | `VARIANT(VT_ARRAY\|VT_BSTR, (...))` |

Hatch boundaries must be closed loops; hatch one loop per call rather than
relying on multi-loop boundaries. `InsertRotatedDimension`’s 4th argument is a
**text override**, not a style name: `"%%c23"` renders the diameter sign, `""`
keeps the measured value. `InsertSimpleNote(x, y, z, height, angle, text)` takes an
explicit height, so notes do not depend on drafting variables.

## Sizing: drafting variables, with proof

`dsCommandOptionDouble_e` members via `doc.SetCommandOptionDouble(var, value)` /
`GetCommandOptionDouble(var)` (returns 1 on success, values in drawing units):

| Option | Meaning | Classic variable |
|---|---|---|
| 29 | arrow size | DIMASZ |
| 31 | dimension-line extension | DIMGAP-adjacent |
| 32 | extension-line overrun past dimension line | DEXO |
| 33 | extension-line offset from origin | DSE |
| 37 | **dimension text height** | DIMTXT |
| 38 | text offset from dimension line | DTM |
| 41 | linetype scale | LTSCALE |
| 60 | default note text height | TEXTSIZE-ish |

Numeric proof, glyph heights measured on exported rasters:

| Option 37 | Measured glyph height |
|---|---|
| 2.5 | 2.6 mm |
| 14.0 | 14.0 mm |

Linear 1:1. `doc.SetCommandOptionDouble(41, 0.35)` turned a `CENTER` centerline
from 32 mm dashes (looks solid at part scale) into 11 mm dashes with 3 mm gaps.
Set the variables **before** creating entities.

### What does not work, so do not retry it

- `GetDimensionStyleTextOptions().Height`, `GetDimensionStyleArrowsOptions().Size`,
  `GetDimensionStyleLineOptions().ExtensionLine*`: writes succeed, readback shows
  the new numbers, rendered entities ignore them.
- Creating and activating a new dimension style: rendered size unchanged.
- `IDimensionStyleFitOptions.SetDimensionScale(1, 1.0)` first: no effect on text.
- Per-entity `GetGeneralDimension().LineScale = 1..8`: no effect on text. Apparent
  differences before were `Zoom` fit artifacts, not style application.
- `doc.DimensionScale`: 1.0 by default, not the cause.
- `IDimensionStyleTextOptions` has no text-angle member; per-entity `TextRotation`
  is a no-op. To get horizontal text on vertical dimensions, place `note()` text
  instead of trusting dimension text rotation.

Useful per-entity dimension access: `dimension.GetGeneralDimension()` exposes
`Measurement`, `TextOverride`, `DimensionStyle`, `LineScale`, `TextRotation`,
`Layer`. The dispatch returned by `InsertRotatedDimension` alone does not.

## Styling entities

```python
eh = app.GetEntityHelper()
eh.SetLayer(entity, "DIM"); eh.SetLineStyle(entity, "CENTER")
```

`entity.Layer = "DIM"` fails on dimensions, hatches, and leaders under late
binding. `ISketchManager.SetObjectLayer(obj, layer)` also exists. Load a linetype
once per document: `lsm.LoadLineStyle("CENTER", LINETYPES)` where `LINETYPES` is
`...\Default Files\Linestyles\MM.LIN` (that file also carries `HIDDEN`,
`PHANTOM`, `CENTER2`, `CENTERX2`). `GetLineStyle(name)` returns `None` when absent.

## Crash and silent-failure modes

| Action | Result |
|---|---|
| `SaveAs2(path)` where another **open** document owns `path` | returns success, writes nothing. Verify `doc.GetPathName() == path` and mtime advanced. A stale `os.path.exists()` check hides this |
| `CloseDocument(path)` on the **active** document | COM server dies: the call faults, the next `NewDocument` raises a server exception, the process exits |
| `NewDocument(template)` with zero documents open | server exception; use `OpenDocument(template)` instead |
| `doc.Name` | `AttributeError`; use `GetPathName()` |
| `GetDocuments().GetCount()` | it is a Python tuple; iterate it directly |
| `CreateLayer(name)` | returns `(Layer, Result)`, not a bare `Layer` |
| `GetEntities(...)` | needs a selection filter plus out-arrays; prefer exported images for verification |
| `app.RunCommand` | exists, but system-variable strings are unnecessary once the drafting variables above are used |

If a call raises `pywintypes.com_error` `-2147023170` (RPC failed) the process is
gone. `Get-Process DraftSight` to confirm, then `ds.connect(timeout=120)` to
relaunch. Nothing unsaved survives; keep drawing scripts idempotent.

## Measurement rigs (how the numbers above were obtained)

Guessing at rendered size is what makes DraftSight styling painful. Two
instruments, and they agree:

**SVG, exact text height.** `exp.ExportToSvg(path)`. Each `<text>` carries
`transform="matrix(s 0 -0 s tx ty)"` with `font-size:1.397`. Text height in mm is
`s / px_per_mm` **of that same export** — matrix scales are not comparable
between exports because the view scale differs.

**Raster, anything visible.** `ExportToPng`, then per-colour masks:

```python
import numpy as np
from PIL import Image
px = np.asarray(Image.open(png).convert("RGB")).astype(int)
r, g, b = px[..., 0], px[..., 1], px[..., 2]
dim_ink = (g > 110) & (r < 110) & (b < 110)          # green DIM layer
# px/mm from a known feature span, e.g. a 100 mm extension-line spacing:
# span_px = np.ptp(np.where(dim_ink.any(axis=0))[0]) / 100.0
# then ASCII-dump a band to read glyphs and dash runs directly
rows = np.where(dim_ink.any(axis=1))[0]
for y in range(rows.min(), rows.min() + 30):
    print("".join("#" if dim_ink[y, x] else "." for x in range(cols.min(), cols.max())))
```

Dash length follows the same recipe on the centerline’s column. `numpy` and
`pillow` are already in the kernel venv.

## Speed

One persistent connection plus batched inserts: setup (new document, layers,
linetypes, drafting variables) ~0.8 s, a 20-entity section with two hatches, four
dimensions and a leader ~0.66 s, save plus PNG export ~1 s. A PowerShell process
per COM call pays process start plus COM registration every time and takes
minutes for the same drawing.
