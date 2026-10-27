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
selm = doc.GetSelectionManager()        GetSelectedObjectCount(1)  GetSelectedObjects(1, out)  -> ret[0]
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
keeps the measured value. Its 5th argument (angle) is in **radians**, not degrees: passing 90.0 renders a dimension tilted ~27° off vertical with a wrong reading (45 measured as 40.23); use `math.pi / 2` for a vertical dimension. `InsertSimpleNote(x, y, z, height, angle, text)` takes an
explicit height, so notes do not depend on drafting variables.

## Circles and arcs

- `InsertCircle(x, y, z, radius)` -> `ICircle`. True circle entity; client wrapper
  `Sheet.circle((x, y), r, ...)`.
- `InsertArc(x, y, z, radius, start, end)` -> `ICircleArc`. **start/end are radians**,
  CCW from +x - same radian convention as `InsertRotatedDimension`. Verified live with a
  3/4-circle (start=`pi/2`, end=`2*pi`); client wrapper `Sheet.arc((x, y), r, start, end, ...)`.
- Both style through the entity helper (`SetLayer` / `SetLineStyle`).
- Do not fake a circle with a closed polyline: a 96-segment circle is visually perfect
  (sagitta < 1 um at D20) but a true circle exports as a real arc, stays exact at any
  zoom, and measures true length in selection code.

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

## Selection and per-entity length

Verified live against three picked entities (two lines, one arc) in a part drawing.

```python
import pythoncom, win32com.client as w32
selm = doc.GetSelectionManager()

n    = selm.GetSelectedObjectCount(1)                     # 3
out  = w32.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_VARIANT, None)
objs = selm.GetSelectedObjects(1, out)[0]                 # tuple of PyIDispatch, len == n
total = sum(float(w32.Dispatch(d).GetLength()) for d in objs)   # 921.571793
```

| Fact | Proof |
|---|---|
| The selector argument must be **1** | With three entities picked, `GetSelectedObjectCount` returns 0 for `0`, `2`, `3` and `-1`, and 3 only for `1`. A zero here is the wrong argument, not an empty selection |
| `GetSelectedObjects` returns entities in **`ret[0]`** | `ret == (tuple_of_PyIDispatch, None)`; `out.value` stays `None` on success. The out-VARIANT is decoration, never the payload |
| Both arguments are required | Calling with fewer raises `-2147352562` (wrong parameter count) or a pywin32 `int()` coercion error |
| Interface name comes from ITypeInfo | `d.GetTypeInfo().GetDocumentation(-1)[0]` gives `ILine`, `ICircleArc`. `GetTypeName()`, `GetEntityName()`, `TypeName`, `ObjectType` all return `None` or raise |
| `GetPickPoints(1, out)` returns picks in **`ret`** | Flat `x,y,z` tuple, one pick per selected entity in selection order; `out.value` stays `None` like `GetSelectedObjects`. This is the only way to recover *where the user clicked*, which entity geometry cannot give |
| `CDispatch(d)` is not a wrapper | It needs an `olerepr`; use `win32com.client.Dispatch(d)` for dynamic attribute access |

`GetLength()` is present on every sketch entity exposing `EvaluateAtPoint` /
`GetClosestPointOn` (`ILine`, `ICircleArc`, polylines, splines). It is **true arc
length**, not chord and not circumference. Cross-checked on an `ICircleArc`:

| Quantity | Value |
|---|---|
| `Radius` | 418.01591455136764 |
| `StartAngle` to `EndAngle` | 135.0004338959052 deg to 179.8049233612266 deg |
| `Radius * sweep` | 326.88253478734964 |
| `GetLength()` | 326.88253478734964 |

Exact to the last digit. `StartAngle` / `EndAngle` are **properties** holding
radians, not methods: `arc.StartAngle()` raises `float object is not callable`,
which reads like a broken API but is just a misplaced parenthesis.

Text, hatch, note, dimension and block entities have **no** `GetLength()` and raise
rather than returning 0, so guard per entity when the selection may contain them.
`GetSelectionManager` also offers `SelectByWindow`, `SelectByPolygon`,
`SelectByCrossline`, `SelectByPoint`, `GetPickPoints`, `ClearSelections` and
`PaperSpaceSelectionEnabled` for driving selection programmatically.

### Reading a signature off IDispatch

When a method's shape is unknown, `ITypeInfo` yields it without the typelib, which
beats guessing. pywin32 exposes `PyFUNCDESC` as `.args`, `.rettype`, `.invkind`,
`.memid`, `.cParamsOpt` - there is **no** `.cParams`:

```python
ti = obj._oleobj_.GetTypeInfo()
for i in range(ti.GetTypeAttr().cFuncs):
    f = ti.GetFuncDesc(i)
    print(ti.GetDocumentation(f.memid)[0], f.args, f.rettype)
```

`f.args` is `[(tdesc, wFlags, name), ...]` with `wFlags` 1 = in, 2 = out. Only
`GetDocumentation(memid)[0]` gives names; the boilerplate (`QueryInterface`,
`GetIDsOfNames`, `Invoke`, `get_*` / `put_*`) comes back too. `dir()` on a late-bound
wrapper proves nothing about what the server implements - enumerate ITypeInfo instead.

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
| `GetSelectedObjectCount(0)` / `GetSelectedObjects(0, out)` | returns 0 and no entities while entities are picked; the selector is **1** |
| `GetSelectedObjects` out-VARIANT | stays `None`; the entities arrive in `ret[0]` |

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
