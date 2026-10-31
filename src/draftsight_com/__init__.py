"""DraftSight automation over COM with one persistent win32com connection.

Late binding is enough -- every DraftSight interface answers IDispatch. Geometry
is pushed in-process into the running application, so a whole drawing costs
seconds instead of one process launch per COM call.

Windows only, and DraftSight must be licensed and reachable: `connect` attaches
to the running instance (launching one if asked) and never spawns a second.
"""

import math
import os
import time
from importlib.metadata import PackageNotFoundError, version as _dist_version

import pythoncom
import win32com.client
from win32com.client import VARIANT

# The version lives in pyproject.toml, so read it from the install record instead of restating it.
# A cell answering "is this the code I installed?" should see a number, not just a path. The
# fallback covers a bare source checkout dropped on sys.path with no install record behind it.
try:
    __version__ = _dist_version("draftsight-com")
except PackageNotFoundError:
    __version__ = "0.0.0+no-install-record"

PROGID = "DraftSight.Application"
DRAFTSIGHT_EXE = r"C:\Program Files\Dassault Systemes\DraftSight\bin\DraftSight.exe"
TEMPLATE = r"C:\Program Files\Dassault Systemes\DraftSight\Default Files\Template\standardiso.dwt"
LINETYPES = r"C:\Program Files\Dassault Systemes\DraftSight\Default Files\Linestyles\MM.LIN"
MILLIMETERS = 8

# dsCommandOptionDouble_e members. These drafting variables, not dimension
# styles, are what DraftSight reads when it renders dimensions and notes.
DM_ARROW_SIZE = 29
DM_DIM_LINE_EXTEND = 31
DM_EXT_OVERRUN = 32
DM_EXT_OFFSET = 33
DM_TEXT_HEIGHT = 37          # DIMTXT, verified 1:1 in millimetres
DM_TEXT_OFFSET = 38
LINETYPE_SCALE = 41          # scales absolute linetype dash units
NOTE_TEXT_HEIGHT = 60

R8 = pythoncom.VT_ARRAY | pythoncom.VT_R8
I4 = pythoncom.VT_ARRAY | pythoncom.VT_I4
BSTR = pythoncom.VT_ARRAY | pythoncom.VT_BSTR


def connect(timeout=0.0, exe=DRAFTSIGHT_EXE):
    """Attach to DraftSight. With timeout > 0 start it first and wait that long.

    Raises RuntimeError naming the COM error when no instance answers.
    """
    deadline = time.monotonic() + timeout
    started = False
    while True:
        try:
            return win32com.client.GetActiveObject(PROGID)
        except pythoncom.com_error as exc:
            if not started and timeout > 0:
                if not os.path.exists(exe):
                    raise RuntimeError("DraftSight is not running and this path does not exist: " + exe) from exc
                os.startfile(exe)  # noqa: S606 - GUI launch, not a shell command
                started = True
            elif time.monotonic() >= deadline:
                raise RuntimeError(
                    f"no DraftSight COM instance after {timeout:.0f}s; start the app or pass its exe path"
                ) from exc
        time.sleep(2.0)


def new_document(app, template=TEMPLATE):
    """Create a document from `template` and return it as a Sheet.

    NewDocument faults on the server when DraftSight holds no document at all,
    so fall back to opening the template file as a document.
    """
    try:
        app.GetActiveDocument()
        app.NewDocument(template)
    except (AttributeError, pythoncom.com_error):
        _open(app, template)
    doc = app.GetActiveDocument()
    doc.ScaleUnit = MILLIMETERS
    return Sheet(doc, app)


def _open(app, path):
    """OpenDocument(PathName, Option) under the early-bound gen_py wrapper this
    machine registers; the Option argument is not accepted by every build, so
    fall back to the one-argument late-bound form.
    """
    path = os.path.abspath(path)
    try:
        app.OpenDocument(path, 0)
    except (TypeError, pythoncom.com_error):
        app.OpenDocument(path)
    return path


def open_document(app, path):
    _open(app, path)
    doc = app.GetActiveDocument()
    doc.ScaleUnit = MILLIMETERS
    return Sheet(doc, app)


def close_document(app, path):
    """Close the document DraftSight holds open on `path`, if any.

    SaveAs2 into a path another open document owns fails silently, so a rerun
    must release the old one. Closing the ACTIVE document takes the whole COM
    server down, so that is refused instead of attempted.
    """
    path = os.path.abspath(path)
    for other in app.GetDocuments():
        if other.GetPathName() == path:
            if app.GetActiveDocument().GetPathName() == path:
                raise RuntimeError(
                    "refusing to close the active document " + path +
                    "; it crashes DraftSight - create the new document first")
            app.CloseDocument(path, False)
            return True
    return False


def activate(app, path):
    """Bring the document DraftSight holds on `path` to the foreground.

    Needed to release a document safely: only an inactive one may be closed.
    """
    path = os.path.abspath(path)
    for doc in app.GetDocuments():
        if doc.GetPathName() == path:
            doc.Activate()
            return True
    return False


def documents(app):
    return [doc.GetPathName() for doc in app.GetDocuments()]


def _pts(*xy):
    return VARIANT(R8, tuple(float(v) for pair in xy for v in (pair[0], pair[1])))


def _pts3(*xy):
    return VARIANT(R8, tuple(float(v) for pair in xy for v in (pair[0], pair[1], 0.0)))


def _counts(*n):
    return VARIANT(I4, tuple(int(v) for v in n))


class Sheet:
    """Model space of one DraftSight document."""

    def __init__(self, doc, app=None):
        self.doc = doc
        self.app = app or connect()
        self.sm = doc.GetModel().GetSketchManager()
        self.eh = self.app.GetEntityHelper()
        self._linetypes = set()

    def drafting_style(self, text_height=14.0, arrow_size=6.0, ext_offset=4.0,
                       ext_overrun=5.0, text_offset=2.5, dim_line_extend=3.75,
                       note_height=14.0, linetype_scale=0.35):
        """Set drafting variables in millimetres. Call before inserting entities.

        IDimensionStyle text/arrow/extension option groups accept writes and
        read back the new numbers, yet are ignored at entity creation; only
        these document variables change rendered size.
        """
        for var, value in ((DM_TEXT_HEIGHT, text_height), (DM_ARROW_SIZE, arrow_size),
                           (DM_EXT_OFFSET, ext_offset), (DM_EXT_OVERRUN, ext_overrun),
                           (DM_TEXT_OFFSET, text_offset), (DM_DIM_LINE_EXTEND, dim_line_extend),
                           (NOTE_TEXT_HEIGHT, note_height), (LINETYPE_SCALE, linetype_scale)):
            self.doc.SetCommandOptionDouble(var, float(value))
        return self

    def drafting_variable(self, var):
        """One drafting variable in millimetres.

        The early-bound wrapper hands back the documented success flag next to
        the value as ``(value, 1)``; late binding hands back the bare value.
        Both read as a float here, so callers can compare against `HOUSE_STYLE`.
        """
        value = self.doc.GetCommandOptionDouble(var)
        if isinstance(value, tuple):
            value = value[0]
        return float(value)

    def drafting_variables(self, *variables):
        """Read drafting variables as {variable: millimetres}; default is every
        variable `drafting_style` writes."""
        return {var: self.drafting_variable(var) for var in (variables or STYLE_VARIABLES)}

    def house_style(self):
        """The live document's drafting sizes as `drafting_style` kwargs.

        Reads a drawing's real sizes so a new drawing matches the existing set:
        `new_sheet.drafting_style(**existing.house_style())`.
        """
        values = self.drafting_variables()
        return {name: values[var] for name, var in zip(STYLE_NAMES, STYLE_VARIABLES)}

    def layers(self):
        """Every layer as (name, ACI colour, linetype).

        `GetLayers()` is the dependable read-back path: window and point selection
        need point marshalling the registered typelib types as IDISPATCH, so an
        already-drawn entity list cannot be enumerated any other way. `IColor` is a
        typelib object too - it answers `GetColorIndex()`, it has no `ColorValue`.
        """
        rows = []
        for layer in self.doc.GetLayerManager().GetLayers():
            try:
                aci = int(layer.Color.GetColorIndex())
            except (AttributeError, TypeError, ValueError, pythoncom.com_error):
                aci = None
            rows.append((layer.Name, aci, getattr(layer, "LineStyle", None)))
        return rows

    def layer_names(self):
        return [name for name, _color, _linetype in self.layers()]

    def layer(self, name, color=7, linetype=None):
        lm = self.doc.GetLayerManager()
        layer = lm.GetLayer(name) or lm.CreateLayer(name)[0]
        layer.Color = self.app.GetColorByIndex(color)
        if linetype:
            self.linetype(linetype)
            layer.LineStyle = linetype
        return layer

    def linetype(self, name):
        if name not in self._linetypes:
            lsm = self.doc.GetLineStyleManager()
            if lsm.GetLineStyle(name) is None:
                lsm.LoadLineStyle(name, LINETYPES)
            self._linetypes.add(name)

    def line(self, a, b, layer=None, linetype=None):
        return self._style(self.sm.InsertLine(a[0], a[1], 0.0, b[0], b[1], 0.0), layer, linetype)

    def polyline(self, points, closed=False, layer=None, linetype=None):
        return self._style(self.sm.InsertPolyline2D(_pts(*points), bool(closed)), layer, linetype)

    def circle(self, center, radius, layer=None, linetype=None):
        return self._style(self.sm.InsertCircle(float(center[0]), float(center[1]), 0.0, float(radius)), layer, linetype)

    def arc(self, center, radius, start, end, layer=None, linetype=None):
        """Start/end are radians, CCW from +x (per the official API docs)."""
        return self._style(self.sm.InsertArc(float(center[0]), float(center[1]), 0.0,
                                               float(radius), float(start), float(end)), layer, linetype)

    def hatch(self, loops, pattern="ANSI31", scale=1.0, angle=0.0, layer=None):
        """Hatch each closed boundary in `loops` (sequences of x,y tuples)."""
        counts, flat = [], []
        for loop in loops:
            counts.append(len(loop))
            flat.extend(v for p in loop for v in (float(p[0]), float(p[1])))
        e = self.sm.InsertHatchByBoundary(_counts(*counts), VARIANT(R8, tuple(flat)),
                                          pattern, float(scale), float(angle))
        if e is None:
            raise RuntimeError("InsertHatchByBoundary returned no hatch")
        return self._style(e, layer)

    def note(self, x, y, height, text, layer=None, angle=0.0):
        return self._style(self.sm.InsertSimpleNote(float(x), float(y), 0.0, float(height),
                                                    float(angle), text), layer)

    def leader(self, vertices, note_width, text, layer=None):
        """Vertices are x,y tuples; `note_width` is the wrap width, so size it to
        the whole string or the note breaks onto extra lines."""
        return self._style(self.sm.InsertLeader(_pts3(*vertices), float(note_width), text), layer)

    def dimension(self, p1, p2, at, angle, text, layer=None):
        """Rotated dimension. `text` is a text override: "%%c23" renders a diameter
        sign, "" keeps the measured value. Size comes from drafting_style()."""
        e = self.sm.InsertRotatedDimension(_pts3(p1), _pts3(p2), _pts3(at), text, float(angle))
        return self._style(e, layer)

    def zoom_fit(self):
        self.app.Zoom(0, None, None)

    def export_png(self, path):
        path = os.path.abspath(path)
        if not self.doc.GetDocumentExporter().ExportToPng(path):
            raise RuntimeError("ExportToPng failed: " + path)
        return path

    def export_svg(self, path):
        """SVG text carries transform="matrix(s 0 -0 s ..)"; height_mm = s / (px/mm)
        of that same export, which measures dimension text exactly."""
        path = os.path.abspath(path)
        if not self.doc.GetDocumentExporter().ExportToSvg(path):
            raise RuntimeError("ExportToSvg failed: " + path)
        return path

    def save(self, path, version=28):
        """SaveAs2 silently no-ops when another open document owns the target
        path, so prove the write with the document path and the mtime."""
        path = os.path.abspath(path)
        before = os.path.getmtime(path) if os.path.exists(path) else 0.0
        self.doc.SaveAs2(path, version, True)
        if self.doc.GetPathName() != path or os.path.getmtime(path) <= before:
            raise RuntimeError("SaveAs2 did not write: " + path)
        return path

    def _style(self, entity, layer=None, linetype=None):
        # Dimensions, hatches and leaders expose no writable Layer; the entity
        # helper is the one styling path that covers every entity type.
        if linetype:
            self.linetype(linetype)
            self.eh.SetLineStyle(entity, linetype)
        if layer:
            self.eh.SetLayer(entity, layer)
        return entity


ENTITY_KINDS = ("line", "polyline", "circle", "arc", "hatch", "note", "leader", "dimension")


# `drafting_style` keyword order, paired with the variable ids it writes.
STYLE_NAMES = ("text_height", "arrow_size", "ext_offset", "ext_overrun", "text_offset",
               "dim_line_extend", "note_height", "linetype_scale")
STYLE_VARIABLES = (DM_TEXT_HEIGHT, DM_ARROW_SIZE, DM_EXT_OFFSET, DM_EXT_OVERRUN,
                   DM_TEXT_OFFSET, DM_DIM_LINE_EXTEND, NOTE_TEXT_HEIGHT, LINETYPE_SCALE)


ISO_COARSE_PITCH = {
    "M1.6": 0.35, "M2": 0.40, "M2.5": 0.45, "M3": 0.50, "M4": 0.70, "M5": 0.80,
    "M6": 1.00, "M7": 1.00, "M8": 1.25, "M10": 1.50, "M12": 1.75, "M14": 2.00,
    "M16": 2.00, "M18": 2.50, "M20": 2.50, "M22": 2.50, "M24": 3.00, "M27": 3.00,
    "M30": 3.50, "M33": 3.50, "M36": 4.00, "M42": 4.50, "M48": 5.00,
}

# The DP family: verified by reading M20DP45 / M30DP45 / M30DP55 back out of the DWGs.
HOUSE_LAYERS = (("OUTLINE", 7), ("MAJOR", 8), ("THREAD", 2), ("DIM", 3),
                ("CENTER", 1, "CENTER"), ("TEXT", 7))
HOUSE_STYLE = {"text_height": 7.0, "arrow_size": 3.5, "ext_offset": 2.0, "ext_overrun": 2.0,
               "text_offset": 1.5, "dim_line_extend": 3.75, "note_height": 7.0,
               "linetype_scale": 0.35}

# Literal diameter sign for COM text arguments ("%%%%c" is not substituted there).
DIAMETER = "\u00d8"

# Note text is about this wide per character per note height; 0.83 is measured
# ("SECTION VIEW" at 7 mm renders 69.7 mm) and a leader shelf needs the width to
# clear the text, because a shelf narrower than the text wraps it onto two lines.
NOTE_WIDTH_FACTOR = 0.83


def thread_spec(size, pitch=None):
    """(major, pitch, tap drill) in millimetres for "M30", "m30 x 3" or an explicit pitch.

    Tap drill = major - pitch, the ~75 % engagement rule the DP drawings dimension
    (M30 x 3.5 -> 26.5, M20 x 2.5 -> 17.5).
    """
    name = str(size).strip().upper().replace(" ", "")
    if not name.startswith("M"):
        raise ValueError(f"thread size must look like M12 or M12x1.25, got {size!r}")
    if "X" in name[1:]:
        name, fine = name.split("X", 1)
        pitch = float(fine)
    elif pitch is None:
        try:
            pitch = ISO_COARSE_PITCH[name]
        except KeyError:
            raise ValueError(f"no ISO coarse pitch for {size!r}; pass pitch=... - known sizes are "
                             + ", ".join(ISO_COARSE_PITCH)) from None
    major = float(name[1:])
    return major, float(pitch), round(major - float(pitch), 2)


def drill_point_depth(tap_drill, point_angle=118.0):
    """Cone length below the tap depth: (tap drill / 2) / tan(point angle / 2), to the mm.

    M30 (26.5) gives 7.96 -> 8 mm, and M30DP45 / M30DP55 both dimension a drilled
    depth of thread depth + 8, so the rounding is house style, not a shortcut.
    """
    return float(round((float(tap_drill) / 2.0) / math.tan(math.radians(point_angle / 2.0))))


def tapped_hole_layout(size, tap_depth, *, pitch=None, point_angle=118.0, point_depth=None,
                       at=(0.0, 0.0), view_offset=75.0):
    """Every number for one tapped hole, without touching COM.

    `tap_depth` is the DP figure in the callout. Layout, measured off the drawn
    family: the top view sits at `at`, the section axis `view_offset` to its left
    (75 mm in all three existing drawings - depth is not part of the spacing), and
    the section is centred vertically on the top view, so the top face lands at
    +drilled/2 and the drill point at -drilled/2.
    """
    major, pitch, tap_drill = thread_spec(size, pitch)
    cone = float(point_depth) if point_depth is not None else drill_point_depth(tap_drill, point_angle)
    drilled = round(float(tap_depth) + cone, 6)
    ax, ay = float(at[0]), float(at[1])
    y_top = ay + drilled / 2.0
    return {"size": f"M{major:g}", "pitch": pitch, "major": major, "tap_drill": tap_drill,
            "tap_depth": float(tap_depth), "point_depth": cone, "drilled_depth": drilled,
            "top_view": (ax, ay), "section_axis": (ax - float(view_offset), ay),
            "y_top": y_top, "y_tap_bottom": y_top - float(tap_depth), "y_apex": y_top - drilled}


def tapped_hole_spec(size, tap_depth, *, at=(0.0, 0.0), view_offset=75.0, pitch=None,
                     point_angle=118.0, point_depth=None, hatch_pattern="ANSI31", hatch_scale=0.8,
                     dim_right=32.0, dim_left=30.0, dia_dim_drop=None, center_mark=None,
                     title_gap=26.0, label=True, title=("SECTION VIEW", "TOP VIEW"),
                     style=None, size_label=None):
    """Build-spec dict (style, layers, entities) for a tapped hole: section + top view.

    Feed it straight to `build`/`make`, or take the returned dict apart. Geometry is
    the DP family representation read back out of M20DP45, M30DP45 and M30DP55:

    * section - minor/tap-drill lines and the drill point are OUTLINE (thick), the
      two major-diameter lines are MAJOR (thin grey), ANSI31 hatch between the minor
      lines stops at the tap depth so the point stays clear, the centerline overshoots
      the view by 6 mm at both ends;
    * top view - full circle on the minor diameter, and the major diameter drawn from
      90 to 360 degrees, i.e. a quarter arc missing from the upper right;
    * dimensions - thread depth right of the section, drilled depth left of it, tap
      drill diameter under the top view;
    * text - "M30 x 3.5, DP55" ledged off the section, "D26.5 DP63" (D = the diameter
      sign) ledged off the drill point, view titles centred 26 mm below the drill point.

    `DIAMETER` is a literal diameter sign, not the "%%%%c" code: DraftSight substitutes
    that code on the command line but hands COM text arguments straight to the entity,
    so a code in `text=` lands on the drawing as "%c".

    `at` is the top-view centre; the section goes to its left.
    """
    L = tapped_hole_layout(size, tap_depth, pitch=pitch, point_angle=point_angle,
                           point_depth=point_depth, at=at, view_offset=view_offset)
    minor_r = L["tap_drill"] / 2.0
    major_r = L["major"] / 2.0
    ax, ay = L["top_view"]
    cx, _cy = L["section_axis"]
    y0, y1, y2 = L["y_top"], L["y_tap_bottom"], L["y_apex"]
    note_h = (style or HOUSE_STYLE).get("note_height", HOUSE_STYLE["note_height"])
    ents = [
        {"line": [(cx, y0 + 6.0), (cx, y2 - 6.0)], "layer": "CENTER", "linetype": "CENTER"},
        {"line": [(cx - minor_r, y0), (cx - minor_r, y1)], "layer": "OUTLINE"},
        {"line": [(cx + minor_r, y0), (cx + minor_r, y1)], "layer": "OUTLINE"},
        {"line": [(cx - major_r, y0), (cx - major_r, y1)], "layer": "MAJOR"},
        {"line": [(cx + major_r, y0), (cx + major_r, y1)], "layer": "MAJOR"},
        {"line": [(cx - minor_r, y1), (cx + minor_r, y1)], "layer": "OUTLINE"},
        {"line": [(cx - minor_r, y1), (cx, y2)], "layer": "OUTLINE"},
        {"line": [(cx + minor_r, y1), (cx, y2)], "layer": "OUTLINE"},
        {"hatch": [[(cx - minor_r, y0), (cx + minor_r, y0), (cx + minor_r, y1), (cx - minor_r, y1)]],
         "pattern": hatch_pattern, "scale": hatch_scale, "layer": "THREAD"},
        {"circle": {"center": (ax, ay), "radius": minor_r}, "layer": "OUTLINE"},
        {"arc": {"center": (ax, ay), "radius": major_r, "start": math.pi / 2.0, "end": 2.0 * math.pi},
         "layer": "MAJOR"},
    ]
    mark = float(major_r + 5.5) if center_mark is None else float(center_mark)
    ents += [
        {"line": [(ax, ay + mark), (ax, ay - mark)], "layer": "CENTER", "linetype": "CENTER"},
        {"line": [(ax - mark, ay), (ax + mark, ay)], "layer": "CENTER", "linetype": "CENTER"},
        {"dimension": {"p1": (cx + minor_r, y0), "p2": (cx + minor_r, y1),
                       "at": (cx + dim_right, (y0 + y1) / 2.0), "angle": math.pi / 2.0, "text": ""},
         "layer": "DIM"},
        {"dimension": {"p1": (cx - minor_r, y0), "p2": (cx - minor_r, y2),
                       "at": (cx - dim_left, (y0 + y2) / 2.0), "angle": math.pi / 2.0, "text": ""},
         "layer": "DIM"},
        {"dimension": {"p1": (ax - minor_r, ay), "p2": (ax + minor_r, ay),
                       "at": (ax, ay - ((minor_r + 9.25) if dia_dim_drop is None else dia_dim_drop)),
                       "angle": 0.0, "text": "%s%g" % (DIAMETER, L["tap_drill"])}, "layer": "DIM"},
    ]
    if label:
        tap_note = "%s x %g, DP%g" % (size_label or L["size"], L["pitch"], L["tap_depth"])
        drill_note = "%s%g DP%g" % (DIAMETER, L["tap_drill"], L["drilled_depth"])
        ents += [
            {"leader": {"points": [(cx + minor_r, y0 - 16.0), (cx + 45.0, y0 + 17.0),
                                   (cx + 49.0, y0 + 17.0)],
                        "width": len(tap_note) * note_h * NOTE_WIDTH_FACTOR,
                        "text": tap_note}, "layer": "TEXT"},
            {"leader": {"points": [(cx - 0.47 * minor_r, y2 + 0.3), (cx - 38.0, y2 - 16.0),
                                   (cx - 44.0, y2 - 16.0)],
                        "width": len(drill_note) * note_h * NOTE_WIDTH_FACTOR,
                        "text": drill_note}, "layer": "TEXT"},
        ]
    if title:
        for text, centre in zip(title, (cx, ax)):
            half = len(text) * note_h * NOTE_WIDTH_FACTOR / 2.0
            ents.append({"note": [centre - half, y2 - title_gap, note_h, text], "layer": "TEXT"})
    return {"style": dict(HOUSE_STYLE, **(style or {})), "layers": list(HOUSE_LAYERS),
            "entities": ents, "numbers": L}


def tapped_hole(sheet, size, tap_depth, **options):
    """Draw the tapped hole described by `tapped_hole_spec` and return its numbers.

    The numbers dict is the point: assert on `drilled_depth` / `tap_drill` instead of
    re-deriving thread arithmetic at the call site.
    """
    spec = tapped_hole_spec(size, tap_depth, **options)
    build(sheet, spec)
    return spec["numbers"]


def build(sheet, spec):
    """Draw a declarative drawing spec onto `sheet` and return the sheet.

    `spec` is a dict::

        {"style": {drafting_style kwargs},              # applied before entities
         "layers": [(name, color), (name, color, linetype), ...],
         "entities": [
            {"line": [a, b], "layer": "CENTER", "linetype": "CENTER"},
            {"polyline": pts, "closed": True, "layer": "OUTLINE"},
            {"hatch": [loop, ...], "pattern": "ANSI31", "scale": 2.0, "angle": 0.0},
            {"note": [x, y, height, "TEXT"], "layer": "TEXT"},
            {"leader": {"points": [...], "width": 160.0, "text": "M20 TAP,DP35"}},
            {"dimension": {"p1": p, "p2": p, "at": p, "angle": 0.0, "text": "%%c23"}},
         ]}

    Points are x,y tuples. Every entity needs a `layer` and is styled through the
    entity helper. An entity key outside ENTITY_KINDS is a bug in the spec, so it
    raises rather than being skipped.
    """
    for entry in spec.get("layers", ()):
        sheet.layer(*entry)
    if spec.get("style"):
        sheet.drafting_style(**spec["style"])
    for index, item in enumerate(spec.get("entities", ())):
        kinds = [key for key in item if key in ENTITY_KINDS]
        if len(kinds) != 1:
            raise ValueError(f"entity {index} needs exactly one of {ENTITY_KINDS}, got keys {sorted(item)}")
        kind = kinds[0]
        opts = {k: v for k, v in item.items() if k != kind}
        if "layer" not in opts:
            raise ValueError(f"entity {index} ({kind}) has no layer; create it in spec['layers'] as well")
        if kind == "line":
            a, b = item["line"]
            sheet.line(a, b, **opts)
        elif kind == "polyline":
            sheet.polyline(item["polyline"], **opts)
        elif kind == "circle":
            c = item["circle"]
            sheet.circle(c["center"], c["radius"], **opts)
        elif kind == "arc":
            a = item["arc"]
            sheet.arc(a["center"], a["radius"], a["start"], a["end"], **opts)
        elif kind == "hatch":
            sheet.hatch(item["hatch"], **opts)
        elif kind == "note":
            x, y, height, text = item["note"]
            sheet.note(x, y, height, text, **opts)
        elif kind == "leader":
            pts, width, text = item["leader"]["points"], item["leader"]["width"], item["leader"]["text"]
            sheet.leader(pts, width, text, **opts)
        else:
            d = item["dimension"]
            sheet.dimension(d["p1"], d["p2"], d["at"], d.get("angle", 0.0), d.get("text", ""), **opts)
    return sheet


def make(spec, out, png="", template=None, app=None):
    """Build a whole document from `spec`: new document, entities, zoom-fit, save, export.

    Order matters: the new document becomes active *before* the stale target-path
    document is released, because closing the active document kills DraftSight and
    an owned target path makes SaveAs2 write nothing.
    """
    app = app or connect()
    sheet = new_document(app, template or TEMPLATE)
    close_document(app, out)
    build(sheet, spec)
    sheet.zoom_fit()
    sheet.save(out)
    if png:
        sheet.export_png(png)
    return sheet


def run(spec: str, out: str, png: str = "", template: str = TEMPLATE) -> str:
    """Draw a DWG from a Python spec file and return a one-line summary.

    `spec` is a .py file defining `SPEC` in the dict form documented on `build`.
    Pass `png` for a raster proof of the finished drawing. DraftSight must be
    running or startable at the default install path.
    """
    module = _load_spec(spec)
    make(module.SPEC, out, png=png, template=template)
    proof = ", png=" + os.path.abspath(png) if png else ""
    return (f"{os.path.abspath(out)}: {os.path.getsize(out)} bytes, "
            f"{len(module.SPEC.get('entities', ()))} entities{proof}")


def main(argv=None):
    """Entry point for the `draftsight_com` shell command: draw a DWG from a spec file."""
    import argparse

    ap = argparse.ArgumentParser(prog="draftsight_com",
                                 description="Draw a DWG from a Python spec dict and export a proof PNG.")
    ap.add_argument("--spec", required=True, help="python file defining SPEC (see build)")
    ap.add_argument("--out", required=True, help="dwg path to write")
    ap.add_argument("--png", default="", help="optional png proof path")
    ap.add_argument("--template", default=TEMPLATE, help="dwt template for the new document")
    a = ap.parse_args(argv)
    print(run(a.spec, a.out, png=a.png, template=a.template))
    return 0


def _load_spec(path):
    import importlib.util

    path = os.path.abspath(path)
    if not os.path.exists(path):
        raise RuntimeError("no such spec file: " + path)
    loader = importlib.util.spec_from_file_location("draftsight_spec", path)
    module = importlib.util.module_from_spec(loader)
    loader.loader.exec_module(module)
    if not hasattr(module, "SPEC"):
        raise RuntimeError(path + " defines no SPEC dict")
    return module
