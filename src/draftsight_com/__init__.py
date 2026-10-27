"""DraftSight automation over COM with one persistent win32com connection.

Late binding is enough -- every DraftSight interface answers IDispatch. Geometry
is pushed in-process into the running application, so a whole drawing costs
seconds instead of one process launch per COM call.

Windows only, and DraftSight must be licensed and reachable: `connect` attaches
to the running instance (launching one if asked) and never spawns a second.
"""

import os
import time

import pythoncom
import win32com.client
from win32com.client import VARIANT

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
        app.OpenDocument(template)
    doc = app.GetActiveDocument()
    doc.ScaleUnit = MILLIMETERS
    return Sheet(doc, app)


def open_document(app, path):
    app.OpenDocument(os.path.abspath(path))
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
        return self.doc.GetCommandOptionDouble(var)

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


ENTITY_KINDS = ("line", "polyline", "hatch", "note", "leader", "dimension")


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
