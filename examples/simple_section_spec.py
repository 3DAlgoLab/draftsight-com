"""Worked drawing spec: a sectioned L-bracket, 90 x 90 x 12 mm.

Both entry points use this file's SPEC (flags, not positionals):

    draftsight_com --spec examples/simple_section_spec.py --out C:\\temp\\bracket.dwg --png C:\\temp\\bracket.png

    import draftsight_com as ds
    ds.make(SPEC, r"C:\\temp\\bracket.dwg", png=r"C:\\temp\\bracket.png")
"""

PROFILE = [(0.0, 0.0), (90.0, 0.0), (90.0, 12.0), (12.0, 12.0), (12.0, 90.0), (0.0, 90.0)]

SPEC = {
    "style": {"text_height": 5.0, "note_height": 5.0, "arrow_size": 2.5,
              "ext_offset": 1.5, "ext_overrun": 2.0, "text_offset": 1.0,
              "linetype_scale": 0.25},
    "layers": [("OUTLINE", 7), ("HATCH", 8), ("CENTER", 1, "CENTER"), ("DIM", 3), ("TEXT", 2)],
    "entities": [
        {"polyline": PROFILE, "closed": True, "layer": "OUTLINE"},
        {"hatch": [PROFILE], "pattern": "ANSI31", "scale": 1.5, "angle": 0.0, "layer": "HATCH"},
        {"line": [(36.0, -10.0), (36.0, 100.0)], "layer": "CENTER", "linetype": "CENTER"},
        {"line": [(0.0, 45.0), (100.0, 45.0)], "layer": "CENTER", "linetype": "CENTER"},
        {"dimension": {"p1": (0.0, 0.0), "p2": (90.0, 0.0), "at": (45.0, -28.0),
                       "angle": 0.0, "text": ""}, "layer": "DIM"},
        {"dimension": {"p1": (0.0, 0.0), "p2": (0.0, 90.0), "at": (-30.0, 45.0),
                       "angle": 1.5707963267948966, "text": "90 TYP"}, "layer": "DIM"},
        {"leader": {"points": [(90.0, 12.0), (118.0, 40.0), (128.0, 40.0)],
                    "width": 70.0, "text": "BREAK EDGE"}, "layer": "TEXT"},
        {"note": [0.0, -62.0, 8.0, "SECTION B-B"], "layer": "TEXT"},
    ],
}
