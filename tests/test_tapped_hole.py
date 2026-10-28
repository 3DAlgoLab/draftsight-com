"""Pin the tapped-hole arithmetic and layout - no COM, no DraftSight needed.

Run with `python tests/test_tapped_hole.py` or any pytest. The numbers come from the user's
own drawings (M30DP45, M30DP55), so a change here means a drawing in the set stops matching.
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), os.pardir, "src"))

try:
    import draftsight_com as ds
except ImportError as exc:                       # pywin32 is required at import, Windows-only
    print("SKIP - draftsight_com needs pywin32:", exc)
    sys.exit(0)


def test_thread_spec():
    assert ds.thread_spec("M30") == (30.0, 3.5, 26.5)
    assert ds.thread_spec(" m20 ") == (20.0, 2.5, 17.5)
    assert ds.thread_spec("M30x3") == (30.0, 3.0, 27.0)
    assert ds.thread_spec("M12", pitch=1.25) == (12.0, 1.25, 10.75)


def test_unknown_size_offers_the_pitch_option():
    try:
        ds.thread_spec("M99")
    except ValueError as exc:
        assert "pitch=" in str(exc)
    else:
        raise AssertionError("M99 must not resolve to a coarse pitch")


def test_drill_point_depth_matches_the_dimensioned_cone():
    assert ds.drill_point_depth(26.5) == 8.0      # M30: drawings dimension depth + 8
    assert ds.drill_point_depth(17.5) == 5.0      # M20
    assert ds.drill_point_depth(26.5, 118.0) == ds.drill_point_depth(26.5)


def test_layout_matches_the_drawn_family():
    n = ds.tapped_hole_layout("M30", 55)
    assert (n["tap_drill"], n["point_depth"], n["drilled_depth"]) == (26.5, 8.0, 63.0)
    assert (n["y_top"], n["y_tap_bottom"], n["y_apex"]) == (31.5, -23.5, -31.5)
    assert n["section_axis"] == (-75.0, 0.0)
    n = ds.tapped_hole_layout("M30", 45)
    assert n["drilled_depth"] == 53.0 and (n["y_top"], n["y_tap_bottom"]) == (26.5, -18.5)
    n = ds.tapped_hole_layout("M20", 45)
    assert (n["tap_drill"], n["point_depth"], n["drilled_depth"]) == (17.5, 5.0, 50.0)


def test_layout_shifts_with_at_and_view_offset():
    n = ds.tapped_hole_layout("M30", 55, at=(200.0, 40.0), view_offset=120.0)
    assert n["top_view"] == (200.0, 40.0) and n["section_axis"] == (80.0, 40.0)
    assert n["y_top"] == 40.0 + 63.0 / 2.0


def test_spec_uses_kinds_the_serializer_understands():
    spec = ds.tapped_hole_spec("M30", 55)
    assert set(spec) == {"style", "layers", "entities", "numbers"}
    assert spec["style"] == ds.HOUSE_STYLE
    layers = [name for name, *_rest in ds.HOUSE_LAYERS]
    assert [name for name, *_rest in spec["layers"]] == layers
    for entity in spec["entities"]:
        kinds = [key for key in entity if key in ds.ENTITY_KINDS]
        assert len(kinds) == 1, entity                # exactly one kind per entity
        assert entity["layer"] in layers, entity      # and its layer is declared
    drawn = [next(k for k in e if k != "layer") for e in spec["entities"]]
    for kind in ("line", "circle", "arc", "hatch", "dimension", "leader", "note"):
        assert kind in drawn, kind


def test_spec_text_uses_a_literal_diameter_sign():
    # "%%c" is a command-line code; through COM text arguments it draws literally.
    spec = ds.tapped_hole_spec("M30", 55)
    notes = [e["leader"]["text"] for e in spec["entities"] if "leader" in e]
    assert notes == ["M30 x 3.5, DP55", "\u00d826.5 DP63"]
    dims = [e["dimension"]["text"] for e in spec["entities"] if "dimension" in e]
    assert dims == ["", "", "\u00d826.5"]
    assert all("%c" not in t for t in notes + dims)


def test_leader_shelf_clears_its_text():
    for size, depth in (("M30", 55), ("M20", 45), ("M12", 20)):
        for entity in ds.tapped_hole_spec(size, depth)["entities"]:
            if "leader" in entity:
                text = entity["leader"]["text"]
                assert entity["leader"]["width"] > len(text) * ds.HOUSE_STYLE["note_height"] * 0.7


def test_hatch_stops_at_thread_depth():
    spec = ds.tapped_hole_spec("M30", 55)
    hatch = [e for e in spec["entities"] if "hatch" in e][0]
    ys = [p[1] for p in hatch["hatch"][0]]
    n = spec["numbers"]
    assert max(ys) == n["y_top"] and min(ys) == n["y_tap_bottom"]


def main():
    tests = [(name, func) for name, func in sorted(globals().items())
             if name.startswith("test_") and callable(func)]
    for name, func in tests:
        func()
        print("ok", name)
    print("%d passed" % len(tests))


if __name__ == "__main__":
    main()
