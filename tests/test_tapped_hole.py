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
                # The shelf has to clear the text it carries (a short shelf wraps the note). The
                # floor used to be a flat 0.7 per character, which was only ever a stand-in for
                # the render width; compare against the measured sum itself.
                need = ds.note_text_width(text, ds.HOUSE_STYLE["note_height"], safety=1.0)
                assert entity["leader"]["width"] >= need, (text, entity["leader"]["width"], need)
                # ...and stay near it: a shelf three times the text is a collision waiting to happen.
                assert entity["leader"]["width"] < need * 1.25, (text, entity["leader"]["width"], need)


def test_hatch_stops_at_thread_depth():
    spec = ds.tapped_hole_spec("M30", 55)
    hatch = [e for e in spec["entities"] if "hatch" in e][0]
    ys = [p[1] for p in hatch["hatch"][0]]
    n = spec["numbers"]
    assert max(ys) == n["y_top"] and min(ys) == n["y_tap_bottom"]


def test_note_text_width_reproduces_the_rendered_widths():
    # Widths below are notes inserted at 8 mm in a live DraftSight and read back through the
    # entity's own GetBoundingBox(). safety=1.0 compares the class table against the render.
    # Single-class strings pin the constants themselves; the long mixed callout pins the tolerance a
    # real note gets (one class has to cover every capital and digit glyph, so it drifts a little).
    for text, rendered, tol in (
        ("ABCDEFGHIJ", 69.94, 0.01),
        ("0123456789", 61.20, 0.01),
        ("aaaaaaaaaa", 61.28, 0.01),
        ("          ", 31.05, 0.01),
        (":--.,/;()", 29.00, 0.01),
        ("\u00d8\u00d8\u00d8\u00d8\u00d8\u00d8\u00d8\u00d8\u00d8\u00d8", 86.06, 0.01),
        ("6-020xDP13.5-014 DRLL PCD120", 169.89, 0.04),
    ):
        got = ds.note_text_width(text, 8.0, safety=1.0)
        assert abs(got - rendered) < tol * rendered, (text, got, rendered)
    # The default safety puts a shelf just over the render, never under it (wrapping is the failure).
    assert ds.note_text_width("6-020xDP13.5-014 DRLL PCD120", 8.0) > 169.89


def test_flat_note_width_factor_is_wrong_in_both_directions():
    # Why note_text_width exists: the old single 0.83 per character under-sizes a Hangul shop note
    # and over-sizes a digit-heavy callout, so a shelf sized from it can wrap the note or collide
    # with the geometry it was meant to clear.
    hangul = "\uc55e\ub3c4\uc870\ub3c4\ud655\uc778\ud544\uc694"
    assert ds.note_text_width(hangul, 8.0, safety=1.0) > len(hangul) * 8.0 * ds.NOTE_WIDTH_FACTOR
    assert ds.note_text_width("0123456789", 8.0, safety=1.0) < 10 * 8.0 * ds.NOTE_WIDTH_FACTOR
    assert ds.note_advance(ds.DIAMETER) > ds.note_advance("D")
    assert abs(ds.note_text_width("ABCDEFGHIJ", 8.0) / ds.note_text_width("ABCDEFGHIJ", 12.0) - 8.0 / 12.0) < 1e-9


def main():
    tests = [(name, func) for name, func in sorted(globals().items())
             if name.startswith("test_") and callable(func)]
    for name, func in tests:
        func()
        print("ok", name)
    print("%d passed" % len(tests))


if __name__ == "__main__":
    main()
