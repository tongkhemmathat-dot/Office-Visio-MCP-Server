import base64

import pytest

import math

from visio_mcp_server.svg_import import (Ellipse, Image, Path, Polyline, Rect, SvgError, Text, image_to_png,
                                         parse_color, parse_path, parse_svg, parse_transform)

PNG_1PX = base64.b64encode(
    bytes.fromhex("89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4890000000d4944415478"
                  "9c6360000002000001e221bc330000000049454e44ae426082")).decode()

SAMPLE = f"""<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink"
  width="200" height="100" viewBox="0 0 200 100" font-family="Tahoma, Arial">
<rect width="200" height="100" fill="#FFFFFF"/>
<defs><marker id="m"/></defs>
<rect x="10" y="20" width="50" height="30" fill="none" stroke="#00A4A6" stroke-width="2" stroke-dasharray="6,4"/>
<polyline points="10,70 90.5,70" fill="none" stroke="#232F3E" stroke-width="2" marker-end="url(#m)"/>
<text x="100" y="40" font-size="13" text-anchor="middle" fill="#232F3E" font-weight="bold">
<tspan x="100" dy="0">Line one</tspan><tspan x="100" dy="16">Line two</tspan></text>
<text x="5" y="9" font-size="12" fill="#000" stroke="#FFFFFF" stroke-width="4" paint-order="stroke">Halo</text>
<image x="1" y="2" width="30" height="40" xlink:href="data:image/png;base64,{PNG_1PX}"/>
<circle cx="5" cy="5" r="2"/>
</svg>"""


@pytest.fixture
def svg_file(tmp_path):
    p = tmp_path / "t.svg"
    p.write_text(SAMPLE, encoding="utf8")
    return str(p)


def test_page_size_and_page_background_is_dropped(svg_file):
    doc = parse_svg(svg_file)
    assert (doc.width, doc.height) == (200, 100)
    assert not any(isinstance(i, Rect) and i.x == 0 and i.width == 200 for i in doc.items)


def test_items_keep_document_order_and_types(svg_file):
    kinds = [type(i) for i in parse_svg(svg_file).items]
    assert kinds == [Rect, Polyline, Text, Text, Image, Ellipse]


def test_rect_style(svg_file):
    rect = parse_svg(svg_file).items[0]
    assert (rect.x, rect.y, rect.width, rect.height) == (10, 20, 50, 30)
    assert rect.fill is None and rect.stroke == "#00A4A6" and rect.stroke_width == 2 and rect.dash == "dash"


def test_polyline_points_and_arrow(svg_file):
    line = parse_svg(svg_file).items[1]
    assert line.points == [(10, 70), (90.5, 70)]
    assert line.arrow_end and not line.arrow_start and line.dash is None


def test_text_lines_anchor_and_halo(svg_file):
    text, halo = parse_svg(svg_file).items[2:4]
    assert text.lines == ["Line one", "Line two"] and text.line_height == 16
    assert text.anchor == "middle" and text.bold
    assert halo.lines == ["Halo"] and halo.halo and halo.fill == "#000000"


def test_sample_has_no_skipped_elements(svg_file):
    assert parse_svg(svg_file).skipped == []


def test_viewbox_is_used_when_width_missing(tmp_path):
    p = tmp_path / "v.svg"
    p.write_text('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 30 40"/>')
    assert (parse_svg(str(p)).width, parse_svg(str(p)).height) == (30, 40)


@pytest.mark.parametrize("content", ["not xml", "<html/>", '<svg xmlns="http://www.w3.org/2000/svg"/>'])
def test_bad_svg_raises(tmp_path, content):
    p = tmp_path / "bad.svg"
    p.write_text(content)
    with pytest.raises(SvgError):
        parse_svg(str(p))


def test_missing_file_raises(tmp_path):
    with pytest.raises(SvgError):
        parse_svg(str(tmp_path / "nope.svg"))


@pytest.mark.parametrize("value,expected", [
    ("#abc", "#AABBCC"), ("#00a4a6", "#00A4A6"), ("rgb(1, 2, 255)", "#0102FF"),
    ("none", None), ("url(#g)", None), (None, None), ("white", "#FFFFFF"),
])
def test_parse_color(value, expected):
    assert parse_color(value) == expected


def test_png_data_uri_is_written_as_is(tmp_path):
    path = image_to_png(f"data:image/png;base64,{PNG_1PX}", str(tmp_path), str(tmp_path), "i0")
    assert path.endswith("i0.png") and open(path, "rb").read(4) == b"\x89PNG"


def test_relative_image_file_is_resolved(tmp_path):
    (tmp_path / "icon.png").write_bytes(base64.b64decode(PNG_1PX))
    path = image_to_png("icon.png", str(tmp_path), str(tmp_path / "w") if False else str(tmp_path), "i1")
    assert open(path, "rb").read(4) == b"\x89PNG"


def test_missing_image_file_raises(tmp_path):
    with pytest.raises(SvgError):
        image_to_png("nope.png", str(tmp_path), str(tmp_path), "i2")


# --- groups, transforms, shapes, styles ----------------------------------------------------------

def svg_of(tmp_path, body, attrs='width="200" height="100"'):
    p = tmp_path / "x.svg"
    p.write_text(f'<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" {attrs}>'
                 f"{body}</svg>", encoding="utf8")
    return parse_svg(str(p))


def test_group_translate_and_inherited_style(tmp_path):
    doc = svg_of(tmp_path, '<g transform="translate(10,20)" fill="#112233" stroke="#FF0000" stroke-width="3">'
                           '<rect x="1" y="2" width="5" height="6"/></g>')
    r = doc.items[0]
    assert (r.x, r.y, r.width, r.height) == (11, 22, 5, 6)
    assert r.fill == "#112233" and r.stroke == "#FF0000" and r.stroke_width == 3


def test_scale_scales_geometry_and_stroke(tmp_path):
    r = svg_of(tmp_path, '<g transform="scale(2)"><rect x="1" y="1" width="3" height="4" stroke="#000" '
                         'stroke-width="2" fill="none"/></g>').items[0]
    assert (r.x, r.y, r.width, r.height, r.stroke_width) == (2, 2, 6, 8, 4)


def test_rotated_rect_becomes_closed_polygon(tmp_path):
    p = svg_of(tmp_path, '<rect x="0" y="0" width="10" height="10" transform="rotate(90)" fill="#000"/>').items[0]
    assert isinstance(p, Polyline) and p.closed and len(p.points) == 5
    assert p.points[1] == pytest.approx((0, 10), abs=1e-9)


def test_rounded_rect(tmp_path):
    r = svg_of(tmp_path, '<rect width="20" height="10" x="1" y="1" rx="3" fill="#000"/>').items[0]
    assert r.radius == 3


def test_circle_ellipse_line_polygon(tmp_path):
    items = svg_of(tmp_path, '<circle cx="10" cy="10" r="5" fill="#00F"/>'
                             '<ellipse cx="30" cy="10" rx="8" ry="4" fill="#0F0"/>'
                             '<line x1="0" y1="0" x2="10" y2="10" stroke="#000"/>'
                             '<polygon points="0,0 10,0 5,8" fill="#F00"/>').items
    assert (items[0].cx, items[0].rx, items[0].ry) == (10, 5, 5)
    assert (items[1].rx, items[1].ry) == (8, 4)
    assert items[2].points == [(0, 0), (10, 10)] and items[2].stroke == "#000000"
    assert items[3].closed and items[3].points[0] == items[3].points[-1] and items[3].fill == "#FF0000"


def test_path_lines_and_close():
    (pts, closed), = parse_path("M10 10 H 30 V 30 h-20 z")
    assert closed and pts == [(10, 10), (30, 10), (30, 30), (10, 30)]


def test_path_cubic_is_flattened_and_ends_exactly():
    (pts, closed), = parse_path("M0 0 C 0 50, 100 50, 100 0")
    assert not closed and len(pts) > 8 and pts[-1] == (100, 0)


def test_path_arc_reaches_endpoint_and_stays_on_circle():
    (pts, _), = parse_path("M 0 0 A 10 10 0 0 1 20 0")
    assert pts[-1] == (20, 0)
    assert all(math.dist((10, 0), p) == pytest.approx(10, abs=1e-6) for p in pts)
    assert min(y for _, y in pts) < -9   # sweep=1 goes over the top (y is down)


def test_path_arc_flags_without_separators():
    (pts, _), = parse_path("M0 0a10 10 0 01 20 0")
    assert pts[-1] == pytest.approx((20, 0))


def test_path_multiple_subpaths_become_one_path(tmp_path):
    item, = svg_of(tmp_path, '<path d="M0 0 L5 5 M10 10 L20 20" stroke="#000" fill="none"/>').items
    assert isinstance(item, Path) and item.subpaths == [[(0, 0), (5, 5)], [(10, 10), (20, 20)]]


def test_ring_path_keeps_both_subpaths_closed(tmp_path):
    item, = svg_of(tmp_path, '<path d="M0 0H20V20H0Z M5 5H15V15H5Z" fill="#000"/>').items
    assert isinstance(item, Path) and len(item.subpaths) == 2
    assert all(sub[0] == sub[-1] for sub in item.subpaths) and item.fill == "#000000"


def test_filled_open_path_is_closed(tmp_path):
    p = svg_of(tmp_path, '<path d="M0 0 L10 0 L5 8" fill="#123456"/>').items[0]
    assert p.closed and p.points[0] == p.points[-1]


def test_malformed_path_is_skipped_not_fatal(tmp_path):
    doc = svg_of(tmp_path, '<path d="L 1" stroke="#000"/><rect width="5" height="5" x="1" y="1" fill="#000"/>')
    assert len(doc.items) == 1 and any("path" in s for s in doc.skipped)


def test_style_attribute_and_css_rules_and_precedence(tmp_path):
    doc = svg_of(tmp_path, '<style>.a{fill:#00FF00;stroke:#0000FF} #r{stroke-width:5}</style>'
                           '<rect class="a" id="r" x="1" y="1" width="4" height="4" fill="#FF0000" '
                           'style="stroke:#000000"/>')
    r = doc.items[0]
    assert r.fill == "#00FF00"          # CSS beats the presentation attribute
    assert r.stroke == "#000000"        # style attribute beats CSS
    assert r.stroke_width == 5          # #id rule


def test_opacity_multiplies_through_groups(tmp_path):
    r = svg_of(tmp_path, '<g opacity="0.5"><rect x="1" y="1" width="4" height="4" fill="#000" '
                         'fill-opacity="0.5"/></g>').items[0]
    assert r.fill_opacity == pytest.approx(0.25)


def test_gradient_fill_uses_average_colour(tmp_path):
    r = svg_of(tmp_path, '<defs><linearGradient id="g"><stop offset="0" stop-color="#000000"/>'
                         '<stop offset="1" stop-color="#FFFFFF"/></linearGradient></defs>'
                         '<rect x="1" y="1" width="4" height="4" fill="url(#g)"/>').items[0]
    assert r.fill == "#808080"


def test_use_of_shape_and_symbol(tmp_path):
    doc = svg_of(tmp_path, '<defs><rect id="s" width="5" height="5" fill="#000"/>'
                           '<symbol id="y"><circle cx="2" cy="2" r="2" fill="#F00"/></symbol></defs>'
                           '<use xlink:href="#s" x="10" y="20"/><use href="#y" x="50"/>')
    assert (doc.items[0].x, doc.items[0].y) == (10, 20)
    assert doc.items[1].cx == 52


def test_use_cycle_does_not_hang(tmp_path):
    doc = svg_of(tmp_path, '<g id="a"><use href="#a"/></g>')
    assert doc.items == [] and any("deeply" in s for s in doc.skipped)


def test_display_none_and_defs_are_not_drawn(tmp_path):
    doc = svg_of(tmp_path, '<defs><rect width="5" height="5"/></defs><rect width="5" height="5" display="none"/>'
                           '<g style="display:none"><rect width="5" height="5"/></g>')
    assert doc.items == []


def test_unsupported_features_are_reported(tmp_path):
    doc = svg_of(tmp_path, '<rect x="1" y="1" width="5" height="5" fill="#000" filter="url(#f)"/><foreignObject/>')
    assert len(doc.items) == 1
    assert any("filter" in s for s in doc.skipped) and any("foreignObject" in s for s in doc.skipped)


def test_viewbox_scale_and_offset_apply(tmp_path):
    doc = svg_of(tmp_path, '<rect x="10" y="10" width="10" height="10" fill="#000"/>',
                 'width="200" height="200" viewBox="10 10 100 100"')
    r = doc.items[0]
    assert (r.x, r.y, r.width) == (0, 0, 20)


def test_units_in_lengths(tmp_path):
    doc = svg_of(tmp_path, "", 'width="1in" height="2.54cm"')
    assert (doc.width, doc.height) == (96, pytest.approx(96))


def test_text_in_scaled_group_scales_font_and_position(tmp_path):
    t = svg_of(tmp_path, '<g transform="translate(5,5) scale(2)"><text x="10" y="10" font-size="8">Hi</text></g>').items[0]
    assert (t.x, t.y, t.font_size) == (25, 25, 16)


def test_rotated_text_has_rotation(tmp_path):
    t = svg_of(tmp_path, '<text x="10" y="10" transform="rotate(-90 10 10)">Y</text>').items[0]
    assert t.rotation == pytest.approx(-90)
    assert (t.x, t.y) == pytest.approx((10, 10))


def test_text_tspans_with_absolute_y(tmp_path):
    t = svg_of(tmp_path, '<text x="5" font-size="10"><tspan x="5" y="20">a</tspan>'
                         '<tspan x="5" y="35">b</tspan></text>').items[0]
    assert t.lines == ["a", "b"] and t.y == 20 and t.line_height == 15


def test_marker_start_and_end(tmp_path):
    p = svg_of(tmp_path, '<polyline points="0,0 10,0" stroke="#000" fill="none" marker-start="url(#m)" '
                         'marker-end="url(#m)"/>').items[0]
    assert p.arrow_start and p.arrow_end


@pytest.mark.parametrize("text,expected", [
    ("translate(5 6)", (1, 0, 0, 1, 5, 6)), ("scale(2)", (2, 0, 0, 2, 0, 0)),
    ("translate(1,2) scale(3)", (3, 0, 0, 3, 1, 2)), ("matrix(1 2 3 4 5 6)", (1, 2, 3, 4, 5, 6)),
])
def test_parse_transform(text, expected):
    assert parse_transform(text) == pytest.approx(expected)


def test_rotate_about_point():
    from visio_mcp_server.svg_import import mat_apply
    assert mat_apply(parse_transform("rotate(90 10 10)"), 20, 10) == pytest.approx((10, 20))


def test_switch_without_readable_label_uses_fallback_text(tmp_path):
    doc = svg_of(tmp_path, '<switch><foreignObject width="10" height="10"><div xmlns="http://www.w3.org/1999/xhtml">'
                           '</div></foreignObject><text x="5" y="9" font-size="10">plain</text></switch>')
    assert [type(i) for i in doc.items] == [Text] and doc.items[0].lines == ["plain"]
    assert doc.skipped == []


def test_switch_without_fallback_reports_foreign_object(tmp_path):
    doc = svg_of(tmp_path, '<switch><foreignObject width="10" height="10"/></switch>')
    assert doc.items == [] and any("foreignObject" in s for s in doc.skipped)


def test_editor_namespaces_are_ignored_silently(tmp_path):
    p = tmp_path / "e.svg"
    p.write_text('<svg xmlns="http://www.w3.org/2000/svg" xmlns:sodipodi="http://sodipodi.sourceforge.net/DTD/'
                 'sodipodi-0.dtd" width="10" height="10"><sodipodi:namedview id="n"/>'
                 '<rect x="1" y="1" width="3" height="3" fill="#000"/></svg>')
    doc = parse_svg(str(p))
    assert len(doc.items) == 1 and doc.skipped == []


def test_switch_prefers_the_html_label_over_truncated_fallback(tmp_path):
    body = ('<g transform="translate(0.5,0.5)"><switch><foreignObject width="100%" height="100%">'
            '<div xmlns="http://www.w3.org/1999/xhtml" style="display: flex; align-items: unsafe center; '
            'justify-content: unsafe center; width: 100px; height: 1px; padding-top: 40px; margin-left: 10px;">'
            '<div style="text-align: center; font-size: 0px;"><div style="display: inline-block; font-size: 12px; '
            'font-family: Helvetica; color: #112233; line-height: 1.2; white-space: nowrap;">find<br/>targets</div>'
            '</div></div></foreignObject><text x="60" y="44" font-size="12">find...</text></switch></g>')
    doc = svg_of(tmp_path, body)
    t, = doc.items
    assert t.lines == ["find", "targets"] and t.anchor == "middle" and t.fill == "#112233"
    assert t.x == pytest.approx(60.5) and t.line_height == pytest.approx(14.4)
    assert t.font_family == "Helvetica"


def test_switch_draws_only_the_first_accepted_child(tmp_path):
    doc = svg_of(tmp_path, '<switch><g requiredFeatures="http://www.w3.org/TR/SVG11/feature#Extensibility"/>'
                           '<a><text x="1" y="9">Text is not SVG - cannot display</text></a></switch>')
    assert doc.items == []


def test_switch_skips_unsupported_extension_and_language(tmp_path):
    doc = svg_of(tmp_path, '<switch><text requiredExtensions="http://example.com/x" x="1" y="9">a</text>'
                           '<text systemLanguage="de" x="1" y="9">b</text>'
                           '<text systemLanguage="en-US" x="1" y="9">c</text><text x="1" y="9">d</text></switch>')
    assert [t.lines for t in doc.items] == [["c"]]


MARKER = ('<defs><marker id="a" markerWidth="{w}" markerHeight="8" refX="9" refY="4" orient="auto" '
          'markerUnits="{u}"><path d="M0,0 L10,4 L0,8 z" fill="{f}"/></marker></defs>')


def arrow_line(tmp_path, w, u, f, width=2):
    return svg_of(tmp_path, MARKER.format(w=w, u=u, f=f) +
                  f'<polyline points="0,0 50,0" stroke="#000" stroke-width="{width}" fill="none" '
                  'marker-end="url(#a)"/>').items[0]


def test_marker_size_in_user_space_units(tmp_path):
    p = arrow_line(tmp_path, 10, "userSpaceOnUse", "#000")
    assert p.arrow_end and p.arrow_size == 10 and not p.arrow_open


def test_marker_size_in_stroke_width_units(tmp_path):
    assert arrow_line(tmp_path, 4, "strokeWidth", "#000", width=3).arrow_size == 12


def test_marker_without_fill_is_open(tmp_path):
    assert arrow_line(tmp_path, 10, "userSpaceOnUse", "none").arrow_open


def test_missing_marker_has_unknown_size(tmp_path):
    p = svg_of(tmp_path, '<polyline points="0,0 9,0" stroke="#000" fill="none" marker-end="url(#zzz)"/>').items[0]
    assert p.arrow_end and p.arrow_size == 0


ICON_SVG = ('<svg xmlns="http://www.w3.org/2000/svg" width="80" height="40" viewBox="0 0 80 40">'
            '<rect x="20" y="10" width="40" height="20" fill="#FF0000"/></svg>')


def icon_href(svg=ICON_SVG):
    return "data:image/svg+xml;base64," + base64.b64encode(svg.encode()).decode()


def test_svg_icon_is_rasterised_without_edge(tmp_path, monkeypatch):
    pytest.importorskip("resvg_py")
    from PIL import Image as PILImage
    monkeypatch.setattr("visio_mcp_server.svg_import.find_edge", lambda: None)
    path = image_to_png(icon_href(), str(tmp_path), str(tmp_path), "icon")
    img = PILImage.open(path).convert("RGBA")
    assert img.width == 320 and img.height == 160                      # aspect kept
    assert img.getpixel((2, 2))[3] == 0                                 # transparent outside the drawing
    assert img.getpixel((img.width // 2, img.height // 2)) == (255, 0, 0, 255)


def test_svg_icon_without_any_rasteriser_is_a_clear_error(tmp_path, monkeypatch):
    monkeypatch.setattr("visio_mcp_server.svg_import._rasterize_with_resvg", lambda *a: False)
    monkeypatch.setattr("visio_mcp_server.svg_import.find_edge", lambda: None)
    with pytest.raises(SvgError, match="resvg-py"):
        image_to_png(icon_href(), str(tmp_path), str(tmp_path), "icon")
