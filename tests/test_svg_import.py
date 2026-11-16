import base64

import pytest

from visio_mcp_server.svg_import import (Image, Polyline, Rect, SvgError, Text, image_to_png, parse_color,
                                         parse_svg)

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
    assert kinds == [Rect, Polyline, Text, Text, Image]


def test_rect_style(svg_file):
    rect = parse_svg(svg_file).items[0]
    assert (rect.x, rect.y, rect.width, rect.height) == (10, 20, 50, 30)
    assert rect.fill is None and rect.stroke == "#00A4A6" and rect.stroke_width == 2 and rect.dashed


def test_polyline_points_and_arrow(svg_file):
    line = parse_svg(svg_file).items[1]
    assert line.points == [(10, 70), (90.5, 70)]
    assert line.arrow_end and not line.dashed


def test_text_lines_anchor_and_halo(svg_file):
    text, halo = parse_svg(svg_file).items[2:4]
    assert text.lines == ["Line one", "Line two"] and text.line_height == 16
    assert text.anchor == "middle" and text.bold
    assert halo.lines == ["Halo"] and halo.halo and halo.fill == "#000000"


def test_unsupported_elements_are_reported(svg_file):
    assert parse_svg(svg_file).skipped == ["<circle>"]


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
