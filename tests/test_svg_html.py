import xml.etree.ElementTree as ET

import pytest

from visio_mcp_server.svg_html import parse_label

NS = 'xmlns="http://www.w3.org/2000/svg" xmlns:h="http://www.w3.org/1999/xhtml"'


def label(inner_div: str, flex="width: 100px; height: 1px; padding-top: 50px; margin-left: 20px;",
          justify="center", items="center"):
    xml = (f'<foreignObject {NS} width="100%" height="100%"><div xmlns="http://www.w3.org/1999/xhtml" '
           f'style="display: flex; align-items: unsafe {items}; justify-content: unsafe {justify}; {flex}">'
           f'<div style="box-sizing: border-box; font-size: 0px; text-align: center;">{inner_div}</div></div>'
           f'</foreignObject>')
    return parse_label(ET.fromstring(xml))


def span(text, extra=""):
    return (f'<div style="display: inline-block; font-size: 12px; font-family: Helvetica; color: rgb(10, 20, 30); '
            f'line-height: 1.2; white-space: normal; {extra}">{text}</div>')


def test_centered_label_position_and_style():
    lb = label(span("Hello"))
    assert lb.lines == ["Hello"] and lb.font_size == 12 and lb.font_family == "Helvetica"
    assert lb.color == "rgb(10, 20, 30)" and not lb.bold and lb.align == "center"
    assert lb.anchor_x == pytest.approx(70)              # margin-left 20 + width 100 / 2
    assert lb.line_height == pytest.approx(14.4)
    assert lb.baseline_y == pytest.approx(50 + 0.5 - 7.2 + 7.2 + 0.35 * 12)   # centred on padding-top + 0.5


def test_br_and_nested_divs_make_lines():
    lb = label(span("one<br/>two<div>three</div>"))
    assert lb.lines == ["one", "two", "three"]


def test_text_wraps_to_box_width():
    lb = label(span("alpha beta gamma delta epsilon zeta eta theta"), flex="width: 60px; height: 1px; "
               "padding-top: 50px; margin-left: 20px;")
    assert len(lb.lines) > 2 and " ".join(lb.lines) == "alpha beta gamma delta epsilon zeta eta theta"


def test_nowrap_keeps_one_line_and_overflows_symmetrically():
    lb = label(span("a fairly long label that does not wrap", "white-space: nowrap;"),
               flex="width: 1px; height: 1px; padding-top: 10px; margin-left: 100px;")
    assert len(lb.lines) == 1 and lb.anchor_x == pytest.approx(100.5)


def test_left_aligned_block_anchors_at_left_edge():
    xml = ('<foreignObject xmlns="http://www.w3.org/2000/svg"><div xmlns="http://www.w3.org/1999/xhtml" '
           'style="display: flex; align-items: unsafe flex-start; justify-content: unsafe flex-start; '
           'width: 100px; height: 1px; padding-top: 5px; margin-left: 30px;"><div style="text-align: left; '
           'font-size: 0px;"><div style="font-size: 12px; text-align: left;">left text</div></div></div></foreignObject>')
    lb = parse_label(ET.fromstring(xml))
    assert lb.align == "left" and lb.anchor_x == pytest.approx(30) and lb.baseline_y > 5


def test_bold_and_font_colour():
    lb = label(span('<b><font color="#ffffff">White bold</font></b>'))
    assert lb.bold and lb.color == "#ffffff"


def test_empty_label_is_none():
    assert label(span("<br/>")) is None


def test_list_items_get_bullets():
    lb = label(span("<ul><li>DNS</li><li>Consul</li></ul>", "text-align: left;"))
    assert lb.lines == ["• DNS", "• Consul"]
