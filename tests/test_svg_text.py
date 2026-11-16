import sys

import pytest

from visio_mcp_server import svg_text


def test_generic_families_expand_to_windows_fonts():
    assert "Times New Roman" in svg_text.font_candidates("serif")
    assert "Times New Roman" in svg_text.font_candidates("Georgia, serif")
    assert "Arial" in svg_text.font_candidates("Helvetica Neue, Helvetica, sans-serif")
    assert "Consolas" in svg_text.font_candidates("'Fira Code', monospace")


def test_candidates_keep_order_and_always_end_with_a_default():
    c = svg_text.font_candidates("Leelawadee UI, Tahoma")
    assert c[:2] == ["Leelawadee UI", "Tahoma"] and c[-1] == "Arial"
    assert svg_text.font_candidates("") == ["Arial"]


def test_unknown_font_falls_back_to_estimate():
    w = svg_text.text_width("hello", "NoSuchFontAnywhere123", 10)
    assert 20 <= w <= 40   # 5 chars * ~0.6 em, or Arial if installed


def test_empty_text_has_no_width():
    assert svg_text.text_width("", "Arial", 12) == 0


@pytest.mark.skipif(sys.platform != "win32", reason="needs Windows fonts")
def test_real_font_metrics_scale_and_bold_is_wider():
    if svg_text.find_font_file("Arial") is None:
        pytest.skip("Arial is not installed")
    w10, w20 = svg_text.text_width("Network", "Arial", 10), svg_text.text_width("Network", "Arial", 20)
    assert w20 == pytest.approx(2 * w10, rel=0.02)
    assert svg_text.text_width("Network", "Arial", 10, bold=True) > w10
    assert svg_text.text_width("iiii", "Arial", 10) < svg_text.text_width("WWWW", "Arial", 10)   # not len-based


@pytest.mark.skipif(sys.platform != "win32", reason="needs Windows fonts")
def test_thai_text_is_measured_with_a_thai_font():
    if svg_text.find_font_file("Leelawadee UI") is None:
        pytest.skip("Leelawadee UI is not installed")
    assert svg_text.text_width("ผู้ใช้งานเว็บ", "Leelawadee UI", 13) > 20
