import pytest

pytest.importorskip("win32com")

from visio_mcp_server import svg_import  # noqa: E402
from visio_mcp_server.visio_server import svg_group_runs, svg_layer_name, svg_unique_name  # noqa: E402


def test_unique_names_are_safe_and_do_not_collide():
    used = set()
    assert svg_unique_name("web server", used) == "web_server"
    assert svg_unique_name("Web Server", used) == "Web_Server_2"      # case-insensitive clash
    assert svg_unique_name("web server", used) == "web_server_3"
    assert svg_unique_name("", used) == "" and svg_unique_name("***", used) == ""


def test_group_runs_only_consecutive_and_two_or_more():
    groups = ["", "a", "a", "a", "b", "", "a", "a", "c"]
    assert svg_group_runs(groups) == [("a", 1, 3), ("a", 6, 7)]
    assert svg_group_runs([]) == [] and svg_group_runs(["x"]) == []


def test_layers():
    poly = lambda **k: svg_import.Polyline([(0, 0), (1, 1)], "#000000", 1, None, **k)
    assert svg_layer_name(poly()) == "Connectors"
    assert svg_layer_name(poly(closed=True, fill="#FFFFFF")) == "Shapes"
    assert svg_layer_name(svg_import.Text(["a"], 0, 0, 12, "start", "#000000", False, 14, False)) == "Labels"
    assert svg_layer_name(svg_import.Image(0, 0, 1, 1, "x")) == "Icons"
    assert svg_layer_name(svg_import.Rect(0, 0, 1, 1, None, None, 1, None)) == "Shapes"
