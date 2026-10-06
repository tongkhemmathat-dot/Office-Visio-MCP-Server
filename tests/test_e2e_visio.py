"""End-to-end tests that drive a real Visio. They are skipped when Visio is not installed.

They start (or attach to) Visio, so close any unsaved work before running them.
"""

import asyncio
import json
import os
import re
import sys

import pytest

pytestmark = pytest.mark.skipif(sys.platform != "win32", reason="Visio automation needs Windows")

if sys.platform == "win32":
    import visio_mcp_server.visio_server as server
    visio_missing = not server.check_visio_installed()
else:  # pragma: no cover
    server = None
    visio_missing = True

pytestmark = [pytestmark, pytest.mark.skipif(visio_missing, reason="Microsoft Visio is not installed")]


def run(coro):
    return asyncio.run(coro)


def shape_id(result: str) -> int:
    return int(re.search(r"ID (\d+)", result).group(1))


@pytest.fixture(scope="module", autouse=True)
def visio_session():
    yield
    server.close_visio_app()


@pytest.fixture
def diagram(tmp_path):
    path = str(tmp_path / "test.vsdx")
    run(server.create_visio_file(save_path=path))
    yield path
    run(server.close_document(path))


def test_new_files_get_a_custom_a4_landscape_page(diagram):
    sheet = server.get_page(server.get_document(diagram)).PageSheet
    assert sheet.Cells("PageWidth").ResultIU == 11.69
    assert sheet.Cells("DrawingSizeType").ResultIU == 2       # not tied to the printer paper
    assert sheet.Cells("PrintPageOrientation").ResultIU == 2   # landscape


def test_shapes_and_connectors(diagram):
    a = shape_id(run(server.add_shape(diagram, "Rectangle", 1, 5, fill_color="#2563EB", text="A")))
    b = shape_id(run(server.add_shape(diagram.upper(), "Circle", 5, 5)))   # same file, different spelling
    assert len(server.open_documents) == 1
    assert "connected" in run(server.connect_shapes(diagram, a, b, "Curved", end_arrow=True, label="x"))
    assert "connected" in run(server.connect_shapes(diagram, a, a, "Straight"))
    assert run(server.connect_shapes(diagram, a, 9999)).startswith("Error connecting shapes: Could not find")
    assert run(server.add_shape(diagram, "Rectangle", 0, 0, fill_color="nope")).startswith("Error adding shape")
    assert run(server.delete_shape(diagram, b)) == f"Shape {b} deleted"
    assert len(json.loads(run(server.list_shapes(diagram)))) >= 3


def test_build_diagram_and_exports(diagram, tmp_path):
    spec = {
        "title": "Test", "direction": "top-down",
        "nodes": [{"id": "fw", "label": "FW", "tier": 0, "stencil": "PERIPH_U", "master": "Firewall"},
                  {"id": "s1", "label": "S1", "tier": 1}, {"id": "s2", "label": "S2", "tier": 1}],
        "links": [{"from": "fw", "to": "s1", "arrow": "end"}, {"from": "fw", "to": "s2", "label": "x"}],
        "groups": [{"label": "Servers", "nodes": ["s1", "s2"]}],
    }
    result = json.loads(run(server.build_diagram(diagram, spec)))
    assert set(result["nodes"]) == {"fw", "s1", "s2"} and result["links"] == 2

    bad = run(server.build_diagram(diagram, {"nodes": [{"id": "a"}], "links": [{"from": "a", "to": "zzz"}]}))
    assert bad.startswith("Error building diagram:") and "zzz" in bad

    for name in ("out.png", "out.pdf", "out.svg"):
        out = str(tmp_path / name)
        assert run(server.export_diagram(diagram, out)) == f"Exported to: {out}"
        assert os.path.getsize(out) > 1000
    assert "Unsupported format" in run(server.export_diagram(diagram, str(tmp_path / "x.txt")))


def test_rack_workflow_and_persistence(diagram):
    rack = json.loads(run(server.add_rack(diagram, 30, 10, name="Rack A01")))["rack_id"]
    sw = run(server.add_rack_device(diagram, rack, "Juniper EX Series", "EX4300-24T Front", 42, 1, "SW1"))
    run(server.add_rack_device(diagram, rack, "HPE-ProLiant-DL", "DL380 Gen10 SFF front", 38, 2, "DL380 #1"))
    assert json.loads(sw)["u_end"] == 42

    overlap = run(server.add_rack_device(diagram, rack, "HPE-ProLiant-DL", "DL380 Gen10 SFF front", 39, 2))
    assert "overlaps 'DL380 #1'" in overlap
    assert "does not fit" in run(server.add_rack_device(diagram, rack, "HPE-ProLiant-DL",
                                                         "DL380 Gen10 SFF front", 42, 2))

    # the rack layout must survive closing and reopening the file
    run(server.close_document(diagram))
    info = json.loads(run(server.list_rack(diagram, rack)))
    assert info["u_height"] == 42 and info["used_u"] == 3 and info["free_u"] == 39
    assert [(d["label"], d["u_start"], d["u_end"]) for d in info["devices"]] == [
        ("DL380 #1", 38, 39), ("SW1", 42, 42)]
    assert info["free_ranges"] == [{"u_start": 1, "u_end": 37}, {"u_start": 40, "u_end": 41}]
    assert "not a rack" in run(server.list_rack(diagram, info["devices"][0]["shape_id"]))


def test_stencil_search():
    found = json.loads(run(server.search_stencil_shapes("DL380 Gen10", "HPE-ProLiant-DL")))
    assert found and all("DL380 Gen10" in m["master"] for m in found)
    assert run(server.search_stencil_shapes("x", "no-such-stencil")).startswith("Error searching stencil shapes")
