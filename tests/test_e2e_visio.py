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


def test_connectors_to_one_d_shapes_and_images_do_not_shift_the_page(diagram):
    """Rack-equipment masters are 1-D (no glue target) and a Synology master is an embedded image.

    Connecting to them must work and must not move existing shapes. Note: the page-origin shift
    that motivated dropping connectors on their first shape was only seen on a large, dense page
    (26 x 38 in, ~75 shapes); this small page does not reproduce it, so that part is not guarded here.
    """
    a = shape_id(run(server.add_shape(diagram, "Rectangle", 2, 2, 1, 1, text="A")))
    one_d = shape_id(run(server.add_stencil_shape(diagram, "HPE-ProLiant-DL", "DL380 Gen10 SFF front", 6, 5)))
    image = shape_id(run(server.add_stencil_shape(diagram, "Synology", "SA3400 Front", 6, 3)))

    page = server.get_page(server.get_document(diagram))
    before = page.Shapes.ItemFromID(a).Cells("PinX").ResultIU
    for target in (one_d, image):
        assert "connected" in run(server.connect_shapes(diagram, a, target, "Straight", label="link"))
    assert page.Shapes.ItemFromID(a).Cells("PinX").ResultIU == before


def test_connector_labels_have_an_opaque_background(diagram):
    a = shape_id(run(server.add_shape(diagram, "Rectangle", 1, 5)))
    b = shape_id(run(server.add_shape(diagram, "Rectangle", 6, 5)))
    run(server.connect_shapes(diagram, a, b, label="HTTPS 443"))
    page = server.get_page(server.get_document(diagram))
    connectors = [s for s in page.Shapes if s.OneD and s.Text == "HTTPS 443"]
    assert connectors and connectors[0].Cells("TextBkgndTrans").FormulaU == "0%"


def test_raster_export_failure_points_to_pdf(diagram, tmp_path, monkeypatch):
    class Broken:
        def Export(self, path):
            raise RuntimeError("boom")
    monkeypatch.setattr(server, "get_page", lambda doc: Broken())
    result = run(server.export_diagram(diagram, str(tmp_path / "x.png")))
    assert result.startswith("Error exporting diagram:") and "pdf or .svg" in result


def test_svg_to_visio_creates_native_shapes(tmp_path):
    svg = tmp_path / "d.svg"
    svg.write_text(
        '<svg xmlns="http://www.w3.org/2000/svg" width="480" height="288" viewBox="0 0 480 288">'
        '<rect width="480" height="288" fill="#FFFFFF"/>'
        '<rect x="96" y="96" width="192" height="96" fill="#E8F6F7" stroke="#00A4A6" stroke-width="2"/>'
        '<polyline points="96,48 288,48" fill="none" stroke="#232F3E" stroke-width="2" marker-end="url(#m)"/>'
        '<text x="192" y="150" font-size="13" text-anchor="middle" fill="#232F3E">'
        '<tspan x="192" dy="0">Hello</tspan><tspan x="192" dy="16">World</tspan></text></svg>',
        encoding="utf8")
    out = str(tmp_path / "d.vsdx")
    result = run(server.svg_to_visio(str(svg), out, export_png=str(tmp_path / "d.png")))
    try:
        assert "1 rect, 1 polyline, 1 text" in result
        page = server.get_page(server.get_document(out))
        assert abs(page.PageSheet.Cells("PageWidth").ResultIU - 5.0) < 1e-6
        assert abs(page.PageSheet.Cells("PageHeight").ResultIU - 3.0) < 1e-6
        rect = next(s for s in page.Shapes if s.Text == "" and s.CellsU("FillPattern").ResultIU == 1)
        # SVG y=96..192 of 288 -> Visio y (flipped) = 1.0..2.0 in; x = 1.0..3.0 in
        assert abs(rect.CellsU("PinX").ResultIU - 2.0) < 1e-6 and abs(rect.CellsU("PinY").ResultIU - 1.5) < 1e-6
        assert any(s.Text == "Hello\nWorld" for s in page.Shapes)
        assert os.path.exists(tmp_path / "d.png")
    finally:
        run(server.close_document(out))


def test_svg_to_visio_groups_transforms_paths_and_opacity(tmp_path):
    svg = tmp_path / "g.svg"
    svg.write_text(
        '<svg xmlns="http://www.w3.org/2000/svg" width="480" height="288">'
        '<g transform="translate(96,96)" opacity="0.5">'
        '<rect width="96" height="48" rx="10" fill="#FF0000"/>'
        '<circle cx="144" cy="24" r="24" fill="#0000FF"/></g>'
        '<path d="M 0 288 L 96 192 L 192 288 Z" fill="#00FF00" stroke="#000000" stroke-width="2"/>'
        '<text x="400" y="200" font-size="13" transform="rotate(-90 400 200)">Rot</text></svg>',
        encoding="utf8")
    out = str(tmp_path / "g.vsdx")
    result = run(server.svg_to_visio(str(svg), out))
    try:
        assert "1 rect, 1 ellipse, 1 polyline, 1 text" in result
        page = server.get_page(server.get_document(out))
        shapes = list(page.Shapes)
        rect = next(s for s in shapes if abs(s.CellsU("Rounding").ResultIU - 10 / 96) < 1e-6)
        # translate(96,96) -> x 1..2 in, y 96..144 of 288 -> Visio y 1.0..1.5? (3 - 1.5 .. 3 - 1)
        assert abs(rect.CellsU("PinX").ResultIU - 1.5) < 1e-6 and abs(rect.CellsU("PinY").ResultIU - 1.75) < 1e-6
        assert abs(rect.CellsU("FillForegndTrans").ResultIU - 0.5) < 1e-6
        text = next(s for s in shapes if s.Text == "Rot")
        assert abs(text.CellsU("Angle").ResultIU - 3.141592653589793 / 2) < 1e-6   # -90 deg clockwise
    finally:
        run(server.close_document(out))


def test_svg_to_visio_cuts_holes_in_multi_subpath_paths(tmp_path):
    from PIL import Image
    svg = tmp_path / "ring.svg"
    svg.write_text('<svg xmlns="http://www.w3.org/2000/svg" width="288" height="288">'
                   '<path d="M48 48H240V240H48Z M96 96H192V192H96Z" fill="#FF0000"/></svg>', encoding="utf8")
    out, png = str(tmp_path / "ring.vsdx"), str(tmp_path / "ring.png")
    result = run(server.svg_to_visio(str(svg), out, export_png=png))
    try:
        assert "1 path" in result
        img = Image.open(png).convert("RGB")
        w, h = img.size
        assert img.getpixel((w // 2, h // 2)) == (255, 255, 255)           # the hole
        row = [img.getpixel((x, h // 2)) for x in range(w)]                   # a red ring exists on both sides
        red = [x for x, (r, g, b) in enumerate(row) if r > 200 and g < 60 and b < 60]
        assert red and min(red) < w // 2 < max(red)
    finally:
        run(server.close_document(out))


def test_svg_to_visio_names_layers_and_groups(tmp_path):
    svg = tmp_path / "n.svg"
    svg.write_text(
        '<svg xmlns="http://www.w3.org/2000/svg" width="480" height="288">'
        '<rect id="lonely" x="10" y="10" width="40" height="20" fill="#CCCCCC"/>'
        '<g id="tier"><rect id="box" x="100" y="100" width="60" height="40" fill="#E8F6F7"/>'
        '<text id="caption" x="130" y="125" font-size="13" text-anchor="middle">Web</text></g>'
        '<polyline points="160,120 300,120" stroke="#000" fill="none"/></svg>', encoding="utf8")
    out = str(tmp_path / "n.vsdx")
    result = run(server.svg_to_visio(str(svg), out))
    try:
        assert "Grouped 1 SVG group" in result
        page = server.get_page(server.get_document(out))
        names = {s.NameU: s for s in page.Shapes}
        assert "lonely" in names and "tier" in names
        group = names["tier"]
        assert group.Type == 2 and group.Shapes.Count == 2                      # visTypeGroup, box + caption
        assert {s.NameU for s in group.Shapes} == {"box", "caption"}
        assert {layer.Name for layer in page.Layers} >= {"Shapes", "Labels", "Connectors"}
    finally:
        run(server.close_document(out))
