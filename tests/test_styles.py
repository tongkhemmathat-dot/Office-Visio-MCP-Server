import pytest

from visio_mcp_server.styles import (
    GROUP_SHAPE_TYPE,
    apply_shape_style,
    color_formula,
    line_pattern_value,
    style_geometry,
    validate_style,
)


class FakeCell:
    def __init__(self):
        self.FormulaForceU = None


class FakeShape:
    """Records the formulas written to its cells; mimics shape.Cells(name)."""

    def __init__(self, shape_type=1, children=()):
        self.Type = shape_type
        self.Shapes = list(children)
        self.cells = {}

    def Cells(self, name):
        return self.cells.setdefault(name, FakeCell())

    def formula(self, name):
        return self.cells[name].FormulaForceU


@pytest.mark.parametrize("value, expected", [
    ("#2563EB", "RGB(37,99,235)"),
    ("2563eb", "RGB(37,99,235)"),
    ("  Red ", "RGB(220,38,38)"),
    ("white", "RGB(255,255,255)"),
])
def test_color_formula(value, expected):
    assert color_formula(value) == expected


@pytest.mark.parametrize("value", ["notacolor", "#12345", "#GGGGGG", ""])
def test_color_formula_rejects_invalid(value):
    with pytest.raises(ValueError):
        color_formula(value)


def test_line_pattern_value():
    assert line_pattern_value("Dash") == "2"
    assert line_pattern_value("dotted") == "3"
    with pytest.raises(ValueError):
        line_pattern_value("wavy")


def test_validate_style_checks_every_field():
    validate_style("red", "#000000", "dash", "blue")
    with pytest.raises(ValueError):
        validate_style(fill_color="nope")
    with pytest.raises(ValueError):
        validate_style(line_pattern="wavy")


def test_apply_shape_style_only_touches_given_options():
    shape = FakeShape()
    apply_shape_style(shape, fill_color="red", line_weight=2, bold=True)
    assert shape.formula("FillForegnd") == "RGB(220,38,38)"
    assert shape.formula("FillPattern") == "1"
    assert shape.formula("FillGradientEnabled") == "0"
    assert shape.formula("LineWeight") == "2 pt"
    assert shape.formula("Char.Style") == "1"
    assert "LineColor" not in shape.cells
    assert "Char.Size" not in shape.cells


def test_style_geometry_descends_into_groups():
    child_a, child_b = FakeShape(), FakeShape()
    group = FakeShape(GROUP_SHAPE_TYPE, [child_a, FakeShape(GROUP_SHAPE_TYPE, [child_b])])
    style_geometry(group, fill_color="#112233")
    assert child_a.formula("FillForegnd") == "RGB(17,34,51)"
    assert child_b.formula("FillForegnd") == "RGB(17,34,51)"
    assert "FillForegnd" not in group.cells  # the group itself is not styled


def test_style_geometry_skips_sub_shapes_that_refuse_the_cell():
    class Locked(FakeShape):
        def Cells(self, name):
            raise RuntimeError("Cell is guarded.")

    ok = FakeShape()
    group = FakeShape(GROUP_SHAPE_TYPE, [Locked(), ok])
    style_geometry(group, line_color="red")
    assert ok.formula("LineColor") == "RGB(220,38,38)"
