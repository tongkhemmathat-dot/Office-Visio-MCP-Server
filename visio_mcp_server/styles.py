"""Style helpers for Visio shapes and connectors.

Everything here works on plain strings and on shape objects passed in, so the
parsing and validation can be tested without Visio installed.
"""

from typing import Optional

NAMED_COLORS = {
    "black": "000000", "white": "FFFFFF", "red": "DC2626", "green": "16A34A",
    "blue": "2563EB", "yellow": "FACC15", "orange": "EA580C", "purple": "7C3AED",
    "teal": "0D9488", "gray": "6B7280", "grey": "6B7280",
}

# Visio's LinePattern cell values
LINE_PATTERNS = {"solid": 1, "dash": 2, "dashed": 2, "dot": 3, "dotted": 3, "dashdot": 4}

GROUP_SHAPE_TYPE = 2  # Shape.Type value for groups (visTypeGroup)


def color_formula(color: str) -> str:
    """Convert '#RRGGBB', 'RRGGBB' or a basic color name to a Visio RGB() formula."""
    value = NAMED_COLORS.get(color.strip().lower(), color.strip()).lstrip("#")
    if len(value) != 6 or any(c not in "0123456789abcdefABCDEF" for c in value):
        raise ValueError(f"Invalid color '{color}'. Use #RRGGBB or a name like red, blue, teal.")
    r, g, b = (int(value[i:i + 2], 16) for i in (0, 2, 4))
    return f"RGB({r},{g},{b})"


def line_pattern_value(pattern: str) -> str:
    """Convert a pattern name (solid, dash, dot, dashdot) to Visio's LinePattern value."""
    key = pattern.strip().lower()
    if key not in LINE_PATTERNS:
        raise ValueError(f"Invalid line_pattern '{pattern}'. Use one of: {', '.join(sorted(LINE_PATTERNS))}.")
    return str(LINE_PATTERNS[key])


def validate_style(fill_color: Optional[str] = None, line_color: Optional[str] = None,
                   line_pattern: Optional[str] = None, text_color: Optional[str] = None) -> None:
    """Raise ValueError for bad style input before anything is drawn."""
    for color in (fill_color, line_color, text_color):
        if color is not None:
            color_formula(color)
    if line_pattern is not None:
        line_pattern_value(line_pattern)


def apply_shape_style(shape, fill_color=None, line_color=None, line_weight=None, line_pattern=None,
                      text_color=None, font_size=None, bold=None, rounding=None) -> None:
    """Apply the given style options to a shape; options left as None are not touched.

    FormulaForceU is used so guarded cells (common in stencil shapes) can be set too.
    """
    if fill_color is not None:
        shape.Cells("FillForegnd").FormulaForceU = color_formula(fill_color)
        shape.Cells("FillPattern").FormulaForceU = "1"
        # Stencil shapes often use gradient fills, which hide a plain fill color
        try:
            shape.Cells("FillGradientEnabled").FormulaForceU = "0"
        except Exception:
            pass
    if line_color is not None:
        shape.Cells("LineColor").FormulaForceU = color_formula(line_color)
    if line_weight is not None:
        shape.Cells("LineWeight").FormulaForceU = f"{line_weight} pt"
    if line_pattern is not None:
        shape.Cells("LinePattern").FormulaForceU = line_pattern_value(line_pattern)
    if text_color is not None:
        shape.Cells("Char.Color").FormulaForceU = color_formula(text_color)
    if font_size is not None:
        shape.Cells("Char.Size").FormulaForceU = f"{font_size} pt"
    if bold is not None:
        shape.Cells("Char.Style").FormulaForceU = "1" if bold else "0"
    if rounding is not None:
        shape.Cells("Rounding").FormulaForceU = f"{rounding} in"


def style_geometry(shape, fill_color=None, line_color=None, line_weight=None, line_pattern=None) -> None:
    """Apply fill/line options to a shape, descending into groups.

    Stencil shapes are usually groups, so the options go on every sub-shape.
    Sub-shapes without the relevant cells (e.g. locked ones) are skipped.
    """
    if shape.Type == GROUP_SHAPE_TYPE:
        for sub in shape.Shapes:
            style_geometry(sub, fill_color, line_color, line_weight, line_pattern)
        return
    try:
        apply_shape_style(shape, fill_color=fill_color, line_color=line_color,
                          line_weight=line_weight, line_pattern=line_pattern)
    except Exception:
        pass
