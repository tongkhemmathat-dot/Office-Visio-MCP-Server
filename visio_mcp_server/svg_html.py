"""Read the HTML label inside an SVG <foreignObject> (draw.io / diagrams.net export) into plain text lines.

draw.io writes labels as XHTML: a flex box that positions the label (width, margin-left, padding-top,
justify-content, align-items) around a div that carries the font, colour and alignment. This module
extracts the text, wraps it to the box width with real font metrics and works out where the lines go.
It has no Visio or SVG dependency beyond the element tree it is given.
"""
import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional

from visio_mcp_server import svg_text

XHTML = "{http://www.w3.org/1999/xhtml}"
BLOCKS = {"div", "p", "li", "h1", "h2", "h3", "h4", "h5", "h6", "tr", "ul", "ol", "table"}
BOLD_TAGS = {"b", "strong"}
DEFAULT_LINE_HEIGHT = 1.2
BASELINE_FROM_CENTRE = 0.35   # first baseline sits this many font sizes below the middle of its line box


@dataclass
class HtmlLabel:
    lines: List[str]
    font_size: float
    font_family: str
    color: Optional[str]          # raw CSS colour, None = inherit default
    bold: bool
    line_height: float            # px between baselines
    align: str                    # left | center | right  (text-align inside the label block)
    anchor_x: float               # x of the anchor for `align`
    baseline_y: float             # y of the first baseline
    box_width: float = 0.0
    extra: Dict[str, str] = field(default_factory=dict)


def _style(el) -> Dict[str, str]:
    out = {}
    for decl in (el.get("style") or "").split(";"):
        k, _, v = decl.partition(":")
        if v.strip():
            out[k.strip().lower()] = v.strip()
    return out


def _px(value: Optional[str], default: Optional[float] = None) -> Optional[float]:
    m = re.match(r"\s*(-?\d+(?:\.\d+)?)\s*(px)?\s*$", value or "")
    return float(m.group(1)) if m else default


def _local(tag) -> str:
    return tag.rsplit("}", 1)[-1].lower() if isinstance(tag, str) else ""


def _collect(el, lines: List[str], cur: List[str], pre: bool, tags: set) -> None:
    """Depth-first text extraction: <br> and block boundaries end a line, whitespace collapses."""
    tag = _local(el.tag)
    tags.add(tag)
    if tag in ("script", "style"):
        return
    block = tag in BLOCKS
    if block and "".join(cur).strip():
        lines.append("".join(cur))
        cur.clear()
    if tag == "br":
        lines.append("".join(cur))
        cur.clear()
        return
    is_pre = pre or "pre" in _style(el).get("white-space", "")
    if tag == "li":
        cur.append("• ")
    if el.text:
        cur.append(el.text if is_pre else re.sub(r"\s+", " ", el.text))
    for child in el:
        _collect(child, lines, cur, is_pre, tags)
        if child.tail:
            cur.append(child.tail if is_pre else re.sub(r"\s+", " ", child.tail))
    if block and "".join(cur).strip():
        lines.append("".join(cur))
        cur.clear()


def _wrap(text: str, family: str, size: float, bold: bool, width: float) -> List[str]:
    """Greedy word wrap to `width` px; a single over-long word stays on its own line."""
    words = text.split(" ")
    lines, cur = [], ""
    for word in words:
        trial = f"{cur} {word}" if cur else word
        if cur and svg_text.text_width(trial, family, size, bold) > width + 0.5:
            lines.append(cur)
            cur = word
        else:
            cur = trial
    lines.append(cur)
    return lines


def parse_label(fo) -> Optional[HtmlLabel]:
    """Turn a <foreignObject> into an HtmlLabel, or None if it holds no usable text."""
    root = next((c for c in fo if _local(c.tag) == "div" or _local(c.tag) == "span" or _local(c.tag) == "p"), None)
    if root is None:
        return None

    # Inherit the text properties down the first chain of elements; the last non-zero font-size wins.
    props: Dict[str, str] = {}
    chain = []
    node = root
    while node is not None:
        chain.append(node)
        for k, v in _style(node).items():
            if k == "font-size" and _px(v, 0) == 0:
                continue
            props[k] = v
        kids = [c for c in node if isinstance(c.tag, str)]
        node = kids[0] if len(kids) == 1 and not (node.text or "").strip() else None

    lines: List[str] = []
    tags: set = set()
    cur: List[str] = []
    _collect(root, lines, cur, False, tags)
    if "".join(cur).strip():
        lines.append("".join(cur))
    # <br> inside an otherwise empty block yields blank lines, which draw.io keeps as spacing
    while lines and not lines[-1].strip():
        lines.pop()
    while lines and not lines[0].strip():
        lines.pop(0)
    if not lines:
        return None

    size = _px(props.get("font-size"), 12.0)
    family = props.get("font-family", "Helvetica")
    bold = props.get("font-weight", "normal") in ("bold", "600", "700", "800", "900") or bool(tags & BOLD_TAGS)
    lh_raw = props.get("line-height", "")
    if lh_raw.endswith("px"):
        lh = _px(lh_raw, size * DEFAULT_LINE_HEIGHT)
    else:
        try:
            lh = float(lh_raw) * size
        except ValueError:
            lh = size * DEFAULT_LINE_HEIGHT
    align = props.get("text-align", "center")
    align = {"start": "left", "end": "right"}.get(align, align)
    nowrap = "nowrap" in props.get("white-space", "")

    # Geometry: draw.io's flex wrapper gives the box; otherwise use the foreignObject's own x/y/width/height.
    st = _style(root)
    if "display" in st and "flex" in st["display"]:
        box_w = _px(st.get("width"), 1.0)
        box_h = _px(st.get("height"), 1.0)
        left = _px(st.get("margin-left"), 0.0)
        top = _px(st.get("padding-top"), 0.0)
        justify = st.get("justify-content", "flex-start").replace("unsafe", "").strip()
        items = st.get("align-items", "flex-start").replace("unsafe", "").strip()
    else:
        box_w = _px(fo.get("width"), None) or _px(props.get("width"), None) or _px(props.get("max-width"), 1e6)
        box_h = _px(fo.get("height"), 1.0)
        left = _px(fo.get("x"), 0.0) + (_px(props.get("margin-left"), 0.0))
        top = _px(fo.get("y"), 0.0) + (_px(props.get("padding-top"), 0.0))
        justify = {"center": "center", "right": "flex-end"}.get(align, "flex-start")
        items = "flex-start"

    wrap_w = _px(props.get("max-width"), None) or box_w
    out: List[str] = []
    for line in lines:
        out.extend([line.strip()] if nowrap or wrap_w <= 2 else _wrap(line.strip(), family, size, bold, wrap_w))
    block_w = max(svg_text.text_width(l, family, size, bold) for l in out)

    if justify == "center":
        block_left = left + (box_w - block_w) / 2
    elif justify in ("flex-end", "end", "right"):
        block_left = left + box_w - block_w
    else:
        block_left = left
    total_h = lh * len(out)
    if items == "center":
        block_top = top + box_h / 2 - total_h / 2
    elif items in ("flex-end", "end"):
        block_top = top + box_h - total_h
    else:
        block_top = top
    anchor_x = {"left": block_left, "center": block_left + block_w / 2, "right": block_left + block_w}.get(
        align, block_left + block_w / 2)
    color = props.get("color")
    font_el = next((e for e in root.iter() if _local(e.tag) == "font" and e.get("color")), None)
    if font_el is not None:   # <font color=...> wrapping all the text overrides the box colour
        squash = lambda t: re.sub(r"\s+", "", t)
        if squash("".join(out)) == squash("".join(font_el.itertext())):
            color = font_el.get("color")
    return HtmlLabel(out, size, family, color, bold, lh, align, anchor_x,
                     block_top + lh / 2 + BASELINE_FROM_CENTRE * size, block_w)
