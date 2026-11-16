"""Read an SVG into a flat list of drawing primitives (no Visio, no COM).

The Visio side lives in visio_server.svg_to_visio. Keeping the parser separate makes it testable
and lets the same primitives feed other back ends.

Everything is resolved here, so the primitives are in page coordinates (SVG px, y down):
  * <g>, <use>/<symbol> and nested transforms (translate, scale, rotate, skew, matrix),
  * presentation attributes, the style attribute and simple <style> rules (tag, .class, #id),
    with inheritance, opacity and fill/stroke gradients (reduced to their average colour),
  * rect (incl. rx), circle, ellipse, line, polyline, polygon, path (curves and arcs are flattened
    into line segments), text/tspan and image (data URIs or files next to the SVG).
Not supported (reported in SvgDocument.skipped): filters, masks, clip-paths, patterns, markers other
than start/end arrows, nested <svg>, foreignObject. A path with several subpaths becomes one Path
(drawn even-odd, so holes are cut out; SVG's default nonzero rule differs only for overlapping same-direction subpaths).
"""
import base64
import math
import os
import re
import subprocess
import urllib.parse
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from visio_mcp_server import svg_html

SVG_NS_PREFIX = "{http://www.w3.org/2000/svg}"
XLINK_HREF = "{http://www.w3.org/1999/xlink}href"
DEFAULT_FONT_SIZE = 16.0
EDGE_PATHS = (
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
)
ICON_RENDER_PX = 320  # rasterised SVG icons are 4x their usual 80 px
CURVE_STEP_PX = 5.0   # target length of a flattened curve segment
MAX_USE_DEPTH = 20

Matrix = Tuple[float, float, float, float, float, float]
IDENTITY: Matrix = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)


class SvgError(Exception):
    """The SVG is missing, malformed or uses something this importer cannot handle."""


@dataclass
class Rect:
    x: float
    y: float
    width: float
    height: float
    fill: Optional[str]          # "#RRGGBB" or None
    stroke: Optional[str]
    stroke_width: float
    dash: Optional[str]          # None | "dash" | "dot"
    radius: float = 0.0
    fill_opacity: float = 1.0
    stroke_opacity: float = 1.0


@dataclass
class Ellipse:
    cx: float
    cy: float
    rx: float
    ry: float
    fill: Optional[str]
    stroke: Optional[str]
    stroke_width: float
    dash: Optional[str]
    fill_opacity: float = 1.0
    stroke_opacity: float = 1.0


@dataclass
class Polyline:
    points: List[Tuple[float, float]]
    stroke: Optional[str]
    stroke_width: float
    dash: Optional[str]
    arrow_end: bool = False
    arrow_start: bool = False
    closed: bool = False
    fill: Optional[str] = None
    fill_opacity: float = 1.0
    stroke_opacity: float = 1.0


@dataclass
class Path:
    """Several subpaths drawn as ONE shape, so that holes (glyph counters, ring icons) are cut out."""
    subpaths: List[List[Tuple[float, float]]]   # a closed subpath repeats its first point at the end
    stroke: Optional[str]
    stroke_width: float
    dash: Optional[str]
    fill: Optional[str] = None
    fill_opacity: float = 1.0
    stroke_opacity: float = 1.0


@dataclass
class Text:
    lines: List[str]
    x: float
    y: float                     # baseline of the first line
    font_size: float
    anchor: str                  # start | middle | end
    fill: str
    bold: bool
    line_height: float           # distance between baselines
    halo: bool                   # white outline behind the glyphs (paint-order stroke)
    rotation: float = 0.0        # degrees, clockwise (SVG convention), about (x, y)
    opacity: float = 1.0
    font_family: str = ""        # CSS font-family list as written in the SVG


@dataclass
class Image:
    x: float
    y: float
    width: float
    height: float
    href: str
    opacity: float = 1.0


@dataclass
class SvgDocument:
    width: float
    height: float
    font_family: str
    items: list = field(default_factory=list)
    skipped: List[str] = field(default_factory=list)
    base_dir: str = "."


# --- Small helpers -----------------------------------------------------------

def _local(tag) -> str:
    return tag.rsplit("}", 1)[-1] if isinstance(tag, str) else ""


NUMBER = r"[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?"
UNITS = {"px": 1.0, "pt": 96 / 72, "pc": 16.0, "mm": 96 / 25.4, "cm": 96 / 2.54, "in": 96.0, "": 1.0}


def parse_length(value, default=0.0, font_size=DEFAULT_FONT_SIZE) -> float:
    """A CSS length in px. Percentages are not resolvable here and fall back to the default."""
    if value is None:
        return default
    m = re.match(rf"\s*({NUMBER})\s*(px|pt|pc|mm|cm|in|em|ex|%)?\s*$", str(value))
    if not m:
        return default
    n, unit = float(m.group(1)), m.group(2) or ""
    if unit == "%":
        return default
    if unit == "em":
        return n * font_size
    if unit == "ex":
        return n * font_size / 2
    return n * UNITS[unit]


def _floats(text: str) -> List[float]:
    return [float(n) for n in re.findall(NUMBER, text or "")]


def parse_color(value: Optional[str]) -> Optional[str]:
    """Return '#RRGGBB' for hex, rgb() and a few named colours; None for none/transparent/url(...)."""
    if not value:
        return None
    v = value.strip()
    if v.lower() in ("none", "transparent") or v.startswith("url("):
        return None
    m = re.fullmatch(r"#([0-9a-fA-F]{3})", v)
    if m:
        return "#" + "".join(c * 2 for c in m.group(1)).upper()
    if re.fullmatch(r"#[0-9a-fA-F]{6}", v):
        return v.upper()
    m = re.fullmatch(r"rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)", v)
    if m:
        return "#%02X%02X%02X" % tuple(min(255, int(g)) for g in m.groups())
    named = {"black": "#000000", "white": "#FFFFFF", "red": "#FF0000", "gray": "#808080", "grey": "#808080",
             "blue": "#0000FF", "green": "#008000", "orange": "#FFA500", "yellow": "#FFFF00",
             "purple": "#800080", "navy": "#000080", "teal": "#008080"}
    return named.get(v.lower())


def _average(colors: List[str]) -> Optional[str]:
    if not colors:
        return None
    rgb = [[int(c[i:i + 2], 16) for i in (1, 3, 5)] for c in colors]
    return "#%02X%02X%02X" % tuple(round(sum(ch) / len(ch)) for ch in zip(*rgb))


# --- Transforms --------------------------------------------------------------

def mat_mul(m: Matrix, n: Matrix) -> Matrix:
    """m applied after n (point -> m(n(point)))."""
    a1, b1, c1, d1, e1, f1 = m
    a2, b2, c2, d2, e2, f2 = n
    return (a1 * a2 + c1 * b2, b1 * a2 + d1 * b2, a1 * c2 + c1 * d2, b1 * c2 + d1 * d2,
            a1 * e2 + c1 * f2 + e1, b1 * e2 + d1 * f2 + f1)


def mat_apply(m: Matrix, x: float, y: float) -> Tuple[float, float]:
    return m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]


def parse_transform(text: Optional[str]) -> Matrix:
    """Parse an SVG transform list into one matrix."""
    result = IDENTITY
    for name, args in re.findall(r"(\w+)\s*\(([^)]*)\)", text or ""):
        a = _floats(args)
        if name == "translate" and a:
            m = (1, 0, 0, 1, a[0], a[1] if len(a) > 1 else 0)
        elif name == "scale" and a:
            m = (a[0], 0, 0, a[1] if len(a) > 1 else a[0], 0, 0)
        elif name == "rotate" and a:
            r = math.radians(a[0])
            c, s = math.cos(r), math.sin(r)
            m = (c, s, -s, c, 0, 0)
            if len(a) >= 3:
                m = mat_mul(mat_mul((1, 0, 0, 1, a[1], a[2]), m), (1, 0, 0, 1, -a[1], -a[2]))
        elif name == "matrix" and len(a) == 6:
            m = tuple(a)
        elif name == "skewX" and a:
            m = (1, 0, math.tan(math.radians(a[0])), 1, 0, 0)
        elif name == "skewY" and a:
            m = (1, math.tan(math.radians(a[0])), 0, 1, 0, 0)
        else:
            continue
        result = mat_mul(result, m)
    return result


def _axis_aligned(m: Matrix) -> bool:
    return abs(m[1]) < 1e-9 and abs(m[2]) < 1e-9


def _scale_of(m: Matrix) -> float:
    return math.sqrt(abs(m[0] * m[3] - m[1] * m[2]))


# --- Paths -------------------------------------------------------------------

def _flatten_cubic(p0, p1, p2, p3) -> List[Tuple[float, float]]:
    length = (math.dist(p0, p1) + math.dist(p1, p2) + math.dist(p2, p3))
    n = max(4, min(64, math.ceil(length / CURVE_STEP_PX)))
    pts = []
    for i in range(1, n + 1):
        t = i / n
        u = 1 - t
        pts.append((u ** 3 * p0[0] + 3 * u * u * t * p1[0] + 3 * u * t * t * p2[0] + t ** 3 * p3[0],
                    u ** 3 * p0[1] + 3 * u * u * t * p1[1] + 3 * u * t * t * p2[1] + t ** 3 * p3[1]))
    return pts


def _flatten_arc(p0, rx, ry, rot, large, sweep, p1) -> List[Tuple[float, float]]:
    """SVG endpoint arc -> points (SVG implementation notes, F.6.5)."""
    if p0 == p1:
        return []
    rx, ry = abs(rx), abs(ry)
    if rx == 0 or ry == 0:
        return [p1]
    phi = math.radians(rot)
    cp, sp = math.cos(phi), math.sin(phi)
    dx, dy = (p0[0] - p1[0]) / 2, (p0[1] - p1[1]) / 2
    x1, y1 = cp * dx + sp * dy, -sp * dx + cp * dy
    lam = x1 * x1 / (rx * rx) + y1 * y1 / (ry * ry)
    if lam > 1:
        rx, ry = rx * math.sqrt(lam), ry * math.sqrt(lam)
    num = rx * rx * ry * ry - rx * rx * y1 * y1 - ry * ry * x1 * x1
    den = rx * rx * y1 * y1 + ry * ry * x1 * x1
    coef = math.sqrt(max(0.0, num / den)) * (-1 if large == sweep else 1)
    cx1, cy1 = coef * rx * y1 / ry, -coef * ry * x1 / rx
    cx = cp * cx1 - sp * cy1 + (p0[0] + p1[0]) / 2
    cy = sp * cx1 + cp * cy1 + (p0[1] + p1[1]) / 2

    def angle(ux, uy, vx, vy):
        a = math.atan2(ux * vy - uy * vx, ux * vx + uy * vy)
        return a

    t1 = angle(1, 0, (x1 - cx1) / rx, (y1 - cy1) / ry)
    dt = angle((x1 - cx1) / rx, (y1 - cy1) / ry, (-x1 - cx1) / rx, (-y1 - cy1) / ry)
    if not sweep and dt > 0:
        dt -= 2 * math.pi
    elif sweep and dt < 0:
        dt += 2 * math.pi
    n = max(6, min(96, math.ceil(abs(dt) * max(rx, ry) / CURVE_STEP_PX)))
    pts = []
    for i in range(1, n + 1):
        t = t1 + dt * i / n
        px, py = rx * math.cos(t), ry * math.sin(t)
        pts.append((cp * px - sp * py + cx, sp * px + cp * py + cy))
    pts[-1] = p1
    return pts


def parse_path(d: str) -> List[Tuple[List[Tuple[float, float]], bool]]:
    """Flatten an SVG path 'd' into subpaths: [(points, closed), ...]."""
    tokens = re.findall(rf"[MmLlHhVvCcSsQqTtAaZz]|{NUMBER}", d or "")
    pos = 0
    subpaths: List[Tuple[List[Tuple[float, float]], bool]] = []
    cur: List[Tuple[float, float]] = []
    x = y = sx = sy = 0.0
    last_ctrl = None   # (x, y, kind) of the previous curve's control point for S/T reflection
    cmd = None

    def num():
        nonlocal pos
        if pos >= len(tokens) or re.match(r"[A-Za-z]", tokens[pos]):
            raise SvgError("Malformed path data.")
        pos += 1
        return float(tokens[pos - 1])

    def flag():
        # Arc flags may be written without separators ("a1 1 0 01 1 1"): the tokeniser then glues
        # the flag to the next number, so peel one digit off.
        nonlocal pos
        tok = tokens[pos]
        if tok in ("0", "1"):
            pos += 1
            return tok == "1"
        if tok[0] in "01" and len(tok) > 1:
            tokens[pos] = tok[1:]
            return tok[0] == "1"
        raise SvgError("Malformed arc flag in path data.")

    while pos < len(tokens):
        if re.match(r"[A-Za-z]", tokens[pos]):
            cmd = tokens[pos]
            pos += 1
            if cmd in "Zz":
                if cur:
                    subpaths.append((cur, True))
                    cur = []
                x, y = sx, sy
                last_ctrl = None
                continue
        elif cmd is None:
            raise SvgError("Path data must start with a command.")
        rel = cmd.islower()
        c = cmd.upper()
        ox, oy = (x, y) if rel else (0.0, 0.0)
        if c == "M":
            if cur:
                subpaths.append((cur, False))
            x, y = ox + num(), oy + num()
            sx, sy = x, y
            cur = [(x, y)]
            cmd = "l" if rel else "L"
            last_ctrl = None
            continue
        if not cur:
            cur = [(x, y)]
        if c == "L":
            x, y = ox + num(), oy + num()
            cur.append((x, y))
            last_ctrl = None
        elif c == "H":
            x = ox + num()
            cur.append((x, y))
            last_ctrl = None
        elif c == "V":
            y = oy + num()
            cur.append((x, y))
            last_ctrl = None
        elif c == "C":
            p1 = (ox + num(), oy + num())
            p2 = (ox + num(), oy + num())
            p3 = (ox + num(), oy + num())
            cur.extend(_flatten_cubic((x, y), p1, p2, p3))
            x, y = p3
            last_ctrl = (p2, "C")
        elif c == "S":
            p1 = (2 * x - last_ctrl[0][0], 2 * y - last_ctrl[0][1]) if last_ctrl and last_ctrl[1] == "C" else (x, y)
            p2 = (ox + num(), oy + num())
            p3 = (ox + num(), oy + num())
            cur.extend(_flatten_cubic((x, y), p1, p2, p3))
            x, y = p3
            last_ctrl = (p2, "C")
        elif c in "QT":
            if c == "Q":
                q = (ox + num(), oy + num())
            else:
                q = (2 * x - last_ctrl[0][0], 2 * y - last_ctrl[0][1]) if last_ctrl and last_ctrl[1] == "Q" else (x, y)
            p3 = (ox + num(), oy + num())
            p1 = (x + 2 / 3 * (q[0] - x), y + 2 / 3 * (q[1] - y))
            p2 = (p3[0] + 2 / 3 * (q[0] - p3[0]), p3[1] + 2 / 3 * (q[1] - p3[1]))
            cur.extend(_flatten_cubic((x, y), p1, p2, p3))
            x, y = p3
            last_ctrl = (q, "Q")
        elif c == "A":
            rx, ry, rot = num(), num(), num()
            large, sweep = flag(), flag()
            p1 = (ox + num(), oy + num())
            cur.extend(_flatten_arc((x, y), rx, ry, rot, large, sweep, p1))
            x, y = p1
            last_ctrl = None
        else:
            raise SvgError(f"Unsupported path command '{cmd}'.")
    if cur:
        subpaths.append((cur, False))
    return [(pts, closed) for pts, closed in subpaths if len(pts) >= 2]


# --- Styles ------------------------------------------------------------------

INHERITED = {"fill", "stroke", "stroke-width", "stroke-dasharray", "stroke-opacity", "fill-opacity",
             "font-size", "font-weight", "font-family", "text-anchor", "marker-start", "marker-end",
             "color", "visibility"}
PRESENTATION = INHERITED | {"opacity", "display", "paint-order", "clip-path", "mask", "filter", "fill-rule"}
DEFAULT_STYLE = {"fill": "#000000", "stroke": "none", "stroke-width": "1", "font-size": "16",
                 "fill-opacity": "1", "stroke-opacity": "1"}
BOLD_WEIGHTS = {"bold", "bolder", "600", "700", "800", "900"}


def parse_css(text: str) -> List[Tuple[int, str, Dict[str, str]]]:
    """Very small CSS reader: comma lists of tag, .class, #id and tag.class selectors."""
    rules = []
    text = re.sub(r"/\*.*?\*/", "", text, flags=re.S)
    for selectors, body in re.findall(r"([^{}]+)\{([^}]*)\}", text):
        props = _declarations(body)
        for sel in (s.strip() for s in selectors.split(",")):
            if re.fullmatch(r"\*|[\w-]+|[\w-]*\.[\w-]+|#[\w-]+", sel):
                spec = (100 if sel.startswith("#") else 0) + 10 * sel.count(".") + (1 if re.match(r"\w", sel) else 0)
                rules.append((spec, sel, props))
    return rules


def _declarations(body: str) -> Dict[str, str]:
    return {k.strip(): v.strip() for k, _, v in (p.partition(":") for p in (body or "").split(";")) if v.strip()}


def _matches(selector: str, el) -> bool:
    if selector == "*":
        return True
    if selector.startswith("#"):
        return el.get("id") == selector[1:]
    tag, _, cls = selector.partition(".")
    if tag and tag != _local(el.tag):
        return False
    return not cls or cls in (el.get("class") or "").split()


@dataclass
class Context:
    matrix: Matrix
    style: Dict[str, str]
    opacity: float = 1.0
    depth: int = 0


class _Parser:
    def __init__(self, root, doc: SvgDocument):
        self.root = root
        self.doc = doc
        self.ids = {el.get("id"): el for el in root.iter() if el.get("id")}
        self.rules = []
        for el in root.iter():
            if _local(el.tag) == "style":
                self.rules += parse_css("".join(el.itertext()))
        self.rules.sort(key=lambda r: r[0])
        self.gradients: Dict[str, Optional[str]] = {}

    # -- style resolution
    def computed(self, el, parent: Dict[str, str]) -> Dict[str, str]:
        style = {k: v for k, v in parent.items() if k in INHERITED}
        for k in PRESENTATION:
            if el.get(k) is not None:
                style[k] = el.get(k).strip()
        for _, sel, props in self.rules:
            if _matches(sel, el):
                style.update({k: v for k, v in props.items() if k in PRESENTATION})
        style.update({k: v for k, v in _declarations(el.get("style")).items() if k in PRESENTATION})
        return style

    def paint(self, style: Dict[str, str], key: str) -> Optional[str]:
        v = style.get(key, DEFAULT_STYLE.get(key))
        if v is None:
            return None
        if v.strip().lower() == "currentcolor":
            v = style.get("color", "#000000")
        m = re.match(r"url\(\s*#([^)\s]+)\s*\)", v)
        if m:
            return self.gradient_color(m.group(1))
        return parse_color(v)

    def gradient_color(self, gid: str) -> Optional[str]:
        if gid not in self.gradients:
            el = self.ids.get(gid)
            seen = set()
            stops = []
            while el is not None and id(el) not in seen:
                seen.add(id(el))
                stops = [s for s in el if _local(s.tag) == "stop"]
                if stops:
                    break
                href = el.get(XLINK_HREF) or el.get("href")
                el = self.ids.get(href[1:]) if href and href.startswith("#") else None
            colors = []
            for s in stops:
                decl = _declarations(s.get("style"))
                c = parse_color(decl.get("stop-color") or s.get("stop-color") or "#000000")
                if c:
                    colors.append(c)
            self.gradients[gid] = _average(colors)
            if not stops:
                self.doc.skipped.append(f"gradient #{gid} not found")
        return self.gradients[gid]

    @staticmethod
    def dash(style: Dict[str, str], width: float) -> Optional[str]:
        d = style.get("stroke-dasharray", "none")
        vals = _floats(d) if d.strip().lower() != "none" else []
        if not vals or not any(vals):
            return None
        return "dot" if vals[0] <= max(width, 1.0) * 1.5 else "dash"

    def stroke_and_fill(self, style, m: Matrix, ctx_opacity: float, filled=True):
        stroke = self.paint(style, "stroke")
        width = parse_length(style.get("stroke-width"), 1.0) * _scale_of(m)
        if stroke is not None and width <= 0:
            stroke = None
        fill = self.paint(style, "fill") if filled else None
        return {
            "stroke": stroke, "stroke_width": width, "dash": self.dash(style, width) if stroke else None,
            "fill": fill,
            "fill_opacity": max(0.0, min(1.0, parse_length(style.get("fill-opacity"), 1.0) * ctx_opacity)),
            "stroke_opacity": max(0.0, min(1.0, parse_length(style.get("stroke-opacity"), 1.0) * ctx_opacity)),
        }

    # -- walking
    def walk(self, el, ctx: Context) -> None:
        tag = _local(el.tag)
        if el.tag.startswith("{") and not el.tag.startswith(SVG_NS_PREFIX):
            return  # editor metadata (sodipodi, inkscape, rdf, ...), not drawing content
        if tag in ("defs", "title", "desc", "metadata", "style", "symbol", "clipPath", "mask", "marker",
                   "linearGradient", "radialGradient", "pattern", "stop", "script", "filter", ""):
            return
        style = self.computed(el, ctx.style)
        if style.get("display", "").lower() == "none" or style.get("visibility", "").lower() in ("hidden", "collapse"):
            return
        for k in ("clip-path", "mask", "filter"):
            if style.get(k, "none") != "none":
                self.doc.skipped.append(f"{k} on <{tag}> ignored")
        m = mat_mul(ctx.matrix, parse_transform(el.get("transform")))
        opacity = ctx.opacity * max(0.0, min(1.0, parse_length(style.get("opacity"), 1.0)))
        sub = Context(m, style, opacity, ctx.depth)

        if tag == "g" or tag == "a":
            for child in el:
                self.walk(child, sub)
        elif tag == "switch":
            # Only the first child whose conditions pass is drawn. draw.io writes
            #   <switch><foreignObject>html label</foreignObject><text>truncated copy</text></switch>
            # and a closing <switch><g requiredFeatures=.../><a><text>Text is not SVG...</text></a></switch>
            # whose warning browsers never show. The plain <text> is used only if the label is unreadable.
            drawn = unreadable = False
            for child in el:
                if not self.switch_accepts(child):
                    continue
                if _local(child.tag) == "foreignObject":
                    if self.foreign(child, sub):
                        drawn = True
                        break
                    unreadable = True
                    continue
                self.walk(child, sub)
                drawn = True
                break
            if unreadable and not drawn:
                self.doc.skipped.append("<foreignObject>")
        elif tag == "foreignObject":
            if not self.foreign(el, sub):
                self.doc.skipped.append("<foreignObject>")
        elif tag == "use":
            self.use(el, sub)
        elif tag == "rect":
            self.rect(el, sub)
        elif tag in ("circle", "ellipse"):
            self.ellipse(el, sub, tag)
        elif tag in ("line", "polyline", "polygon", "path"):
            self.lines(el, sub, tag)
        elif tag == "text":
            self.text(el, sub)
        elif tag == "image":
            self.image(el, sub)
        else:
            self.doc.skipped.append(f"<{tag}>")

    @staticmethod
    def switch_accepts(child) -> bool:
        """Conditional processing for <switch> children (feature strings count as supported, as in browsers)."""
        if not isinstance(child.tag, str) or (child.tag.startswith("{") and not child.tag.startswith(SVG_NS_PREFIX)):
            return False
        ext = child.get("requiredExtensions")
        if ext is not None and "xhtml" not in ext.lower():
            return False
        lang = child.get("systemLanguage")
        return lang is None or any(l.strip().lower().startswith("en") for l in lang.split(","))

    def foreign(self, fo, ctx: Context) -> bool:
        """Turn an HTML label in a <foreignObject> into a Text. False if it holds no readable text."""
        label = svg_html.parse_label(fo)
        if label is None:
            return False
        m = ctx.matrix
        x, y = mat_apply(m, label.anchor_x, label.baseline_y)
        scale = _scale_of(m)
        self.doc.items.append(Text(
            label.lines, x, y, label.font_size * scale,
            {"left": "start", "right": "end"}.get(label.align, "middle"),
            parse_color(label.color) or "#000000", label.bold, label.line_height * scale, False,
            opacity=ctx.opacity, font_family=label.font_family))
        return True

    def use(self, el, ctx: Context) -> None:
        href = el.get(XLINK_HREF) or el.get("href") or ""
        target = self.ids.get(href[1:]) if href.startswith("#") else None
        if target is None:
            self.doc.skipped.append(f"<use> of {href or 'nothing'}")
            return
        if ctx.depth >= MAX_USE_DEPTH:
            self.doc.skipped.append("<use> nested too deeply")
            return
        offset = (1, 0, 0, 1, parse_length(el.get("x")), parse_length(el.get("y")))
        sub = Context(mat_mul(ctx.matrix, offset), ctx.style, ctx.opacity, ctx.depth + 1)
        if _local(target.tag) == "symbol":
            if target.get("viewBox"):
                self.doc.skipped.append("<symbol> viewBox ignored")
            for child in target:
                self.walk(child, sub)
        else:
            # walk the target as if it were a child of the <use> (it re-reads its own transform)
            self.walk(target, sub)

    def rect(self, el, ctx: Context) -> None:
        x, y = parse_length(el.get("x")), parse_length(el.get("y"))
        w, h = parse_length(el.get("width")), parse_length(el.get("height"))
        if w <= 0 or h <= 0:
            return
        if el.get("x") is None and el.get("y") is None and ctx.matrix == IDENTITY and self.is_page_background(w, h, ctx):
            return
        st = self.stroke_and_fill(ctx.style, ctx.matrix, ctx.opacity)
        m = ctx.matrix
        if _axis_aligned(m):
            x1, y1 = mat_apply(m, x, y)
            x2, y2 = mat_apply(m, x + w, y + h)
            rx = parse_length(el.get("rx"), parse_length(el.get("ry"))) * abs(m[0])
            self.doc.items.append(Rect(min(x1, x2), min(y1, y2), abs(x2 - x1), abs(y2 - y1), radius=rx, **st))
        else:
            pts = [mat_apply(m, px, py) for px, py in ((x, y), (x + w, y), (x + w, y + h), (x, y + h))]
            self.doc.items.append(Polyline(pts + [pts[0]], closed=True, **{
                k: v for k, v in st.items()}))

    def is_page_background(self, w, h, ctx) -> bool:
        return w >= self.doc.width and h >= self.doc.height and (
            self.paint(ctx.style, "stroke") is None) and self.paint(ctx.style, "fill") == "#FFFFFF"

    def ellipse(self, el, ctx: Context, tag: str) -> None:
        cx, cy = parse_length(el.get("cx")), parse_length(el.get("cy"))
        if tag == "circle":
            rx = ry = parse_length(el.get("r"))
        else:
            rx, ry = parse_length(el.get("rx")), parse_length(el.get("ry"))
        if rx <= 0 or ry <= 0:
            return
        st = self.stroke_and_fill(ctx.style, ctx.matrix, ctx.opacity)
        m = ctx.matrix
        if _axis_aligned(m):
            x, y = mat_apply(m, cx, cy)
            self.doc.items.append(Ellipse(x, y, rx * abs(m[0]), ry * abs(m[3]), **st))
        else:
            n = max(16, min(96, math.ceil(2 * math.pi * max(rx, ry) / CURVE_STEP_PX)))
            pts = [mat_apply(m, cx + rx * math.cos(2 * math.pi * i / n), cy + ry * math.sin(2 * math.pi * i / n))
                   for i in range(n)]
            self.doc.items.append(Polyline(pts + [pts[0]], closed=True, **st))

    def lines(self, el, ctx: Context, tag: str) -> None:
        if tag == "line":
            subpaths = [([(parse_length(el.get("x1")), parse_length(el.get("y1"))),
                          (parse_length(el.get("x2")), parse_length(el.get("y2")))], False)]
        elif tag == "path":
            try:
                subpaths = parse_path(el.get("d", ""))
            except SvgError as e:
                self.doc.skipped.append(f"<path> {e}")
                return
        else:
            raw = [float(n) for n in re.findall(NUMBER, el.get("points", ""))]
            pts = list(zip(raw[0::2], raw[1::2]))
            if len(pts) < 2:
                self.doc.skipped.append(f"{tag} with fewer than 2 points")
                return
            subpaths = [(pts, tag == "polygon")]
        filled = tag != "line"
        st = self.stroke_and_fill(ctx.style, ctx.matrix, ctx.opacity, filled=filled)
        if st["fill"] is None and st["stroke"] is None:
            return
        placed = []
        for pts, closed in subpaths:
            pts = [mat_apply(ctx.matrix, px, py) for px, py in pts]
            if st["fill"] is not None:
                closed = True
            if closed and pts[0] != pts[-1]:
                pts = pts + [pts[0]]
            placed.append((pts, closed))
        if len(placed) > 1:
            self.doc.items.append(Path([pts for pts, _ in placed], **st))
            return
        for pts, closed in placed:
            self.doc.items.append(Polyline(
                pts, closed=closed,
                arrow_end=ctx.style.get("marker-end", "none") != "none" and not closed,
                arrow_start=ctx.style.get("marker-start", "none") != "none" and not closed, **st))

    def text(self, el, ctx: Context) -> None:
        style, m = ctx.style, ctx.matrix
        size = parse_length(style.get("font-size"), DEFAULT_FONT_SIZE)
        tspans = [t for t in el if _local(t.tag) == "tspan"]
        lines = ["".join(t.itertext()) for t in tspans] or ["".join(el.itertext())]
        if not any(l.strip() for l in lines):
            return
        x = _floats(el.get("x") or (tspans[0].get("x") if tspans else None) or "0")[:1] or [0.0]
        y = _floats(el.get("y") or (tspans[0].get("y") if tspans else None) or "0")[:1] or [0.0]
        if len(tspans) > 1:
            if tspans[1].get("y") is not None and tspans[0].get("y") is not None:
                dy = parse_length(tspans[1].get("y")) - parse_length(tspans[0].get("y"))
            else:
                dy = parse_length(tspans[1].get("dy"), size * 1.2, size)
        else:
            dy = size * 1.2
        px, py = mat_apply(m, x[0], y[0])
        scale = _scale_of(m)
        fill = self.paint(style, "fill") or "#000000"
        stroke = self.paint(style, "stroke")
        weight = style.get("font-weight", "normal").lower()
        self.doc.items.append(Text(
            lines, px, py, size * scale, style.get("text-anchor", "start"), fill, weight in BOLD_WEIGHTS,
            dy * scale, stroke == "#FFFFFF" and parse_length(style.get("stroke-width"), 1.0) > 0,
            rotation=math.degrees(math.atan2(m[1], m[0])) if not _axis_aligned(m) or m[0] < 0 else 0.0,
            opacity=max(0.0, min(1.0, parse_length(style.get("fill-opacity"), 1.0) * ctx.opacity)),
            font_family=style.get("font-family", "")))
        if not self.doc.font_family and style.get("font-family"):
            self.doc.font_family = style["font-family"]

    def image(self, el, ctx: Context) -> None:
        href = el.get(XLINK_HREF) or el.get("href")
        if not href:
            self.doc.skipped.append("image without href")
            return
        m = ctx.matrix
        x, y = parse_length(el.get("x")), parse_length(el.get("y"))
        w, h = parse_length(el.get("width")), parse_length(el.get("height"))
        if not _axis_aligned(m):
            self.doc.skipped.append("rotated image drawn unrotated")
        x1, y1 = mat_apply(m, x, y)
        x2, y2 = mat_apply(m, x + w, y + h)
        self.doc.items.append(Image(min(x1, x2), min(y1, y2), abs(x2 - x1), abs(y2 - y1), href, ctx.opacity))


def _viewport(root) -> Tuple[float, float, Matrix]:
    vb = _floats(root.get("viewBox") or "")
    width, height = parse_length(root.get("width"), 0), parse_length(root.get("height"), 0)
    if len(vb) == 4 and vb[2] > 0 and vb[3] > 0:
        width = width or vb[2]
        height = height or vb[3]
        sx, sy = width / vb[2], height / vb[3]
        s = min(sx, sy)  # default preserveAspectRatio="xMidYMid meet"
        return width, height, (s, 0, 0, s, (width - vb[2] * s) / 2 - vb[0] * s, (height - vb[3] * s) / 2 - vb[1] * s)
    if width <= 0 or height <= 0:
        raise SvgError("The SVG has no width/height or viewBox.")
    return width, height, IDENTITY


def parse_svg(path: str) -> SvgDocument:
    """Parse an SVG file into an SvgDocument."""
    if not os.path.isfile(path):
        raise SvgError(f"SVG file not found: {path}")
    try:
        root = ET.parse(path).getroot()
    except ET.ParseError as e:
        raise SvgError(f"Not a valid SVG: {e}")
    if _local(root.tag) != "svg":
        raise SvgError("The file's root element is not <svg>.")

    width, height, matrix = _viewport(root)
    doc = SvgDocument(width, height, root.get("font-family", ""), base_dir=os.path.dirname(os.path.abspath(path)))
    parser = _Parser(root, doc)
    ctx = Context(matrix, parser.computed(root, {}))
    for child in root:
        parser.walk(child, ctx)
    return doc


# --- Icons -------------------------------------------------------------------

def find_edge() -> Optional[str]:
    return next((p for p in EDGE_PATHS if os.path.exists(p)), None)


def image_to_png(href: str, base_dir: str, work_dir: str, name: str) -> str:
    """Return an image file Visio can import for an <image> href, rasterising SVG icons with headless Edge."""
    m = re.match(r"data:image/([\w+.-]+);base64,(.*)", href, re.S)
    if m:
        kind, data = m.group(1).lower(), base64.b64decode(m.group(2))
    else:
        file_path = urllib.parse.unquote(href)
        if not os.path.isabs(file_path):
            file_path = os.path.join(base_dir, file_path)
        if not os.path.isfile(file_path):
            raise SvgError(f"Image file not found: {href[:80]}")
        kind = os.path.splitext(file_path)[1].lstrip(".").lower()
        with open(file_path, "rb") as f:
            data = f.read()
        if kind == "svg":
            kind = "svg+xml"

    if kind in ("png", "jpeg", "jpg", "gif", "bmp"):
        ext = "jpg" if kind == "jpeg" else kind
        out = os.path.join(work_dir, f"{name}.{ext}")
        with open(out, "wb") as f:
            f.write(data)
        return out
    if kind != "svg+xml":
        raise SvgError(f"Unsupported image type '{kind}'.")

    png_path = os.path.join(work_dir, f"{name}.png")
    edge = find_edge()
    if not edge:
        raise SvgError("Embedded SVG icons need Microsoft Edge (headless) to be rasterised, but it was not found.")
    svg_path = os.path.join(work_dir, f"{name}.svg")
    html_path = os.path.join(work_dir, f"{name}.html")
    with open(svg_path, "wb") as f:
        f.write(data)
    with open(html_path, "w", encoding="utf8") as f:
        f.write(f'<html><body style="margin:0;background:transparent"><img src="{name}.svg" '
                f'width="{ICON_RENDER_PX}" height="{ICON_RENDER_PX}" style="display:block"></body></html>')
    subprocess.run([edge, "--headless", "--disable-gpu", "--hide-scrollbars",
                    "--default-background-color=00000000",
                    f"--window-size={ICON_RENDER_PX},{ICON_RENDER_PX}",
                    "--screenshot=" + png_path, "file:///" + html_path.replace("\\", "/")],
                   check=True, capture_output=True, timeout=60)
    if not os.path.exists(png_path):
        raise SvgError(f"Edge did not produce a PNG for icon {name}.")
    return png_path
