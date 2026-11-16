"""Read an SVG into a flat list of drawing primitives (no Visio, no COM).

The Visio side lives in visio_server.svg_to_visio. Keeping the parser separate makes it testable
and lets the same primitives feed other back ends.

Supported today: top-level <rect>, <polyline>, <text> (with <tspan> lines) and <image> (PNG or
SVG data URIs, or files next to the SVG). Anything else is reported in SvgDocument.skipped.
Units are SVG user units (px); y grows downwards.
"""
import base64
import os
import re
import subprocess
import urllib.parse
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from typing import List, Optional, Tuple

SVG_NS = "http://www.w3.org/2000/svg"
XLINK_HREF = "{http://www.w3.org/1999/xlink}href"
DEFAULT_FONT_SIZE = 16.0
EDGE_PATHS = (
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
)
ICON_RENDER_PX = 320  # rasterised SVG icons are 4x their usual 80 px


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
    dashed: bool


@dataclass
class Polyline:
    points: List[Tuple[float, float]]
    stroke: Optional[str]
    stroke_width: float
    dashed: bool
    arrow_end: bool


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


@dataclass
class Image:
    x: float
    y: float
    width: float
    height: float
    href: str


@dataclass
class SvgDocument:
    width: float
    height: float
    font_family: str
    items: list = field(default_factory=list)
    skipped: List[str] = field(default_factory=list)
    base_dir: str = "."


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _num(value, default=0.0) -> float:
    if value is None:
        return default
    m = re.match(r"\s*(-?\d+(?:\.\d+)?(?:[eE][-+]?\d+)?)", value)
    return float(m.group(1)) if m else default


def parse_color(value: Optional[str]) -> Optional[str]:
    """Return '#RRGGBB' for hex and rgb() colours, None for none/transparent/unsupported (url(...))."""
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
    named = {"black": "#000000", "white": "#FFFFFF", "red": "#FF0000", "gray": "#808080", "grey": "#808080"}
    return named.get(v.lower())


def _viewport(root) -> Tuple[float, float]:
    width, height = _num(root.get("width"), 0), _num(root.get("height"), 0)
    vb = root.get("viewBox")
    if vb and (not width or not height):
        parts = [float(p) for p in re.split(r"[ ,]+", vb.strip())]
        if len(parts) == 4:
            width, height = parts[2], parts[3]
    if width <= 0 or height <= 0:
        raise SvgError("The SVG has no width/height or viewBox.")
    return width, height


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

    width, height = _viewport(root)
    doc = SvgDocument(width, height, root.get("font-family", ""), base_dir=os.path.dirname(os.path.abspath(path)))

    for el in root:
        tag = _local(el.tag)
        if tag in ("defs", "title", "desc", "metadata", "style"):
            continue
        if tag == "rect":
            if el.get("x") is None and el.get("y") is None and _num(el.get("width")) >= width \
                    and _num(el.get("height")) >= height:
                continue  # full-page background: the Visio page is already white
            doc.items.append(Rect(_num(el.get("x")), _num(el.get("y")), _num(el.get("width")),
                                  _num(el.get("height")), parse_color(el.get("fill", "#000000")),
                                  parse_color(el.get("stroke")), _num(el.get("stroke-width"), 1.0),
                                  bool(el.get("stroke-dasharray"))))
        elif tag == "polyline":
            pts = [tuple(float(n) for n in re.split(r"[ ,]+", p.strip()) if n)
                   for p in re.findall(r"-?[\d.]+\s*,\s*-?[\d.]+", el.get("points", ""))]
            pts = [(p[0], p[1]) for p in pts if len(p) == 2]
            if len(pts) < 2:
                doc.skipped.append("polyline with fewer than 2 points")
                continue
            doc.items.append(Polyline(pts, parse_color(el.get("stroke")), _num(el.get("stroke-width"), 1.0),
                                      bool(el.get("stroke-dasharray")), bool(el.get("marker-end"))))
        elif tag == "text":
            tspans = [t for t in el if _local(t.tag) == "tspan"]
            lines = ["".join(t.itertext()) for t in tspans] or ["".join(el.itertext())]
            size = _num(el.get("font-size"), DEFAULT_FONT_SIZE)
            dy = _num(tspans[1].get("dy"), size * 1.2) if len(tspans) > 1 else size * 1.2
            doc.items.append(Text(lines, _num(el.get("x")), _num(el.get("y")), size,
                                  el.get("text-anchor", "start"), parse_color(el.get("fill")) or "#000000",
                                  el.get("font-weight") in ("bold", "700", "800", "900"), dy,
                                  parse_color(el.get("stroke")) == "#FFFFFF"))
        elif tag == "image":
            href = el.get(XLINK_HREF) or el.get("href")
            if not href:
                doc.skipped.append("image without href")
                continue
            doc.items.append(Image(_num(el.get("x")), _num(el.get("y")), _num(el.get("width")),
                                   _num(el.get("height")), href))
        else:
            doc.skipped.append(f"<{tag}>")
    return doc


# --- Icons -------------------------------------------------------------------

def find_edge() -> Optional[str]:
    return next((p for p in EDGE_PATHS if os.path.exists(p)), None)


def image_to_png(href: str, base_dir: str, work_dir: str, name: str) -> str:
    """Return a PNG file for an <image> href, rasterising SVG icons with headless Edge."""
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

    png_path = os.path.join(work_dir, f"{name}.png")
    if kind in ("png", "jpeg", "jpg", "gif", "bmp"):
        ext = "jpg" if kind == "jpeg" else kind
        png_path = os.path.join(work_dir, f"{name}.{ext}")
        with open(png_path, "wb") as f:
            f.write(data)
        return png_path
    if kind != "svg+xml":
        raise SvgError(f"Unsupported image type '{kind}'.")

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
