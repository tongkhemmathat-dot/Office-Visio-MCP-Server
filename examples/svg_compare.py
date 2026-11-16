"""Convert SVG files with svg_to_visio and compare each result with how a browser draws the SVG.

Usage: python examples/svg_compare.py <out_dir> <file.svg> [<file.svg> ...]

For every SVG this writes, into out_dir:
  <name>.vsdx          the converted Visio file
  <name>.visio.png     Visio's rendering of page 1
  <name>.edge.png      headless Edge's rendering of the SVG
  <name>.side.png      Edge (left) next to Visio (right)
and prints one line per file: shape counts, skipped features and a difference score
(mean absolute grey difference, 0 = identical, lower is better).
Needs Windows, Visio and Edge.
"""
import asyncio
import os
import subprocess
import sys

from PIL import Image, ImageChops, ImageStat

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import visio_mcp_server.visio_server as server  # noqa: E402
from visio_mcp_server import svg_import  # noqa: E402

MAX_SIDE = 2400


def edge_render(svg: str, png: str, width: int, height: int) -> None:
    subprocess.run([svg_import.find_edge(), "--headless", "--disable-gpu", "--hide-scrollbars",
                    f"--window-size={width},{height}", "--screenshot=" + png,
                    "file:///" + os.path.abspath(svg).replace("\\", "/")],
                   check=True, capture_output=True, timeout=120)


def score(a: Image.Image, b: Image.Image) -> float:
    b = b.resize(a.size)
    return ImageStat.Stat(ImageChops.difference(a.convert("L"), b.convert("L"))).mean[0]


def main() -> None:
    out_dir = os.path.abspath(sys.argv[1])
    os.makedirs(out_dir, exist_ok=True)
    try:
        for svg in sys.argv[2:]:
            name = os.path.splitext(os.path.basename(svg))[0]
            base = os.path.join(out_dir, name)
            try:
                doc = svg_import.parse_svg(svg)
                edge_render(svg, base + ".edge.png", min(MAX_SIDE, int(doc.width)), min(MAX_SIDE, int(doc.height)))
                result = asyncio.run(server.svg_to_visio(svg, base + ".vsdx", export_png=base + ".visio.png"))
                asyncio.run(server.close_document(base + ".vsdx"))
                a, b = Image.open(base + ".edge.png").convert("RGB"), Image.open(base + ".visio.png").convert("RGB")
                side = Image.new("RGB", (a.width + b.width + 10, max(a.height, b.height)), "#888888")
                side.paste(a, (0, 0))
                side.paste(b, (a.width + 10, 0))
                side.save(base + ".side.png")
                print(f"{name[:48]:48} score {score(a, b):6.2f}  {result.split(': ', 1)[-1][:160]}")
            except Exception as e:  # keep going: one bad file should not stop the run
                print(f"{name[:48]:48} FAILED {type(e).__name__}: {str(e)[:150]}")
    finally:
        server.close_visio_app()


if __name__ == "__main__":
    main()
