"""Build a Visio diagram from a JSON spec without going through Claude.

    python examples/build_from_spec.py examples/hci_network.json out/hci.vsdx --png --pdf

Run it with the same Python that runs the server (the project's .venv). The spec format is
described in docs/TEAM_GUIDE.md and in the build_diagram tool.
"""

import argparse
import asyncio
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import visio_mcp_server.visio_server as server  # noqa: E402


async def build(spec_path: str, output: str, png: bool, pdf: bool, replace: bool) -> int:
    with open(spec_path, encoding="utf-8") as f:
        spec = json.load(f)
    output = os.path.abspath(output)
    os.makedirs(os.path.dirname(output), exist_ok=True)

    results = [await server.build_diagram(output, spec, replace_existing=replace)]
    base = os.path.splitext(output)[0]
    if png:
        results.append(await server.export_diagram(output, base + ".png"))
    if pdf:
        results.append(await server.export_diagram(output, base + ".pdf"))
    await server.close_document(output)

    for r in results:
        print(r)
    return 1 if any(r.startswith("Error") for r in results) else 0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("spec", help="JSON diagram spec")
    parser.add_argument("output", help="Visio file to create or update (.vsdx)")
    parser.add_argument("--png", action="store_true", help="also export a PNG next to the .vsdx")
    parser.add_argument("--pdf", action="store_true", help="also export a PDF next to the .vsdx")
    parser.add_argument("--replace", action="store_true", help="clear the page before drawing")
    args = parser.parse_args()
    try:
        return asyncio.run(build(args.spec, args.output, args.png, args.pdf, args.replace))
    finally:
        server.close_visio_app()


if __name__ == "__main__":
    sys.exit(main())
