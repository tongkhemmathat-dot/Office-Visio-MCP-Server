# Visio MCP Server

**English** | [ไทย](README.th.md)

An MCP server that lets an AI assistant (or any MCP client) create and edit Microsoft Visio diagrams
through Visio's COM automation interface: add shapes, connect them, use stencils, build a whole
network diagram from one JSON spec, mount devices in racks, export to PNG/PDF/SVG, and turn an SVG
file into an editable `.vsdx`.

![](https://badge.mcpx.dev?type=server "MCP Server")

![demo](./public/demo.gif)

## Contents

1. [Requirements](#requirements)
2. [Installation](#installation)
3. [Connecting a client](#connecting-a-client)
4. [How it works (read this once)](#how-it-works-read-this-once)
5. [Quick start](#quick-start)
6. [Tool reference](#tool-reference)
7. [Recipes](#recipes)
   - [A. Flowchart shape by shape](#a-flowchart-shape-by-shape)
   - [B. A network diagram from one spec](#b-a-network-diagram-from-one-spec)
   - [C. A rack diagram](#c-a-rack-diagram)
   - [D. SVG to Visio](#d-svg-to-visio)
   - [E. Run without an AI client](#e-run-without-an-ai-client)
8. [SVG import in depth](#svg-import-in-depth)
9. [Troubleshooting](#troubleshooting)
10. [Development](#development)
11. [License](#license)

More documents: [team guide](docs/TEAM_GUIDE.md) (Thai: prompts, stencils, racks),
[SVG import status](docs/SVG_IMPORT_TODO.md), [development notes](docs/DEVELOPMENT_NOTES.md),
[changelog](CHANGELOG.md), and ready-made specs in [`examples/`](examples).

## Requirements

- Windows (the server drives Visio through COM)
- Microsoft Visio desktop (Standard or Professional). Visio 2013 or later has been used
- Python 3.12 or later
- Python packages (installed automatically with `pip install -e .`):
  `mcp[cli]<2`, `pywin32`, `Pillow` (text measuring for SVG import), `resvg-py` (rasterises SVG icons)

## Installation

```powershell
git clone https://github.com/tongkhemmathat-dot/Office-Visio-MCP-Server.git
cd Office-Visio-MCP-Server
python -m venv .venv
.venv\Scripts\activate
pip install -e .
```

Check that it starts (it waits for an MCP client on stdin, so it will look idle; press Ctrl+C to stop):

```powershell
python -m visio_mcp_server.visio_server
# or, after pip install -e .
visio_mcp_server
```

If Visio is not installed the server prints a message and exits.

## Connecting a client

### Claude Code

```powershell
claude mcp add visio -- D:\path\to\repo\.venv\Scripts\python.exe -m visio_mcp_server.visio_server
```

### Claude Desktop or any client that uses a JSON config

```json
{
  "mcpServers": {
    "visio": {
      "command": "D:\\path\\to\\repo\\.venv\\Scripts\\python.exe",
      "args": ["-m", "visio_mcp_server.visio_server"]
    }
  }
}
```

### Without a local clone (uvx)

```json
{
  "mcpServers": {
    "visio": {
      "command": "uvx",
      "args": ["--from", "office-visio-mcp-server", "visio_mcp_server"]
    }
  }
}
```

Use the Python that has the packages installed (the project's `.venv`). Restart the client after
editing its config, then ask it something like *"list the Visio stencils on this machine"* to confirm.

## How it works (read this once)

- **One Visio, shared.** The server attaches to a Visio that is already running
  (`GetActiveObject`). If none is running it starts its own and quits it when the server exits.
  It never quits a Visio it did not start. Diagrams you build appear in your own Visio window,
  so close unsaved work you care about before running experiments.
- **Files are addressed by path.** Every tool takes `file_path`. A bare file name (`demo.vsdx`)
  resolves to your **Documents** folder. Most tools save the file after each change.
- **Units are inches** on the page. Origin is the bottom-left corner; y grows upward.
  A4 landscape is 11.69 x 8.27 in (the default page).
- **Shape IDs connect the steps.** `add_shape` and `add_stencil_shape` return a shape ID.
  Use it in `connect_shapes`, `add_text`, `delete_shape`. `list_shapes` shows all IDs again.
- **Colours** accept `#RRGGBB` or a name (`red`, `blue`, `teal`, ...). Line patterns are
  `solid`, `dash`, `dot`, `dashdot`. `line_weight` and `font_size` are in points.
- **Tests that run Visio** create documents in your Visio. Pure-parsing tests do not.

## Quick start

Ask your assistant, in plain words:

> Create `C:\Diagrams\hello.vsdx`, add a rectangle "Start" at (2, 4) and a circle "End" at (6, 4),
> connect them with an arrow, then export a PNG.

Behind the scenes the assistant calls `create_visio_file`, `add_shape` twice,
`connect_shapes` and `export_diagram`. The sections below show each tool's arguments.

## Tool reference

All arguments are JSON. Optional ones can be left out.

| Group | Tools |
|---|---|
| Files | `create_visio_file`, `open_visio_file`, `close_document`, `set_page_setup`, `export_diagram` |
| Shapes | `add_shape`, `add_text`, `connect_shapes`, `list_shapes`, `delete_shape` |
| Stencils | `list_stencils`, `list_stencil_masters`, `search_stencil_shapes`, `add_stencil_shape` |
| Whole diagrams | `build_diagram` |
| Racks | `add_rack`, `add_rack_device`, `list_rack` |
| SVG | `svg_to_visio` |

### create_visio_file

Creates a new diagram. Blank documents get a fixed page size so Visio does not reset it to the printer paper.

```json
{
  "save_path": "C:\\Diagrams\\MyDiagram.vsdx",
  "template_path": "optional .vstx/.vst template",
  "page_width": 11.69,
  "page_height": 8.27
}
```

`page_width/page_height` are ignored when a template is used.

### open_visio_file / close_document

```json
{ "file_path": "C:\\Diagrams\\MyDiagram.vsdx" }
```

```json
{ "file_path": "C:\\Diagrams\\MyDiagram.vsdx", "save_changes": true }
```

### set_page_setup

Sets the size, and optionally the drawing scale, of the first page.

```json
{ "file_path": "C:\\Diagrams\\rack.vsdx", "width": 11.69, "height": 8.27, "drawing_scale": 12 }
```

`drawing_scale` is the real inches one page inch stands for (12 means 1:12). Stencils from
VisioCafe (HPE, Dell, ...) are drawn in real sizes and need a scaled page. After this,
coordinates you pass to the shape tools are in real (drawing) inches.

### add_shape

```json
{
  "file_path": "C:\\Diagrams\\flow.vsdx",
  "shape_type": "Rectangle",
  "x": 2.0, "y": 4.0, "width": 1.5, "height": 0.8,
  "text": "Start",
  "fill_color": "#DBEAFE",
  "line_color": "#2563EB",
  "line_weight": 1.5,
  "line_pattern": "solid",
  "text_color": "#1E293B",
  "font_size": 12,
  "bold": true,
  "rounding": 0.1
}
```

`shape_type` examples: `Rectangle`, `Circle`, `Line`. An unknown type falls back to a rectangle.
`x, y` is the shape's centre. `rounding` is the corner radius in inches. Returns the shape ID.

### connect_shapes

```json
{
  "file_path": "C:\\Diagrams\\flow.vsdx",
  "shape1_id": 1,
  "shape2_id": 2,
  "connector_type": "Straight",
  "line_color": "#475569",
  "line_weight": 2,
  "line_pattern": "dash",
  "begin_arrow": false,
  "end_arrow": true,
  "label": "HTTPS",
  "text_color": "#475569",
  "font_size": 10
}
```

`connector_type` is `Dynamic` (re-routes at right angles), `Straight` or `Curved`.
Straight and curved connectors run centre to centre.

### add_text

```json
{ "file_path": "C:\\Diagrams\\flow.vsdx", "shape_id": 1, "text": "Start" }
```

### list_shapes / delete_shape

```json
{ "file_path": "C:\\Diagrams\\flow.vsdx" }
```

```json
{ "file_path": "C:\\Diagrams\\flow.vsdx", "shape_id": 3 }
```

### list_stencils / list_stencil_masters / search_stencil_shapes

Stencils hold ready-made shapes ("masters"): servers, switches, firewalls, clouds, racks.

```json
{ "query": "server" }
```

```json
{ "stencil": "SERVER_U" }
```

```json
{ "query": "DL380", "stencil": "HPE-ProLiant", "limit": 20 }
```

`list_stencils` lists the files found, `list_stencil_masters` lists the shapes in one stencil
(name or full path), and `search_stencil_shapes` searches shape names across stencils whose file
name contains `stencil`.

### add_stencil_shape

```json
{
  "file_path": "C:\\Diagrams\\net.vsdx",
  "stencil": "SERVER_U",
  "master": "Database server",
  "x": 4.0, "y": 5.0,
  "text": "DB-01",
  "width": 1.2, "height": 1.2,
  "fill_color": "#E2E8F0",
  "font_size": 11,
  "bold": true
}
```

Position is the centre. A fill colour replaces the stencil's 3D shading with a flat colour.
Returns the shape ID.

### build_diagram

Draws a whole diagram from one spec. Nodes are arranged in tiers (rows, or columns for
`left-right`) so everything fits the page. See [recipe B](#b-a-network-diagram-from-one-spec).

```json
{
  "file_path": "C:\\Diagrams\\net.vsdx",
  "replace_existing": false,
  "spec": {
    "title": "Heading", "subtitle": "Optional line under it",
    "page": { "width": 11.69, "height": 8.27 },
    "direction": "top-down",
    "icon_size": 1.0,
    "nodes": [ { "id": "fw", "label": "Firewall", "tier": 1, "stencil": "PERIPH_U", "master": "Firewall" } ],
    "links": [ { "from": "fw", "to": "sw", "label": "10GbE", "type": "straight", "arrow": "end" } ],
    "groups": [ { "label": "Core", "nodes": ["fw", "sw"], "color": "#3B82F6" } ]
  }
}
```

| Field | Meaning |
|---|---|
| `direction` | `top-down` (tier 0 on top, default) or `left-right` |
| `nodes[].tier` | 0 is the first row/column. Nodes with the same tier share a row |
| `nodes[].stencil`, `master` | Optional icon. Without them the node is a labelled box |
| `nodes[].fill_color`, `text_color`, `width`, `height`, `icon_size` | Optional per-node style/size |
| `links[].type` | `straight` (default), `curved`, `dynamic` |
| `links[].arrow` | `none` (default), `end`, `begin`, `both` |
| `links[].color`, `pattern`, `weight`, `label` | Optional line style and text |
| `groups` | A dashed box behind the listed node ids, with an upper-case label |
| `replace_existing` | `true` clears the page first; default adds to it |

Returns JSON with the Visio shape ID of every node.

### add_rack / add_rack_device / list_rack

Rack stencils are in real inches, so `add_rack` sets the page to a 1:12 scale (A4 landscape then
spans about 140 x 99 in). Coordinates are in real inches.

```json
{ "file_path": "C:\\Diagrams\\rack.vsdx", "x": 30, "y": 10, "name": "Rack A01" }
```

Returns the rack's shape ID. Other options: `stencil` (default `Dell-Racks`), `master`
(default `4220 Rack Frame`), `u_height` (42), `u1_offset` (4.0), `drawing_scale` (12).

```json
{
  "file_path": "C:\\Diagrams\\rack.vsdx",
  "rack_id": 1,
  "stencil": "HPE-ProLiant-DL",
  "master": "DL380 Gen10 SFF front",
  "u_position": 38,
  "u_size": 2,
  "label": "app-01",
  "callout": true
}
```

`u_position` is the lowest unit the device uses (U1 is the bottom). Overlaps and positions outside
the rack are rejected. `callout` draws a leader line and a "app-01 (U38-39)" label.

```json
{ "file_path": "C:\\Diagrams\\rack.vsdx", "rack_id": 1 }
```

`list_rack` returns the devices, the free U ranges and used/free counts.

### export_diagram

```json
{ "file_path": "C:\\Diagrams\\net.vsdx", "output_path": "C:\\Diagrams\\net.png" }
```

The extension picks the format: `.png .jpg .jpeg .gif .bmp .svg .pdf`. For very large pages with many
images, export `.pdf` or `.svg` instead of a bitmap.

### svg_to_visio

```json
{
  "svg_path": "C:\\Diagrams\\architecture.svg",
  "output_path": "C:\\Diagrams\\architecture.vsdx",
  "export_png": "C:\\Diagrams\\architecture.check.png",
  "group_shapes": true
}
```

`export_png` and `group_shapes` are optional (grouping is on by default). Details in
[SVG import in depth](#svg-import-in-depth).

## Recipes

### A. Flowchart shape by shape

The assistant's calls, in order (paths shortened):

```json
// 1. create_visio_file
{ "save_path": "C:\\Diagrams\\FlowChart.vsdx" }

// 2. add_shape  -> returns ID 1
{ "file_path": "C:\\Diagrams\\FlowChart.vsdx", "shape_type": "Rectangle",
  "x": 2, "y": 4, "width": 1.5, "height": 0.8, "text": "Start", "fill_color": "#DBEAFE" }

// 3. add_shape  -> returns ID 2
{ "file_path": "C:\\Diagrams\\FlowChart.vsdx", "shape_type": "Circle",
  "x": 6, "y": 4, "width": 1, "height": 1, "text": "End", "fill_color": "#DCFCE7" }

// 4. connect_shapes
{ "file_path": "C:\\Diagrams\\FlowChart.vsdx", "shape1_id": 1, "shape2_id": 2,
  "connector_type": "Straight", "end_arrow": true, "label": "next" }

// 5. export_diagram
{ "file_path": "C:\\Diagrams\\FlowChart.vsdx", "output_path": "C:\\Diagrams\\FlowChart.png" }
```

### B. A network diagram from one spec

`build_diagram` needs one call. This is [`examples/three_tier_app.json`](examples/three_tier_app.json),
shortened:

```json
{
  "title": "Three-tier Web Application",
  "direction": "left-right",
  "nodes": [
    {"id": "users", "label": "Users",         "tier": 0, "fill_color": "#E2E8F0"},
    {"id": "lb",    "label": "Load Balancer", "tier": 1, "fill_color": "#FEF3C7"},
    {"id": "web1",  "label": "Web 1", "tier": 2, "stencil": "SERVER_U", "master": "Web server"},
    {"id": "web2",  "label": "Web 2", "tier": 2, "stencil": "SERVER_U", "master": "Web server"},
    {"id": "db",    "label": "Database", "tier": 3, "stencil": "SERVER_U", "master": "Database server"}
  ],
  "links": [
    {"from": "users", "to": "lb",   "arrow": "end", "label": "HTTPS"},
    {"from": "lb",    "to": "web1", "arrow": "end"},
    {"from": "lb",    "to": "web2", "arrow": "end"},
    {"from": "web1",  "to": "db",   "arrow": "end"},
    {"from": "web2",  "to": "db",   "arrow": "end"}
  ],
  "groups": [ {"label": "Web tier", "nodes": ["web1", "web2"], "color": "#16A34A"} ]
}
```

Finding the stencil and master names: ask for `list_stencils` with a query (`server`), then
`list_stencil_masters` for the one you want (`SERVER_U` has Web server, Database server, ...), or
`search_stencil_shapes` when you only know part of a product name.

### C. A rack diagram

```json
// 1. add_rack -> returns rack_id
{ "file_path": "C:\\Diagrams\\rack.vsdx", "x": 30, "y": 10, "name": "Rack A01" }

// 2. add_rack_device (repeat for each device)
{ "file_path": "C:\\Diagrams\\rack.vsdx", "rack_id": 1,
  "stencil": "Juniper EX Series", "master": "EX4300-24T Front", "u_position": 40 }
{ "file_path": "C:\\Diagrams\\rack.vsdx", "rack_id": 1,
  "stencil": "HPE-ProLiant-DL", "master": "DL380 Gen10 SFF front",
  "u_position": 38, "u_size": 2, "label": "app-01" }

// 3. list_rack: what is mounted and which units are free
{ "file_path": "C:\\Diagrams\\rack.vsdx", "rack_id": 1 }
```

### D. SVG to Visio

Ask: *"Convert `C:\Diagrams\architecture.svg` to `C:\Diagrams\architecture.vsdx` and export a PNG to compare."*
The tool call:

```json
{
  "svg_path": "C:\\Diagrams\\architecture.svg",
  "output_path": "C:\\Diagrams\\architecture.vsdx",
  "export_png": "C:\\Diagrams\\architecture.check.png"
}
```

A typical result (the counts depend on the SVG):

```
Created C:\Diagrams\architecture.vsdx (13.33 x 7.50 in): 150 lines, 31 pictures, 60 texts. Grouped 12 SVG group(s). Skipped or approximated 2 thing(s): filter, gradient. Exported to: C:\Diagrams\architecture.check.png
```

Open the `.vsdx` in Visio: every element is a native, editable shape named after its SVG `id`.

### E. Run without an AI client

The scripts in `examples/` call the same functions directly (use the project's `.venv`):

```powershell
# JSON spec -> .vsdx (+ PNG and PDF)
python examples\build_from_spec.py examples\hci_network.json out\hci.vsdx --png --pdf

# SVG -> .vsdx, and an Edge-vs-Visio comparison image per file
python examples\svg_compare.py out my1.svg my2.svg
```

`svg_compare.py` writes `.vsdx`, `.visio.png`, `.edge.png` and `.side.png` (the two renderings side by
side). Judge by the side-by-side picture; the printed difference score is only a rough hint.

## SVG import in depth

**Goal:** put an SVG in, get an editable Visio file out, with the same layout, colours, text and
layering. Each SVG element becomes a Visio shape you can select, restyle and move.

### What is converted

| SVG | Visio result |
|---|---|
| `<g>`, transforms (translate, scale, rotate, skew, matrix), `viewBox` | Positions are resolved; runs of shapes from one `<g>` become Visio groups |
| `rect` (with `rx/ry`), `circle`, `ellipse` | Rectangle (rounded) / ellipse shapes |
| `line`, `polyline`, `polygon` | Line shapes with dash patterns and arrowheads |
| `path` (lines, curves, arcs) | Curves are flattened to line segments. Several subpaths in one `path` are one shape, even-odd fill, so holes stay holes |
| `text`, `tspan` | Text shapes with per-text font family, size, bold, anchor and measured width |
| draw.io `<foreignObject>` labels | Read as text with wrapping, alignment and bullets; the "Text is not SVG" fallback is not drawn |
| `<use>`, `<symbol>`, `<switch>` | Expanded; `switch` picks the same branch a browser would |
| `<style>`, classes, `#id`, `style=""`, inheritance | Resolved before drawing |
| opacity, `fill-opacity`, `stroke-opacity` | Transparency on fill, line and text |
| `<marker>` arrowheads | Arrowhead type and size follow the marker |
| `<image>` PNG/JPEG, SVG data URIs, files next to the SVG | Pictures. Embedded SVG icons are rasterised with `resvg-py` (headless Edge is only a fallback) |

### Naming, layers, groups

- Shapes are named after the SVG `id` (made unique).
- Layers: `Shapes`, `Connectors`, `Labels`, `Icons`. Toggle or lock them in Visio's Layer Properties.
- With `group_shapes: true` (default), consecutive shapes from one `<g>` are grouped and the group
  takes the `<g>` id. Set it to `false` for a flat shape list.

### Not supported (reported in the result)

- `clip-path`, `mask`, `filter`, `<pattern>`, nested `<svg>`, `<symbol>` with its own `viewBox`
- Gradients are approximated by their average colour
- Rotated images are drawn unrotated; `preserveAspectRatio` on images is not applied
- `textLength`, per-character `dx/dy`, and `<tspan>` with a different style inside one line
- Connectors are ordinary line shapes, not glued to the shapes at their ends
- Nonzero fill is drawn as even-odd (only differs for overlapping same-direction subpaths)

### Performance and tips

- Time grows with the number of elements: a diagram with about 200 lines, 16 paths and 14 pictures
  takes around 90 seconds, because Visio is written to through COM one cell at a time.
- Pass `export_png` and compare against the SVG when fidelity matters.
- For a faster, flatter result use `group_shapes: false`.
- Fonts: families map to installed fonts (serif to Times New Roman, sans-serif to Arial,
  monospace to Consolas, Helvetica to Arial, system-ui to Segoe UI). Install the fonts the SVG uses
  for the best text fit.

## Troubleshooting

| Symptom | What to do |
|---|---|
| Server exits at start | Visio is not installed or not registered for COM. Open Visio manually once |
| `Visio could not render this page as .png` | The page is very large. Export `.pdf` or `.svg` |
| COM errors, tools hang | Close dialogs in Visio (a modal dialog blocks automation). Check Task Manager for a stuck `VISIO.EXE` that you started, and end it |
| Shape type not recognised | It falls back to a rectangle. Check the spelling, or use a stencil master |
| Stencil or master not found | Use `list_stencils`, `list_stencil_masters` or `search_stencil_shapes` for the exact names |
| Stencil shapes are tiny or huge | They are drawn in real units. Use `set_page_setup` with a `drawing_scale` (or `add_rack`, which sets it) |
| My open diagram changed | The server works inside the Visio you have open. Save and close your own work before running experiments |
| SVG text wraps or sits differently | Install the SVG's fonts, or compare with `examples/svg_compare.py` |
| SVG result looks flat | Gradients, filters and masks are not supported; see the list above |
| `svg_to_visio` says output must end with `.vsdx` | Change `output_path` |

## Development

```powershell
pip install -e ".[test]"
python -m pytest tests --deselect tests/test_e2e_visio.py   # fast, does not touch Visio
python -m pytest tests                                       # full run, needs Visio; close your own Visio first
```

The style parsing, tier layout, rack maths and SVG parsing are tested without Visio. The e2e tests
create documents in a real Visio.

Layout of the code:

| File | Role |
|---|---|
| `visio_mcp_server/visio_server.py` | MCP tools and the Visio COM calls |
| `visio_mcp_server/styles.py` | Colour, line and text style parsing |
| `visio_mcp_server/diagram.py` | Spec validation, tier layout and rack maths (no Visio needed) |
| `visio_mcp_server/svg_import.py` | SVG parser to flat primitives (rect, ellipse, polyline, path, text, image) |
| `visio_mcp_server/svg_html.py` | draw.io HTML label reader |
| `visio_mcp_server/svg_text.py` | Font lookup and text width measuring |
| `examples/` | Specs and scripts that run without an AI client |

See [CHANGELOG.md](CHANGELOG.md) and [docs/DEVELOPMENT_NOTES.md](docs/DEVELOPMENT_NOTES.md) for the
Visio automation pitfalls behind the code.

## License

MIT. See the LICENSE file.
