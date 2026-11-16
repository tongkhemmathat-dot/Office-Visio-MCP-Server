---
name: visio-mcp
description: Use when creating or editing Microsoft Visio diagrams (.vsdx) through the visio MCP server tools - flowcharts, network or system diagrams, rack layouts, stencil shapes, exporting to PNG/PDF/SVG, or converting an SVG into an editable Visio file. Triggers on "Visio", ".vsdx", "วาด diagram", "สร้างไดอะแกรม", "แปลง SVG เป็น Visio", "rack diagram", "stencil".
---

# Using the Visio MCP server

Tools drive a real Visio on Windows through COM. Pick the tool by the job, build, then look at the result before saying it is done.

## 1. Pick the path

| Request | Use |
|---|---|
| A network / system / tiered diagram with many nodes | `build_diagram` (one call, automatic layout) |
| A few shapes, or a precise manual layout | `create_visio_file` then `add_shape` / `add_stencil_shape`, then `connect_shapes` |
| Servers mounted in a rack | `add_rack`, `add_rack_device`, `list_rack` |
| An existing SVG to turn into Visio | `svg_to_visio` |
| Edit a diagram that already exists | `open_visio_file`, `list_shapes`, then add/delete/connect |
| A picture or PDF of the result | `export_diagram` |

Default to `build_diagram` for anything with more than about six nodes. Hand-placing shapes costs many calls and the layout is worse.

## 2. Ground rules

- **Units are inches**, origin bottom-left, y grows up. Default page is A4 landscape, 11.69 x 8.27 in. Positions are shape centres.
- **A bare file name goes to the Documents folder.** Pass a full path when the user named a location, and tell the user the final path.
- **Shape IDs link the steps.** Keep the ID returned by `add_shape`, `add_stencil_shape` and `add_rack`. If it is lost, call `list_shapes`.
- **Colours**: `#RRGGBB` or a name. **Line patterns**: `solid`, `dash`, `dot`, `dashdot`. Weights and font sizes are points.
- **Never guess stencil or master names.** Find them first with `list_stencils` (query), `list_stencil_masters`, or `search_stencil_shapes` when only part of a model name is known. A wrong name fails.
- **The server uses the user's open Visio.** It attaches to a running Visio and creates documents inside it. Do not close or kill Visio processes, and warn the user before a large run if they may have unsaved work open. If a modal dialog is open in Visio, tools hang until it is dismissed.
- **Icons**: use stencil masters or real icon files (for example official AWS or Kubernetes sets). Do not hand-draw icons unless none exist.

## 3. Workflows

### Tiered diagram (`build_diagram`)

1. Find stencils for the node types (`list_stencils` with `server`, then `list_stencil_masters`). Nodes without `stencil` and `master` become labelled boxes, which is fine for logical views.
2. Write the spec: `nodes` (`id`, `label`, `tier`, optional `stencil`/`master`), `links` (`from`, `to`, optional `label`, `arrow`, `pattern`, `color`), optional `groups`.
3. Call `build_diagram` with `file_path` and the `spec`.
4. `export_diagram` to PNG and look at it. Fix and call again with `replace_existing: true`.

Spec rules that matter:
- `tier` 0 is the first row (or column when `direction` is `left-right`). Nodes with equal tier share a row.
- `arrow` is `none` (default), `end`, `begin` or `both`. Set `end` for flows that have a direction.
- `groups` draw a dashed box behind the listed node ids. Use them for zones, VLANs, tiers.
- Keep about 20 nodes per page. Split larger systems into several diagrams.

Minimal example:

```json
{
  "file_path": "C:\\Diagrams\\app.vsdx",
  "spec": {
    "title": "Web app",
    "direction": "left-right",
    "nodes": [
      {"id": "users", "label": "Users", "tier": 0},
      {"id": "lb", "label": "Load Balancer", "tier": 1},
      {"id": "web", "label": "Web", "tier": 2, "stencil": "SERVER_U", "master": "Web server"},
      {"id": "db", "label": "Database", "tier": 3, "stencil": "SERVER_U", "master": "Database server"}
    ],
    "links": [
      {"from": "users", "to": "lb", "arrow": "end", "label": "HTTPS"},
      {"from": "lb", "to": "web", "arrow": "end"},
      {"from": "web", "to": "db", "arrow": "end"}
    ]
  }
}
```

### Manual diagram

1. `create_visio_file` with `save_path` (add `page_width`/`page_height` for a non-A4 page).
2. `add_shape` for each box (or `add_stencil_shape` for icons). Style in the same call: `fill_color`, `line_color`, `text_color`, `font_size`, `bold`, `rounding`.
3. `connect_shapes` with the two IDs. Use `Straight` for clean centre-to-centre lines, `Dynamic` for right-angle routing, `Curved` for arcs. Set `end_arrow: true` for direction and `label` for a caption.
4. Plan coordinates first on a grid (for example 2 in between columns, 1.5 in between rows) so shapes do not overlap.

### Rack diagram

1. `add_rack` with `x` and `y` in **real inches**. It sets the page to a 1:12 scale, so A4 landscape is about 140 x 99 in. `x = 30, y = 10` puts a 42U rack on the left. Keep the returned rack ID.
2. `add_rack_device` per device: `stencil`, `master`, `u_position` (lowest unit, U1 at the bottom) and `u_size` (a 2U server needs `u_size: 2`). Overlaps and out-of-range positions are rejected, so read the error and pick another U.
3. `list_rack` to report used and free units.
4. Rack and server stencils from HPE, Dell and similar are drawn in real inches. Do not size them by hand on an unscaled page; use `set_page_setup` with `drawing_scale` first.

### SVG to Visio (`svg_to_visio`)

1. Call with `svg_path` and `output_path` ending in `.vsdx`. Add `export_png` to get a picture to compare with the SVG.
2. Read the result text. It lists counts, groups made, and anything skipped or approximated.
3. Report skipped items to the user. Not supported: `clip-path`, `mask`, `filter`, `<pattern>`, nested `<svg>`. Gradients become their average colour. Connectors are plain lines, not glued to shapes.
4. Large files are slow (around 90 seconds for 200 lines plus pictures). Say so before starting rather than retrying.
5. Result shapes are named after SVG ids, on layers `Shapes`, `Connectors`, `Labels`, `Icons`, and grouped per SVG `<g>`. Set `group_shapes: false` for a flat list.
6. Judge fidelity by viewing the exported PNG next to the SVG, not by assumption.

## 4. Verify before finishing

1. `export_diagram` to PNG (or PDF for very large pages) and look at it. Check: nothing overlaps, labels are readable, arrows point the right way, nothing is off the page.
2. `list_shapes` when counts or IDs matter.
3. Tell the user the file path and what you built. Do not close the document unless asked; `close_document` saves by default.

## 5. Common failures

| Symptom | Fix |
|---|---|
| Stencil or master not found | Look it up with `list_stencils`, `list_stencil_masters`, `search_stencil_shapes`; copy the exact name |
| Stencil shapes look tiny or huge | Real-unit stencil on an unscaled page: `set_page_setup` with `drawing_scale: 12`, or start with `add_rack` |
| Unknown `shape_type` | It silently becomes a rectangle; check the spelling |
| Rack device rejected | The U range overlaps another device or exceeds the rack; use `list_rack` to see free ranges |
| PNG export fails on a huge page | Export `.pdf` or `.svg` instead |
| Tools hang or return COM errors | A dialog is open in Visio; ask the user to dismiss it. Do not kill Visio |
| `svg_to_visio` output error | `output_path` must end with `.vsdx` |
| Added shapes land on an old drawing | `build_diagram` adds to the page by default; pass `replace_existing: true` to start clean |

## 6. Style defaults when the user does not specify

- Boxes `#DBEAFE` fill with `#2563EB` line, text `#1E293B`; success/neutral `#DCFCE7` / `#E2E8F0`; warning `#FEF3C7`.
- Link lines `#475569`, 1.5 pt. Use `dash` for backup or management links, a colour for a special path (for example `#0D9488` for a LAG), and a label for protocol or bandwidth.
- Labels short and in the user's language. Put model or spec on a second line, for example `DB-01\nPostgreSQL 16`.
