# Development notes: Visio automation pitfalls

Things that cost time while building this server. Each one was reproduced against a real Visio
(version 15.0); the fix is in `visio_mcp_server/` unless noted.

## Shapes and geometry

| Symptom | Cause | What the code does |
|---|---|---|
| `Inappropriate target object` when gluing a connector | Rack-equipment masters are 1-D shapes | `connectable()` draws an invisible 2-D anchor at the centre and glues to that |
| Every shape moved by the same vector after adding connectors | `page.Drop(connector, 0, 0)` on a large, dense page (26 x 38 in, ~75 shapes) shifted the page origin. Not reproducible on small pages | `make_connector` drops the connector on its first shape |
| Same shift when dropping an image master at the page corner | A shape hanging off the page | `drop_stencil_shape` drops at the page centre, then the caller moves it |
| `BoundingBox` returns `(0, 0, -1, -1)` | Seen after text or style is set, and on documents that were just reopened | Layout uses cells (`PinX`, `LocPinX`, `Width`) via `center_on`; group boxes come from the layout, not from bounding boxes |
| Placing by pin put icons off-centre | Pins differ per stencil | `center_on` / `move_by_bbox` place by frame centre or bottom edge |
| Text of rack shapes is invisible | 1-D shapes do not draw their text | Labels are separate text boxes |
| Setting `PinX` on an image master had no effect | `FormulaForceU` is overridden by the master's own formula | `set_formula` tries `FormulaU` first and forces only if the cell is guarded |

## Styling

- `FormulaForceU` is needed for stencil shapes: many cells are guarded (`Cell is guarded`).
- Stencil fills are often gradients; set `FillGradientEnabled = 0` or `FillForegnd` shows nothing.
  This flattens the icon's 3-D shading.
- Group shapes (`Shape.Type == 2`) need fill and line set on every sub-shape.
- `TextBkgnd = RGB(255,255,255)` renders grey on connector labels; set only `TextBkgndTrans = 0%`.
- `ConLineRouteExt` alone does not make a straight line: also set `ShapeRouteStyle = 16`
  (centre to centre), otherwise the connector is routed at right angles.

## Pages and scale

- `DrawingSizeType = 0` means "same as the printer paper": Visio resets a custom page size on a
  machine with different paper. Set `DrawingSizeType = 2` and the orientation (`apply_page_setup`).
- VisioCafe stencils (HPE, Dell, Juniper, ...) are drawn in real units (19 in rack width, 1U = 1.75 in).
  `add_rack` sets the page scale to 1:12 and all coordinates passed to rack tools are real inches.
- `DrawRectangle` and `Drop` take drawing units, so on a scaled page a "1 inch" box is 1 real inch.

## Stencils

- Only `.vss`, `.vssx` and `.vssm` are discovered by name. Some vendors ship stencils as `.vsd`
  (F5): pass the full path.
- Some stencil files are corrupt (a file in the VMware folder), so searching all stencils can raise
  `Visio cannot open the file`; `search_stencil_shapes` requires a stencil name fragment for this reason.
- Stencil names are not what you expect: `PERIPH_U` holds switch/firewall/router/user,
  `RCKEQP_U` the Visio rack shapes, `SERVER_U` server icons.

## Export

- `Page.Export` to PNG/JPG can fail ("unable to complete exporting") on very large pages with many
  embedded images. `ExportAsFixedFormat` (PDF) and SVG export are reliable.
- PDF pages are limited to 200 in; a shifted page origin shows up as a 14400 x 14400 pt PDF.

## Process and testing

- The server prefers `GetActiveObject` so the user's own Visio is never quit; it only calls `Quit()`
  on an instance it started (`owns_visio_app`).
- COM calls are synchronous inside `async def` tools. This works because everything runs on one
  thread, but a slow Visio call blocks the event loop.
- Visio start-up can take several seconds; MCP clients with short start-up timeouts (Codex:
  `startup_timeout_sec`) may need a larger value, or Visio opened first.
- End-to-end tests (`tests/test_e2e_visio.py`) start Visio and close it afterwards; they are skipped
  when Visio is missing. The Visio-free logic lives in `diagram.py` and `styles.py` so CI can run it.
- The page-origin shift above is not guarded by a test: it needs a large, dense page to appear.

## Data handling

- Real diagrams contain hostnames and IPs. Keep generators and outputs for internal systems in
  `private/` (git-ignored) and never commit vendor stencils: their licences forbid redistribution.
