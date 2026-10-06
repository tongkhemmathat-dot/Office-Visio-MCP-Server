# Changelog

## 1.1.0 - 2026-10-06

Focus: making the server useful for system and network engineers (real hardware stencils,
whole-diagram builds, rack diagrams) and fixing the problems found along the way.

### Added
- `build_diagram`: draw a complete diagram from one JSON spec (tiered auto-layout, stencil icons,
  labelled links, group boxes, title). `examples/` holds two specs and a `build_from_spec.py` CLI.
- Rack diagrams: `add_rack`, `add_rack_device` (mount by U position, rejects overlaps and out-of-range
  positions) and `list_rack` (devices and free units). The layout is stored in the `.vsdx`.
- Stencils: `list_stencils`, `list_stencil_masters`, `search_stencil_shapes`, `add_stencil_shape`.
  Searches Visio Content, `Documents\My Shapes` (recursively) and the user's stencil paths.
- `export_diagram` (PNG, JPG, GIF, BMP, SVG, PDF), `delete_shape`, `set_page_setup`.
- Style options on `add_shape`, `connect_shapes` and `add_stencil_shape` (colors, line weight and
  pattern, arrowheads, labels, fonts).
- `docs/TEAM_GUIDE.md` (Thai): setup for Claude Code/Desktop, Antigravity, Cursor, VS Code, Gemini CLI
  and Codex CLI, stencil sources and licences, prompts, spec format, rack workflow, troubleshooting.
- `docs/DEVELOPMENT_NOTES.md`: Visio automation pitfalls found while building this.
- Tests: unit tests for the Visio-free logic and end-to-end tests that drive a real Visio
  (skipped when Visio is not installed).

### Changed
- New files get a custom A4 landscape page that Visio no longer resets to the printer's paper
  (`create_visio_file` takes `page_width` / `page_height`).
- Tools are wrapped by one decorator that reports failures as `Error <action>: <reason>`.
- Paths are normalised, so one file reached by different spellings is one open document.
- The server attaches to a running Visio and only quits an instance it started itself.
- `mcp` is pinned below 2.0 (2.x removed `mcp.server.fastmcp`).
- Shape lookup uses `Shapes.ItemFromID`; every operation uses the document's first page instead of
  whatever page is active.

### Fixed
- Straight and curved connectors were invisible (`LinePattern = 0`) and Straight/Curved were routed at
  right angles.
- Connectors could not glue to 1-D shapes (rack equipment): an invisible anchor is used.
- A connector dropped at the page corner shifted every shape on large, dense pages: connectors are
  now dropped on their first shape.
- Connector labels rendered with a grey background.
- Fill colors had no effect on stencil shapes (guarded cells and gradient fills).
- `add_shape` overwrote the shape's text with its type; it can now be given `text`.
- A shape could not be connected to itself.
- Failed bitmap exports now explain that PDF or SVG should be used instead.

### Known issues
- Bitmap export (PNG/JPG) can fail on very large pages with many images; PDF and SVG work.
- Stencil images from Synology are embedded bitmaps: they cannot be recoloured.
- Rack U positions use a fixed offset (`u1_offset`, 4 in for the Dell 4220 frame), not the real
  rail holes.
- File paths are not restricted: the server opens and writes any path the client names.
- Tested on Visio 2013 (version 15.0) with Claude Code only; other MCP clients and Visio versions
  were checked against documentation, not run.

## 1.0.1

Upstream release (GongRzhe/Office-Visio-MCP-Server): create/open files, rectangle/circle/line shapes,
connectors, text, shape listing.
