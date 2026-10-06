"""
Visio MCP Server

This MCP server provides tools for creating and editing Visio files.
It drives Microsoft Visio through its COM automation API using pywin32.
"""

import atexit
import functools
import json
import os
import sys
import time
import winreg
from typing import List, Optional

import win32com.client
import win32com.client.dynamic
from mcp.server.fastmcp import FastMCP

try:
    from visio_mcp_server.diagram import (RACK_UNIT, check_rack_fit, free_ranges, group_bounds,
                                          normalize_spec, rack_device_bottom, tier_layout)
    from visio_mcp_server.styles import apply_shape_style, style_geometry, validate_style
except ImportError:
    # Started as a plain script without the package on PYTHONPATH
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    from visio_mcp_server.diagram import (RACK_UNIT, check_rack_fit, free_ranges, group_bounds,
                                          normalize_spec, rack_device_bottom, tier_layout)
    from visio_mcp_server.styles import apply_shape_style, style_geometry, validate_style

mcp = FastMCP("visio-server")

# --- Constants ---------------------------------------------------------------

DEFAULT_SAVE_PATH = os.path.expandvars(r"%USERPROFILE%\Documents")
A4_LANDSCAPE = (11.69, 8.27)  # inches

STENCIL_EXTENSIONS = (".vssx", ".vss", ".vssm")
MY_SHAPES_DIR = os.path.join(DEFAULT_SAVE_PATH, "My Shapes")

# Visio enumeration values used below
VIS_OPEN_RO = 2
VIS_OPEN_HIDDEN = 64
VIS_SECTION_USER = 242      # User-defined cells section
DRAWING_SIZE_CUSTOM = 2      # page size is not tied to the printer's paper
ORIENTATION_PORTRAIT = 1
ORIENTATION_LANDSCAPE = 2
ROUTE_CENTER_TO_CENTER = 16  # ShapeRouteStyle: straight line between shape centres
LINE_ROUTE_STRAIGHT = 1      # ConLineRouteExt
LINE_ROUTE_CURVED = 2
ARROW_TRIANGLE = 4
ARROW_NONE = 0

# --- Application and document state ------------------------------------------

visio_app = None
owns_visio_app = False       # True only if this server started Visio itself
open_documents = {}          # doc_key -> Visio Document
stencil_documents = {}       # normalised stencil path -> Visio Document


class ToolError(Exception):
    """A problem with the tool's input that should be reported to the caller as-is."""


def check_visio_installed() -> bool:
    """Check if Visio is installed without launching it."""
    try:
        winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Visio.Application")
        return True
    except OSError:
        return False


def get_visio_app():
    """Return the Visio Application object, attaching to a running Visio or starting one.

    Only an instance started here is quit on shutdown, so the user's own Visio
    session is never closed by the server.
    """
    global visio_app, owns_visio_app

    if visio_app is not None:
        return visio_app

    try:
        visio_app = win32com.client.GetActiveObject("Visio.Application")
        owns_visio_app = False
        return visio_app
    except Exception:
        pass  # not running; start it below

    errors = []
    for factory in (win32com.client.Dispatch, win32com.client.dynamic.Dispatch, win32com.client.DispatchEx):
        try:
            app = factory("Visio.Application")
            app.Visible = True
            visio_app = app
            owns_visio_app = True
            return visio_app
        except Exception as e:
            errors.append(str(e))
    raise RuntimeError(f"Failed to initialize Visio application after multiple attempts. Errors: {errors}")


def close_visio_app() -> None:
    """Close the documents opened by this server and quit Visio if the server started it."""
    global visio_app, owns_visio_app

    for doc in list(open_documents.values()) + list(stencil_documents.values()):
        try:
            doc.Close()
        except Exception:
            pass
    open_documents.clear()
    stencil_documents.clear()

    if visio_app is not None and owns_visio_app:
        try:
            visio_app.Quit()
        except Exception:
            pass
    visio_app = None
    owns_visio_app = False


def tool(action: str):
    """Register an MCP tool whose exceptions are returned as 'Error <action>: <reason>' text."""
    def decorator(fn):
        @functools.wraps(fn)
        async def wrapper(*args, **kwargs):
            try:
                return await fn(*args, **kwargs)
            except Exception as e:
                return f"Error {action}: {e}"
        return mcp.tool()(wrapper)
    return decorator


# --- Paths and documents -----------------------------------------------------

def resolve_path(file_path: str) -> str:
    """Return the canonical absolute path for a document.

    A bare filename is placed in the default save directory, so every tool
    derives the same path (and the same open_documents key) for the same input.
    """
    if os.path.dirname(file_path) == '':
        file_path = os.path.join(DEFAULT_SAVE_PATH, file_path)
    return os.path.abspath(file_path)


def doc_key(file_path: str) -> str:
    """Key for open_documents: case- and slash-insensitive form of the path."""
    return os.path.normcase(resolve_path(file_path))


def apply_page_setup(doc, width: float, height: float, drawing_scale: Optional[float] = None) -> None:
    """Give the document's first page a custom size that Visio will not reset to the printer paper."""
    sheet = get_page(doc).PageSheet
    sheet.Cells("DrawingSizeType").FormulaForceU = str(DRAWING_SIZE_CUSTOM)
    sheet.Cells("PrintPageOrientation").FormulaForceU = str(
        ORIENTATION_LANDSCAPE if width > height else ORIENTATION_PORTRAIT)
    sheet.Cells("PageWidth").FormulaForceU = f"{width} in"
    sheet.Cells("PageHeight").FormulaForceU = f"{height} in"
    if drawing_scale is not None:
        sheet.Cells("PageScale").FormulaForceU = "1 in"
        sheet.Cells("DrawingScale").FormulaForceU = f"{drawing_scale} in"


def create_document(template_path: Optional[str] = None, save_path: Optional[str] = None,
                    page_size: Optional[tuple] = A4_LANDSCAPE):
    """Create and save a new document; return (document, path).

    page_size is ignored when a template is used, since templates define their own page.
    """
    app = get_visio_app()

    if not save_path:
        save_path = f"New_Diagram_{int(time.time())}.vsdx"
    save_path = resolve_path(save_path)
    os.makedirs(os.path.dirname(save_path), exist_ok=True)

    use_template = bool(template_path) and os.path.exists(template_path)
    doc = app.Documents.Add(template_path if use_template else "")
    if not use_template and page_size:
        apply_page_setup(doc, *page_size)
    doc.SaveAs(save_path)
    open_documents[doc_key(save_path)] = doc
    return doc, save_path


def get_document(file_path: str, create_if_missing: bool = False):
    """Return the open Visio document for file_path, opening (or creating) it if needed."""
    path = resolve_path(file_path)
    key = doc_key(file_path)

    doc = open_documents.get(key)
    if doc is not None:
        try:
            _ = doc.Name  # raises if the document was closed externally
            return doc
        except Exception:
            del open_documents[key]

    if not os.path.exists(path):
        if create_if_missing:
            return create_document(save_path=path)[0]
        raise FileNotFoundError(f"File does not exist at path: {path}")
    doc = get_visio_app().Documents.Open(path)
    open_documents[key] = doc
    return doc


def get_page(doc):
    """Return the first page of doc, independent of which window is active."""
    return doc.Pages.Item(1)


def get_shape(page, shape_id: int):
    """Return the shape with the given ID on the page."""
    try:
        return page.Shapes.ItemFromID(shape_id)
    except Exception:
        raise LookupError(f"Could not find shape with ID {shape_id}")


# --- Stencils ----------------------------------------------------------------

def stencil_search_dirs() -> List[str]:
    """Directories searched for stencil files: Visio's own content, My Shapes and any custom paths."""
    app = get_visio_app()
    dirs = [os.path.join(app.Path, "Visio Content"), MY_SHAPES_DIR]
    dirs.extend(p for p in (app.StencilPaths or "").split(";") if p)
    return [d for d in dirs if os.path.isdir(d)]


def find_stencil_files() -> List[str]:
    """All stencil files found in the search directories."""
    files = []
    for directory in stencil_search_dirs():
        for root, _, names in os.walk(directory):
            files.extend(os.path.join(root, n) for n in names if n.lower().endswith(STENCIL_EXTENSIONS))
    return sorted(files)


def resolve_stencil(stencil: str) -> str:
    """Resolve a stencil file path or a name such as 'SERVER_U' / 'server_u.vssx'."""
    if os.path.isfile(stencil):
        return os.path.abspath(stencil)
    wanted = stencil.lower()
    files = find_stencil_files()
    exact = [f for f in files if os.path.basename(f).lower() == wanted
             or os.path.splitext(os.path.basename(f))[0].lower() == wanted]
    # Prefer the .vssx file when both formats exist
    exact.sort(key=lambda f: not f.lower().endswith(".vssx"))
    if exact:
        return exact[0]
    partial = [f for f in files if wanted in os.path.basename(f).lower()]
    if len(partial) == 1:
        return partial[0]
    if partial:
        names = ", ".join(os.path.basename(f) for f in partial[:10])
        raise ToolError(f"Stencil '{stencil}' is ambiguous. Matches: {names}")
    raise FileNotFoundError(f"Stencil not found: {stencil}. Use list_stencils to see what is available.")


def get_stencil(stencil: str):
    """Open a stencil read-only and hidden (cached) and return its Document."""
    path = resolve_stencil(stencil)
    key = os.path.normcase(path)
    doc = stencil_documents.get(key)
    if doc is not None:
        try:
            _ = doc.Name
            return doc
        except Exception:
            del stencil_documents[key]
    doc = get_visio_app().Documents.OpenEx(path, VIS_OPEN_RO | VIS_OPEN_HIDDEN)
    stencil_documents[key] = doc
    return doc


def find_master(stencil_doc, master: str):
    """Find a master by local or universal name, case-insensitively."""
    try:
        return stencil_doc.Masters.Item(master)
    except Exception:
        pass
    wanted = master.lower()
    for m in stencil_doc.Masters:
        if m.Name.lower() == wanted or m.NameU.lower() == wanted:
            return m
    raise KeyError(f"Master '{master}' not found in stencil '{stencil_doc.Name}'. Use list_stencil_masters.")


# --- Tools: documents --------------------------------------------------------

@tool("creating Visio file")
async def create_visio_file(template_path: Optional[str] = None, save_path: Optional[str] = None,
                            page_width: float = A4_LANDSCAPE[0], page_height: float = A4_LANDSCAPE[1]) -> str:
    """Create a new Visio file.

    Args:
        template_path: Path to the Visio template file (.vstx, .vst, etc.) to use.
                      If not provided, a blank document is created.
        save_path: Path where the file should be saved. If not provided,
                  it will be saved in the user's Documents folder with a default name.
        page_width: Page width in inches for a blank document (default: 11.69, A4 landscape).
        page_height: Page height in inches for a blank document (default: 8.27, A4 landscape).
                     The page is a custom size, so Visio does not reset it to the printer's paper.
                     Ignored when a template is used.

    Returns:
        The path to the created Visio file.
    """
    _, path = create_document(template_path, save_path, (page_width, page_height))
    return f"Visio file created successfully at: {path}"


@tool("opening Visio file")
async def open_visio_file(file_path: str) -> str:
    """Open an existing Visio file.

    Args:
        file_path: Path to the Visio file to open.

    Returns:
        Result message indicating success or failure.
    """
    path = resolve_path(file_path)
    already_open = doc_key(path) in open_documents
    get_document(path)
    if already_open:
        return f"Visio file is already open: {path}"
    return f"Visio file opened successfully: {path}"


@tool("closing document")
async def close_document(file_path: str, save_changes: Optional[bool] = True) -> str:
    """Close a Visio document.

    Args:
        file_path: Path to the Visio file.
        save_changes: Whether to save changes before closing (default: True).

    Returns:
        Result message indicating success or failure.
    """
    path = resolve_path(file_path)
    doc = open_documents.get(doc_key(file_path))
    if doc is None:
        return f"Document {path} is not currently open"
    if save_changes:
        doc.Save()
    doc.Close()
    del open_documents[doc_key(file_path)]
    return f"Document {path} closed successfully"


@tool("setting page size")
async def set_page_setup(file_path: str, width: float = A4_LANDSCAPE[0], height: float = A4_LANDSCAPE[1],
                         drawing_scale: Optional[float] = None) -> str:
    """Set the page size (and optionally the drawing scale) of a Visio document's first page.

    The page is set to a custom size and its orientation follows the width/height, so Visio
    no longer resizes it to the printer's paper. The default is A4 landscape.

    Args:
        file_path: Path to the Visio file.
        width: Page width in inches (default: 11.69, A4 landscape).
        height: Page height in inches (default: 8.27, A4 landscape).
        drawing_scale: Optional drawing scale: real inches represented by one page inch
                       (e.g. 12 for 1:12). Shapes added afterwards are placed in real inches.

    Returns:
        Result message indicating success or failure.
    """
    doc = get_document(file_path)
    apply_page_setup(doc, width, height, drawing_scale)
    doc.Save()
    scale_note = f", scale 1:{drawing_scale:g}" if drawing_scale is not None else ""
    return f"Page set to {width} x {height} in{scale_note}"


# --- Tools: shapes -----------------------------------------------------------

SHAPE_DRAWERS = {
    "rectangle": lambda page, x1, y1, x2, y2: page.DrawRectangle(x1, y1, x2, y2),
    "circle": lambda page, x1, y1, x2, y2: page.DrawOval(x1, y1, x2, y2),
    "ellipse": lambda page, x1, y1, x2, y2: page.DrawOval(x1, y1, x2, y2),
    "line": lambda page, x1, y1, x2, y2: page.DrawLine(x1, y1, x2, y2),
}


@tool("adding shape to Visio file")
async def add_shape(file_path: str, shape_type: str, x: float, y: float,
                    width: Optional[float] = 1.0, height: Optional[float] = 1.0,
                    text: Optional[str] = None,
                    fill_color: Optional[str] = None, line_color: Optional[str] = None,
                    line_weight: Optional[float] = None, line_pattern: Optional[str] = None,
                    text_color: Optional[str] = None, font_size: Optional[float] = None,
                    bold: Optional[bool] = None, rounding: Optional[float] = None) -> str:
    """Add a shape to a Visio document (the file is created if it does not exist).

    Args:
        file_path: Path to the Visio file.
        shape_type: Type of shape to add: "Rectangle", "Circle", "Ellipse" or "Line".
                    Anything else is drawn as a rectangle.
        x: X-coordinate for the shape.
        y: Y-coordinate for the shape.
        width: Width of the shape (default: 1.0).
        height: Height of the shape (default: 1.0).
        text: Label for the shape (default: the shape type).
        fill_color: Fill color as #RRGGBB or a name (red, blue, teal, ...).
        line_color: Outline color, same formats as fill_color.
        line_weight: Outline thickness in points.
        line_pattern: Outline style: solid, dash, dot or dashdot.
        text_color: Label color, same formats as fill_color.
        font_size: Label font size in points.
        bold: Whether the label is bold.
        rounding: Corner radius in inches (rectangles).

    Returns:
        Result message indicating success or failure.
    """
    validate_style(fill_color, line_color, line_pattern, text_color)

    doc = get_document(file_path, create_if_missing=True)
    page = get_page(doc)

    draw = SHAPE_DRAWERS.get(shape_type.lower(), SHAPE_DRAWERS["rectangle"])
    shape = draw(page, x, y, x + width, y + height)
    shape.Text = text if text is not None else shape_type
    apply_shape_style(shape, fill_color, line_color, line_weight, line_pattern,
                      text_color, font_size, bold, rounding)
    doc.Save()

    return f"Shape '{shape_type}' added to the Visio file at ({x}, {y}) with ID {shape.ID}"


@tool("connecting shapes")
async def connect_shapes(file_path: str, shape1_id: int, shape2_id: int,
                         connector_type: Optional[str] = "Dynamic",
                         line_color: Optional[str] = None, line_weight: Optional[float] = None,
                         line_pattern: Optional[str] = None,
                         begin_arrow: Optional[bool] = None, end_arrow: Optional[bool] = None,
                         label: Optional[str] = None,
                         text_color: Optional[str] = None, font_size: Optional[float] = None) -> str:
    """Connect two shapes in a Visio document.

    Args:
        file_path: Path to the Visio file.
        shape1_id: ID of the first shape.
        shape2_id: ID of the second shape.
        connector_type: Type of connector (options: "Dynamic", "Straight", "Curved").
        line_color: Line color as #RRGGBB or a name (red, blue, teal, ...).
        line_weight: Line thickness in points.
        line_pattern: Line style: solid, dash, dot or dashdot.
        begin_arrow: Draw an arrowhead at the first shape.
        end_arrow: Draw an arrowhead at the second shape.
        label: Text shown on the connector.
        text_color: Label color, same formats as line_color.
        font_size: Label font size in points.

    Returns:
        Result message indicating success or failure.
    """
    validate_style(None, line_color, line_pattern, text_color)
    app = get_visio_app()
    doc = get_document(file_path)
    page = get_page(doc)
    make_connector(app, page, get_shape(page, shape1_id), get_shape(page, shape2_id), connector_type,
                   line_color, line_weight, line_pattern, begin_arrow, end_arrow, label, text_color, font_size)
    doc.Save()

    return f"Shapes {shape1_id} and {shape2_id} connected successfully with {connector_type} connector"


@tool("adding text to shape")
async def add_text(file_path: str, shape_id: int, text: str) -> str:
    """Add text to a shape in a Visio document.

    Args:
        file_path: Path to the Visio file.
        shape_id: ID of the shape to add text to.
        text: Text to add to the shape.

    Returns:
        Result message indicating success or failure.
    """
    doc = get_document(file_path)
    get_shape(get_page(doc), shape_id).Text = text
    doc.Save()
    return f"Text added to shape {shape_id} successfully"


@tool("listing shapes")
async def list_shapes(file_path: str) -> str:
    """List all shapes in a Visio document.

    Args:
        file_path: Path to the Visio file.

    Returns:
        JSON string containing information about all shapes in the document.
    """
    page = get_page(get_document(file_path))
    shapes_info = [
        {
            "ID": shape.ID,
            "Name": shape.Name,
            "Text": shape.Text,
            "Type": shape.Type,
            "Position": {"X": shape.Cells("PinX").Result(""), "Y": shape.Cells("PinY").Result("")},
            "Size": {"Width": shape.Cells("Width").Result(""), "Height": shape.Cells("Height").Result("")},
        }
        for shape in page.Shapes
    ]
    return json.dumps(shapes_info, indent=2)


# --- Tools: stencils ---------------------------------------------------------

@tool("listing stencils")
async def list_stencils(query: Optional[str] = None) -> str:
    """List the Visio stencils available on this machine.

    Args:
        query: Optional case-insensitive text to filter by file name
               (e.g. "server", "network", "flowchart").

    Returns:
        JSON list of stencils with their name and path.
    """
    files = find_stencil_files()
    if query:
        files = [f for f in files if query.lower() in os.path.basename(f).lower()]
    return json.dumps([{"name": os.path.splitext(os.path.basename(f))[0], "path": f} for f in files], indent=2)


@tool("listing stencil masters")
async def list_stencil_masters(stencil: str) -> str:
    """List the shapes (masters) in a stencil.

    Args:
        stencil: Stencil name (e.g. "SERVER_U") or full path to a stencil file.

    Returns:
        JSON list of master names in the stencil.
    """
    stencil_doc = get_stencil(stencil)
    return json.dumps({"stencil": stencil_doc.Name, "masters": [m.Name for m in stencil_doc.Masters]}, indent=2)


@tool("adding stencil shape")
async def add_stencil_shape(file_path: str, stencil: str, master: str, x: float, y: float,
                            text: Optional[str] = None,
                            width: Optional[float] = None, height: Optional[float] = None,
                            fill_color: Optional[str] = None, line_color: Optional[str] = None,
                            line_weight: Optional[float] = None, line_pattern: Optional[str] = None,
                            text_color: Optional[str] = None, font_size: Optional[float] = None,
                            bold: Optional[bool] = None) -> str:
    """Add a shape from a stencil (e.g. a server, router or flowchart symbol) to a document.

    Args:
        file_path: Path to the Visio file (created if it does not exist).
        stencil: Stencil name (e.g. "SERVER_U") or full path to a stencil file.
        master: Name of the shape in the stencil (see list_stencil_masters).
        x: X-coordinate of the shape's pin (usually its centre or bottom centre), in drawing inches.
        y: Y-coordinate of the shape's pin, in drawing inches.
        text: Optional label for the shape.
        width: Optional width in inches (default: the stencil's own size).
        height: Optional height in inches (default: the stencil's own size).
        fill_color: Optional fill color as #RRGGBB or a name (red, blue, teal, ...).
                    Applied to every part of the shape, which flattens built-in shading.
        line_color: Optional outline color, same formats as fill_color.
        line_weight: Optional outline thickness in points.
        line_pattern: Optional outline style: solid, dash, dot or dashdot.
        text_color: Optional label color, same formats as fill_color.
        font_size: Optional label font size in points.
        bold: Optional, whether the label is bold.

    Returns:
        Result message including the new shape's ID.
    """
    validate_style(fill_color, line_color, line_pattern, text_color)

    doc = get_document(file_path, create_if_missing=True)
    page = get_page(doc)
    stencil_doc = get_stencil(stencil)
    master_obj = find_master(stencil_doc, master)

    shape = page.Drop(master_obj, x, y)
    if width is not None:
        shape.Cells("Width").FormulaForceU = f"{width} in"
    if height is not None:
        shape.Cells("Height").FormulaForceU = f"{height} in"
    if text is not None:
        shape.Text = text
    style_geometry(shape, fill_color, line_color, line_weight, line_pattern)
    apply_shape_style(shape, text_color=text_color, font_size=font_size, bold=bold)
    doc.Save()

    return f"Shape '{master_obj.Name}' from '{stencil_doc.Name}' added at ({x}, {y}) with ID {shape.ID}"


# --- Drawing helpers shared by the high-level tools ---------------------------

def drawing_factor(page) -> float:
    """Real (drawing) inches per page inch, e.g. 12 on a 1:12 page."""
    sheet = page.PageSheet
    return sheet.Cells("DrawingScale").ResultIU / sheet.Cells("PageScale").ResultIU


def page_size(page) -> tuple:
    """Page width and height in page inches."""
    sheet = page.PageSheet
    return sheet.Cells("PageWidth").ResultIU, sheet.Cells("PageHeight").ResultIU


def bbox(shape) -> tuple:
    """(left, bottom, right, top) of a shape in drawing inches."""
    left, bottom, right, top = shape.BoundingBox(0)[:4]
    return left, bottom, right, top


def move_by_bbox(shape, center_x: Optional[float] = None, bottom_y: Optional[float] = None,
                 center_y: Optional[float] = None) -> None:
    """Shift a shape so its visible bounding box lands on the given centre x / bottom y / centre y.

    Stencil shapes have different pin positions, so placing by pin is unreliable.
    """
    left, bottom, right, top = bbox(shape)
    dx = 0.0 if center_x is None else center_x - (left + right) / 2
    if bottom_y is not None:
        dy = bottom_y - bottom
    elif center_y is not None:
        dy = center_y - (bottom + top) / 2
    else:
        dy = 0.0
    if dx:
        shape.Cells("PinX").FormulaForceU = f"{shape.Cells('PinX').ResultIU + dx} in"
    if dy:
        shape.Cells("PinY").FormulaForceU = f"{shape.Cells('PinY').ResultIU + dy} in"


def center_on(shape, x: float, y: float) -> None:
    """Move a shape so the centre of its own frame (not its pin or visible ink) is at (x, y)."""
    cells = shape.Cells
    center_x = cells("PinX").ResultIU - cells("LocPinX").ResultIU + cells("Width").ResultIU / 2
    center_y = cells("PinY").ResultIU - cells("LocPinY").ResultIU + cells("Height").ResultIU / 2
    cells("PinX").FormulaForceU = f"{cells('PinX').ResultIU + x - center_x} in"
    cells("PinY").FormulaForceU = f"{cells('PinY').ResultIU + y - center_y} in"


def fit_to_size(shape, size: float) -> None:
    """Scale a shape, keeping its proportions, so its longest side equals size (drawing inches)."""
    width, height = shape.Cells("Width").ResultIU, shape.Cells("Height").ResultIU
    if width <= 0 or height <= 0:
        return
    factor = size / max(width, height)
    shape.Cells("Width").FormulaForceU = f"{width * factor} in"
    shape.Cells("Height").FormulaForceU = f"{height * factor} in"


def draw_text_box(page, x: float, y: float, width: float, height: float, text: str,
                  size: float = 10, bold: bool = False, color: str = "#1E293B", align: int = 0):
    """Draw a borderless, unfilled text box with its bottom-left corner at (x, y)."""
    box = page.DrawRectangle(x, y, x + width, y + height)
    box.Text = text
    box.Cells("FillPattern").FormulaForceU = "0"
    box.Cells("LinePattern").FormulaForceU = "0"
    box.Cells("Para.HorzAlign").FormulaForceU = str(align)
    apply_shape_style(box, text_color=color, font_size=size, bold=bold)
    return box


def set_user_cell(shape, name: str, value) -> None:
    """Store a value on the shape (User-defined cell) so it survives saving and reopening."""
    if not shape.CellExistsU(f"User.{name}", 0):
        shape.AddNamedRow(VIS_SECTION_USER, name, 0)
    formula = f'"{value}"' if isinstance(value, str) else repr(value)
    shape.CellsU(f"User.{name}").FormulaU = formula


def get_user_cell(shape, name: str, number: bool = False):
    """Read a value stored with set_user_cell, or None if the shape has none."""
    if not shape.CellExistsU(f"User.{name}", 0):
        return None
    cell = shape.CellsU(f"User.{name}")
    return cell.ResultIU if number else cell.ResultStr("")


def make_connector(app, page, shape1, shape2, connector_type: str = "Dynamic", line_color=None,
                   line_weight=None, line_pattern=None, begin_arrow=None, end_arrow=None, label=None,
                   text_color=None, font_size=None):
    """Draw a connector glued between two shapes and style it."""
    route = {"dynamic": None, "straight": LINE_ROUTE_STRAIGHT, "curved": LINE_ROUTE_CURVED}.get(
        connector_type.lower())
    connector = page.Drop(app.ConnectorToolDataObject, 0, 0)
    if route is not None:
        # Centre-to-centre routing keeps straight/curved lines from being re-routed at right angles
        connector.Cells("ShapeRouteStyle").FormulaU = str(ROUTE_CENTER_TO_CENTER)
        connector.Cells("ConLineRouteExt").FormulaU = str(route)
    connector.Cells("BeginX").GlueTo(shape1.Cells("PinX"))
    connector.Cells("EndX").GlueTo(shape2.Cells("PinX"))

    apply_shape_style(connector, line_color=line_color, line_weight=line_weight,
                      line_pattern=line_pattern, text_color=text_color, font_size=font_size)
    if begin_arrow is not None:
        connector.Cells("BeginArrow").FormulaU = str(ARROW_TRIANGLE if begin_arrow else ARROW_NONE)
    if end_arrow is not None:
        connector.Cells("EndArrow").FormulaU = str(ARROW_TRIANGLE if end_arrow else ARROW_NONE)
    if label is not None:
        connector.Text = label
    return connector


def drop_stencil_shape(page, stencil: str, master: str):
    """Drop a stencil master on the page at the origin and return the shape and its master."""
    stencil_doc = get_stencil(stencil)
    master_obj = find_master(stencil_doc, master)
    return page.Drop(master_obj, 0, 0), master_obj, stencil_doc


# --- Tools: whole diagrams ---------------------------------------------------

@tool("building diagram")
async def build_diagram(file_path: str, spec: dict, replace_existing: bool = False) -> str:
    """Draw a complete diagram from a declarative spec in one call (network, system or HCI diagrams).

    Nodes are arranged automatically in rows (tiers) so the diagram fits the page.
    Use stencil shapes (server, switch, firewall icons) or plain labelled boxes.

    Args:
        file_path: Path to the Visio file (created if it does not exist).
        spec: The diagram:
            title, subtitle   optional heading text
            page              optional {"width": 11.69, "height": 8.27} in inches (default A4 landscape)
            direction         "top-down" (tier 0 on top, default) or "left-right"
            icon_size         longest side of stencil icons in inches (default 1.0)
            nodes             list of {"id", "label", "tier" (0 = first row), "stencil" + "master"
                              (optional, from list_stencils / list_stencil_masters),
                              "fill_color", "text_color", "width", "height"}
            links             list of {"from", "to" (node ids), "label", "type" (straight|curved|dynamic),
                              "color", "pattern" (solid|dash|dot|dashdot), "weight", "arrow" (none|end|begin|both)}
            groups            list of {"label", "nodes": [ids], "color"} drawn as a dashed box behind members
        replace_existing: Remove everything already on the page first (default: add to it).

    Returns:
        JSON with the Visio shape ID of every node.
    """
    spec = normalize_spec(spec)
    doc = get_document(file_path, create_if_missing=True)
    if spec.get("page"):
        apply_page_setup(doc, float(spec["page"].get("width", A4_LANDSCAPE[0])),
                         float(spec["page"].get("height", A4_LANDSCAPE[1])))
    page = get_page(doc)
    app = get_visio_app()
    if replace_existing:
        for existing in list(page.Shapes):
            existing.Delete()

    k = drawing_factor(page)
    page_w, page_h = (v * k for v in page_size(page))
    margin = 0.8 * k
    heading = bool(spec.get("title") or spec.get("subtitle"))
    positions = tier_layout(spec["nodes"], page_w, page_h, spec["direction"], margin,
                            top_reserved=(1.0 * k if heading else 0.0))

    if spec.get("title"):
        draw_text_box(page, margin, page_h - margin - 0.1 * k, page_w - 2 * margin, 0.5 * k,
                      spec["title"], size=22, bold=True)
    if spec.get("subtitle"):
        draw_text_box(page, margin, page_h - margin - 0.55 * k, page_w - 2 * margin, 0.35 * k,
                      spec["subtitle"], size=11, color="#64748B")

    shapes = {}
    sizes = {}  # node id -> (width, height) in drawing inches, used to size group boxes
    for node in spec["nodes"]:
        x, y = positions[node["id"]]
        if node.get("stencil"):
            shape, _, _ = drop_stencil_shape(page, node["stencil"], node["master"])
            node_icon = float(node.get("icon_size", spec["icon_size"])) * k
            if node.get("width") or node.get("height"):
                if node.get("width"):
                    shape.Cells("Width").FormulaForceU = f"{float(node['width']) * k} in"
                if node.get("height"):
                    shape.Cells("Height").FormulaForceU = f"{float(node['height']) * k} in"
                sizes[node["id"]] = (float(node.get("width", node.get("icon_size", 1.0))) * k,
                                     float(node.get("height", node.get("icon_size", 1.0))) * k)
            else:
                fit_to_size(shape, node_icon)
                sizes[node["id"]] = (node_icon, node_icon)
            center_on(shape, x, y)
            shape.Text = node["label"]
            style_geometry(shape, node.get("fill_color"))
            apply_shape_style(shape, text_color=node.get("text_color"))
        else:
            w = float(node.get("width", 1.8 * spec["icon_size"])) * k
            h = float(node.get("height", 0.7 * spec["icon_size"])) * k
            shape = page.DrawRectangle(x - w / 2, y - h / 2, x + w / 2, y + h / 2)
            sizes[node["id"]] = (w, h)
            shape.Text = node["label"]
            apply_shape_style(shape, fill_color=node.get("fill_color", "#DBEAFE"), line_color="#2563EB",
                              line_weight=1.25, text_color=node.get("text_color", "#1E293B"),
                              font_size=11, bold=True, rounding=0.1 * k)
        shapes[node["id"]] = shape

    group_shapes = []
    for group in spec["groups"]:
        left, bottom, right, top = group_bounds(group["nodes"], positions, sizes, pad=0.15 * k,
                                                label_space=0.3 * k)
        color = group.get("color", "#3B82F6")
        box = page.DrawRectangle(left, bottom, right, top)
        apply_shape_style(box, fill_color="#F1F5F9", line_color=color, line_weight=1.25,
                          line_pattern="dash", rounding=0.15 * k)
        box.Cells("FillForegndTrans").FormulaForceU = "40%"
        group_shapes.append((box, group, left, top, right))
    for box, group, left, top, right in group_shapes:
        if group.get("label"):
            draw_text_box(page, left + 0.1 * k, top - 0.32 * k, right - left - 0.2 * k, 0.28 * k,
                          group["label"].upper(), size=9, bold=True, color=group.get("color", "#1D4ED8"))
        box.SendToBack()

    arrows = {"none": (False, False), "end": (False, True), "begin": (True, False), "both": (True, True)}
    for link in spec["links"]:
        begin, end = arrows[str(link.get("arrow", "none")).lower()]
        connector = make_connector(
            app, page, shapes[link["from"]], shapes[link["to"]], link.get("type", "straight"),
            line_color=link.get("color", "#475569"), line_weight=link.get("weight", 1.5),
            line_pattern=link.get("pattern"), begin_arrow=begin, end_arrow=end, label=link.get("label"),
            font_size=9 if link.get("label") else None)
        connector.SendToBack()
    for box, *_ in group_shapes:
        box.SendToBack()

    doc.Save()
    return json.dumps({"nodes": {nid: s.ID for nid, s in shapes.items()},
                       "links": len(spec["links"]), "groups": len(spec["groups"])}, indent=2)


# --- Tools: racks ------------------------------------------------------------

@tool("adding rack")
async def add_rack(file_path: str, x: float, y: float, name: Optional[str] = None,
                   stencil: str = "Dell-Racks", master: str = "4220 Rack Frame",
                   u_height: int = 42, u1_offset: float = 4.0,
                   drawing_scale: Optional[float] = 12) -> str:
    """Add a rack frame from a stencil so devices can be mounted by U position.

    Rack and server stencils from VisioCafe (Dell, HPE, ...) are drawn in real inches, so the page
    is set to a drawing scale (default 1:12) and x, y are in real inches: on A4 landscape at 1:12
    the page is about 140 x 99 inches. x = 30, y = 10 puts a 42U rack on the left of the page.

    Args:
        file_path: Path to the Visio file (created if it does not exist).
        x: Centre of the rack, in drawing inches.
        y: Bottom of the rack, in drawing inches.
        name: Optional rack name shown above the rack (e.g. "Rack A01").
        stencil: Stencil holding the rack frame (default: Dell-Racks).
        master: Rack frame shape (default: "4220 Rack Frame", a Dell 42U rack).
        u_height: Number of rack units (default: 42).
        u1_offset: Distance from the rack's bottom edge to the bottom of U1, in inches
                   (default 4.0, which fits the Dell 4220 frame).
        drawing_scale: Page scale, real inches per page inch (default 12). None keeps the page as is.

    Returns:
        The rack's shape ID, to pass to add_rack_device and list_rack.
    """
    if u_height < 1:
        raise ToolError("u_height must be at least 1")
    doc = get_document(file_path, create_if_missing=True)
    page = get_page(doc)
    if drawing_scale is not None:
        width, height = page_size(page)
        apply_page_setup(doc, width, height, drawing_scale)

    shape, master_obj, _ = drop_stencil_shape(page, stencil, master)
    move_by_bbox(shape, center_x=x, bottom_y=y)
    left, bottom, right, top = bbox(shape)
    set_user_cell(shape, "RackUHeight", int(u_height))
    set_user_cell(shape, "RackU1Y", bottom + u1_offset)
    set_user_cell(shape, "RackCenterX", (left + right) / 2)
    set_user_cell(shape, "RackRight", right)
    if name:
        set_user_cell(shape, "RackName", name)
        draw_text_box(page, left, top + 2.5, right - left, 3.5, name, size=14, bold=True, align=1)
    doc.Save()
    return json.dumps({"rack_id": shape.ID, "u_height": u_height, "rack": master_obj.Name,
                       "bounds": [round(left, 2), round(bottom, 2), round(right, 2), round(top, 2)]})


def rack_contents(page, rack_id: int) -> list:
    """The devices mounted in a rack: (first_u, size, label, shape_id), lowest first."""
    devices = []
    for shape in page.Shapes:
        if shape.CellExistsU("User.RackId", 0) and int(get_user_cell(shape, "RackId", number=True)) == rack_id:
            devices.append((int(get_user_cell(shape, "RackU0", number=True)),
                            int(get_user_cell(shape, "RackUSize", number=True)),
                            get_user_cell(shape, "RackLabel") or shape.Name, shape.ID))
    return sorted(devices)


def get_rack(page, rack_id: int):
    """Return the rack shape and its stored geometry, or raise if rack_id is not a rack."""
    rack = get_shape(page, rack_id)
    if get_user_cell(rack, "RackUHeight", number=True) is None:
        raise ToolError(f"Shape {rack_id} is not a rack created with add_rack")
    return rack, {
        "u_height": int(get_user_cell(rack, "RackUHeight", number=True)),
        "u1_y": get_user_cell(rack, "RackU1Y", number=True),
        "center_x": get_user_cell(rack, "RackCenterX", number=True),
        "right": get_user_cell(rack, "RackRight", number=True),
    }


@tool("adding rack device")
async def add_rack_device(file_path: str, rack_id: int, stencil: str, master: str, u_position: int,
                          u_size: int = 1, label: Optional[str] = None, callout: bool = True) -> str:
    """Mount a device in a rack at a U position, rejecting overlaps and positions outside the rack.

    Args:
        file_path: Path to the Visio file.
        rack_id: Shape ID returned by add_rack.
        stencil: Stencil holding the device (e.g. "HPE-ProLiant-DL", "Juniper EX Series").
        master: Device shape (e.g. "DL380 Gen10 SFF front", "EX4300-24T Front").
        u_position: Lowest rack unit the device occupies (U1 is the bottom).
        u_size: Height of the device in rack units (default 1; a DL380 is 2).
        label: Name shown in the callout on the right of the rack (default: the shape name).
        callout: Draw a leader line and a "label (U38-39)" text next to the device (default True).

    Returns:
        The device's shape ID and the U range it occupies.
    """
    doc = get_document(file_path)
    page = get_page(doc)
    rack, geometry = get_rack(page, rack_id)
    occupied = [(first, size, name) for first, size, name, _ in rack_contents(page, rack_id)]
    check_rack_fit(geometry["u_height"], u_position, u_size, occupied)

    shape, master_obj, _ = drop_stencil_shape(page, stencil, master)
    bottom_y = rack_device_bottom(geometry["u1_y"], u_position)
    move_by_bbox(shape, center_x=geometry["center_x"], bottom_y=bottom_y)
    name = label or master_obj.Name
    last_u = u_position + u_size - 1
    set_user_cell(shape, "RackId", int(rack_id))
    set_user_cell(shape, "RackU0", int(u_position))
    set_user_cell(shape, "RackUSize", int(u_size))
    set_user_cell(shape, "RackLabel", name)

    if callout:
        k = drawing_factor(page)
        y_mid = bottom_y + u_size * RACK_UNIT / 2
        line = page.DrawLine(geometry["center_x"] + 10, y_mid, geometry["right"] + 2 * k, y_mid)
        apply_shape_style(line, line_color="#94A3B8", line_weight=0.75)
        rng = f"U{u_position}" if u_size == 1 else f"U{u_position}-{last_u}"
        draw_text_box(page, geometry["right"] + 2.1 * k, y_mid - 0.14 * k, 6 * k, 0.28 * k,
                      f"{name}   ({rng})", size=10)
    doc.Save()
    return json.dumps({"shape_id": shape.ID, "u_start": u_position, "u_end": last_u})


@tool("listing rack")
async def list_rack(file_path: str, rack_id: int) -> str:
    """Show what is mounted in a rack and which units are free.

    Args:
        file_path: Path to the Visio file.
        rack_id: Shape ID returned by add_rack.

    Returns:
        JSON with the devices (lowest U first), the free U ranges and used/free counts.
    """
    page = get_page(get_document(file_path))
    _, geometry = get_rack(page, rack_id)
    contents = rack_contents(page, rack_id)
    free = free_ranges(geometry["u_height"], [(f, s, n) for f, s, n, _ in contents])
    used = sum(size for _, size, _, _ in contents)
    return json.dumps({
        "u_height": geometry["u_height"],
        "used_u": used,
        "free_u": geometry["u_height"] - used,
        "devices": [{"shape_id": sid, "label": name, "u_start": first, "u_end": first + size - 1}
                    for first, size, name, sid in contents],
        "free_ranges": [{"u_start": a, "u_end": b} for a, b in free],
    }, indent=2)


# --- Tools: export, cleanup, search ------------------------------------------

IMAGE_FORMATS = (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".svg")
VIS_FIXED_FORMAT_PDF = 1
VIS_DOC_EX_INTENT_PRINT = 1
VIS_PRINT_ALL = 0


@tool("exporting diagram")
async def export_diagram(file_path: str, output_path: str) -> str:
    """Export a diagram's first page as PNG, JPG, GIF, BMP, SVG or PDF (chosen by the file extension).

    Args:
        file_path: Path to the Visio file.
        output_path: Where to write the image or PDF, e.g. C:\\Diagrams\\network.png or network.pdf.
                     A bare file name is written to the Documents folder.

    Returns:
        The path of the exported file.
    """
    out = resolve_path(output_path)
    extension = os.path.splitext(out)[1].lower()
    if extension != ".pdf" and extension not in IMAGE_FORMATS:
        raise ToolError(f"Unsupported format '{extension}'. Use one of: .pdf, {', '.join(IMAGE_FORMATS)}")
    os.makedirs(os.path.dirname(out), exist_ok=True)

    doc = get_document(file_path)
    if extension == ".pdf":
        doc.ExportAsFixedFormat(VIS_FIXED_FORMAT_PDF, out, VIS_DOC_EX_INTENT_PRINT, VIS_PRINT_ALL)
    else:
        get_page(doc).Export(out)
    return f"Exported to: {out}"


@tool("deleting shape")
async def delete_shape(file_path: str, shape_id: int) -> str:
    """Delete a shape (or connector) from a Visio document.

    Args:
        file_path: Path to the Visio file.
        shape_id: ID of the shape to delete.

    Returns:
        Result message indicating success or failure.
    """
    doc = get_document(file_path)
    get_shape(get_page(doc), shape_id).Delete()
    doc.Save()
    return f"Shape {shape_id} deleted"


@tool("searching stencil shapes")
async def search_stencil_shapes(query: str, stencil: str, limit: int = 40) -> str:
    """Find shapes by name inside the stencils whose file name contains `stencil`.

    Example: query "DL380", stencil "HPE-ProLiant" finds every DL380 shape in the HPE ProLiant stencils.

    Args:
        query: Case-insensitive text to look for in shape names.
        stencil: Part of the stencil file name (e.g. "HPE-ProLiant", "Juniper EX", "Dell-Racks").
        limit: Maximum number of shapes to return (default 40).

    Returns:
        JSON list of {"stencil", "master"} matches.
    """
    wanted_stencil = stencil.lower()
    files = [f for f in find_stencil_files() if wanted_stencil in os.path.basename(f).lower()]
    if not files:
        raise ToolError(f"No stencil file name contains '{stencil}'. Use list_stencils to see what is available.")
    matches = []
    seen = set()
    for path in files[:10]:
        if os.path.splitext(os.path.basename(path))[0].lower() in seen:
            continue  # same stencil saved as both .vss and .vssx
        seen.add(os.path.splitext(os.path.basename(path))[0].lower())
        stencil_doc = get_stencil(path)
        for m in stencil_doc.Masters:
            if query.lower() in m.Name.lower():
                matches.append({"stencil": os.path.splitext(os.path.basename(path))[0], "master": m.Name})
                if len(matches) >= limit:
                    return json.dumps(matches, indent=2)
    return json.dumps(matches, indent=2)


# --- Entry point -------------------------------------------------------------

atexit.register(close_visio_app)


def main():
    """Entry point for the MCP server."""
    if not check_visio_installed():
        sys.stderr.write("Microsoft Visio is not installed. This MCP server requires Visio to function.\n")
        sys.exit(1)

    try:
        # Start (or attach to) Visio up front so the first request is fast
        get_visio_app()
        mcp.run(transport='stdio')
    except Exception as e:
        sys.stderr.write(f"Error initializing Visio MCP Server: {str(e)}\n")
        sys.exit(1)


if __name__ == "__main__":
    main()
