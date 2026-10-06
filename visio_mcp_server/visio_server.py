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
    from visio_mcp_server.styles import (apply_shape_style, style_geometry, validate_style)
except ImportError:
    # Started as a plain script without the package on PYTHONPATH
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    from visio_mcp_server.styles import (apply_shape_style, style_geometry, validate_style)

mcp = FastMCP("visio-server")

# --- Constants ---------------------------------------------------------------

DEFAULT_SAVE_PATH = os.path.expandvars(r"%USERPROFILE%\Documents")
A4_LANDSCAPE = (11.69, 8.27)  # inches

STENCIL_EXTENSIONS = (".vssx", ".vss", ".vssm")
MY_SHAPES_DIR = os.path.join(DEFAULT_SAVE_PATH, "My Shapes")

# Visio enumeration values used below
VIS_OPEN_RO = 2
VIS_OPEN_HIDDEN = 64
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
    route = {"dynamic": None, "straight": LINE_ROUTE_STRAIGHT, "curved": LINE_ROUTE_CURVED}.get(
        connector_type.lower(), None)

    app = get_visio_app()
    doc = get_document(file_path)
    page = get_page(doc)
    shape1 = get_shape(page, shape1_id)
    shape2 = get_shape(page, shape2_id)

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
