"""Font lookup and text measuring for the SVG importer (Windows fonts, no Visio needed).

* font_candidates(): turn a CSS font-family list into concrete font names, expanding generic
  families (serif, sans-serif, monospace) and common web fonts (Helvetica, system-ui) to Windows fonts.
* text_width(): width of a string in px using the real font file through PIL, or an estimate when the
  font cannot be found.
"""
import functools
import os
import re
import sys
from typing import Dict, List, Optional

ALIASES = {
    "sans-serif": ["Arial", "Segoe UI"],
    "serif": ["Times New Roman", "Georgia"],
    "monospace": ["Consolas", "Courier New"],
    "cursive": ["Comic Sans MS"],
    "fantasy": ["Impact"],
    "helvetica": ["Arial"],
    "helvetica neue": ["Arial"],
    "times": ["Times New Roman"],
    "courier": ["Courier New"],
    "system-ui": ["Segoe UI"],
    "ui-sans-serif": ["Segoe UI"],
    "-apple-system": ["Segoe UI"],
    "blinkmacsystemfont": ["Segoe UI"],
    "roboto": ["Segoe UI", "Arial"],
    "open sans": ["Segoe UI", "Arial"],
    "inter": ["Segoe UI", "Arial"],
    "lucida grande": ["Segoe UI", "Arial"],
}
DEFAULT_FONTS = ["Arial"]
ESTIMATE_EM = 0.6  # average glyph width / font size when nothing can be measured


def font_candidates(family: str) -> List[str]:
    """Concrete font names to try, best first, for a CSS font-family value."""
    names: List[str] = []
    for part in (family or "").split(","):
        name = part.strip().strip("'\"")
        if not name:
            continue
        names.append(name)
        names.extend(ALIASES.get(name.lower(), []))
    return names + [n for n in DEFAULT_FONTS if n not in names]


@functools.lru_cache(maxsize=1)
def _font_files() -> Dict[str, str]:
    """Map lower-case font name ('arial', 'arial bold') to its file, from the Windows font registry."""
    files: Dict[str, str] = {}
    if sys.platform != "win32":
        return files
    import winreg
    font_dir = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
    for hive, sub in ((winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"),
                      (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts")):
        try:
            with winreg.OpenKey(hive, sub) as key:
                i = 0
                while True:
                    try:
                        name, value, _ = winreg.EnumValue(key, i)
                    except OSError:
                        break
                    i += 1
                    name = re.sub(r"\s*\((TrueType|OpenType)\)\s*$", "", name).lower()
                    if str(value).lower().endswith((".ttf", ".otf")):
                        files.setdefault(name, value if os.path.isabs(value) else os.path.join(font_dir, value))
        except OSError:
            continue
    return files


def find_font_file(family: str, bold: bool = False) -> Optional[str]:
    """Path of the first installed font from a CSS font-family list (bold variant if available)."""
    files = _font_files()
    for name in font_candidates(family):
        key = name.lower()
        if bold and key + " bold" in files:
            return files[key + " bold"]
        if key in files:
            return files[key]
    return None


@functools.lru_cache(maxsize=256)
def _load(path: str, size: int):
    from PIL import ImageFont
    return ImageFont.truetype(path, size)


def text_width(text: str, family: str, size_px: float, bold: bool = False) -> float:
    """Width in px of one line of text; falls back to an estimate if the font cannot be loaded."""
    if not text:
        return 0.0
    path = find_font_file(family, bold)
    if path and size_px > 0:
        try:
            # measure at a large integer size and scale, so fractional font sizes stay accurate
            scale = 64.0 / size_px
            return _load(path, 64).getlength(text) / scale
        except Exception:
            pass
    return len(text) * size_px * ESTIMATE_EM * (1.08 if bold else 1.0)
