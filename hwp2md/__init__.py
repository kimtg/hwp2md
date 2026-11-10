"""
hwp2md - Standalone HWP / HWPX to Markdown and Microsoft Word (.docx) Converter
Without Hancom Office / Hangul application dependency.
Supports tables, cell spans, headings, text styles, and embedded images.
"""

from .converter import convert, convert_to_docx, detect_format
from .hwpx_parser import convert_hwpx, convert_hwpx_to_docx, HWPXParser
from .hwp_parser import convert_hwp, convert_hwp_to_docx, HWPParser
from .table_builder import TableGrid, sanitize_cell_text, get_east_asian_width
from .docx_builder import DocxBuilder

__version__ = "1.1.0"

__all__ = [
    "convert",
    "convert_to_docx",
    "convert_hwpx",
    "convert_hwpx_to_docx",
    "convert_hwp",
    "convert_hwp_to_docx",
    "detect_format",
    "HWPXParser",
    "HWPParser",
    "TableGrid",
    "DocxBuilder",
    "sanitize_cell_text",
    "get_east_asian_width",
]
