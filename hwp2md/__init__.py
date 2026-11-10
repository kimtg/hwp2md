"""
hwp2md - Standalone HWP / HWPX to Markdown, Microsoft Word (.docx), and Plain Text Converter
Without Hancom Office / MS Office dependency.
Supports tables, cell spans, headings, text styles, embedded images, and metadata inspection.
"""

from .converter import convert, convert_to_docx, convert_to_text, get_metadata, detect_format
from .hwpx_parser import convert_hwpx, convert_hwpx_to_docx, HWPXParser
from .hwp_parser import convert_hwp, convert_hwp_to_docx, HWPParser
from .table_builder import TableGrid, sanitize_cell_text, get_east_asian_width
from .docx_builder import DocxBuilder
from .text_builder import markdown_to_text, extract_metadata

__version__ = "1.2.0"

__all__ = [
    "convert",
    "convert_to_docx",
    "convert_to_text",
    "get_metadata",
    "convert_hwpx",
    "convert_hwpx_to_docx",
    "convert_hwp",
    "convert_hwp_to_docx",
    "detect_format",
    "HWPXParser",
    "HWPParser",
    "TableGrid",
    "DocxBuilder",
    "markdown_to_text",
    "extract_metadata",
    "sanitize_cell_text",
    "get_east_asian_width",
]
