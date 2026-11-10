"""
Unified converter module for HWP and HWPX to Markdown, Microsoft Word (.docx), and plain text.
Automatically identifies file format by extension and magic bytes.
"""

from __future__ import annotations
from pathlib import Path
from typing import Dict, Any, Optional, Union

from .hwpx_parser import convert_hwpx, convert_hwpx_to_docx, HWPXParser
from .hwp_parser import convert_hwp, convert_hwp_to_docx, HWPParser, PYHWP_AVAILABLE
from .text_builder import markdown_to_text, extract_metadata

# File magic bytes signatures
ZIP_MAGIC = b"PK\x03\x04"
OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def detect_format(file_path: Union[str, Path]) -> str:
    """
    Detect whether file is 'hwpx' or 'hwp' using extension and magic bytes.
    Returns 'hwpx', 'hwp', or raises ValueError if format is unrecognized.
    """
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")

    # Check extension first
    ext = path.suffix.lower()
    if ext == ".hwpx":
        return "hwpx"
    elif ext == ".hwp":
        return "hwp"

    # Inspect magic bytes if extension is missing or non-standard
    with open(path, "rb") as f:
        header = f.read(8)

    if header.startswith(ZIP_MAGIC):
        return "hwpx"
    elif header.startswith(OLE_MAGIC):
        return "hwp"

    raise ValueError(f"Unknown or unsupported document format for: {path}")


def convert_to_docx(
    input_path: Union[str, Path],
    output_path: Optional[Union[str, Path]] = None,
    extract_images: bool = True,
    images_dir: Optional[Union[str, Path]] = None,
    box_style: str = "table",
) -> Path:
    """
    Convert an HWP or HWPX file to Microsoft Word (.docx).

    Args:
        input_path: Path to the .hwp or .hwpx file.
        output_path: Optional path to save the .docx output file.
                     If None, saves next to input file with .docx extension.
        extract_images: Whether to extract embedded images. Default True.
        images_dir: Optional directory to store extracted images.
        box_style: Formatting for 1x1 callout boxes ('table', 'quote', 'plain').

    Returns:
        Path to the saved .docx file.
    """
    fmt = detect_format(input_path)

    if fmt == "hwpx":
        return convert_hwpx_to_docx(
            hwpx_path=input_path,
            output_path=output_path,
            extract_images=extract_images,
            images_dir=images_dir,
            box_style=box_style,
        )
    elif fmt == "hwp":
        return convert_hwp_to_docx(
            hwp_path=input_path,
            output_path=output_path,
            extract_images=extract_images,
            images_dir=images_dir,
            box_style=box_style,
        )
    else:
        raise ValueError(f"Unsupported format: {fmt}")


def convert_to_text(
    input_path: Union[str, Path],
    output_path: Optional[Union[str, Path]] = None,
) -> str:
    """
    Convert an HWP or HWPX file to clean plain text (optimized for LLM / RAG).

    Args:
        input_path: Path to the .hwp or .hwpx file.
        output_path: Optional path to save the .txt output file.

    Returns:
        Clean plain text string.
    """
    md = convert(input_path=input_path, extract_images=False)
    plain_text = markdown_to_text(md)

    if output_path:
        out_p = Path(output_path)
        out_p.parent.mkdir(parents=True, exist_ok=True)
        out_p.write_text(plain_text, encoding="utf-8")

    return plain_text


def get_metadata(input_path: Union[str, Path]) -> Dict[str, Any]:
    """Inspect and return document metadata."""
    return extract_metadata(input_path)


def convert(
    input_path: Union[str, Path],
    output_path: Optional[Union[str, Path]] = None,
    extract_images: bool = True,
    images_dir: Optional[Union[str, Path]] = None,
    table_format: str = "markdown",
    box_style: str = "table",
    to_docx: bool = False,
    to_text: bool = False,
) -> Union[str, Path]:
    """
    Convert an HWP or HWPX file to Markdown, Word (.docx), or plain text.

    Args:
        input_path: Path to the .hwp or .hwpx file.
        output_path: Optional path to save the output file (.md, .docx, or .txt).
        extract_images: Whether to extract embedded images to a directory. Default True.
        images_dir: Optional directory to save extracted images. Default '<output_stem>_images/'.
        table_format: 'markdown' for GitHub Flavored Markdown (GFM) tables,
                      or 'html' for HTML <table> tags with rowspan/colspan.
        box_style: How to format 1x1 single-cell boxes:
                   'table' (default Markdown table), 'quote' (blockquote >), or 'plain'.
        to_docx: If True or if output_path ends with '.docx', converts to .docx.
        to_text: If True or if output_path ends with '.txt', converts to clean plain text.

    Returns:
        Converted Markdown string, plain text string, or Path to docx file.
    """
    is_text = to_text or (output_path is not None and str(output_path).lower().endswith(".txt"))
    if is_text:
        return convert_to_text(input_path=input_path, output_path=output_path)

    is_docx = to_docx or (output_path is not None and str(output_path).lower().endswith(".docx"))
    if is_docx:
        return convert_to_docx(
            input_path=input_path,
            output_path=output_path,
            extract_images=extract_images,
            images_dir=images_dir,
            box_style=box_style,
        )

    fmt = detect_format(input_path)

    if fmt == "hwpx":
        return convert_hwpx(
            hwpx_path=input_path,
            output_path=output_path,
            extract_images=extract_images,
            images_dir=images_dir,
            table_format=table_format,
            box_style=box_style,
        )
    elif fmt == "hwp":
        return convert_hwp(
            hwp_path=input_path,
            output_path=output_path,
            extract_images=extract_images,
            images_dir=images_dir,
            table_format=table_format,
            box_style=box_style,
        )
    else:
        raise ValueError(f"Unsupported format: {fmt}")
