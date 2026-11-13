"""
Text extraction and metadata module for hwp2md.
Provides clean plain-text extraction optimized for LLM / RAG pipelines,
and document metadata inspection.
"""

from __future__ import annotations
import datetime
import os
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, Any, Optional, Union

_ZIP_MAGIC = b"PK\x03\x04"
_OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"

_HWP_PROP_MAP = {
    2: "title",
    3: "subject",
    4: "creator",
    5: "keywords",
    6: "comments",
    8: "last_saved_by",
    12: "created_at",
    13: "modified_at",
    20: "date",
}


def _decode_ole_prop(value: Any) -> Optional[str]:
    """Decode an OLE property value (str, bytes, or datetime) to clean string."""
    if value is None:
        return None
    if isinstance(value, datetime.datetime):
        return value.isoformat()
    if isinstance(value, bytes):
        for enc in ("utf-16-le", "utf-8", "cp949"):
            try:
                decoded = value.decode(enc).rstrip("\x00").strip()
                if decoded:
                    return decoded
            except Exception:
                continue
        return None
    s = str(value).strip()
    return s if s else None

# Regex patterns for stripping markdown syntax
RE_IMAGE = re.compile(r"!\[(.*?)\]\(.*?\)")
RE_LINK = re.compile(r"\[(.*?)\]\((.*?)\)")
RE_BOLD = re.compile(r"\*\*(.*?)\*\*")
RE_ITALIC = re.compile(r"\*(.*?)\*")
RE_STRIKE = re.compile(r"~~(.*?)~~")
RE_HTML_TAGS = re.compile(r"<[^>]+>")
RE_HEADING = re.compile(r"^\s*#{1,6}\s+", re.MULTILINE)
RE_BLOCKQUOTE = re.compile(r"^\s*>\s*", re.MULTILINE)
RE_HR = re.compile(r"^\s*---\s*$", re.MULTILINE)


def markdown_to_text(md_content: str) -> str:
    """
    Convert Markdown content to clean plain text.
    - Converts images to their alt text
    - Converts hyperlinks to just text
    - Strips markdown formatting (bold, italic, strike, headings, quotes)
    - Normalizes table separators
    - Removes HTML tags (converting <br> to newline)
    """
    if not md_content:
        return ""

    text = md_content

    # Replace <br> tags with newlines before stripping other HTML tags
    text = re.sub(r"(?i)<br\s*/?>", "\n", text)
    text = RE_HTML_TAGS.sub("", text)

    # Convert images ![alt](url) -> alt
    text = RE_IMAGE.sub(r"\1", text)

    # Convert links [text](url) -> text
    text = RE_LINK.sub(r"\1", text)

    # Strip inline styles
    text = RE_BOLD.sub(r"\1", text)
    text = RE_ITALIC.sub(r"\1", text)
    text = RE_STRIKE.sub(r"\1", text)

    # Strip headings and blockquotes
    text = RE_HEADING.sub("", text)
    text = RE_BLOCKQUOTE.sub("", text)
    text = RE_HR.sub("", text)

    # Convert GFM tables to tab-delimited text
    clean_lines = []
    for line in text.splitlines():
        line_s = line.strip()
        # Delimiter row like | --- | --- |
        if re.match(r"^\|(?:\s*:?-+:?\s*\|)+$", line_s):
            continue
        if line_s.startswith("|") and line_s.endswith("|"):
            cells = [c.strip().replace(r"\|", "|") for c in line_s[1:-1].split("|")]
            clean_lines.append("\t".join(cells))
        else:
            clean_lines.append(line.replace(r"\|", "|"))

    # Normalize excessive newlines
    result = "\n".join(clean_lines)
    result = re.sub(r"\n{3,}", "\n\n", result).strip() + "\n"
    return result


def extract_metadata(file_path: Union[str, Path]) -> Dict[str, Any]:
    """
    Extract document metadata from an HWP or HWPX file.
    Returns dictionary with file info, format, size, and available document properties.
    """
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")

    stat = path.stat()
    ext = path.suffix.lower()
    fmt = ext.lstrip(".")
    if ext not in (".hwpx", ".hwp"):
        with open(path, "rb") as f:
            header = f.read(8)
        if header.startswith(_ZIP_MAGIC):
            fmt = "hwpx"
        elif header.startswith(_OLE_MAGIC):
            fmt = "hwp"

    meta: Dict[str, Any] = {
        "filename": path.name,
        "filesize_bytes": stat.st_size,
        "format": fmt,
    }

    if fmt == "hwpx":
        try:
            with zipfile.ZipFile(path, "r") as zf:
                # Count sections and images
                sections = [n for n in zf.namelist() if n.startswith("Contents/section") and n.endswith(".xml")]
                bindata = [n for n in zf.namelist() if n.startswith("BinData/")]
                meta["section_count"] = len(sections)
                meta["embedded_images_count"] = len(bindata)

                # Read Content.hpf metadata if available
                if "Contents/content.hpf" in zf.namelist():
                    try:
                        root = ET.fromstring(zf.read("Contents/content.hpf"))
                        title = root.find(".//{*}title")
                        creator = root.find(".//{*}creator")
                        date = root.find(".//{*}date")
                        if title is not None and title.text:
                            meta["title"] = title.text.strip()
                        if creator is not None and creator.text:
                            meta["creator"] = creator.text.strip()
                        if date is not None and date.text:
                            meta["date"] = date.text.strip()
                    except Exception:
                        pass
        except Exception as e:
            meta["error"] = str(e)
    elif fmt == "hwp":
        try:
            import olefile
            if olefile.isOleFile(str(path)):
                with olefile.OleFileIO(str(path)) as ole:
                    streams = ole.listdir()
                    meta["ole_streams_count"] = len(streams)
                    bindata_streams = [s for s in streams if s[0] == "BinData"]
                    meta["embedded_images_count"] = len(bindata_streams)
                    section_streams = [s for s in streams if s[0] == "BodyText" and len(s) > 1 and s[1].startswith("Section")]
                    if section_streams:
                        meta["section_count"] = len(section_streams)

                    # Read SummaryInformation if exists
                    if ole.exists("\x05HwpSummaryInformation"):
                        try:
                            props = ole.getproperties("\x05HwpSummaryInformation")
                            for prop_id, key in _HWP_PROP_MAP.items():
                                if prop_id in props:
                                    val = _decode_ole_prop(props[prop_id])
                                    if val:
                                        meta[key] = val
                        except Exception:
                            pass
        except Exception as e:
            meta["error"] = str(e)

    return meta
