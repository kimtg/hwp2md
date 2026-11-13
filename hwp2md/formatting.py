"""
Shared inline formatting, run merging, and heading detection helpers
used by both HWP and HWPX parsers.
"""

from __future__ import annotations
import re
from typing import List, Optional, Sequence, Tuple

# (text, bold, italic, strikeout, underline)
FormattedRun = Tuple[str, bool, bool, bool, bool]

_RE_INLINE_MARKERS = re.compile(r"</?u>|\*\*|\*|~~")


def merge_runs(runs: Sequence[Tuple[str, ...]]) -> List[FormattedRun]:
    """Merge adjacent runs that share identical formatting into 5-tuples."""
    merged: List[FormattedRun] = []
    for item in runs:
        text = item[0]
        if not text:
            continue
        bold = bool(item[1]) if len(item) > 1 else False
        italic = bool(item[2]) if len(item) > 2 else False
        strikeout = bool(item[3]) if len(item) > 3 else False
        underline = bool(item[4]) if len(item) > 4 else False
        fmt = (bold, italic, strikeout, underline)

        if merged and merged[-1][1:] == fmt:
            merged[-1] = (merged[-1][0] + text, *fmt)
        else:
            merged.append((text, *fmt))
    return merged


def apply_inline_formatting(
    text: str,
    bold: bool = False,
    italic: bool = False,
    strikeout: bool = False,
    underline: bool = False,
) -> str:
    """Apply bold/italic/strikeout/underline markdown formatting preserving leading/trailing whitespace."""
    if not text or not (bold or italic or strikeout or underline):
        return text

    match = re.match(r"^(\s*)(.*?)(\s*)$", text, flags=re.DOTALL)
    if not match:
        return text
    leading, core, trailing = match.groups()
    if not core:
        return text

    if underline:
        core = f"<u>{core}</u>"
    if strikeout:
        core = f"~~{core}~~"
    if bold:
        core = f"**{core}**"
    if italic:
        core = f"*{core}*"

    return f"{leading}{core}{trailing}"


def strip_inline_markers(text: str) -> str:
    """Remove inline markdown/HTML formatting markers for clean heading text."""
    return _RE_INLINE_MARKERS.sub("", text).strip()


def detect_heading_level(
    clean_text: str,
    max_font_size: float,
    is_all_bold: bool,
) -> Optional[int]:
    """Determine heading level (1-3) from font size and boldness, or None for normal text."""
    if not clean_text or "\n" in clean_text:
        return None
    if max_font_size >= 18.0 and len(clean_text) < 100:
        return 1
    if max_font_size >= 15.0 and len(clean_text) < 120:
        return 2
    if max_font_size >= 13.0 and is_all_bold and len(clean_text) < 120:
        return 3
    return None


def format_paragraph_block(
    paragraph_text: str,
    max_font_size: float,
    is_all_bold: bool,
) -> str:
    """Format a paragraph as a Markdown heading if it matches heading criteria."""
    clean_text = strip_inline_markers(paragraph_text)
    level = detect_heading_level(clean_text, max_font_size, is_all_bold)
    if level is not None:
        return f"{'#' * level} {clean_text}"
    return paragraph_text
