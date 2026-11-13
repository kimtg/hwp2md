"""
Table builder module for converting HWP/HWPX tables into GitHub Flavored Markdown (GFM)
or HTML tables with proper span handling and column alignment.
"""

from __future__ import annotations
import html
import unicodedata
from typing import List, Tuple, Optional


def get_east_asian_width(s: str) -> int:
    """Calculate the visual display width of a string considering East Asian wide characters."""
    width = 0
    for ch in s:
        # Fullwidth or Wide characters take 2 columns; others take 1
        ea = unicodedata.east_asian_width(ch)
        if ea in ("W", "F"):
            width += 2
        else:
            width += 1
    return width


def sanitize_cell_text(text: str) -> str:
    """
    Sanitize text for a markdown table cell:
    - Escapes pipe '|' characters
    - Converts newlines to '<br>'
    - Trims leading/trailing whitespace
    """
    if not text:
        return ""
    # Normalize newlines
    normalized = text.replace("\r\n", "\n").replace("\r", "\n")
    # Split into lines and strip each line
    lines = [line.strip() for line in normalized.split("\n") if line.strip()]
    if not lines:
        return ""
    joined = "<br>".join(lines)
    # Escape markdown table pipe delimiter
    joined = joined.replace("|", "\\|")
    return joined


class TableGrid:
    """
    A 2-dimensional grid representing a table.
    Handles cell spanning (rowspan and colspan), cell padding, and markdown/HTML formatting.
    """

    def __init__(self, num_rows: int = 1, num_cols: int = 1):
        self.num_rows = max(1, num_rows)
        self.num_cols = max(1, num_cols)
        self.grid: List[List[str]] = [["" for _ in range(self.num_cols)] for _ in range(self.num_rows)]
        self.occupied: List[List[bool]] = [[False for _ in range(self.num_cols)] for _ in range(self.num_rows)]
        self.cells_info: List[Tuple[int, int, int, int, str]] = []  # (r, c, row_span, col_span, text)

    def _ensure_capacity(self, target_row: int, target_col: int):
        """Expand grid dimensions if a cell exceeds initial bounds."""
        if target_row >= self.num_rows:
            needed_rows = target_row + 1
            for _ in range(needed_rows - self.num_rows):
                self.grid.append(["" for _ in range(self.num_cols)])
                self.occupied.append([False for _ in range(self.num_cols)])
            self.num_rows = needed_rows

        if target_col >= self.num_cols:
            needed_cols = target_col + 1
            for r in range(self.num_rows):
                self.grid[r].extend(["" for _ in range(needed_cols - self.num_cols)])
                self.occupied[r].extend([False for _ in range(needed_cols - self.num_cols)])
            self.num_cols = needed_cols

    def set_cell(self, row: int, col: int, text: str, row_span: int = 1, col_span: int = 1):
        """
        Place a cell at (row, col) with specified row_span and col_span.
        Spanned positions are filled with empty string to maintain rectangular shape.
        """
        row_span = max(1, row_span)
        col_span = max(1, col_span)
        self._ensure_capacity(row + row_span - 1, col + col_span - 1)

        clean = sanitize_cell_text(text)
        self.grid[row][col] = clean
        self.occupied[row][col] = True
        self.cells_info = [item for item in self.cells_info if not (item[0] == row and item[1] == col)]
        self.cells_info.append((row, col, row_span, col_span, clean))

        # Mark all positions covered by this cell's span
        for dr in range(row_span):
            for dc in range(col_span):
                r = row + dr
                c = col + dc
                self.occupied[r][c] = True
                if dr != 0 or dc != 0:
                    self.grid[r][c] = ""

    def is_empty(self) -> bool:
        """Check if table has any non-empty text."""
        return not any(any(cell.strip() for cell in row) for row in self.grid)

    def to_markdown(self, box_style: str = "table", align_columns: bool = True) -> str:
        """
        Convert table to GitHub Flavored Markdown (GFM).
        
        box_style options for 1x1 tables:
            - 'table': Render as standard 1-cell markdown table.
            - 'quote': Render as a markdown blockquote (useful for callout/announcement boxes).
            - 'plain': Render as a regular markdown paragraph.
        """
        if self.is_empty():
            return ""

        # Handle 1x1 table special styles
        if self.num_rows == 1 and self.num_cols == 1:
            cell_text = self.grid[0][0]
            if box_style == "quote":
                # Convert <br> back to newlines prefixed with '>'
                raw = cell_text.replace("\\|", "|")
                lines = raw.split("<br>")
                return "\n".join(f"> {line.strip()}" for line in lines if line.strip())
            elif box_style == "plain":
                raw = cell_text.replace("\\|", "|")
                return raw.replace("<br>", "\n\n")

        # Calculate visual column widths
        col_widths = [3 for _ in range(self.num_cols)]
        for r in range(self.num_rows):
            for c in range(self.num_cols):
                w = get_east_asian_width(self.grid[r][c])
                if w > col_widths[c]:
                    col_widths[c] = w

        def pad_cell(text: str, width: int) -> str:
            if not align_columns:
                return text
            w = get_east_asian_width(text)
            pad = max(0, width - w)
            return text + (" " * pad)

        lines: List[str] = []

        # Header row (Row 0)
        header_cells = [pad_cell(self.grid[0][c], col_widths[c]) for c in range(self.num_cols)]
        lines.append("| " + " | ".join(header_cells) + " |")

        # Delimiter row
        sep_cells = ["-" * col_widths[c] for c in range(self.num_cols)]
        lines.append("| " + " | ".join(sep_cells) + " |")

        # Data rows (Row 1..num_rows-1)
        for r in range(1, self.num_rows):
            row_cells = [pad_cell(self.grid[r][c], col_widths[c]) for c in range(self.num_cols)]
            lines.append("| " + " | ".join(row_cells) + " |")

        return "\n".join(lines)

    def to_html(self) -> str:
        """
        Convert table to an HTML table with exact rowspan and colspan preserved.
        Supported by all major Markdown viewers.
        """
        if self.is_empty():
            return ""

        # Build grid of cell objects
        # We only output <td> at the starting (row, col) of a span
        cell_by_pos = {}
        for r, c, rs, cs, text in self.cells_info:
            cell_by_pos[(r, c)] = (rs, cs, text)

        lines = ["<table>"]
        for r in range(self.num_rows):
            lines.append("  <tr>")
            for c in range(self.num_cols):
                if (r, c) in cell_by_pos:
                    rs, cs, text = cell_by_pos[(r, c)]
                    attrs = []
                    if rs > 1:
                        attrs.append(f'rowspan="{rs}"')
                    if cs > 1:
                        attrs.append(f'colspan="{cs}"')
                    attr_str = (" " + " ".join(attrs)) if attrs else ""
                    # Unescape pipe and HTML-escape cell content while preserving <br> line breaks
                    raw_content = text.replace("\\|", "|")
                    html_content = "<br>".join(
                        html.escape(part, quote=False) for part in raw_content.split("<br>")
                    )
                    lines.append(f"    <td{attr_str}>{html_content}</td>")
            lines.append("  </tr>")
        lines.append("</table>")
        return "\n".join(lines)

    def to_text(self) -> str:
        """
        Convert table to plain text with tab-separated columns and newline-separated rows.
        Ideal for LLM processing, RAG indexing, and search.
        """
        if self.is_empty():
            return ""

        rows: List[str] = []
        for r in range(self.num_rows):
            cells = []
            for c in range(self.num_cols):
                raw = self.grid[r][c].replace("<br>", " ").replace("\\|", "|").strip()
                cells.append(raw)
            if any(c for c in cells):
                rows.append("\t".join(cells))
        return "\n".join(rows)
