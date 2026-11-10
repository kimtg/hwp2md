"""
DocxBuilder module for converting parsed HWP/HWPX documents into Microsoft Word (.docx).
Maintains typography, font styles, headings, tables with rowspan/colspan, and images.
"""

from __future__ import annotations
import os
import re
from pathlib import Path
from typing import List, Optional, Tuple, Union

try:
    import docx
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
    from docx.oxml import OxmlElement, parse_xml
    from docx.oxml.ns import nsdecls, qn
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

from .table_builder import TableGrid


def parse_inline_markdown(text: str) -> List[Tuple[str, str, Optional[str]]]:
    """
    Parse inline markdown tokens from text:
    Returns list of (token_type, content, extra)
    token_type can be: 'text', 'bold', 'italic', 'strike', 'image'
    """
    pattern = re.compile(
        r'(!\[.*?\]\(.*?\))|'
        r'(\*\*.*?\*\*)|'
        r'(\*.*?\*)|'
        r'(~~.*?~~)'
    )
    tokens: List[Tuple[str, str, Optional[str]]] = []
    last_idx = 0

    for match in pattern.finditer(text):
        start, end = match.span()
        if start > last_idx:
            raw_text = text[last_idx:start]
            tokens.append(('text', raw_text.replace(r"\|", "|"), None))

        token_str = match.group(0)
        if token_str.startswith('!['):
            img_match = re.match(r'!\[(.*?)\]\((.*?)\)', token_str)
            if img_match:
                tokens.append(('image', img_match.group(1), img_match.group(2)))
        elif token_str.startswith('**'):
            inner = token_str[2:-2].replace(r"\|", "|")
            tokens.append(('bold', inner, None))
        elif token_str.startswith('*'):
            inner = token_str[1:-1].replace(r"\|", "|")
            tokens.append(('italic', inner, None))
        elif token_str.startswith('~~'):
            inner = token_str[2:-2].replace(r"\|", "|")
            tokens.append(('strike', inner, None))

        last_idx = end

    if last_idx < len(text):
        raw_text = text[last_idx:]
        tokens.append(('text', raw_text.replace(r"\|", "|"), None))

    return tokens


class DocxBuilder:
    """Builder for constructing Microsoft Word (.docx) documents."""

    def __init__(self, title: Optional[str] = None):
        if not DOCX_AVAILABLE:
            raise ImportError(
                "python-docx is required for docx conversion.\n"
                "Please install it via: pip install python-docx"
            )

        self.doc = docx.Document()
        self._setup_document_styles()

        if title:
            self.add_heading(title, level=1)

    def _setup_document_styles(self):
        """Configure default Korean font and typography styles."""
        # Set Normal style
        style = self.doc.styles["Normal"]
        style.font.name = "맑은 고딕"
        style.font.size = Pt(10)
        style.font.color.rgb = RGBColor(0x22, 0x22, 0x22)
        style.paragraph_format.line_spacing = 1.15
        style.paragraph_format.space_after = Pt(4)

        # Set East Asian font in rFonts xml
        rPr = style.element.get_or_add_rPr()
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = OxmlElement("w:rFonts")
            rPr.append(rFonts)
        rFonts.set(qn("w:eastAsia"), "맑은 고딕")

        # Set default page margins (0.75 in)
        for section in self.doc.sections:
            section.top_margin = Inches(0.75)
            section.bottom_margin = Inches(0.75)
            section.left_margin = Inches(0.75)
            section.right_margin = Inches(0.75)

    def add_heading(self, text: str, level: int = 1):
        """Add a heading with clean styling and appropriate hierarchy."""
        p = self.doc.add_paragraph()
        run = p.add_run(text.strip())
        run.bold = True
        run.font.name = "맑은 고딕"

        # Explicit East Asian font for heading
        rPr = run.element.get_or_add_rPr()
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = OxmlElement("w:rFonts")
            rPr.append(rFonts)
        rFonts.set(qn("w:eastAsia"), "맑은 고딕")

        if level == 1:
            run.font.size = Pt(18)
            run.font.color.rgb = RGBColor(0x11, 0x18, 0x27)
            p.paragraph_format.space_before = Pt(14)
            p.paragraph_format.space_after = Pt(6)
        elif level == 2:
            run.font.size = Pt(14)
            run.font.color.rgb = RGBColor(0x1F, 0x29, 0x37)
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after = Pt(4)
        else:
            run.font.size = Pt(12)
            run.font.color.rgb = RGBColor(0x37, 0x41, 0x51)
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(3)

    def add_paragraph_runs(
        self,
        runs: List[Tuple[str, bool, bool, bool]],
        is_heading: bool = False,
        heading_level: int = 1,
    ):
        """Add a paragraph constructed from raw (text, bold, italic, strikeout) runs."""
        if not runs:
            return

        if is_heading:
            full_text = "".join(t for t, _, _, _ in runs)
            self.add_heading(full_text, level=heading_level)
            return

        p = self.doc.add_paragraph()
        for text, bold, italic, strikeout in runs:
            if not text:
                continue
            run = p.add_run(text)
            run.bold = bold
            run.italic = italic
            if strikeout:
                run.font.strike = True
            run.font.name = "맑은 고딕"

    def add_markdown_paragraph(self, text: str, base_dir: Optional[Path] = None):
        """Add a paragraph parsed from markdown formatting (headings, quotes, bold, images)."""
        stripped = text.strip()
        if not stripped:
            return

        # Heading detection
        if stripped.startswith("### "):
            self.add_heading(stripped[4:], level=3)
            return
        elif stripped.startswith("## "):
            self.add_heading(stripped[3:], level=2)
            return
        elif stripped.startswith("# "):
            self.add_heading(stripped[2:], level=1)
            return

        # Blockquote detection
        if stripped.startswith(">"):
            lines = [l.lstrip("> ").strip() for l in stripped.splitlines()]
            p = self.doc.add_paragraph()
            p.paragraph_format.left_indent = Inches(0.25)
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(4)
            for i, line in enumerate(lines):
                if i > 0:
                    p = self.doc.add_paragraph()
                    p.paragraph_format.left_indent = Inches(0.25)
                self._render_line_to_paragraph(p, line, base_dir=base_dir)
            return

        # Regular paragraph
        p = self.doc.add_paragraph()
        self._render_line_to_paragraph(p, stripped, base_dir=base_dir)

    def _render_line_to_paragraph(
        self,
        paragraph,
        text_line: str,
        base_dir: Optional[Path] = None,
        in_table: bool = False,
    ):
        """Render tokens in a single line to a paragraph."""
        tokens = parse_inline_markdown(text_line)
        for token_type, content, extra in tokens:
            if token_type == "text":
                run = paragraph.add_run(content)
                run.font.name = "맑은 고딕"
            elif token_type == "bold":
                run = paragraph.add_run(content)
                run.bold = True
                run.font.name = "맑은 고딕"
            elif token_type == "italic":
                run = paragraph.add_run(content)
                run.italic = True
                run.font.name = "맑은 고딕"
            elif token_type == "strike":
                run = paragraph.add_run(content)
                run.font.strike = True
                run.font.name = "맑은 고딕"
            elif token_type == "image":
                alt_text = content
                img_path_str = extra or ""
                self._insert_image_to_paragraph(
                    paragraph,
                    img_path_str,
                    base_dir=base_dir,
                    max_width_inches=1.8 if in_table else 5.5,
                )

    def _insert_image_to_paragraph(
        self,
        paragraph,
        image_path_str: str,
        base_dir: Optional[Path] = None,
        max_width_inches: float = 5.5,
    ):
        """Insert an image file into a paragraph with automatic dimension scaling."""
        candidates = [
            Path(image_path_str),
        ]
        if base_dir:
            candidates.append(base_dir / image_path_str)
            candidates.append(base_dir / Path(image_path_str).name)

        found_path: Optional[Path] = None
        for cand in candidates:
            if cand.exists() and cand.is_file():
                found_path = cand
                break

        if not found_path:
            # Fallback to text label if image file not on disk
            run = paragraph.add_run(f"[{image_path_str}]")
            run.italic = True
            return

        try:
            run = paragraph.add_run()
            shape = run.add_picture(str(found_path))
            max_w = Inches(max_width_inches)
            if shape.width > max_w:
                ratio = shape.height / shape.width
                shape.width = max_w
                shape.height = int(max_w * ratio)
        except Exception:
            run = paragraph.add_run(f"[{found_path.name}]")
            run.italic = True

    def add_image(
        self,
        image_path: Union[str, Path],
        max_width_inches: float = 5.5,
    ):
        """Add a standalone centered image paragraph."""
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(6)
        self._insert_image_to_paragraph(p, str(image_path), max_width_inches=max_width_inches)

    def add_table_from_grid(
        self,
        grid: TableGrid,
        box_style: str = "table",
        base_dir: Optional[Path] = None,
    ):
        """
        Convert a TableGrid into a formatted Word Table,
        properly applying rowspan/colspan merging and cell formatting.
        """
        if grid.is_empty():
            return

        # Special handling for 1x1 callout boxes when box_style is 'quote' or 'plain'
        if grid.num_rows == 1 and grid.num_cols == 1:
            raw_text = grid.grid[0][0]
            if box_style == "plain":
                lines = raw_text.split("<br>")
                for line in lines:
                    self.add_markdown_paragraph(line.strip(), base_dir=base_dir)
                return
            elif box_style == "quote":
                lines = raw_text.split("<br>")
                quote_md = "\n".join(f"> {line.strip()}" for line in lines if line.strip())
                self.add_markdown_paragraph(quote_md, base_dir=base_dir)
                return

        # Create table
        table = self.doc.add_table(rows=grid.num_rows, cols=grid.num_cols)
        table.style = "Table Grid"
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        # Perform cell merges based on cells_info
        for r, c, row_span, col_span, _ in grid.cells_info:
            if row_span > 1 or col_span > 1:
                try:
                    target_r = min(r + row_span - 1, grid.num_rows - 1)
                    target_c = min(c + col_span - 1, grid.num_cols - 1)
                    cell_a = table.cell(r, c)
                    cell_b = table.cell(target_r, target_c)
                    if cell_a != cell_b:
                        cell_a.merge(cell_b)
                except Exception:
                    pass

        # Populate cell contents
        for r, c, _, _, text in grid.cells_info:
            try:
                cell = table.cell(r, c)
                self._populate_cell(cell, text, base_dir=base_dir)
            except Exception:
                pass

        # Add spacing after table
        spacer = self.doc.add_paragraph()
        spacer.paragraph_format.space_before = Pt(2)
        spacer.paragraph_format.space_after = Pt(4)

    def _populate_cell(self, cell, text: str, base_dir: Optional[Path] = None):
        """Populate a table cell with multiple lines and inline formatting."""
        if not text:
            cell.text = ""
            return

        lines = [line.strip() for line in text.split("<br>") if line.strip()]
        if not lines:
            cell.text = ""
            return

        p = cell.paragraphs[0]
        p.text = ""
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = 1.05

        for i, line in enumerate(lines):
            if i > 0:
                p = cell.add_paragraph()
                p.paragraph_format.space_before = Pt(1)
                p.paragraph_format.space_after = Pt(1)
                p.paragraph_format.line_spacing = 1.05
            self._render_line_to_paragraph(p, line, base_dir=base_dir, in_table=True)

    def save(self, output_path: Union[str, Path]):
        """Save the document to file."""
        out = Path(output_path)
        out.parent.mkdir(parents=True, exist_ok=True)
        self.doc.save(str(out))
