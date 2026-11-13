"""
DocxBuilder module for converting parsed HWP/HWPX documents into Microsoft Word (.docx).
Maintains typography, font styles, headings, tables with rowspan/colspan, images,
hyperlinks, and advanced callout boxes.
"""

from __future__ import annotations
import os
import re
from pathlib import Path
from typing import List, Optional, Tuple, Union
from xml.sax.saxutils import escape as xml_escape

try:
    import docx
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
    from docx.oxml import OxmlElement, parse_xml
    from docx.oxml.ns import nsdecls, qn
    import docx.opc.constants
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

from .table_builder import TableGrid, get_east_asian_width

# XML Namespaces
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def parse_inline_markdown(text: str) -> List[Tuple[str, str, Optional[str]]]:
    """
    Parse inline markdown tokens from text:
    Returns list of (token_type, content, extra)
    token_type can be: 'text', 'bold', 'italic', 'strike', 'underline', 'image', 'link'
    """
    pattern = re.compile(
        r'(!\[.*?\]\(.*?\))|'
        r'(\[.*?\]\(.*?\))|'
        r'(<u>.*?</u>)|'
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
        elif token_str.startswith('['):
            link_match = re.match(r'\[(.*?)\]\((.*?)\)', token_str)
            if link_match:
                tokens.append(('link', link_match.group(1), link_match.group(2)))
        elif token_str.startswith('<u>'):
            inner = token_str[3:-4].replace(r"\|", "|")
            tokens.append(('underline', inner, None))
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


def _apply_east_asian_font(run, font_name: str = "맑은 고딕"):
    """Set run font name and explicit East Asian font attribute in w:rPr."""
    run.font.name = font_name
    rPr = run.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:eastAsia"), font_name)


def add_hyperlink_to_paragraph(paragraph, url: str, text: str, color_hex: str = "1D4ED8"):
    """Insert a clickable external hyperlink into a python-docx paragraph."""
    part = paragraph.part
    try:
        r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
        safe_text = xml_escape(text)
        hyperlink_xml = (
            f'<w:hyperlink xmlns:w="{W_NS}" xmlns:r="{R_NS}" r:id="{r_id}">'
            f'  <w:r>'
            f'    <w:rPr>'
            f'      <w:rFonts w:eastAsia="맑은 고딕"/>'
            f'      <w:color w:val="{color_hex}"/>'
            f'      <w:u w:val="single"/>'
            f'    </w:rPr>'
            f'    <w:t xml:space="preserve">{safe_text}</w:t>'
            f'  </w:r>'
            f'</w:hyperlink>'
        )
        paragraph._p.append(parse_xml(hyperlink_xml))
    except Exception:
        # Fallback to plain formatted run
        run = paragraph.add_run(f"{text} ({url})")
        _apply_east_asian_font(run)
        run.underline = True


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
        _apply_east_asian_font(run)

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
        runs: List[Union[Tuple[str, bool, bool], Tuple[str, bool, bool, bool], Tuple[str, bool, bool, bool, bool]]],
        is_heading: bool = False,
        heading_level: int = 1,
    ):
        """
        Add a paragraph constructed from raw runs.
        Supports 3-tuple (text, bold, italic),
                 4-tuple (text, bold, italic, strikeout),
                 or 5-tuple (text, bold, italic, strikeout, underline).
        """
        if not runs:
            return

        if is_heading:
            full_text = "".join(item[0] for item in runs)
            self.add_heading(full_text, level=heading_level)
            return

        p = self.doc.add_paragraph()
        for item in runs:
            text = item[0]
            if not text:
                continue
            bold = item[1] if len(item) > 1 else False
            italic = item[2] if len(item) > 2 else False
            strikeout = item[3] if len(item) > 3 else False
            underline = item[4] if len(item) > 4 else False

            if "](" in text and "[" in text:
                last_pos = 0
                for m in re.finditer(r"\[([^\]]+)\]\(([^)]+)\)", text):
                    if m.start() > last_pos:
                        r = p.add_run(text[last_pos:m.start()])
                        r.bold = bold
                        r.italic = italic
                        if strikeout:
                            r.font.strike = True
                        if underline:
                            r.underline = True
                        _apply_east_asian_font(r)
                    add_hyperlink_to_paragraph(p, url=m.group(2), text=m.group(1))
                    last_pos = m.end()
                if last_pos < len(text):
                    r = p.add_run(text[last_pos:])
                    r.bold = bold
                    r.italic = italic
                    if strikeout:
                        r.font.strike = True
                    if underline:
                        r.underline = True
                    _apply_east_asian_font(r)
                continue

            run = p.add_run(text)
            run.bold = bold
            run.italic = italic
            if strikeout:
                run.font.strike = True
            if underline:
                run.underline = True
            _apply_east_asian_font(run)

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
                _apply_east_asian_font(run)
            elif token_type == "bold":
                run = paragraph.add_run(content)
                run.bold = True
                _apply_east_asian_font(run)
            elif token_type == "italic":
                run = paragraph.add_run(content)
                run.italic = True
                _apply_east_asian_font(run)
            elif token_type == "strike":
                run = paragraph.add_run(content)
                run.font.strike = True
                _apply_east_asian_font(run)
            elif token_type == "underline":
                run = paragraph.add_run(content)
                run.underline = True
                _apply_east_asian_font(run)
            elif token_type == "link":
                url = extra or ""
                add_hyperlink_to_paragraph(paragraph, url=url, text=content)
            elif token_type == "image":
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

    def _add_callout_box(self, text: str, base_dir: Optional[Path] = None):
        """Render a 1x1 callout box with light shaded background and an accent left border."""
        tbl = self.doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = tbl.cell(0, 0)
        cell.width = Inches(6.8)

        # Style borders and shading: Left thick border, subtle light gray background
        tcPr = cell._tc.get_or_add_tcPr()
        shd_xml = f'<w:shd xmlns:w="{W_NS}" w:val="clear" w:color="auto" w:fill="F8FAFC"/>'
        tcPr.append(parse_xml(shd_xml))

        borders_xml = (
            f'<w:tcBorders xmlns:w="{W_NS}">'
            f'  <w:top w:val="single" w:sz="4" w:space="0" w:color="E2E8F0"/>'
            f'  <w:left w:val="single" w:sz="24" w:space="0" w:color="2563EB"/>'
            f'  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="E2E8F0"/>'
            f'  <w:right w:val="single" w:sz="4" w:space="0" w:color="E2E8F0"/>'
            f'</w:tcBorders>'
        )
        tcPr.append(parse_xml(borders_xml))

        # Internal cell margins (padding)
        mar_xml = (
            f'<w:tcMar xmlns:w="{W_NS}">'
            f'  <w:top w:w="120" w:type="dxa"/>'
            f'  <w:left w:w="180" w:type="dxa"/>'
            f'  <w:bottom w:w="120" w:type="dxa"/>'
            f'  <w:right w:w="180" w:type="dxa"/>'
            f'</w:tcMar>'
        )
        tcPr.append(parse_xml(mar_xml))

        self._populate_cell(cell, text, base_dir=base_dir)

        spacer = self.doc.add_paragraph()
        spacer.paragraph_format.space_before = Pt(2)
        spacer.paragraph_format.space_after = Pt(4)

    def add_table_from_grid(
        self,
        grid: TableGrid,
        box_style: str = "table",
        base_dir: Optional[Path] = None,
    ):
        """
        Convert a TableGrid into a formatted Word Table,
        properly applying rowspan/colspan merging, intelligent column widths,
        cantSplit, repeat headers (tblHeader), and cell padding.
        """
        if grid.is_empty():
            return

        # Special handling for 1x1 callout boxes
        if grid.num_rows == 1 and grid.num_cols == 1:
            raw_text = grid.grid[0][0]
            if box_style == "plain":
                lines = raw_text.split("<br>")
                for line in lines:
                    self.add_markdown_paragraph(line.strip(), base_dir=base_dir)
                return
            elif box_style in ("quote", "callout", "table"):
                self._add_callout_box(raw_text, base_dir=base_dir)
                return

        # Create table
        table = self.doc.add_table(rows=grid.num_rows, cols=grid.num_cols)
        table.style = "Table Grid"
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False

        # Set table cell padding (margins) across the table
        tblPr = table._tbl.tblPr
        tblCellMar = (
            f'<w:tblCellMar xmlns:w="{W_NS}">'
            f'  <w:top w:w="80" w:type="dxa"/>'
            f'  <w:left w:w="120" w:type="dxa"/>'
            f'  <w:bottom w:w="80" w:type="dxa"/>'
            f'  <w:right w:w="120" w:type="dxa"/>'
            f'</w:tblCellMar>'
        )
        tblPr.append(parse_xml(tblCellMar))

        # Calculate proportional column widths
        total_avail_width_inches = 6.8  # Default page width minus margins
        col_text_weights = [1 for _ in range(grid.num_cols)]
        for r in range(grid.num_rows):
            for c in range(grid.num_cols):
                w = get_east_asian_width(grid.grid[r][c])
                if w > col_text_weights[c]:
                    col_text_weights[c] = min(w, 80)

        total_weight = sum(col_text_weights) or 1
        col_widths_inches = []
        for weight in col_text_weights:
            calc_w = (weight / total_weight) * total_avail_width_inches
            col_widths_inches.append(max(0.3, calc_w))

        # Normalize column widths to fit total_avail_width
        norm_sum = sum(col_widths_inches)
        if norm_sum > 0:
            col_widths_inches = [(w / norm_sum) * total_avail_width_inches for w in col_widths_inches]

        # Apply column widths to cells
        for r_idx, row in enumerate(table.rows):
            # cantSplit: prevent row breaking across pages
            trPr = row._tr.get_or_add_trPr()
            trPr.append(parse_xml(f'<w:cantSplit xmlns:w="{W_NS}"/>'))

            # tblHeader: repeat header row on subsequent pages
            if r_idx == 0 and grid.num_rows > 1:
                trPr.append(parse_xml(f'<w:tblHeader xmlns:w="{W_NS}"/>'))

            for c_idx, cell in enumerate(row.cells):
                if c_idx < len(col_widths_inches):
                    cell.width = Inches(col_widths_inches[c_idx])

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
