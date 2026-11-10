"""
HWP 5.0 (Binary) parser and converter to Markdown.
Converts .hwp files using pyhwp and olefile (without Hancom Office or COM automation).
Extracts text, formatted runs, images, and converts tables into GFM or HTML tables.
"""

from __future__ import annotations
import io
import os
import re
import shutil
import tempfile
import xml.etree.ElementTree as ET
from contextlib import closing, redirect_stderr
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

from .table_builder import TableGrid

try:
    from hwp5.xmlmodel import Hwp5File
    from hwp5.hwp5html import HTMLTransform
    PYHWP_AVAILABLE = True
except ImportError:
    PYHWP_AVAILABLE = False


class HWPParser:
    """Parser for HWP (HWP 5.0 binary) files."""

    def __init__(
        self,
        hwp_path: str | Path,
        extract_images: bool = True,
        images_dir: Optional[str | Path] = None,
        table_format: str = "markdown",
        box_style: str = "table",
    ):
        if not PYHWP_AVAILABLE:
            raise ImportError(
                "pyhwp and olefile are required for .hwp conversion.\n"
                "Please install them via: pip install pyhwp olefile"
            )

        self.hwp_path = Path(hwp_path)
        self.extract_images = extract_images
        self.images_dir = Path(images_dir) if images_dir else None
        self.table_format = table_format
        self.box_style = box_style

        self.bold_classes: Set[str] = set()
        self.italic_classes: Set[str] = set()
        self.underline_classes: Set[str] = set()
        self.font_sizes: Dict[str, float] = {}

    def parse(self, output_md_path: Optional[str | Path] = None) -> str:
        """Parse HWP file and return Markdown string. Writes to file if output_md_path is provided."""
        if not self.hwp_path.exists():
            raise FileNotFoundError(f"HWP file not found: {self.hwp_path}")

        if self.extract_images and not self.images_dir:
            base_stem = Path(output_md_path).stem if output_md_path else self.hwp_path.stem
            base_parent = Path(output_md_path).parent if output_md_path else self.hwp_path.parent
            self.images_dir = base_parent / f"{base_stem}_images"

        temp_dir = tempfile.mkdtemp(prefix="hwp2md_")
        try:
            # Transform HWP to XHTML directory with pyhwp (suppressing noisy internal warnings)
            with redirect_stderr(io.StringIO()):
                with closing(Hwp5File(str(self.hwp_path))) as hwp5file:
                    transformer = HTMLTransform()
                    transformer.transform_hwp5_to_dir(hwp5file, temp_dir)

            # Copy bindata images if enabled
            bindata_dir = Path(temp_dir) / "bindata"
            if self.extract_images and bindata_dir.exists() and any(bindata_dir.iterdir()):
                assert self.images_dir is not None
                self.images_dir.mkdir(parents=True, exist_ok=True)
                for img_file in bindata_dir.iterdir():
                    shutil.copy2(img_file, self.images_dir / img_file.name)

            # Parse CSS styles
            css_file = Path(temp_dir) / "styles.css"
            if css_file.exists():
                self._parse_css(css_file.read_text(encoding="utf-8", errors="ignore"))

            # Parse XHTML content
            xhtml_file = Path(temp_dir) / "index.xhtml"
            if not xhtml_file.exists():
                raise RuntimeError(f"Failed to generate XHTML from {self.hwp_path}")

            root = ET.fromstring(xhtml_file.read_text(encoding="utf-8", errors="ignore"))
            body = root.find(".//{http://www.w3.org/1999/xhtml}body")
            if body is None:
                body = root

            md_blocks = self._parse_element(body)
            full_md = "\n\n".join([b for b in md_blocks if b.strip()])
            full_md = re.sub(r"\n{3,}", "\n\n", full_md).strip() + "\n"

            if output_md_path:
                out_file = Path(output_md_path)
                out_file.parent.mkdir(parents=True, exist_ok=True)
                out_file.write_text(full_md, encoding="utf-8")

            return full_md
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

    def _parse_css(self, css_text: str):
        """Extract bold, italic, underline, and font size properties from styles.css."""
        for match in re.finditer(r"(\.[a-zA-Z0-9_-]+)\s*\{([^}]+)\}", css_text):
            cls_name = match.group(1).lstrip(".")
            rules = match.group(2)
            if "font-weight: bold" in rules or "font-weight: 700" in rules:
                self.bold_classes.add(cls_name)
            if "font-style: italic" in rules:
                self.italic_classes.add(cls_name)
            if "text-decoration: underline" in rules:
                self.underline_classes.add(cls_name)
            fs_match = re.search(r"font-size:\s*([\d\.]+)pt", rules)
            if fs_match:
                try:
                    self.font_sizes[cls_name] = float(fs_match.group(1))
                except ValueError:
                    pass

    def _parse_element(self, elem: ET.Element) -> List[str]:
        """Recursively parse XHTML elements into Markdown blocks."""
        tag = elem.tag.split("}")[-1]

        if tag == "table":
            # Check for inner/descendant tables (layout container table)
            inner_tables = [t for t in elem.iter("{http://www.w3.org/1999/xhtml}table") if t != elem]
            if inner_tables:
                # Layout container: recursively process child elements
                blocks: List[str] = []
                for child in elem:
                    blocks.extend(self._parse_element(child))
                return blocks
            else:
                table_md = self._parse_table(elem)
                return [table_md] if table_md else []

        elif tag == "p":
            # In pyhwp XHTML, tables are wrapped inside <p><span class="TableControl"><table>
            tbl = elem.find(".//{http://www.w3.org/1999/xhtml}table")
            if tbl is not None:
                inner_tables = [t for t in tbl.iter("{http://www.w3.org/1999/xhtml}table") if t != tbl]
                if inner_tables:
                    blocks: List[str] = []
                    for child in tbl:
                        blocks.extend(self._parse_element(child))
                    return blocks
                else:
                    table_md = self._parse_table(tbl)
                    return [table_md] if table_md else []
            else:
                p_md = self._parse_paragraph(elem)
                return [p_md] if p_md else []

        elif tag in ("h1", "h2", "h3", "h4", "h5", "h6"):
            level = int(tag[1])
            text = " ".join("".join(elem.itertext()).split())
            if text:
                return [f"{'#' * level} {text}"]
            return []

        elif tag == "hr":
            return ["---"]

        else:
            blocks = []
            for child in elem:
                blocks.extend(self._parse_element(child))
            return blocks

    def _parse_paragraph_details(
        self, p_elem: ET.Element
    ) -> Tuple[List[Tuple[str, bool, bool, bool]], List[str], float]:
        """Extract inline runs (text, bold, italic, underline), images, and max font size from a <p> element."""
        runs: List[Tuple[str, bool, bool, bool]] = []
        max_font_size = 10.0
        images: List[str] = []

        def walk_inline(node: ET.Element):
            nonlocal max_font_size
            tag = node.tag.split("}")[-1]

            if tag == "img":
                src = node.attrib.get("src", "")
                alt = node.attrib.get("alt", os.path.basename(src))
                if src:
                    img_name = os.path.basename(src)
                    if self.extract_images and self.images_dir:
                        rel_path = f"{self.images_dir.name}/{img_name}"
                    else:
                        rel_path = src
                    images.append(f"![{alt}]({rel_path})")

            classes = node.attrib.get("class", "").split()
            is_bold = any(c in self.bold_classes for c in classes) or tag in ("b", "strong")
            is_italic = any(c in self.italic_classes for c in classes) or tag in ("i", "em")
            is_underline = any(c in self.underline_classes for c in classes) or tag in ("u", "ins")

            for c in classes:
                if c in self.font_sizes and self.font_sizes[c] > max_font_size:
                    max_font_size = self.font_sizes[c]

            if node.text:
                runs.append((node.text, is_bold, is_italic, is_underline))

            for child in node:
                walk_inline(child)
                if child.tail:
                    runs.append((child.tail, is_bold, is_italic, is_underline))

        walk_inline(p_elem)

        # Merge adjacent runs with identical styling
        merged: List[Tuple[str, bool, bool, bool]] = []
        for text, b, it, u in runs:
            if not text:
                continue
            if merged and merged[-1][1:] == (b, it, u):
                prev_text, prev_b, prev_it, prev_u = merged[-1]
                merged[-1] = (prev_text + text, prev_b, prev_it, prev_u)
            else:
                merged.append((text, b, it, u))

        return merged, images, max_font_size

    def _parse_paragraph(self, p_elem: ET.Element) -> str:
        """Parse XHTML <p> element, converting spans and inline elements to Markdown."""
        merged, images, max_font_size = self._parse_paragraph_details(p_elem)

        parts = []
        is_all_bold = True if merged else False
        for text, b, it, u in merged:
            if not b:
                is_all_bold = False
            t = text.replace("\r", "").replace("\n", " ")
            if not t.strip():
                parts.append(t)
                continue
            match = re.match(r"^(\s*)(.*?)(\s*)$", t, flags=re.DOTALL)
            if match:
                lead, core, trail = match.groups()
                if core:
                    if b:
                        core = f"**{core}**"
                    if it:
                        core = f"*{core}*"
                    if u:
                        core = f"<u>{core}</u>"
                    parts.append(f"{lead}{core}{trail}")
                    parts.append(f"{lead}{core}{trail}")

        full_text = " ".join("".join(parts).split())
        if not full_text and not images:
            return ""

        result_blocks = []
        if images:
            result_blocks.extend(images)
        if full_text:
            clean_text = full_text.replace("*", "").strip()
            if max_font_size >= 18.0 and len(clean_text) < 100:
                result_blocks.append(f"# {clean_text}")
            elif max_font_size >= 15.0 and len(clean_text) < 120:
                result_blocks.append(f"## {clean_text}")
            elif max_font_size >= 13.0 and is_all_bold and len(clean_text) < 120:
                result_blocks.append(f"### {clean_text}")
            else:
                result_blocks.append(full_text)

        return "\n\n".join(result_blocks)

    def _extract_table_grid(self, tbl_elem: ET.Element) -> Optional[TableGrid]:
        """Extract TableGrid from XHTML <table> element with rowspan and colspan."""
        rows = tbl_elem.findall("{http://www.w3.org/1999/xhtml}tr")
        if not rows:
            return None

        grid = TableGrid()
        row_spans: Dict[int, Tuple[int, str]] = {}

        for r_idx, tr in enumerate(rows):
            cells = (
                tr.findall("{http://www.w3.org/1999/xhtml}td")
                + tr.findall("{http://www.w3.org/1999/xhtml}th")
            )
            cell_iter = iter(cells)
            c_idx = 0

            while True:
                if c_idx in row_spans:
                    rem, val = row_spans[c_idx]
                    grid.set_cell(row=r_idx, col=c_idx, text="", row_span=1, col_span=1)
                    if rem <= 1:
                        del row_spans[c_idx]
                    else:
                        row_spans[c_idx] = (rem - 1, val)
                    c_idx += 1
                    continue

                try:
                    cell = next(cell_iter)
                except StopIteration:
                    break

                rs = int(cell.attrib.get("rowspan", 1))
                cs = int(cell.attrib.get("colspan", 1))

                cell_text = self._extract_cell_text(cell)
                grid.set_cell(row=r_idx, col=c_idx, text=cell_text, row_span=rs, col_span=cs)

                if rs > 1:
                    row_spans[c_idx] = (rs - 1, cell_text)
                c_idx += 1

                for _ in range(cs - 1):
                    if rs > 1:
                        row_spans[c_idx] = (rs - 1, cell_text)
                    c_idx += 1

        return grid

    def _parse_table(self, tbl_elem: ET.Element) -> str:
        """Parse XHTML <table> element into Markdown or HTML table with rowspan/colspan support."""
        grid = self._extract_table_grid(tbl_elem)
        if not grid or grid.is_empty():
            return ""

        if self.table_format == "html":
            return grid.to_html()
        else:
            return grid.to_markdown(box_style=self.box_style)

    def _extract_cell_text(self, cell_elem: ET.Element) -> str:
        """Extract multi-line cell text joined with <br>."""
        ps = cell_elem.findall(".//{http://www.w3.org/1999/xhtml}p")
        if not ps:
            ps = [cell_elem]

        lines = []
        for p in ps:
            img = p.find(".//{http://www.w3.org/1999/xhtml}img")
            img_md = ""
            if img is not None:
                src = img.attrib.get("src", "")
                alt = img.attrib.get("alt", os.path.basename(src))
                if src:
                    img_name = os.path.basename(src)
                    rel_path = f"{self.images_dir.name}/{img_name}" if self.images_dir else src
                    img_md = f"![{alt}]({rel_path}) "

            t = " ".join("".join(p.itertext()).split())
            if img_md or t:
                lines.append(f"{img_md}{t}".strip())

        return "<br>".join(lines)

    def parse_docx(self, output_docx_path: Optional[str | Path] = None) -> Path:
        """Parse HWP file and export as Microsoft Word (.docx)."""
        from .docx_builder import DocxBuilder

        if not self.hwp_path.exists():
            raise FileNotFoundError(f"HWP file not found: {self.hwp_path}")

        out_file = Path(output_docx_path) if output_docx_path else self.hwp_path.with_suffix(".docx")

        if self.extract_images and not self.images_dir:
            self.images_dir = out_file.parent / f"{out_file.stem}_images"

        builder = DocxBuilder()
        temp_dir = tempfile.mkdtemp(prefix="hwp2md_")

        try:
            with redirect_stderr(io.StringIO()):
                with closing(Hwp5File(str(self.hwp_path))) as hwp5file:
                    transformer = HTMLTransform()
                    transformer.transform_hwp5_to_dir(hwp5file, temp_dir)

            bindata_dir = Path(temp_dir) / "bindata"
            if self.extract_images and bindata_dir.exists() and any(bindata_dir.iterdir()):
                assert self.images_dir is not None
                self.images_dir.mkdir(parents=True, exist_ok=True)
                for img_file in bindata_dir.iterdir():
                    shutil.copy2(img_file, self.images_dir / img_file.name)

            css_file = Path(temp_dir) / "styles.css"
            if css_file.exists():
                self._parse_css(css_file.read_text(encoding="utf-8", errors="ignore"))

            xhtml_file = Path(temp_dir) / "index.xhtml"
            if not xhtml_file.exists():
                raise RuntimeError(f"Failed to generate XHTML from {self.hwp_path}")

            root = ET.fromstring(xhtml_file.read_text(encoding="utf-8", errors="ignore"))
            body = root.find(".//{http://www.w3.org/1999/xhtml}body")
            if body is None:
                body = root

            self._render_element_to_docx(body, builder)
            builder.save(out_file)
            return out_file
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

    def _render_element_to_docx(self, elem: ET.Element, builder):
        """Recursively render XHTML elements into DocxBuilder."""
        tag = elem.tag.split("}")[-1]

        if tag == "table":
            inner_tables = [t for t in elem.iter("{http://www.w3.org/1999/xhtml}table") if t != elem]
            if inner_tables:
                for child in elem:
                    self._render_element_to_docx(child, builder)
            else:
                grid = self._extract_table_grid(elem)
                if grid and not grid.is_empty():
                    builder.add_table_from_grid(grid, box_style=self.box_style, base_dir=self.images_dir)

        elif tag == "p":
            tbl = elem.find(".//{http://www.w3.org/1999/xhtml}table")
            if tbl is not None:
                inner_tables = [t for t in tbl.iter("{http://www.w3.org/1999/xhtml}table") if t != tbl]
                if inner_tables:
                    for child in tbl:
                        self._render_element_to_docx(child, builder)
                else:
                    grid = self._extract_table_grid(tbl)
                    if grid and not grid.is_empty():
                        builder.add_table_from_grid(grid, box_style=self.box_style, base_dir=self.images_dir)
            else:
                self._render_paragraph_to_docx(elem, builder)

        elif tag in ("h1", "h2", "h3", "h4", "h5", "h6"):
            level = int(tag[1])
            text = " ".join("".join(elem.itertext()).split())
            if text:
                builder.add_heading(text, level=min(level, 3))

        elif tag == "hr":
            pass

        else:
            for child in elem:
                self._render_element_to_docx(child, builder)

    def _render_paragraph_to_docx(self, p_elem: ET.Element, builder):
        """Render a single XHTML <p> paragraph into DocxBuilder."""
        merged, images, max_font_size = self._parse_paragraph_details(p_elem)

        # Add images
        for img_md in images:
            builder.add_markdown_paragraph(img_md, base_dir=self.images_dir)

        if not merged:
            return

        clean_text = " ".join("".join(t for t, _, _, _ in merged).split())
        if not clean_text:
            return

        is_all_bold = all(b for _, b, _, _ in merged)
        if max_font_size >= 18.0 and len(clean_text) < 100:
            builder.add_heading(clean_text, level=1)
        elif max_font_size >= 15.0 and len(clean_text) < 120:
            builder.add_heading(clean_text, level=2)
        elif max_font_size >= 13.0 and is_all_bold and len(clean_text) < 120:
            builder.add_heading(clean_text, level=3)
        else:
            runs_5 = [(t, b, it, False, u) for t, b, it, u in merged]
            builder.add_paragraph_runs(runs_5)


def convert_hwp(
    hwp_path: str | Path,
    output_path: Optional[str | Path] = None,
    extract_images: bool = True,
    images_dir: Optional[str | Path] = None,
    table_format: str = "markdown",
    box_style: str = "table",
) -> str:
    """Convenience function to convert an HWP file to Markdown."""
    parser = HWPParser(
        hwp_path=hwp_path,
        extract_images=extract_images,
        images_dir=images_dir,
        table_format=table_format,
        box_style=box_style,
    )
    return parser.parse(output_md_path=output_path)


def convert_hwp_to_docx(
    hwp_path: str | Path,
    output_path: Optional[str | Path] = None,
    extract_images: bool = True,
    images_dir: Optional[str | Path] = None,
    box_style: str = "table",
) -> Path:
    """Convenience function to convert an HWP file to Microsoft Word (.docx)."""
    parser = HWPParser(
        hwp_path=hwp_path,
        extract_images=extract_images,
        images_dir=images_dir,
        box_style=box_style,
    )
    return parser.parse_docx(output_docx_path=output_path)
