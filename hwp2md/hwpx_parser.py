"""
Standalone HWPX parser and converter to Markdown.
Parses HWPX (.hwpx) files directly using standard library zipfile and xml.etree.ElementTree.
No Hancom Office or Windows COM dependency required.
"""

from __future__ import annotations
import os
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any

from .formatting import (
    FormattedRun,
    apply_inline_formatting,
    detect_heading_level,
    format_paragraph_block,
    merge_runs,
)
from .table_builder import TableGrid

# XML Namespaces used in HWPX (KS X 6101 / OCF)
NS = {
    "hp": "http://www.hancom.co.kr/hwpml/2011/paragraph",
    "hs": "http://www.hancom.co.kr/hwpml/2011/section",
    "hc": "http://www.hancom.co.kr/hwpml/2011/core",
    "hh": "http://www.hancom.co.kr/hwpml/2011/head",
}


class HWPXParser:
    """Parser for HWPX files."""

    def __init__(
        self,
        hwpx_path: str | Path,
        extract_images: bool = True,
        images_dir: Optional[str | Path] = None,
        table_format: str = "markdown",
        box_style: str = "table",
    ):
        self.hwpx_path = Path(hwpx_path)
        self.extract_images = extract_images
        self.images_dir = Path(images_dir) if images_dir else None
        self.table_format = table_format
        self.box_style = box_style

        self.char_properties: Dict[str, Dict[str, Any]] = {}
        self.styles: Dict[str, Dict[str, Any]] = {}
        self.manifest: Dict[str, str] = {}
        self.section_hrefs: List[str] = []

    def parse(self, output_md_path: Optional[str | Path] = None) -> str:
        """Parse HWPX file and return Markdown string. Writes to file if output_md_path is provided."""
        if not self.hwpx_path.exists():
            raise FileNotFoundError(f"HWPX file not found: {self.hwpx_path}")

        # Set up default images directory next to output or input file
        if self.extract_images and not self.images_dir:
            base_stem = Path(output_md_path).stem if output_md_path else self.hwpx_path.stem
            base_parent = Path(output_md_path).parent if output_md_path else self.hwpx_path.parent
            self.images_dir = base_parent / f"{base_stem}_images"

        with zipfile.ZipFile(self.hwpx_path, "r") as zf:
            self._parse_manifest(zf)
            self._parse_header(zf)
            markdown_sections = self._parse_sections(zf)

        full_md = "\n\n".join([s for s in markdown_sections if s.strip()])
        # Normalize redundant newlines
        full_md = re.sub(r"\n{3,}", "\n\n", full_md).strip() + "\n"

        if output_md_path:
            out_file = Path(output_md_path)
            out_file.parent.mkdir(parents=True, exist_ok=True)
            out_file.write_text(full_md, encoding="utf-8")

        return full_md

    def _parse_manifest(self, zf: zipfile.ZipFile):
        """Read Contents/content.hpf to find section order and binary items."""
        hpf_path = "Contents/content.hpf"
        if hpf_path not in zf.namelist():
            # Fallback: scan zip entries for section*.xml
            self.section_hrefs = sorted(
                [name for name in zf.namelist() if name.startswith("Contents/section") and name.endswith(".xml")]
            )
            return

        try:
            root = ET.fromstring(zf.read(hpf_path))
            for item in root.findall(".//{*}item"):
                item_id = item.attrib.get("id", "")
                href = item.attrib.get("href", "")
                if item_id and href:
                    if not href.startswith("Contents/") and not href.startswith("BinData/"):
                        if "Contents/" + href in zf.namelist():
                            href = "Contents/" + href
                    self.manifest[item_id] = href

            spine_items = []
            for itemref in root.findall(".//{*}itemref"):
                idref = itemref.attrib.get("idref", "")
                if idref in self.manifest:
                    spine_items.append(self.manifest[idref])

            if spine_items:
                self.section_hrefs = spine_items
            else:
                self.section_hrefs = sorted(
                    [name for name in zf.namelist() if name.startswith("Contents/section") and name.endswith(".xml")]
                )
        except Exception:
            self.section_hrefs = sorted(
                [name for name in zf.namelist() if name.startswith("Contents/section") and name.endswith(".xml")]
            )

    def _parse_header(self, zf: zipfile.ZipFile):
        """Parse character properties and styles from Contents/header.xml."""
        header_path = "Contents/header.xml"
        if header_path not in zf.namelist():
            return

        try:
            root = ET.fromstring(zf.read(header_path))
            for cp in root.findall(".//{*}charPr"):
                cp_id = cp.attrib.get("id", "")
                height = int(cp.attrib.get("height", 1000)) / 100.0  # e.g. 1000 -> 10.0pt

                # Bold & Italic are present as child tags in HWPX
                bold = any(c.tag.endswith("bold") for c in cp)
                italic = any(c.tag.endswith("italic") for c in cp)

                # Underline and Strikeout are property tags with type / shape attributes
                ul = cp.find(".//{*}underline")
                underline = (ul is not None and ul.attrib.get("type", "NONE") != "NONE")

                so = cp.find(".//{*}strikeout")
                # Strikeout only when explicitly SOLID or DOUBLE
                strikeout = (so is not None and so.attrib.get("shape", "NONE") in ("SOLID", "DOUBLE"))

                self.char_properties[cp_id] = {
                    "height": height,
                    "bold": bold,
                    "italic": italic,
                    "underline": underline,
                    "strikeout": strikeout,
                }

            for st in root.findall(".//{*}style"):
                st_id = st.attrib.get("id", "")
                name = st.attrib.get("name", "")
                eng_name = st.attrib.get("engName", "")
                self.styles[st_id] = {
                    "name": name,
                    "engName": eng_name,
                    "charPrIDRef": st.attrib.get("charPrIDRef", ""),
                }
        except Exception:
            pass

    def _parse_sections(self, zf: zipfile.ZipFile) -> List[str]:
        """Parse all section XMLs and convert paragraphs and tables to markdown blocks."""
        blocks: List[str] = []
        for sec_href in self.section_hrefs:
            if sec_href not in zf.namelist():
                continue
            try:
                xml_data = zf.read(sec_href)
                root = ET.fromstring(xml_data)
                for p in root.findall("hp:p", NS):
                    block = self._parse_paragraph(p, zf)
                    if block and block.strip():
                        blocks.append(block)
            except Exception as e:
                blocks.append(f"<!-- Error reading section {sec_href}: {e} -->")
        return blocks

    def _extract_runs_and_images(
        self, p_elem: ET.Element, zf: zipfile.ZipFile, for_docx: bool = False
    ) -> Tuple[List[FormattedRun], List[Any], float]:
        """Extract formatted runs (with hyperlink support), images, and max font size from <hp:p>."""
        raw_runs: List[FormattedRun] = []
        max_font_size = 10.0
        images: List[Any] = []

        active_link_url: Optional[str] = None
        link_text_parts: List[str] = []
        link_fmt: Optional[Tuple[bool, bool, bool, bool]] = None

        def flush_link():
            nonlocal active_link_url, link_text_parts, link_fmt
            if active_link_url and link_text_parts:
                link_text = "".join(link_text_parts).strip()
                if link_text:
                    fmt = link_fmt or (False, False, False, False)
                    raw_runs.append((f"[{link_text}]({active_link_url})", *fmt))
            active_link_url = None
            link_text_parts = []
            link_fmt = None

        for run in p_elem.findall("hp:run", NS):
            # Check for hyperlink fieldBegin / fieldEnd inside <hp:ctrl>
            for ctrl in run.findall("hp:ctrl", NS):
                fb = ctrl.find("hp:fieldBegin", NS)
                if fb is not None and fb.attrib.get("type") == "HYPERLINK":
                    flush_link()
                    url = ""
                    for sp in fb.findall(".//hp:stringParam", NS):
                        name = sp.attrib.get("name", "")
                        val = (sp.text or "").strip()
                        if name == "Path" and val:
                            url = val
                            break
                        if name == "Command" and val and not url:
                            url = val.split(";")[0].strip()
                    if url:
                        active_link_url = url
                fe = ctrl.find("hp:fieldEnd", NS)
                if fe is not None and active_link_url is not None:
                    flush_link()

            # Check for embedded picture
            pic = run.find("hp:pic", NS)
            if pic is not None:
                if for_docx:
                    img_path, _ = self._extract_picture_file(pic, zf)
                    if img_path:
                        images.append(img_path)
                else:
                    img_md = self._parse_picture(pic, zf)
                    if img_md:
                        images.append(img_md)

            char_pr_id = run.attrib.get("charPrIDRef", "")
            cp = self.char_properties.get(char_pr_id, {})
            font_size = cp.get("height", 10.0)
            if font_size > max_font_size:
                max_font_size = font_size

            run_text = ""
            for t in run.findall("hp:t", NS):
                t_str = self._extract_text_from_t(t)
                if t_str:
                    run_text += t_str

            if run_text:
                bold = cp.get("bold", False)
                italic = cp.get("italic", False)
                strikeout = cp.get("strikeout", False)
                underline = cp.get("underline", False)
                if active_link_url is not None:
                    link_text_parts.append(run_text)
                    if link_fmt is None:
                        link_fmt = (bold, italic, strikeout, underline)
                else:
                    raw_runs.append((run_text, bold, italic, strikeout, underline))

        flush_link()
        return raw_runs, images, max_font_size

    def _parse_paragraph(self, p_elem: ET.Element, zf: zipfile.ZipFile) -> str:
        """Parse a single paragraph element, which may contain text, images, or tables."""
        tbl = p_elem.find(".//hp:tbl", NS)
        if tbl is not None:
            return self._parse_table(tbl, zf)

        raw_runs, images, max_font_size = self._extract_runs_and_images(p_elem, zf, for_docx=False)
        merged_runs = merge_runs(raw_runs)

        formatted_parts: List[str] = []
        is_all_bold = bool(merged_runs)

        for text, bold, italic, strikeout, underline in merged_runs:
            if not bold:
                is_all_bold = False
            part = apply_inline_formatting(
                text, bold=bold, italic=italic, strikeout=strikeout, underline=underline
            )
            formatted_parts.append(part)

        paragraph_text = "".join(formatted_parts).strip()

        result_parts = []
        if images:
            result_parts.extend(images)
        if paragraph_text:
            result_parts.append(format_paragraph_block(paragraph_text, max_font_size, is_all_bold))

        return "\n\n".join(result_parts)

    def _extract_text_from_t(self, t_elem: ET.Element) -> str:
        """Extract text from <hp:t>, properly handling lineBreak and fwSpace."""
        parts = []
        if t_elem.text:
            parts.append(t_elem.text)
        for child in t_elem:
            tag = child.tag.split("}")[-1]
            if tag == "lineBreak":
                parts.append("\n")
            elif tag == "fwSpace":
                parts.append(" ")
            if child.text:
                parts.append(child.text)
            if child.tail:
                parts.append(child.tail)
        return "".join(parts)

    def _extract_table_grid(self, tbl_elem: ET.Element, zf: zipfile.ZipFile) -> Optional[TableGrid]:
        """Extract TableGrid from <hp:tbl> element. Returns None if nested table container."""
        descendant_tbls = [t for t in tbl_elem.findall(".//hp:tbl", NS) if t != tbl_elem]
        if descendant_tbls:
            return None

        row_cnt = int(tbl_elem.attrib.get("rowCnt", 1))
        col_cnt = int(tbl_elem.attrib.get("colCnt", 1))
        grid = TableGrid(num_rows=row_cnt, num_cols=col_cnt)

        for tr in tbl_elem.findall("hp:tr", NS):
            for tc in tr.findall("hp:tc", NS):
                addr = tc.find("hp:cellAddr", NS)
                span = tc.find("hp:cellSpan", NS)
                r = int(addr.attrib.get("rowAddr", 0)) if addr is not None else 0
                c = int(addr.attrib.get("colAddr", 0)) if addr is not None else 0
                rs = int(span.attrib.get("rowSpan", 1)) if span is not None else 1
                cs = int(span.attrib.get("colSpan", 1)) if span is not None else 1

                cell_text = self._parse_table_cell(tc, zf)
                grid.set_cell(row=r, col=c, text=cell_text, row_span=rs, col_span=cs)

        return grid

    def _parse_table(self, tbl_elem: ET.Element, zf: zipfile.ZipFile) -> str:
        """Parse <hp:tbl> element into Markdown or HTML table."""
        grid = self._extract_table_grid(tbl_elem, zf)
        if grid is None:
            # Nested tables (layout container)
            unwrapped_blocks: List[str] = []
            for tc in tbl_elem.findall(".//hp:tc", NS):
                sublist = tc.find("hp:subList", NS)
                if sublist is not None:
                    for p in sublist.findall("hp:p", NS):
                        block = self._parse_paragraph(p, zf)
                        if block and block.strip():
                            unwrapped_blocks.append(block)
            return "\n\n".join(unwrapped_blocks)

        if self.table_format == "html":
            return grid.to_html()
        else:
            return grid.to_markdown(box_style=self.box_style)

    def _parse_table_cell(self, tc_elem: ET.Element, zf: zipfile.ZipFile) -> str:
        """Parse content of a single table cell <hp:tc>."""
        sublist = tc_elem.find("hp:subList", NS)
        if sublist is None:
            return ""

        cell_lines: List[str] = []
        for p in sublist.findall("hp:p", NS):
            line_parts: List[str] = []
            raw_runs: List[FormattedRun] = []
            for run in p.findall("hp:run", NS):
                # Check for embedded picture in cell
                pic = run.find("hp:pic", NS)
                if pic is not None:
                    img_md = self._parse_picture(pic, zf)
                    if img_md:
                        line_parts.append(img_md)

                char_pr_id = run.attrib.get("charPrIDRef", "")
                cp = self.char_properties.get(char_pr_id, {})
                run_text = ""
                for t in run.findall("hp:t", NS):
                    t_str = self._extract_text_from_t(t)
                    if t_str:
                        run_text += t_str
                if run_text:
                    bold = cp.get("bold", False)
                    italic = cp.get("italic", False)
                    strikeout = cp.get("strikeout", False)
                    underline = cp.get("underline", False)
                    raw_runs.append((run_text, bold, italic, strikeout, underline))

            merged = merge_runs(raw_runs)
            for t, b, it, s, u in merged:
                part = apply_inline_formatting(t, bold=b, italic=it, strikeout=s, underline=u)
                if part:
                    line_parts.append(part)

            p_text = "".join(line_parts).strip()
            if p_text:
                cell_lines.append(p_text)

        return "<br>".join(cell_lines)

    def _extract_picture_file(self, pic_elem: ET.Element, zf: zipfile.ZipFile) -> Tuple[Optional[Path], str]:
        """Extract picture binary and return (dest_file_path, alt_text)."""
        img_elem = pic_elem.find(".//{*}img")
        if img_elem is None:
            return None, ""

        bin_id = img_elem.attrib.get("binaryItemIDRef", "")
        if not bin_id:
            return None, ""

        bin_path = self.manifest.get(bin_id)
        if not bin_path:
            possible = [n for n in zf.namelist() if bin_id in n]
            if possible:
                bin_path = possible[0]

        if not bin_path or bin_path not in zf.namelist():
            return None, ""

        alt_text = bin_id
        comment = pic_elem.find(".//hp:shapeComment", NS)
        if comment is not None and comment.text:
            lines = comment.text.strip().split("\n")
            for line in lines:
                if "원본 그림의 이름:" in line:
                    alt_text = line.split(":", 1)[-1].strip()
                    break

        if not self.extract_images:
            return None, alt_text

        assert self.images_dir is not None
        self.images_dir.mkdir(parents=True, exist_ok=True)
        img_ext = os.path.splitext(bin_path)[1]
        dest_filename = f"{bin_id}{img_ext}"
        dest_filepath = self.images_dir / dest_filename

        try:
            with open(dest_filepath, "wb") as f_out:
                f_out.write(zf.read(bin_path))
            return dest_filepath, alt_text
        except Exception:
            return None, alt_text

    def _parse_picture(self, pic_elem: ET.Element, zf: zipfile.ZipFile) -> str:
        """Extract picture and return markdown image link."""
        dest_filepath, alt_text = self._extract_picture_file(pic_elem, zf)
        if dest_filepath:
            rel_path = f"{self.images_dir.name}/{dest_filepath.name}"
            return f"![{alt_text}]({rel_path})"
        elif alt_text:
            return f"<!-- Image: {alt_text} -->"
        return ""

    def parse_docx(self, output_docx_path: Optional[str | Path] = None) -> Path:
        """Parse HWPX file and export as Microsoft Word (.docx)."""
        from .docx_builder import DocxBuilder

        if not self.hwpx_path.exists():
            raise FileNotFoundError(f"HWPX file not found: {self.hwpx_path}")

        out_path = Path(output_docx_path) if output_docx_path else self.hwpx_path.with_suffix(".docx")

        if self.extract_images and not self.images_dir:
            self.images_dir = out_path.parent / f"{out_path.stem}_images"

        builder = DocxBuilder()

        with zipfile.ZipFile(self.hwpx_path, "r") as zf:
            self._parse_manifest(zf)
            self._parse_header(zf)
            for sec_href in self.section_hrefs:
                if sec_href not in zf.namelist():
                    continue
                try:
                    xml_data = zf.read(sec_href)
                    root = ET.fromstring(xml_data)
                    for p in root.findall("hp:p", NS):
                        self._render_paragraph_to_docx(p, zf, builder)
                except Exception:
                    pass

        builder.save(out_path)
        return out_path

    def _render_paragraph_to_docx(self, p_elem: ET.Element, zf: zipfile.ZipFile, builder):
        """Render a single HWPX paragraph or table into DocxBuilder."""
        tbl = p_elem.find(".//hp:tbl", NS)
        if tbl is not None:
            grid = self._extract_table_grid(tbl, zf)
            if grid is not None:
                builder.add_table_from_grid(grid, box_style=self.box_style, base_dir=self.images_dir)
            else:
                for tc in tbl.findall(".//hp:tc", NS):
                    sublist = tc.find("hp:subList", NS)
                    if sublist is not None:
                        for sub_p in sublist.findall("hp:p", NS):
                            self._render_paragraph_to_docx(sub_p, zf, builder)
            return

        raw_runs, images, max_font_size = self._extract_runs_and_images(p_elem, zf, for_docx=True)

        # Add images
        for img in images:
            builder.add_image(img)

        if not raw_runs:
            return

        merged_runs = merge_runs(raw_runs)
        is_all_bold = all(item[1] for item in merged_runs)
        clean_text = "".join(item[0] for item in merged_runs).strip()

        if not clean_text:
            return

        # Heading detection
        level = detect_heading_level(clean_text, max_font_size, is_all_bold)
        if level is not None:
            builder.add_heading(clean_text, level=level)
        else:
            builder.add_paragraph_runs(merged_runs)


def convert_hwpx(
    hwpx_path: str | Path,
    output_path: Optional[str | Path] = None,
    extract_images: bool = True,
    images_dir: Optional[str | Path] = None,
    table_format: str = "markdown",
    box_style: str = "table",
) -> str:
    """Convenience function to convert an HWPX file to Markdown."""
    parser = HWPXParser(
        hwpx_path=hwpx_path,
        extract_images=extract_images,
        images_dir=images_dir,
        table_format=table_format,
        box_style=box_style,
    )
    return parser.parse(output_md_path=output_path)


def convert_hwpx_to_docx(
    hwpx_path: str | Path,
    output_path: Optional[str | Path] = None,
    extract_images: bool = True,
    images_dir: Optional[str | Path] = None,
    box_style: str = "table",
) -> Path:
    """Convenience function to convert an HWPX file to Microsoft Word (.docx)."""
    parser = HWPXParser(
        hwpx_path=hwpx_path,
        extract_images=extract_images,
        images_dir=images_dir,
        box_style=box_style,
    )
    return parser.parse_docx(output_docx_path=output_path)
