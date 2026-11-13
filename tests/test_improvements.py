"""
Regression and unit tests for hwp2md improvements.
"""

import io
import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from hwp2md.cli import run_cli
from hwp2md.docx_builder import DocxBuilder, add_hyperlink_to_paragraph
from hwp2md.formatting import apply_inline_formatting, format_paragraph_block, merge_runs
from hwp2md.hwp_parser import HWPParser, PYHWP_AVAILABLE
from hwp2md.hwpx_parser import HWPXParser
from hwp2md.table_builder import TableGrid
from hwp2md.text_builder import extract_metadata


class TestFormattingHelpers(unittest.TestCase):
    def test_merge_and_apply_formatting(self):
        runs = [
            ("Hello ", True, False, False, False),
            ("World", True, False, False, False),
            ("!", False, True, False, False),
        ]
        merged = merge_runs(runs)
        self.assertEqual(len(merged), 2)
        self.assertEqual(merged[0][0], "Hello World")
        self.assertEqual(
            apply_inline_formatting(merged[0][0], bold=True),
            "**Hello World**",
        )

    def test_heading_strips_underline_tags(self):
        block = format_paragraph_block("<u>**Title**</u>", max_font_size=18.0, is_all_bold=True)
        self.assertEqual(block, "# Title")


class TestTableGridImprovements(unittest.TestCase):
    def test_html_escapes_special_chars_but_preserves_br(self):
        grid = TableGrid(1, 2)
        grid.set_cell(0, 0, "A & B <tag>\nLine 2")
        grid.set_cell(0, 1, "OK")
        html_out = grid.to_html()
        self.assertIn("<td>A &amp; B &lt;tag&gt;<br>Line 2</td>", html_out)

    def test_set_cell_overwrite_does_not_duplicate_cells_info(self):
        grid = TableGrid(1, 2)
        grid.set_cell(0, 0, "Old")
        grid.set_cell(0, 0, "New")
        self.assertEqual(len(grid.cells_info), 1)
        self.assertEqual(grid.cells_info[0][4], "New")


class TestDocxBuilderImprovements(unittest.TestCase):
    def test_hyperlink_xml_escaping_and_east_asian_font(self):
        builder = DocxBuilder()
        p = builder.doc.add_paragraph()
        add_hyperlink_to_paragraph(p, "https://example.com/?a=1&b=2", "Q&A <Link>")
        xml_str = ET.tostring(p._p, encoding="unicode")
        # Should be a real w:hyperlink element, not fallback plain text
        self.assertIn("hyperlink", xml_str)
        self.assertIn("Q&amp;A &lt;Link&gt;", xml_str)

        builder.add_paragraph_runs([("한글 테스트", True, False, False, False)])
        last_p = builder.doc.paragraphs[-1]
        last_xml = ET.tostring(last_p._p, encoding="unicode")
        self.assertIn("맑은 고딕", last_xml)


class TestParserImprovements(unittest.TestCase):
    @unittest.skipUnless(PYHWP_AVAILABLE, "pyhwp not installed")
    def test_hwp_paragraph_no_duplicate_and_hyperlink(self):
        parser = HWPParser("dummy.hwp")
        parser.bold_classes.add("boldcls")
        xhtml = (
            '<p xmlns="http://www.w3.org/1999/xhtml">'
            '<span class="boldcls">중요</span> 내용 '
            '<a href="https://example.com">링크</a>'
            "</p>"
        )
        elem = ET.fromstring(xhtml)
        md = parser._parse_paragraph(elem)
        self.assertEqual(md.count("**중요**"), 1)
        self.assertIn("[링크](https://example.com)", md)

    def test_hwpx_hyperlink_parsing_and_metadata_extensionless(self):
        with tempfile.TemporaryDirectory() as tmp:
            hwpx_path = Path(tmp) / "no_ext_doc"
            content_hpf = (
                '<?xml version="1.0" encoding="UTF-8"?>'
                '<opf:package xmlns:opf="http://www.idpf.org/2007/opf">'
                '<opf:metadata><opf:title>테스트 제목</opf:title></opf:metadata>'
                '<opf:manifest><opf:item id="sec0" href="Contents/section0.xml"/></opf:manifest>'
                '<opf:spine><opf:itemref idref="sec0"/></opf:spine>'
                "</opf:package>"
            )
            section0 = (
                '<?xml version="1.0" encoding="UTF-8"?>'
                '<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" '
                'xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">'
                "<hp:p>"
                "<hp:run><hp:t>방문: </hp:t></hp:run>"
                '<hp:run><hp:ctrl><hp:fieldBegin type="HYPERLINK">'
                '<hp:parameters><hp:stringParam name="Path">https://hancom.com</hp:stringParam></hp:parameters>'
                "</hp:fieldBegin></hp:ctrl></hp:run>"
                "<hp:run><hp:t>한컴</hp:t></hp:run>"
                "<hp:run><hp:ctrl><hp:fieldEnd/></hp:ctrl></hp:run>"
                "</hp:p>"
                "</hs:sec>"
            )
            with zipfile.ZipFile(hwpx_path, "w") as zf:
                zf.writestr("Contents/content.hpf", content_hpf)
                zf.writestr("Contents/section0.xml", section0)

            meta = extract_metadata(hwpx_path)
            self.assertEqual(meta["format"], "hwpx")
            self.assertEqual(meta["title"], "테스트 제목")

            md = HWPXParser(hwpx_path, extract_images=False).parse()
            self.assertIn("방문: [한컴](https://hancom.com)", md)

    def test_cli_single_file_output_dir_without_extension(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            hwpx_path = tmp_path / "sample.hwpx"
            section0 = (
                '<?xml version="1.0" encoding="UTF-8"?>'
                '<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" '
                'xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">'
                "<hp:p><hp:run><hp:t>Hello CLI</hp:t></hp:run></hp:p>"
                "</hs:sec>"
            )
            with zipfile.ZipFile(hwpx_path, "w") as zf:
                zf.writestr("Contents/section0.xml", section0)

            out_dir = tmp_path / "out_dir"
            rc = run_cli([str(hwpx_path), "-o", str(out_dir), "--no-images"])
            self.assertEqual(rc, 0)
            self.assertTrue((out_dir / "sample.md").exists())


if __name__ == "__main__":
    unittest.main()
