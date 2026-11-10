"""
Unit and integration tests for hwp2md.
"""

import os
import unittest
from pathlib import Path

from hwp2md.table_builder import TableGrid, get_east_asian_width, sanitize_cell_text
from hwp2md.converter import convert, detect_format


class TestTableBuilder(unittest.TestCase):
    def test_east_asian_width(self):
        self.assertEqual(get_east_asian_width("abc"), 3)
        self.assertEqual(get_east_asian_width("가나다"), 6)
        self.assertEqual(get_east_asian_width("a가b"), 4)

    def test_sanitize_cell_text(self):
        raw = "Line 1\nLine 2 | with pipe"
        sanitized = sanitize_cell_text(raw)
        self.assertEqual(sanitized, "Line 1<br>Line 2 \\| with pipe")

    def test_table_grid_basic(self):
        grid = TableGrid(num_rows=2, num_cols=2)
        grid.set_cell(0, 0, "Header A")
        grid.set_cell(0, 1, "Header B")
        grid.set_cell(1, 0, "Val 1")
        grid.set_cell(1, 1, "Val 2")

        md = grid.to_markdown()
        lines = md.splitlines()
        self.assertEqual(len(lines), 3)  # header, delimiter, data
        self.assertTrue(lines[0].startswith("| Header A"))
        self.assertTrue(lines[1].startswith("| ---"))
        self.assertTrue(lines[2].startswith("| Val 1"))

    def test_table_grid_spans(self):
        grid = TableGrid(num_rows=2, num_cols=2)
        # Rowspan 2 in col 0
        grid.set_cell(0, 0, "SpanVertical", row_span=2, col_span=1)
        grid.set_cell(0, 1, "TopRight")
        grid.set_cell(1, 1, "BottomRight")

        md = grid.to_markdown()
        lines = md.splitlines()
        self.assertEqual(len(lines), 3)
        self.assertIn("SpanVertical", lines[0])
        # Second row col 0 should be empty string
        self.assertIn("BottomRight", lines[2])

    def test_box_style_quote(self):
        grid = TableGrid(num_rows=1, num_cols=1)
        grid.set_cell(0, 0, "Callout Notice\nSecond line")
        quote_md = grid.to_markdown(box_style="quote")
        self.assertEqual(quote_md, "> Callout Notice\n> Second line")

    def test_html_table_output(self):
        grid = TableGrid(num_rows=2, num_cols=2)
        grid.set_cell(0, 0, "Merged", row_span=1, col_span=2)
        grid.set_cell(1, 0, "Cell 1")
        grid.set_cell(1, 1, "Cell 2")

        html = grid.to_html()
        self.assertIn('<td colspan="2">Merged</td>', html)
        self.assertIn("<td>Cell 1</td>", html)


class TestHWPXConversion(unittest.TestCase):
    def setUp(self):
        sample_dir = Path(__file__).resolve().parent.parent / "sample"
        self.hwpx_files = list(Path(".").glob("*.hwpx")) + list(sample_dir.glob("*.hwpx"))

    def test_hwpx_conversion(self):
        if not self.hwpx_files:
            self.skipTest("No .hwpx test file found.")

        sample_hwpx = self.hwpx_files[0]
        md = convert(sample_hwpx, extract_images=False)
        self.assertTrue(len(md) > 1000)
        # Check that markdown tables were generated
        self.assertIn("| ---", md)
        # Check that table text exists
        self.assertIn("보도자료", md)


class TestHWPConversion(unittest.TestCase):
    def test_hwp_conversion_if_sample_exists(self):
        hwp_sample = (
            Path(__file__).resolve().parent.parent
            / "sample"
            / "[별지 제21호서식] 일반과세자 부가가치세(예정¸ 확정¸ 기한후과세표준¸ 영세율 등 조기환급)신고서(부가가치세법 시행규칙).hwp"
        )
        if not hwp_sample.exists():
            hwp_sample = Path(
                "sample/[별지 제21호서식] 일반과세자 부가가치세(예정¸ 확정¸ 기한후과세표준¸ 영세율 등 조기환급)신고서(부가가치세법 시행규칙).hwp"
            )

        if not hwp_sample.exists():
            self.skipTest("Sample HWP not found at path.")

        md = convert(hwp_sample, extract_images=False)
        self.assertTrue(len(md) > 500)
        self.assertIn("| ---", md)
        self.assertIn("부가가치세", md)
        self.assertIn("일반과세자", md)


class TestFormatDetection(unittest.TestCase):
    def test_detect_hwpx(self):
        sample_dir = Path(__file__).resolve().parent.parent / "sample"
        hwpx_files = list(Path(".").glob("*.hwpx")) + list(sample_dir.glob("*.hwpx"))
        if hwpx_files:
            self.assertEqual(detect_format(hwpx_files[0]), "hwpx")

    def test_detect_hwp(self):
        hwp_sample = (
            Path(__file__).resolve().parent.parent
            / "sample"
            / "[별지 제21호서식] 일반과세자 부가가치세(예정¸ 확정¸ 기한후과세표준¸ 영세율 등 조기환급)신고서(부가가치세법 시행규칙).hwp"
        )
        if not hwp_sample.exists():
            hwp_sample = Path(
                "sample/[별지 제21호서식] 일반과세자 부가가치세(예정¸ 확정¸ 기한후과세표준¸ 영세율 등 조기환급)신고서(부가가치세법 시행규칙).hwp"
            )

        if hwp_sample.exists():
            self.assertEqual(detect_format(hwp_sample), "hwp")


if __name__ == "__main__":
    unittest.main()
