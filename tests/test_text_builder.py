"""
Unit tests for text extraction and metadata inspection.
"""

import unittest
from pathlib import Path

from hwp2md.text_builder import markdown_to_text, extract_metadata
from hwp2md.table_builder import TableGrid
from hwp2md.converter import convert_to_text, get_metadata


class TestTextBuilder(unittest.TestCase):
    def test_markdown_to_text_stripping(self):
        md = (
            "# Main Title\n\n"
            "> This is a quote\n\n"
            "This is **bold**, *italic*, <u>underline</u>, and ~~strike~~ text with [link](https://example.com).\n\n"
            "![alt text](image.png)\n\n"
            "| Header 1 | Header 2 |\n"
            "| --- | --- |\n"
            "| Cell A | Cell B |\n"
        )
        plain = markdown_to_text(md)
        self.assertNotIn("#", plain)
        self.assertNotIn(">", plain)
        self.assertNotIn("**", plain)
        self.assertNotIn("<u>", plain)
        self.assertNotIn("~~", plain)
        self.assertIn("Main Title", plain)
        self.assertIn("This is a quote", plain)
        self.assertIn("link", plain)
        self.assertNotIn("https://example.com", plain)
        self.assertIn("Cell A\tCell B", plain)

    def test_table_grid_to_text(self):
        grid = TableGrid(num_rows=2, num_cols=2)
        grid.set_cell(0, 0, "Col1")
        grid.set_cell(0, 1, "Col2")
        grid.set_cell(1, 0, "Val1")
        grid.set_cell(1, 1, "Val2")

        text = grid.to_text()
        lines = text.splitlines()
        self.assertEqual(len(lines), 2)
        self.assertEqual(lines[0], "Col1\tCol2")
        self.assertEqual(lines[1], "Val1\tVal2")

    def test_extract_metadata(self):
        sample_dir = Path(__file__).resolve().parent.parent / "sample"
        hwpx_files = list(sample_dir.glob("*.hwpx"))
        if hwpx_files:
            meta = extract_metadata(hwpx_files[0])
            self.assertEqual(meta["format"], "hwpx")
            self.assertIn("filesize_bytes", meta)
            self.assertGreater(meta["filesize_bytes"], 0)

        hwp_files = list(sample_dir.glob("*.hwp"))
        if hwp_files:
            meta_hwp = get_metadata(hwp_files[0])
            self.assertEqual(meta_hwp["format"], "hwp")
            self.assertIn("ole_streams_count", meta_hwp)

    def test_convert_to_text(self):
        sample_dir = Path(__file__).resolve().parent.parent / "sample"
        hwpx_files = list(sample_dir.glob("*.hwpx"))
        if hwpx_files:
            txt = convert_to_text(hwpx_files[0])
            self.assertGreater(len(txt), 100)
            self.assertIn("코리아 핀테크 위크", txt)


if __name__ == "__main__":
    unittest.main()
