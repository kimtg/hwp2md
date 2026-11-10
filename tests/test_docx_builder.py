"""
Unit tests for DocxBuilder and inline markdown parser.
"""

import unittest
from pathlib import Path
import tempfile
import docx

from hwp2md.docx_builder import DocxBuilder, parse_inline_markdown
from hwp2md.table_builder import TableGrid


class TestDocxBuilder(unittest.TestCase):
    def test_parse_inline_markdown(self):
        text = "Hello **bold** and *italic* and ~~strike~~ and plain"
        tokens = parse_inline_markdown(text)
        types = [t[0] for t in tokens]
        self.assertIn("bold", types)
        self.assertIn("italic", types)
        self.assertIn("strike", types)
        self.assertIn("text", types)

    def test_docx_builder_elements(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            out_path = Path(tmpdir) / "test.docx"
            builder = DocxBuilder(title="Test Document")

            builder.add_heading("Section 1", level=2)
            builder.add_paragraph_runs([
                ("This is ", False, False, False),
                ("bold text", True, False, False),
                (" and ", False, False, False),
                ("italic text", False, True, False),
            ])

            builder.add_markdown_paragraph("> Blockquote line 1<br>> Blockquote line 2")

            # Table test
            grid = TableGrid(num_rows=2, num_cols=2)
            grid.set_cell(0, 0, "Top Left", row_span=1, col_span=2)
            grid.set_cell(1, 0, "Bottom Left", row_span=1, col_span=1)
            grid.set_cell(1, 1, "Bottom Right", row_span=1, col_span=1)

            builder.add_table_from_grid(grid)
            builder.save(out_path)

            self.assertTrue(out_path.exists())
            doc = docx.Document(str(out_path))
            self.assertGreater(len(doc.paragraphs), 0)
            self.assertEqual(len(doc.tables), 1)
            self.assertEqual(len(doc.tables[0].rows), 2)
            self.assertEqual(len(doc.tables[0].columns), 2)


if __name__ == "__main__":
    unittest.main()
