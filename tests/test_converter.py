"""
Tests for converter with sample HWP and HWPX files.
"""

import unittest
from pathlib import Path
import tempfile
import docx

from hwp2md.converter import convert, convert_to_docx, detect_format
from hwp2md.hwpx_parser import convert_hwpx_to_docx
from hwp2md.hwp_parser import convert_hwp_to_docx


class TestSampleConversions(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.sample_dir = Path(__file__).parent.parent / "sample"
        cls.hwpx_files = list(cls.sample_dir.glob("*.hwpx"))
        cls.hwp_files = list(cls.sample_dir.glob("*.hwp"))

    def test_detect_format(self):
        if self.hwpx_files:
            self.assertEqual(detect_format(self.hwpx_files[0]), "hwpx")
        if self.hwp_files:
            self.assertEqual(detect_format(self.hwp_files[0]), "hwp")

    def test_hwpx_to_docx(self):
        if not self.hwpx_files:
            self.skipTest("No .hwpx sample file found.")

        hwpx_file = self.hwpx_files[0]
        with tempfile.TemporaryDirectory() as tmpdir:
            out_file = Path(tmpdir) / "output.docx"
            result = convert_to_docx(hwpx_file, output_path=out_file)

            self.assertTrue(result.exists())
            self.assertGreater(result.stat().st_size, 1000)

            # Check docx structure
            doc = docx.Document(str(result))
            self.assertGreater(len(doc.paragraphs), 10)
            self.assertGreater(len(doc.tables), 0)

    def test_hwp_to_docx(self):
        if not self.hwp_files:
            self.skipTest("No .hwp sample file found.")

        hwp_file = self.hwp_files[0]
        with tempfile.TemporaryDirectory() as tmpdir:
            out_file = Path(tmpdir) / "output.docx"
            result = convert_to_docx(hwp_file, output_path=out_file)

            self.assertTrue(result.exists())
            self.assertGreater(result.stat().st_size, 1000)

            # Check docx structure
            doc = docx.Document(str(result))
            self.assertGreater(len(doc.tables), 0)

    def test_convert_auto_docx_by_extension(self):
        if not self.hwpx_files:
            self.skipTest("No .hwpx sample file found.")

        hwpx_file = self.hwpx_files[0]
        with tempfile.TemporaryDirectory() as tmpdir:
            out_file = Path(tmpdir) / "auto.docx"
            result = convert(hwpx_file, output_path=out_file)

            self.assertEqual(Path(result), out_file)
            self.assertTrue(out_file.exists())


if __name__ == "__main__":
    unittest.main()
