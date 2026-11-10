#!/usr/bin/env python3
"""
hwp2docx - Standalone HWP and HWPX to Microsoft Word (.docx) Converter.
Run directly with: python hwp2docx.py <file.hwpx | file.hwp>
"""

import sys
from hwp2md.cli import main_docx

if __name__ == "__main__":
    sys.exit(main_docx())
