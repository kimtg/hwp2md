#!/usr/bin/env python3
"""
hwp2md - Standalone HWP and HWPX to Markdown Converter.
Run directly with: python hwp2md.py <file.hwpx | file.hwp>
"""

import sys
from hwp2md.cli import main

if __name__ == "__main__":
    sys.exit(main())
