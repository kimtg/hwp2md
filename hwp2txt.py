#!/usr/bin/env python3
"""
hwp2txt - Standalone HWP and HWPX to Plain Text (.txt) Converter.
Optimized for LLM prompting, RAG indexing, and search.
Run directly with: python hwp2txt.py <file.hwpx | file.hwp>
"""

import sys
from hwp2md.cli import main_text

if __name__ == "__main__":
    sys.exit(main_text())
