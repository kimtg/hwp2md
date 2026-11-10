"""
Command-line interface (CLI) for hwp2md and hwp2docx.
"""

from __future__ import annotations
import argparse
import glob
import sys
from pathlib import Path
from typing import List, Optional

from .converter import convert, convert_to_docx, detect_format
from . import __version__


def parse_args(args: List[str] = None, default_to_docx: bool = False) -> argparse.Namespace:
    prog_name = "hwp2docx" if default_to_docx else "hwp2md"
    desc = (
        "Convert HWP / HWPX documents to Microsoft Word (.docx) without Hancom Office dependency."
        if default_to_docx
        else "Convert HWP / HWPX documents to Markdown (.md) or Microsoft Word (.docx) without Hancom Office dependency."
    )

    parser = argparse.ArgumentParser(
        prog=prog_name,
        description=desc,
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Convert to markdown (default)
  hwp2md document.hwpx
  hwp2md document.hwp

  # Convert to Word (.docx)
  hwp2md document.hwpx --docx
  hwp2md document.hwp -o report.docx
  hwp2docx document.hwpx

  # Specify output file
  hwp2md report.hwpx -o output.md

  # Convert all files in current folder to Word (.docx)
  hwp2md *.hwpx --docx

  # Convert entire directory recursively
  hwp2md ./documents/ -r

  # Convert with single-cell boxes as blockquotes
  hwp2md document.hwpx --box-style quote

  # Convert tables as HTML for maximum layout preservation (Markdown mode)
  hwp2md document.hwpx --table-format html

  # Disable image extraction
  hwp2md document.hwpx --no-images
""",
    )

    parser.add_argument(
        "inputs",
        nargs="+",
        help="Input .hwp or .hwpx file(s) or directories (wildcards supported).",
    )
    parser.add_argument(
        "-o",
        "--output",
        help="Output file or directory. If extension is .docx, automatically converts to Word.",
    )
    parser.add_argument(
        "--docx",
        action="store_true",
        default=default_to_docx,
        help="Convert to Microsoft Word (.docx) format instead of Markdown.",
    )
    parser.add_argument(
        "--format",
        choices=["md", "markdown", "docx"],
        default="docx" if default_to_docx else None,
        help="Explicitly set output format: 'md' (Markdown) or 'docx' (Word).",
    )
    parser.add_argument(
        "--table-format",
        choices=["markdown", "html"],
        default="markdown",
        help="Table conversion format (Markdown mode): 'markdown' (GFM table, default) or 'html' (HTML <table>).",
    )
    parser.add_argument(
        "--box-style",
        choices=["table", "quote", "plain"],
        default="table",
        help="Formatting for 1x1 callout boxes: 'table' (default), 'quote' (> blockquote), or 'plain'.",
    )
    parser.add_argument(
        "--no-images",
        action="store_true",
        help="Do not extract embedded images.",
    )
    parser.add_argument(
        "--images-dir",
        help="Custom directory to store extracted images (default: '<docname>_images/').",
    )
    parser.add_argument(
        "-r",
        "--recursive",
        action="store_true",
        help="Recursively search directories for .hwp and .hwpx files.",
    )
    parser.add_argument(
        "--stdout",
        action="store_true",
        help="Print converted markdown directly to standard output (Markdown mode only).",
    )
    parser.add_argument(
        "-v",
        "--version",
        action="version",
        version=f"%(prog)s {__version__}",
    )

    return parser.parse_args(args)


def collect_files(input_patterns: List[str], recursive: bool = False) -> List[Path]:
    """Collect all matching .hwp and .hwpx files."""
    matched_files: List[Path] = []
    seen = set()

    for pattern in input_patterns:
        p = Path(pattern)
        if p.is_dir():
            search_pattern = "**/*" if recursive else "*"
            for ext in (".hwpx", ".hwp", ".HWPX", ".HWP"):
                for f in p.glob(f"{search_pattern}{ext}"):
                    if f.is_file() and f.resolve() not in seen:
                        seen.add(f.resolve())
                        matched_files.append(f)
        else:
            # Handle glob patterns or exact files
            expanded = glob.glob(pattern, recursive=recursive)
            if not expanded and p.exists():
                expanded = [str(p)]

            for exp in expanded:
                ep = Path(exp)
                if ep.is_file() and ep.resolve() not in seen:
                    if ep.suffix.lower() in (".hwp", ".hwpx"):
                        seen.add(ep.resolve())
                        matched_files.append(ep)

    return matched_files


def run_cli(args: List[str] = None, default_to_docx: bool = False) -> int:
    parsed = parse_args(args, default_to_docx=default_to_docx)
    files = collect_files(parsed.inputs, recursive=parsed.recursive)

    if not files:
        print("[!] No .hwp or .hwpx files found matching input patterns.", file=sys.stderr)
        return 1

    # Determine if docx mode
    is_docx = (
        parsed.docx
        or parsed.format == "docx"
        or (parsed.output and Path(parsed.output).suffix.lower() == ".docx")
    )

    if is_docx and parsed.stdout:
        print("[!] --stdout is not supported for binary .docx output. Generating file instead.", file=sys.stderr)
        parsed.stdout = False

    target_ext = ".docx" if is_docx else ".md"
    success_count = 0
    fail_count = 0

    for file_path in files:
        try:
            # Determine output destination
            if parsed.stdout:
                out_path = None
            elif parsed.output:
                out_p = Path(parsed.output)
                if out_p.is_dir() or len(files) > 1:
                    out_p.mkdir(parents=True, exist_ok=True)
                    out_path = out_p / f"{file_path.stem}{target_ext}"
                else:
                    out_path = out_p
            else:
                out_path = file_path.with_suffix(target_ext)

            result = convert(
                input_path=file_path,
                output_path=out_path,
                extract_images=not parsed.no_images,
                images_dir=parsed.images_dir,
                table_format=parsed.table_format,
                box_style=parsed.box_style,
                to_docx=is_docx,
            )

            if parsed.stdout:
                sys.stdout.write(str(result))
            else:
                dest_name = out_path.name if out_path else str(result)
                print(f"[✓] Converted: {file_path.name} -> {dest_name}")
            success_count += 1
        except Exception as e:
            print(f"[✗] Failed to convert {file_path.name}: {e}", file=sys.stderr)
            fail_count += 1

    if not parsed.stdout and len(files) > 1:
        print(f"\nDone: {success_count} succeeded, {fail_count} failed.")

    return 0 if fail_count == 0 else 1


def main(args: List[str] = None) -> int:
    return run_cli(args, default_to_docx=False)


def main_docx(args: List[str] = None) -> int:
    return run_cli(args, default_to_docx=True)


if __name__ == "__main__":
    sys.exit(main())
