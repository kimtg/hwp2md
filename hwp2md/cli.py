"""
Command-line interface (CLI) for hwp2md, hwp2docx, and hwp2txt.
"""

from __future__ import annotations
import argparse
import glob
import json
import sys
from pathlib import Path
from typing import List, Optional

from .converter import convert, convert_to_docx, convert_to_text, get_metadata, detect_format
from . import __version__


def parse_args(
    args: List[str] = None,
    default_mode: str = "md",
) -> argparse.Namespace:
    prog_map = {
        "docx": "hwp2docx",
        "txt": "hwp2txt",
        "md": "hwp2md",
    }
    prog_name = prog_map.get(default_mode, "hwp2md")

    desc = (
        f"Convert HWP / HWPX documents to Markdown (.md), Microsoft Word (.docx), or plain text (.txt) "
        f"without Hancom Office or MS Office dependency."
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

  # Extract clean plain text for LLM / RAG
  hwp2md document.hwpx --text
  hwp2txt document.hwpx

  # Inspect document metadata
  hwp2md document.hwpx --meta

  # Convert entire directory recursively
  hwp2md ./documents/ -r --docx
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
        help="Output file or directory. Extension (.docx, .txt, .md) determines format automatically.",
    )
    parser.add_argument(
        "--docx",
        action="store_true",
        default=(default_mode == "docx"),
        help="Convert to Microsoft Word (.docx) format.",
    )
    parser.add_argument(
        "--text",
        action="store_true",
        default=(default_mode == "txt"),
        help="Extract clean plain text (.txt) for LLM / RAG pipelines.",
    )
    parser.add_argument(
        "--meta",
        action="store_true",
        help="Display document metadata in JSON format and exit.",
    )
    parser.add_argument(
        "--format",
        choices=["md", "markdown", "docx", "txt", "text"],
        default=default_mode if default_mode != "md" else None,
        help="Explicitly set output format: 'md', 'docx', or 'txt'.",
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
        help="Print converted text/markdown directly to standard output.",
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


def _safe_symbol(symbol: str, fallback: str) -> str:
    """Return symbol if stdout encoding supports it, otherwise ASCII fallback."""
    enc = sys.stdout.encoding or "utf-8"
    try:
        symbol.encode(enc)
        return symbol
    except (UnicodeEncodeError, LookupError):
        return fallback


def run_cli(args: List[str] = None, default_mode: str = "md") -> int:
    parsed = parse_args(args, default_mode=default_mode)
    files = collect_files(parsed.inputs, recursive=parsed.recursive)

    if not files:
        print("[!] No .hwp or .hwpx files found matching input patterns.", file=sys.stderr)
        return 1

    # Metadata mode
    if parsed.meta:
        meta_list = [get_metadata(f) for f in files]
        print(json.dumps(meta_list if len(meta_list) > 1 else meta_list[0], indent=2, ensure_ascii=False))
        return 0

    # Determine mode
    out_ext = Path(parsed.output).suffix.lower() if parsed.output else ""
    is_docx = parsed.docx or parsed.format == "docx" or out_ext == ".docx"
    is_text = parsed.text or parsed.format in ("txt", "text") or out_ext == ".txt"

    if is_text:
        target_ext = ".txt"
    elif is_docx:
        target_ext = ".docx"
    else:
        target_ext = ".md"

    if is_docx and parsed.stdout:
        print("[!] --stdout is not supported for binary .docx output. Generating file instead.", file=sys.stderr)
        parsed.stdout = False

    success_count = 0
    fail_count = 0
    total = len(files)

    for idx, file_path in enumerate(files, 1):
        try:
            if total > 1 and not parsed.stdout:
                prefix = f"[{idx}/{total}] "
            else:
                prefix = ""

            # Determine output destination
            if parsed.stdout:
                out_path = None
            elif parsed.output:
                out_p = Path(parsed.output)
                is_out_dir = (
                    out_p.is_dir()
                    or len(files) > 1
                    or str(parsed.output).endswith(("/", "\\"))
                    or not out_p.suffix
                )
                if is_out_dir:
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
                to_text=is_text,
            )

            if parsed.stdout:
                sys.stdout.write(str(result))
            else:
                dest_name = out_path.name if out_path else str(result)
                ok_sym = _safe_symbol("✓", "OK")
                print(f"{prefix}[{ok_sym}] Converted: {file_path.name} -> {dest_name}")
            success_count += 1
        except Exception as e:
            err_sym = _safe_symbol("✗", "ERR")
            print(f"{prefix}[{err_sym}] Failed to convert {file_path.name}: {e}", file=sys.stderr)
            fail_count += 1

    if not parsed.stdout and len(files) > 1:
        print(f"\nDone: {success_count} succeeded, {fail_count} failed.")

    return 0 if fail_count == 0 else 1


def main(args: List[str] = None) -> int:
    return run_cli(args, default_mode="md")


def main_docx(args: List[str] = None) -> int:
    return run_cli(args, default_mode="docx")


def main_text(args: List[str] = None) -> int:
    return run_cli(args, default_mode="txt")


if __name__ == "__main__":
    sys.exit(main())
