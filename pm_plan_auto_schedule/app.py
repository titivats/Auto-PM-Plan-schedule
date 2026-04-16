from __future__ import annotations

import argparse

from .generator import GenerationError, ensure_excel_available, generate_year_files
from .gui import run_ui


def run_cli(template: str, output_dir: str, year: int) -> int:
    try:
        ensure_excel_available()
        results = generate_year_files(
            template_path=template,
            output_dir=output_dir,
            year=year,
            log=lambda message: print(message),
        )
    except GenerationError as exc:
        print(f"ERROR: {exc}")
        return 1

    print(f"Generated {len(results)} file(s) in {output_dir}")
    return 0


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate yearly PM plan Excel files.")
    parser.add_argument("--template", help="Path to the Excel template.")
    parser.add_argument("--output", help="Output directory for generated files.")
    parser.add_argument("--year", type=int, help="Target year.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    if args.template and args.output and args.year:
        return run_cli(args.template, args.output, args.year)

    run_ui()
    return 0
