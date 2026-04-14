"""CLI entry points for workbook indexing."""

import argparse
from pathlib import Path

from vector_search_app.db import ensure_schema
from vector_search_app.service import index_uploaded_workbook


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--workbook-path",
        required=True,
        help="Path to the Excel workbook to index.",
    )
    parser.add_argument(
        "--worksheet-name",
        default=None,
        help="Optional worksheet name to index. Defaults to all sheets.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    workbook_path = Path(args.workbook_path)
    workbook_bytes = workbook_path.read_bytes()
    ensure_schema()
    result = index_uploaded_workbook(
        workbook_bytes,
        workbook_path.name,
        args.worksheet_name,
    )
    print(
        f"Indexed {result['indexed_rows']} rows from {result['workbook_name']}"
    )


if __name__ == "__main__":
    main()
