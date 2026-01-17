#!/usr/bin/env python3
"""
HWP Table Creation Script

Creates a table in an HWP document with optional data.

Usage:
    python hwp_create_table.py --rows ROWS --cols COLS [options]

Examples:
    python hwp_create_table.py --rows 5 --cols 3 --output table.hwp
    python hwp_create_table.py --rows 3 --cols 2 --data "Name,Age\\nJohn,25\\nJane,30" --output table.hwp
"""

import argparse
import sys
from pathlib import Path


def create_table(rows, cols, data=None, output_path=None, visible=True):
    """
    Create a table in an HWP document.

    Args:
        rows: Number of rows
        cols: Number of columns
        data: Optional 2D list or CSV string with cell data
        output_path: Path to save the document (optional)
        visible: Whether HWP should be visible (default: True)

    Returns:
        True if successful, False otherwise
    """
    try:
        from hwpapi.core import App

        # Create or connect to HWP
        app = App(new_app=True, is_visible=visible)

        # Create table
        print(f"Creating table with {rows} rows and {cols} columns...")
        app.table.create(rows=rows, cols=cols)

        # Fill data if provided
        if data:
            print("Filling table with data...")

            # Parse data if it's a string (CSV format)
            if isinstance(data, str):
                rows_data = [row.split(',') for row in data.strip().split('\\n')]
            else:
                rows_data = data

            # Fill cells
            for row_idx in range(min(len(rows_data), rows)):
                for col_idx in range(min(len(rows_data[row_idx]), cols)):
                    cell_text = rows_data[row_idx][col_idx]
                    app.cell.move(row_idx, col_idx)
                    app.cell.text = str(cell_text)
                    print(f"  Cell ({row_idx}, {col_idx}): {cell_text}")

        # Save if output path provided
        if output_path:
            output_path = Path(output_path).absolute()
            app.save_as(str(output_path))
            print(f"Document saved to: {output_path}")
        else:
            print("Table created (document not saved)")

        return True

    except ImportError:
        print("Error: hwpapi library not found. Install with: pip install hwpapi")
        return False
    except Exception as e:
        print(f"Error creating table: {e}")
        return False


def parse_data_string(data_str):
    """Parse data string into 2D list."""
    if not data_str:
        return None

    rows = []
    for row in data_str.strip().split('\\n'):
        cells = [cell.strip() for cell in row.split(',')]
        rows.append(cells)

    return rows


def main():
    parser = argparse.ArgumentParser(description="Create a table in an HWP document")
    parser.add_argument("--rows", "-r", type=int, required=True, help="Number of rows")
    parser.add_argument("--cols", "-c", type=int, required=True, help="Number of columns")
    parser.add_argument(
        "--data", "-d",
        help='Table data in CSV format (use \\n for newlines). Example: "A,B\\nC,D"'
    )
    parser.add_argument(
        "--output", "-o",
        help="Output file path (e.g., table.hwp)"
    )
    parser.add_argument(
        "--visible", "-v",
        type=bool,
        default=True,
        help="Show HWP window (default: True)"
    )

    args = parser.parse_args()

    # Validate table dimensions
    if args.rows <= 0 or args.cols <= 0:
        print("Error: Rows and columns must be positive numbers")
        sys.exit(1)

    # Parse data
    data = None
    if args.data:
        data = parse_data_string(args.data)
        if len(data) > args.rows:
            print(f"Warning: Data has {len(data)} rows but table has only {args.rows} rows")

    success = create_table(
        rows=args.rows,
        cols=args.cols,
        data=data,
        output_path=args.output,
        visible=args.visible
    )

    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
