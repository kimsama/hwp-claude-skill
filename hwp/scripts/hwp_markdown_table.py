#!/usr/bin/env python3
"""
HWP Markdown Table Converter Script

Converts markdown tables to HWP tables and inserts them into a document.

Usage:
    python hwp_markdown_table.py --markdown "TABLE" [options]

Examples:
    python hwp_markdown_table.py --markdown "| A | B |\\n|---|---|\\n| 1 | 2 |" --output table.hwp
    python hwp_markdown_table.py --input table.md --output table.hwp
"""

import argparse
import re
import sys
from pathlib import Path


def parse_markdown_table(markdown_text):
    """
    Parse a markdown table into a 2D array.

    Args:
        markdown_text: Markdown table as string

    Returns:
        List of lists (rows of cells)
    """
    lines = [
        line.strip()
        for line in markdown_text.strip().split('\n')
        if line.strip()
    ]

    # Remove separator line (e.g., |---|---| or |:---|:---:|)
    lines = [
        line for line in lines
        if not re.match(r'^\|?[\s\-:]+\|?$', line)
    ]

    # Parse each row
    table_data = []
    for line in lines:
        # Remove leading/trailing pipes and split
        cells = [
            cell.strip()
            for cell in line.strip('|').split('|')
        ]
        if cells:  # Skip empty rows
            table_data.append(cells)

    return table_data


def create_hwp_table_from_markdown(markdown_text, output_path=None, visible=True):
    """
    Create an HWP document with a table from markdown.

    Args:
        markdown_text: Markdown table as string
        output_path: Path to save the document (optional)
        visible: Whether HWP should be visible (default: True)

    Returns:
        True if successful, False otherwise
    """
    try:
        from hwpapi.core import App

        # Parse markdown table
        data = parse_markdown_table(markdown_text)

        if not data:
            print("Error: No table data found in markdown")
            return False

        rows = len(data)
        cols = len(data[0]) if data else 0

        print(f"Parsed markdown table: {rows} rows x {cols} columns")

        # Create HWP document
        app = App(new_app=True, is_visible=visible)

        # Create table
        app.table.create(rows=rows, cols=cols)

        # Fill data
        print("Filling table...")
        for row_idx, row_data in enumerate(data):
            for col_idx, cell_text in enumerate(row_data):
                app.cell.move(row_idx, col_idx)
                app.cell.text = cell_text

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
        print(f"Error creating table from markdown: {e}")
        return False


def extract_first_markdown_table(text):
    """
    Extract the first markdown table from text.

    Args:
        text: Text containing markdown table(s)

    Returns:
        First markdown table as string, or None if not found
    """
    # Pattern to match markdown tables
    table_pattern = r'(?:^\|[^\n]+\|?\n)+'

    matches = re.findall(table_pattern, text, re.MULTILINE)

    if matches:
        return matches[0]
    return None


def main():
    parser = argparse.ArgumentParser(
        description="Convert markdown tables to HWP tables"
    )
    parser.add_argument(
        "--markdown", "-m",
        help='Markdown table as string (use \\n for newlines)'
    )
    parser.add_argument(
        "--input", "-i",
        help="Input markdown file path"
    )
    parser.add_argument(
        "--output", "-o",
        help="Output HWP file path"
    )
    parser.add_argument(
        "--visible", "-v",
        type=bool,
        default=True,
        help="Show HWP window (default: True)"
    )

    args = parser.parse_args()

    # Get markdown text
    markdown_text = None

    if args.markdown:
        markdown_text = args.markdown
    elif args.input:
        input_path = Path(args.input)
        if not input_path.exists():
            print(f"Error: Input file not found: {args.input}")
            sys.exit(1)
        with open(input_path, 'r', encoding='utf-8') as f:
            content = f.read()
            # Try to extract first table
            markdown_text = extract_first_markdown_table(content)
            if not markdown_text:
                # Use entire content
                markdown_text = content
    else:
        print("Error: Must provide either --markdown or --input")
        sys.exit(1)

    # Convert markdown newlines to actual newlines
    markdown_text = markdown_text.replace('\\n', '\n')

    success = create_hwp_table_from_markdown(
        markdown_text=markdown_text,
        output_path=args.output,
        visible=args.visible
    )

    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
