#!/usr/bin/env python3
"""
HWP Document Creation Script

Creates a new HWP document with optional initial content.

Usage:
    python hwp_create_document.py [--output OUTPUT] [--title TITLE] [--content CONTENT]

Examples:
    python hwp_create_document.py --output report.hwp
    python hwp_create_document.py --output report.hwp --title "Monthly Report" --content "This is a report"
"""

import argparse
import sys
from pathlib import Path


def create_document(output_path=None, title=None, content=None, visible=True):
    """
    Create a new HWP document.

    Args:
        output_path: Path to save the document (optional)
        title: Document title (optional)
        content: Initial content (optional)
        visible: Whether HWP should be visible (default: True)

    Returns:
        True if successful, False otherwise
    """
    try:
        from hwpapi.core import App

        # Create new HWP application
        app = App(new_app=True, is_visible=visible)

        # Add title if provided
        if title:
            app.set_charshape(fontname="맑은 고딕", height=18, bold=True)
            app.insert_text(title)
            app.insert_text("\r\n")

        # Add content if provided
        if content:
            app.set_charshape(height=11, bold=False)
            app.insert_text(content)

        # Save if output path provided
        if output_path:
            app.save_as(str(output_path))
            print(f"Document saved to: {output_path}")
        else:
            print("New document created (not saved)")

        return True

    except ImportError:
        print("Error: hwpapi library not found. Install with: pip install hwpapi")
        return False
    except Exception as e:
        print(f"Error creating document: {e}")
        return False


def main():
    parser = argparse.ArgumentParser(description="Create a new HWP document")
    parser.add_argument("--output", "-o", help="Output file path (e.g., report.hwp)")
    parser.add_argument("--title", "-t", help="Document title")
    parser.add_argument("--content", "-c", help="Initial content")
    parser.add_argument("--visible", "-v", type=bool, default=True, help="Show HWP window (default: True)")

    args = parser.parse_args()

    output_path = Path(args.output) if args.output else None

    # Validate output path
    if output_path:
        output_path = output_path.absolute()
        if output_path.suffix.lower() != '.hwp':
            print("Warning: Output file should have .hwp extension")

    success = create_document(
        output_path=output_path,
        title=args.title,
        content=args.content,
        visible=args.visible
    )

    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
