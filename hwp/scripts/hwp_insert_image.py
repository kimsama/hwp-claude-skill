#!/usr/bin/env python3
"""
HWP Image Insertion Script

Inserts images into HWP documents at cursor position.

Usage:
    python hwp_insert_image.py --image IMAGE_PATH [options]

Examples:
    python hwp_insert_image.py --image logo.png --output document.hwp
    python hwp_insert_image.py --image photo.jpg --document existing.hwp --width 100 --height 100
"""

import argparse
import sys
from pathlib import Path


def insert_image(image_path, document_path=None, output_path=None, width=None, height=None, maintain_ratio=True, visible=True):
    """
    Insert an image into an HWP document.

    Args:
        image_path: Path to the image file
        document_path: Path to existing HWP document (optional, creates new if not provided)
        output_path: Path to save the document (optional)
        width: Image width in millimeters (optional)
        height: Image height in millimeters (optional)
        maintain_ratio: Whether to maintain aspect ratio (default: True)
        visible: Whether HWP should be visible (default: True)

    Returns:
        True if successful, False otherwise
    """
    try:
        from hwpapi.core import App
        from hwpapi.functions import mili2unit

        # Validate image path
        image_path = Path(image_path).absolute()
        if not image_path.exists():
            print(f"Error: Image file not found: {image_path}")
            return False

        # Open or create document
        if document_path:
            app = App(is_visible=visible)
            doc_path = Path(document_path).absolute()
            if not doc_path.exists():
                print(f"Error: Document not found: {document_path}")
                return False
            app.open(str(doc_path))
            print(f"Opened document: {doc_path}")
        else:
            app = App(new_app=True, is_visible=visible)
            print("Created new document")

        # Insert picture
        print(f"Inserting image: {image_path}")

        if width or height:
            # Use action for sized insertion
            action = app.actions.InsertPicture
            action.pset.FileName = str(image_path)

            if width:
                action.pset.Width = mili2unit(width)
                print(f"  Width: {width}mm")

            if height:
                action.pset.Height = mili2unit(height)
                print(f"  Height: {height}mm")

            if maintain_ratio:
                action.pset.SizeManipulate = 1  # Maintain aspect ratio
                print(f"  Maintain aspect ratio: Yes")

            action.run()
        else:
            # Simple insertion
            app.insert_picture(str(image_path))

        # Save if output path provided
        if output_path:
            output_path = Path(output_path).absolute()
            app.save_as(str(output_path))
            print(f"Document saved to: {output_path}")
        else:
            print("Image inserted (document not saved)")

        return True

    except ImportError:
        print("Error: hwpapi library not found. Install with: pip install hwpapi")
        return False
    except Exception as e:
        print(f"Error inserting image: {e}")
        return False


def main():
    parser = argparse.ArgumentParser(description="Insert an image into an HWP document")
    parser.add_argument("--image", "-i", required=True, help="Path to the image file")
    parser.add_argument("--document", "-d", help="Path to existing HWP document")
    parser.add_argument("--output", "-o", help="Path to save the document")
    parser.add_argument("--width", "-w", type=float, help="Image width in millimeters")
    parser.add_argument("--height", "-h", type=float, help="Image height in millimeters")
    parser.add_argument(
        "--no-maintain-ratio",
        action="store_true",
        help="Do not maintain aspect ratio when resizing"
    )
    parser.add_argument(
        "--visible", "-v",
        type=bool,
        default=True,
        help="Show HWP window (default: True)"
    )

    args = parser.parse_args()

    success = insert_image(
        image_path=args.image,
        document_path=args.document,
        output_path=args.output,
        width=args.width,
        height=args.height,
        maintain_ratio=not args.no_maintain_ratio,
        visible=args.visible
    )

    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
