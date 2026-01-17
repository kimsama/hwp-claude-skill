#!/usr/bin/env python3
"""
HWP Template Filling Script

Fills placeholders in an HWP template with data from a dictionary or JSON file.

Usage:
    python hwp_template_fill.py --template TEMPLATE --output OUTPUT [options]

Examples:
    python hwp_template_fill.py --template proposal.hwp --output filled.hwp --data company:"ABC Corp"
    python hwp_template_fill.py --template proposal.hwp --output filled.hwp --json data.json
"""

import argparse
import json
import sys
from pathlib import Path


def fill_template(template_path, output_path, data, visible=True):
    """
    Fill an HWP template with data.

    Args:
        template_path: Path to the template HWP file
        output_path: Path to save the filled document
        data: Dictionary of placeholder -> value mappings
        visible: Whether HWP should be visible (default: True)

    Returns:
        True if successful, False otherwise
    """
    try:
        from hwpapi.core import App

        # Open HWP
        app = App(is_visible=visible)

        # Open template
        template_path = Path(template_path).absolute()
        if not template_path.exists():
            print(f"Error: Template file not found: {template_path}")
            return False

        app.open(str(template_path))
        print(f"Opened template: {template_path}")

        # Replace placeholders
        replaced_count = 0
        for placeholder, value in data.items():
            # Support both {{key}} and {{ key }} formats
            placeholder_patterns = [
                f"{{{{{placeholder}}}}}",
                f"{{{{ {placeholder} }}}}"
            ]

            for pattern in placeholder_patterns:
                try:
                    app.replace_all(pattern, str(value))
                    replaced_count += 1
                    print(f"Replaced: {pattern} -> {value}")
                except Exception as e:
                    print(f"Warning: Could not replace {pattern}: {e}")

        print(f"Total replacements: {replaced_count}")

        # Save filled document
        output_path = Path(output_path).absolute()
        app.save_as(str(output_path))
        print(f"Saved filled document to: {output_path}")

        return True

    except ImportError:
        print("Error: hwpapi library not found. Install with: pip install hwpapi")
        return False
    except Exception as e:
        print(f"Error filling template: {e}")
        return False


def parse_data_args(data_args):
    """Parse data arguments in key:value format."""
    data = {}
    for arg in data_args:
        if ':' in arg:
            key, value = arg.split(':', 1)
            data[key.strip()] = value.strip()
        else:
            print(f"Warning: Invalid data format: {arg} (use key:value)")
    return data


def main():
    parser = argparse.ArgumentParser(
        description="Fill placeholders in an HWP template with data"
    )
    parser.add_argument(
        "--template", "-t",
        required=True,
        help="Path to the template HWP file"
    )
    parser.add_argument(
        "--output", "-o",
        required=True,
        help="Path to save the filled document"
    )
    parser.add_argument(
        "--data", "-d",
        action="append",
        help="Data in key:value format (can be used multiple times)"
    )
    parser.add_argument(
        "--json", "-j",
        help="Path to JSON file containing data"
    )
    parser.add_argument(
        "--visible", "-v",
        type=bool,
        default=True,
        help="Show HWP window (default: True)"
    )

    args = parser.parse_args()

    # Load data
    data = {}

    # From JSON file
    if args.json:
        json_path = Path(args.json)
        if not json_path.exists():
            print(f"Error: JSON file not found: {args.json}")
            sys.exit(1)
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
            print(f"Loaded {len(data)} items from {args.json}")

    # From command line arguments
    if args.data:
        cli_data = parse_data_args(args.data)
        data.update(cli_data)
        print(f"Added {len(cli_data)} items from command line")

    if not data:
        print("Error: No data provided. Use --data or --json")
        sys.exit(1)

    # Fill template
    success = fill_template(
        template_path=args.template,
        output_path=args.output,
        data=data,
        visible=args.visible
    )

    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
