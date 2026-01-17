---
name: hwp
description: HWP (Hangul Word Processor) document manipulation skill. Use when Claude needs to create, edit, format, or manipulate HWP documents including: (1) Creating new documents or opening existing HWP files, (2) Inserting and editing text with font formatting, (3) Creating and manipulating tables, (4) Working with templates and filling placeholders, (5) Converting markdown tables to HWP tables, (6) Inserting images, (7) Batch processing multiple HWP operations, (8) Applying document formatting (fonts, sizes, bold, italic, paragraphs)
---

# HWP Document Manipulation

This skill provides comprehensive capabilities for working with HWP (Hangul Word Processor) documents using the hwpapi Python library.

## Prerequisites

- Windows OS with Hancom Office HWP installed
- hwpapi Python library (`pip install hwpapi`)
- pywin32 for COM automation

## Quick Start

### Initialize HWP Application

```python
from hwpapi.core import App

# Open HWP (connects to running instance or starts new)
app = App()

# Run in background (invisible)
app = App(is_visible=False)
```

### Basic Document Operations

```python
# Create new document
app = App(new_app=True)

# Open existing file
app.open("path/to/document.hwp")

# Save document
app.save()  # Save to current path
app.save_as("path/to/new_document.hwp")  # Save to new path

# Close document (not quit HWP)
app.close()
```

## Core Operations

### Text Manipulation

#### Insert Text

```python
# Insert text at cursor position
app.insert_text("Hello, World!")

# Insert with line break
app.insert_text("First line\r\nSecond line")

# Insert Korean text
app.insert_text("안녕하세요 한글입니다.")
```

#### Set Font Properties

```python
# Simple font setting
app.set_charshape(fontname="맑은 고딕", height=12)

# Full font properties
app.set_charshape(
    fontname="맑은 고딕",
    height=14,
    bold=True,
    italic=False,
    underline=False,
    color="#000000"
)

# Font size in points (1pt = 100 HWPUNIT)
app.set_charshape(height=12)  # 12pt
```

#### Get Text

```python
# Get text from current position
text = app.get_text()

# Get selected text
selected = app.get_selected_text()
```

#### Find and Replace

```python
# Find text
found = app.find_text("search term")

# Replace all occurrences
app.replace_all("old text", "new text")
```

### Paragraph Formatting

```python
# Set paragraph properties
app.set_parashape(
    line_spacing=160,  # 160% line spacing
    left_margin=0,
    right_margin=0,
    indent=0
)

# Align text (0=left, 1=center, 2=right, 3=justify)
from hwpapi.constants import Alignment
app.actions.ParaShape.pset.Align = Alignment.CENTER
app.actions.ParaShape.run()
```

### Navigation

```python
# Move to document positions
app.move.top_of_file()
app.move.bottom()
app.move.page_start()
app.move.page_end()

# Move by list/paragraph
app.move.current_list(para=1, pos=0)
```

## Table Operations

### Create Table

```python
# Using table accessor (simplest)
app.table.create(rows=5, cols=3)

# Using action directly (more control)
app.actions.TableCreate.pset.Rows = 5
app.actions.TableCreate.pset.Cols = 3
app.actions.TableCreate.pset.Width = 5000  # Width in HWPUNIT
app.actions.TableCreate.pset.TableSizeMode = 0  # Fixed width
app.actions.TableCreate.run()
```

### Fill Table Cells

```python
# Navigate to cell and set content
app.cell.text = "Cell content"

# Move to specific cell
app.cell.move(0, 0)  # Row 0, Col 0
app.cell.text = "First cell"

# Fill from list data
data = [
    ["Header 1", "Header 2", "Header 3"],
    ["Data 1", "Data 2", "Data 3"],
    ["Data 4", "Data 5", "Data 6"]
]

for row_idx, row in enumerate(data):
    for col_idx, cell_data in enumerate(row):
        app.cell.move(row_idx, col_idx)
        app.cell.text = cell_data
```

### Format Cells

```python
# Set cell border
app.cell.set_border(type="solid", width=1)

# Set cell background color
app.cell.set_color("#FFFF00")  # Yellow background

# Merge cells
app.actions.CellMerge.run()
```

## Template Operations

### Fill Template with Placeholders

Template files should use `{{placeholder_name}}` format for placeholders.

```python
# Open template
app.open("path/to/proposal_template.hwp")

# Replace placeholders
app.replace_all("{{company_name}}", "ABC Corporation")
app.replace_all("{{date}}", "2024-01-18")
app.replace_all("{{amount}}", "1,000,000")

# Save filled document
app.save_as("path/to/filled_proposal.hwp")
```

### Template Data from Dictionary

```python
def fill_template(app, template_path, data):
    """Fill template with data from dictionary."""
    app.open(template_path)

    for placeholder, value in data.items():
        app.replace_all(f"{{{{{placeholder}}}}}", str(value))

    return app

# Usage
data = {
    "company_name": "ABC Corp",
    "representative": "John Doe",
    "date": "2024-01-18",
    "amount": "5,000,000원"
}

fill_template(app, "template.hwp", data)
app.save_as("output.hwp")
```

## Markdown to HWP Table Conversion

```python
import re

def parse_markdown_table(markdown):
    """Parse markdown table into 2D array."""
    lines = [line.strip() for line in markdown.strip().split('\n') if line.strip()]

    # Remove separator line (e.g., |---|---|)
    lines = [line for line in lines if not re.match(r'^\|?[\s\-:]+\|?$', line)]

    # Parse each row
    table_data = []
    for line in lines:
        # Remove leading/trailing pipes and split
        cells = [cell.strip() for cell in line.strip('|').split('|')]
        table_data.append(cells)

    return table_data

def create_hwp_table_from_markdown(app, markdown):
    """Create HWP table from markdown table."""
    data = parse_markdown_table(markdown)

    # Create table
    rows = len(data)
    cols = len(data[0]) if data else 0
    app.table.create(rows=rows, cols=cols)

    # Fill data
    for row_idx, row_data in enumerate(data):
        for col_idx, cell_text in enumerate(row_data):
            app.cell.move(row_idx, col_idx)
            app.cell.text = cell_text

# Example markdown table
markdown_table = """
| Header 1 | Header 2 | Header 3 |
|----------|----------|----------|
| Data 1   | Data 2   | Data 3   |
| Data 4   | Data 5   | Data 6   |
"""

create_hwp_table_from_markdown(app, markdown_table)
```

## Image Insertion

```python
# Insert picture at cursor position
app.insert_picture("path/to/image.png")

# Insert with sizing (via action)
app.actions.InsertPicture.pset.FileName = "path/to/image.png"
app.actions.InsertPicture.pset.Width = 5000  # Width in HWPUNIT
app.actions.InsertPicture.pset.Height = 3000  # Height in HWPUNIT
app.actions.InsertPicture.pset.SizeManipulate = 1  # Maintain aspect ratio
app.actions.InsertPicture.run()
```

## Document Formatting

### Page Setup

```python
# Set page properties
app.actions.PageSetup.pset.PaperWidth = 21000  # 210mm in HWPUNIT (283 per mm)
app.actions.PageSetup.pset.PaperHeight = 29700  # 297mm (A4)
app.actions.PageSetup.run()
```

### Section Formatting

```python
# Set section properties
app.actions.SecDef.pset.StartPage = 0
app.actions.SecDef.run()
```

## Batch Operations

### Execute Multiple Operations

```python
def batch_create_reports(app, template_path, data_list, output_dir):
    """Create multiple reports from template."""
    results = []

    for i, data in enumerate(data_list):
        # Open template
        app.open(template_path)

        # Fill placeholders
        for key, value in data.items():
            app.replace_all(f"{{{{{key}}}}}", str(value))

        # Save
        output_path = f"{output_dir}/report_{i+1}.hwp"
        app.save_as(output_path)
        results.append(output_path)

    return results

# Usage
data_list = [
    {"name": "John", "score": 95},
    {"name": "Jane", "score": 87},
    {"name": "Bob", "score": 92}
]

batch_create_reports(app, "template.hwp", data_list, "output/")
```

## Color Reference

### Text Colors
```python
# Hex color format (BBGGRR to RRGGBB conversion handled by hwpapi)
app.set_charshape(color="#FF0000")  # Red
app.set_charshape(color="#00FF00")  # Green
app.set_charshape(color="#0000FF")  # Blue
app.set_charshape(color="#000000")  # Black
```

### Cell Background Colors
```python
app.cell.set_color("#FFFF00")  # Yellow
app.cell.set_color("#FFC0CB")  # Pink
```

## Font Reference

### Common Korean Fonts
```python
# Standard Korean fonts
app.set_charshape(fontname="맑은 고딕")  # Malgun Gothic
app.set_charshape(fontname="돋움")      # Dotum
app.set_charshape(fontname="굴림")      # Gulim
app.set_charshape(fontname="바탕")      # Batang

# English fonts
app.set_charshape(fontname="Arial")
app.set_charshape(fontname="Times New Roman")
app.set_charshape(fontname="Courier New")
```

### Font Size Reference
```python
# Common sizes (in points)
app.set_charshape(height=9)   # Small
app.set_charshape(height=11)  # Body text
app.set_charshape(height=14)  # Subheading
app.set_charshape(height=18)  # Heading
app.set_charshape(height=24)  # Title
```

## Common Workflows

### 1. Create Simple Document

```python
from hwpapi.core import App

app = App(new_app=True)

# Add title
app.set_charshape(fontname="맑은 고딕", height=24, bold=True)
app.insert_text("보고서\r\n")

# Add content
app.set_charshape(height=11, bold=False)
app.insert_text("이것은 보고서 내용입니다.\r\n")

app.save_as("report.hwp")
```

### 2. Create Document with Table

```python
app = App(new_app=True)

# Add title
app.set_charshape(height=18, bold=True)
app.insert_text("판매 현황\r\n")

# Create table
app.table.create(rows=4, cols=3)

# Fill header row
headers = ["품목", "수량", "금액"]
for col, text in enumerate(headers):
    app.cell.move(0, col)
    app.cell.text = text

# Fill data
data = [["상품A", "100", "1,000,000원"], ["상품B", "50", "500,000원"], ["합계", "150", "1,500,000원"]]
for row_idx, row_data in enumerate(data, start=1):
    for col_idx, text in enumerate(row_data):
        app.cell.move(row_idx, col_idx)
        app.cell.text = text

app.save_as("sales_report.hwp")
```

### 3. Fill Proposal Template

```python
def create_proposal(template_path, output_path, proposal_data):
    app = App()
    app.open(template_path)

    # Fill placeholders
    app.replace_all("{{proposal_title}}", proposal_data["title"])
    app.replace_all("{{client_name}}", proposal_data["client"])
    app.replace_all("{{project_period}}", proposal_data["period"])
    app.replace_all("{{total_amount}}", proposal_data["amount"])
    app.replace_all("{{submission_date}}", proposal_data["date"])

    # Fill table data if present
    if "items" in proposal_data:
        # Navigate to table and fill
        # ... (specific table filling logic)

    app.save_as(output_path)
    app.close()

# Usage
proposal = {
    "title": "시스템 구축 제안서",
    "client": "ABC 주식회사",
    "period": "2024.01 ~ 2024.06 (6개월)",
    "amount": "150,000,000원",
    "date": "2024년 1월 18일"
}

create_proposal("proposal_template.hwp", "output_proposal.hwp", proposal)
```

## Utility Functions

### Convert Units

```python
from hwpapi.functions import point2unit, unit2point, mili2unit, unit2mili

# Points to HWPUNIT (1pt = 100 HWPUNIT)
hwpunit = point2unit(12)  # 1200 HWPUNIT

# HWPUNIT to points
points = unit2point(1200)  # 12pt

# Millimeters to HWPUNIT (1mm = 283 HWPUNIT)
hwpunit = mili2unit(10)  # 2830 HWPUNIT

# HWPUNIT to millimeters
mm = unit2mili(2830)  # 10mm
```

### Get Font List

```python
fonts = app.get_font_list()
print("Available fonts:", fonts)
```

## Error Handling

```python
from hwpapi.core import App

try:
    app = App()
    app.open("document.hwp")
    # ... operations ...
except Exception as e:
    print(f"Error: {e}")
    # HWP may not be installed or running
finally:
    if 'app' in locals():
        app.close()
```

## Important Notes

1. **HWP Must Be Installed**: hwpapi requires Hancom Office HWP to be installed on Windows
2. **COM Automation**: Uses pywin32 for COM communication with HWP
3. **Running Instance**: By default, App() connects to running HWP or starts new
4. **Unit System**: HWP uses HWPUNIT (1mm = 283, 1pt = 100)
5. **Korean Text**: HWP handles Korean text natively
6. **Template Placeholders**: Use `{{placeholder}}` format for templates

## See Also

- [references/hwpapi_api_reference.md](references/hwpapi_api_reference.md) - Complete hwpapi API reference
- [references/action_reference.md](references/action_reference.md) - Common HWP actions
- [references/template_workflow.md](references/template_workflow.md) - Template-based workflow guide
