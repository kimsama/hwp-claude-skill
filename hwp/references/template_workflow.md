# Template-Based Document Workflow

Complete workflow for creating documents from HWP templates.

## Overview

Template-based workflow allows you to:
1. Create HWP templates with placeholders
2. Fill templates with data programmatically
3. Generate multiple documents from a single template

## Placeholder Format

Use `{{placeholder_name}}` format for placeholders in your HWP templates.

### Example Template Content

```
제안서

제목: {{title}}
고객사: {{client_name}}
작성일: {{date}}

1. 개요
본 제안서는 {{project_name}} 프로젝트에 대한 제안 내용입니다.

2. 프로젝트 기간
{{project_period}}

3. 예산
총 예산: {{total_amount}}원

4. 담당자
이름: {{manager_name}}
연락처: {{manager_contact}}
```

## Step-by-Step Workflow

### Step 1: Create Template

1. Open HWP and create your document structure
2. Insert placeholders in `{{placeholder_name}}` format
3. Save as template (e.g., `proposal_template.hwp`)

### Step 2: Prepare Data

Prepare your data as a Python dictionary:

```python
data = {
    "title": "시스템 구축 제안서",
    "client_name": "ABC 주식회사",
    "date": "2024년 1월 18일",
    "project_name": "ERP 시스템 구축",
    "project_period": "2024.01 ~ 2024.06 (6개월)",
    "total_amount": "150,000,000",
    "manager_name": "홍길동",
    "manager_contact": "02-1234-5678"
}
```

### Step 3: Fill Template

```python
from hwpapi.core import App

def fill_template(template_path, output_path, data):
    """Fill template with data."""
    app = App(is_visible=False)  # Run in background

    # Open template
    app.open(template_path)

    # Replace all placeholders
    for placeholder, value in data.items():
        pattern = f"{{{{{placeholder}}}}}"
        app.replace_all(pattern, str(value))
        print(f"Replaced: {pattern}")

    # Save filled document
    app.save_as(output_path)
    app.close()

    print(f"Created: {output_path}")

# Usage
fill_template(
    "proposal_template.hwp",
    "output_proposal.hwp",
    data
)
```

## Batch Processing

### Multiple Documents from One Template

```python
from hwpapi.core import App
from pathlib import Path

def batch_fill_templates(template_path, data_list, output_dir):
    """Create multiple documents from template."""
    output_dir = Path(output_dir)
    output_dir.mkdir(exist_ok=True)

    app = App(is_visible=False)

    for i, data in enumerate(data_list):
        # Open template
        app.open(template_path)

        # Fill placeholders
        for placeholder, value in data.items():
            pattern = f"{{{{{placeholder}}}}}"
            app.replace_all(pattern, str(value))

        # Save with unique name
        output_path = output_dir / f"document_{i+1}.hwp"
        app.save_as(str(output_path))
        print(f"Created: {output_path}")

    app.close()
    print(f"Created {len(data_list)} documents")

# Usage
data_list = [
    {"name": "John", "score": 95},
    {"name": "Jane", "score": 87},
    {"name": "Bob", "score": 92}
]

batch_fill_templates("report_template.hwp", data_list, "output/")
```

## Working with Tables in Templates

### Method 1: Fill After Template

```python
from hwpapi.core import App

def fill_template_with_table(template_path, output_path, data, table_data):
    """Fill template and add table data."""
    app = App(is_visible=False)

    # Open and fill template
    app.open(template_path)

    for placeholder, value in data.items():
        app.replace_all(f"{{{{{placeholder}}}}}", str(value))

    # Navigate to table position (user marks this)
    # For example, find "{{table_data}}" marker
    app.find_text("{{table_data}}")
    app.replace_all("{{table_data}}", "")  # Remove marker

    # Create/fill table
    rows = len(table_data)
    cols = len(table_data[0])
    app.table.create(rows=rows, cols=cols)

    for row_idx, row in enumerate(table_data):
        for col_idx, cell_value in enumerate(row):
            app.cell.move(row_idx, col_idx)
            app.cell.text = str(cell_value)

    app.save_as(output_path)
    app.close()

# Usage
data = {
    "title": "월간 판매 보고서",
    "month": "2024년 1월"
}

table_data = [
    ["품목", "수량", "금액"],
    ["상품A", "100", "1,000,000원"],
    ["상품B", "50", "500,000원"]
]

fill_template_with_table("template.hwp", "output.hwp", data, table_data)
```

### Method 2: Pre-Created Table in Template

If template has a table with placeholders in cells:

```python
def fill_template_table_cells(template_path, output_path, table_data):
    """Fill pre-existing table in template."""
    app = App(is_visible=False)
    app.open(template_path)

    # Navigate to table (first table in document)
    app.move.top_of_file()
    # User should place cursor at table start or use bookmark

    # Fill table cells
    for row_idx, row in enumerate(table_data):
        for col_idx, cell_value in enumerate(row):
            app.cell.move(row_idx, col_idx)
            # Replace placeholder in cell if exists
            if "{{" in app.cell.text:
                app.cell.text = str(cell_value)

    app.save_as(output_path)
    app.close()
```

## Conditional Content

```python
def fill_template_with_conditionals(template_path, output_path, data):
    """Fill template with conditional content."""
    app = App(is_visible=False)
    app.open(template_path)

    for placeholder, value in data.items():
        pattern = f"{{{{{placeholder}}}}}"

        # Handle boolean/conditional values
        if value is None or value == "":
            # Remove placeholder and optional surrounding text
            app.replace_all(pattern, "")
        else:
            app.replace_all(pattern, str(value))

    # Remove unused sections
    app.replace_all("{{if:.*?}}.*?{{endif}}", "", regex=True)

    app.save_as(output_path)
    app.close()
```

## Number Formatting

```python
def format_number(value, style="korean"):
    """Format numbers for Korean documents."""
    if style == "korean":
        # 1,000,000원
        return f"{value:,}원"
    elif style == "decimal":
        return f"{value:,.2f}"
    else:
        return str(value)

# Usage
data = {
    "amount": format_number(150000000, "korean"),  # "150,000,000원"
    "rate": format_number(0.0525, "decimal")       # "0.53"
}
```

## Date Formatting

```python
from datetime import datetime

def format_korean_date(date_value):
    """Format date in Korean style."""
    if isinstance(date_value, str):
        # Parse if needed
        date_value = datetime.strptime(date_value, "%Y-%m-%d")
    return date_value.strftime("%Y년 %m월 %d일")

# Usage
data = {
    "date": format_korean_date(datetime.now())
    # Output: "2024년 01월 18일"
}
```

## Error Handling

```python
from hwpapi.core import App
import logging

def safe_fill_template(template_path, output_path, data):
    """Fill template with error handling."""
    try:
        app = App(is_visible=False)

        # Check template exists
        if not Path(template_path).exists():
            raise FileNotFoundError(f"Template not found: {template_path}")

        app.open(template_path)

        replaced_count = 0
        for placeholder, value in data.items():
            pattern = f"{{{{{placeholder}}}}}"
            try:
                app.replace_all(pattern, str(value))
                replaced_count += 1
            except Exception as e:
                logging.warning(f"Failed to replace {pattern}: {e}")

        app.save_as(output_path)
        app.close()

        print(f"Success: {replaced_count} placeholders replaced")
        return True

    except Exception as e:
        logging.error(f"Template filling failed: {e}")
        return False
```

## Best Practices

1. **Use Clear Placeholder Names**
   - Good: `{{client_name}}`, `{{project_amount}}`
   - Bad: `{{name}}`, `{{amount}}` (too generic)

2. **Document Placeholders**
   - Keep a list of all placeholders in your template
   - Add comments in HWP describing each placeholder

3. **Validate Data**
   - Check all required keys are present
   - Validate data types before filling

4. **Test Templates**
   - Create a test data set
   - Verify all placeholders are replaced

5. **Backup Templates**
   - Keep original templates safe
   - Use copy for filling

## Complete Example

```python
#!/usr/bin/env python3
"""
Proposal Generator
Generate proposals from template with client data.
"""

from hwpapi.core import App
from pathlib import Path
from datetime import datetime
import json


def load_data(json_path):
    """Load client data from JSON."""
    with open(json_path, 'r', encoding='utf-8') as f:
        return json.load(f)


def format_proposal_data(data):
    """Format and validate proposal data."""
    # Add formatted fields
    data['formatted_date'] = format_korean_date(datetime.now())
    data['formatted_amount'] = format_number(data['amount'], "korean")
    return data


def generate_proposal(template_path, output_dir, client_data):
    """Generate proposal for a client."""
    output_dir = Path(output_dir)
    output_dir.mkdir(exist_ok=True)

    # Format data
    data = format_proposal_data(client_data)

    # Generate filename
    filename = f"proposal_{data['client_name']}_{datetime.now().strftime('%Y%m%d')}.hwp"
    output_path = output_dir / filename

    # Fill template
    app = App(is_visible=False)
    app.open(template_path)

    for placeholder, value in data.items():
        pattern = f"{{{{{placeholder}}}}}"
        app.replace_all(pattern, str(value))

    app.save_as(str(output_path))
    app.close()

    print(f"Generated: {output_path}")
    return output_path


def main():
    template_path = "templates/proposal_template.hwp"
    output_dir = "output/proposals"
    data_path = "data/clients.json"

    # Load client data
    clients = load_data(data_path)

    # Generate proposals
    for client in clients:
        generate_proposal(template_path, output_dir, client)


if __name__ == "__main__":
    main()
```

## Tips

1. Run HWP invisible for batch processing (`is_visible=False`)
2. Use absolute paths for templates and outputs
3. Handle Korean encoding properly (use utf-8)
4. Test with small batches before large runs
5. Keep templates simple and well-structured
