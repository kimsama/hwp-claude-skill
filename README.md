# HWP Skill for Claude Code

A Claude Code skill for working with HWP (Hangul Word Processor) documents using the hwpapi library.

## What This Skill Does

This skill enables Claude Code to:

- **문서 생성 및 관리** - Create, open, and save HWP documents
- **텍스트 편집** - Insert text, set fonts, format paragraphs
- **테이블 작업** - Create tables, fill cell data
- **완성된 문서 생성** - Generate documents from templates
- **Markdown 변환** - Convert markdown tables to HWP tables
- **이미지 삽입** - Insert images into documents
- **일괄 작업** - Batch process multiple documents

## Prerequisites

- Windows OS with Hancom Office HWP installed
- Python 3.7+ with `hwpapi` installed:
  ```bash
  pip install hwpapi
  ```

## Installation

### Option 1: Install from .skill File

```bash
# Copy to Claude Code skills directory
cp hwp.skill ~/.config/claude-code/skills/

# Extract
cd ~/.config/claude-code/skills
unzip hwp.skill -d hwp
```

### Option 2: Install from Directory

```bash
# Copy entire directory
cp -r hwp ~/.config/claude-code/skills/
```

## Usage with Claude Code

Once installed, you can ask Claude to work with HWP documents directly:

### Basic Document Creation

```
Create a new HWP document with the title "월간 보고서" and save it as monthly_report.hwp
```

```
Create a document with a title (bold, 18pt) and add the following content:
[your content here]
Save as report.hwp
```

### Template Filling

```
Fill out the proposal_template.hwp file with these values:
- Company: ABC 주식회사
- Date: 2024년 1월 18일
- Amount: 5,000,000원
Save as filled_proposal.hwp
```

```
I have a JSON file with client data. Use the proposal_template.hwp to create proposals for all clients.
```

### Table Operations

```
Create a table with 5 rows and 3 columns in a new HWP document. Fill it with:
Header row: 품목, 수량, 금액
Data rows as shown in this data: [...]
```

```
Convert this markdown table to an HWP table:
| 제품 | 가격 | 수량 |
|------|------|------|
| A | 1000 | 5 |
| B | 2000 | 3 |
```

### Image Insertion

```
Insert logo.png at the top of document.hwp
```

```
Insert the chart image and resize it to 100mm width
```

### Document Formatting

```
Open report.hwp and:
1. Make all headings bold (14pt 맑은 고딕)
2. Set body text to 11pt
3. Add 1.5 line spacing
```

### Batch Processing

```
Create 10 reports from template.hwp using the data from clients.csv
```

## Available Scripts

The skill includes reusable Python scripts that Claude can import:

| Script | Purpose |
|--------|---------|
| `hwp_create_document.py` | Create new documents |
| `hwp_template_fill.py` | Fill template placeholders |
| `hwp_create_table.py` | Create tables with data |
| `hwp_markdown_table.py` | Convert markdown tables to HWP |
| `hwp_insert_image.py` | Insert images with sizing |

## Placeholder Format for Templates

Templates use `{{placeholder_name}}` format:

```
제안서

제목: {{title}}
고객사: {{client_name}}
작성일: {{date}}

프로젝트 기간: {{project_period}}
예산: {{total_amount}}원
```

## Example Workflow

### 1. Create a Proposal from Template

**Your prompt:**
```
Fill out proposal_template.hwp with:
- title: "시스템 구축 제안"
- client: "ABC Corporation"
- date: today's date in Korean format
- amount: 150,000,000원 formatted as Korean currency
```

**Claude will:**
1. Import `fill_template` from the skill
2. Load the template
3. Replace all placeholders
4. Save the filled document

### 2. Generate Multiple Reports

**Your prompt:**
```
Generate monthly reports for all employees in employees.csv
Use report_template.hwp and save to reports/ directory
```

**Claude will:**
1. Read the CSV file
2. Loop through each employee
3. Fill the template with their data
4. Save each report with a unique filename

### 3. Create Document with Table

**Your prompt:**
```
Create a sales report document with:
- Title: "1월 판매 현황" (18pt, bold)
- A table showing: Product, Quantity, Amount
- Data from this list: [...]
- Save as sales_report.hwp
```

## Template Examples

### Simple Proposal Template

```
제  안  서

{{title}}

1. 개요
본 {{project_type}} 프로젝트에 대한 제안서입니다.

2. 기간
{{project_period}}

3. 예산
총 예산: {{total_amount}}원

4. 담당자
{{manager_name}} ({{manager_contact}}})

작성일: {{date}}
```

### Report Template with Table Marker

```
{{report_title}}

작성일: {{date}}

{{report_content}}

판매 현황:
{{table_data}}
```

## Tips

1. **Use Korean naturally** - The skill handles Korean text natively
2. **Be specific about formatting** - "Make it bold, 14pt" is better than "Make it nice"
3. **Provide template structure** - Describe your template placeholders clearly
4. **Specify file paths** - Use absolute paths when possible
5. **Batch operations** - Claude can process multiple files at once

## Limitations

- Requires Windows with HWP installed
- HWP must be runnable (COM automation)
- Some advanced HWP features may require direct API calls

## Troubleshooting

**HWP doesn't start:**
- Ensure HWP is installed and accessible
- Try running HWP manually first

**Scripts not found:**
- Verify skill is in `~/.config/claude-code/skills/hwp/`
- Check scripts directory exists and contains .py files

**Placeholders not replaced:**
- Verify placeholder format: `{{name}}`
- Check for extra spaces or typos

## Reference Documentation

- `SKILL.md` - Complete skill documentation
- `references/hwpapi_api_reference.md` - Full hwpapi API reference
- `references/template_workflow.md` - Template workflow guide

## Contributing

To add new functionality:
1. Create a new script in `scripts/`
2. Add documentation to `SKILL.md`
3. Update this README with usage examples

## License

Same as hwpapi library.
