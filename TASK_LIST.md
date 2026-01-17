# HWP Skill Development Task List

## Overview
Create a Claude Code skill for working with HWP (Hangul Word Processor) documents using the hwpapi library.

## Requirements Summary

### Core Functionality
1. **문서 생성 및 관리** (Document Creation & Management)
   - Create new documents
   - Open existing HWP files
   - Save documents

2. **텍스트 편집** (Text Editing)
   - Insert text
   - Set font properties (font family, size, bold, italic, color)
   - Add paragraphs
   - Text selection and replacement

3. **테이블 작업** (Table Operations)
   - Create tables
   - Fill data into cells
   - Set cell content

4. **완성된 문서 생성** (Complete Document Generation)
   - Template-based report generation
   - Template-based letter generation

5. **일괄 작업** (Batch Operations)
   - Execute multiple operations at once

### Additional Requirements
6. **Template Handling**
   - Read and understand HWP proposal template files
   - Fill in content into proposal templates
   - Convert markdown tables to HWP tables
   - Insert images into files

7. **Document Formatting**
   - Font family and size changes
   - Bold, italic formatting
   - Paragraph formatting

## Skill Structure Planning

### 1. Scripts Directory (`scripts/`)

#### Core HWP Operations
- `hwp_create.py` - Create new HWP document
- `hwp_open.py` - Open existing HWP file
- `hwp_save.py` - Save HWP document
- `hwp_close.py` - Close HWP document

#### Text Operations
- `hwp_insert_text.py` - Insert text at cursor position
- `hwp_set_font.py` - Set font properties (family, size, bold, italic, color)
- `hwp_find_replace.py` - Find and replace text
- `hwp_select_text.py` - Select text for manipulation

#### Table Operations
- `hwp_create_table.py` - Create table with specified dimensions
- `hwp_fill_table.py` - Fill table with data
- `hwp_set_cell.py` - Set individual cell content

#### Template Operations
- `hwp_template_fill.py` - Fill template with data
- `hwp_markdown_to_hwp.py` - Convert markdown tables to HWP tables
- `hwp_insert_image.py` - Insert image into document
- `hwp_format_document.py` - Apply document formatting

#### Batch Operations
- `hwp_batch_process.py` - Execute multiple HWP operations

### 2. References Directory (`references/`)

#### API Documentation
- `hwpapi_api_reference.md` - Complete hwpapi API reference
- `action_reference.md` - Common HWP actions and their usage
- `property_reference.md` - Character and paragraph shape properties

#### Workflow Guides
- `template_workflow.md` - Template-based document creation workflow
- `table_workflow.md` - Table creation and manipulation workflow
- `formatting_workflow.md` - Document formatting workflow
- `batch_workflow.md` - Batch processing workflow

#### Examples
- `proposal_template_example.md` - Example proposal template structure
- `markdown_table_example.md` - Markdown table conversion examples

### 3. Assets Directory (`assets/`)

#### Templates
- `proposal_template.hwp` - Sample proposal template (to be added by user)
- `letter_template.hwp` - Sample letter template (to be added by user)

#### Resources
- `default_font_config.json` - Default font configuration
- `table_style_config.json` - Default table style configuration

## Implementation Steps

### Phase 1: Core Operations
- [ ] Create `scripts/hwp_create.py`
- [ ] Create `scripts/hwp_open.py`
- [ ] Create `scripts/hwp_save.py`
- [ ] Create `scripts/hwp_close.py`
- [ ] Test basic document operations

### Phase 2: Text Operations
- [ ] Create `scripts/hwp_insert_text.py`
- [ ] Create `scripts/hwp_set_font.py`
- [ ] Create `scripts/hwp_find_replace.py`
- [ ] Create `scripts/hwp_select_text.py`
- [ ] Test text manipulation

### Phase 3: Table Operations
- [ ] Create `scripts/hwp_create_table.py`
- [ ] Create `scripts/hwp_fill_table.py`
- [ ] Create `scripts/hwp_set_cell.py`
- [ ] Test table operations

### Phase 4: Template & Advanced Features
- [ ] Create `scripts/hwp_template_fill.py`
- [ ] Create `scripts/hwp_markdown_to_hwp.py`
- [ ] Create `scripts/hwp_insert_image.py`
- [ ] Create `scripts/hwp_format_document.py`
- [ ] Test template and advanced features

### Phase 5: Batch Operations
- [ ] Create `scripts/hwp_batch_process.py`
- [ ] Test batch processing

### Phase 6: Documentation
- [ ] Create `references/hwpapi_api_reference.md`
- [ ] Create `references/action_reference.md`
- [ ] Create `references/property_reference.md`
- [ ] Create workflow guides
- [ ] Create example documentation

### Phase 7: SKILL.md
- [ ] Write SKILL.md frontmatter
- [ ] Write skill body with usage instructions
- [ ] Add examples and workflows

### Phase 8: Package & Validate
- [ ] Run `package_skill.py` to validate
- [ ] Fix any validation errors
- [ ] Create final .skill package

## Key hwpapi API Reference

### App Class Methods
```python
from hwpapi.core import App

# Initialize
app = App()  # Opens HWP or connects to running instance
app = App(is_visible=False)  # Run in background

# Document operations
app.open(path)           # Open HWP file
app.save()               # Save current document
app.save_as(path)        # Save to specific path
app.close()              # Close document
app.quit()               # Quit HWP application

# Text operations
app.insert_text(text)    # Insert text at cursor
app.get_text()           # Get text from current position
app.get_selected_text()  # Get selected text
app.find_text(text)      # Find text
app.replace_all(old, new) # Replace all occurrences

# Formatting
app.set_charshape(fontname="맑은 고딕", height=12, bold=True, italic=False, color="#000000")
app.set_parashape(line_spacing=160, left_margin=0)

# File insertion
app.insert_picture(path) # Insert image
app.insert_file(path)    # Insert another HWP file

# Navigation (via move accessor)
app.move.top_of_file()
app.move.bottom()
app.move.page_start()
app.move.page_end()

# Table operations (via cell accessor)
app.cell.text = "Cell content"
app.cell.select()
app.cell.set_border()
app.cell.set_color()

# Table operations (via table accessor)
app.table.create(rows=5, cols=3)
```

### Actions
```python
# Access actions via app.actions
app.actions.InsertText.pset.Text = "Hello"
app.actions.InsertText.run()

# Character formatting
app.actions.CharShape.pset.Height = 1200  # Font size in HWPUNIT (100 = 1pt)
app.actions.CharShape.pset.FontNameHangul = "맑은 고딕"
app.actions.CharShape.pset.Bold = 1
app.actions.CharShape.run()

# Paragraph formatting
app.actions.ParaShape.pset.LineSpacing = 160  # 160%
app.actions.ParaShape.pset.LeftMargin = 1000
app.actions.ParaShape.run()

# Table creation
app.actions.TableCreate.pset.Rows = 5
app.actions.TableCreate.pset.Cols = 3
app.actions.TableCreate.pset.Width = 5000  # Table width in HWPUNIT
app.actions.TableCreate.run()
```

### Parameter Sets
Key parameter sets from hwpapi:
- `CharShape` - Character formatting (font, size, color, bold, italic)
- `ParaShape` - Paragraph formatting (alignment, spacing, margins)
- `TableCreation` - Table creation parameters (rows, cols, width)
- `BorderFill` - Border and fill settings
- `PageDef` - Page definition (size, margins)

## Design Decisions

### Script Design
- Each script should be standalone and executable
- Scripts should accept command-line arguments
- Scripts should return meaningful exit codes
- Scripts should log important operations

### Error Handling
- Handle HWP not installed
- Handle file not found
- Handle template placeholders not found
- Handle invalid data formats

### Template Placeholder Format
- Use `{{placeholder_name}}` format for placeholders
- Support nested placeholders
- Support conditional placeholders

### Markdown to HWP Table Conversion
- Parse markdown tables using standard syntax
- Preserve column alignments
- Handle merged cells if specified

## Testing Strategy
- Test each script individually
- Test with sample HWP files
- Test with Korean text
- Test with various font settings
- Test table creation and manipulation
- Test template filling

## Dependencies
- hwpapi (Python wrapper for HWP COM)
- Windows OS with Hancom Office HWP installed
- pywin32 (for COM automation)

## Notes
- hwpapi requires HWP to be installed on Windows
- DLL registration may be required for some operations
- HWP must be running or be able to start
- Consider running HWP in background for automation

## Progress Tracking
- [ ] Phase 1: Core Operations
- [ ] Phase 2: Text Operations
- [ ] Phase 3: Table Operations
- [ ] Phase 4: Template & Advanced Features
- [ ] Phase 5: Batch Operations
- [ ] Phase 6: Documentation
- [ ] Phase 7: SKILL.md
- [ ] Phase 8: Package & Validate
