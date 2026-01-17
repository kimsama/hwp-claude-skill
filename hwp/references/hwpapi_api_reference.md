# hwpapi API Reference

Complete reference for the hwpapi Python library for HWP (Hangul Word Processor) automation.

## Core Classes

### App

The main entry point for HWP automation.

```python
from hwpapi.core import App

# Create or connect to HWP
app = App()                    # Connect to running or start new
app = App(new_app=True)        # Always start new instance
app = App(is_visible=False)    # Run in background
```

#### App Properties

| Property | Type | Description |
|----------|------|-------------|
| `api` | COM object | Direct access to HWP COM object |
| `actions` | _Actions | Collection of all available actions |
| `parameters` | HParameterSet | Parameter set collection |
| `move` | MoveAccessor | Navigation accessor |
| `cell` | CellAccessor | Cell operations accessor |
| `table` | TableAccessor | Table operations accessor |
| `page` | PageAccessor | Page operations accessor |
| `engine` | Engine | Underlying engine object |

#### App Methods

##### Document Operations

```python
app.open(path)              # Open HWP file
app.save()                 # Save current document
app.save_as(path)          # Save to specific path
app.close()                # Close document (don't quit HWP)
app.quit()                 # Quit HWP application
app.reload()               # Reload HWP object
```

##### Text Operations

```python
app.insert_text(text)      # Insert text at cursor
app.get_text()             # Get text from current position
app.get_selected_text()    # Get selected text
app.find_text(text)        # Find text, returns True/False
app.replace_all(old, new)  # Replace all occurrences
app.select_text(option)    # Select text (SelectionOption enum)
```

##### Formatting

```python
# Character formatting
app.set_charshape(
    fontname="맑은 고딕",    # Font family
    height=12,               # Font size in points
    bold=True,               # Bold
    italic=False,            # Italic
    underline=False,         # Underline
    color="#000000"          # Color (hex)
)

# Paragraph formatting
app.set_parashape(
    line_spacing=160,        # Line spacing (%)
    left_margin=0,           # Left margin
    right_margin=0,          # Right margin
    indent=0                 # Indent
)

# Get current formatting
char_shape = app.get_charshape()
para_shape = app.get_parashape()
```

##### File Insertion

```python
app.insert_picture(path)    # Insert image
app.insert_file(path)       # Insert another HWP file
```

##### Utility

```python
app.get_filepath()          # Get current file path
app.get_font_list()         # Get list of available fonts
app.get_hwnd()              # Get window handle
app.set_visible(bool)       # Set window visibility
app.scan()                  # Scan document
app.save_block()            # Save block
```

### Accessors

#### MoveAccessor (app.move)

Navigation within the document.

```python
# Document navigation
app.move.top_of_file()
app.move.bottom()
app.move.screen_top()
app.move.screen_bottom()

# Page navigation
app.move.page_start()
app.move.page_end()
app.move.next_page()
app.move.prev_page()

# List/paragraph navigation
app.move.current_list(para=0, pos=0)  # Move to specific paragraph
app.move.next_list()
app.move.prev_list()
```

#### CellAccessor (app.cell)

Table cell operations.

```python
# Move to cell
app.cell.move(row, col)

# Get/set cell text
app.cell.text = "Content"
text = app.cell.text

# Cell formatting
app.cell.set_border(type="solid", width=1)
app.cell.set_color("#FFFF00")

# Selection
app.cell.select()
```

#### TableAccessor (app.table)

Table operations.

```python
# Create table
app.table.create(rows=5, cols=3)
app.table.create(rows=5, cols=3, width=15000)  # Width in HWPUNIT

# Get table info
row_count = app.table.row_count
col_count = app.table.col_count
```

#### PageAccessor (app.page)

Page operations.

```python
# Navigation
app.page.next()
app.page.prev()

# Get page info
current_page = app.page.current
total_pages = app.page.count
```

## Actions

Actions are accessed via `app.actions.ActionName`.

### Common Actions

#### InsertText

```python
action = app.actions.InsertText
action.pset.Text = "Hello World"
action.run()
```

#### CharShape (Character Formatting)

```python
action = app.actions.CharShape
action.pset.Height = 1200          # Font size (100 = 1pt)
action.pset.FontNameHangul = "맑은 고딕"
action.pset.FontNameLatin = "Arial"
action.pset.Bold = 1               # 0=off, 1=on
action.pset.Italic = 1
action.pset.Underline = 1
action.pset.TextColor = 0           # BBGGRR format
action.run()
```

#### ParaShape (Paragraph Formatting)

```python
action = app.actions.ParaShape
action.pset.LineSpacing = 160       # 160%
action.pset.LeftMargin = 1000       # In HWPUNIT
action.pset.RightMargin = 1000
action.pset.Indent = 500
action.pset.Align = 0               # 0=left, 1=center, 2=right, 3=justify
action.run()
```

#### TableCreate

```python
action = app.actions.TableCreate
action.pset.Rows = 5
action.pset.Cols = 3
action.pset.Width = 15000           # Table width in HWPUNIT
action.pset.TableSizeMode = 0       # 0=fixed, 1=auto
action.pset.AutoCreate = 1          # Auto-create table
action.run()
```

#### InsertPicture

```python
action = app.actions.InsertPicture
action.pset.FileName = "path/to/image.png"
action.pset.Width = 5000            # Width in HWPUNIT
action.pset.Height = 3000           # Height in HWPUNIT
action.pset.SizeManipulate = 1      # Maintain aspect ratio
action.run()
```

#### FindReplace

```python
action = app.actions.FindReplace
action.pset.FindString = "old text"
action.pset.ReplaceString = "new text"
action.pset.Direction = 0           # 0=down, 1=up
action.pset.FindType = 1            # 1=exact match
action.pset.Replace = 1             # 1=replace, 0=find only
action.run()
```

#### PageSetup

```python
action = app.actions.PageSetup
action.pset.PaperWidth = 21000      # 210mm in HWPUNIT
action.pset.PaperHeight = 29700     # 297mm (A4)
action.run()
```

## Parameter Sets

Parameter sets configure action behavior.

### Common Parameter Sets

#### CharShape
- `Height` - Font size in HWPUNIT (100 = 1pt)
- `FontNameHangul` - Korean font name
- `FontNameLatin` - Latin/English font name
- `Bold` - Bold (0/1)
- `Italic` - Italic (0/1)
- `Underline` - Underline (0/1)
- `TextColor` - Text color in BBGGRR format

#### ParaShape
- `LineSpacing` - Line spacing percentage
- `LeftMargin` - Left margin in HWPUNIT
- `RightMargin` - Right margin in HWPUNIT
- `Indent` - First line indent in HWPUNIT
- `Align` - Alignment (0=left, 1=center, 2=right, 3=justify)

#### TableCreation
- `Rows` - Number of rows
- `Cols` - Number of columns
- `Width` - Table width in HWPUNIT
- `TableSizeMode` - 0=fixed width, 1=auto width
- `AutoCreate` - Auto-create table (0/1)

## Constants

### Alignment

```python
from hwpapi.constants import Alignment

Alignment.LEFT      # 0
Alignment.CENTER    # 1
Alignment.RIGHT     # 2
Alignment.JUSTIFY   # 3
```

### SelectionOption

```python
from hwpapi.constants import SelectionOption

SelectionOption.LINE      # Current line
SelectionOption.PARA      # Current paragraph
SelectionOption.DOCUMENT  # Entire document
```

## Utility Functions

### Unit Conversion

```python
from hwpapi.functions import point2unit, unit2point, mili2unit, unit2mili

# Points to HWPUNIT (1pt = 100 HWPUNIT)
hwpunit = point2unit(12)    # 1200

# HWPUNIT to points
points = unit2point(1200)   # 12

# Millimeters to HWPUNIT (1mm = 283 HWPUNIT)
hwpunit = mili2unit(10)     # 2830

# HWPUNIT to millimeters
mm = unit2mili(2830)        # 10
```

### Color Conversion

```python
from hwpapi.functions import hex_to_rgb, get_rgb_tuple

# Hex to RGB tuple
rgb = get_rgb_tuple("#FF0000")  # (255, 0, 0)

# RGB to hex
hex_color = hex_to_rgb((255, 0, 0))
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
finally:
    if 'app' in locals():
        app.close()
```

## Common Issues

### HWP Not Found
- Ensure HWP is installed
- Check HWP is in PATH

### COM Errors
- Run as administrator if needed
- Check HWP version compatibility

### File Path Issues
- Use absolute paths
- Ensure paths use forward slashes or escaped backslashes

## Notes

- All file paths should be absolute or properly escaped
- HWP must be installed on Windows
- Korean text is handled natively
- 1mm = 283 HWPUNIT, 1pt = 100 HWPUNIT
