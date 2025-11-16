# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a collection of Python automation tools for insurance brokerage workflows. The tools process policy documents, generate renewal letters, manage Excel data, and automate document handling tasks. All scripts are designed to be packaged as standalone Windows executables using `auto-py-to-exe`.

## Setup and Installation

```bash
# Install dependencies
pip install -r requirements.txt

# Build standalone executables (launches GUI)
python -m auto_py_to_exe
```

## Core Architecture

### Entry Point System

The main automation dispatcher is `py/file_completion_tool.py`, which:
- Reads configuration from `config.xlsx` (in `py/` directory)
- Routes tasks based on the `event` field in cell B3
- Supported events: "manual renewal letter", "sort renewal list"

### Key Configuration File

`config.xlsx` structure (defined in `utils.py:CONFIG_FIELDS`):
- B3: Task/event type
- B7: Broker name
- B9: On behalf of
- B13: Risk type
- B15: Named insured
- B16: Insurer
- B17: Policy number
- B18: Effective date
- B20-B22: Mailing address lines
- B24: Risk address
- B27: Number of PDFs
- B29: Drive letter
- B33-B49: Producer range

### Shared Utilities (`utils.py`)

Common functions used across all tools:

**File Operations:**
- `safe_filename(name)`: Sanitizes filenames by removing invalid characters
- `unique_file_name(path)`: Generates unique filenames with (n) suffix pattern
- `load_excel_mapping(mapping_path, default_mappings, excel_mappings, sheet_name)`: Loads configuration from Excel

**PDF Processing:**
- `build_index(doc)`: Creates page/block/line index for PyMuPDF documents
- Returns `(page_index, text_to_location)` for efficient text searching

**Document Generation:**
- `write_to_new_docx(template_path, data, output_dir)`: Renders Word templates using `docxtpl`
- Template location: `assets/Renewal Letter.docx`
- Auto-generates filename from `named_insured` field
- Default output: Desktop

**Name Formatting:**
- `clean_name(name)`: Removes special characters, normalizes whitespace
- `format_name(name, lessor)`: Reformats names (handles corporate entities, lessors, individuals)

**UI:**
- `progressbar(iterable, prefix, size)`: Terminal progress bar with time estimates

### PDF Extraction System

#### Constants (`constants.py`)

The `RECTS` dictionary defines PDF field extraction rules for multiple insurance companies:
- **Policy Type Detection:** `RECTS['policy_type']` contains keyword + rect pairs for auto-detecting insurers
- **Field Mappings:** Each insurer (Aviva, Family, Intact, Wawanesa) has field extraction rules with:
  - `pattern`: Regex to find keywords
  - `rect`: `fitz.Rect` for text clipping (can be offset from pattern match)

#### PDFExtractor Class (`auto_renewal_letter.py`)

**Workflow:**
1. `detect_policy_type()`: Auto-detects insurer using `RECTS['policy_type']`
2. `extract_fields()`: Extracts all fields for detected insurer
3. `find_by_keyword(page, pattern, rect)`: Locates text using pattern + optional rect offset

**Pattern + Rect Logic:**
- If `pattern` only: Returns text in keyword bounding box
- If `pattern` + `rect`: Returns text from `keyword_position + rect_offset`
- If `rect` only: Direct extraction from absolute coordinates

### Main Tools

#### 1. Auto Renewal Letter (`auto_renewal_letter.py`)
- Processes all PDFs in `~/Downloads`
- Auto-detects insurer type using RECTS
- Extracts policy details from declarations
- Generates renewal letters to `~/Desktop`
- Uses template: `assets/Renewal Letter.docx`

**Data Flow:**
```
PDF → PDFExtractor.auto_from_pdf() → map_extracted_data_for_renewal() → write_to_new_docx()
```

#### 2. Manual Renewal Letter (`manual_renewal_letter.py`)
- Uses data from `config.xlsx` instead of PDF extraction
- Triggered when config.xlsx event = "manual renewal letter"
- Maps config fields using `constants.py:EXCEL_CELL_MAPPING`
- Fallback: Uses mailing address for risk address if empty

#### 3. Sort Renewal List (`sort_renewal_list.py`)
- Processes the 2 most recent Excel files in `~/Downloads`
- Combines data, removes duplicates by `policynum`
- Sorts by: insurer → renewal date → name
- Adds blank rows between insurers
- Creates formatted table with borders on Pulled/D/L columns
- Output: `~/Desktop/renewal_list.xlsx`

**Column Order:**
`policynum, ccode, name, pcode, csrcode, insurer, buscode, renewal, Pulled, D/L`

#### 4. Reconciler Tool (`reconciller_tool.py`)
- Compares PDF tables to find matching policy numbers and premiums
- Auto-detects policy and premium columns in PDF tables
- Uses pattern matching: `^[A-Z0-9]{6,}$` for policy numbers
- Input: `../input/` folder
- Output: `../output/` folder

## Development Patterns

### Working with PDFs
- All PDF operations use PyMuPDF (`fitz`)
- Text extraction uses `page.get_text("text", clip=rect)` for precise regions
- `page.search_for(pattern)` returns list of `fitz.Rect` for keyword locations
- When adding new insurers: Update `RECTS` in `constants.py` with keyword + rect definitions

### Adding New Automation Tasks
1. Create new script in `py/` directory
2. Add event trigger to `file_completion_tool.py:main()`
3. Update `config.xlsx` if new configuration fields needed
4. Use shared utilities from `utils.py` for file operations

### Date Handling
- Standard format: `"%B %d, %Y"` (e.g., "January 01, 2025")
- Always parse with fallback formats: `%Y-%m-%d`, `%m/%d/%Y`
- Store raw datetime objects until final rendering

### Executable Packaging
- Use PyInstaller via `auto-py-to-exe` GUI
- Scripts must handle paths relative to executable location
- Test with `Path.home()` for Desktop/Downloads access
- Include all assets in build (Word templates, etc.)

## File Organization

```
Python-Automations/
├── py/                          # All Python source files
│   ├── file_completion_tool.py  # Main dispatcher
│   ├── auto_renewal_letter.py   # PDF-based renewal letters
│   ├── manual_renewal_letter.py # Config-based renewal letters
│   ├── sort_renewal_list.py     # Excel list processor
│   ├── reconciller_tool.py      # PDF table comparison
│   ├── utils.py                 # Shared utilities
│   ├── constants.py             # RECTS and mappings
│   ├── config.xlsx              # Configuration file
│   └── debug.py                 # Testing/debugging
├── assets/                      # Templates (if exists)
│   └── Renewal Letter.docx      # Word template
└── requirements.txt             # Dependencies
```

## Common Input/Output Paths

- **Input PDFs:** `Path.home() / "Downloads"`
- **Input Excel:** `Path.home() / "Downloads"` (for sort_renewal_list)
- **Output Documents:** `Path.home() / "Desktop"`
- **Template:** `Path.cwd() / "assets" / "Renewal Letter.docx"`

## Testing

- Each script has `if __name__ == "__main__"` block with test data
- `debug.py` for development testing
- Test with sample PDFs from each insurer before building executables
