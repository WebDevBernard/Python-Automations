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

### Dependencies (requirements.txt)
- `pyinstaller`: Package Python scripts as standalone executables
- `auto-py-to-exe`: GUI for PyInstaller configuration
- `openpyxl`: Read/write Excel files (.xlsx format)
- `docxtpl`: Render Word templates with Jinja2 syntax
- `pymupdf` (fitz): PDF parsing, text extraction, table detection
- `pymupdf-fonts`: Font support for PyMuPDF
- `xlrd`: Read legacy Excel files (.xls format)
- `tabulate`: Format tabular data (used in reconciler and debug tools)
- `pandas`: Data manipulation (used in reconciler and sort tools)
- `pikepdf`: PDF permission/signature removal (used by `unlocked.py`) — **not yet in requirements.txt**

## Core Architecture

### Entry Point System

The main automation dispatcher is `py/file_completion_tool.py`, which:
- Reads configuration from `config.xlsx` (in `py/` directory)
- Routes tasks based on the `event` field in cell B3
- Supported events: "manual renewal letter", "auto renewal letter", "sort renewal list", "reconciller"

### Key Configuration File

`config.xlsx` structure (defined in `constants.py:EXCEL_CELL_MAPPING`):
- B3: Task/event type
- B7: Broker name
- B9: On behalf of
- B13: Risk type
- B15: Named insured (maps to `insured_name`)
- B16: Insurer
- B17: Policy number
- B18: Effective date
- B20-B22: Mailing address lines (street, city/province, postal)
- B24: Risk address

### Shared Utilities (`utils.py`)

Common functions used across all tools:

**File Operations:**
- `safe_filename(name)`: Sanitizes filenames by removing invalid characters
- `unique_file_name(path)`: Generates unique filenames with (n) suffix pattern
- `load_excel_mapping(mapping_path, default_mappings, excel_mappings, sheet_name)`: Loads configuration from Excel

**Document Generation:**
- `write_to_new_docx(template_path=None, data=None, output_dir=None)`: Renders Word templates using `docxtpl`
- Auto-detects first .docx file in `assets/` folder
- Returns `True` on success, `False` if template not found
- Auto-generates output filename from `named_insured` field
- Default output: Desktop

**PDF Processing:**
- `build_index(doc)`: Creates page/block/line index for PyMuPDF documents
- Returns `(page_index, text_to_location)` for efficient text searching

**UI:**
- `progressbar(iterable, prefix, size)`: Terminal progress bar with time estimates

### PDF Extraction System

#### Constants (`constants.py`)

The `RECTS` dictionary defines PDF field extraction rules for multiple insurance companies:
- **Policy Type Detection:** `POLICY_TYPE_DETECTION` contains keyword + rect pairs for auto-detecting insurers
- **Field Mappings:** Each insurer (Aviva, Family, Intact, Wawanesa) has field extraction rules with:
  - `pattern`: Regex to find keywords (optional)
  - `rect`: `fitz.Rect` for text clipping (can be offset from pattern match or absolute coordinates)
  - `return_all`: Boolean flag to extract all matches instead of just first match

**Helper Functions for Field Definitions:**
- `absolute_rect_field(x0, y0, x1, y1)`: Extract from absolute coordinates
- `pattern_only_field(regex_str, return_all=False)`: Extract by pattern matching only
- `pattern_with_offset_field(regex_str, dx0, dy0, dx1, dy1, return_all=False)`: Pattern + relative offset

**Regex Patterns (`REGEX_PATTERNS`):**
- `postal_code`: Canadian postal code format (e.g., V6B 1A1)
- `dollar`: Extracts dollar amounts with $ prefix
- `date`: Matches various date formats including "DD Mon YYYY" and "Month DD, YYYY"
- `address`: Matches street addresses (excludes "Ltd.")
- `and`: Matches "&" or word "and" (case-insensitive)

#### PDF Extraction Functions (`auto_renewal_letter.py`)

**Core Extraction Logic:**
1. `detect_insurer(doc, rects)`: Auto-detects insurer using keyword matching in `POLICY_TYPE_DETECTION`
2. `extract_fields(doc, field_mapping, insurer)`: Extracts all fields for detected insurer
3. `extract_single_field(pages_dict, cfg, doc)`: Handles extraction based on field config (pattern only, rect only, or pattern + offset)

**Extraction Patterns:**
- If `pattern` only: Returns text in keyword bounding box
- If `pattern` + `rect`: Returns text from `keyword_position + rect_offset`
- If `rect` only: Direct extraction from absolute coordinates
- If `return_all=True`: Returns list of all matches instead of first match

## Main Tools

### 1. Auto Renewal Letter (`auto_renewal_letter.py`)
- Processes the **2 most recent** PDFs in `~/Downloads`
- Auto-detects insurer type using `detect_insurer()`
- Skips Wawanesa "Statement of Account" PDFs (detected via `wawanesa_statement` field)
- Extracts policy details and generates renewal letters to `~/Desktop`
- Uses template: first `.docx` file found in `assets/`

**Data Flow:**
```
PDF → detect_insurer() → extract_fields() → format_fields() → check_glass_policy() → write_to_new_docx()
```

**Glass Policy Matching:**
- Loads glass policy data from Excel files in `assets/` folder (any `.xls` or `.xlsx`)
- For "home" risk types only: matches by postal code + renewal month/day against REL insurer rows
- When a match is found: adds `glass_policynum` field and sums the glass premium into `premium_amount`
- Warns if glass policy Excel files are older than 1 year
- Removes duplicate policy numbers from the glass dataset (keeps only uniquely-occurring policy numbers)

**Key Functions:**
- `get_text(doc, structured=True)`: Extracts text from PDF as pages → blocks
- `search_text(pages_dict, regex, return_all=False)`: Search regex across all blocks
- `deduplicate_field(value)`: Removes duplicates from multi-value fields (Intact-specific)
- `format_fields(raw_data, insurer)`: Formats all extracted fields with insurer-specific logic
- `address_one_title_case(sentence)`: Title case with ordinal numbers in lowercase
- `address_two_title_case(strings_list)`: Title case with province codes uppercased
- `check_glass_policy(fields, glass_policies)`: Matches and merges glass policy data

### 2. Manual Renewal Letter (`manual_renewal_letter.py`)
- Uses data from `config.xlsx` instead of PDF extraction
- Triggered when config.xlsx event = "manual renewal letter"
- Maps config fields via `map_config_for_renewal()` using keys from `EXCEL_CELL_MAPPING`
- Fallback: Uses mailing address for risk address if empty
- Handles date parsing with multiple fallback formats and NaN handling
- Output format standardized to `DATE_FORMAT = "%B %d, %Y"`

### 3. Sort Renewal List (`sort_renewal_list.py`)
- Processes all Excel files in `~/Downloads`
- Combines data, removes duplicates by `policynum` (keep=False — removes all instances of duplicates)
- Sorts by: insurer → renewal date → name
- Adds blank rows between insurers for visual separation
- Creates formatted Excel table with TableStyleLight1, borders on Pulled/D/L columns, and page setup for printing
- Output: `~/Desktop/renewal_list.xlsx`

**Column Order:**
`policynum, ccode, name, pcode, csrcode, insurer, buscode, renewal, Pulled, D/L`

### 4. Reconciler (`reconciller.py`)
- Compares PDF tables to find matching policy numbers and premiums across multiple PDFs
- Auto-detects policy columns (pattern: `^[A-Z0-9]{6,}`) and premium columns (numeric `.00` values with highest max)
- Highlights unbalanced (non-matching) policies in output PDFs
- Has Intact-specific policy number cleaning (strips spaces, handles letter prefixes, removes trailing "H")
- **Input:** `~/Downloads` (default, or configurable via `config_data`)
- **Output:** `~/Desktop` (default, or configurable via `config_data`)

### 5. PDF Unlocker (`unlocked.py`)
- Removes password protection, signature fields, and XFA forms from PDFs using `pikepdf`
- Reads all PDFs from `~/Downloads`, outputs to `~/Downloads/unlocked/`
- Strips `/Perms`, `/NeedsRendering`, `/OpenAction`, `/AA`, `/Names` from PDF root
- Removes signature fields and XFA data from AcroForms
- Adds AcroForm date formatting JavaScript for date widget fields that lost XFA validation

### 6. PDF Redaction (`redact.py`)
- Redacts specific words and their adjacent dollar/percentage values from PDFs
- Reads target words from `input/config.txt` (comma-separated)
- Only redacts if the adjacent text is a monetary value or percentage (validated via regex)
- **Input:** `input/` folder
- **Output:** `output/` folder

### 7. Debug Tool (`debug.py`)
- Development utility for PDF reverse-engineering with toggleable features:
  - Text extraction (blocks mode)
  - Table extraction with cell coordinates
  - Text search with name extraction
  - Region highlighting (saves page as PNG)
  - **Offset calculation:** Computes dx0/dy0/dx1/dy1 between a pattern rect and target rect, outputting values ready to paste into `constants.py` field definitions. Supports above/below/right directions.
- All features controlled via `CONFIG` dict at the top of the file

## Development Patterns

### Working with PDFs
- All PDF operations use PyMuPDF (`fitz`)
- Text extraction methods:
  - `page.get_text("text", clip=rect)`: Extract text from specific region
  - `page.get_text("blocks")`: Returns blocks with coordinates `(x0, y0, x1, y1, text, block_num, block_type)`
  - `page.get_textbox(rect)`: Extract text within a rect
  - `page.find_tables(strategy="text")`: Extract tables as pandas DataFrames
- `page.search_for(pattern)` returns list of `fitz.Rect` for keyword locations

### Adding New Insurers to Auto Renewal Letter
1. Add detection entry to `POLICY_TYPE_DETECTION` with unique keyword and rect
2. Create field mapping dictionary (e.g., `NEWINSURER_FIELDS`) using helper functions:
   - Use `absolute_rect_field()` for fixed positions (e.g., name/address blocks)
   - Use `pattern_only_field()` for regex-based extraction
   - Use `pattern_with_offset_field()` for relative positioning from keywords
   - Add `return_all=True` for fields that can have multiple values
3. Add insurer to `RECTS` dictionary
4. Test with sample PDF from the insurer — use `debug.py` offset calculations to find correct rect positions

### Adding New Automation Tasks
1. Create new script in `py/` directory with a main function that accepts `config_data` dict
2. Add event trigger to `file_completion_tool.py:main()`:
   ```python
   elif event == "new task name":
       new_task_function(config_data)
   ```
3. Update `config.xlsx` if new configuration fields needed:
   - Add new cells to column B
   - Update `DEFAULT_MAPPING` and `EXCEL_CELL_MAPPING` in `constants.py`
4. Use shared utilities from `utils.py` for file operations
5. Follow pattern: Print success message at end, use 3-second countdown before exit

### Date Handling
- Standard output format: `DATE_FORMAT = "%B %d, %Y"` (e.g., "January 01, 2025")
- Always parse with fallback formats: `%Y-%m-%d`, `%m/%d/%Y`, and handle datetime objects
- Store raw datetime objects until final rendering
- For sorting: Convert to MMDD format using `strftime("%m%d")`

### Executable Packaging
- Use PyInstaller via `auto-py-to-exe` GUI: `python -m auto_py_to_exe`
- Scripts must handle paths relative to executable location using `Path.cwd()`
- Use `Path.home()` for Desktop/Downloads access (not hardcoded paths)
- Include all assets in build (Word templates, icons, etc.)
- Assets are located in `assets/` directory relative to working directory

### Template Variables (docxtpl)
The Word template uses Jinja2 syntax for variable substitution. Fields passed to the template include: `named_insured`, `broker_name`, `on_behalf`, `policy_number`, `effective_date`, `address_line_one`, `address_line_two`, `address_line_three`, `risk_address_1`, `insurer`, `today`, and optional fields like `premium_amount`, `glass_policynum`, `number_of_families_1`, etc.

## File Organization

```
Python-Automations/
├── py/                          # All Python source files (working directory)
│   ├── file_completion_tool.py  # Main dispatcher (entry point)
│   ├── auto_renewal_letter.py   # PDF-based renewal letters
│   ├── manual_renewal_letter.py # Config-based renewal letters
│   ├── sort_renewal_list.py     # Excel list processor
│   ├── reconciller.py           # PDF table comparison & highlighting
│   ├── unlocked.py              # PDF password/restriction removal
│   ├── redact.py                # PDF word redaction
│   ├── utils.py                 # Shared utilities
│   ├── constants.py             # RECTS, REGEX_PATTERNS, and field mappings
│   ├── config.xlsx              # Configuration file (in py/ directory)
│   ├── debug.py                 # Development debugging & offset calculation
│   ├── test.py                  # Minimal test (prints sys.prefix)
│   └── assets/                  # Templates and resources
│       ├── *.docx               # Word template (any .docx file, auto-detected)
│       ├── Glass Polcies.xls    # Reference file for glass policy matching
│       └── Sonya-Swarm-Gameboy.ico  # Application icon
├── requirements.txt             # Python dependencies
└── CLAUDE.md
```

**Note:** Working directory is `py/`. All scripts run from this directory, so `Path.cwd()` resolves to `py/`.

## Common Input/Output Paths

- **Input PDFs:** `Path.home() / "Downloads"` (auto_renewal_letter.py, unlocked.py, reconciller.py)
- **Input Excel:** `Path.home() / "Downloads"` (sort_renewal_list.py)
- **Output Documents:** `Path.home() / "Desktop"` (renewal letters, sorted lists, highlighted PDFs)
- **Config File:** `config.xlsx` (in `py/` directory)
- **Word Template:** First `.docx` file found in `assets/` folder (auto-detected)
- **Redaction Input:** `input/` folder + `config.txt` (relative to `py/`)
- **Redaction/Unlock Output:** `output/` or `~/Downloads/unlocked/` (relative to `py/`)
