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
- `tabulate`: Format tabular data (used in reconciler_tool.py)

## Core Architecture

### Entry Point System

The main automation dispatcher is `py/file_completion_tool.py`, which:
- Reads configuration from `config.xlsx` (in `py/` directory)
- Routes tasks based on the `event` field in cell B3
- Supported events: "manual renewal letter", "auto renewal letter", "sort renewal list"

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
- `write_to_new_docx(template_path=None, data=None, output_dir=None)`: Renders Word templates using `docxtpl`
- Auto-detects first .docx file in `assets/` folder (no need to specify template name)
- Returns `True` on success, `False` if template not found
- If template not found: Prints GitHub message and returns False gracefully
- Auto-generates output filename from `named_insured` field
- Default output: Desktop

**UI:**
- `progressbar(iterable, prefix, size)`: Terminal progress bar with time estimates

**Configuration Loading:**
- `CONFIG_FIELDS`: Defines Excel cell mappings for all configuration fields (rows 2-29, column B)

### PDF Extraction System

#### Constants (`constants.py`)

The `RECTS` dictionary defines PDF field extraction rules for multiple insurance companies:
- **Policy Type Detection:** `RECTS['policy_type']` contains keyword + rect pairs for auto-detecting insurers
- **Field Mappings:** Each insurer (Aviva, Family, Intact, Wawanesa) has field extraction rules with:
  - `pattern`: Regex to find keywords (optional)
  - `rect`: `fitz.Rect` for text clipping (can be offset from pattern match or absolute coordinates)
  - `return_all`: Boolean flag to extract all matches instead of just first match

**Helper Functions for Field Definitions:**
- `absolute_rect_field(x0, y0, x1, y1)`: Extract from absolute coordinates
- `pattern_only_field(regex_str, return_all=False)`: Extract by pattern matching only
- `pattern_with_offset_field(regex_str, dx0, dy0, dx1, dy1, return_all=False)`: Pattern + relative offset

#### PDF Extraction Functions (`auto_renewal_letter.py`)

**Core Extraction Logic:**
1. `detect_insurer(doc, rects)`: Auto-detects insurer using keyword matching in `RECTS['policy_type']`
2. `extract_fields(doc, field_mapping, insurer)`: Extracts all fields for detected insurer
3. `extract_single_field(pages_dict, cfg, doc)`: Handles extraction based on field config:
   - Pattern only: Uses `search_text()` to find matching text
   - Rect only: Uses `extract_text_from_absolute_rect()` for fixed coordinates
   - Pattern + rect offset: Uses `extract_with_pattern_and_offset()` to find pattern, then extract from offset position
   - Supports `return_all=True` for fields that can have multiple values (e.g., multiple locations)

**Extraction Patterns:**
- If `pattern` only: Returns text in keyword bounding box
- If `pattern` + `rect`: Returns text from `keyword_position + rect_offset`
- If `rect` only: Direct extraction from absolute coordinates
- If `return_all=True`: Returns list of all matches instead of first match

### Main Tools

#### 1. Auto Renewal Letter (`auto_renewal_letter.py`)
- Processes all PDFs in `~/Downloads`
- Auto-detects insurer type using `detect_insurer()` and `RECTS['policy_type']`
- Extracts policy details from declarations using `extract_fields()`
- Generates renewal letters to `~/Desktop`
- Uses template: `assets/Renewal Letter.docx`
- Handles multiple risk addresses and form types (returns list if multiple matches)
- Applies special formatting to addresses (title case, province codes, etc.)

**Data Flow:**
```
PDF → detect_insurer() → extract_fields() → format_extracted_data() → write_to_new_docx()
```

**Key Functions:**
- `auto_renewal_letter(config_data)`: Main entry point that processes all PDFs
- `get_text(doc, structured=True)`: Extracts text from PDF as pages → blocks
- `search_text(pages_dict, regex, return_all=False)`: Search regex across all blocks
- `deduplicate_field(value)`: Removes duplicates from multi-value fields (Intact-specific)
- `address_one_title_case(sentence)`: Title case with ordinal numbers in lowercase
- `address_two_title_case(strings_list)`: Title case with province codes uppercased

#### 2. Manual Renewal Letter (`manual_renewal_letter.py`)
- Uses data from `config.xlsx` instead of PDF extraction
- Triggered when config.xlsx event = "manual renewal letter"
- Maps config fields using `constants.py:EXCEL_CELL_MAPPING`
- Fallback: Uses mailing address for risk address if empty
- Handles date parsing with multiple fallback formats (`%Y-%m-%d`, `%m/%d/%Y`)
- Output format standardized to `DATE_FORMAT = "%B %d, %Y"`

#### 3. Sort Renewal List (`sort_renewal_list.py`)
- Processes the 2 most recent Excel files in `~/Downloads`
- Combines data, removes duplicates by `policynum` (keeps only unique policy numbers)
- Sorts by: insurer → renewal date → name
- Adds blank rows between insurers for visual separation
- Creates formatted Excel table with TableStyleLight1
- Adds borders specifically to Pulled/D/L columns
- Auto-adjusts column widths based on content
- Configures page setup for printing (fit to width, header row on each page)
- Output: `~/Desktop/renewal_list.xlsx`

**Column Order:**
`policynum, ccode, name, pcode, csrcode, insurer, buscode, renewal, Pulled, D/L`

**Sorting Logic:**
- Uses `get_sort_key(date_value)` to handle both datetime objects and string dates
- Date format: MM-DD for sorting (returns "9999" for empty dates to push to end)

#### 4. Reconciler Tool (`reconciller_tool.py`)
- Compares PDF tables to find matching policy numbers and premiums across multiple PDFs
- Auto-detects policy and premium columns in PDF tables using intelligent heuristics
- Uses pattern matching: `^[A-Z0-9]{6,}$` for policy numbers (min 6 alphanumeric characters)
- Premium column detection: Finds numeric columns ending in ".00" with highest max value
- Input: `../input/` folder
- Output: `../output/` folder

**Key Functions:**
- `find_policy_column(df, min_unique=2)`: Finds column with most unique policy numbers matching pattern
- `find_premium_column(df)`: Finds numeric column with values ending in .00 and highest max value
- `extract_policies_and_premiums(pdf_path)`: Extracts all policy/premium pairs from all tables in PDF
- Uses `page.find_tables(strategy="text")` from PyMuPDF for table detection

## Development Patterns

### Working with PDFs
- All PDF operations use PyMuPDF (`fitz`)
- Text extraction methods:
  - `page.get_text("text", clip=rect)`: Extract text from specific region
  - `page.get_text("blocks")`: Returns blocks with coordinates `(x0, y0, x1, y1, text, block_num, block_type)`
  - `page.get_textbox(rect)`: Extract text within a rect (used in auto_renewal_letter.py)
  - `page.find_tables(strategy="text")`: Extract tables as pandas DataFrames (used in reconciller_tool.py)
- `page.search_for(pattern)` returns list of `fitz.Rect` for keyword locations
- When adding new insurers: Update `RECTS` in `constants.py` with keyword + rect definitions

### Adding New Insurers to Auto Renewal Letter
1. Add detection entry to `POLICY_TYPE_DETECTION` with unique keyword and rect
2. Create field mapping dictionary (e.g., `NEWINSURER_FIELDS`) using helper functions:
   - Use `absolute_rect_field()` for fixed positions (e.g., name/address blocks)
   - Use `pattern_only_field()` for regex-based extraction
   - Use `pattern_with_offset_field()` for relative positioning from keywords
   - Add `return_all=True` for fields that can have multiple values
3. Add insurer to `RECTS` dictionary at the bottom of `constants.py`
4. Test with sample PDF from the insurer to verify all fields extract correctly

### Adding New Automation Tasks
1. Create new script in `py/` directory with main function that accepts config_data dict
2. Add event trigger to `file_completion_tool.py:main()`:
   ```python
   elif event == "new task name":
       new_task_function(config_data)
   ```
3. Update `config.xlsx` if new configuration fields needed:
   - Add new cells to column B
   - Update `EXCEL_CELL_MAPPING` in `constants.py`
   - Update `DEFAULT_MAPPING` in `constants.py`
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

## File Organization

```
Python-Automations/
├── py/                          # All Python source files (working directory)
│   ├── file_completion_tool.py  # Main dispatcher (entry point)
│   ├── auto_renewal_letter.py   # PDF-based renewal letters
│   ├── manual_renewal_letter.py # Config-based renewal letters
│   ├── sort_renewal_list.py     # Excel list processor
│   ├── reconciller_tool.py      # PDF table comparison
│   ├── utils.py                 # Shared utilities
│   ├── constants.py             # RECTS and field mappings
│   ├── config.xlsx              # Configuration file (in py/ directory)
│   ├── debug.py                 # Testing/debugging
│   └── assets/                  # Templates and resources
│       ├── *.docx               # Word template (any .docx file, auto-detected)
│       ├── Glass Polcies.xls    # Reference file
│       └── Sonya-Swarm-Gameboy.ico  # Application icon
└── requirements.txt             # Python dependencies
```

**Note:** Working directory is `py/`. All scripts run from this directory, so `Path.cwd()` resolves to `py/`.

## Common Input/Output Paths

- **Input PDFs:** `Path.home() / "Downloads"` (auto_renewal_letter.py)
- **Input Excel:** `Path.home() / "Downloads"` (sort_renewal_list.py - picks 2 most recent files)
- **Output Documents:** `Path.home() / "Desktop"` (renewal letters, sorted lists)
- **Config File:** `config.xlsx` (in `py/` directory)
- **Word Template:** First `.docx` file found in `assets/` folder (auto-detected)
- **Reconciler Input:** `../input/` (relative to `py/`)
- **Reconciler Output:** `../output/` (relative to `py/`)

## Important Implementation Details

### Regex Patterns in constants.py
- `postal_code`: Canadian postal code format (e.g., V6B 1A1)
- `dollar`: Extracts dollar amounts with $ prefix
- `date`: Matches various date formats including "DD Mon YYYY" and "Month DD, YYYY"
- `address`: Matches street addresses (excludes "Ltd.")
- `and`: Matches "&" or word "and" (case-insensitive)

### Template Auto-Detection
- The `write_to_new_docx` function automatically finds the first `.docx` file in the `assets/` folder
- Template can have any filename (e.g., "Renewal Letter.docx", "Template.docx", etc.)
- If no template found, displays message: "Template not found. Visit https://github.com/yourusername/Python-Automations to download the template."
- Function returns `False` on missing template (allows graceful error handling)

### Template Variables (docxtpl)
The Word template uses Jinja2 syntax for variable substitution:
- `{{ named_insured }}`: Client name
- `{{ broker_name }}`: Broker name
- `{{ on_behalf }}`: On behalf of field
- `{{ policy_number }}`: Policy number
- `{{ effective_date }}`: Policy effective date
- `{{ mailing_address }}`: Full mailing address (concatenated)
- `{{ risk_address_1 }}`: Risk address
- `{{ insurer }}`: Insurance company name
- `{{ today }}`: Current date

## Testing

- Each script has `if __name__ == "__main__"` block with test data
- `debug.py` for development testing and experimentation
- Test with sample PDFs from each insurer before building executables
- Verify field extraction by checking intermediate values during development
