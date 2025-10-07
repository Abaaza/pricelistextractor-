# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

This is a Python-based pricelist extraction tool that extracts pricing data from the MJD-PRICELIST.xlsx Excel file. It processes 26+ sheets with different structures and generates structured JSON/CSV outputs with cell references for rate updates.

## Key Commands

### Installation
```bash
pip install -r requirements.txt
```

### Run Individual Sheet Extractors
Each sheet has its own specialized extractor:
```bash
python extract_drainage.py
python extract_groundworks.py
python extract_rc_works.py
python extract_external_works.py
python extract_services.py
python extract_underpinning.py
```

### Combine All Extracts
After running individual extractors:
```bash
python combine_all_extracts.py
```

### Optional: Set OpenAI API Key
For AI-generated keyword extraction (optional):
```bash
# Windows
set OPENAI_API_KEY=your-api-key-here

# Mac/Linux
export OPENAI_API_KEY=your-api-key-here
```

## Architecture

### Base Extractor Pattern
All sheet extractors inherit from `BaseExtractor` (extractor_base.py) which provides:
- **Cell reference tracking**: Converts (row, col) indices to Excel references (e.g., "F20")
- **Sheet-qualified references**: Formats as "SheetName!F20" for global uniqueness
- **Code preservation**: Extracts and preserves original Excel codes
- **Unit standardization**: Maps m2/m²/sqm → "m2", m3/m³/cum → "m3", etc.
- **Rate extraction**: Scans columns D-J for valid numeric rates, defaulting to column F (index 5)
- **Item creation**: Standardized schema with id, code, description, unit, category, subcategory, rate, cellRate_reference, keywords

### Sheet-Specific Extractors
Each extractor (DrainageExtractor, GroundworksExtractor, RCWorksExtractor, etc.) extends BaseExtractor with:
- **Column detection logic**: Handles varying layouts per sheet
- **Bold text detection**: Uses openpyxl to detect bold rows as subcategory headers
- **Description enhancement**: Expands abbreviations (ne → not exceeding, thk → thick, exc → excavation)
- **Subcategory inference**: Keyword-based or header-based categorization
- **Keyword generation**: Extracts measurements and technical terms

### Key Implementation Details

#### Drainage Extractor (extract_drainage.py)
- Combines header descriptions with range data (e.g., "depth to invert: 2.5 - 3.5m")
- Checks column O (15) for rates, falls back to column T (20)
- Headers have empty column A but text in column B

#### Groundworks/RC Works Extractors
- Load workbook twice: pandas for data, openpyxl for formatting (bold detection)
- Start extraction from row 10 (index 9)
- Bold rows become subcategory headers for subsequent items
- Description from columns B+C, unit from column E (index 4)
- Rate from columns D-J with validation (0 ≤ rate < 1,000,000)

#### Combiner (combine_all_extracts.py)
- Reads all `Files/*_extracted.{csv,json}` files
- Reassigns sequential IDs across all items
- Generates summary statistics by category/subcategory

## Output Schema

Every extracted item follows this structure:
```json
{
  "id": "code_value",           // Uses actual Excel code as ID
  "code": "code_value",          // Original Excel code
  "description": "text",         // Enhanced with expanded abbreviations
  "unit": "m2",                  // Standardized unit
  "category": "Groundworks",     // Sheet name
  "subcategory": "Excavation",   // Section within sheet
  "rate": 45.50,                 // Numeric rate value
  "cellRate_reference": "Groundworks!F20",  // Full cell reference
  "cellRate_rate": 45.50,        // Same as rate
  "excelCellReference": "Groundworks!A20",  // Item's row reference
  "sourceSheetName": "Groundworks",
  "keywords": ["excavate", "150mm"]  // Search terms
}
```

## Important Patterns

### Cell Reference Format
- All cell references include sheet name: `"SheetName!ColumnRow"`
- Example: `"RC works!F42"`, `"Drainage!O156"`
- This enables direct Excel updates via cell reference lookup

### Code Preservation
- Original Excel codes are preserved exactly as-is
- Used as both `id` and `code` fields
- If no code exists in Excel, sequential numbers are assigned

### Rate Column Detection
The `extract_rate()` method in BaseExtractor:
1. Scans columns D-J (indices 3-9) for numeric values
2. Validates: 0 ≤ rate < 1,000,000
3. Returns both rate value and column index
4. Defaults to column F (index 5) if none found
5. Returns rate as 0 if no valid rate exists

### Bold Header Detection
Many extractors use openpyxl to detect bold text:
```python
def is_row_bold(self, row_idx):
    # Checks if all non-empty cells in first 5 columns are bold
    # Used to identify subcategory headers
```

## File Organization

- **Root directory**: Individual extractor scripts (extract_*.py)
- **Files/**: Output directory for extracted CSVs and JSONs
- **Files copy/**: Backup of previous extractions
- **shit/**: Archive of old/experimental extraction scripts
- **MJD-PRICELIST.xlsx**: Source Excel file (26+ sheets)

## Data Flow

1. Individual extractors read MJD-PRICELIST.xlsx
2. Each generates `{sheet_name}_extracted.{csv,json}` in Files/
3. `combine_all_extracts.py` merges all Files/*_extracted.* into:
   - `Files/pricelist_combined_all.csv`
   - `Files/pricelist_combined_all.json`
   - `Files/pricelist_summary_report.txt`

## Special Considerations

- **Multiple workbook loads**: Some extractors load the file twice (pandas + openpyxl) to access both data and formatting
- **Row indexing**: pandas uses 0-based indexing, Excel uses 1-based, openpyxl cells use 1-based
- **Unit handling**: Never treat numbers as units; validate against known unit list before accepting
- **Rate validation**: Allows rate=0 but filters out negative or unreasonably large values
- **Description length**: Minimum 5 characters required to be considered valid
