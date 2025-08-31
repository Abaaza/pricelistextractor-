# MJD Pricelist Extraction Tool

## Overview
This tool extracts pricing data from the MJD-PRICELIST.xlsx file with custom logic for each of the 26 sheets. It generates a structured pricelist with cell references for rate updates and optional AI-generated keywords.

## Features

### Sheet-Specific Extractors
The tool includes specialized extractors for different sheet types:

1. **GroundworksExtractor** - For Groundworks sheet
   - Handles excavation, earthworks, and foundation items
   - Identifies subcategories like "Site Clearance", "Excavation", "Filling"

2. **RCWorksExtractor** - For Reinforced Concrete works
   - Processes concrete, formwork, and reinforcement items
   - Subcategories: "In-situ Concrete", "Formwork", "Reinforcement"

3. **DrainageExtractor** - For Drainage systems
   - Handles pipes, manholes, and drainage infrastructure
   - Adapts to varying column layouts in drainage sheets

4. **ServicesExtractor** - For M&E Services
   - Processes mechanical and electrical items
   - Subcategories: "Electrical", "Mechanical", "Plumbing", "HVAC"

5. **ExternalWorksExtractor** - For External Works
   - Handles paving, landscaping, and site works
   - Subcategories: "Paving", "Landscaping", "Fencing", "Roads"

6. **PrelimsExtractor** - For Preliminaries (3 sheets)
   - Processes time-based preliminary items
   - Handles weekly/monthly duration-based pricing

7. **GenericExtractor** - For remaining sheets
   - Auto-detects column positions
   - Flexible extraction for non-standard sheets

## Output Schema

Each extracted item contains:

```json
{
  "id": "GW_0001_Groundworks",        // Unique identifier
  "code": "GW0001",                   // Short code
  "ref": "optional reference",        // Optional reference
  "description": "Item description",   // Full description
  "unit": "m3",                       // Unit of measurement
  "category": "Groundworks",          // Sheet name
  "subCategory": "Excavation",        // Section within sheet
  "keywords": ["excavate", "dig"],    // AI-generated keywords (optional)
  "cellRates": {                      // Excel cell references for rates
    "cellRate1": {
      "reference": "F15",            // Cell reference (e.g., F15)
      "sheetName": "Groundworks",    // Source sheet
      "rate": 45.50                  // Current rate value
    },
    "cellRate2": {...},              // Additional rates if present
    "cellRate3": {...},
    "cellRate4": {...}
  },
  "patterns": []                      // Placeholder for matching patterns
}
```

## Installation

### Prerequisites
- Python 3.8 or higher
- Excel file: MJD-PRICELIST.xlsx

### Setup

1. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

2. **Optional: Set OpenAI API key for keyword generation:**
   ```bash
   # Windows
   set OPENAI_API_KEY=your-api-key-here
   
   # Mac/Linux
   export OPENAI_API_KEY=your-api-key-here
   ```

## Usage

### Quick Start (Windows)
Double-click `setup_and_run.bat` to:
1. Install dependencies
2. Optionally set OpenAI API key
3. Run the extraction

### Manual Execution
```bash
python pricelist_extractor.py
```

### Programmatic Usage
```python
from pricelist_extractor import PricelistExtractor

# Initialize extractor
extractor = PricelistExtractor(
    file_path="MJD-PRICELIST.xlsx",
    use_openai=True  # Enable AI keywords
)

# Extract all sheets
items = extractor.extract_all_sheets()

# Export results
extractor.export_to_csv("output.csv")
extractor.export_to_json("output.json")
```

## Output Files

The tool generates two output files:

1. **pricelist_extracted.csv** - CSV format for spreadsheet tools
   - All fields in flat structure
   - Cell rates as separate columns
   - Keywords pipe-delimited

2. **pricelist_extracted.json** - JSON format for applications
   - Nested structure
   - Preserves all relationships
   - Direct import ready

## Extraction Logic Details

### Column Detection
Each extractor uses intelligent column detection:
- Searches for headers like "Description", "Rate", "Unit"
- Adapts to varying column positions
- Validates data types before extraction

### Rate Extraction
- Identifies valid rate values (numeric, reasonable range)
- Captures multiple rate columns when present
- Stores exact cell references for updates

### Subcategory Detection
- Bold text indicates subcategories
- Keyword matching for common categories
- Hierarchical structure preservation

### Data Validation
- Skips total/subtotal rows
- Filters empty rows
- Validates rate ranges (0 < rate < 1,000,000)

## Customization

### Adding New Sheet Extractors

Create a new extractor class:

```python
class CustomSheetExtractor(SheetExtractor):
    def extract_items(self) -> List[PriceItem]:
        items = []
        # Custom extraction logic
        return items

# Register in SHEET_EXTRACTORS
PricelistExtractor.SHEET_EXTRACTORS['SheetName'] = CustomSheetExtractor
```

### Modifying Extraction Rules

Edit extractor methods:
- `is_valid_rate()` - Change rate validation
- `clean_value()` - Modify data cleaning
- `extract_header_info()` - Adjust header detection

## Sheet Processing Summary

| Sheet Name | Extractor Type | Items Expected | Special Logic |
|------------|---------------|----------------|---------------|
| Groundworks | GroundworksExtractor | ~200-300 | Subcategory headers |
| RC works | RCWorksExtractor | ~400-500 | Concrete/formwork sections |
| Drainage | DrainageExtractor | ~300-400 | Variable column positions |
| Services | ServicesExtractor | ~100-200 | M&E categorization |
| External Works | ExternalWorksExtractor | ~150-250 | Landscaping items |
| Prelims (full) | PrelimsExtractor | ~50-100 | Duration-based |
| Prelims (consoltd) | PrelimsExtractor | ~30-50 | Consolidated items |
| Tower Cranes | GenericExtractor | ~20-40 | Equipment rates |
| Hoists | GenericExtractor | ~10-20 | Equipment rates |
| Others | GenericExtractor | Varies | Auto-detection |

## Troubleshooting

### Common Issues

1. **"File not found" error**
   - Ensure MJD-PRICELIST.xlsx is in the same directory
   - Check file path in the script

2. **No items extracted from a sheet**
   - Sheet may have non-standard structure
   - Check data_start_row in the specific extractor
   - Verify column positions

3. **OpenAI keywords not generated**
   - Verify API key is set correctly
   - Check internet connection
   - API quota may be exceeded

4. **Rate extraction issues**
   - Some sheets may have rates in different columns
   - Check the rate validation range
   - Verify numeric formatting in Excel

### Debug Mode

Enable detailed logging:

```python
# In main() function
extractor = PricelistExtractor(file_path, use_openai=False)
extractor.debug = True  # Add debug flag
```

## Performance

- Extraction time: ~10-30 seconds for all sheets
- With OpenAI keywords: +2-3 seconds per 100 items
- Memory usage: ~50-100 MB for full extraction
- Output file sizes:
  - CSV: ~2-5 MB
  - JSON: ~3-7 MB

## Future Enhancements

Planned improvements:
1. Pattern learning from historical matches
2. Confidence scoring for extracted items
3. Duplicate detection and merging
4. Rate change tracking over versions
5. Export to database formats
6. Web interface for extraction configuration

## Support

For issues or questions:
1. Check this README first
2. Review the extraction logs
3. Validate Excel file structure
4. Contact support with error details