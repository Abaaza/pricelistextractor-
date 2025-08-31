"""
Base extractor class with common functionality for all sheet extractors
Ensures consistent format across all extractors
"""

import pandas as pd
import numpy as np
import json
import re
from pathlib import Path
import string

class BaseExtractor:
    def __init__(self, excel_file='MJD-PRICELIST.xlsx', sheet_name=''):
        self.excel_file = excel_file
        self.sheet_name = sheet_name
        self.df = None
        self.extracted_items = []
        
    def load_sheet(self):
        """Load the sheet"""
        print(f"Loading {self.sheet_name} sheet...")
        self.df = pd.read_excel(self.excel_file, sheet_name=self.sheet_name, header=None)
        print(f"Loaded {len(self.df)} rows x {len(self.df.columns)} columns")
        return self.df
    
    def get_cell_reference(self, row_idx, col_idx):
        """Convert row and column index to Excel cell reference"""
        if col_idx < 26:
            col_letter = string.ascii_uppercase[col_idx]
        else:
            col_letter = string.ascii_uppercase[col_idx // 26 - 1] + string.ascii_uppercase[col_idx % 26]
        
        return f"{col_letter}{row_idx + 1}"
    
    def get_sheet_cell_reference(self, row_idx, col_idx):
        """Get full cell reference with sheet name (e.g., 'Groundworks!F20')"""
        cell_ref = self.get_cell_reference(row_idx, col_idx)
        return f"{self.sheet_name}!{cell_ref}"
    
    def extract_code(self, row, col_idx=0):
        """Extract code from row - tries to get actual Excel code"""
        if col_idx < len(row) and pd.notna(row[col_idx]):
            code = str(row[col_idx]).strip()
            # Clean up code but keep the actual value
            if code and not code.lower() in ['nan', 'none', '']:
                # Remove only excessive whitespace, keep the code as-is
                code = ' '.join(code.split())
                return code
        return None
    
    def is_unit(self, value):
        """Check if value is a unit"""
        if pd.isna(value):
            return False
        
        value_str = str(value).strip().lower()
        
        # Don't treat numbers as units
        try:
            float(value_str)
            return False
        except:
            pass
        
        units = ['m', 'm2', 'm²', 'm3', 'm³', 'mm', 'nr', 'no', 'item', 'sum',
                 'kg', 'tonnes', 't', 'lm', 'sqm', 'cum', 'each', 'set',
                 'l.s.', 'ls', 'hour', 'hr', 'day', 'week', 'month']
        
        return value_str in units or any(unit in value_str for unit in units)
    
    def standardize_unit(self, unit):
        """Standardize unit format"""
        if not unit:
            return 'item'
            
        unit_map = {
            'm2': 'm²', 'sqm': 'm²', 'sq.m': 'm²', 'sq m': 'm²',
            'm3': 'm³', 'cum': 'm³', 'cu.m': 'm³', 'cu m': 'm³',
            'no': 'nr', 'no.': 'nr', 'each': 'nr', 'number': 'nr',
            't': 'tonnes', 'ton': 'tonnes', 'tonne': 'tonnes',
            'lm': 'm', 'lin.m': 'm', 'l.m': 'm', 'lin m': 'm',
            'l.s.': 'sum', 'ls': 'sum', 'lump sum': 'sum',
            'hr': 'hour', 'hrs': 'hour',
        }
        
        unit_lower = unit.lower().strip()
        return unit_map.get(unit_lower, unit_lower)
    
    def extract_rate(self, row, start_col=3, end_col=10):
        """Extract rate value from typical rate columns"""
        for col_idx in range(start_col, min(end_col, len(row))):
            if pd.notna(row[col_idx]):
                value = str(row[col_idx]).strip()
                # Check if it's a number
                value_clean = value.replace(',', '').replace('£', '').replace('$', '')
                try:
                    rate = float(value_clean)
                    if rate > 0 and rate < 1000000:  # Sanity check
                        return rate, col_idx  # Return both rate and column index
                except:
                    continue
        return None, None
    
    def create_item(self, row_idx, row, code=None, description='', unit='item', 
                   subcategory='', rate=None, rate_col_idx=None, keywords=None):
        """Create standardized item dictionary"""
        # Use actual code for both id and code
        if not code:
            code = f"{self.sheet_name[:2].upper()}{row_idx}"
        
        # Get cell references
        excel_ref = self.get_sheet_cell_reference(row_idx, 0)
        
        # Get rate cell reference with sheet name
        rate_cell_ref = None
        if rate_col_idx is not None:
            rate_cell_ref = self.get_sheet_cell_reference(row_idx, rate_col_idx)
        
        item = {
            'id': code,  # Use the actual code as ID
            'code': code,  # Same code
            'description': description,
            'unit': self.standardize_unit(unit),
            'category': self.sheet_name,
            'subcategory': subcategory,
            'rate': rate if rate else 0,
            'cellRate_reference': rate_cell_ref,
            'cellRate_rate': rate if rate else 0,
            'excelCellReference': excel_ref,
            'sourceSheetName': self.sheet_name,
            'keywords': keywords if keywords else []
        }
        
        return item
    
    def save_output(self, output_prefix=None):
        """Save extracted data"""
        if not self.extracted_items:
            print("No items to save")
            return
        
        if not output_prefix:
            output_prefix = self.sheet_name.lower().replace(' ', '_')
        
        # Save JSON
        json_file = f"{output_prefix}_extracted.json"
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(self.extracted_items, f, indent=2, ensure_ascii=False)
        print(f"Saved JSON: {json_file}")
        
        # Save CSV
        df = pd.DataFrame(self.extracted_items)
        df['keywords'] = df['keywords'].apply(lambda x: '|'.join(x) if isinstance(x, list) else '')
        csv_file = f"{output_prefix}_extracted.csv"
        df.to_csv(csv_file, index=False)
        print(f"Saved CSV: {csv_file}")
        
        return json_file, csv_file