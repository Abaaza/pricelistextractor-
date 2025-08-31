"""
Comprehensive Full Pricelist Extraction
Extracts ALL items from all sheets with proper formatting
"""

import pandas as pd
import numpy as np
import json
import re
import string
from pathlib import Path

class FullPricelistExtractor:
    def __init__(self, excel_file='MJD-PRICELIST.xlsx'):
        self.excel_file = excel_file
        self.all_items = []
        
    def get_cell_reference(self, row_idx, col_idx):
        """Convert to Excel cell reference"""
        if col_idx < 26:
            col_letter = string.ascii_uppercase[col_idx]
        else:
            col_letter = string.ascii_uppercase[col_idx // 26 - 1] + string.ascii_uppercase[col_idx % 26]
        return f"{col_letter}{row_idx + 1}"
    
    def is_valid_code(self, value):
        """Check if value could be a valid item code"""
        if pd.isna(value):
            return False
        
        str_val = str(value).strip()
        
        # Skip common headers
        if str_val.lower() in ['', 'nan', 'none', 'oliver connell', 'client', 'site', 
                               'schedule of works', 'item', 'description', 'unit', 'rate']:
            return False
        
        # Skip if it's just text headers
        if any(x in str_val.lower() for x in ['oliver connell', 'schedule', 'client:', 'site:']):
            return False
            
        # Accept numeric codes or alphanumeric codes
        try:
            # Could be a number
            float(str_val)
            return True
        except:
            # Or could be alphanumeric code
            if len(str_val) < 50:  # Reasonable code length
                return True
        
        return False
    
    def is_unit(self, value):
        """Check if value is a unit"""
        if pd.isna(value):
            return False
        
        str_val = str(value).strip().lower()
        
        # Common units
        units = ['m', 'm2', 'm²', 'm3', 'm³', 'mm', 'nr', 'no', 'item', 'sum',
                 'kg', 'tonnes', 't', 'lm', 'sqm', 'cum', 'each', 'set',
                 'l.s.', 'ls', 'hour', 'hr', 'day', 'week', 'month',
                 'lin.m', 'sq.m', 'cu.m', 'no.', 'l/s']
        
        # Direct match
        if str_val in units:
            return True
        
        # Check if it contains unit patterns
        if any(unit in str_val for unit in units):
            return True
            
        return False
    
    def extract_description(self, row, start_idx=1, end_idx=10):
        """Extract description from row"""
        desc_parts = []
        
        for i in range(start_idx, min(end_idx, len(row))):
            if pd.notna(row[i]):
                val = str(row[i]).strip()
                
                # Skip if it's a number (likely rate) or unit
                if not self.is_unit(val):
                    try:
                        # Skip large numbers (rates)
                        if float(val) > 0:
                            continue
                    except:
                        # It's text, add it
                        if val and len(val) > 1:
                            desc_parts.append(val)
        
        return ' '.join(desc_parts)
    
    def find_unit(self, row):
        """Find unit in row"""
        # Common unit column positions
        for idx in [2, 3, 4, 5]:
            if idx < len(row) and pd.notna(row[idx]):
                if self.is_unit(row[idx]):
                    return str(row[idx]).strip()
        
        # Search more broadly
        for idx in range(2, min(10, len(row))):
            if pd.notna(row[idx]):
                val = str(row[idx]).strip()
                if self.is_unit(val):
                    return val
        
        return 'item'  # Default
    
    def find_rate(self, row, start_idx=3):
        """Find rate value and column"""
        for idx in range(start_idx, min(20, len(row))):
            if pd.notna(row[idx]):
                try:
                    val = float(str(row[idx]).replace(',', '').replace('£', ''))
                    # Reasonable rate range
                    if 0 < val < 100000:
                        return val, idx
                except:
                    continue
        return 0, None
    
    def standardize_unit(self, unit):
        """Standardize unit format"""
        if not unit:
            return 'item'
        
        unit_map = {
            'm2': 'm²', 'sqm': 'm²', 'sq.m': 'm²',
            'm3': 'm³', 'cum': 'm³', 'cu.m': 'm³',
            'no': 'nr', 'no.': 'nr', 'each': 'nr',
            't': 'tonnes', 'ton': 'tonnes',
            'lm': 'm', 'lin.m': 'm',
            'l.s.': 'sum', 'ls': 'sum', 'l/s': 'sum',
        }
        
        unit_lower = unit.lower().strip()
        return unit_map.get(unit_lower, unit_lower)
    
    def extract_sheet(self, sheet_name):
        """Extract all items from a sheet"""
        print(f"\nProcessing {sheet_name}...")
        
        try:
            df = pd.read_excel(self.excel_file, sheet_name=sheet_name, header=None)
            print(f"  Sheet has {len(df)} rows")
            
            items = []
            consecutive_empty = 0
            data_started = False
            
            for idx, row in df.iterrows():
                # Check first column for code
                first_val = row[0] if 0 < len(row) else None
                
                if self.is_valid_code(first_val):
                    code = str(first_val).strip()
                    
                    # Try to extract description
                    description = self.extract_description(row.values)
                    
                    # Skip if no description
                    if not description or len(description) < 3:
                        # Try harder - maybe description is in first column after code
                        if pd.notna(row[1]):
                            description = str(row[1]).strip()
                        else:
                            continue
                    
                    # Find unit
                    unit = self.find_unit(row.values)
                    
                    # Find rate
                    rate, rate_col = self.find_rate(row.values)
                    
                    # Determine subcategory based on description
                    desc_lower = description.lower()
                    subcategory = sheet_name  # Default
                    
                    if sheet_name == 'Groundworks':
                        if 'demolish' in desc_lower or 'demolition' in desc_lower:
                            subcategory = 'Demolition'
                        elif 'excavat' in desc_lower:
                            subcategory = 'Excavation'
                        elif 'fill' in desc_lower:
                            subcategory = 'Filling'
                        elif 'disposal' in desc_lower:
                            subcategory = 'Disposal'
                        elif 'piling' in desc_lower:
                            subcategory = 'Piling'
                    elif sheet_name == 'RC works':
                        if 'concrete' in desc_lower:
                            subcategory = 'Concrete'
                        elif 'reinforcement' in desc_lower or 'rebar' in desc_lower:
                            subcategory = 'Reinforcement'
                        elif 'formwork' in desc_lower:
                            subcategory = 'Formwork'
                    elif sheet_name == 'Drainage':
                        if 'pipe' in desc_lower:
                            subcategory = 'Pipes'
                        elif 'manhole' in desc_lower:
                            subcategory = 'Manholes'
                        elif 'gully' in desc_lower:
                            subcategory = 'Gullies'
                    
                    # Create item
                    item = {
                        'id': code,
                        'code': code,
                        'description': description,
                        'unit': self.standardize_unit(unit),
                        'category': sheet_name,
                        'subcategory': subcategory,
                        'rate': rate,
                        'cellRate_reference': f"{sheet_name}!{self.get_cell_reference(idx, rate_col)}" if rate_col is not None else None,
                        'cellRate_rate': rate,
                        'excelCellReference': f"{sheet_name}!{self.get_cell_reference(idx, 0)}",
                        'sourceSheetName': sheet_name,
                        'keywords': []
                    }
                    
                    items.append(item)
                    data_started = True
                    consecutive_empty = 0
                else:
                    # Track empty rows to stop if we hit too many
                    if data_started:
                        consecutive_empty += 1
                        if consecutive_empty > 50:  # Stop after 50 empty rows
                            break
            
            print(f"  Extracted {len(items)} items from {sheet_name}")
            return items
            
        except Exception as e:
            print(f"  Error processing {sheet_name}: {e}")
            return []
    
    def extract_all(self):
        """Extract from all sheets"""
        sheets = ['Groundworks', 'RC works', 'Drainage', 'Services', 'External Works', 'Underpinning']
        
        print("="*80)
        print("COMPREHENSIVE PRICELIST EXTRACTION")
        print("="*80)
        
        all_items = []
        
        for sheet in sheets:
            items = self.extract_sheet(sheet)
            all_items.extend(items)
        
        self.all_items = all_items
        return all_items
    
    def save_outputs(self, prefix='full_pricelist'):
        """Save extracted data"""
        if not self.all_items:
            print("No items to save!")
            return
        
        print(f"\nSaving {len(self.all_items)} items...")
        
        # Save JSON
        json_file = f"{prefix}.json"
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(self.all_items, f, indent=2, ensure_ascii=False)
        print(f"Saved JSON: {json_file}")
        
        # Save CSV
        df = pd.DataFrame(self.all_items)
        csv_file = f"{prefix}.csv"
        df.to_csv(csv_file, index=False)
        print(f"Saved CSV: {csv_file}")
        
        return json_file, csv_file
    
    def show_statistics(self):
        """Show extraction statistics"""
        if not self.all_items:
            return
        
        print("\n" + "="*80)
        print("EXTRACTION STATISTICS")
        print("="*80)
        
        df = pd.DataFrame(self.all_items)
        
        print(f"\nTotal items extracted: {len(self.all_items)}")
        print(f"Items with rates: {sum(1 for i in self.all_items if i['rate'] > 0)}")
        print(f"Items with cell references: {sum(1 for i in self.all_items if i['cellRate_reference'])}")
        
        print("\nItems per category:")
        for cat in df['category'].unique():
            count = len(df[df['category'] == cat])
            with_rates = len(df[(df['category'] == cat) & (df['rate'] > 0)])
            print(f"  {cat:20} - {count:5} items ({with_rates} with rates)")
        
        # Sample items
        print("\nSample items:")
        samples = df[df['rate'] > 0].head(5)
        for _, item in samples.iterrows():
            print(f"\n  Code: {item['code']}")
            print(f"  Desc: {item['description'][:60]}...")
            print(f"  Unit: {item['unit']}")
            print(f"  Rate: {item['rate']}")
            print(f"  Cell: {item['cellRate_reference']}")

def main():
    extractor = FullPricelistExtractor()
    
    # Extract all items
    items = extractor.extract_all()
    
    if items:
        # Save outputs
        extractor.save_outputs()
        
        # Show statistics
        extractor.show_statistics()
        
        print("\n✓ Full extraction complete!")
    else:
        print("\n✗ No items extracted!")

if __name__ == "__main__":
    main()