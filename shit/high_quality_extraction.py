"""
High-Quality Pricelist Extraction
Preserves original descriptions and properly handles Drainage range items
"""

import pandas as pd
import numpy as np
import json
import re
import string
from pathlib import Path

class HighQualityExtractor:
    def __init__(self, excel_file='MJD-PRICELIST.xlsx'):
        self.excel_file = excel_file
        self.all_items = []
        
    def get_cell_reference(self, sheet_name, row_idx, col_idx):
        """Get full cell reference with sheet name"""
        if col_idx < 26:
            col_letter = string.ascii_uppercase[col_idx]
        else:
            col_letter = string.ascii_uppercase[col_idx // 26 - 1] + string.ascii_uppercase[col_idx % 26]
        return f"{sheet_name}!{col_letter}{row_idx + 1}"
    
    def is_header_row(self, value):
        """Check if this looks like a header/title row"""
        if pd.isna(value):
            return False
        str_val = str(value).strip().lower()
        
        # Common headers to skip
        headers = ['oliver connell', 'client', 'site', 'schedule of works', 
                   'item', 'description', 'unit', 'rate', 'quantity', 'amount']
        return any(h in str_val for h in headers)
    
    def is_valid_code(self, value):
        """Check if value is a valid item code"""
        if pd.isna(value):
            return False
        
        str_val = str(value).strip()
        
        # Skip empty or header values
        if not str_val or str_val.lower() in ['', 'nan', 'none', '0']:
            return False
            
        # Skip if it's a header
        if self.is_header_row(str_val):
            return False
        
        # Accept numeric codes
        try:
            code_num = float(str_val)
            if 0 < code_num < 10000:  # Reasonable code range
                return True
        except:
            pass
        
        # Accept alphanumeric codes
        if re.match(r'^[A-Z0-9]+[-/]?\d*$', str_val, re.I):
            return True
            
        return False
    
    def is_unit(self, value):
        """Check if value is a unit"""
        if pd.isna(value):
            return False
        
        str_val = str(value).strip().lower()
        
        # Don't treat pure numbers as units
        try:
            float(str_val.replace(',', ''))
            return False
        except:
            pass
        
        units = ['m', 'm2', 'm²', 'm3', 'm³', 'mm', 'nr', 'no', 'item', 'sum',
                 'kg', 'tonnes', 't', 'lm', 'sqm', 'cum', 'each', 'set',
                 'l.s.', 'ls', 'hour', 'hr', 'day', 'week', 'month',
                 'lin.m', 'sq.m', 'cu.m', 'no.', 'l/s', '%']
        
        return str_val in units or any(unit in str_val for unit in units)
    
    def standardize_unit(self, unit):
        """Standardize unit format"""
        if not unit:
            return 'item'
        
        unit_map = {
            'm2': 'm²', 'sqm': 'm²', 'sq.m': 'm²', 'sq m': 'm²',
            'm3': 'm³', 'cum': 'm³', 'cu.m': 'm³', 'cu m': 'm³',
            'no': 'nr', 'no.': 'nr', 'each': 'nr', 'number': 'nr',
            't': 'tonnes', 'ton': 'tonnes', 'tonne': 'tonnes',
            'lm': 'm', 'lin.m': 'm', 'l.m': 'm',
            'l.s.': 'sum', 'ls': 'sum', 'l/s': 'sum',
        }
        
        unit_lower = str(unit).lower().strip()
        return unit_map.get(unit_lower, unit_lower)
    
    def extract_description(self, row_vals, start_col=1, max_cols=5):
        """Extract description from row, preserving original text"""
        parts = []
        
        for col in range(start_col, min(start_col + max_cols, len(row_vals))):
            if pd.notna(row_vals[col]):
                val = str(row_vals[col]).strip()
                
                # Skip if it's a unit or a rate
                if self.is_unit(val):
                    continue
                    
                # Skip if it's likely a rate (number > 10)
                try:
                    num_val = float(val.replace(',', '').replace('£', ''))
                    if num_val > 10:
                        continue
                except:
                    # It's text, keep it
                    if val and len(val) > 1:
                        parts.append(val)
        
        return ' '.join(parts)
    
    def find_rate_and_column(self, row_vals, start_col=3):
        """Find rate value and its column index"""
        # Try common rate column positions first
        for col in [5, 6, 7, 8, 14, 15]:  # Common rate columns
            if col < len(row_vals) and pd.notna(row_vals[col]):
                try:
                    val = float(str(row_vals[col]).replace(',', '').replace('£', ''))
                    if 0 < val < 100000:
                        return val, col
                except:
                    pass
        
        # Search more broadly
        for col in range(start_col, min(20, len(row_vals))):
            if pd.notna(row_vals[col]):
                try:
                    val = float(str(row_vals[col]).replace(',', '').replace('£', ''))
                    if 0 < val < 100000:
                        return val, col
                except:
                    pass
        
        return 0, None
    
    def extract_drainage_with_ranges(self):
        """Extract Drainage sheet with proper range handling"""
        sheet_name = 'Drainage'
        print(f"\nProcessing {sheet_name} with range support...")
        
        try:
            df = pd.read_excel(self.excel_file, sheet_name=sheet_name, header=None)
            items = []
            current_header = None
            
            for idx, row in df.iterrows():
                row_vals = row.values
                
                # Check for header in column 0 or 1 (long excavation descriptions)
                for col in [0, 1]:
                    if pd.notna(row_vals[col]):
                        val = str(row_vals[col])
                        if len(val) > 50 and ('excavat' in val.lower() or 'trench' in val.lower()):
                            current_header = val.strip()
                            # Remove trailing colon if present
                            if current_header.endswith(':'):
                                current_header = current_header[:-1]
                            break
                
                # Check if this is a range row
                is_range = False
                depth_range = ""
                
                # Check for range pattern (col2=value1, col3='-', col4=value2)
                if len(row_vals) > 4 and pd.notna(row_vals[2]) and pd.notna(row_vals[3]) and pd.notna(row_vals[4]):
                    if str(row_vals[3]).strip() == '-':
                        try:
                            val1 = str(row_vals[2]).strip()
                            val2 = str(row_vals[4]).strip()
                            
                            # Handle "ne" (not exceeding)
                            if val1.lower() == 'ne':
                                depth_range = f"not exceeding {val2}m"
                            else:
                                # Try to convert to numbers
                                try:
                                    num1 = float(val1)
                                    num2 = float(val2)
                                    depth_range = f"{num1}-{num2}m"
                                except:
                                    depth_range = f"{val1}-{val2}"
                            
                            is_range = True
                        except:
                            pass
                
                if is_range and self.is_valid_code(row_vals[0]):
                    code = str(row_vals[0]).strip()
                    
                    # Build complete description
                    if current_header:
                        description = f"{current_header}; depth to invert: {depth_range}"
                    else:
                        description = f"Excavate trenches; depth to invert: {depth_range}"
                    
                    # Find unit (usually in column 5)
                    unit = 'm'  # Default for excavation
                    if len(row_vals) > 5 and pd.notna(row_vals[5]):
                        if self.is_unit(row_vals[5]):
                            unit = str(row_vals[5]).strip()
                    
                    # Find rate (often in column 14 for Drainage ranges)
                    rate, rate_col = self.find_rate_and_column(row_vals, 10)
                    
                    item = {
                        'id': code,
                        'code': code,
                        'description': description,
                        'unit': self.standardize_unit(unit),
                        'category': sheet_name,
                        'subcategory': 'Excavation',
                        'rate': rate,
                        'cellRate_reference': self.get_cell_reference(sheet_name, idx, rate_col) if rate_col else None,
                        'cellRate_rate': rate,
                        'excelCellReference': self.get_cell_reference(sheet_name, idx, 0),
                        'sourceSheetName': sheet_name,
                        'keywords': []
                    }
                    items.append(item)
                
                # Also extract normal drainage items
                elif self.is_valid_code(row_vals[0]) and not is_range:
                    code = str(row_vals[0]).strip()
                    
                    # Get description
                    description = self.extract_description(row_vals)
                    if not description or len(description) < 3:
                        continue
                    
                    # Find unit
                    unit = 'item'
                    for col in range(2, min(10, len(row_vals))):
                        if pd.notna(row_vals[col]) and self.is_unit(row_vals[col]):
                            unit = str(row_vals[col]).strip()
                            break
                    
                    # Find rate
                    rate, rate_col = self.find_rate_and_column(row_vals)
                    
                    # Determine subcategory
                    desc_lower = description.lower()
                    if 'pipe' in desc_lower:
                        subcategory = 'Pipes'
                    elif 'manhole' in desc_lower:
                        subcategory = 'Manholes'
                    elif 'gully' in desc_lower:
                        subcategory = 'Gullies'
                    elif 'channel' in desc_lower:
                        subcategory = 'Channels'
                    elif 'excavat' in desc_lower:
                        subcategory = 'Excavation'
                    else:
                        subcategory = 'Drainage'
                    
                    item = {
                        'id': code,
                        'code': code,
                        'description': description,
                        'unit': self.standardize_unit(unit),
                        'category': sheet_name,
                        'subcategory': subcategory,
                        'rate': rate,
                        'cellRate_reference': self.get_cell_reference(sheet_name, idx, rate_col) if rate_col else None,
                        'cellRate_rate': rate,
                        'excelCellReference': self.get_cell_reference(sheet_name, idx, 0),
                        'sourceSheetName': sheet_name,
                        'keywords': []
                    }
                    items.append(item)
            
            print(f"  Extracted {len(items)} items from {sheet_name}")
            return items
            
        except Exception as e:
            print(f"  Error processing {sheet_name}: {e}")
            import traceback
            traceback.print_exc()
            return []
    
    def extract_standard_sheet(self, sheet_name):
        """Extract items from standard sheets"""
        print(f"\nProcessing {sheet_name}...")
        
        try:
            df = pd.read_excel(self.excel_file, sheet_name=sheet_name, header=None)
            items = []
            
            for idx, row in df.iterrows():
                row_vals = row.values
                
                # Skip empty rows
                if pd.isna(row_vals).all() or row.notna().sum() < 2:
                    continue
                
                # Check for valid code
                if not self.is_valid_code(row_vals[0]):
                    continue
                
                code = str(row_vals[0]).strip()
                
                # Get description
                description = self.extract_description(row_vals)
                if not description or len(description) < 3:
                    continue
                
                # Find unit
                unit = 'item'
                for col in range(2, min(10, len(row_vals))):
                    if pd.notna(row_vals[col]) and self.is_unit(row_vals[col]):
                        unit = str(row_vals[col]).strip()
                        break
                
                # Find rate
                rate, rate_col = self.find_rate_and_column(row_vals)
                
                # Determine subcategory based on sheet and description
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
                    elif 'piling' in desc_lower or 'pile' in desc_lower:
                        subcategory = 'Piling'
                    elif 'compact' in desc_lower:
                        subcategory = 'Compaction'
                        
                elif sheet_name == 'RC works':
                    if 'concrete' in desc_lower:
                        if 'pour' in desc_lower:
                            subcategory = 'Concrete Pouring'
                        else:
                            subcategory = 'Concrete'
                    elif 'reinforcement' in desc_lower or 'rebar' in desc_lower:
                        subcategory = 'Reinforcement'
                    elif 'formwork' in desc_lower or 'shutter' in desc_lower:
                        subcategory = 'Formwork'
                    elif 'mesh' in desc_lower:
                        subcategory = 'Mesh Reinforcement'
                        
                elif sheet_name == 'External Works':
                    if 'paving' in desc_lower:
                        subcategory = 'Paving'
                    elif 'kerb' in desc_lower:
                        subcategory = 'Kerbs'
                    elif 'fence' in desc_lower or 'gate' in desc_lower:
                        subcategory = 'Fencing'
                    elif 'road' in desc_lower:
                        subcategory = 'Roads'
                    elif 'landscape' in desc_lower:
                        subcategory = 'Landscaping'
                        
                elif sheet_name == 'Services':
                    if 'electrical' in desc_lower or 'cable' in desc_lower:
                        subcategory = 'Electrical'
                    elif 'plumbing' in desc_lower or 'water' in desc_lower:
                        subcategory = 'Plumbing'
                    elif 'hvac' in desc_lower or 'ventilation' in desc_lower:
                        subcategory = 'HVAC'
                    elif 'fire' in desc_lower:
                        subcategory = 'Fire Systems'
                        
                elif sheet_name == 'Underpinning':
                    if 'excavat' in desc_lower:
                        subcategory = 'Excavation'
                    elif 'concrete' in desc_lower:
                        subcategory = 'Concrete'
                    elif 'support' in desc_lower:
                        subcategory = 'Support Systems'
                
                item = {
                    'id': code,
                    'code': code,
                    'description': description,
                    'unit': self.standardize_unit(unit),
                    'category': sheet_name.replace(' works', ' Works'),
                    'subcategory': subcategory,
                    'rate': rate,
                    'cellRate_reference': self.get_cell_reference(sheet_name, idx, rate_col) if rate_col else None,
                    'cellRate_rate': rate,
                    'excelCellReference': self.get_cell_reference(sheet_name, idx, 0),
                    'sourceSheetName': sheet_name,
                    'keywords': []
                }
                items.append(item)
            
            print(f"  Extracted {len(items)} items from {sheet_name}")
            return items
            
        except Exception as e:
            print(f"  Error processing {sheet_name}: {e}")
            return []
    
    def extract_all(self):
        """Extract all sheets with high quality"""
        print("="*80)
        print("HIGH-QUALITY COMPREHENSIVE PRICELIST EXTRACTION")
        print("="*80)
        
        all_items = []
        
        # Standard sheets
        sheets = ['Groundworks', 'RC works', 'Services', 'External Works', 'Underpinning']
        for sheet in sheets:
            items = self.extract_standard_sheet(sheet)
            all_items.extend(items)
        
        # Drainage with special range handling
        drainage_items = self.extract_drainage_with_ranges()
        all_items.extend(drainage_items)
        
        self.all_items = all_items
        return all_items
    
    def generate_keywords(self):
        """Generate keywords for all items"""
        for item in self.all_items:
            keywords = []
            desc = item['description'].lower()
            
            # Extract measurements
            measurements = re.findall(r'\d+(?:mm|m|kg|tonnes?)\b', desc)
            keywords.extend(measurements[:2])
            
            # Extract key terms based on category
            if item['category'] == 'Drainage':
                terms = ['pipe', 'manhole', 'gully', 'excavate', 'trench']
            elif item['category'] == 'Groundworks':
                terms = ['excavate', 'fill', 'disposal', 'demolish', 'compact']
            elif item['category'] == 'RC Works':
                terms = ['concrete', 'reinforcement', 'formwork', 'rebar', 'mesh']
            else:
                terms = []
            
            for term in terms:
                if term in desc:
                    keywords.append(term)
            
            item['keywords'] = keywords[:5]
    
    def save_outputs(self, prefix='high_quality_pricelist'):
        """Save the extracted data"""
        if not self.all_items:
            print("No items to save!")
            return
        
        # Generate keywords
        self.generate_keywords()
        
        print(f"\nSaving {len(self.all_items)} items...")
        
        # Save JSON
        json_file = f"{prefix}.json"
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(self.all_items, f, indent=2, ensure_ascii=False)
        print(f"Saved JSON: {json_file}")
        
        # Save CSV
        df = pd.DataFrame(self.all_items)
        df['keywords'] = df['keywords'].apply(lambda x: '|'.join(x) if isinstance(x, list) else '')
        csv_file = f"{prefix}.csv"
        df.to_csv(csv_file, index=False)
        print(f"Saved CSV: {csv_file}")
        
        return json_file, csv_file
    
    def show_statistics(self):
        """Display extraction statistics"""
        if not self.all_items:
            return
        
        print("\n" + "="*80)
        print("EXTRACTION STATISTICS")
        print("="*80)
        
        df = pd.DataFrame(self.all_items)
        
        print(f"\nTotal items: {len(self.all_items)}")
        print(f"Items with rates: {sum(1 for i in self.all_items if i['rate'] > 0)}")
        print(f"Items with cell references: {sum(1 for i in self.all_items if i['cellRate_reference'])}")
        
        print("\nBreakdown by category:")
        for cat in df['category'].unique():
            count = len(df[df['category'] == cat])
            with_rates = len(df[(df['category'] == cat) & (df['rate'] > 0)])
            print(f"  {cat:20} - {count:5} items ({with_rates} with rates)")
        
        # Show sample Drainage range items
        print("\nSample Drainage excavation items with proper descriptions:")
        drainage_excavation = df[(df['category'] == 'Drainage') & 
                                 (df['description'].str.contains('depth to invert:', case=False, na=False))]
        
        for _, item in drainage_excavation.head(5).iterrows():
            print(f"\n  Code: {item['code']}")
            print(f"  Description: {item['description'][:100]}...")
            print(f"  Unit: {item['unit']}")
            print(f"  Rate: {item['rate']}")
            print(f"  Cell: {item['cellRate_reference']}")

def main():
    extractor = HighQualityExtractor()
    
    # Extract all items
    items = extractor.extract_all()
    
    if items:
        # Save outputs
        extractor.save_outputs()
        
        # Show statistics
        extractor.show_statistics()
        
        print("\n[OK] High-quality extraction complete!")
        print("\nFeatures:")
        print("  - Proper Drainage range descriptions from headers")
        print("  - Original descriptions preserved")
        print("  - All cell references in SheetName!Cell format")
        print("  - Actual Excel codes used for ID and code")
    else:
        print("\n[X] No items extracted!")

if __name__ == "__main__":
    main()