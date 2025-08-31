"""
Enhanced Full Pricelist Extraction with Drainage Range Support
Properly extracts range-based items like "0.5-0.75" with their headers
"""

import pandas as pd
import numpy as np
import json
import re
import string
from pathlib import Path

class EnhancedPricelistExtractor:
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
            float(str_val)
            return True
        except:
            if len(str_val) < 50:
                return True
        
        return False
    
    def is_unit(self, value):
        """Check if value is a unit"""
        if pd.isna(value):
            return False
        
        str_val = str(value).strip().lower()
        
        units = ['m', 'm2', 'm²', 'm3', 'm³', 'mm', 'nr', 'no', 'item', 'sum',
                 'kg', 'tonnes', 't', 'lm', 'sqm', 'cum', 'each', 'set',
                 'l.s.', 'ls', 'hour', 'hr', 'day', 'week', 'month',
                 'lin.m', 'sq.m', 'cu.m', 'no.', 'l/s']
        
        if str_val in units:
            return True
        
        if any(unit in str_val for unit in units):
            return True
            
        return False
    
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
        
        unit_lower = str(unit).lower().strip()
        return unit_map.get(unit_lower, unit_lower)
    
    def extract_drainage_enhanced(self):
        """Enhanced extraction for Drainage sheet with range support"""
        sheet_name = 'Drainage'
        print(f"\nProcessing {sheet_name} (Enhanced)...")
        
        try:
            df = pd.read_excel(self.excel_file, sheet_name=sheet_name, header=None)
            print(f"  Sheet has {len(df)} rows")
            
            items = []
            current_header = None
            
            for idx, row in df.iterrows():
                row_vals = row.values
                
                # Check if this is an excavation header
                first_val = str(row_vals[0]) if pd.notna(row_vals[0]) else ''
                if 'excavat' in first_val.lower() or 'trench' in first_val.lower():
                    current_header = first_val.strip()
                    continue
                
                # Check for range pattern in columns 2-4 (e.g., "0.5 - 0.75")
                is_range_row = False
                range_str = ""
                
                # Pattern 1: separate columns (e.g., col2=0.5, col3='-', col4=0.75)
                if (pd.notna(row_vals[2]) and pd.notna(row_vals[3]) and pd.notna(row_vals[4]) and
                    str(row_vals[3]).strip() == '-'):
                    try:
                        val1 = float(str(row_vals[2]))
                        val2 = float(str(row_vals[4]))
                        range_str = f"{val1}-{val2}"
                        is_range_row = True
                    except:
                        pass
                
                # Pattern 2: combined in one column
                for col in range(1, min(5, len(row_vals))):
                    if pd.notna(row_vals[col]):
                        val_str = str(row_vals[col]).strip()
                        if re.match(r'^[\d.]+\s*-\s*[\d.]+$', val_str):
                            range_str = val_str
                            is_range_row = True
                            break
                
                if is_range_row:
                    # Extract code
                    code = str(row_vals[0]) if pd.notna(row_vals[0]) else None
                    if not code or not self.is_valid_code(code):
                        continue
                    
                    # Build description
                    if current_header:
                        # Add range to header
                        description = f"{current_header} Depth: {range_str}m"
                    else:
                        description = f"Excavation depth: {range_str}m"
                    
                    # Find unit (usually 'm' for linear meter of trench)
                    unit = 'm'
                    for col in range(5, min(10, len(row_vals))):
                        if pd.notna(row_vals[col]) and self.is_unit(row_vals[col]):
                            unit = str(row_vals[col]).strip()
                            break
                    
                    # Find rate - typically in columns 10-20
                    rate = 0
                    rate_col = None
                    for col in range(10, min(20, len(row_vals))):
                        if pd.notna(row_vals[col]):
                            try:
                                val = float(str(row_vals[col]).replace(',', '').replace('£', ''))
                                if 0 < val < 10000:
                                    rate = val
                                    rate_col = col
                                    break
                            except:
                                continue
                    
                    # Create item
                    item = {
                        'id': code,
                        'code': code,
                        'description': description,
                        'unit': self.standardize_unit(unit),
                        'category': sheet_name,
                        'subcategory': 'Excavation',
                        'rate': rate,
                        'cellRate_reference': f"{sheet_name}!{self.get_cell_reference(idx, rate_col)}" if rate_col else None,
                        'cellRate_rate': rate,
                        'excelCellReference': f"{sheet_name}!{self.get_cell_reference(idx, 0)}",
                        'sourceSheetName': sheet_name,
                        'keywords': ['excavation', 'trench', f'depth_{range_str}m']
                    }
                    
                    items.append(item)
                
                # Also handle normal drainage items
                elif self.is_valid_code(row_vals[0]):
                    code = str(row_vals[0]).strip()
                    
                    # Extract description
                    description = ""
                    for col in range(1, min(5, len(row_vals))):
                        if pd.notna(row_vals[col]) and not self.is_unit(row_vals[col]):
                            val = str(row_vals[col]).strip()
                            try:
                                float(val.replace(',', ''))
                                if float(val.replace(',', '')) > 10:
                                    continue
                            except:
                                if val and len(val) > 2:
                                    description += " " + val
                    
                    description = description.strip()
                    if not description or len(description) < 3:
                        continue
                    
                    # Find unit
                    unit = 'item'
                    for col in range(2, min(10, len(row_vals))):
                        if pd.notna(row_vals[col]) and self.is_unit(row_vals[col]):
                            unit = str(row_vals[col]).strip()
                            break
                    
                    # Find rate
                    rate = 0
                    rate_col = None
                    for col in range(3, min(20, len(row_vals))):
                        if pd.notna(row_vals[col]):
                            try:
                                val = float(str(row_vals[col]).replace(',', '').replace('£', ''))
                                if 0 < val < 100000:
                                    rate = val
                                    rate_col = col
                                    break
                            except:
                                continue
                    
                    # Determine subcategory
                    desc_lower = description.lower()
                    if 'pipe' in desc_lower:
                        subcategory = 'Pipes'
                    elif 'manhole' in desc_lower:
                        subcategory = 'Manholes'
                    elif 'gully' in desc_lower:
                        subcategory = 'Gullies'
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
                        'cellRate_reference': f"{sheet_name}!{self.get_cell_reference(idx, rate_col)}" if rate_col else None,
                        'cellRate_rate': rate,
                        'excelCellReference': f"{sheet_name}!{self.get_cell_reference(idx, 0)}",
                        'sourceSheetName': sheet_name,
                        'keywords': []
                    }
                    
                    items.append(item)
            
            print(f"  Extracted {len(items)} items from {sheet_name} (including range items)")
            return items
            
        except Exception as e:
            print(f"  Error processing {sheet_name}: {e}")
            import traceback
            traceback.print_exc()
            return []
    
    def extract_standard_sheet(self, sheet_name):
        """Standard extraction for other sheets"""
        print(f"\nProcessing {sheet_name}...")
        
        try:
            df = pd.read_excel(self.excel_file, sheet_name=sheet_name, header=None)
            print(f"  Sheet has {len(df)} rows")
            
            items = []
            
            for idx, row in df.iterrows():
                if row.notna().sum() < 2:
                    continue
                
                # Extract code
                code = None
                if pd.notna(row[0]):
                    code_val = str(row[0]).strip()
                    if self.is_valid_code(code_val):
                        code = code_val
                
                if not code:
                    continue
                
                # Extract description
                description = ""
                for col in range(1, min(5, len(row))):
                    if pd.notna(row[col]) and not self.is_unit(row[col]):
                        val = str(row[col]).strip()
                        try:
                            if float(val.replace(',', '')) > 10:
                                continue
                        except:
                            if val and len(val) > 2:
                                description += " " + val
                
                description = description.strip()
                if not description or len(description) < 3:
                    continue
                
                # Find unit
                unit = 'item'
                for col in range(2, min(10, len(row))):
                    if pd.notna(row[col]) and self.is_unit(row[col]):
                        unit = str(row[col]).strip()
                        break
                
                # Find rate
                rate = 0
                rate_col = None
                for col in range(3, min(20, len(row))):
                    if pd.notna(row[col]):
                        try:
                            val = float(str(row[col]).replace(',', '').replace('£', ''))
                            if 0 < val < 100000:
                                rate = val
                                rate_col = col
                                break
                        except:
                            continue
                
                # Determine subcategory
                subcategory = sheet_name
                desc_lower = description.lower()
                
                if sheet_name == 'Groundworks':
                    if 'demolish' in desc_lower:
                        subcategory = 'Demolition'
                    elif 'excavat' in desc_lower:
                        subcategory = 'Excavation'
                    elif 'fill' in desc_lower:
                        subcategory = 'Filling'
                elif sheet_name == 'RC works':
                    if 'concrete' in desc_lower:
                        subcategory = 'Concrete'
                    elif 'reinforcement' in desc_lower:
                        subcategory = 'Reinforcement'
                    elif 'formwork' in desc_lower:
                        subcategory = 'Formwork'
                
                item = {
                    'id': code,
                    'code': code,
                    'description': description,
                    'unit': self.standardize_unit(unit),
                    'category': sheet_name.replace(' works', ' Works') if 'works' in sheet_name.lower() else sheet_name,
                    'subcategory': subcategory,
                    'rate': rate,
                    'cellRate_reference': f"{sheet_name}!{self.get_cell_reference(idx, rate_col)}" if rate_col else None,
                    'cellRate_rate': rate,
                    'excelCellReference': f"{sheet_name}!{self.get_cell_reference(idx, 0)}",
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
        """Extract from all sheets with enhanced Drainage support"""
        print("="*80)
        print("ENHANCED COMPREHENSIVE PRICELIST EXTRACTION")
        print("="*80)
        
        all_items = []
        
        # Standard sheets
        standard_sheets = ['Groundworks', 'RC works', 'Services', 'External Works', 'Underpinning']
        for sheet in standard_sheets:
            items = self.extract_standard_sheet(sheet)
            all_items.extend(items)
        
        # Enhanced Drainage extraction
        drainage_items = self.extract_drainage_enhanced()
        all_items.extend(drainage_items)
        
        self.all_items = all_items
        return all_items
    
    def save_outputs(self, prefix='full_pricelist_enhanced'):
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
        df['keywords'] = df['keywords'].apply(lambda x: '|'.join(x) if isinstance(x, list) else '')
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
        
        # Show some range items from Drainage
        print("\nSample Drainage range items:")
        drainage_items = [i for i in self.all_items if i['category'] == 'Drainage' and 'depth_' in str(i.get('keywords', []))]
        for item in drainage_items[:5]:
            print(f"\n  Code: {item['code']}")
            print(f"  Desc: {item['description'][:80]}...")
            print(f"  Unit: {item['unit']}")
            print(f"  Rate: {item['rate']}")
            print(f"  Cell: {item['cellRate_reference']}")

def main():
    extractor = EnhancedPricelistExtractor()
    
    # Extract all items
    items = extractor.extract_all()
    
    if items:
        # Save outputs
        extractor.save_outputs()
        
        # Show statistics
        extractor.show_statistics()
        
        print("\n[OK] Enhanced extraction complete!")
        print("Now includes Drainage range items (0.5-0.75, etc.) with proper headers!")
    else:
        print("\n[X] No items extracted!")

if __name__ == "__main__":
    main()