"""
Fixed Master Extraction Script
- Uses actual Excel codes for id and code
- Cell references include sheet name (e.g., "Groundworks!F20")
- Removed work_type and original_code fields
- Fixed all extraction issues
"""

import pandas as pd
import numpy as np
import json
import re
import string
from datetime import datetime
from pathlib import Path

class UnifiedExtractor:
    """Single extractor that handles all sheets with their specific logic"""
    
    def __init__(self, excel_file='MJD-PRICELIST.xlsx'):
        self.excel_file = excel_file
        self.all_items = []
        
    def get_cell_reference(self, row_idx, col_idx):
        """Convert row and column index to Excel cell reference"""
        if col_idx < 26:
            col_letter = string.ascii_uppercase[col_idx]
        else:
            col_letter = string.ascii_uppercase[col_idx // 26 - 1] + string.ascii_uppercase[col_idx % 26]
        return f"{col_letter}{row_idx + 1}"
    
    def get_sheet_cell_reference(self, sheet_name, row_idx, col_idx):
        """Get full cell reference with sheet name"""
        cell_ref = self.get_cell_reference(row_idx, col_idx)
        return f"{sheet_name}!{cell_ref}"
    
    def is_unit(self, value):
        """Check if value is a unit"""
        if pd.isna(value) or value is None:
            return False
        
        value_str = str(value).strip().lower()
        
        # Don't treat numbers as units
        try:
            float(value_str.replace(',', ''))
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
        
        unit_lower = str(unit).lower().strip()
        return unit_map.get(unit_lower, unit_lower)
    
    def extract_code(self, row, col_idx=0):
        """Extract actual code from Excel"""
        if col_idx < len(row) and pd.notna(row[col_idx]):
            code = str(row[col_idx]).strip()
            if code and code.lower() not in ['nan', 'none', '']:
                # Keep the actual Excel value
                return code
        return None
    
    def extract_rate(self, row, start_col=3, end_col=15):
        """Extract rate value and column index"""
        for col_idx in range(start_col, min(end_col, len(row))):
            if pd.notna(row[col_idx]):
                value = str(row[col_idx]).strip()
                value_clean = value.replace(',', '').replace('£', '').replace('$', '')
                try:
                    rate = float(value_clean)
                    # Sanity check - rates should be reasonable
                    if 0 < rate < 1000000:
                        return rate, col_idx
                except:
                    continue
        return None, None
    
    def extract_description(self, row, start_col=1, max_cols=5):
        """Extract and clean description"""
        description_parts = []
        
        for col_idx in range(start_col, min(start_col + max_cols, len(row))):
            if pd.notna(row[col_idx]):
                part = str(row[col_idx]).strip()
                # Skip if it's just a number or unit
                if not re.match(r'^[\d,\.]+$', part) and not self.is_unit(part):
                    # Skip if it looks like a rate
                    try:
                        val = float(part.replace(',', '').replace('£', ''))
                        if val > 10:
                            continue
                    except:
                        pass
                    description_parts.append(part)
        
        description = ' '.join(description_parts)
        
        # Clean common abbreviations
        replacements = {
            ' exc ': ' excavation ', ' ne ': ' not exceeding ', ' n.e. ': ' not exceeding ',
            ' thk ': ' thick ', ' dia ': ' diameter ', ' incl ': ' including ',
            ' excl ': ' excluding ', ' c/w ': ' complete with ', ' conc ': ' concrete ',
            ' reinf ': ' reinforcement ', ' fwk ': ' formwork ', ' rc ': ' reinforced concrete ',
        }
        
        for old, new in replacements.items():
            description = description.replace(old, new)
        
        return ' '.join(description.split())
    
    def extract_groundworks(self):
        """Extract Groundworks sheet"""
        sheet_name = 'Groundworks'
        print(f"\nProcessing {sheet_name}...")
        
        try:
            df = pd.read_excel(self.excel_file, sheet_name=sheet_name, header=None)
            print(f"Loaded {len(df)} rows")
            
            items = []
            for idx, row in df.iterrows():
                # Skip empty rows
                if row.notna().sum() < 2:
                    continue
                
                code = self.extract_code(row)
                if not code:
                    continue
                
                description = self.extract_description(row)
                if not description or len(description) < 5:
                    continue
                
                # Find unit
                unit = 'item'
                for col_idx in range(2, min(5, len(row))):
                    if pd.notna(row[col_idx]) and self.is_unit(row[col_idx]):
                        unit = str(row[col_idx]).strip()
                        break
                
                # Extract rate
                rate, rate_col_idx = self.extract_rate(row)
                
                # Determine subcategory
                desc_lower = description.lower()
                if 'excavat' in desc_lower:
                    subcategory = 'Excavation'
                elif 'fill' in desc_lower:
                    subcategory = 'Filling'
                elif 'disposal' in desc_lower:
                    subcategory = 'Disposal'
                else:
                    subcategory = 'Groundworks'
                
                # Create item
                item = {
                    'id': code,
                    'code': code,
                    'description': description,
                    'unit': self.standardize_unit(unit),
                    'category': sheet_name,
                    'subcategory': subcategory,
                    'rate': rate if rate else 0,
                    'cellRate_reference': self.get_sheet_cell_reference(sheet_name, idx, rate_col_idx) if rate_col_idx else None,
                    'cellRate_rate': rate if rate else 0,
                    'excelCellReference': self.get_sheet_cell_reference(sheet_name, idx, 0),
                    'sourceSheetName': sheet_name,
                    'keywords': []
                }
                items.append(item)
            
            print(f"Extracted {len(items)} items from {sheet_name}")
            return items
            
        except Exception as e:
            print(f"Error processing {sheet_name}: {e}")
            return []
    
    def extract_rc_works(self):
        """Extract RC Works sheet"""
        sheet_name = 'RC works'
        print(f"\nProcessing {sheet_name}...")
        
        try:
            df = pd.read_excel(self.excel_file, sheet_name=sheet_name, header=None)
            print(f"Loaded {len(df)} rows")
            
            items = []
            for idx, row in df.iterrows():
                if row.notna().sum() < 2:
                    continue
                
                code = self.extract_code(row)
                if not code:
                    continue
                
                description = self.extract_description(row)
                if not description or len(description) < 5:
                    continue
                
                # Find unit
                unit = 'item'
                for col_idx in range(2, min(5, len(row))):
                    if pd.notna(row[col_idx]) and self.is_unit(row[col_idx]):
                        unit = str(row[col_idx]).strip()
                        break
                
                # Extract rate
                rate, rate_col_idx = self.extract_rate(row)
                
                # Determine subcategory
                desc_lower = description.lower()
                if 'concrete' in desc_lower:
                    subcategory = 'Concrete Works'
                elif 'reinforcement' in desc_lower or 'rebar' in desc_lower:
                    subcategory = 'Reinforcement'
                elif 'formwork' in desc_lower:
                    subcategory = 'Formwork'
                else:
                    subcategory = 'RC Works'
                
                item = {
                    'id': code,
                    'code': code,
                    'description': description,
                    'unit': self.standardize_unit(unit),
                    'category': 'RC Works',
                    'subcategory': subcategory,
                    'rate': rate if rate else 0,
                    'cellRate_reference': self.get_sheet_cell_reference(sheet_name, idx, rate_col_idx) if rate_col_idx else None,
                    'cellRate_rate': rate if rate else 0,
                    'excelCellReference': self.get_sheet_cell_reference(sheet_name, idx, 0),
                    'sourceSheetName': sheet_name,
                    'keywords': []
                }
                items.append(item)
            
            print(f"Extracted {len(items)} items from RC Works")
            return items
            
        except Exception as e:
            print(f"Error processing RC Works: {e}")
            return []
    
    def extract_drainage(self):
        """Extract Drainage sheet"""
        sheet_name = 'Drainage'
        print(f"\nProcessing {sheet_name}...")
        
        try:
            df = pd.read_excel(self.excel_file, sheet_name=sheet_name, header=None)
            print(f"Loaded {len(df)} rows")
            
            items = []
            for idx, row in df.iterrows():
                if row.notna().sum() < 2:
                    continue
                
                code = self.extract_code(row)
                if not code:
                    continue
                
                description = self.extract_description(row)
                if not description or len(description) < 5:
                    continue
                
                # Find unit
                unit = 'item'
                for col_idx in range(2, min(5, len(row))):
                    if pd.notna(row[col_idx]) and self.is_unit(row[col_idx]):
                        unit = str(row[col_idx]).strip()
                        break
                
                # Extract rate
                rate, rate_col_idx = self.extract_rate(row)
                
                # Determine subcategory
                desc_lower = description.lower()
                if 'pipe' in desc_lower:
                    subcategory = 'Pipes'
                elif 'manhole' in desc_lower:
                    subcategory = 'Manholes'
                elif 'gully' in desc_lower:
                    subcategory = 'Gullies'
                else:
                    subcategory = 'Drainage'
                
                item = {
                    'id': code,
                    'code': code,
                    'description': description,
                    'unit': self.standardize_unit(unit),
                    'category': sheet_name,
                    'subcategory': subcategory,
                    'rate': rate if rate else 0,
                    'cellRate_reference': self.get_sheet_cell_reference(sheet_name, idx, rate_col_idx) if rate_col_idx else None,
                    'cellRate_rate': rate if rate else 0,
                    'excelCellReference': self.get_sheet_cell_reference(sheet_name, idx, 0),
                    'sourceSheetName': sheet_name,
                    'keywords': []
                }
                items.append(item)
            
            print(f"Extracted {len(items)} items from {sheet_name}")
            return items
            
        except Exception as e:
            print(f"Error processing {sheet_name}: {e}")
            return []
    
    def extract_services(self):
        """Extract Services sheet"""
        sheet_name = 'Services'
        print(f"\nProcessing {sheet_name}...")
        
        try:
            df = pd.read_excel(self.excel_file, sheet_name=sheet_name, header=None)
            print(f"Loaded {len(df)} rows")
            
            items = []
            for idx, row in df.iterrows():
                if row.notna().sum() < 2:
                    continue
                
                code = self.extract_code(row)
                if not code:
                    continue
                
                description = self.extract_description(row)
                if not description or len(description) < 5:
                    continue
                
                # Find unit
                unit = 'item'
                for col_idx in range(2, min(5, len(row))):
                    if pd.notna(row[col_idx]) and self.is_unit(row[col_idx]):
                        unit = str(row[col_idx]).strip()
                        break
                
                # Extract rate
                rate, rate_col_idx = self.extract_rate(row)
                
                # Determine subcategory
                desc_lower = description.lower()
                if 'electrical' in desc_lower or 'cable' in desc_lower:
                    subcategory = 'Electrical'
                elif 'plumbing' in desc_lower or 'water' in desc_lower:
                    subcategory = 'Plumbing'
                elif 'hvac' in desc_lower or 'air' in desc_lower:
                    subcategory = 'HVAC'
                else:
                    subcategory = 'Services'
                
                item = {
                    'id': code,
                    'code': code,
                    'description': description,
                    'unit': self.standardize_unit(unit),
                    'category': sheet_name,
                    'subcategory': subcategory,
                    'rate': rate if rate else 0,
                    'cellRate_reference': self.get_sheet_cell_reference(sheet_name, idx, rate_col_idx) if rate_col_idx else None,
                    'cellRate_rate': rate if rate else 0,
                    'excelCellReference': self.get_sheet_cell_reference(sheet_name, idx, 0),
                    'sourceSheetName': sheet_name,
                    'keywords': []
                }
                items.append(item)
            
            print(f"Extracted {len(items)} items from {sheet_name}")
            return items
            
        except Exception as e:
            print(f"Error processing {sheet_name}: {e}")
            return []
    
    def extract_external_works(self):
        """Extract External Works sheet"""
        sheet_name = 'External Works'
        print(f"\nProcessing {sheet_name}...")
        
        try:
            df = pd.read_excel(self.excel_file, sheet_name=sheet_name, header=None)
            print(f"Loaded {len(df)} rows")
            
            items = []
            for idx, row in df.iterrows():
                if row.notna().sum() < 2:
                    continue
                
                code = self.extract_code(row)
                if not code:
                    continue
                
                description = self.extract_description(row)
                if not description or len(description) < 5:
                    continue
                
                # Find unit
                unit = 'item'
                for col_idx in range(2, min(5, len(row))):
                    if pd.notna(row[col_idx]) and self.is_unit(row[col_idx]):
                        unit = str(row[col_idx]).strip()
                        break
                
                # Extract rate
                rate, rate_col_idx = self.extract_rate(row)
                
                # Determine subcategory
                desc_lower = description.lower()
                if 'paving' in desc_lower:
                    subcategory = 'Paving'
                elif 'kerb' in desc_lower:
                    subcategory = 'Kerbs'
                elif 'fence' in desc_lower:
                    subcategory = 'Fencing'
                else:
                    subcategory = 'External Works'
                
                item = {
                    'id': code,
                    'code': code,
                    'description': description,
                    'unit': self.standardize_unit(unit),
                    'category': sheet_name,
                    'subcategory': subcategory,
                    'rate': rate if rate else 0,
                    'cellRate_reference': self.get_sheet_cell_reference(sheet_name, idx, rate_col_idx) if rate_col_idx else None,
                    'cellRate_rate': rate if rate else 0,
                    'excelCellReference': self.get_sheet_cell_reference(sheet_name, idx, 0),
                    'sourceSheetName': sheet_name,
                    'keywords': []
                }
                items.append(item)
            
            print(f"Extracted {len(items)} items from {sheet_name}")
            return items
            
        except Exception as e:
            print(f"Error processing {sheet_name}: {e}")
            return []
    
    def extract_underpinning(self):
        """Extract Underpinning sheet"""
        sheet_name = 'Underpinning'
        print(f"\nProcessing {sheet_name}...")
        
        try:
            df = pd.read_excel(self.excel_file, sheet_name=sheet_name, header=None)
            print(f"Loaded {len(df)} rows")
            
            items = []
            for idx, row in df.iterrows():
                if row.notna().sum() < 2:
                    continue
                
                code = self.extract_code(row)
                if not code:
                    continue
                
                description = self.extract_description(row)
                if not description or len(description) < 5:
                    continue
                
                # Find unit
                unit = 'item'
                for col_idx in range(2, min(5, len(row))):
                    if pd.notna(row[col_idx]) and self.is_unit(row[col_idx]):
                        unit = str(row[col_idx]).strip()
                        break
                
                # Extract rate
                rate, rate_col_idx = self.extract_rate(row)
                
                # Determine subcategory
                desc_lower = description.lower()
                if 'excavat' in desc_lower:
                    subcategory = 'Excavation'
                elif 'concrete' in desc_lower:
                    subcategory = 'Concrete'
                elif 'support' in desc_lower:
                    subcategory = 'Support'
                else:
                    subcategory = 'Underpinning'
                
                item = {
                    'id': code,
                    'code': code,
                    'description': description,
                    'unit': self.standardize_unit(unit),
                    'category': sheet_name,
                    'subcategory': subcategory,
                    'rate': rate if rate else 0,
                    'cellRate_reference': self.get_sheet_cell_reference(sheet_name, idx, rate_col_idx) if rate_col_idx else None,
                    'cellRate_rate': rate if rate else 0,
                    'excelCellReference': self.get_sheet_cell_reference(sheet_name, idx, 0),
                    'sourceSheetName': sheet_name,
                    'keywords': []
                }
                items.append(item)
            
            print(f"Extracted {len(items)} items from {sheet_name}")
            return items
            
        except Exception as e:
            print(f"Error processing {sheet_name}: {e}")
            return []
    
    def extract_all(self):
        """Extract from all sheets"""
        print("="*80)
        print("FIXED MASTER PRICELIST EXTRACTION")
        print("="*80)
        
        all_items = []
        
        # Extract from each sheet
        all_items.extend(self.extract_groundworks())
        all_items.extend(self.extract_rc_works())
        all_items.extend(self.extract_drainage())
        all_items.extend(self.extract_services())
        all_items.extend(self.extract_external_works())
        all_items.extend(self.extract_underpinning())
        
        self.all_items = all_items
        return all_items
    
    def save_outputs(self, prefix='master_pricelist_fixed'):
        """Save the extracted data"""
        if not self.all_items:
            print("No items to save!")
            return
        
        print(f"\nSaving {len(self.all_items)} items...")
        
        # Save JSON
        json_file = f"{prefix}.json"
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(self.all_items, f, indent=2, ensure_ascii=False)
        print(f"[OK] Saved JSON: {json_file}")
        
        # Save CSV
        df = pd.DataFrame(self.all_items)
        df['keywords'] = df['keywords'].apply(lambda x: '|'.join(x) if isinstance(x, list) else '')
        csv_file = f"{prefix}.csv"
        df.to_csv(csv_file, index=False)
        print(f"[OK] Saved CSV: {csv_file}")
        
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
        print(f"Items with cell refs: {sum(1 for i in self.all_items if i['cellRate_reference'])}")
        
        print("\nItems per category:")
        for cat in df['category'].unique():
            count = len(df[df['category'] == cat])
            print(f"  {cat:20} - {count:4} items")
        
        print("\nSample items with cell references:")
        sample = [i for i in self.all_items if i['cellRate_reference']][:5]
        for item in sample:
            print(f"\n  ID: {item['id']}")
            print(f"  Description: {item['description'][:50]}...")
            print(f"  Rate: {item['rate']}")
            print(f"  Cell Ref: {item['cellRate_reference']}")

def main():
    """Main execution"""
    extractor = UnifiedExtractor()
    
    # Extract all items
    items = extractor.extract_all()
    
    if items:
        # Save outputs
        extractor.save_outputs()
        
        # Show statistics
        extractor.show_statistics()
        
        print("\n[OK] Extraction complete!")
        print("\nOutput format:")
        print("  - id: Actual Excel code")
        print("  - code: Same as id")
        print("  - cellRate_reference: Sheet!Cell format (e.g., Groundworks!F20)")
        print("  - No work_type or original_code fields")
    else:
        print("\n[X] No items extracted!")

if __name__ == "__main__":
    main()