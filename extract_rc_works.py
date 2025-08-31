"""
Extraction script for RC Works sheet
Handles the specific structure and format of the Reinforced Concrete works pricelist
"""

import pandas as pd
import numpy as np
import json
import re
from datetime import datetime
from pathlib import Path
import string

class RCWorksExtractor:
    def __init__(self, excel_file='MJD-PRICELIST.xlsx'):
        self.excel_file = excel_file
        self.sheet_name = 'RC works'
        self.df = None
        self.extracted_items = []
        
    def load_sheet(self):
        """Load the RC Works sheet"""
        print(f"Loading {self.sheet_name} sheet...")
        self.df = pd.read_excel(self.excel_file, sheet_name=self.sheet_name, header=None)
        print(f"Loaded {len(self.df)} rows x {len(self.df.columns)} columns")
        return self.df
    
    def identify_data_rows(self):
        """Identify rows containing actual pricelist data"""
        data_rows = []
        
        for idx, row in self.df.iterrows():
            # Skip if row is mostly empty
            if row.notna().sum() < 3:
                continue
                
            # Look for patterns that indicate RC works data
            # Check if first column might be a code
            code_col = row[0] if 0 < len(row) else None
            if pd.notna(code_col):
                code_str = str(code_col).strip()
                # RC codes often start with RC, C, or numbers
                if (re.match(r'^\d+', code_str) or 
                    re.match(r'^RC', code_str, re.I) or
                    re.match(r'^C\d+', code_str, re.I) or
                    re.match(r'^[A-Z]+\d+', code_str)):
                    data_rows.append(idx)
                    continue
            
            # Check if row has RC-related content
            for col_idx in range(1, min(5, len(row))):
                cell = row[col_idx]
                if pd.notna(cell):
                    cell_str = str(cell).strip().lower()
                    # RC works keywords
                    if any(keyword in cell_str for keyword in 
                           ['concrete', 'reinforce', 'rebar', 'formwork', 'shutter',
                            'pour', 'cast', 'slab', 'beam', 'column', 'wall',
                            'foundation', 'mesh', 'bar', 'stirrup', 'link',
                            'grade', 'c25', 'c30', 'c35', 'c40', 'cement',
                            'aggregate', 'vibrat', 'cure', 'finish']):
                        data_rows.append(idx)
                        break
        
        return data_rows
    
    def extract_code(self, row, col_idx=0):
        """Extract code from row"""
        if col_idx < len(row) and pd.notna(row[col_idx]):
            code = str(row[col_idx]).strip()
            # Clean up code
            code = re.sub(r'\s+', '', code)
            if code and not code.lower() in ['nan', 'none', '-', '']:
                return code
        return None
    
    def extract_description(self, row, start_col=1):
        """Extract and clean description for RC works"""
        description_parts = []
        
        # Collect description from multiple columns
        for col_idx in range(start_col, min(start_col + 3, len(row))):
            if pd.notna(row[col_idx]):
                part = str(row[col_idx]).strip()
                # Skip if it's a number or unit
                if not re.match(r'^[\d,\.]+$', part) and not self.is_unit(part):
                    description_parts.append(part)
        
        description = ' '.join(description_parts)
        
        # Clean and expand RC-specific abbreviations
        replacements = {
            ' rc ': ' reinforced concrete ',
            ' r.c. ': ' reinforced concrete ',
            ' conc ': ' concrete ',
            ' reinf ': ' reinforcement ',
            ' fwk ': ' formwork ',
            ' ne ': ' not exceeding ',
            ' n.e. ': ' not exceeding ',
            ' thk ': ' thick ',
            ' dia ': ' diameter ',
            ' c/c ': ' centers ',
            ' bwys ': ' both ways ',
            ' ew ': ' each way ',
            ' t&b ': ' top and bottom ',
            ' u/s ': ' underside ',
            ' o/a ': ' overall ',
            ' incl ': ' including ',
            ' excl ': ' excluding ',
            ' horiz ': ' horizontal ',
            ' vert ': ' vertical ',
            ' adj ': ' adjacent ',
            ' struct ': ' structural ',
            ' fdn ': ' foundation ',
            ' col ': ' column ',
            ' bm ': ' beam ',
            ' slab ': ' slab ',
            ' r/f ': ' reinforcement ',
            ' ms ': ' mild steel ',
            ' hy ': ' high yield ',
            ' tor ': ' deformed bars ',
            ' cwk ': ' concrete work ',
            ' sfwk ': ' soffit formwork ',
            ' vfwk ': ' vertical formwork ',
            ' c25/30': ' Grade C25/30 ',
            ' c30/37': ' Grade C30/37 ',
            ' c35/45': ' Grade C35/45 ',
            ' c40/50': ' Grade C40/50 ',
        }
        
        for old, new in replacements.items():
            description = description.replace(old, new)
            description = description.replace(old.upper(), new)
        
        # Fix patterns
        description = re.sub(r'(\d+)thk', r'\1mm thick', description)
        description = re.sub(r'(\d+)dia', r'\1mm diameter', description)
        description = re.sub(r'T(\d+)', r'\1mm diameter bars', description)
        description = re.sub(r'Y(\d+)', r'\1mm diameter high yield bars', description)
        description = re.sub(r'R(\d+)', r'\1mm diameter mild steel bars', description)
        
        # Fix concrete grades
        description = re.sub(r'\bC(\d+)\b', r'Grade C\1', description)
        description = re.sub(r'\bC(\d+)/(\d+)\b', r'Grade C\1/\2', description)
        
        # Clean up spaces
        description = ' '.join(description.split())
        
        return description
    
    def is_unit(self, value):
        """Check if value is a unit"""
        if pd.isna(value):
            return False
        
        value_str = str(value).strip().lower()
        
        units = ['m', 'm2', 'm²', 'm3', 'm³', 'nr', 'no', 'item', 'sum',
                 'kg', 'tonnes', 't', 'lm', 'sqm', 'cum', 'each', 'ton']
        
        return value_str in units
    
    def extract_unit(self, row, expected_col=None):
        """Extract unit from row"""
        # Try expected column first
        if expected_col is not None and expected_col < len(row):
            if pd.notna(row[expected_col]):
                value = str(row[expected_col]).strip()
                if self.is_unit(value):
                    return self.standardize_unit(value)
        
        # Search for unit in other columns
        for col_idx in range(2, min(6, len(row))):
            if pd.notna(row[col_idx]):
                value = str(row[col_idx]).strip()
                if self.is_unit(value):
                    return self.standardize_unit(value)
        
        # Infer from description
        return self.infer_unit_from_description(row)
    
    def standardize_unit(self, unit):
        """Standardize unit format"""
        unit_map = {
            'm2': 'm²', 'sqm': 'm²', 'sq.m': 'm²',
            'm3': 'm³', 'cum': 'm³', 'cu.m': 'm³',
            'no': 'nr', 'no.': 'nr', 'each': 'nr',
            't': 'tonnes', 'ton': 'tonnes', 'tonne': 'tonnes',
            'lm': 'm', 'lin.m': 'm', 'l.m': 'm',
        }
        
        unit_lower = unit.lower()
        return unit_map.get(unit_lower, unit_lower)
    
    def infer_unit_from_description(self, row):
        """Infer unit from description content for RC works"""
        desc = self.extract_description(row)
        desc_lower = desc.lower()
        
        # RC works specific patterns
        if 'reinforcement' in desc_lower or 'rebar' in desc_lower or 'steel' in desc_lower:
            if 'mesh' in desc_lower:
                return 'm²'
            return 'kg'
        elif 'concrete' in desc_lower:
            if any(word in desc_lower for word in ['slab', 'surface', 'topping', 'screed']):
                if 'thick' in desc_lower:
                    return 'm²'
            elif any(word in desc_lower for word in ['beam', 'column', 'wall', 'foundation']):
                return 'm³'
            return 'm³'
        elif 'formwork' in desc_lower or 'shutter' in desc_lower:
            if 'edge' in desc_lower or 'linear' in desc_lower:
                return 'm'
            return 'm²'
        elif any(word in desc_lower for word in ['bar', 'rod', 'dowel']):
            if 'each' in desc_lower or 'number' in desc_lower:
                return 'nr'
            return 'kg'
        elif 'mesh' in desc_lower:
            return 'm²'
        elif any(word in desc_lower for word in ['joint', 'groove', 'chase']):
            return 'm'
        
        return 'item'
    
    def extract_rate(self, row, start_col=3):
        """Extract rate value"""
        for col_idx in range(start_col, min(start_col + 4, len(row))):
            if pd.notna(row[col_idx]):
                value = str(row[col_idx]).strip()
                # Check if it's a number
                value_clean = value.replace(',', '').replace('£', '').replace('$', '')
                try:
                    rate = float(value_clean)
                    if rate > 0:  # Valid rate
                        return rate
                except:
                    continue
        return None
    
    def get_cell_reference(self, row_idx, col_idx):
        """Convert row and column index to Excel cell reference"""
        if col_idx < 26:
            col_letter = string.ascii_uppercase[col_idx]
        else:
            col_letter = string.ascii_uppercase[col_idx // 26 - 1] + string.ascii_uppercase[col_idx % 26]
        
        return f"{col_letter}{row_idx + 1}"
    
    def determine_subcategory(self, description):
        """Determine subcategory based on description for RC works"""
        desc_lower = description.lower()
        
        # RC works subcategories
        if 'concrete' in desc_lower:
            if 'foundation' in desc_lower:
                return 'Concrete Foundations'
            elif 'slab' in desc_lower:
                if 'ground' in desc_lower:
                    return 'Ground Slabs'
                elif 'suspend' in desc_lower:
                    return 'Suspended Slabs'
                else:
                    return 'Concrete Slabs'
            elif 'beam' in desc_lower:
                return 'Concrete Beams'
            elif 'column' in desc_lower:
                return 'Concrete Columns'
            elif 'wall' in desc_lower:
                if 'retaining' in desc_lower:
                    return 'Retaining Walls'
                else:
                    return 'Concrete Walls'
            elif 'stair' in desc_lower:
                return 'Concrete Stairs'
            else:
                return 'Concrete Works'
        elif 'reinforcement' in desc_lower or 'rebar' in desc_lower or 'steel' in desc_lower:
            if 'mesh' in desc_lower:
                return 'Mesh Reinforcement'
            elif 'bar' in desc_lower:
                return 'Bar Reinforcement'
            else:
                return 'Steel Reinforcement'
        elif 'formwork' in desc_lower or 'shutter' in desc_lower:
            if 'slab' in desc_lower or 'soffit' in desc_lower:
                return 'Slab Formwork'
            elif 'wall' in desc_lower or 'vertical' in desc_lower:
                return 'Wall Formwork'
            elif 'beam' in desc_lower:
                return 'Beam Formwork'
            elif 'column' in desc_lower:
                return 'Column Formwork'
            else:
                return 'Formwork'
        elif 'waterproof' in desc_lower:
            return 'Waterproofing'
        elif 'joint' in desc_lower:
            return 'Joints and Accessories'
        elif 'finish' in desc_lower:
            return 'Concrete Finishes'
        else:
            return 'General RC Works'
    
    def determine_work_type(self, description, subcategory):
        """Determine work type for RC works"""
        desc_lower = description.lower()
        
        if 'concrete' in desc_lower and 'pour' not in desc_lower:
            return 'Concrete'
        elif 'pour' in desc_lower or 'cast' in desc_lower:
            return 'Concrete Pouring'
        elif 'reinforcement' in desc_lower or 'rebar' in desc_lower or 'steel' in desc_lower:
            return 'Reinforcement'
        elif 'formwork' in desc_lower or 'shutter' in desc_lower:
            return 'Formwork'
        elif 'finish' in desc_lower:
            return 'Finishing'
        elif 'waterproof' in desc_lower:
            return 'Waterproofing'
        elif 'cure' in desc_lower or 'curing' in desc_lower:
            return 'Curing'
        else:
            return 'RC Works'
    
    def generate_keywords(self, description, subcategory):
        """Generate search keywords for RC works"""
        keywords = []
        desc_lower = description.lower()
        
        # Extract concrete grades
        grades = re.findall(r'c\d+(?:/\d+)?', desc_lower)
        keywords.extend(grades[:2])
        
        # Extract bar sizes
        bar_sizes = re.findall(r'\d+mm\s*(?:diameter|dia)', desc_lower)
        keywords.extend([size.replace(' ', '') for size in bar_sizes[:2]])
        
        # Extract thickness
        thickness = re.findall(r'\d+mm\s*thick', desc_lower)
        keywords.extend([t.replace(' ', '') for t in thickness[:1]])
        
        # Key RC terms
        terms = ['concrete', 'reinforcement', 'formwork', 'rebar', 'mesh',
                 'slab', 'beam', 'column', 'wall', 'foundation', 'pour',
                 'steel', 'bar', 'stirrup', 'link']
        
        for term in terms:
            if term in desc_lower:
                keywords.append(term)
        
        # Add subcategory keyword
        if subcategory:
            keywords.append(subcategory.lower().replace(' ', '_'))
        
        # Limit and remove duplicates
        seen = set()
        unique_keywords = []
        for kw in keywords:
            if kw not in seen:
                seen.add(kw)
                unique_keywords.append(kw)
        
        return unique_keywords[:6]
    
    def extract_items(self):
        """Main extraction method"""
        if self.df is None:
            self.load_sheet()
        
        print(f"\nExtracting items from {self.sheet_name}...")
        data_rows = self.identify_data_rows()
        print(f"Found {len(data_rows)} potential data rows")
        
        items = []
        current_id = 1
        
        for row_idx in data_rows:
            row = self.df.iloc[row_idx]
            
            # Extract basic fields
            code = self.extract_code(row)
            description = self.extract_description(row)
            
            # Skip if no valid description
            if not description or len(description) < 10:
                continue
            
            unit = self.extract_unit(row)
            rate = self.extract_rate(row)
            
            # Determine categories
            subcategory = self.determine_subcategory(description)
            work_type = self.determine_work_type(description, subcategory)
            
            # Generate keywords
            keywords = self.generate_keywords(description, subcategory)
            
            # Get cell references
            excel_ref = self.get_cell_reference(row_idx, 0)
            rate_cell_ref = None
            rate_value = None
            
            # Find rate cell reference
            for col_idx in range(3, min(7, len(row))):
                if pd.notna(row[col_idx]):
                    try:
                        value = float(str(row[col_idx]).replace(',', '').replace('£', ''))
                        if value > 0:
                            rate_cell_ref = self.get_cell_reference(row_idx, col_idx)
                            rate_value = value
                            break
                    except:
                        continue
            
            # Create item
            item = {
                'id': f"RC{current_id:04d}",
                'code': code if code else f"RC{current_id:04d}",
                'original_code': code,
                'description': description,
                'unit': unit,
                'category': 'RC Works',
                'subcategory': subcategory,
                'work_type': work_type,
                'rate': rate,
                'cellRate_reference': rate_cell_ref,
                'cellRate_rate': rate_value,
                'excelCellReference': excel_ref,
                'sourceSheetName': self.sheet_name,
                'keywords': keywords
            }
            
            items.append(item)
            current_id += 1
        
        self.extracted_items = items
        print(f"Extracted {len(items)} valid items from {self.sheet_name}")
        return items
    
    def save_output(self, output_prefix='rc_works'):
        """Save extracted data"""
        if not self.extracted_items:
            print("No items to save")
            return
        
        # Save JSON
        json_file = f"{output_prefix}_extracted.json"
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(self.extracted_items, f, indent=2, ensure_ascii=False)
        print(f"Saved JSON: {json_file}")
        
        # Save CSV
        df = pd.DataFrame(self.extracted_items)
        df['keywords'] = df['keywords'].apply(lambda x: '|'.join(x) if x else '')
        csv_file = f"{output_prefix}_extracted.csv"
        df.to_csv(csv_file, index=False)
        print(f"Saved CSV: {csv_file}")
        
        return json_file, csv_file

def main():
    print("="*60)
    print("RC WORKS SHEET EXTRACTION")
    print("="*60)
    
    extractor = RCWorksExtractor()
    items = extractor.extract_items()
    
    if items:
        # Show sample
        print("\nSample extracted items:")
        for item in items[:3]:
            print(f"\nID: {item['id']}")
            print(f"  Code: {item['code']}")
            print(f"  Description: {item['description'][:60]}...")
            print(f"  Unit: {item['unit']}")
            print(f"  Subcategory: {item['subcategory']}")
            print(f"  Rate: {item['rate']}")
            print(f"  Cell Ref: {item['cellRate_reference']}")
            print(f"  Keywords: {', '.join(item['keywords'][:3])}")
        
        extractor.save_output()
        
        # Statistics
        print("\n" + "="*60)
        print("EXTRACTION STATISTICS")
        print("="*60)
        print(f"Total items: {len(items)}")
        print(f"Items with rates: {sum(1 for i in items if i['rate'])}")
        print(f"Items with cell references: {sum(1 for i in items if i['cellRate_reference'])}")
        
        # Subcategory distribution
        subcats = {}
        for item in items:
            subcat = item['subcategory']
            subcats[subcat] = subcats.get(subcat, 0) + 1
        
        print("\nSubcategory distribution:")
        for subcat, count in sorted(subcats.items(), key=lambda x: x[1], reverse=True)[:5]:
            print(f"  {subcat}: {count}")

if __name__ == "__main__":
    main()