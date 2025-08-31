"""
Extraction script for External Works sheet
Handles the specific structure and format of the External Works pricelist
"""

from extractor_base import BaseExtractor
import pandas as pd
import re
from openpyxl import load_workbook

class ExternalWorksExtractor(BaseExtractor):
    def __init__(self, excel_file='MJD-PRICELIST.xlsx'):
        super().__init__(excel_file, 'External Works')
        self.workbook = None
        self.worksheet = None
        self.current_subcategory = 'External Works'  # Default subcategory
        
    def load_workbook_for_formatting(self):
        """Load workbook with openpyxl to access formatting"""
        try:
            self.workbook = load_workbook(self.excel_file, data_only=True)
            self.worksheet = self.workbook[self.sheet_name]
            print(f"Loaded workbook for formatting detection")
        except Exception as e:
            print(f"Warning: Could not load workbook for formatting: {e}")
            self.worksheet = None
    
    def is_row_bold(self, row_idx):
        """Check if all non-empty cells in a row are bold"""
        if not self.worksheet:
            return False
        
        try:
            # Excel rows are 1-indexed
            excel_row = row_idx + 1
            row_is_bold = False
            has_content = False
            
            # Check first 5 columns for bold text
            for col in range(1, 6):
                cell = self.worksheet.cell(row=excel_row, column=col)
                if cell.value:
                    has_content = True
                    if cell.font and cell.font.bold:
                        row_is_bold = True
                    else:
                        # If any cell with content is not bold, row is not fully bold
                        return False
            
            return has_content and row_is_bold
        except Exception as e:
            return False
    
    def identify_data_rows(self):
        """Identify rows containing actual pricelist data"""
        data_rows = []
        
        for idx, row in self.df.iterrows():
            # Skip if row is mostly empty
            if row.notna().sum() < 3:
                continue
                
            # Look for patterns that indicate a data row
            # External Works typically has: Code | Description | Unit | Rate columns
            
            # Check if first column might be a code
            code_col = row[0] if 0 < len(row) else None
            if pd.notna(code_col):
                code_str = str(code_col).strip()
                # External Works codes often start with numbers or specific prefixes
                if (re.match(r'^\d+', code_str) or 
                    re.match(r'^[A-Z]\d+', code_str) or
                    re.match(r'^EW', code_str, re.I)):
                    data_rows.append(idx)
                    continue
            
            # Check if row has description-like content
            for col_idx in range(1, min(5, len(row))):
                cell = row[col_idx]
                if pd.notna(cell):
                    cell_str = str(cell).strip().lower()
                    # External Works keywords
                    if any(keyword in cell_str for keyword in 
                           ['paving', 'kerb', 'edging', 'fence', 'gate', 
                            'tarmac', 'asphalt', 'concrete', 'block', 'slab',
                            'drainage', 'channel', 'gulley', 'bollard', 'signage']):
                        data_rows.append(idx)
                        break
        
        return data_rows
    
    def extract_code(self, row, col_idx=0):
        """Extract code from row"""
        if col_idx < len(row) and pd.notna(row[col_idx]):
            code = str(row[col_idx]).strip()
            # Clean up code
            code = re.sub(r'\s+', '', code)  # Remove spaces
            if code and not code.lower() in ['nan', 'none', '-', '']:
                return code
        return None
    
    def extract_description(self, row, start_col=1):
        """Extract and clean description from columns B and C primarily"""
        description_parts = []
        
        # Column B (index 1) is usually the main description
        if len(row) > 1 and pd.notna(row[1]):
            desc = str(row[1]).strip()
            if desc and not self.is_unit(desc):
                description_parts.append(desc)
        
        # Column C (index 2) might have continuation or additional info
        if len(row) > 2 and pd.notna(row[2]):
            part = str(row[2]).strip()
            # Only add if it's not a unit and not a number
            if part and not self.is_unit(part) and not re.match(r'^[\d,\.]+$', part):
                description_parts.append(part)
        
        description = ' '.join(description_parts)
        
        # Clean common abbreviations
        replacements = {
            ' ne ': ' not exceeding ',
            ' n.e. ': ' not exceeding ',
            ' incl ': ' including ',
            ' excl ': ' excluding ',
            ' thk ': ' thick ',
            ' dp ': ' deep ',
            ' w ': ' wide ',
            ' h ': ' high ',
        }
        
        for old, new in replacements.items():
            description = description.replace(old, new)
        
        # Fix patterns
        description = re.sub(r'(\d+)thk', r'\1mm thick', description)
        description = re.sub(r'(\d+)dp', r'\1m deep', description)
        
        # Clean up spaces
        description = ' '.join(description.split())
        
        return description
    
    def is_unit(self, value):
        """Check if value is a unit"""
        if pd.isna(value):
            return False
        
        value_str = str(value).strip().lower()
        
        units = ['m', 'm2', 'm²', 'm3', 'm³', 'nr', 'no', 'item', 'sum',
                 'kg', 'tonnes', 't', 'lm', 'sqm', 'cum', 'each']
        
        return value_str in units
    
    def extract_unit(self, row):
        """Extract unit from row - primarily column E (index 4)"""
        # Check column E first (index 4) - this is the primary unit column
        if len(row) > 4 and pd.notna(row[4]):
            value = str(row[4]).strip()
            if self.is_unit(value):
                return self.standardize_unit(value)
        
        # Then check columns C and D as fallback
        for col_idx in [2, 3]:
            if col_idx < len(row) and pd.notna(row[col_idx]):
                value = str(row[col_idx]).strip()
                if self.is_unit(value):
                    return self.standardize_unit(value)
        
        # Infer from description if not found
        desc = self.extract_description(row)
        desc_lower = desc.lower()
        
        if any(word in desc_lower for word in ['paving', 'surfacing', 'tarmac']):
            return 'm2'
        elif any(word in desc_lower for word in ['kerb', 'edging', 'channel']):
            return 'm'
        elif any(word in desc_lower for word in ['bollard', 'sign', 'gate', 'post']):
            return 'nr'
        elif 'fence' in desc_lower or 'fencing' in desc_lower:
            return 'm'
        
        return 'item'
    
    def standardize_unit(self, unit):
        """Standardize unit format"""
        unit_map = {
            'm2': 'm2', 'sqm': 'm2', 'm²': 'm2',
            'm3': 'm3', 'cum': 'm3', 'm³': 'm3',
            'no': 'nr', 'no.': 'nr', 'each': 'nr',
            't': 'tonnes', 'tonne': 'tonnes',
            'lm': 'm', 'lin.m': 'm', 'l.m': 'm',
        }
        
        unit_lower = unit.lower()
        return unit_map.get(unit_lower, unit_lower)
    
    def infer_unit_from_description(self, row):
        """Infer unit from description content"""
        desc = self.extract_description(row)
        desc_lower = desc.lower()
        
        # Groundworks specific patterns
        if any(word in desc_lower for word in ['excavat', 'disposal', 'fill', 'earthwork']):
            if 'surface' in desc_lower or 'strip' in desc_lower:
                return 'm²'
            return 'm³'
        elif any(word in desc_lower for word in ['trench', 'drain', 'edge', 'kerb']):
            return 'm'
        elif any(word in desc_lower for word in ['area', 'slab', 'bed']):
            return 'm²'
        elif any(word in desc_lower for word in ['volume', 'bulk', 'mass']):
            return 'm³'
        
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
        """Determine subcategory based on description"""
        desc_lower = description.lower()
        
        # Groundworks subcategories
        if 'excavat' in desc_lower:
            if 'reduced level' in desc_lower:
                return 'Reduced Level Excavation'
            elif 'foundation' in desc_lower or 'fdn' in desc_lower:
                return 'Foundation Excavation'
            elif 'trench' in desc_lower:
                return 'Trench Excavation'
            elif 'basement' in desc_lower:
                return 'Basement Excavation'
            else:
                return 'General Excavation'
        elif 'fill' in desc_lower:
            if 'hardcore' in desc_lower:
                return 'Hardcore Filling'
            elif 'selected' in desc_lower:
                return 'Selected Fill'
            elif 'imported' in desc_lower:
                return 'Imported Fill'
            else:
                return 'General Filling'
        elif 'disposal' in desc_lower or 'cart away' in desc_lower:
            return 'Disposal'
        elif 'compact' in desc_lower:
            return 'Compaction'
        elif 'level' in desc_lower or 'grade' in desc_lower:
            return 'Leveling and Grading'
        elif 'surface' in desc_lower or 'strip' in desc_lower:
            return 'Surface Preparation'
        elif 'rock' in desc_lower:
            return 'Rock Excavation'
        elif 'support' in desc_lower or 'shore' in desc_lower:
            return 'Earthwork Support'
        else:
            return 'General Groundworks'
    
    def determine_work_type(self, description, subcategory):
        """Determine work type"""
        desc_lower = description.lower()
        
        if 'excavat' in desc_lower:
            return 'Excavation'
        elif 'fill' in desc_lower:
            return 'Filling'
        elif 'disposal' in desc_lower:
            return 'Disposal'
        elif 'compact' in desc_lower:
            return 'Compaction'
        elif 'level' in desc_lower or 'grade' in desc_lower:
            return 'Site Preparation'
        elif 'support' in desc_lower:
            return 'Temporary Works'
        else:
            return 'Groundworks'
    
    def generate_keywords(self, description, subcategory):
        """Generate search keywords"""
        keywords = []
        desc_lower = description.lower()
        
        # Extract measurements
        measurements = re.findall(r'\d+(?:mm|m|kg|tonnes?)\b', desc_lower)
        keywords.extend(measurements[:2])
        
        # Extract depths
        depths = re.findall(r'(?:ne|not exceeding)\s*(\d+)m?\b', desc_lower)
        for depth in depths[:1]:
            keywords.append(f"depth_{depth}m")
        
        # Key groundworks terms
        terms = ['excavation', 'filling', 'disposal', 'hardcore', 'topsoil', 
                 'foundation', 'trench', 'basement', 'rock', 'compact']
        
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
            excel_ref = self.get_cell_reference(row_idx, 0)  # Reference to code cell
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
                'id': f"GW{current_id:04d}",
                'code': code if code else f"GW{current_id:04d}",
                'original_code': code,
                'description': description,
                'unit': unit,
                'category': 'Groundworks',
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
    
    def save_output(self, output_prefix='groundworks'):
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
    print("EXTERNAL WORKS SHEET EXTRACTION")
    print("="*60)
    
    extractor = ExternalWorksExtractor()
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