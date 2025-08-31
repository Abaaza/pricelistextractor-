"""
Fixed Extraction script for Groundworks sheet
Uses actual Excel codes and includes sheet name in cell references
"""

from extractor_base import BaseExtractor
import pandas as pd
import re

class GroundworksExtractor(BaseExtractor):
    def __init__(self, excel_file='MJD-PRICELIST.xlsx'):
        super().__init__(excel_file, 'Groundworks')
        
    def identify_data_rows(self):
        """Identify rows containing actual pricelist data"""
        data_rows = []
        
        for idx, row in self.df.iterrows():
            # Skip if row is mostly empty
            if row.notna().sum() < 2:
                continue
            
            # Check if first column has a code-like value
            first_col = row[0] if 0 < len(row) else None
            if pd.notna(first_col):
                first_str = str(first_col).strip()
                # Accept any non-empty value that could be a code
                if first_str and not first_str.lower() in ['', 'nan', 'none']:
                    # Check if there's a description in nearby columns
                    has_description = False
                    for col_idx in range(1, min(4, len(row))):
                        if pd.notna(row[col_idx]):
                            desc = str(row[col_idx]).strip()
                            if len(desc) > 5 and not self.is_unit(desc):
                                has_description = True
                                break
                    
                    if has_description:
                        data_rows.append(idx)
                        continue
            
            # Also check if row has groundworks keywords
            for col_idx in range(0, min(5, len(row))):
                cell = row[col_idx]
                if pd.notna(cell):
                    cell_str = str(cell).strip().lower()
                    if any(keyword in cell_str for keyword in 
                           ['excavat', 'fill', 'disposal', 'earthwork', 'trench', 
                            'foundation', 'hardcore', 'topsoil', 'subsoil', 'rock',
                            'compact', 'level', 'grade', 'strip']):
                        data_rows.append(idx)
                        break
        
        return data_rows
    
    def extract_description(self, row, start_col=1):
        """Extract and clean description"""
        description_parts = []
        
        # Collect description from columns, skipping numbers and units
        for col_idx in range(start_col, min(start_col + 4, len(row))):
            if pd.notna(row[col_idx]):
                part = str(row[col_idx]).strip()
                # Skip if it's just a number or a unit
                if not re.match(r'^[\d,\.]+$', part) and not self.is_unit(part):
                    # Skip if it looks like a rate
                    try:
                        float(part.replace(',', '').replace('£', ''))
                        if float(part.replace(',', '').replace('£', '')) > 10:
                            continue
                    except:
                        pass
                    description_parts.append(part)
        
        description = ' '.join(description_parts)
        
        # Clean common abbreviations
        replacements = {
            ' exc ': ' excavation ',
            ' ne ': ' not exceeding ',
            ' n.e. ': ' not exceeding ',
            ' disp ': ' disposal ',
            ' fdn ': ' foundation ',
            ' u/s ': ' underside ',
            ' c/away': ' cart away',
            ' incl ': ' including ',
            ' excl ': ' excluding ',
            ' thk ': ' thick ',
            ' dp ': ' deep ',
        }
        
        for old, new in replacements.items():
            description = description.replace(old, new)
        
        # Fix patterns
        description = re.sub(r'(\d+)thk', r'\1mm thick', description)
        description = re.sub(r'(\d+)dp', r'\1m deep', description)
        
        # Clean up spaces
        description = ' '.join(description.split())
        
        return description
    
    def extract_unit(self, row):
        """Extract unit from row"""
        # Look for unit in typical unit columns (2-4)
        for col_idx in range(2, min(5, len(row))):
            if pd.notna(row[col_idx]):
                value = str(row[col_idx]).strip()
                if self.is_unit(value):
                    return value
        
        # Infer from description if not found
        desc = self.extract_description(row)
        desc_lower = desc.lower()
        
        if any(word in desc_lower for word in ['excavat', 'disposal', 'fill']):
            if 'surface' in desc_lower or 'strip' in desc_lower:
                return 'm²'
            return 'm³'
        elif 'trench' in desc_lower or 'drain' in desc_lower:
            return 'm'
        elif 'area' in desc_lower or 'slab' in desc_lower:
            return 'm²'
        
        return 'item'
    
    def determine_subcategory(self, description):
        """Determine subcategory based on description"""
        desc_lower = description.lower()
        
        if 'excavat' in desc_lower:
            if 'reduced level' in desc_lower:
                return 'Reduced Level Excavation'
            elif 'foundation' in desc_lower:
                return 'Foundation Excavation'
            elif 'trench' in desc_lower:
                return 'Trench Excavation'
            else:
                return 'General Excavation'
        elif 'fill' in desc_lower:
            if 'hardcore' in desc_lower:
                return 'Hardcore Filling'
            else:
                return 'General Filling'
        elif 'disposal' in desc_lower:
            return 'Disposal'
        elif 'compact' in desc_lower:
            return 'Compaction'
        else:
            return 'Groundworks'
    
    def generate_keywords(self, description):
        """Generate search keywords"""
        keywords = []
        desc_lower = description.lower()
        
        # Extract measurements
        measurements = re.findall(r'\d+(?:mm|m|kg|tonnes?)\b', desc_lower)
        keywords.extend(measurements[:2])
        
        # Key terms
        terms = ['excavation', 'filling', 'disposal', 'hardcore', 'topsoil', 
                 'foundation', 'trench', 'compact']
        
        for term in terms:
            if term in desc_lower:
                keywords.append(term)
        
        return keywords[:5]
    
    def extract_items(self):
        """Main extraction method"""
        if self.df is None:
            self.load_sheet()
        
        print(f"\nExtracting items from {self.sheet_name}...")
        data_rows = self.identify_data_rows()
        print(f"Found {len(data_rows)} potential data rows")
        
        items = []
        
        for row_idx in data_rows:
            row = self.df.iloc[row_idx]
            
            # Extract code (actual Excel value)
            code = self.extract_code(row)
            
            # Extract description
            description = self.extract_description(row)
            
            # Skip if no valid description
            if not description or len(description) < 5:
                continue
            
            # Extract unit
            unit = self.extract_unit(row)
            
            # Extract rate and column index
            rate, rate_col_idx = self.extract_rate(row)
            
            # Determine subcategory
            subcategory = self.determine_subcategory(description)
            
            # Generate keywords
            keywords = self.generate_keywords(description)
            
            # Create item with actual code
            item = self.create_item(
                row_idx=row_idx,
                row=row,
                code=code,
                description=description,
                unit=unit,
                subcategory=subcategory,
                rate=rate,
                rate_col_idx=rate_col_idx,
                keywords=keywords
            )
            
            items.append(item)
        
        self.extracted_items = items
        print(f"Extracted {len(items)} valid items from {self.sheet_name}")
        return items

def main():
    print("="*60)
    print("GROUNDWORKS SHEET EXTRACTION (FIXED)")
    print("="*60)
    
    extractor = GroundworksExtractor()
    items = extractor.extract_items()
    
    if items:
        # Show sample
        print("\nSample extracted items:")
        for item in items[:3]:
            print(f"\nID: {item['id']}")
            print(f"  Code: {item['code']}")
            print(f"  Description: {item['description'][:60]}...")
            print(f"  Unit: {item['unit']}")
            print(f"  Rate: {item['rate']}")
            print(f"  Cell Ref: {item['cellRate_reference']}")
        
        extractor.save_output()
        
        print(f"\nTotal items: {len(items)}")
        print(f"Items with rates: {sum(1 for i in items if i['rate'])}")
        print(f"Items with cell refs: {sum(1 for i in items if i['cellRate_reference'])}")

if __name__ == "__main__":
    main()