"""
Master Extraction Script
Combines all 6 sheet extractions into a single high-quality pricelist
Returns the standardized format with all required fields including cell mappings
"""

import pandas as pd
import json
from datetime import datetime
from pathlib import Path
import sys

# Import all extractors
from extract_groundworks import GroundworksExtractor
from extract_rc_works import RCWorksExtractor
from extract_drainage import DrainageExtractor
from extract_services import ServicesExtractor
from extract_external_works import ExternalWorksExtractor
from extract_underpinning import UnderpinningExtractor

class MasterPricelistExtractor:
    def __init__(self, excel_file='MJD-PRICELIST.xlsx'):
        self.excel_file = excel_file
        self.all_items = []
        self.extractors = {
            'Groundworks': GroundworksExtractor,
            'RC Works': RCWorksExtractor,
            'Drainage': DrainageExtractor,
            'Services': ServicesExtractor,
            'External Works': ExternalWorksExtractor,
            'Underpinning': UnderpinningExtractor
        }
        
    def extract_all_sheets(self):
        """Extract data from all 6 sheets"""
        print("="*80)
        print("MASTER PRICELIST EXTRACTION")
        print("="*80)
        print(f"\nExtracting from: {self.excel_file}")
        print("-"*80)
        
        all_extracted = []
        extraction_stats = {}
        
        for sheet_name, ExtractorClass in self.extractors.items():
            print(f"\n>>> Processing {sheet_name}...")
            print("-"*40)
            
            try:
                # Create extractor instance
                extractor = ExtractorClass(self.excel_file)
                
                # Extract items
                items = extractor.extract_items()
                
                if items:
                    all_extracted.extend(items)
                    extraction_stats[sheet_name] = {
                        'total': len(items),
                        'with_rates': sum(1 for i in items if i.get('rate')),
                        'with_cells': sum(1 for i in items if i.get('cellRate_reference'))
                    }
                    print(f"[OK] Extracted {len(items)} items from {sheet_name}")
                else:
                    extraction_stats[sheet_name] = {'total': 0, 'with_rates': 0, 'with_cells': 0}
                    print(f"[X] No items extracted from {sheet_name}")
                    
            except Exception as e:
                print(f"[ERROR] Error processing {sheet_name}: {str(e)}")
                extraction_stats[sheet_name] = {'error': str(e)}
        
        self.all_items = all_extracted
        return all_extracted, extraction_stats
    
    def standardize_items(self):
        """Ensure all items have the standardized format"""
        print("\nStandardizing all items...")
        
        standardized = []
        for item in self.all_items:
            # Ensure all required fields are present
            std_item = {
                'id': item.get('id', ''),
                'code': item.get('code', ''),
                'original_code': item.get('original_code', item.get('code', '')),
                'description': item.get('description', ''),
                'unit': item.get('unit', 'item'),
                'category': item.get('category', 'General'),
                'subcategory': item.get('subcategory', ''),
                'work_type': item.get('work_type', ''),
                'rate': item.get('rate'),
                'cellRate_reference': item.get('cellRate_reference'),
                'cellRate_rate': item.get('cellRate_rate'),
                'excelCellReference': item.get('excelCellReference'),
                'sourceSheetName': item.get('sourceSheetName'),
                'keywords': item.get('keywords', [])
            }
            
            # Clean None values
            for key, value in std_item.items():
                if value is None:
                    std_item[key] = '' if key != 'keywords' else []
            
            standardized.append(std_item)
        
        self.all_items = standardized
        return standardized
    
    def add_metadata(self):
        """Add metadata to all items"""
        timestamp = int(datetime.now().timestamp() * 1000)
        
        for item in self.all_items:
            item['isActive'] = True
            item['createdAt'] = timestamp
            item['updatedAt'] = timestamp
            item['createdBy'] = 'master_extraction'
            item['extractionDate'] = datetime.now().isoformat()
    
    def save_outputs(self, prefix='master_pricelist'):
        """Save the combined extraction results"""
        if not self.all_items:
            print("No items to save!")
            return None, None
        
        print(f"\nSaving {len(self.all_items)} items...")
        
        # Save JSON
        json_file = f"{prefix}.json"
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(self.all_items, f, indent=2, ensure_ascii=False)
        print(f"[OK] Saved JSON: {json_file}")
        
        # Prepare CSV data
        csv_data = []
        for item in self.all_items:
            csv_row = {
                'id': item['id'],
                'code': item['code'],
                'original_code': item['original_code'],
                'description': item['description'],
                'unit': item['unit'],
                'category': item['category'],
                'subcategory': item['subcategory'],
                'work_type': item['work_type'],
                'rate': item['rate'] if item['rate'] else '',
                'cellRate_reference': item['cellRate_reference'],
                'cellRate_rate': item['cellRate_rate'] if item['cellRate_rate'] else '',
                'excelCellReference': item['excelCellReference'],
                'sourceSheetName': item['sourceSheetName'],
                'keywords': '|'.join(item['keywords']) if item['keywords'] else ''
            }
            csv_data.append(csv_row)
        
        # Save CSV
        df = pd.DataFrame(csv_data)
        csv_file = f"{prefix}.csv"
        df.to_csv(csv_file, index=False)
        print(f"[OK] Saved CSV: {csv_file}")
        
        return json_file, csv_file
    
    def generate_statistics(self, extraction_stats):
        """Generate and display extraction statistics"""
        print("\n" + "="*80)
        print("EXTRACTION STATISTICS")
        print("="*80)
        
        total_items = len(self.all_items)
        print(f"\nTotal items extracted: {total_items}")
        
        # Per-sheet statistics
        print("\nPer-sheet breakdown:")
        print("-"*40)
        for sheet, stats in extraction_stats.items():
            if 'error' in stats:
                print(f"  {sheet:20} - ERROR: {stats['error']}")
            else:
                print(f"  {sheet:20} - Total: {stats['total']:4} | "
                      f"With rates: {stats['with_rates']:4} | "
                      f"With cells: {stats['with_cells']:4}")
        
        # Category distribution
        print("\nCategory distribution:")
        print("-"*40)
        categories = {}
        for item in self.all_items:
            cat = item['category']
            categories[cat] = categories.get(cat, 0) + 1
        
        for cat, count in sorted(categories.items(), key=lambda x: x[1], reverse=True):
            percentage = (count / total_items * 100) if total_items > 0 else 0
            print(f"  {cat:20} - {count:4} items ({percentage:.1f}%)")
        
        # Quality metrics
        print("\nQuality metrics:")
        print("-"*40)
        items_with_rates = sum(1 for i in self.all_items if i.get('rate'))
        items_with_cells = sum(1 for i in self.all_items if i.get('cellRate_reference'))
        items_with_keywords = sum(1 for i in self.all_items if i.get('keywords'))
        items_with_subcategory = sum(1 for i in self.all_items if i.get('subcategory'))
        
        print(f"  Items with rates:        {items_with_rates:4} ({items_with_rates/total_items*100:.1f}%)")
        print(f"  Items with cell refs:    {items_with_cells:4} ({items_with_cells/total_items*100:.1f}%)")
        print(f"  Items with keywords:     {items_with_keywords:4} ({items_with_keywords/total_items*100:.1f}%)")
        print(f"  Items with subcategory:  {items_with_subcategory:4} ({items_with_subcategory/total_items*100:.1f}%)")
        
        # Unit distribution
        print("\nUnit distribution (top 10):")
        print("-"*40)
        units = {}
        for item in self.all_items:
            unit = item.get('unit', 'unknown')
            units[unit] = units.get(unit, 0) + 1
        
        for unit, count in sorted(units.items(), key=lambda x: x[1], reverse=True)[:10]:
            print(f"  {unit:10} - {count:4} items")
    
    def show_samples(self, n=5):
        """Display sample extracted items"""
        print("\n" + "="*80)
        print(f"SAMPLE ITEMS (First {n})")
        print("="*80)
        
        for i, item in enumerate(self.all_items[:n], 1):
            print(f"\n[{i}] ID: {item['id']}")
            print(f"    Code: {item['code']}")
            print(f"    Original Code: {item['original_code']}")
            print(f"    Description: {item['description'][:80]}...")
            print(f"    Unit: {item['unit']}")
            print(f"    Category: {item['category']}")
            print(f"    Subcategory: {item['subcategory']}")
            print(f"    Work Type: {item['work_type']}")
            print(f"    Rate: {item['rate']}")
            print(f"    Cell Ref: {item['cellRate_reference']}")
            print(f"    Excel Cell: {item['excelCellReference']}")
            print(f"    Sheet: {item['sourceSheetName']}")
            if item['keywords']:
                print(f"    Keywords: {', '.join(item['keywords'][:5])}")

def main():
    """Main execution function"""
    print("\n" + "#"*80)
    print("#" + " "*78 + "#")
    print("#" + " "*20 + "HIGH-QUALITY PRICELIST EXTRACTION" + " "*25 + "#")
    print("#" + " "*78 + "#")
    print("#"*80)
    
    # Create master extractor
    master = MasterPricelistExtractor()
    
    try:
        # Extract from all sheets
        items, stats = master.extract_all_sheets()
        
        if not items:
            print("\n[X] No items were extracted from any sheet!")
            return
        
        # Standardize items
        master.standardize_items()
        
        # Add metadata
        master.add_metadata()
        
        # Save outputs
        json_file, csv_file = master.save_outputs()
        
        # Generate statistics
        master.generate_statistics(stats)
        
        # Show samples
        master.show_samples()
        
        # Final summary
        print("\n" + "="*80)
        print("EXTRACTION COMPLETE!")
        print("="*80)
        print(f"\n[OK] Successfully extracted {len(master.all_items)} items")
        print(f"[OK] Output files:")
        print(f"  - JSON: {json_file}")
        print(f"  - CSV:  {csv_file}")
        print("\nAll items include the following fields:")
        print("  - id, code, original_code")
        print("  - description, unit")
        print("  - category, subcategory, work_type")
        print("  - rate, cellRate_reference, cellRate_rate")
        print("  - excelCellReference, sourceSheetName")
        print("  - keywords")
        print("\n[OK] High-quality pricelist ready for use!")
        
    except Exception as e:
        print(f"\n[X] Fatal error during extraction: {str(e)}")
        import traceback
        traceback.print_exc()
        return 1
    
    return 0

if __name__ == "__main__":
    sys.exit(main())