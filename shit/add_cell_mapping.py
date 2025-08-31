"""
Add cell mapping and all required fields to the final perfect pricelist
Ensures all Excel cell references are preserved for rate updates
"""

import json
import pandas as pd
from pathlib import Path

def merge_cell_mappings():
    """Merge cell mappings from clean data into perfect data"""
    
    print("="*60)
    print("ADDING CELL MAPPINGS TO PERFECT PRICELIST")
    print("="*60)
    
    # Load the perfect descriptions
    perfect_file = "pricelist_final_perfect.json"
    print(f"\nLoading {perfect_file}...")
    with open(perfect_file, 'r', encoding='utf-8') as f:
        perfect_data = json.load(f)
    print(f"Loaded {len(perfect_data)} items with perfect descriptions")
    
    # Load the clean data with cell mappings
    clean_file = "pricelist_final_clean.json"
    print(f"\nLoading {clean_file} for cell mappings...")
    with open(clean_file, 'r', encoding='utf-8') as f:
        clean_data = json.load(f)
    print(f"Loaded {len(clean_data)} items with cell mappings")
    
    # Create lookup for clean data by ID
    clean_lookup = {item['id']: item for item in clean_data}
    
    # Merge cell mappings into perfect data
    print("\nMerging cell mappings...")
    items_with_cells = 0
    items_with_keywords = 0
    
    for item in perfect_data:
        item_id = item['id']
        
        if item_id in clean_lookup:
            clean_item = clean_lookup[item_id]
            
            # Add all cell mapping fields
            if 'cellRate' in clean_item:
                item['cellRate'] = clean_item['cellRate']
                item['cellRate_reference'] = clean_item['cellRate'].get('reference')
                item['cellRate_rate'] = clean_item['cellRate'].get('rate')
                items_with_cells += 1
            
            # Add Excel reference fields
            if 'excelCellReference' in clean_item:
                item['excelCellReference'] = clean_item['excelCellReference']
            
            if 'sourceSheetName' in clean_item:
                item['sourceSheetName'] = clean_item['sourceSheetName']
            
            if 'sourceRowNumber' in clean_item:
                item['sourceRowNumber'] = clean_item['sourceRowNumber']
            
            if 'sourceColumnLetter' in clean_item:
                item['sourceColumnLetter'] = clean_item['sourceColumnLetter']
            
            # Add keywords if available
            if 'keywords' in clean_item:
                item['keywords'] = clean_item['keywords']
                if clean_item['keywords']:
                    items_with_keywords += 1
            
            # Add work_type if available
            if 'work_type' in clean_item:
                item['work_type'] = clean_item['work_type']
            
            # Ensure original_code is preserved
            if 'original_code' in clean_item:
                item['original_code'] = clean_item['original_code']
    
    print(f"Merged {items_with_cells} cell mappings")
    print(f"Found {items_with_keywords} items with keywords")
    
    # Save updated JSON with all fields
    output_json = "pricelist_complete_final.json"
    with open(output_json, 'w', encoding='utf-8') as f:
        json.dump(perfect_data, f, indent=2, ensure_ascii=False)
    print(f"\nSaved complete JSON: {output_json}")
    
    # Create comprehensive CSV with all columns
    print("\nCreating comprehensive CSV...")
    df_data = []
    
    for item in perfect_data:
        row = {
            'id': item.get('id'),
            'code': item.get('code'),
            'original_code': item.get('original_code'),
            'description': item.get('description'),
            'unit': item.get('unit'),
            'category': item.get('category'),
            'subcategory': item.get('subcategory'),
            'work_type': item.get('work_type', ''),
            'rate': item.get('rate'),
            'cellRate_reference': item.get('cellRate', {}).get('reference', ''),
            'cellRate_rate': item.get('cellRate', {}).get('rate', ''),
            'excelCellReference': item.get('excelCellReference', ''),
            'sourceSheetName': item.get('sourceSheetName', ''),
            'keywords': '|'.join(item.get('keywords', [])) if item.get('keywords') else ''
        }
        df_data.append(row)
    
    df = pd.DataFrame(df_data)
    
    # Save comprehensive CSV
    output_csv = "pricelist_complete_final.csv"
    df.to_csv(output_csv, index=False)
    print(f"Saved comprehensive CSV: {output_csv}")
    
    # Show sample with all fields
    print("\n" + "="*60)
    print("SAMPLE ITEMS WITH ALL FIELDS")
    print("="*60)
    
    sample_items = df[df['cellRate_reference'] != ''].head(3)
    for idx, row in sample_items.iterrows():
        print(f"\nID: {row['id']}")
        print(f"  Code: {row['code']}")
        print(f"  Original Code: {row['original_code']}")
        print(f"  Description: {row['description'][:50]}...")
        print(f"  Unit: {row['unit']}")
        print(f"  Category: {row['category']}")
        print(f"  Subcategory: {row['subcategory']}")
        print(f"  Work Type: {row['work_type']}")
        print(f"  Rate: {row['rate']}")
        print(f"  Cell Reference: {row['cellRate_reference']}")
        print(f"  Excel Cell: {row['excelCellReference']}")
        print(f"  Sheet Name: {row['sourceSheetName']}")
        print(f"  Keywords: {row['keywords']}")
    
    # Final statistics
    print("\n" + "="*60)
    print("FINAL STATISTICS")
    print("="*60)
    
    total_items = len(df)
    items_with_cells = len(df[df['cellRate_reference'] != ''])
    items_with_original_codes = len(df[df['original_code'] != ''])
    items_with_keywords = len(df[df['keywords'] != ''])
    items_with_work_type = len(df[df['work_type'] != ''])
    
    print(f"Total items: {total_items}")
    print(f"Items with cell references: {items_with_cells} ({items_with_cells/total_items*100:.1f}%)")
    print(f"Items with original codes: {items_with_original_codes} ({items_with_original_codes/total_items*100:.1f}%)")
    print(f"Items with keywords: {items_with_keywords} ({items_with_keywords/total_items*100:.1f}%)")
    print(f"Items with work types: {items_with_work_type} ({items_with_work_type/total_items*100:.1f}%)")
    
    # Check for quality
    if items_with_cells < total_items * 0.8:
        print("\nWARNING: Less than 80% of items have cell references!")
        print("This may affect rate update functionality.")
    else:
        print(f"\n✓ EXCELLENT! {items_with_cells/total_items*100:.1f}% of items have cell references for rate updates!")
    
    print("\n" + "="*60)
    print("COMPLETE PRICELIST WITH ALL FIELDS READY!")
    print("="*60)
    print("\nFinal files:")
    print(f"  - {output_json} (Complete JSON with all fields)")
    print(f"  - {output_csv} (Complete CSV with all columns)")
    print("\nThese files include:")
    print("  • Perfect descriptions (expanded from short)")
    print("  • Cell references for rate updates")
    print("  • Original Excel codes preserved")
    print("  • Keywords for searching")
    print("  • All required database fields")
    
    return output_json, output_csv

if __name__ == "__main__":
    merge_cell_mappings()