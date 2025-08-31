"""
Deduplication script for extracted pricelist
Removes duplicates while preserving important information
"""

import pandas as pd
import json
import hashlib
from pathlib import Path
from typing import Dict, List, Any

def generate_unique_id(description: str, category: str, unit: str = None) -> str:
    """Generate a unique ID based on content"""
    # Create a unique hash based on description, category, and unit
    content = f"{description}_{category}_{unit or ''}"
    hash_value = hashlib.md5(content.encode()).hexdigest()[:8]
    
    # Create readable ID
    cat_prefix = ''.join([c for c in category[:3].upper() if c.isalpha()])[:2] or 'XX'
    return f"{cat_prefix}_{hash_value}"

def deduplicate_csv(input_file: str, output_file: str = None):
    """Deduplicate CSV file"""
    
    if not output_file:
        output_file = input_file.replace('.csv', '_dedup.csv')
    
    print(f"Loading {input_file}...")
    df = pd.read_csv(input_file)
    original_count = len(df)
    print(f"Original items: {original_count}")
    
    # Strategy 1: Remove exact duplicates (all columns)
    df = df.drop_duplicates()
    after_exact = len(df)
    print(f"After removing exact duplicates: {after_exact} (-{original_count - after_exact})")
    
    # Strategy 2: Create composite key for uniqueness
    # Items are unique based on: description + category + unit + subcategory
    df['composite_key'] = df.apply(
        lambda row: f"{row['description']}|{row['category']}|{row.get('unit', '')}|{row.get('subcategory', '')}",
        axis=1
    )
    
    # Strategy 3: For duplicates, keep the one with most information
    # Sort by amount of non-null values (descending)
    df['info_score'] = df.notna().sum(axis=1)
    df = df.sort_values(['composite_key', 'info_score'], ascending=[True, False])
    
    # Keep first occurrence of each composite key (which has most info)
    df_dedup = df.drop_duplicates(subset=['composite_key'], keep='first')
    
    # Strategy 4: Regenerate IDs to ensure uniqueness
    print("\nRegenerating unique IDs...")
    df_dedup['id'] = df_dedup.apply(
        lambda row: generate_unique_id(
            row['description'], 
            row['category'],
            row.get('unit', '')
        ),
        axis=1
    )
    
    # Add sequential suffix if IDs still duplicate
    id_counts = {}
    new_ids = []
    for idx, row in df_dedup.iterrows():
        base_id = row['id']
        if base_id in id_counts:
            id_counts[base_id] += 1
            new_id = f"{base_id}_{id_counts[base_id]:02d}"
        else:
            id_counts[base_id] = 0
            new_id = base_id
        new_ids.append(new_id)
    
    df_dedup['id'] = new_ids
    
    # Strategy 5: Regenerate codes to ensure uniqueness within category
    print("Regenerating unique codes...")
    category_counters = {}
    new_codes = []
    
    for idx, row in df_dedup.iterrows():
        category = row['category']
        
        # Get category prefix
        cat_prefix = ''.join([c for c in category[:3].upper() if c.isalpha()])[:2]
        if not cat_prefix:
            cat_prefix = 'XX'
        
        # Increment counter for this category
        if category not in category_counters:
            category_counters[category] = 1
        else:
            category_counters[category] += 1
        
        # Generate new code
        new_code = f"{cat_prefix}{category_counters[category]:04d}"
        new_codes.append(new_code)
    
    df_dedup['code'] = new_codes
    
    # Drop temporary columns
    df_dedup = df_dedup.drop(columns=['composite_key', 'info_score'])
    
    # Final stats
    final_count = len(df_dedup)
    print(f"\nFinal items: {final_count}")
    print(f"Removed duplicates: {original_count - final_count}")
    print(f"Reduction: {((original_count - final_count) / original_count * 100):.1f}%")
    
    # Verify no duplicate IDs or codes
    duplicate_ids = df_dedup[df_dedup.duplicated(subset=['id'], keep=False)]
    duplicate_codes = df_dedup[df_dedup.duplicated(subset=['code'], keep=False)]
    print(f"\nVerification:")
    print(f"  Duplicate IDs remaining: {len(duplicate_ids)}")
    print(f"  Duplicate codes remaining: {len(duplicate_codes)}")
    
    # Category breakdown
    print(f"\nItems per category:")
    category_counts = df_dedup['category'].value_counts()
    for cat, count in category_counts.items():
        original_cat_count = len(df[df['category'] == cat])
        print(f"  {cat}: {count} (was {original_cat_count})")
    
    # Save deduplicated data
    df_dedup.to_csv(output_file, index=False)
    print(f"\nSaved deduplicated data to: {output_file}")
    
    return df_dedup

def deduplicate_json(input_file: str, output_file: str = None):
    """Deduplicate JSON file"""
    
    if not output_file:
        output_file = input_file.replace('.json', '_dedup.json')
    
    print(f"Loading {input_file}...")
    with open(input_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    original_count = len(data)
    print(f"Original items: {original_count}")
    
    # Create a dictionary to track unique items
    unique_items = {}
    
    for item in data:
        # Create composite key
        composite_key = f"{item.get('description', '')}|{item.get('category', '')}|{item.get('unit', '')}|{item.get('subcategory', '')}"
        
        # If not seen before, or this one has more information, keep it
        if composite_key not in unique_items:
            unique_items[composite_key] = item
        else:
            # Count non-null fields
            existing_info = sum(1 for v in unique_items[composite_key].values() if v is not None)
            new_info = sum(1 for v in item.values() if v is not None)
            
            if new_info > existing_info:
                unique_items[composite_key] = item
    
    # Convert back to list
    dedup_data = list(unique_items.values())
    
    # Regenerate IDs and codes
    category_counters = {}
    id_counts = {}
    
    for item in dedup_data:
        # Generate unique ID
        base_id = generate_unique_id(
            item.get('description', ''),
            item.get('category', ''),
            item.get('unit', '')
        )
        
        if base_id in id_counts:
            id_counts[base_id] += 1
            item['id'] = f"{base_id}_{id_counts[base_id]:02d}"
        else:
            id_counts[base_id] = 0
            item['id'] = base_id
        
        # Generate unique code
        category = item.get('category', 'Unknown')
        cat_prefix = ''.join([c for c in category[:3].upper() if c.isalpha()])[:2] or 'XX'
        
        if category not in category_counters:
            category_counters[category] = 1
        else:
            category_counters[category] += 1
        
        item['code'] = f"{cat_prefix}{category_counters[category]:04d}"
    
    final_count = len(dedup_data)
    print(f"\nFinal items: {final_count}")
    print(f"Removed duplicates: {original_count - final_count}")
    print(f"Reduction: {((original_count - final_count) / original_count * 100):.1f}%")
    
    # Save deduplicated data
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(dedup_data, f, indent=2, ensure_ascii=False)
    
    print(f"Saved deduplicated data to: {output_file}")
    
    return dedup_data

def main():
    """Main execution"""
    print("="*60)
    print("PRICELIST DEDUPLICATION TOOL")
    print("="*60)
    
    # Find available files
    csv_files = list(Path('.').glob('pricelist*.csv'))
    json_files = list(Path('.').glob('pricelist*.json'))
    
    # Skip already deduplicated files
    csv_files = [f for f in csv_files if '_dedup' not in f.name]
    json_files = [f for f in json_files if '_dedup' not in f.name]
    
    if not csv_files and not json_files:
        print("No pricelist files found!")
        return
    
    print("\nAvailable files:")
    all_files = []
    for i, f in enumerate(csv_files + json_files, 1):
        print(f"  {i}. {f.name}")
        all_files.append(f)
    
    # Process each file
    choice = input("\nEnter file number to deduplicate (or 'all' for all files): ").strip()
    
    if choice.lower() == 'all':
        files_to_process = all_files
    else:
        try:
            idx = int(choice) - 1
            files_to_process = [all_files[idx]]
        except (ValueError, IndexError):
            print("Invalid choice!")
            return
    
    # Process selected files
    for file_path in files_to_process:
        print(f"\n{'='*60}")
        print(f"Processing: {file_path.name}")
        print('='*60)
        
        if file_path.suffix == '.csv':
            deduplicate_csv(str(file_path))
        elif file_path.suffix == '.json':
            deduplicate_json(str(file_path))
    
    print(f"\n{'='*60}")
    print("DEDUPLICATION COMPLETE!")
    print('='*60)
    print("\nDeduplicated files have been created with '_dedup' suffix")
    print("These files are ready for import without duplicates!")

if __name__ == "__main__":
    main()