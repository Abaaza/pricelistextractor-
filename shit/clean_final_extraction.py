"""
Final Clean Extraction - Fixes all data quality issues
Creates perfectly clean JSON with proper units and data types
"""

import pandas as pd
import numpy as np
import json
import re
from datetime import datetime
from pathlib import Path

def is_valid_unit(value):
    """Check if value is a valid unit of measurement"""
    if pd.isna(value) or value is None:
        return False
    
    value_str = str(value).strip().lower()
    
    # Check if it's a number (not a unit)
    try:
        float(value_str)
        return False  # Pure numbers are not units
    except:
        pass
    
    # Valid construction units
    valid_units = [
        'm', 'm2', 'm²', 'm3', 'm³', 'mm', 'lm',
        'nr', 'no', 'each', 'item', 
        'kg', 'tonnes', 't', 'g',
        'hour', 'hr', 'day', 'week', 'month',
        'sum', 'ls', '%', 'set', 'pair',
        'sqm', 'cum', 'lin.m', 'sq.m', 'cu.m'
    ]
    
    return value_str in valid_units or any(unit in value_str for unit in valid_units)

def clean_unit(value):
    """Clean and standardize unit values"""
    if pd.isna(value) or value is None:
        return "item"  # Default unit
    
    value_str = str(value).strip()
    
    # Check if it's a number
    try:
        float(value_str)
        return "item"  # Default for numeric values
    except:
        pass
    
    # Standardization map
    unit_map = {
        # Area units
        'm2': 'm²', 'M2': 'm²', 'sqm': 'm²', 'SQM': 'm²', 
        'sq.m': 'm²', 'Sq.m': 'm²', 'sq m': 'm²',
        
        # Volume units
        'm3': 'm³', 'M3': 'm³', 'cum': 'm³', 'CUM': 'm³',
        'cu.m': 'm³', 'Cu.m': 'm³', 'cu m': 'm³',
        
        # Length units
        'M': 'm', 'lm': 'm', 'LM': 'm', 'lin.m': 'm',
        'Lin.m': 'm', 'linear m': 'm', 'l.m': 'm',
        
        # Count units
        'no': 'nr', 'No': 'nr', 'NO': 'nr', 'No.': 'nr',
        'no.': 'nr', 'each': 'nr', 'EACH': 'nr', 'Each': 'nr',
        
        # Weight units
        'KG': 'kg', 'Kg': 'kg', 'kgs': 'kg',
        'TONNES': 'tonnes', 'Tonnes': 'tonnes', 'T': 'tonnes',
        't': 'tonnes', 'tonne': 'tonnes',
        
        # Time units
        'HOUR': 'hr', 'Hour': 'hr', 'hour': 'hr', 'hours': 'hr',
        'HR': 'hr', 'Hr': 'hr', 'hrs': 'hr',
        'DAY': 'day', 'Day': 'day', 'days': 'day',
        'WEEK': 'week', 'Week': 'week', 'wk': 'week', 'WK': 'week',
        'MONTH': 'month', 'Month': 'month', 'mth': 'month', 'MTH': 'month',
        
        # Other units
        'ITEM': 'item', 'Item': 'item',
        'SUM': 'sum', 'Sum': 'sum', 'LS': 'sum', 'ls': 'sum',
        'SET': 'set', 'Set': 'set',
        'PAIR': 'pair', 'Pair': 'pair',
        '%': '%', 'percent': '%', 'PERCENT': '%'
    }
    
    # Apply mapping
    if value_str in unit_map:
        return unit_map[value_str]
    
    # Check lowercase version
    if value_str.lower() in unit_map:
        return unit_map[value_str.lower()]
    
    # If it's already a valid unit, return as lowercase
    if is_valid_unit(value_str):
        return value_str.lower()
    
    # Default
    return "item"

def clean_description(desc):
    """Clean and enhance description"""
    if pd.isna(desc) or not desc:
        return ""
    
    desc = str(desc).strip()
    
    # Expand common abbreviations
    replacements = {
        ' ne ': ' not exceeding ',
        ' n.e. ': ' not exceeding ',
        ' thk': ' thick',
        ' THK': ' thick',
        ' exc ': ' excavation ',
        ' exc.': ' excavation',
        ' inc ': ' including ',
        ' inc.': ' including',
        ' incl ': ' including ',
        ' excl ': ' excluding ',
        ' reinf ': ' reinforcement ',
        ' conc ': ' concrete ',
        ' fdn ': ' foundation ',
        ' fwk ': ' formwork ',
        ' u/s ': ' underside ',
        ' o/a ': ' overall ',
        ' c/c ': ' center to center ',
        ' dia ': ' diameter ',
        ' adj ': ' adjacent ',
        ' horiz ': ' horizontal ',
        ' vert ': ' vertical ',
        ' approx ': ' approximately ',
        ' temp ': ' temporary ',
        ' perm ': ' permanent ',
        ' struct ': ' structural ',
        ' bwk ': ' brickwork ',
        ' blk ': ' blockwork ',
        ' r.c.': ' reinforced concrete',
        ' rc ': ' reinforced concrete ',
        ' c/w ': ' complete with ',
        ' w/ ': ' with ',
        ' w/o ': ' without ',
    }
    
    for old, new in replacements.items():
        desc = desc.replace(old, new)
        desc = desc.replace(old.upper(), new)
    
    # Fix patterns like "150thk" -> "150mm thick"
    desc = re.sub(r'(\d+)thk', r'\1mm thick', desc)
    desc = re.sub(r'(\d+)THK', r'\1mm thick', desc)
    
    # Fix "ne150" -> "not exceeding 150"
    desc = re.sub(r'\bne(\d+)', r'not exceeding \1', desc)
    desc = re.sub(r'\bNE(\d+)', r'not exceeding \1', desc)
    
    # Clean up extra spaces
    desc = ' '.join(desc.split())
    
    return desc

def infer_unit_from_description(desc):
    """Infer unit from description if missing"""
    if not desc:
        return "item"
    
    desc_lower = str(desc).lower()
    
    # Check for area indicators
    if any(word in desc_lower for word in ['area', 'square', 'floor', 'ceiling', 'wall face', 'surface']):
        return "m²"
    
    # Check for volume indicators
    if any(word in desc_lower for word in ['volume', 'cubic', 'excavat', 'concrete', 'fill', 'pour']):
        return "m³"
    
    # Check for linear indicators
    if any(word in desc_lower for word in ['length', 'linear', 'perimeter', 'edge', 'joint', 'pipe']):
        return "m"
    
    # Check for weight indicators
    if any(word in desc_lower for word in ['weight', 'steel', 'reinforcement', 'rebar']):
        return "kg"
    
    # Check for count indicators
    if any(word in desc_lower for word in ['number', 'quantity', 'each', 'unit', 'fixing', 'bolt']):
        return "nr"
    
    # Check for time indicators
    if any(word in desc_lower for word in ['hour', 'day', 'week', 'month', 'duration']):
        if 'hour' in desc_lower:
            return "hr"
        elif 'day' in desc_lower:
            return "day"
        elif 'week' in desc_lower:
            return "week"
        elif 'month' in desc_lower:
            return "month"
    
    # Default
    return "item"

def generate_keywords(desc, category):
    """Generate relevant keywords for searching"""
    if not desc:
        return []
    
    keywords = []
    desc_lower = str(desc).lower()
    
    # Extract key terms based on category
    if 'groundwork' in str(category).lower():
        terms = ['excavation', 'filling', 'earthwork', 'disposal', 'foundation', 'trench', 'pit']
    elif 'rc work' in str(category).lower() or 'concrete' in str(category).lower():
        terms = ['concrete', 'formwork', 'reinforcement', 'rebar', 'mesh', 'shuttering', 'pour']
    elif 'drain' in str(category).lower():
        terms = ['drainage', 'pipe', 'manhole', 'gully', 'sewer', 'channel', 'outlet']
    elif 'external' in str(category).lower():
        terms = ['paving', 'kerb', 'fence', 'landscape', 'road', 'path', 'car park']
    else:
        terms = ['construction', 'building', 'structure', 'install', 'supply', 'fix']
    
    # Add matching terms
    for term in terms:
        if term in desc_lower:
            keywords.append(term)
    
    # Extract measurements
    measurements = re.findall(r'\d+mm|\d+m\b|\d+kg|\d+t\b', desc_lower)
    keywords.extend(measurements[:2])
    
    # Extract material types
    materials = ['concrete', 'steel', 'timber', 'block', 'brick', 'stone', 
                'aggregate', 'sand', 'cement', 'mortar', 'plaster']
    for mat in materials:
        if mat in desc_lower:
            keywords.append(mat)
    
    # Remove duplicates and limit
    keywords = list(dict.fromkeys(keywords))[:6]
    
    return keywords

def main():
    print("="*60)
    print("FINAL CLEAN EXTRACTION")
    print("="*60)
    
    # Load the original extracted data
    input_file = "pricelist_v2.csv"
    if not Path(input_file).exists():
        print(f"Error: {input_file} not found!")
        return
    
    print(f"Loading {input_file}...")
    df = pd.read_csv(input_file)
    print(f"Loaded {len(df)} items")
    
    # Clean and fix all data
    print("\nCleaning data...")
    
    # 1. Clean descriptions
    df['clean_description'] = df['description'].apply(clean_description)
    
    # 2. Fix units
    print("  Fixing units...")
    df['clean_unit'] = df['unit'].apply(clean_unit)
    
    # For items with invalid units, infer from description
    invalid_units = df['clean_unit'] == 'item'
    print(f"  Found {invalid_units.sum()} items with default units")
    
    # Infer better units where possible
    df.loc[invalid_units, 'clean_unit'] = df.loc[invalid_units, 'clean_description'].apply(
        infer_unit_from_description
    )
    
    # 3. Generate keywords
    print("  Generating keywords...")
    df['keywords_list'] = df.apply(
        lambda row: generate_keywords(row['clean_description'], row['category']), 
        axis=1
    )
    
    # 4. Determine work type
    print("  Assigning work types...")
    def get_work_type(row):
        desc = str(row['clean_description']).lower()
        cat = str(row['category']).lower()
        
        if 'excavat' in desc:
            return 'Excavation'
        elif 'concrete' in desc or 'conc' in desc:
            return 'Concrete'
        elif 'formwork' in desc or 'shutter' in desc:
            return 'Formwork'
        elif 'reinforc' in desc or 'rebar' in desc:
            return 'Reinforcement'
        elif 'block' in desc or 'brick' in desc:
            return 'Masonry'
        elif 'drain' in desc or 'pipe' in desc:
            return 'Drainage'
        elif 'paving' in desc or 'road' in desc:
            return 'Paving'
        elif 'external' in cat:
            return 'External Works'
        elif 'prelim' in cat:
            return 'Preliminaries'
        elif 'plant' in cat or 'crane' in cat or 'hoist' in cat:
            return 'Plant & Equipment'
        else:
            return 'General Construction'
    
    df['work_type'] = df.apply(get_work_type, axis=1)
    
    # 5. Create clean JSON
    print("\nCreating clean JSON...")
    current_timestamp = int(datetime.now().timestamp() * 1000)
    
    json_data = []
    for idx, row in df.iterrows():
        # Prepare cell rate
        cell_rate = None
        if pd.notna(row.get('cellRate_reference')) and pd.notna(row.get('cellRate_rate')):
            try:
                cell_rate = {
                    'reference': str(row['cellRate_reference']),
                    'rate': float(row['cellRate_rate'])
                }
            except:
                pass
        
        # Create clean item
        item = {
            # Core fields
            'id': str(row['id']),
            'code': str(row['code']) if pd.notna(row['code']) else None,
            'original_code': str(row['original_code']) if pd.notna(row.get('original_code')) else None,
            'description': row['clean_description'],
            
            # Enhanced fields
            'keywords': row['keywords_list'] if row['keywords_list'] else [],
            'category': str(row['category']) if pd.notna(row['category']) else 'General',
            'subcategory': str(row['subcategory']) if pd.notna(row.get('subcategory')) else None,
            'work_type': row['work_type'],
            'unit': row['clean_unit'],
            
            # Rate information
            'rate': float(row['rate']) if pd.notna(row.get('rate')) else None,
            'cellRate': cell_rate,
            
            # Excel mapping
            'excelCellReference': str(row['excelCellReference']) if pd.notna(row.get('excelCellReference')) else None,
            'sourceSheetName': str(row['sourceSheetName']) if pd.notna(row.get('sourceSheetName')) else None,
            'sourceRowNumber': int(row['sourceRowNumber']) if pd.notna(row.get('sourceRowNumber')) else None,
            'sourceColumnLetter': str(row['sourceColumnLetter']) if pd.notna(row.get('sourceColumnLetter')) else None,
            
            # Metadata
            'isActive': True,
            'createdAt': current_timestamp,
            'updatedAt': current_timestamp,
            'createdBy': 'system'
        }
        
        # Remove None values for cleaner JSON
        item = {k: v for k, v in item.items() if v is not None}
        json_data.append(item)
    
    # Save files
    print("\nSaving files...")
    
    # Save clean CSV
    csv_output = "pricelist_final_clean.csv"
    df_export = df[[
        'id', 'code', 'original_code', 'clean_description', 'clean_unit',
        'category', 'subcategory', 'work_type', 'rate',
        'cellRate_reference', 'cellRate_rate',
        'excelCellReference', 'sourceSheetName'
    ]].copy()
    df_export.columns = [
        'id', 'code', 'original_code', 'description', 'unit',
        'category', 'subcategory', 'work_type', 'rate',
        'cellRate_reference', 'cellRate_rate',
        'excelCellReference', 'sourceSheetName'
    ]
    df_export['keywords'] = df['keywords_list'].apply(lambda x: '|'.join(x) if x else '')
    df_export.to_csv(csv_output, index=False)
    print(f"  Saved CSV: {csv_output}")
    
    # Save clean JSON
    json_output = "pricelist_final_clean.json"
    with open(json_output, 'w', encoding='utf-8') as f:
        json.dump(json_data, f, indent=2, ensure_ascii=False)
    print(f"  Saved JSON: {json_output}")
    
    # Statistics
    print("\n" + "="*60)
    print("EXTRACTION STATISTICS")
    print("="*60)
    print(f"Total items: {len(json_data)}")
    print(f"Items with original codes: {sum(1 for item in json_data if 'original_code' in item)}")
    print(f"Items with keywords: {sum(1 for item in json_data if item.get('keywords'))}")
    print(f"Items with proper units: {sum(1 for item in json_data if item.get('unit') != 'item')}")
    
    # Unit distribution
    unit_dist = {}
    for item in json_data:
        unit = item.get('unit', 'unknown')
        unit_dist[unit] = unit_dist.get(unit, 0) + 1
    
    print("\nUnit distribution:")
    for unit, count in sorted(unit_dist.items(), key=lambda x: x[1], reverse=True)[:10]:
        print(f"  {unit}: {count}")
    
    # Sample output
    print("\n" + "="*60)
    print("SAMPLE CLEAN ITEMS")
    print("="*60)
    
    for item in json_data[:3]:
        print(f"\nID: {item['id']}")
        print(f"  Code: {item.get('code', 'N/A')}")
        print(f"  Description: {item['description'][:60]}...")
        print(f"  Unit: {item['unit']}")
        print(f"  Category: {item['category']}")
        print(f"  Work Type: {item.get('work_type', 'N/A')}")
        if item.get('keywords'):
            print(f"  Keywords: {', '.join(item['keywords'][:3])}")
    
    print("\n✅ Clean extraction complete!")
    print(f"Files created:")
    print(f"  - {csv_output}")
    print(f"  - {json_output}")

if __name__ == "__main__":
    main()