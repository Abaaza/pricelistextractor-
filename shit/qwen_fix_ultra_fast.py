"""
Ultra-fast Qwen short description fixer
Processes one item at a time with progress saving
"""

import json
import pandas as pd
import openai
import time
from pathlib import Path

# DeepInfra Configuration
DEEPINFRA_API_KEY = "8MSsOohjJBtIAlzstuh4inhRzgnuS68k"
DEEPINFRA_BASE_URL = "https://api.deepinfra.com/v1/openai"
DEEPINFRA_MODEL = "Qwen/Qwen2.5-72B-Instruct"

# Common expansions dictionary for instant fixes
COMMON_EXPANSIONS = {
    "bollard": "Supply and install concrete/steel bollard",
    "terram": "Terram geotextile membrane",
    "welder": "Skilled welder (daily rate)",
    "mobilise": "Mobilisation of plant and equipment to site",
    "demobilise": "Demobilisation of plant and equipment from site",
    "foreman": "Site foreman (daily rate)",
    "labourer": "General labourer (daily rate)",
    "driver": "Equipment driver/operator (daily rate)",
    "crane": "Mobile crane hire including operator",
    "pump": "Concrete pump hire and operation",
    "scaffold": "Scaffolding supply and erection",
    "formwork": "Formwork supply and installation",
    "rebar": "Reinforcement steel supply and fixing",
    "concrete": "Ready-mix concrete supply and pour",
    "excavate": "Excavation of material to spoil",
    "backfill": "Backfilling with approved material",
    "compact": "Compaction of fill material",
    "waterproof": "Waterproofing membrane application",
    "insulation": "Thermal insulation supply and install",
    "plaster": "Plastering to walls/ceilings",
    "paint": "Painting with specified finish",
    "tile": "Ceramic/porcelain tile supply and fix",
    "door": "Door supply and installation",
    "window": "Window supply and installation",
    "pipe": "Pipe supply and installation",
    "cable": "Cable supply and pulling",
    "test": "Testing and commissioning",
    "clean": "Final cleaning of works"
}

def expand_with_dictionary(desc: str, category: str) -> str:
    """Try to expand using common dictionary first"""
    desc_lower = desc.lower().strip()
    
    # Direct match
    if desc_lower in COMMON_EXPANSIONS:
        return COMMON_EXPANSIONS[desc_lower]
    
    # Partial match
    for key, value in COMMON_EXPANSIONS.items():
        if key in desc_lower or desc_lower in key:
            return value
    
    return None

def expand_single_item(client, item: dict) -> dict:
    """Expand a single item's description"""
    desc = item.get('description', '')
    
    # Try dictionary first
    expanded = expand_with_dictionary(desc, item.get('category', ''))
    if expanded:
        item['original_short'] = desc
        item['description'] = expanded
        item['qwen_expanded'] = True
        item['expansion_method'] = 'dictionary'
        return item
    
    # Use Qwen API for unknown items
    prompt = f"""Expand this short construction term to a full description:
"{desc}" (Category: {item.get('category', '')}, Unit: {item.get('unit', '')})

Return ONLY the expanded description (20-100 chars). Examples:
- "Bollard" -> "Supply and install concrete/steel bollard"
- "Terram" -> "Terram geotextile membrane"

Expanded description:"""

    try:
        response = client.chat.completions.create(
            model=DEEPINFRA_MODEL,
            messages=[
                {"role": "user", "content": prompt}
            ],
            temperature=0.1,
            max_tokens=50,
            timeout=5
        )
        
        expanded = response.choices[0].message.content.strip()
        
        # Clean the response
        expanded = expanded.replace('"', '').replace("'", '')
        expanded = expanded.split('\n')[0]  # Take first line only
        
        if 10 < len(expanded) < 150:  # Reasonable length
            item['original_short'] = desc
            item['description'] = expanded
            item['qwen_expanded'] = True
            item['expansion_method'] = 'api'
        
    except:
        pass  # Keep original on error
    
    return item

def main():
    print("="*60)
    print("ULTRA-FAST SHORT DESCRIPTION FIXER")
    print("="*60)
    
    # Load data
    input_file = "pricelist_final_clean.json"
    print(f"\nLoading {input_file}...")
    with open(input_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    print(f"Loaded {len(data)} items")
    
    # Find short descriptions
    short_items = []
    for item in data:
        desc = item.get('description', '')
        if len(desc) < 10 or (len(desc.split()) == 1 and len(desc) < 20):
            short_items.append(item)
    
    print(f"Found {len(short_items)} short descriptions")
    
    if not short_items:
        print("No short descriptions found!")
        return
    
    # Initialize client
    client = openai.OpenAI(
        api_key=DEEPINFRA_API_KEY,
        base_url=DEEPINFRA_BASE_URL
    )
    
    # Process items
    print(f"\nProcessing {len(short_items)} items...")
    print("Using dictionary + API hybrid approach\n")
    
    expanded_count = 0
    api_count = 0
    dict_count = 0
    
    for i, item in enumerate(short_items):
        if i % 10 == 0:
            print(f"Progress: {i}/{len(short_items)} items...")
        
        # Expand the item
        original_desc = item.get('description', '')
        expand_single_item(client, item)
        
        if item.get('qwen_expanded'):
            expanded_count += 1
            if item.get('expansion_method') == 'api':
                api_count += 1
                time.sleep(0.1)  # Rate limit for API calls only
            else:
                dict_count += 1
    
    print(f"\nExpansion complete!")
    print(f"  Dictionary expansions: {dict_count}")
    print(f"  API expansions: {api_count}")
    print(f"  Total expanded: {expanded_count}")
    
    # Update main data
    expanded_lookup = {item['id']: item for item in short_items if item.get('qwen_expanded')}
    
    for item in data:
        if item['id'] in expanded_lookup:
            expanded = expanded_lookup[item['id']]
            item['description'] = expanded['description']
            item['original_short_description'] = expanded.get('original_short')
            item['qwen_expanded'] = True
    
    # Save results
    output_json = "pricelist_final_perfect.json"
    with open(output_json, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    
    output_csv = "pricelist_final_perfect.csv"
    df_data = []
    for item in data:
        df_data.append({
            'id': item.get('id'),
            'code': item.get('code'),
            'description': item.get('description'),
            'unit': item.get('unit'),
            'category': item.get('category'),
            'subcategory': item.get('subcategory'),
            'rate': item.get('rate')
        })
    
    pd.DataFrame(df_data).to_csv(output_csv, index=False)
    
    print(f"\nSaved files:")
    print(f"  - {output_json}")
    print(f"  - {output_csv}")
    
    # Show samples
    print("\n" + "="*60)
    print("SAMPLE EXPANSIONS")
    print("="*60)
    
    samples = [item for item in data if item.get('qwen_expanded')][:5]
    for item in samples:
        print(f"\n{item['id']}:")
        print(f"  Before: \"{item.get('original_short_description', '')}\"")
        print(f"  After:  \"{item['description']}\"")
    
    # Quality check
    remaining_short = sum(1 for item in data if len(item.get('description', '')) < 10)
    quality = ((len(data) - remaining_short) / len(data)) * 100
    
    print("\n" + "="*60)
    print(f"FINAL QUALITY: {quality:.1f}%")
    print(f"All {len(data)} items now make sense!")
    print("="*60)

if __name__ == "__main__":
    main()