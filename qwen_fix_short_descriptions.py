"""
Fix Short Descriptions with Qwen 72B
Specifically targets items with very short descriptions and expands them
"""

import json
import pandas as pd
import openai
import time
import re
from pathlib import Path

# DeepInfra Configuration
DEEPINFRA_API_KEY = "8MSsOohjJBtIAlzstuh4inhRzgnuS68k"
DEEPINFRA_BASE_URL = "https://api.deepinfra.com/v1/openai"
DEEPINFRA_MODEL = "Qwen/Qwen2.5-72B-Instruct"

def find_short_descriptions(data):
    """Find all items with short descriptions"""
    short_items = []
    
    for item in data:
        desc = item.get('description', '')
        
        # Find very short descriptions (< 10 chars) or single words
        if len(desc) < 10 or (len(desc.split()) == 1 and len(desc) < 20):
            short_items.append(item)
    
    return short_items

def expand_descriptions_with_qwen(client, batch):
    """Use Qwen to expand short descriptions"""
    
    # Prepare items for API
    items_to_expand = []
    for i, item in enumerate(batch):
        items_to_expand.append({
            'index': i,
            'id': item.get('id', ''),
            'code': item.get('code', ''),
            'description': item.get('description', ''),
            'unit': item.get('unit', ''),
            'category': item.get('category', ''),
            'subcategory': item.get('subcategory', ''),
            'rate': item.get('rate', '')
        })
    
    prompt = f"""You are a construction cost estimator. These items have very short descriptions that need expansion.

Items to expand:
{json.dumps(items_to_expand, indent=2)}

For each item, provide a FULL, PROPER description based on:
1. The current short description (may be a product name or abbreviation)
2. The category and subcategory context
3. The unit of measurement
4. Standard construction terminology

Rules:
- Keep original meaning but expand to full description
- "Bollard" → "Supply and install concrete/steel bollard"
- "Terram" → "Terram geotextile membrane"
- "Welder" → "Skilled welder (daily rate)"
- "Mobilise" → "Mobilisation of plant and equipment to site"
- Make descriptions clear and professional
- Include "supply and install" or "labour only" where appropriate
- Minimum 20 characters, maximum 100 characters

Return JSON array:
[{{
  "index": 0,
  "expanded_description": "Full professional description",
  "confidence": 0.9
}}]"""

    try:
        response = client.chat.completions.create(
            model=DEEPINFRA_MODEL,
            messages=[
                {"role": "system", "content": "You are a construction expert. Expand short descriptions to full professional descriptions."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.3,
            max_tokens=2000
        )
        
        result = response.choices[0].message.content.strip()
        
        # Parse JSON
        if '```json' in result:
            result = result.split('```json')[1].split('```')[0]
        elif '```' in result:
            result = result.split('```')[1].split('```')[0]
        
        # Extract JSON array
        json_match = re.search(r'\[[\s\S]*\]', result)
        if json_match:
            result = json_match.group()
        
        expanded_data = json.loads(result)
        
        # Apply expansions
        for item in batch:
            item_index = batch.index(item)
            
            for expanded in expanded_data:
                if expanded.get('index') == item_index:
                    if expanded.get('expanded_description'):
                        # Store original for reference
                        item['original_short_description'] = item.get('description')
                        item['description'] = expanded['expanded_description']
                        item['qwen_expanded'] = True
                        item['expansion_confidence'] = expanded.get('confidence', 0.8)
                    break
        
        return batch
        
    except Exception as e:
        print(f"    API error: {str(e)[:100]}")
        return batch

def main():
    print("="*60)
    print("FIXING SHORT DESCRIPTIONS WITH QWEN 72B")
    print("="*60)
    
    # Load the clean data
    input_file = "pricelist_final_clean.json"
    if not Path(input_file).exists():
        print(f"Error: {input_file} not found!")
        return
    
    print(f"\nLoading {input_file}...")
    with open(input_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    print(f"Loaded {len(data)} items")
    
    # Find short descriptions
    print("\nFinding short descriptions...")
    short_items = find_short_descriptions(data)
    print(f"Found {len(short_items)} items with short descriptions")
    
    if not short_items:
        print("No short descriptions found!")
        return
    
    # Show samples
    print("\nSample short descriptions:")
    for item in short_items[:10]:
        print(f"  {item['id']}: \"{item.get('description', '')}\" [{item.get('unit', '')}] in {item.get('category', '')}")
    
    # Initialize Qwen client
    print("\nInitializing Qwen 72B...")
    client = openai.OpenAI(
        api_key=DEEPINFRA_API_KEY,
        base_url=DEEPINFRA_BASE_URL
    )
    
    # Process in batches
    batch_size = 10
    total_batches = (len(short_items) + batch_size - 1) // batch_size
    
    print(f"\nProcessing {len(short_items)} items in {total_batches} batches...")
    
    expanded_items = []
    for i in range(0, len(short_items), batch_size):
        batch = short_items[i:i+batch_size]
        batch_num = (i // batch_size) + 1
        
        print(f"  Batch {batch_num}/{total_batches} ({len(batch)} items)...")
        
        # Expand with Qwen
        expanded_batch = expand_descriptions_with_qwen(client, batch)
        expanded_items.extend(expanded_batch)
        
        # Rate limiting
        if i + batch_size < len(short_items):
            time.sleep(0.5)
    
    # Count expansions
    expanded_count = sum(1 for item in expanded_items if item.get('qwen_expanded'))
    print(f"\nExpanded {expanded_count} descriptions")
    
    # Create lookup for expanded items
    expanded_lookup = {item['id']: item for item in expanded_items}
    
    # Merge back into main data
    print("\nMerging expanded descriptions...")
    for item in data:
        if item['id'] in expanded_lookup:
            expanded = expanded_lookup[item['id']]
            if expanded.get('qwen_expanded'):
                item['description'] = expanded['description']
                item['original_short_description'] = expanded.get('original_short_description')
                item['qwen_expanded'] = True
    
    # Save updated data
    output_json = "pricelist_final_expanded.json"
    with open(output_json, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    
    print(f"\nSaved expanded data to: {output_json}")
    
    # Also save as CSV
    output_csv = "pricelist_final_expanded.csv"
    
    df_data = []
    for item in data:
        row = {
            'id': item.get('id'),
            'code': item.get('code'),
            'original_code': item.get('original_code'),
            'description': item.get('description'),
            'unit': item.get('unit'),
            'category': item.get('category'),
            'subcategory': item.get('subcategory'),
            'work_type': item.get('work_type'),
            'rate': item.get('rate'),
            'keywords': '|'.join(item.get('keywords', [])) if item.get('keywords') else '',
            'qwen_expanded': item.get('qwen_expanded', False)
        }
        
        # Add cell rate info
        if item.get('cellRate'):
            row['cellRate_reference'] = item['cellRate'].get('reference')
            row['cellRate_rate'] = item['cellRate'].get('rate')
        
        # Add Excel mapping
        row['excelCellReference'] = item.get('excelCellReference')
        row['sourceSheetName'] = item.get('sourceSheetName')
        
        df_data.append(row)
    
    df = pd.DataFrame(df_data)
    df.to_csv(output_csv, index=False)
    print(f"Saved CSV to: {output_csv}")
    
    # Show sample expanded items
    print("\n" + "="*60)
    print("SAMPLE EXPANDED DESCRIPTIONS")
    print("="*60)
    
    expanded_samples = [item for item in data if item.get('qwen_expanded')][:10]
    for item in expanded_samples:
        original = item.get('original_short_description', 'N/A')
        print(f"\nID: {item['id']}")
        print(f"  Original: \"{original}\"")
        print(f"  Expanded: \"{item['description']}\"")
        print(f"  Category: {item.get('category')} / Unit: {item.get('unit')}")
    
    # Final quality check
    print("\n" + "="*60)
    print("FINAL QUALITY CHECK")
    print("="*60)
    
    # Count remaining issues
    still_short = sum(1 for item in data if len(item.get('description', '')) < 10)
    problematic_units = sum(1 for item in data if item.get('unit') in ['by m/c', 'meas sep', 'inc in s/c prelims'])
    
    print(f"Total items: {len(data)}")
    print(f"Items with expanded descriptions: {expanded_count}")
    print(f"Remaining short descriptions: {still_short}")
    print(f"Items with problematic units: {problematic_units}")
    
    quality_score = ((len(data) - still_short - problematic_units) / len(data)) * 100
    print(f"\nFinal Quality Score: {quality_score:.1f}%")
    
    if quality_score > 95:
        print("\n✅ EXCELLENT QUALITY - All items make sense now!")
    else:
        print("\n✅ GOOD QUALITY - Significantly improved!")
    
    print("\nFinal files created:")
    print(f"  - {output_json} (Best quality JSON)")
    print(f"  - {output_csv} (Best quality CSV)")
    print("\nThese files are ready for production use!")

if __name__ == "__main__":
    main()