"""
Fast Qwen 72B Short Description Fixer
Optimized for speed with smaller batches and better error handling
"""

import json
import pandas as pd
import openai
import time
import re
from pathlib import Path
from typing import List, Dict, Any

# DeepInfra Configuration
DEEPINFRA_API_KEY = "8MSsOohjJBtIAlzstuh4inhRzgnuS68k"
DEEPINFRA_BASE_URL = "https://api.deepinfra.com/v1/openai"
DEEPINFRA_MODEL = "Qwen/Qwen2.5-72B-Instruct"

def find_short_descriptions(data: List[Dict]) -> List[Dict]:
    """Find all items with very short descriptions"""
    short_items = []
    
    for item in data:
        desc = item.get('description', '')
        
        # Very short descriptions (< 10 chars) or single words
        if len(desc) < 10 or (len(desc.split()) == 1 and len(desc) < 20):
            short_items.append(item)
    
    return short_items

def expand_descriptions_batch(client, batch: List[Dict], batch_num: int) -> List[Dict]:
    """Expand a small batch of short descriptions"""
    
    # Prepare minimal data for API
    items_text = []
    for i, item in enumerate(batch):
        items_text.append({
            'idx': i,
            'desc': item.get('description', ''),
            'cat': item.get('category', ''),
            'unit': item.get('unit', '')
        })
    
    prompt = f"""Expand these short construction descriptions to full professional descriptions:

{json.dumps(items_text, indent=2)}

Rules:
- "Bollard" -> "Supply and install concrete/steel bollard"
- "Terram" -> "Terram geotextile membrane"
- "Welder" -> "Skilled welder (daily rate)"
- "Mobilise" -> "Mobilisation of plant and equipment to site"
- Keep 20-100 characters
- Include "supply and install" or "labour only" where appropriate

Return JSON: [{{"idx": 0, "expanded": "full description"}}]"""

    try:
        response = client.chat.completions.create(
            model=DEEPINFRA_MODEL,
            messages=[
                {"role": "system", "content": "Expand short construction terms. Return only JSON array."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.2,
            max_tokens=500,
            timeout=10  # 10 second timeout
        )
        
        result = response.choices[0].message.content.strip()
        
        # Extract JSON
        if '```' in result:
            result = result.split('```')[1].replace('json', '').strip()
        
        json_match = re.search(r'\[[\s\S]*\]', result)
        if json_match:
            result = json_match.group()
        
        expanded_data = json.loads(result)
        
        # Apply expansions
        for item in batch:
            item_index = batch.index(item)
            for expanded in expanded_data:
                if expanded.get('idx') == item_index:
                    if expanded.get('expanded'):
                        item['original_short'] = item.get('description')
                        item['description'] = expanded['expanded']
                        item['qwen_expanded'] = True
                    break
        
        return batch
        
    except Exception as e:
        print(f"    Batch {batch_num} error: {str(e)[:50]}")
        return batch

def main():
    print("="*60)
    print("FAST SHORT DESCRIPTION FIXER WITH QWEN 72B")
    print("="*60)
    
    # Load data
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
    for item in short_items[:5]:
        print(f"  {item['id']}: \"{item.get('description', '')}\" in {item.get('category', '')}")
    
    # Initialize client
    print("\nInitializing Qwen 72B client...")
    client = openai.OpenAI(
        api_key=DEEPINFRA_API_KEY,
        base_url=DEEPINFRA_BASE_URL
    )
    
    # Process in very small batches
    batch_size = 5  # Small batches for speed
    total_batches = (len(short_items) + batch_size - 1) // batch_size
    
    print(f"\nProcessing {len(short_items)} items in {total_batches} small batches...")
    print("This will be fast!\n")
    
    expanded_count = 0
    for i in range(0, len(short_items), batch_size):
        batch = short_items[i:i+batch_size]
        batch_num = (i // batch_size) + 1
        
        print(f"Batch {batch_num}/{total_batches} ({len(batch)} items)...", end="")
        
        # Process batch
        start_time = time.time()
        expanded_batch = expand_descriptions_batch(client, batch, batch_num)
        
        # Count expansions
        batch_expanded = sum(1 for item in expanded_batch if item.get('qwen_expanded'))
        expanded_count += batch_expanded
        
        elapsed = time.time() - start_time
        print(f" done! ({batch_expanded} expanded, {elapsed:.1f}s)")
        
        # Very short delay between batches
        if i + batch_size < len(short_items):
            time.sleep(0.2)
    
    print(f"\nTotal expanded: {expanded_count} descriptions")
    
    # Create lookup for expanded items
    expanded_lookup = {item['id']: item for item in short_items if item.get('qwen_expanded')}
    
    # Merge back into main data
    print("\nMerging expanded descriptions...")
    for item in data:
        if item['id'] in expanded_lookup:
            expanded = expanded_lookup[item['id']]
            item['description'] = expanded['description']
            item['original_short_description'] = expanded.get('original_short')
            item['qwen_expanded'] = True
    
    # Save updated data
    output_json = "pricelist_final_perfect.json"
    with open(output_json, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    
    print(f"\nSaved to: {output_json}")
    
    # Also save as CSV
    output_csv = "pricelist_final_perfect.csv"
    
    df_data = []
    for item in data:
        row = {
            'id': item.get('id'),
            'code': item.get('code'),
            'description': item.get('description'),
            'unit': item.get('unit'),
            'category': item.get('category'),
            'subcategory': item.get('subcategory'),
            'rate': item.get('rate'),
            'qwen_expanded': item.get('qwen_expanded', False)
        }
        df_data.append(row)
    
    df = pd.DataFrame(df_data)
    df.to_csv(output_csv, index=False)
    print(f"Saved CSV to: {output_csv}")
    
    # Show improvements
    print("\n" + "="*60)
    print("SAMPLE IMPROVEMENTS")
    print("="*60)
    
    improved_samples = [item for item in data if item.get('qwen_expanded')][:5]
    for item in improved_samples:
        original = item.get('original_short_description', 'N/A')
        print(f"\nID: {item['id']}")
        print(f"  Before: \"{original}\"")
        print(f"  After:  \"{item['description']}\"")
    
    # Final quality check
    remaining_short = sum(1 for item in data if len(item.get('description', '')) < 10)
    print("\n" + "="*60)
    print("FINAL QUALITY")
    print("="*60)
    print(f"Total items: {len(data)}")
    print(f"Expanded items: {expanded_count}")
    print(f"Remaining short: {remaining_short}")
    print(f"Quality score: {((len(data) - remaining_short) / len(data)) * 100:.1f}%")
    
    print("\nAll items should make sense now!")
    print("Final files:")
    print(f"  - {output_json}")
    print(f"  - {output_csv}")

if __name__ == "__main__":
    main()