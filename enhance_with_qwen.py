"""
Enhance already extracted pricelist with Qwen 72B
Processes the pricelist_v2.csv file and adds enhancements
"""

import pandas as pd
import json
import openai
import time
from pathlib import Path

# DeepInfra Configuration
DEEPINFRA_API_KEY = "8MSsOohjJBtIAlzstuh4inhRzgnuS68k"
DEEPINFRA_BASE_URL = "https://api.deepinfra.com/v1/openai"
DEEPINFRA_MODEL = "Qwen/Qwen2.5-72B-Instruct"

def enhance_batch(client, items_batch, category):
    """Enhance a batch of items with Qwen"""
    
    # Prepare batch for API
    batch_data = []
    for idx, row in items_batch.iterrows():
        batch_data.append({
            'index': idx,
            'code': row.get('code', ''),
            'description': row.get('description', ''),
            'unit': row.get('unit', ''),
            'subcategory': row.get('subcategory', '')
        })
    
    prompt = f"""You are enhancing construction pricelist items from "{category}" category.

Enhance these items:
{json.dumps(batch_data[:10], indent=2)}

For each item:
1. **Description**: Expand abbreviations (ne=not exceeding, thk=thick, exc=excavation, etc.)
2. **Unit**: Standardize (m, m², m³, nr, kg, tonnes, hour, day, week, sum, item)
3. **Keywords**: Generate 3-5 search keywords
4. **Work Type**: Identify type (Excavation, Concrete, Formwork, etc.)

Return JSON array with same index order:
[{{"index": 0, "description": "enhanced", "unit": "m²", "keywords": ["keyword1"], "work_type": "type"}}]"""

    try:
        response = client.chat.completions.create(
            model=DEEPINFRA_MODEL,
            messages=[
                {"role": "system", "content": "You are a construction expert. Return only valid JSON array."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.2,
            max_tokens=2000
        )
        
        result = response.choices[0].message.content.strip()
        
        # Clean JSON
        if '```json' in result:
            result = result.split('```json')[1].split('```')[0]
        elif '```' in result:
            result = result.split('```')[1].split('```')[0]
        
        # Find JSON array
        import re
        json_match = re.search(r'\[[\s\S]*\]', result)
        if json_match:
            result = json_match.group()
        
        enhanced_data = json.loads(result)
        return enhanced_data
        
    except Exception as e:
        print(f"  Error: {str(e)[:50]}")
        return []

def main():
    print("="*60)
    print("QWEN 72B ENHANCEMENT FOR EXTRACTED PRICELIST")
    print("="*60)
    
    # Load extracted data
    input_file = "pricelist_v2.csv"
    if not Path(input_file).exists():
        print(f"Error: {input_file} not found!")
        print("Run pricelist_extractor_v2.py first")
        return
    
    print(f"Loading {input_file}...")
    df = pd.read_csv(input_file)
    print(f"Loaded {len(df)} items")
    
    # Initialize Qwen client
    client = openai.OpenAI(
        api_key=DEEPINFRA_API_KEY,
        base_url=DEEPINFRA_BASE_URL
    )
    print("Qwen 72B client initialized")
    
    # Process in batches by category
    categories = df['category'].unique()
    total_enhanced = 0
    
    # Add new columns for enhancements
    df['enhanced_description'] = df['description']
    df['enhanced_unit'] = df['unit']
    df['keywords'] = ''
    df['work_type'] = ''
    
    for category in categories[:5]:  # Process first 5 categories
        cat_df = df[df['category'] == category]
        print(f"\nProcessing {category}: {len(cat_df)} items")
        
        # Process in batches of 10
        batch_size = 10
        for i in range(0, min(30, len(cat_df)), batch_size):  # Max 30 items per category
            batch = cat_df.iloc[i:i+batch_size]
            print(f"  Batch {i//batch_size + 1}...")
            
            enhanced = enhance_batch(client, batch, category)
            
            # Apply enhancements
            for item in enhanced:
                if 'index' in item:
                    idx = item['index']
                    if idx in df.index:
                        if 'description' in item:
                            df.at[idx, 'enhanced_description'] = item['description']
                        if 'unit' in item:
                            df.at[idx, 'enhanced_unit'] = item['unit']
                        if 'keywords' in item:
                            df.at[idx, 'keywords'] = '|'.join(item['keywords'])
                        if 'work_type' in item:
                            df.at[idx, 'work_type'] = item['work_type']
                        total_enhanced += 1
            
            time.sleep(0.5)  # Rate limiting
    
    print(f"\n{total_enhanced} items enhanced")
    
    # Save enhanced data
    output_file = "pricelist_v2_enhanced.csv"
    df.to_csv(output_file, index=False)
    print(f"\nSaved to: {output_file}")
    
    # Show samples
    print("\nSample enhanced items:")
    enhanced_items = df[df['keywords'] != ''].head(5)
    for idx, row in enhanced_items.iterrows():
        print(f"\nCode: {row['code']}")
        print(f"  Original: {row['description'][:50]}...")
        print(f"  Enhanced: {row['enhanced_description'][:50]}...")
        print(f"  Keywords: {row['keywords']}")
        print(f"  Work Type: {row['work_type']}")

if __name__ == "__main__":
    main()