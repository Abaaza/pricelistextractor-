"""
Deep Clean with Qwen 72B - Removes nonsense and improves quality
Uses DeepInfra Qwen to intelligently clean problematic items
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

class QwenCleaner:
    """Deep cleaning with Qwen 72B"""
    
    def __init__(self):
        self.client = openai.OpenAI(
            api_key=DEEPINFRA_API_KEY,
            base_url=DEEPINFRA_BASE_URL
        )
        self.api_calls = 0
        self.cleaned_count = 0
        
    def identify_problematic_items(self, items):
        """Identify items that need cleaning"""
        problematic = []
        
        for item in items:
            needs_cleaning = False
            reasons = []
            
            # Check description quality
            desc = item.get('description', '')
            if desc:
                # Too short
                if len(desc) < 10:
                    needs_cleaning = True
                    reasons.append('short_description')
                
                # Contains nonsense patterns
                nonsense_patterns = [
                    r'^[0-9\.\-]+$',  # Just numbers
                    r'^[\W_]+$',      # Just symbols
                    r'^\s*$',         # Empty/whitespace
                    r'^Item \d+$',    # Generic "Item 1"
                    r'^-+$',          # Just dashes
                ]
                
                for pattern in nonsense_patterns:
                    if re.match(pattern, desc):
                        needs_cleaning = True
                        reasons.append('nonsense_pattern')
                        break
                
                # Truncated descriptions
                if desc.endswith('...') or desc.endswith('inc for'):
                    needs_cleaning = True
                    reasons.append('truncated')
            else:
                needs_cleaning = True
                reasons.append('missing_description')
            
            # Check unit quality
            unit = item.get('unit', '')
            problematic_units = ['by m/c', 'meas sep', 'inc in s/c prelims', 'by client', 'm rise']
            if unit in problematic_units:
                needs_cleaning = True
                reasons.append('problematic_unit')
            
            # Check for missing critical fields
            if not item.get('category'):
                needs_cleaning = True
                reasons.append('missing_category')
            
            if needs_cleaning:
                item['cleaning_reasons'] = reasons
                problematic.append(item)
        
        return problematic
    
    def clean_batch_with_qwen(self, batch, attempt_num=1):
        """Clean a batch of problematic items with Qwen"""
        
        # Prepare batch for API
        items_text = []
        for i, item in enumerate(batch):
            items_text.append({
                'index': i,
                'id': item.get('id', ''),
                'code': item.get('code', ''),
                'description': item.get('description', ''),
                'unit': item.get('unit', ''),
                'category': item.get('category', ''),
                'problems': item.get('cleaning_reasons', [])
            })
        
        prompt = f"""You are a construction cost estimator cleaning up problematic pricelist items.

These items have quality issues that need fixing:
{json.dumps(items_text, indent=2)}

For each item, provide a CLEAN version by:

1. **Description**: 
   - If truncated, complete based on construction context
   - If nonsense/empty, infer from code/category or mark as "REMOVE"
   - Fix typos and expand abbreviations
   - Make it a proper construction item description

2. **Unit**: Fix problematic units
   - "by m/c" → "item" or appropriate unit
   - "meas sep" → "item" 
   - "inc in s/c prelims" → "included"
   - "by client" → "provisional"
   - "m rise" → "m"
   - Infer from description if possible

3. **Should Keep**: Decide if item should be kept
   - Mark "REMOVE" if it's genuinely nonsense
   - Keep if it can be cleaned to make sense

4. **Category/Subcategory**: Verify or fix based on description

Return JSON array:
[{{
  "index": 0,
  "description": "cleaned description or REMOVE",
  "unit": "proper unit",
  "should_keep": true/false,
  "category": "verified category",
  "subcategory": "appropriate subcategory",
  "confidence": 0.0-1.0
}}]

Be conservative - only mark REMOVE if truly nonsense. Most construction items can be cleaned."""

        try:
            response = self.client.chat.completions.create(
                model=DEEPINFRA_MODEL,
                messages=[
                    {"role": "system", "content": "You are a construction expert. Fix problematic items or mark them for removal."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.3,
                max_tokens=2000
            )
            
            self.api_calls += 1
            result = response.choices[0].message.content.strip()
            
            # Parse JSON
            if '```json' in result:
                result = result.split('```json')[1].split('```')[0]
            elif '```' in result:
                result = result.split('```')[1].split('```')[0]
            
            # Extract JSON array
            import re
            json_match = re.search(r'\[[\s\S]*\]', result)
            if json_match:
                result = json_match.group()
            
            cleaned_data = json.loads(result)
            
            # Apply cleaning
            cleaned_batch = []
            for item in batch:
                item_index = batch.index(item)
                
                # Find corresponding cleaned version
                for cleaned in cleaned_data:
                    if cleaned.get('index') == item_index:
                        # Apply cleaning
                        if cleaned.get('should_keep', True):
                            item['description'] = cleaned.get('description', item['description'])
                            item['unit'] = cleaned.get('unit', item.get('unit', 'item'))
                            item['category'] = cleaned.get('category', item.get('category'))
                            item['subcategory'] = cleaned.get('subcategory')
                            item['qwen_confidence'] = cleaned.get('confidence', 0.5)
                            item['qwen_cleaned'] = True
                            cleaned_batch.append(item)
                            self.cleaned_count += 1
                        else:
                            # Mark for removal
                            item['should_remove'] = True
                            item['removal_reason'] = 'Qwen identified as nonsense'
                        break
                else:
                    # No cleaning found, keep original
                    cleaned_batch.append(item)
            
            return cleaned_batch
            
        except Exception as e:
            print(f"    Error in batch {attempt_num}: {str(e)[:50]}")
            # Return original on error
            return batch
    
    def deep_clean(self, json_file):
        """Perform deep cleaning on entire dataset"""
        print("="*60)
        print("DEEP CLEANING WITH QWEN 72B")
        print("="*60)
        
        # Load data
        print(f"\nLoading {json_file}...")
        with open(json_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        print(f"Loaded {len(data)} items")
        
        # Identify problematic items
        print("\nIdentifying problematic items...")
        problematic = self.identify_problematic_items(data)
        print(f"Found {len(problematic)} items needing cleaning")
        
        if not problematic:
            print("No problematic items found!")
            return data
        
        # Show sample problems
        print("\nSample problematic items:")
        for item in problematic[:5]:
            print(f"  {item['id']}: {item.get('description', 'NO DESC')[:40]}... [{', '.join(item['cleaning_reasons'])}]")
        
        # Process in batches
        print(f"\nCleaning with Qwen 72B...")
        batch_size = 5
        cleaned_items = []
        items_to_remove = []
        
        for i in range(0, len(problematic), batch_size):
            batch = problematic[i:i+batch_size]
            batch_num = (i // batch_size) + 1
            total_batches = (len(problematic) + batch_size - 1) // batch_size
            
            print(f"  Processing batch {batch_num}/{total_batches}...")
            cleaned_batch = self.clean_batch_with_qwen(batch, batch_num)
            
            for item in cleaned_batch:
                if item.get('should_remove'):
                    items_to_remove.append(item)
                else:
                    cleaned_items.append(item)
            
            # Rate limiting
            if i + batch_size < len(problematic):
                time.sleep(0.5)
        
        print(f"\nCleaning complete!")
        print(f"  Items cleaned: {self.cleaned_count}")
        print(f"  Items marked for removal: {len(items_to_remove)}")
        print(f"  API calls made: {self.api_calls}")
        
        # Merge cleaned items back
        print("\nMerging cleaned items...")
        
        # Create lookup for cleaned items
        cleaned_lookup = {item['id']: item for item in cleaned_items}
        remove_ids = {item['id'] for item in items_to_remove}
        
        # Update original data
        final_data = []
        for item in data:
            item_id = item['id']
            
            # Skip if marked for removal
            if item_id in remove_ids:
                continue
            
            # Use cleaned version if available
            if item_id in cleaned_lookup:
                cleaned = cleaned_lookup[item_id]
                # Update fields
                item['description'] = cleaned['description']
                item['unit'] = cleaned['unit']
                item['category'] = cleaned.get('category', item.get('category'))
                item['subcategory'] = cleaned.get('subcategory', item.get('subcategory'))
                item['qwen_cleaned'] = True
            
            final_data.append(item)
        
        print(f"Final dataset: {len(final_data)} items (removed {len(data) - len(final_data)})")
        
        return final_data

def main():
    """Main execution"""
    
    # Input file
    input_file = "pricelist_final_clean.json"
    if not Path(input_file).exists():
        print(f"Error: {input_file} not found!")
        print("Run clean_final_extraction.py first")
        return
    
    # Initialize cleaner
    cleaner = QwenCleaner()
    
    # Perform deep cleaning
    cleaned_data = cleaner.deep_clean(input_file)
    
    # Save cleaned data
    output_json = "pricelist_qwen_cleaned.json"
    with open(output_json, 'w', encoding='utf-8') as f:
        json.dump(cleaned_data, f, indent=2, ensure_ascii=False)
    
    print(f"\nSaved Qwen-cleaned data to: {output_json}")
    
    # Also save as CSV
    output_csv = "pricelist_qwen_cleaned.csv"
    
    # Convert to DataFrame
    df_data = []
    for item in cleaned_data:
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
            'keywords': '|'.join(item.get('keywords', [])),
            'qwen_cleaned': item.get('qwen_cleaned', False)
        }
        
        # Add cell rate info
        if item.get('cellRate'):
            row['cellRate_reference'] = item['cellRate'].get('reference')
            row['cellRate_rate'] = item['cellRate'].get('rate')
        
        df_data.append(row)
    
    df = pd.DataFrame(df_data)
    df.to_csv(output_csv, index=False)
    print(f"Saved CSV to: {output_csv}")
    
    # Show sample cleaned items
    print("\n" + "="*60)
    print("SAMPLE QWEN-CLEANED ITEMS")
    print("="*60)
    
    qwen_cleaned = [item for item in cleaned_data if item.get('qwen_cleaned')][:5]
    for item in qwen_cleaned:
        print(f"\nID: {item['id']}")
        print(f"  Description: {item['description'][:70]}...")
        print(f"  Unit: {item['unit']}")
        print(f"  Category: {item['category']}")
        if item.get('subcategory'):
            print(f"  Subcategory: {item['subcategory']}")
    
    print("\n✅ Deep cleaning with Qwen complete!")
    print("\nFinal files:")
    print(f"  - {output_json} (JSON format)")
    print(f"  - {output_csv} (CSV format)")

if __name__ == "__main__":
    main()