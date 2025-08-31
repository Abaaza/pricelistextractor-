"""
Create the final high-quality pricelist by combining:
- Full extraction for non-Drainage sheets (more items)
- High-quality extraction for Drainage (with proper range descriptions)
"""

import pandas as pd
import json

print('Creating FINAL HIGH-QUALITY PRICELIST')
print('='*80)

# Load the original full extraction (has more items from other sheets)
df_full = pd.read_csv('full_pricelist.csv')
print(f'Original full extraction: {len(df_full)} items')

# Load the high-quality extraction (has better Drainage with ranges)
df_hq = pd.read_csv('high_quality_pricelist.csv')
print(f'High-quality extraction: {len(df_hq)} items')

# Get non-Drainage from full extraction
df_non_drainage = df_full[df_full['category'] != 'Drainage']
print(f'Non-Drainage items from full: {len(df_non_drainage)}')

# Get Drainage from high-quality extraction
df_drainage_hq = df_hq[df_hq['category'] == 'Drainage']
print(f'Drainage items from HQ: {len(df_drainage_hq)}')

# Combine them
df_final = pd.concat([df_non_drainage, df_drainage_hq], ignore_index=True)
print(f'\nFINAL COMBINED: {len(df_final)} items')

# Save the final high-quality pricelist
df_final.to_csv('final_high_quality_pricelist.csv', index=False)

# Save as JSON too
records = df_final.to_dict('records')
with open('final_high_quality_pricelist.json', 'w', encoding='utf-8') as f:
    json.dump(records, f, indent=2, ensure_ascii=False)

print('\nSaved as:')
print('  - final_high_quality_pricelist.csv')
print('  - final_high_quality_pricelist.json')

# Final statistics
print('\n' + '='*80)
print('FINAL HIGH-QUALITY PRICELIST STATISTICS')
print('='*80)

for cat in df_final['category'].unique():
    count = len(df_final[df_final['category'] == cat])
    with_rates = len(df_final[(df_final['category'] == cat) & (df_final['rate'] > 0)])
    print(f'{cat:20} - {count:5} items ({with_rates:5} with rates)')

print(f'\nTOTAL: {len(df_final)} items')
print(f'Items with rates: {len(df_final[df_final["rate"] > 0])} ({len(df_final[df_final["rate"] > 0])/len(df_final)*100:.1f}%)')
print(f'Items with cell refs: {len(df_final[df_final["cellRate_reference"].notna()])} ({len(df_final[df_final["cellRate_reference"].notna()])/len(df_final)*100:.1f}%)')

# Check sample Drainage range items
print('\n' + '='*80)
print('Sample Drainage items with proper excavation descriptions:')
print('='*80)

drainage_samples = df_final[(df_final['category'] == 'Drainage') & 
                            (df_final['description'].str.contains('depth to invert:', case=False, na=False))]

for _, item in drainage_samples.head(5).iterrows():
    print(f'\nCode: {item["code"]}')
    print(f'Description: {item["description"][:120]}')
    print(f'Rate: {item["rate"]} | Cell: {item["cellRate_reference"]}')

print('\n' + '='*80)
print('FINAL HIGH-QUALITY PRICELIST COMPLETE!')
print('='*80)