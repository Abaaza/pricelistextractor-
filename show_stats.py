import pandas as pd

df = pd.read_csv('groundworks_extracted.csv')
items_with_rates = df[df['rate'] > 0]

print('EXTRACTION STATISTICS')
print('='*60)
print(f'Total items extracted: {len(df)}')
print(f'Items with rates: {len(items_with_rates)}')
print(f'Items without rates: {len(df) - len(items_with_rates)}')
print(f'\nRate range: GBP {items_with_rates["rate"].min():.2f} - GBP {items_with_rates["rate"].max():.2f}')
print(f'Average rate: GBP {items_with_rates["rate"].mean():.2f}')
print(f'\nSubcategories found: {df["subcategory"].nunique()}')
print(f'\nTop 5 subcategories by item count:')
print(df["subcategory"].value_counts().head(5).to_string())

print(f'\n\nSAMPLE ITEMS WITH RATES:')
print('='*60)
sample = items_with_rates[['code', 'description', 'unit', 'rate', 'cellRate_reference', 'subcategory']].head(5)
for idx, row in sample.iterrows():
    print(f'\nCode: {row["code"]}')
    desc = row['description'][:80] + '...' if len(row['description']) > 80 else row['description']
    print(f'Description: {desc}')
    print(f'Unit: {row["unit"]} | Rate: GBP {row["rate"]:.2f}')
    print(f'Cell Reference: {row["cellRate_reference"]}')
    print(f'Subcategory: {row["subcategory"]}')
