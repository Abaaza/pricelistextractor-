"""
Combine all extracted CSV and JSON files into single comprehensive files
Merges all sheet extractions into one master pricelist
"""

import pandas as pd
import json
import os
from datetime import datetime

def combine_csv_files():
    """Combine all CSV files into one master CSV"""
    
    # Define the files to combine
    csv_files = [
        'Files/drainage.csv',
        'Files/external_works_extracted.csv',
        'Files/groundworks_extracted.csv',
        'Files/rc_works_extracted.csv',
        'Files/services_extracted.csv',
        'Files/underpinning_extracted.csv'
    ]
    
    # List to store all dataframes
    all_dfs = []
    
    # Track statistics
    stats = {}
    
    print("="*60)
    print("COMBINING ALL EXTRACTED CSV FILES")
    print("="*60)
    
    # Read each CSV file
    for csv_file in csv_files:
        if os.path.exists(csv_file):
            print(f"\nReading {csv_file}...")
            df = pd.read_csv(csv_file)
            
            # Get sheet name from the file
            sheet_name = os.path.basename(csv_file).replace('_extracted.csv', '').replace('.csv', '')
            
            # Store statistics
            stats[sheet_name] = {
                'total_items': len(df),
                'items_with_rates': len(df[df['rate'] > 0]) if 'rate' in df.columns else 0
            }
            
            # Ensure ID uniqueness by adding prefix based on category
            if 'id' in df.columns:
                # Keep original ID but make sure it's unique across all sheets
                max_id = max(all_dfs[-1]['id'].max() if all_dfs else 0, 0)
                df['original_id'] = df['id']
                df['id'] = df.index + max_id + 1
            
            all_dfs.append(df)
            print(f"  - Loaded {len(df)} items from {sheet_name}")
        else:
            print(f"  - Warning: {csv_file} not found")
    
    # Combine all dataframes
    if all_dfs:
        combined_df = pd.concat(all_dfs, ignore_index=True)
        
        # Reset IDs to be sequential
        combined_df['id'] = range(1, len(combined_df) + 1)
        
        # Drop the temporary original_id column if it exists
        if 'original_id' in combined_df.columns:
            combined_df = combined_df.drop('original_id', axis=1)
        
        # Save combined CSV
        output_csv = 'Files/pricelist_combined_all.csv'
        combined_df.to_csv(output_csv, index=False)
        print(f"\n✅ Combined CSV saved to: {output_csv}")
        
        # Print statistics
        print("\n" + "="*60)
        print("EXTRACTION STATISTICS BY SHEET")
        print("="*60)
        for sheet, stat in stats.items():
            print(f"{sheet:20} - Items: {stat['total_items']:5} | With rates: {stat['items_with_rates']:5}")
        
        print("\n" + "="*60)
        print("COMBINED TOTALS")
        print("="*60)
        print(f"Total items: {len(combined_df)}")
        if 'rate' in combined_df.columns:
            items_with_rates = len(combined_df[combined_df['rate'] > 0])
            print(f"Items with rates: {items_with_rates} ({items_with_rates/len(combined_df)*100:.1f}%)")
            
            # Category distribution
            if 'category' in combined_df.columns:
                print("\nItems by category:")
                category_counts = combined_df['category'].value_counts()
                for cat, count in category_counts.items():
                    print(f"  {cat}: {count}")
        
        return combined_df
    else:
        print("No CSV files found to combine")
        return None

def combine_json_files():
    """Combine all JSON files into one master JSON"""
    
    # Define the files to combine
    json_files = [
        'Files/drainage.json',
        'Files/external_works_extracted.json',
        'Files/groundworks_extracted.json',
        'Files/rc_works_extracted.json',
        'Files/services_extracted.json',
        'Files/underpinning_extracted.json'
    ]
    
    # List to store all items
    all_items = []
    
    print("\n" + "="*60)
    print("COMBINING ALL EXTRACTED JSON FILES")
    print("="*60)
    
    # Read each JSON file
    for json_file in json_files:
        if os.path.exists(json_file):
            print(f"\nReading {json_file}...")
            with open(json_file, 'r', encoding='utf-8') as f:
                items = json.load(f)
                
            sheet_name = os.path.basename(json_file).replace('_extracted.json', '').replace('.json', '')
            
            # Add items to the list
            all_items.extend(items)
            print(f"  - Loaded {len(items)} items from {sheet_name}")
        else:
            print(f"  - Warning: {json_file} not found")
    
    # Update IDs to be sequential
    for idx, item in enumerate(all_items, 1):
        item['id'] = idx
    
    # Save combined JSON
    if all_items:
        output_json = 'Files/pricelist_combined_all.json'
        with open(output_json, 'w', encoding='utf-8') as f:
            json.dump(all_items, f, indent=2, ensure_ascii=False)
        print(f"\n✅ Combined JSON saved to: {output_json}")
        print(f"Total items in JSON: {len(all_items)}")
        
        return all_items
    else:
        print("No JSON files found to combine")
        return None

def create_summary_report(df):
    """Create a summary report of the combined data"""
    
    if df is None or df.empty:
        return
    
    print("\n" + "="*60)
    print("CREATING SUMMARY REPORT")
    print("="*60)
    
    summary = []
    summary.append("PRICELIST COMBINED SUMMARY REPORT")
    summary.append("=" * 60)
    summary.append(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    summary.append("")
    
    # Overall statistics
    summary.append("OVERALL STATISTICS")
    summary.append("-" * 40)
    summary.append(f"Total items: {len(df)}")
    
    if 'rate' in df.columns:
        items_with_rates = len(df[df['rate'] > 0])
        summary.append(f"Items with rates: {items_with_rates} ({items_with_rates/len(df)*100:.1f}%)")
        
        rates = df[df['rate'] > 0]['rate']
        if not rates.empty:
            summary.append(f"Rate range: £{rates.min():.2f} - £{rates.max():.2f}")
            summary.append(f"Average rate: £{rates.mean():.2f}")
            summary.append(f"Median rate: £{rates.median():.2f}")
    
    summary.append("")
    
    # Category breakdown
    if 'category' in df.columns:
        summary.append("ITEMS BY CATEGORY")
        summary.append("-" * 40)
        category_counts = df['category'].value_counts()
        for cat, count in category_counts.items():
            percentage = count / len(df) * 100
            summary.append(f"{cat:20} {count:6} items ({percentage:5.1f}%)")
        summary.append("")
    
    # Subcategory breakdown (top 20)
    if 'subcategory' in df.columns:
        summary.append("TOP 20 SUBCATEGORIES")
        summary.append("-" * 40)
        subcat_counts = df['subcategory'].value_counts().head(20)
        for subcat, count in subcat_counts.items():
            summary.append(f"{subcat:30} {count:6} items")
        summary.append("")
    
    # Unit distribution
    if 'unit' in df.columns:
        summary.append("UNIT DISTRIBUTION")
        summary.append("-" * 40)
        unit_counts = df['unit'].value_counts()
        for unit, count in unit_counts.items():
            percentage = count / len(df) * 100
            summary.append(f"{unit:10} {count:6} items ({percentage:5.1f}%)")
    
    # Save summary report
    summary_text = '\n'.join(summary)
    summary_file = 'Files/pricelist_summary_report.txt'
    with open(summary_file, 'w', encoding='utf-8') as f:
        f.write(summary_text)
    
    print(f"\n✅ Summary report saved to: {summary_file}")
    print("\nReport Preview:")
    print("-" * 40)
    for line in summary[:20]:
        print(line)

def main():
    """Main function to combine all files"""
    
    print("\n" + "🔧 PRICELIST COMBINER TOOL 🔧")
    print("="*60)
    print("Combining all extracted pricelist files...")
    print("="*60)
    
    # Combine CSV files
    combined_df = combine_csv_files()
    
    # Combine JSON files
    combined_json = combine_json_files()
    
    # Create summary report
    if combined_df is not None:
        create_summary_report(combined_df)
    
    print("\n" + "="*60)
    print("✅ ALL FILES COMBINED SUCCESSFULLY!")
    print("="*60)
    print("\nOutput files created:")
    print("  📊 Files/pricelist_combined_all.csv")
    print("  📋 Files/pricelist_combined_all.json")
    print("  📝 Files/pricelist_summary_report.txt")
    print("\n🎉 Process complete!")

if __name__ == "__main__":
    main()