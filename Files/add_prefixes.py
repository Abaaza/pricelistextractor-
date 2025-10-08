"""
Add category prefixes to id and code columns in extracted CSV files
Keeps all other columns untouched
"""

import pandas as pd
import os
from pathlib import Path

# Filename to prefix mapping
FILENAME_PREFIX_MAP = {
    'drainage.csv': 'DRA',
    'groundworks_extracted.csv': 'GR',
    'external_works_extracted.csv': 'EXT',
    'rc_works_extracted.csv': 'RCW',
    'services_extracted.csv': 'SER',
    'underpinning_extracted.csv': 'UND',
}

def add_prefix_to_file(file_path, prefix, overwrite=False):
    """
    Add prefix to id and code columns in a CSV file

    Args:
        file_path: Path to CSV file
        prefix: Prefix to add (e.g., 'GR', 'DRA')
        overwrite: If True, overwrite original file. If False, create new file with _updated suffix
    """
    print(f"\nProcessing: {os.path.basename(file_path)}")
    print(f"Prefix: {prefix}")

    # Read CSV file
    df = pd.read_csv(file_path)
    original_count = len(df)
    print(f"Loaded {original_count} items")

    # Check if required columns exist
    if 'id' not in df.columns or 'code' not in df.columns:
        print(f"ERROR: Missing 'id' or 'code' column in {file_path}")
        return False

    # Store original column order
    original_columns = df.columns.tolist()

    # Add prefix to id and code columns (keep all other columns untouched)
    df['id'] = df['id'].apply(lambda x: f"{prefix}{x}")
    df['code'] = df['code'].apply(lambda x: f"{prefix}{x}")

    # Ensure column order is preserved
    df = df[original_columns]

    # Determine output file path
    if overwrite:
        output_path = file_path
    else:
        output_path = str(file_path).replace('.csv', '_updated.csv')

    # Save updated CSV
    df.to_csv(output_path, index=False)

    print(f"[OK] Saved to: {os.path.basename(output_path)}")
    print(f"  Sample IDs: {df['id'].head(3).tolist()}")

    return True

def main():
    """Main execution"""
    print("="*60)
    print("ADD CATEGORY PREFIXES TO CSV FILES")
    print("="*60)

    # Get Files directory (where script is located)
    files_dir = Path(__file__).parent

    # Find all CSV files to process (only look in Files folder)
    files_to_process = []
    for filename, prefix in FILENAME_PREFIX_MAP.items():
        file_path = files_dir / filename

        if file_path.exists():
            files_to_process.append((file_path, prefix))
            print(f"[OK] Found: {filename} (prefix: {prefix})")
        else:
            print(f"[SKIP] Not found: {filename}")

    if not files_to_process:
        print("\nNo files found to process!")
        return

    print(f"\nFound {len(files_to_process)} files to process")

    # Ask user if they want to overwrite or create new files
    print("\nOptions:")
    print("  1. Overwrite original files (replace in place)")
    print("  2. Create new files with '_updated.csv' suffix (keep originals)")

    choice = input("\nEnter choice (1 or 2, default=2): ").strip()
    overwrite = (choice == '1')

    if overwrite:
        confirm = input("WARNING: This will overwrite original files. Are you sure? (yes/no): ").strip().lower()
        if confirm != 'yes':
            print("Cancelled.")
            return

    # Process each file
    print("\n" + "="*60)
    print("PROCESSING FILES")
    print("="*60)

    success_count = 0
    for file_path, prefix in files_to_process:
        if add_prefix_to_file(file_path, prefix, overwrite):
            success_count += 1

    # Summary
    print("\n" + "="*60)
    print("SUMMARY")
    print("="*60)
    print(f"Successfully processed: {success_count}/{len(files_to_process)} files")

    if overwrite:
        print("\n[SUCCESS] Original files have been updated with prefixes")
    else:
        print("\n[SUCCESS] New files created with '_updated.csv' suffix")
        print("Original files remain unchanged")

    print("\nPrefixes added:")
    for filename, prefix in FILENAME_PREFIX_MAP.items():
        print(f"  {filename:35} -> {prefix}")

if __name__ == "__main__":
    main()
