"""
Simple runner to add prefixes without user input
"""
import sys
sys.path.insert(0, r'C:\code\pricelistextractor-\Files')

from add_prefixes import add_prefix_to_file, FILENAME_PREFIX_MAP
from pathlib import Path

files_dir = Path(r'C:\code\pricelistextractor-\Files')

print("="*60)
print("ADD CATEGORY PREFIXES TO CSV FILES")
print("="*60)

# Find all CSV files to process
files_to_process = []
for filename, prefix in FILENAME_PREFIX_MAP.items():
    file_path = files_dir / filename

    if file_path.exists():
        files_to_process.append((file_path, prefix))
        print(f"[OK] Found: {filename} (prefix: {prefix})")
    else:
        print(f"[SKIP] Not found: {filename}")

print(f"\nFound {len(files_to_process)} files to process")
print("\nCreating files with '_updated.csv' suffix (keeping originals)...")

print("\n" + "="*60)
print("PROCESSING FILES")
print("="*60)

success_count = 0
for file_path, prefix in files_to_process:
    if add_prefix_to_file(file_path, prefix, overwrite=False):
        success_count += 1

print("\n" + "="*60)
print("SUMMARY")
print("="*60)
print(f"Successfully processed: {success_count}/{len(files_to_process)} files")
print("\n[SUCCESS] New files created with '_updated.csv' suffix")
print("Original files remain unchanged")

print("\nPrefixes added:")
for filename, prefix in FILENAME_PREFIX_MAP.items():
    print(f"  {filename:35} -> {prefix}")
