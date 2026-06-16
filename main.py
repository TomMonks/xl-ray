import sys
from pathlib import Path
from rich import print as printr

# Adjust these imports based on your exact package structure
from xl_ray.extractor import (
    load_workbooks, 
    extract_metadata, 
    extract_vba_modules, 
    extract_named_ranges,
    extract_tables  # <-- Import the new function
)

def main():
    test_file = Path("example_audit.xlsm")

    if not test_file.exists():
        print(f"Error: Could not find {test_file}")
        sys.exit(1)

    print(f"Loading '{test_file.name}'...")
    wb_formulas, wb_values = load_workbooks(test_file)

    print("Extracting VBA modules...")
    vba_modules = extract_vba_modules(test_file)

    print("Extracting Named Ranges...")
    named_ranges = extract_named_ranges(wb_formulas)

    # --- NEW: Extracting Tables per Worksheet ---
    print("Extracting Excel Tables...")
    all_tables = []
    for sheet in wb_formulas.worksheets:
        sheet_tables = extract_tables(sheet)
        if sheet_tables:
            printr(f"  Found {len(sheet_tables)} table(s) on sheet '{sheet.title}'")
            all_tables.extend(sheet_tables)

    print("Extracting metadata...")
    metadata = extract_metadata(
        path=test_file, 
        wb_formulas=wb_formulas, 
        vba_modules=vba_modules, 
        named_ranges=named_ranges
    )
    # Update metadata based on our actual table extraction
    metadata.has_data_tables = len(all_tables) > 0

    # --- Display Table Previews ---
    printr("\n--- Excel Tables Preview ---")
    if not all_tables:
        printr("- No tables found in workbook.")
    for table in all_tables:
        printr(f"- [bold]{table.name}[/bold] (Range: {table.range_address})")
        printr(f"  Columns: {table.columns}")

    printr("\n--- Metadata Result ---")
    printr(metadata.model_dump_json(indent=2))

if __name__ == "__main__":
    main()