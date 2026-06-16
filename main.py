import sys
from pathlib import Path
from rich import print as printr

# Adjust these imports based on your exact package structure
from xl_ray.extractor import (
    load_workbooks, 
    extract_metadata, 
    extract_vba_modules, 
    extract_named_ranges
)
from xl_ray.schema import WorkbookMetadata 

def main():
    test_file = Path("example_audit.xlsm")

    if not test_file.exists():
        print(f"Error: Could not find {test_file}")
        sys.exit(1)

    print(f"Loading '{test_file.name}'...")
    
    # 1. Dual-load the workbooks
    wb_formulas, wb_values = load_workbooks(test_file)
    print("Workbooks loaded successfully.")

    # 2. Extract VBA and Named Ranges
    print("Extracting VBA modules...")
    vba_modules = extract_vba_modules(test_file)
    print(f"Found {len(vba_modules)} VBA modules.")

    print("Extracting Named Ranges...")
    named_ranges = extract_named_ranges(wb_formulas)
    print(f"Found {len(named_ranges)} named ranges.")

    # 3. Extract metadata
    print("Extracting metadata...")
    metadata = extract_metadata(
        path=test_file, 
        wb_formulas=wb_formulas, 
        vba_modules=vba_modules, 
        named_ranges=named_ranges
    )

    # 4. Display the results
    printr("\n--- VBA Modules Preview ---")
    for vba in vba_modules:
        printr(f"- {vba.filename}: {len(vba.content)} characters")

    printr("\n--- Named Ranges Preview ---")
    for nr in named_ranges[:5]:  # Just show the first 5 so it doesn't flood the terminal
        printr(f"- {nr.name} ({nr.scope}): {nr.refers_to}")

    printr("\n--- Metadata Result ---")
    printr(metadata.model_dump_json(indent=2))

if __name__ == "__main__":
    main()