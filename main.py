import sys
from pathlib import Path
from rich import print as printr

# Adjust these imports based on your exact package structure
from xl_ray.extractor import load_workbooks, extract_metadata
from xl_ray.schema import WorkbookMetadata 

def main():
    # Update this to point to your actual toy model Excel file
    # Using the path you mentioned in the task.md file earlier
    test_file = Path("example_audit.xlsm")

    if not test_file.exists():
        print(f"Error: Could not find {test_file}")
        print("Please update the test_file path in the script.")
        sys.exit(1)

    print(f"Loading '{test_file.name}'...")
    
    # 1. Dual-load the workbooks
    wb_formulas, wb_values = load_workbooks(test_file)
    print("Workbooks loaded successfully.")

    # We will pass empty lists for these until we write their extraction functions
    vba_modules_mock = [] 
    named_ranges_mock = []

    # 2. Extract metadata
    print("Extracting metadata...")
    metadata = extract_metadata(
        path=test_file, 
        wb_formulas=wb_formulas, 
        vba_modules=vba_modules_mock, 
        named_ranges=named_ranges_mock
    )

    # 3. Display the Pydantic model output
    printr("\n--- Metadata Result ---")
    # model_dump_json() is the standard Pydantic V2 way to export to JSON
    printr(metadata.model_dump_json(indent=2))

if __name__ == "__main__":
    main()