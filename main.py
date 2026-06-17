import sys
from pathlib import Path
from rich import print as printr

import gzip

# Adjust these imports based on your exact package structure
from xl_ray.extractor import (
    load_workbooks, 
    extract_metadata, 
    extract_vba_modules, 
    extract_named_ranges,
    extract_worksheets  # Import it at the top
)
from xl_ray.schema import ExcelModelData

from xl_ray.audit import (
    detect_magic_numbers, 
    detect_hidden_logic, 
    detect_broken_references,
    detect_inconsistent_columns,
    detect_complex_logic,
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

    print("Extracting Worksheets, Cells, and Tables (this may take a moment)...")
    worksheets = extract_worksheets(wb_formulas, wb_values)

    # --- Derived Data for Previews and Metadata ---
    # Flatten tables from all worksheets for our preview/flagging
    all_tables = [table for ws in worksheets.values() for table in ws.tables]
    
    # Check for array formulas across all extracted cells
    has_arrays = any(
        cell.is_array_formula 
        for ws in worksheets.values() 
        for cell in ws.cells.values()
    )

    print("Extracting metadata...")
    metadata = extract_metadata(
        path=test_file, 
        wb_formulas=wb_formulas, 
        vba_modules=vba_modules, 
        named_ranges=named_ranges
    )
    
    # Update dynamic flags
    metadata.has_data_tables = len(all_tables) > 0
    metadata.has_array_formulas = has_arrays

    # --- Display Previews ---
    printr("\n--- Excel Tables Preview ---")
    if not all_tables:
        printr("- No tables found in workbook.")
    for table in all_tables:
        printr(f"- [bold]{table.name}[/bold] (Range: {table.range_address})")
        printr(f"  Columns: {table.columns}")

    printr("\n--- Metadata Result ---")
    printr(metadata.model_dump_json(indent=2))

    # Construct the root schema
    print("Constructing root ExcelModelData...")
    excel_model = ExcelModelData(
        metadata=metadata,
        worksheets=worksheets,
        named_ranges=named_ranges,
        vba_modules=vba_modules
    )

    # --- OPTIONAL: Quick check for precedents ---
    printr("\n--- Precedents Check Preview ---")
    formula_count = 0
    for ws in worksheets.values():
        for cell in ws.cells.values():
            if cell.formula and cell.precedents:
                printr(f"- {ws.name}!{cell.address} -> {cell.formula}")
                printr(f"  Precedents: {cell.precedents}")
                formula_count += 1
                if formula_count >= 5: # Just show the first 5
                    break
        if formula_count >= 5:
            break

    # --- OPTIONAL: Quick check for array formulas ---
    printr("\n--- Array Formulas Check Preview ---")
    array_count = 0
    for ws in worksheets.values():
        for cell in ws.cells.values():
            if cell.is_array_formula:
                if cell.array_range:
                    printr(f"- [bold cyan]Parent Array Cell[/bold cyan]: {ws.name}!{cell.address} | Formula: {cell.formula} | Spills to: {cell.array_range}")
                    array_count += 1
                elif cell.parent_array_cell:
                    # Print a child cell, showing its value and where it points back to
                    printr(f"  - [magenta]Spilled Child Cell[/magenta]: {ws.name}!{cell.address} | Value: {cell.value} | Parent: {cell.parent_array_cell}")
                    array_count += 1
            
            # Stop after showing around 10 examples so we don't flood the terminal
            if array_count >= 10:
                break
        if array_count >= 10:
            break
            
    if array_count == 0:
        printr("- No array formulas found in this test workbook.")

    # --- Run Deterministic Audits ---
    

    printr("\n--- Audit: Magic Numbers ---")
    magic_numbers = detect_magic_numbers(excel_model)
    if magic_numbers:
        printr(f"[red]Found {len(magic_numbers)} cells with hardcoded magic numbers.[/red]")
        for item in magic_numbers[:5]: # Show first 5
            printr(f"  - {item['sheet']}!{item['cell']}: `{item['formula']}` (Found: {item['magic_numbers_found']})")
    else:
        printr("[green]No magic numbers found![/green]")


    printr("\n--- Audit: Hidden Logic ---")
    hidden_logic = detect_hidden_logic(excel_model)
    if hidden_logic:
        printr(f"[red]Found {len(hidden_logic)} cells referencing hidden sheets.[/red]")
        for item in hidden_logic[:5]:
            printr(f"  - {item['sheet']}!{item['cell']}: `{item['formula']}` -> Pulls from {item['hidden_references']}")
    else:
        printr("[green]No hidden logic dependencies found![/green]")

    
    printr("\n--- Audit: Broken References ---")
    broken_refs = detect_broken_references(excel_model)
    if broken_refs:
        printr(f"[red]Found {len(broken_refs)} cells with broken or error-state references.[/red]")
        for item in broken_refs[:5]:
            printr(f"  - {item['sheet']}!{item['cell']}: `{item['formula']}` (Issues: {item['issues_found']})")
    else:
        printr("[green]No broken references found![/green]")

    printr("\n--- Audit: Inconsistent Column Formulas ---")
    inconsistent_cols = detect_inconsistent_columns(excel_model)
    if inconsistent_cols:
        printr(f"[red]Found {len(inconsistent_cols)} cells with inconsistent formula patterns.[/red]")
        for item in inconsistent_cols[:5]:
            printr(f"  - {item['sheet']}!{item['cell']}: `{item['formula']}`")
            printr(f"    Expected: {item['expected_pattern']} | Actual: {item['actual_pattern']}")
    else:
        printr("[green]No column formula inconsistencies found![/green]")    


    printr("\n--- Audit: Logic Complexity & Depth ---")
    complex_logic = detect_complex_logic(excel_model, max_depth_threshold=10, complexity_score_threshold=8)
    
    if complex_logic:
        printr(f"[red]Found {len(complex_logic)} highly complex cells.[/red]")
        for item in complex_logic[:5]:
            printr(f"  - {item['sheet']}!{item['cell']}: `{item['formula']}`")
            printr(f"    Reason: {item['flag_reason']} | Score: {item['complexity_score']} | Depth: {item['chain_depth']}")
    else:
        printr("[green]No overly complex logic chains found![/green]")

        
    # ... (Then save your JSON) ...




    # Save to compressed JSON
    output_dir = Path("output")
    output_dir.mkdir(exist_ok=True) # Ensure output directory exists
    
    # Use .json.gz extension to indicate it's a gzipped JSON file
    output_file = output_dir / "extracted_model.json.gz" 
    
    print(f"\nSaving compressed final output to {output_file}...")
    
    # Dump the model to a JSON string first
    json_data = excel_model.model_dump_json(indent=2)
    
    # Write it using gzip
    with gzip.open(output_file, "wt", encoding="utf-8") as f:
        f.write(json_data)
        
    printr(f"[green]Done! Saved to {output_file}[/green]")    


if __name__ == "__main__":
    main()