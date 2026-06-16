from pathlib import Path
from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula

from .schema import WorkbookMetadata, VBAModule, NamedRange, ExcelTable, WorksheetData, CellData

from oletools.olevba import VBA_Parser

def extract_worksheets(wb_formulas, wb_values) -> dict[str, WorksheetData]:
    """
    Iterates through all worksheets, merges formula and value data, 
    and returns a dictionary of WorksheetData schemas.
    """
    worksheets_data = {}
    
    # Iterate through sheets in both workbooks simultaneously
    for sheet_f, sheet_v in zip(wb_formulas.worksheets, wb_values.worksheets):
        sheet_name = sheet_f.title
        
        # Map sheet_state to schema's expected visibility strings
        state_map = {"visible": "Visible", "hidden": "Hidden", "veryHidden": "VeryHidden"}
        visibility = state_map.get(sheet_f.sheet_state, "Visible")
        
        cells_dict = {}
        sheet_tables = extract_tables(sheet_f)
        
        # We need to iterate over all rows and cols that have data.
        # values_only=False gives us the actual Cell objects.
        for row_f, row_v in zip(sheet_f.iter_rows(values_only=False), sheet_v.iter_rows(values_only=False)):
            for cell_f, cell_v in zip(row_f, row_v):
                
                # Skip empty cells to save memory and JSON bloat
                if cell_f.value is None and cell_v.value is None:
                    continue
                    
                address = cell_f.coordinate
                
                # Default cell parameters
                formula_str = None
                is_array = False
                array_range = None
                
                # Check if it's a formula (data_type 'f')
                if cell_f.data_type == 'f':
                    val_f = cell_f.value
                    if isinstance(val_f, str):
                        formula_str = val_f
                    elif isinstance(val_f, ArrayFormula):
                        is_array = True
                        formula_str = str(getattr(val_f, "text", None) or "")
                        array_range = str(getattr(val_f, "ref", None) or "")
                    else:
                        # Fallback for other formula-bearing objects
                        formula_str = str(getattr(val_f, "value", None) or val_f)
                
                # Determine data type based on the evaluated value
                data_type = cell_v.data_type if cell_v.data_type else cell_f.data_type
                
                # Create the CellData instance
                cells_dict[address] = CellData(
                    address=address,
                    value=cell_v.value,  # Cached value from the data_only workbook
                    formula=formula_str,
                    is_array_formula=is_array,
                    array_range=array_range,
                    data_type=data_type
                    # Note: precedents, dependents, and parent_array_cell would require 
                    # advanced parsing/AST logic which we are bypassing for now.
                )
                
        worksheets_data[sheet_name] = WorksheetData(
            name=sheet_name,
            visibility=visibility,
            cells=cells_dict,
            tables=sheet_tables
        )
        
    return worksheets_data


def extract_vba_modules(path: Path) -> list[VBAModule]:
    """
    Extract VBA macros from the workbook and return them as VBAModule schemas.
    """
    if VBA_Parser is None:
        print("Warning: oletools not installed. Skipping VBA extraction.")
        return []

    modules = []
    try:
        parser = VBA_Parser(str(path))
        if parser.detect_vba_macros():
            # oletools extract_all_macros yields: (filename, stream_path, vba_filename, vba_code)
            for _, _, filename, content in parser.extract_all_macros():
                # Handle potential bytes output from oletools
                if isinstance(content, bytes):
                    content = content.decode("utf-8", errors="ignore")
                else:
                    content = str(content) if content is not None else ""
                
                modules.append(VBAModule(
                    filename=str(filename) or "Unknown",
                    content=content
                ))
    except Exception as e:
        print(f"Warning: Failed to extract VBA modules - {e}")
    finally:
        if 'parser' in locals():
            parser.close()
            
    return modules


def extract_named_ranges(wb) -> list[NamedRange]:
    """
    Extract workbook and worksheet scoped named ranges.
    """
    named_ranges = []
    
    # openpyxl >= 3.1 stores defined names in a dict-like object
    try:
        if hasattr(wb, "defined_names") and hasattr(wb.defined_names, "values"):
            dn_iter = wb.defined_names.values()
        else:
            dn_iter = wb.defined_names.definedName  # openpyxl < 3.1
    except AttributeError:
        return named_ranges

    for dn in dn_iter:
        if isinstance(dn, str): 
            continue  # Safeguard if iteration yields dict keys directly
        
        try:
            name = getattr(dn, "name", None)
            # The formula/reference is usually in 'attr_text' or 'value'
            refers_to = getattr(dn, "attr_text", None) or getattr(dn, "value", None)
            
            if not name or not refers_to:
                continue

            # Determine scope. localSheetId is not None if scoped to a specific sheet.
            localSheetId = getattr(dn, "localSheetId", None)
            if localSheetId is not None:
                try:
                    scope = wb.worksheets[int(localSheetId)].title
                except (IndexError, ValueError, TypeError):
                    scope = f"SheetId_{localSheetId}"
            else:
                scope = "Workbook"

            named_ranges.append(NamedRange(
                name=str(name),
                refers_to=str(refers_to),
                scope=scope
            ))
        except Exception as e:
            print(f"Warning: Failed to parse a named range - {e}")
            continue

    return named_ranges

def extract_tables(sheet) -> list[ExcelTable]:
    """
    Extract structured Excel Data Tables (ListObjects) from a single worksheet.
    """
    tables = []
    
    # openpyxl 3.x stores tables in sheet.tables, accessible via .values()
    try:
        sheet_tables = sheet.tables.values()
    except AttributeError:
        # Fallback for older openpyxl versions
        raw = getattr(sheet, "_tables", None)
        sheet_tables = raw.values() if isinstance(raw, dict) else (raw or [])

    for table in sheet_tables:
        if isinstance(table, str):
            continue  # Safeguard against key leakage
            
        name = getattr(table, "name", None) or getattr(table, "displayName", "UnknownTable")
        ref = getattr(table, "ref", None)
        
        # Extract column headers if available
        columns = []
        if hasattr(table, "tableColumns"):
            for col in table.tableColumns:
                if col.name:
                    columns.append(col.name)

        if name and ref:
            tables.append(ExcelTable(
                name=str(name),
                range_address=str(ref),
                columns=columns
            ))

    return tables

def load_workbooks(path: Path):
    """
    Dual-loads the workbook to extract both structural/formula data 
    and evaluated cached values.
    """
    # Load for structure, formulas, and VBA detection
    wb_formulas = load_workbook(str(path), data_only=False, read_only=False)
    
    # Load strictly for cached cell values
    wb_values = load_workbook(str(path), data_only=True, read_only=False)
    
    return wb_formulas, wb_values

def extract_metadata(path: Path, wb_formulas, vba_modules: list, named_ranges: list) -> WorkbookMetadata:
    """
    Generates the WorkbookMetadata model to provide an immediate summary.
    """
    # Check for hidden sheets
    has_hidden = any(sheet.sheet_state in ('hidden', 'veryHidden') for sheet in wb_formulas.worksheets)
    
    # Check for data tables (ListObjects) across all sheets
    has_tables = any(bool(sheet.tables) for sheet in wb_formulas.worksheets)
    
    # Array formulas require iterating through cells, but we can set this flag 
    # later during the WorksheetData/CellData extraction phase. 
    # Defaulting to False here.
    
    return WorkbookMetadata(
        file_name=path.name,
        file_path=str(path.absolute()),
        file_type=path.suffix.lower(),
        has_macros=len(vba_modules) > 0,
        has_array_formulas=False, # To be updated during cell extraction
        has_data_tables=has_tables,
        has_named_ranges=len(named_ranges) > 0,
        has_hidden_sheets=has_hidden,
        has_external_links=bool(wb_formulas._external_links),
        worksheet_count=len(wb_formulas.worksheets)
    )