from pathlib import Path
from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.formula import Tokenizer
from openpyxl.utils import range_boundaries, get_column_letter

from .schema import WorkbookMetadata, VBAModule, NamedRange, ExcelTable, WorksheetData, CellData

from oletools.olevba import VBA_Parser

def extract_precedents_from_formula(formula_str: str) -> list[str]:
    """
    Parses an Excel formula and returns a list of referenced cell addresses, 
    ranges, or named ranges.
    """
    if not formula_str or not formula_str.startswith('='):
        return []
        
    precedents = set()
    try:
        tok = Tokenizer(formula_str)
        for t in tok.items:
            # Tokenizer identifies ranges (A1, A1:B2, 'Sheet1'!A1) and names
            if t.type == 'OPERAND' and t.subtype == 'RANGE':
                # Strip out absolute reference markers ($) to keep things clean
                clean_ref = t.value.replace('$', '')
                precedents.add(clean_ref)
    except Exception:
        # Fail gracefully on extremely complex or malformed formulas
        pass
        
    return sorted(list(precedents))


def extract_worksheets(wb_formulas, wb_values) -> dict[str, WorksheetData]:
    worksheets_data = {}
    
    for sheet_f, sheet_v in zip(wb_formulas.worksheets, wb_values.worksheets):
        sheet_name = sheet_f.title
        state_map = {"visible": "Visible", "hidden": "Hidden", "veryHidden": "VeryHidden"}
        visibility = state_map.get(sheet_f.sheet_state, "Visible")
        
        cells_dict = {}
        sheet_tables = extract_tables(sheet_f)
        
        for row_f, row_v in zip(sheet_f.iter_rows(values_only=False), sheet_v.iter_rows(values_only=False)):
            for cell_f, cell_v in zip(row_f, row_v):
                if cell_f.value is None and cell_v.value is None:
                    continue
                    
                address = cell_f.coordinate
                formula_str = None
                is_array = False
                array_range = None
                precedents = []
                
                if cell_f.data_type == 'f':
                    val_f = cell_f.value
                    if isinstance(val_f, str):
                        formula_str = val_f
                    elif isinstance(val_f, ArrayFormula):
                        is_array = True
                        formula_str = str(getattr(val_f, "text", None) or "")
                        array_range = str(getattr(val_f, "ref", None) or "")
                    else:
                        formula_str = str(getattr(val_f, "value", None) or val_f)
                
                    precedents = extract_precedents_from_formula(formula_str)

                data_type = cell_v.data_type if cell_v.data_type else cell_f.data_type
                
                cells_dict[address] = CellData(
                    address=address,
                    value=cell_v.value,  
                    formula=formula_str,
                    precedents=precedents,
                    is_array_formula=is_array,
                    array_range=array_range,
                    data_type=data_type
                )
        
        # --- NEW: Array Formula Spill Propagation ---
        # Find all cells we just extracted that act as the "parent" of an array
        array_parents = {addr: cell for addr, cell in cells_dict.items() if cell.is_array_formula and cell.array_range}
        
        for parent_addr, parent_cell in array_parents.items():
            if ':' in parent_cell.array_range:
                # Get the boundaries of the array (e.g., 'A1:B3' -> min_col=1, min_row=1, max_col=2, max_row=3)
                min_col, min_row, max_col, max_row = range_boundaries(parent_cell.array_range)
                
                for row in range(min_row, max_row + 1):
                    for col in range(min_col, max_col + 1):
                        col_letter = get_column_letter(col)
                        child_addr = f"{col_letter}{row}"
                        
                        if child_addr == parent_addr:
                            continue # Skip the parent cell itself
                            
                        if child_addr in cells_dict:
                            # Update existing cell
                            cells_dict[child_addr].is_array_formula = True
                            cells_dict[child_addr].parent_array_cell = parent_addr
                        else:
                            # If the spilled cell is completely empty/None in openpyxl, create it
                            cells_dict[child_addr] = CellData(
                                address=child_addr,
                                is_array_formula=True,
                                parent_array_cell=parent_addr
                            )
        # --- END NEW ---

        worksheets_data[sheet_name] = WorksheetData(
            name=sheet_name,
            visibility=visibility,
            cells=cells_dict,
            tables=sheet_tables
        )
        
    return worksheets_data


def extract_vba_modules(path: Path, wb_formulas) -> list[VBAModule]:
    """
    Extract VBA macros from the workbook and return them as VBAModule schemas.
    """
    if VBA_Parser is None:
        print("Warning: oletools not installed. Skipping VBA extraction.")
        return []

    # --- NEW: Build the CodeName -> Sheet Name map ---
    codename_to_tab = {}
    for sheet in wb_formulas.worksheets:
        try:
            code_name = sheet.sheet_properties.codeName
            if code_name:
                codename_to_tab[code_name.lower()] = sheet.title
        except AttributeError:
            continue

    modules = []
    try:
        parser = VBA_Parser(str(path))
        if parser.detect_vba_macros():
            for _, _, filename, content in parser.extract_all_macros():
                if isinstance(content, bytes):
                    content = content.decode("utf-8", errors="ignore")
                else:
                    content = str(content) if content is not None else ""

                # --- NEW: Check if this file maps to a worksheet ---
                safe_filename = str(filename) or "Unknown"
                base_name = safe_filename.split(".")[0].lower() # e.g., 'Sheet1.cls' -> 'sheet1'
                linked_sheet = codename_to_tab.get(base_name)

                modules.append(VBAModule(
                    filename=safe_filename,
                    content=content,
                    linked_worksheet=linked_sheet  # --- NEW ---
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