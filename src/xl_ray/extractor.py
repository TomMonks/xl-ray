from pathlib import Path
from openpyxl import load_workbook

from .schema import WorkbookMetadata, VBAModule, NamedRange, ExcelTable

from oletools.olevba import VBA_Parser

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