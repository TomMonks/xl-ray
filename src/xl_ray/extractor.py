from pathlib import Path
from openpyxl import load_workbook

from .schema import WorkbookMetadata

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