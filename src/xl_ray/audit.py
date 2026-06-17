import re
from typing import List, Dict, Any
from xl_ray.schema import ExcelModelData

# ---------------------------------------------------------------------------
# Magic Number Detector
# ---------------------------------------------------------------------------

def detect_magic_numbers(model_data: ExcelModelData) -> List[Dict[str, Any]]:
    """
    Scans all worksheets for formulas containing hardcoded 'magic numbers'.
    Ignores common structural numbers like 0, 1, and -1.
    """
    flagged_cells = []
    
    # Regex to find mathematical operators followed by digits
    # Matches things like '+ 5', '* 0.05', '/ 365', '> 10'
    # The negative lookbehind (?<![A-Za-z]) ensures we don't match cell row numbers like 'A15'
    pattern = re.compile(r'[\+\-\*/^<>=]\s*(?<![A-Za-z])(\d+\.?\d*)')
    
    # Numbers we typically want to ignore as they are often used for basic logic
    ignore_list = {"0", "1", "-1", "0.0"}

    for ws_name, ws_data in model_data.worksheets.items():
        for cell_address, cell in ws_data.cells.items():
            if not cell.formula:
                continue
                
            matches = pattern.findall(cell.formula)
            
            # Filter out ignored numbers
            suspicious_numbers = [m for m in matches if m not in ignore_list]
            
            if suspicious_numbers:
                flagged_cells.append({
                    "sheet": ws_name,
                    "cell": cell_address,
                    "formula": cell.formula,
                    "magic_numbers_found": suspicious_numbers
                })
                
    return flagged_cells

# ---------------------------------------------------------------------------
# Hidden Logic Flag
# ---------------------------------------------------------------------------

def detect_hidden_logic(model_data: ExcelModelData) -> List[Dict[str, Any]]:
    """
    Identifies cells on Visible sheets that reference cells on Hidden or VeryHidden sheets.
    """
    flagged_cells = []
    
    # Map out which sheets are hidden
    hidden_sheets = {
        name for name, ws in model_data.worksheets.items() 
        if ws.visibility in ("Hidden", "VeryHidden")
    }
    
    # If no sheets are hidden, we can exit early
    if not hidden_sheets:
        return flagged_cells

    for ws_name, ws_data in model_data.worksheets.items():
        # We only care about visible sheets pulling from hidden ones
        if ws_data.visibility in ("Hidden", "VeryHidden"):
            continue
            
        for cell_address, cell in ws_data.cells.items():
            if not cell.precedents:
                continue
                
            hidden_refs = []
            for precedent in cell.precedents:
                # Precedents look like "'Sheet Name'!A1" or "SheetName!A1"
                if "!" in precedent:
                    target_sheet = precedent.split("!")[0].replace("'", "")
                    if target_sheet in hidden_sheets:
                        hidden_refs.append(precedent)
                        
            if hidden_refs:
                flagged_cells.append({
                    "sheet": ws_name,
                    "cell": cell_address,
                    "formula": cell.formula,
                    "hidden_references": hidden_refs
                })
                
    return flagged_cells