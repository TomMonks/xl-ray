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


def detect_broken_references(model_data: ExcelModelData) -> List[Dict[str, Any]]:
    """
    Identifies cells that contain explicit #REF! errors or rely on 
    precedents that evaluate to Excel error states or are completely empty.
    Handles named ranges and ignores block ranges.
    """
    flagged_cells = []
    error_states = {"#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!", "#REF!", "#VALUE!"}
    
    # Helper to check if a string looks like a single coordinate (A1, Z99)
    # This prevents us from trying to look up block ranges like 'A1:A10'
    coord_pattern = re.compile(r'^[A-Za-z]+\d+$')

    for ws_name, ws_data in model_data.worksheets.items():
        for cell_address, cell in ws_data.cells.items():
            issues = []
            
            if cell.formula and "#REF!" in cell.formula:
                issues.append("Formula contains explicit #REF!")
                
            if cell.precedents:
                for precedent in cell.precedents:
                    target_sheet = ws_name
                    target_cell = precedent
                    
                    # 1. Is it a Named Range?
                    is_named_range = False
                    for nr in model_data.named_ranges:
                        if nr.name == precedent:
                            is_named_range = True
                            # Extract target sheet and cell from the refers_to string
                            # e.g., 'Sheet1'!$A$1 -> Sheet1, A1
                            ref_clean = nr.refers_to.replace('$', '')
                            if "!" in ref_clean:
                                target_sheet = ref_clean.split("!")[0].replace("'", "").replace("=", "")
                                target_cell = ref_clean.split("!")[1]
                            else:
                                target_cell = ref_clean.replace("=", "")
                            break
                            
                    # 2. Handle standard cross-sheet references if it wasn't a named range
                    if not is_named_range and "!" in precedent:
                        parts = precedent.split("!")
                        target_sheet = parts[0].replace("'", "")
                        target_cell = parts[1]

                    # 3. Skip Block Ranges (e.g., D13:D14). We only check single target cells.
                    if ":" in target_cell or not coord_pattern.match(target_cell):
                        continue

                    # 4. Check if the target sheet exists
                    if target_sheet not in model_data.worksheets:
                        issues.append(f"Broken sheet link: {target_sheet}")
                        continue
                        
                    target_ws = model_data.worksheets[target_sheet]
                    
                    # 5. Check if the target cell exists and what its value is
                    if target_cell not in target_ws.cells:
                        issues.append(f"Precedent '{precedent}' (resolved to {target_cell}) is blank/missing")
                    else:
                        target_val = target_ws.cells[target_cell].value
                        if str(target_val) in error_states:
                            issues.append(f"Precedent '{precedent}' evaluates to error: {target_val}")

            if issues:
                flagged_cells.append({
                    "sheet": ws_name,
                    "cell": cell_address,
                    "formula": cell.formula,
                    "value": cell.value,
                    "issues_found": issues
                })
                
    return flagged_cells