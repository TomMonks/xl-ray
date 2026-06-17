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


def detect_inconsistent_columns(model_data: ExcelModelData) -> List[Dict[str, Any]]:
    """
    Scans contiguous columns of formulas to find cells that break the established pattern.
    Useful for finding manual overrides ('fudge factors') in Markov traces or data tables.
    """
    flagged_cells = []
    
    # Helper to strip row numbers out of references (e.g., '=B12 * 2' -> '=B_ROW * 2')
    # This allows us to compare relative formula structures down a column.
    def abstract_formula(f_str):
        if not f_str: 
            return ""
        # Match column letters followed by row numbers and replace the number
        return re.compile(r'([A-Za-z]+)\d+').sub(r'\1_ROW', f_str)

    # Helper to parse A1 coordinate into column letter and row number
    coord_pattern = re.compile(r'^([A-Za-z]+)(\d+)$')

    for ws_name, ws_data in model_data.worksheets.items():
        # 1. Group formula cells by column
        cols = {}
        for addr, cell in ws_data.cells.items():
            if not cell.formula: 
                continue
            
            match = coord_pattern.match(addr)
            if match:
                c_let, r_num = match.groups()
                r_num = int(r_num)
                if c_let not in cols: 
                    cols[c_let] = []
                cols[c_let].append((r_num, cell))
                
        # 2. Check each column for inconsistencies
        for c_let, row_cells in cols.items():
            # Sort by row number
            row_cells.sort(key=lambda x: x[0])
            
            # We need at least 3 contiguous cells to establish a broken pattern
            # (e.g., Row 1 matches Row 3, but Row 2 is different)
            if len(row_cells) < 3: 
                continue
            
            for i in range(1, len(row_cells) - 1):
                prev_r, prev_cell = row_cells[i-1]
                curr_r, curr_cell = row_cells[i]
                next_r, next_cell = row_cells[i+1]
                
                # Ensure the three cells are perfectly contiguous (no blank rows between them)
                if curr_r != prev_r + 1 or next_r != curr_r + 1:
                    continue
                    
                prev_abs = abstract_formula(prev_cell.formula)
                curr_abs = abstract_formula(curr_cell.formula)
                next_abs = abstract_formula(next_cell.formula)
                
                # Detect an override: the cell before and after match, but the current one doesn't
                if prev_abs == next_abs and curr_abs != prev_abs:
                    flagged_cells.append({
                        "sheet": ws_name,
                        "cell": curr_cell.address,
                        "formula": curr_cell.formula,
                        "expected_pattern": prev_abs.replace("_ROW", "[row]"),
                        "actual_pattern": curr_abs.replace("_ROW", "[row]")
                    })
                    
    return flagged_cells

def detect_complex_logic(
    model_data: ExcelModelData, 
    max_depth_threshold: int = 7, 
    complexity_score_threshold: int = 25
) -> List[Dict[str, Any]]:
    """
    Identifies cells with highly complex formulas or deep dependency chains.
    
    This function traverses the parsed Excel model to find formulas that exceed 
    a calculated structural complexity score or a specified dependency chain 
    depth. It deduplicates identical formula patterns across rows to prevent 
    flooding the output with repeating column logic.

    Parameters
    ----------
    model_data : ExcelModelData
        The parsed representation of the Excel workbook containing worksheets, 
        cells, and parsed precedents.
    max_depth_threshold : int, optional
        The maximum number of levels a dependency chain can go before it is 
        flagged as 'Deep Logic'. For example, if A1 relies on A2, which relies 
        on A3, the depth is 2. The default is 7.
    complexity_score_threshold : int, optional
        The maximum structural complexity score allowed before a formula is 
        flagged. The score is calculated based on parenthesis nesting, the total 
        number of functions, and the presence of specific high-risk functions 
        (e.g., VLOOKUP, OFFSET). The default is 25.

    Returns
    -------
    List[Dict[str, Any]]
        A list of dictionaries representing the flagged cells. Each dictionary 
        contains 'sheet', 'cell', 'formula', 'chain_depth', 'complexity_score', 
        and 'flag_reason'. The list is sorted in descending order by complexity 
        score and then by chain depth.
    """
    
    flagged_cells = []
    depth_cache = {}
    seen_patterns = set()

    # Helper to strip row numbers out of references for deduplication
    def abstract_formula(f_str):
        if not f_str: return ""
        return re.compile(r'([A-Za-z]+)\d+').sub(r'\1_ROW', f_str)

    def calculate_depth(sheet_name: str, cell_address: str, current_path: set) -> int:
        cache_key = f"{sheet_name}!{cell_address}"
        if cache_key in current_path:
            return 0 
        if cache_key in depth_cache:
            return depth_cache[cache_key]
            
        ws = model_data.worksheets.get(sheet_name)
        if not ws or cell_address not in ws.cells:
            return 0
            
        cell = ws.cells[cell_address]
        if not cell.precedents:
            depth_cache[cache_key] = 0
            return 0
            
        max_child_depth = 0
        current_path.add(cache_key)
        
        for precedent in cell.precedents:
            target_sheet = sheet_name
            target_cell = precedent
            if "!" in precedent:
                parts = precedent.split("!")
                target_sheet = parts[0].replace("'", "")
                target_cell = parts[1]
                
            if ":" in target_cell:
                continue
                
            depth = calculate_depth(target_sheet, target_cell, current_path)
            if depth > max_child_depth:
                max_child_depth = depth
                
        current_path.remove(cache_key)
        final_depth = max_child_depth + 1
        depth_cache[cache_key] = final_depth
        return final_depth

    def calculate_complexity_score(formula: str) -> int:
        if not formula: return 0
        score = 0
        
        current_nesting = 0
        max_nesting = 0
        for char in formula:
            if char == '(': current_nesting += 1
            elif char == ')': current_nesting -= 1
            if current_nesting > max_nesting: max_nesting = current_nesting
        score += max_nesting * 2 
        
        functions = re.findall(r'([A-Za-z_]+)\(', formula)
        score += len(functions)
        
        complex_funcs = {"VLOOKUP", "INDEX", "MATCH", "OFFSET", "INDIRECT", "IFERROR", "SUMIFS"}
        for func in functions:
            if func.upper() in complex_funcs:
                score += 2
                
        return score

    # Run the checks
    for ws_name, ws_data in model_data.worksheets.items():
        for cell_address, cell in ws_data.cells.items():
            if not cell.formula:
                continue
                
            depth = calculate_depth(ws_name, cell_address, set())
            complexity = calculate_complexity_score(cell.formula)
            
            if depth > max_depth_threshold or complexity > complexity_score_threshold:
                # Check for deduplication
                abs_f = abstract_formula(cell.formula)
                if abs_f in seen_patterns:
                    continue
                seen_patterns.add(abs_f)
                
                flagged_cells.append({
                    "sheet": ws_name,
                    "cell": cell_address,
                    "formula": cell.formula,
                    "chain_depth": depth,
                    "complexity_score": complexity,
                    "flag_reason": "Deep Logic Chain" if depth > max_depth_threshold else "Highly Complex Formula"
                })
                
    flagged_cells.sort(key=lambda x: (x["complexity_score"], x["chain_depth"]), reverse=True)
    return flagged_cells