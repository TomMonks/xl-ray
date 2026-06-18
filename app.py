import tempfile
import gzip
from pathlib import Path
import streamlit as st

# Import extraction logic and schemas
from xl_ray.extractor import (
    load_workbooks, 
    extract_metadata, 
    extract_vba_modules, 
    extract_named_ranges,
    extract_worksheets
)
from xl_ray.schema import ExcelModelData

# --- NEW: Import Audit Functions ---
from xl_ray.audit import (
    detect_magic_numbers,
    detect_hidden_logic,
    detect_broken_references,
    detect_inconsistent_columns,
    detect_complex_logic
)

st.set_page_config(page_title="XL-Ray Auditor", layout="wide")

st.title("🩻 XL-Ray: Excel Model Auditor")
st.markdown("Interrogate Excel-based health economic models (Markov/Decision Trees) and generate LLM prompts for auditing.")

# --- NEW: Sidebar Configuration ---
with st.sidebar:
    st.header("⚙️ Configuration")
    st.markdown("Adjust the thresholds for the Deterministic Audit.")
    
    # Sliders for complexity thresholds
    ui_max_depth = st.slider(
        "Max Formula Depth Threshold", 
        min_value=1, max_value=20, value=7, 
        help="Flags formulas whose dependency chain exceeds this depth."
    )
    ui_max_complexity = st.slider(
        "Max Complexity Score Threshold", 
        min_value=5, max_value=100, value=25, 
        help="Flags formulas scoring above this threshold based on nested parentheses and specific functions."
    )

# --- 1. File Upload ---
uploaded_file = st.file_uploader("Upload an Excel Model (.xlsx, .xlsm)", type=["xlsx", "xlsm"])

if uploaded_file is not None:
    # Trigger full re-extraction only if a new file is uploaded
    if "extracted_data" not in st.session_state or st.session_state.get("current_file") != uploaded_file.name:
        
        with st.spinner(f"Interrogating {uploaded_file.name}... this may take a minute."):
            with tempfile.NamedTemporaryFile(delete=False, suffix=Path(uploaded_file.name).suffix) as tmp:
                tmp.write(uploaded_file.getvalue())
                tmp_path = Path(tmp.name)

            try:
                wb_formulas, wb_values = load_workbooks(tmp_path)
                vba_modules = extract_vba_modules(tmp_path)
                named_ranges = extract_named_ranges(wb_formulas)
                worksheets = extract_worksheets(wb_formulas, wb_values)
                
                metadata = extract_metadata(
                    path=Path(uploaded_file.name),
                    wb_formulas=wb_formulas, 
                    vba_modules=vba_modules, 
                    named_ranges=named_ranges
                )
                
                metadata.has_data_tables = any(ws.tables for ws in worksheets.values())
                metadata.has_array_formulas = any(
                    cell.is_array_formula for ws in worksheets.values() for cell in ws.cells.values()
                )

                excel_model = ExcelModelData(
                    metadata=metadata,
                    worksheets=worksheets,
                    named_ranges=named_ranges,
                    vba_modules=vba_modules
                )
                
                # Save base data to session state
                st.session_state.extracted_data = excel_model
                st.session_state.current_file = uploaded_file.name
                
                # --- Run Initial Audits ---
                st.session_state.audit_results = {
                    "broken_refs": detect_broken_references(excel_model),
                    "magic_numbers": detect_magic_numbers(excel_model),
                    "inconsistent_cols": detect_inconsistent_columns(excel_model),
                    "hidden_logic": detect_hidden_logic(excel_model),
                    # Pass the UI slider values to the initial run
                    "complex_logic": detect_complex_logic(
                        excel_model, 
                        max_depth_threshold=ui_max_depth, 
                        complexity_score_threshold=ui_max_complexity
                    )
                }

            except Exception as e:
                st.error(f"Error interrogating file: {e}")
            finally:
                if tmp_path.exists():
                    tmp_path.unlink()
    
    # --- NEW: Dynamic Complexity Recalculation ---
    # If the user moves the sliders, we only recalculate the complexity audit 
    # instead of re-parsing the whole Excel file.
    elif "extracted_data" in st.session_state:
        st.session_state.audit_results["complex_logic"] = detect_complex_logic(
            st.session_state.extracted_data, 
            max_depth_threshold=ui_max_depth, 
            complexity_score_threshold=ui_max_complexity
        )

    # --- 2. Display Results ---
    if "extracted_data" in st.session_state:
        data: ExcelModelData = st.session_state.extracted_data
        audits = st.session_state.audit_results
        
        # --- Added a new tab for VBA Modules ---
        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "📊 Model Summary",
            "🛠️ Audit Diagnostics",
            "📜 VBA Modules",
            "🔍 Raw JSON",
            "🤖 LLM Copilot Prompt"
        ])
        
        with tab1:
            import pandas as pd
            
            st.subheader("Workbook Metadata")
            st.markdown(f"**File Segmented:** `{uploaded_file.name}`")
            
            cols = st.columns(4)
            cols[0].metric("Worksheets", data.metadata.worksheet_count)
            cols[1].metric("Has Macros?", "Yes" if data.metadata.has_macros else "No")
            cols[2].metric("Has Array Formulas?", "Yes" if data.metadata.has_array_formulas else "No")
            cols[3].metric("Has Hidden Sheets?", "Yes" if data.metadata.has_hidden_sheets else "No")
            
            st.divider()
            st.subheader("Worksheets Overview")
            
            # Restructure worksheet dictionary into a DataFrame
            ws_data = []
            for ws_name, ws in data.worksheets.items():
                visibility_status = "👁️ Visible" if str(ws.visibility).lower() == "visible" else "👻 Hidden"
                ws_data.append({
                    "Sheet Name": ws_name,
                    "Visibility": visibility_status,
                    "Active Cells": len(ws.cells),
                    "Data Tables": len(ws.tables)
                })
            
            if ws_data:
                df_ws = pd.DataFrame(ws_data)
                max_cells = int(df_ws["Active Cells"].max()) if not df_ws.empty else 100
                
                # Render an upgraded dataframe with visual bar charts
                st.dataframe(
                    df_ws,
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "Sheet Name": st.column_config.TextColumn("Sheet Name", width="medium"),
                        "Visibility": st.column_config.TextColumn("Status", width="small"),
                        "Active Cells": st.column_config.ProgressColumn(
                            "Active Cells",
                            help="Total number of parsed cells in this sheet",
                            format="%d",
                            min_value=0,
                            max_value=max_cells
                        ),
                        "Data Tables": st.column_config.NumberColumn("Data Tables", format="%d")
                    }
                )
            else:
                st.info("No worksheets detected in this model.")
                
        # --- NEW: Audit Diagnostics UI ---
        with tab2:
            st.subheader("Deterministic Audit Results")
            st.markdown("These checks run instantly in Python to triage the model before using AI.")

            def display_audit_section(title, flag_data, icon, empty_msg):
                if flag_data:
                    with st.expander(f"{icon} {title} ({len(flag_data)} issues found)"):
                        st.dataframe(flag_data, use_container_width=True)
                else:
                    st.success(f"{icon} {title}: {empty_msg}")

            display_audit_section("Broken References", audits["broken_refs"], "🚨", "No broken references found.")
            display_audit_section("Hardcoded Magic Numbers", audits["magic_numbers"], "🪄", "No magic numbers found.")
            display_audit_section("Inconsistent Column Formulas", audits["inconsistent_cols"], "📉", "No column inconsistencies found.")
            display_audit_section("Hidden Logic Dependencies", audits["hidden_logic"], "👻", "No hidden sheet dependencies found.")
            display_audit_section("High Complexity / Deep Logic", audits["complex_logic"], "🍝", "No overly complex logic found.")

            st.markdown("---")
            st.subheader("🕵️‍♂️ Logic Dependency Tracer")
            st.markdown("Enter a specific cell address (e.g., from the 'High Complexity' table above) to see exactly how its value is calculated.")

            # Interactive form for tracing
            col1, col2, col3 = st.columns([2, 1, 1])
            with col1:
                trace_sheet = st.selectbox("Select Worksheet", list(data.worksheets.keys()))
            with col2:
                trace_cell = st.text_input("Cell Address", placeholder="e.g., C14")
            with col3:
                st.markdown("<br>", unsafe_allow_html=True) # alignment hack
                run_trace = st.button("Trace Logic")

            if run_trace and trace_cell:
                from xl_ray.audit import trace_cell_logic
                import json
                
                trace_cell = trace_cell.upper().strip()
                trace_result = trace_cell_logic(data, trace_sheet, trace_cell)
                
                if trace_result:
                    st.markdown("#### 🌳 Interactive Logic Tree")
                    
                    # --- Symbol Key / Legend ---
                    st.markdown(
                        """
                        <div style="background-color: #262730; color: #fafafa; padding: 10px; border-radius: 5px; border: 1px solid #444; margin-bottom: 15px; font-size: 13px;">
                            <strong>Symbol Key:</strong><br>
                            🎯 <b>Target</b>: Starting cell &nbsp;&nbsp;|&nbsp;&nbsp; 
                            📐 <b>Formula</b>: Calculated cell &nbsp;&nbsp;|&nbsp;&nbsp; 
                            🔹 <b>Constant</b>: Hardcoded value &nbsp;&nbsp;|&nbsp;&nbsp; 
                            📦 <b>Range</b>: Cell block (skipped)<br>
                            ⏭️ <b>Duplicate</b>: Already expanded elsewhere &nbsp;&nbsp;|&nbsp;&nbsp; 
                            🔄 <b>Circular</b>: Loop detected &nbsp;&nbsp;|&nbsp;&nbsp; 
                            ❌ <b>Missing</b>: Broken reference &nbsp;&nbsp;|&nbsp;&nbsp; 
                            ⚠️ <b>Limit</b>: Max depth reached
                        </div>
                        """, 
                        unsafe_allow_html=True
                    )
                    
                    # Recursive function to build HTML details/summary tree from the JSON structure
                    def build_html_tree(node):
                        icons = {
                            "formula": "📐", "constant": "🔹", "duplicate": "⏭️", 
                            "circular": "🔄", "missing": "❌", "limit": "⚠️", "range": "📦"
                        }
                        
                        node_type = node.get("type", "unknown")
                        icon = "🎯" if node.get("is_target") else icons.get(node_type, "▪️")
                        address = node.get("address", "")
                        via = f" <span style='color:#888;font-size:0.9em'>(via {node.get('via')})</span>" if node.get("via") else ""
                        
                        if node_type == "formula":
                            value_html = f"<code style='background-color:#333; color:#ff4b4b; padding:2px 4px; border-radius:4px;'>{node.get('value','')}</code>"
                        elif node_type == "constant":
                            value_html = f"<b>{node.get('value','')}</b>"
                        else:
                            value_html = f"<span style='color:#a0a0a0; font-style:italic;'>{node.get('value','')}</span>"
                            
                        header = f"{icon} <b>{address}</b>{via}: {value_html}"
                        children = node.get("children", [])
                        
                        # Leaf node (no children or stopped tracing)
                        if not children or node_type in ["duplicate", "circular", "missing", "limit"]:
                            if node_type == "duplicate":
                                return f"<div style='padding: 4px 0; color:#a0a0a0;'>{icon} <i>{address}</i>{via} - <span style='font-size:0.9em'>(Already expanded above)</span></div>"
                            return f"<div style='padding: 4px 0;'>{header}</div>"
                            
                        # Parent node with expandable details
                        children_html = "".join([build_html_tree(child) for child in children])
                        
                        # Open the root node and the first level by default for quick visibility
                        open_attr = "open" if node.get("depth", 0) < 2 else ""
                        
                        # Return as a single continuous string to prevent Streamlit markdown parsing errors
                        return f"<details {open_attr} style='margin-top: 4px;'><summary style='cursor: pointer; padding: 4px 0; outline: none;'>{header}</summary><div style='border-left: 1px dashed #555; padding-left: 20px; margin-left: 7px; margin-top: 4px; margin-bottom: 4px;'>{children_html}</div></details>"


                    interactive_tree_html = build_html_tree(trace_result)
                    
                    # Display the HTML tree
                    st.markdown(
                        f"""
                        <div style='background-color:#1e1e1e; color:#d4d4d4; padding:20px; border-radius:5px; 
                                    font-family:monospace; font-size:14px; line-height:1.6; 
                                    overflow-x:auto; max-height:600px; overflow-y:auto;'>
                            {interactive_tree_html}
                        </div>
                        """, 
                        unsafe_allow_html=True
                    )
                    
                    # Add a developer toggle to view the raw JSON tree directly
                    with st.expander("View Raw JSON Payload"):
                        st.json(trace_result)
                    
                    # Pass structured JSON text to the LLM prompt state instead of flat text
                    st.session_state.last_trace = {
                        "target": f"{trace_sheet}!{trace_cell}",
                        "trace_text": json.dumps(trace_result, indent=2) 
                    }
                else:
                    st.warning(f"Could not trace {trace_sheet}!{trace_cell}. Check if the address is valid.")

        # --- NEW: VBA Modules Tab ---
        with tab3:
            st.subheader("📜 VBA Macros & Modules")
            st.markdown("Review the raw VBA code embedded in the workbook.")
            
            if data.vba_modules:
                # Handle whether data.vba_modules is a list of objects or dictionaries
                if isinstance(data.vba_modules, dict):
                    # Just in case it's parsed as a dict map
                    modules_items = [(k, v.get('content', v) if isinstance(v, dict) else v) for k, v in data.vba_modules.items()]
                else:
                    # It's a list of schema objects (like Pydantic models)
                    modules_items = []
                    for i, m in enumerate(data.vba_modules):
                        # Handle both dictionary and object access safely
                        name = m.get('filename', f"Module {i}") if isinstance(m, dict) else getattr(m, 'filename', f"Module {i}")
                        content = m.get('content', "") if isinstance(m, dict) else getattr(m, 'content', "")
                        modules_items.append((name, content))

                # Iterate through the modules and create an expander for each
                for mod_name, mod_code in modules_items:
                    display_code = mod_code if mod_code else "' No code found in this module."
                    
                    with st.expander(f"📝 {mod_name}"):
                        # Streamlit natively supports VBA syntax highlighting
                        st.code(display_code, language="vba")
            else:
                st.info("No VBA macros or modules were found in this workbook.")
                
        with tab4:
            st.markdown("Download the fully extracted structured JSON data.")
            json_bytes = gzip.compress(data.model_dump_json(indent=2).encode('utf-8'))
            st.download_button(
                label="Download Compressed JSON (.json.gz)",
                data=json_bytes,
                file_name=f"{data.metadata.file_name}_xlray.json.gz",
                mime="application/gzip"
            )
            
            with st.expander("View JSON Snippet"):
                st.json(data.metadata.model_dump())
                
        with tab5:
            st.subheader("Generate Prompt for Enterprise Copilot")
            st.markdown("Copy this prompt and paste it directly into Microsoft Copilot.")
            
            system_instructions = f"""You are an expert health economist auditing an Excel model.
Please review the following extracted model logic for potential errors.

Model Name: {data.metadata.file_name}
Has VBA Macros: {data.metadata.has_macros}
Worksheets: {', '.join(data.worksheets.keys())}
"""
            st.code(system_instructions, language="markdown")