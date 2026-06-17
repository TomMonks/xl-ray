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

st.title("🩺 XL-Ray: Excel Model Auditor")
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
        
        # --- NEW: Added the Audit Diagnostics Tab ---
        tab1, tab2, tab3, tab4 = st.tabs([
            "📊 Model Summary", 
            "🛠️ Audit Diagnostics", 
            "🔍 Raw JSON", 
            "🤖 LLM Copilot Prompt"
        ])
        
        with tab1:
            st.subheader("Workbook Metadata")
            cols = st.columns(4)
            cols[0].metric("Worksheets", data.metadata.worksheet_count)
            cols[1].metric("Has Macros?", "Yes" if data.metadata.has_macros else "No")
            cols[2].metric("Has Array Formulas?", "Yes" if data.metadata.has_array_formulas else "No")
            cols[3].metric("Has Hidden Sheets?", "Yes" if data.metadata.has_hidden_sheets else "No")
            
            st.subheader("Worksheets")
            for ws_name, ws in data.worksheets.items():
                st.write(f"- **{ws_name}** ({ws.visibility}): {len(ws.cells)} active cells, {len(ws.tables)} tables")
                
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
                
        with tab3:
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
                
        with tab4:
            st.subheader("Generate Prompt for Enterprise Copilot")
            st.markdown("Copy this prompt and paste it directly into Microsoft Copilot.")
            
            system_instructions = f"""You are an expert health economist auditing an Excel model.
Please review the following extracted model logic for potential errors.

Model Name: {data.metadata.file_name}
Has VBA Macros: {data.metadata.has_macros}
Worksheets: {', '.join(data.worksheets.keys())}
"""
            st.code(system_instructions, language="markdown")