import tempfile
import gzip
from pathlib import Path
import streamlit as st

# Import your extraction logic and schemas
from xl_ray.extractor import (
    load_workbooks, 
    extract_metadata, 
    extract_vba_modules, 
    extract_named_ranges,
    extract_worksheets
)
from xl_ray.schema import ExcelModelData

st.set_page_config(page_title="XL-Ray Auditor", layout="wide")

st.title("🩺 XL-Ray: Excel Model Auditor")
st.markdown("Interrogate Excel-based health economic models (Markov/Decision Trees) and generate LLM prompts for auditing.")

# --- 1. File Upload ---
uploaded_file = st.file_uploader("Upload an Excel Model (.xlsx, .xlsm)", type=["xlsx", "xlsm"])

if uploaded_file is not None:
    # Use session state to avoid re-extracting if the user interacts with the app
    if "extracted_data" not in st.session_state or st.session_state.current_file != uploaded_file.name:
        
        with st.spinner(f"Interrogating {uploaded_file.name}... this may take a minute."):
            # Save uploaded file to a temporary file so oletools/openpyxl can read it from disk
            with tempfile.NamedTemporaryFile(delete=False, suffix=Path(uploaded_file.name).suffix) as tmp:
                tmp.write(uploaded_file.getvalue())
                tmp_path = Path(tmp.name)

            try:
                # Run the pipeline
                wb_formulas, wb_values = load_workbooks(tmp_path)
                vba_modules = extract_vba_modules(tmp_path)
                named_ranges = extract_named_ranges(wb_formulas)
                worksheets = extract_worksheets(wb_formulas, wb_values)
                
                metadata = extract_metadata(
                    path=Path(uploaded_file.name), # Use original name for metadata
                    wb_formulas=wb_formulas, 
                    vba_modules=vba_modules, 
                    named_ranges=named_ranges
                )
                
                # Update flags
                metadata.has_data_tables = any(ws.tables for ws in worksheets.values())
                metadata.has_array_formulas = any(
                    cell.is_array_formula for ws in worksheets.values() for cell in ws.cells.values()
                )

                # Build Root Model
                excel_model = ExcelModelData(
                    metadata=metadata,
                    worksheets=worksheets,
                    named_ranges=named_ranges,
                    vba_modules=vba_modules
                )
                
                # Save to session state
                st.session_state.extracted_data = excel_model
                st.session_state.current_file = uploaded_file.name
                
            except Exception as e:
                st.error(f"Error interrogating file: {e}")
            finally:
                # Clean up the temp file
                if tmp_path.exists():
                    tmp_path.unlink()

    # --- 2. Display Results ---
    if "extracted_data" in st.session_state:
        data: ExcelModelData = st.session_state.extracted_data
        
        # Create tabs for different views
        tab1, tab2, tab3 = st.tabs(["📊 Model Summary", "🔍 Raw JSON", "🤖 LLM Copilot Prompt"])
        
        with tab1:
            st.subheader("Workbook Metadata")
            # Display high-level flags as a grid of metrics
            cols = st.columns(4)
            cols[0].metric("Worksheets", data.metadata.worksheet_count)
            cols[1].metric("Has Macros?", "Yes" if data.metadata.has_macros else "No")
            cols[2].metric("Has Array Formulas?", "Yes" if data.metadata.has_array_formulas else "No")
            cols[3].metric("Has Hidden Sheets?", "Yes" if data.metadata.has_hidden_sheets else "No")
            
            st.subheader("Worksheets")
            for ws_name, ws in data.worksheets.items():
                st.write(f"- **{ws_name}** ({ws.visibility}): {len(ws.cells)} active cells, {len(ws.tables)} tables")
                
        with tab2:
            st.markdown("Download the fully extracted structured JSON data.")
            # Streamlit download button for the gzipped JSON
            json_bytes = gzip.compress(data.model_dump_json(indent=2).encode('utf-8'))
            st.download_button(
                label="Download Compressed JSON (.json.gz)",
                data=json_bytes,
                file_name=f"{data.metadata.file_name}_xlray.json.gz",
                mime="application/gzip"
            )
            
            with st.expander("View JSON Snippet"):
                st.json(data.metadata.model_dump())
                
        with tab3:
            st.subheader("Generate Prompt for Enterprise Copilot")
            st.markdown("Copy this prompt and paste it directly into Microsoft Copilot.")
            
            # This is where you construct the actual prompt. 
            # We can build a dedicated prompt builder module later, but here is a basic example.
            system_instructions = f"""You are an expert health economist auditing an Excel model.
Please review the following extracted model logic for potential errors, focusing on array formulas and hidden sheets.

Model Name: {data.metadata.file_name}
Has VBA Macros: {data.metadata.has_macros}
Worksheets: {', '.join(data.worksheets.keys())}
"""
            # st.code automatically includes a "Copy" button in the top right corner!
            st.code(system_instructions, language="markdown")