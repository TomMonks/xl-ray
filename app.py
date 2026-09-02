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

from xl_ray import __version__

# --- NEW: Import Audit Functions ---
from xl_ray.audit import (
    detect_magic_numbers,
    detect_hidden_logic,
    detect_broken_references,
    detect_inconsistent_columns,
    detect_complex_logic,
    detect_array_formulas
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

    # --- NEW: Version info at the bottom of the sidebar ---
    st.divider()
    st.caption(f"Powered by `xl-ray` v{__version__}")

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
                vba_modules = extract_vba_modules(tmp_path, wb_formulas)
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

                # --- NEW: v0.2.0 Extract Terminal Cells for UI ---
                terminal_cells_list = []
                for ws_name, ws_data in excel_model.worksheets.items():
                    for cell_addr, cell in ws_data.cells.items():
                        if getattr(cell, 'is_terminal', False):
                            terminal_cells_list.append({
                                "sheet": ws_name,
                                "cell": cell_addr,
                                "formula": cell.formula,
                                "value": str(cell.value)
                            })
                

                # --- Run Initial Audits ---
                st.session_state.audit_results = {
                    "terminal_cells": terminal_cells_list, # TM added v0.2.0
                    "broken_refs": detect_broken_references(excel_model),
                    "magic_numbers": detect_magic_numbers(excel_model),
                    "inconsistent_cols": detect_inconsistent_columns(excel_model),
                    "hidden_logic": detect_hidden_logic(excel_model),
                    # Pass the UI slider values to the initial run
                    "complex_logic": detect_complex_logic(
                        excel_model, 
                        max_depth_threshold=ui_max_depth, 
                        complexity_score_threshold=ui_max_complexity
                    ),
                    "array_formulas": detect_array_formulas(excel_model),
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
        tab1, tab2, tab3, tab4 = st.tabs([
            "📊 Model Summary",
            "🛠️ Audit Diagnostics",
            "📜 VBA Modules",
            "🔍 Raw JSON",
            #"🤖 LLM Copilot Prompt"
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

            display_audit_section(
                    "Terminal Cells (Model Outputs)", 
                    audits.get("terminal_cells", []), 
                    "🛑", 
                    "No terminal cells found."
                )
            display_audit_section("Broken References", audits["broken_refs"], "🚨", "No broken references found.")
            display_audit_section("Hardcoded Magic Numbers", audits["magic_numbers"], "🪄", "No magic numbers found.")
            display_audit_section("Inconsistent Column Formulas", audits["inconsistent_cols"], "📉", "No column inconsistencies found.")
            display_audit_section("Array Formulas", audits.get('array_formulas'), "📊", "No array formulas found.")
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
                for mod in data.vba_modules:
                    display_title = f"📝 {mod.filename}"
                    
                    # Use the new schema property directly!
                    if mod.linked_worksheet:
                        display_title += f" — 🏷️ Worksheet: '{mod.linked_worksheet}'"
                    elif mod.filename.lower().endswith(".cls") and "thisworkbook" not in mod.filename.lower():
                        display_title += " — 🏷️ Worksheet module"

                    if mod.content:
                        line_count = len(mod.content.splitlines())
                        display_title += f" | {line_count} lines"

                    display_code = mod.content if mod.content else "' No code found in this module."
                    
                    with st.expander(display_title):
                        st.code(display_code, language="vba")
            else:
                st.info("No VBA macros or modules were found in this workbook.")
                
        # --- NEW: Raw JSON Export Tab with ZIP Compression ---
        with tab4:
            st.subheader("🔍 Extracted JSON Data")
            st.markdown("Download the fully extracted structural data. This file is required for the 'Full Workbook' LLM prompt.")
            
            # 1. Get the JSON string
            json_data = data.model_dump_json(indent=2)
            
            # 2. Create the raw JSON download button for small models
            st.download_button(
                label="📄 Download Raw JSON (Uncompressed)",
                data=json_data,
                file_name=f"{st.session_state.current_file}_xlray.json",
                mime="application/json"
            )
            
            st.markdown("---")
            st.markdown("### LLM Export (Recommended)")
            st.markdown("For large models, Copilot/ChatGPT may reject 20MB+ files. Use this ZIP format for LLM uploads.")
            
            # 3. Create the ZIP file in memory
            import zipfile
            import io
            
            zip_buffer = io.BytesIO()
            # Compress the JSON string into the zip buffer
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, compresslevel=9) as zf:
                zip_filename = f"{st.session_state.current_file}_xlray.json"
                zf.writestr(zip_filename, json_data)
                
            # 4. Provide the ZIP download button
            st.download_button(
                label="🗜️ Download ZIP (For LLM Upload)",
                data=zip_buffer.getvalue(),
                file_name=f"{st.session_state.current_file}_xlray.zip",
                mime="application/zip",
                type="primary" # Highlight this button
            )
                
    #     # --- NEW: LLM Copilot Prompt Builder ---
    #     with tab5: # (or whichever tab number it is)
    #         st.subheader("🤖 LLM Copilot Prompt Builder")
    #         st.markdown("Generate targeted prompts to copy/paste into ChatGPT, Claude, or a local LLM.")
            
    #         # 1. Select the type of analysis
    #         prompt_type = st.selectbox(
    #             "Select Analysis Type",
    #             [
    #                 "Full Workbook Evaluation",
    #                 "VBA: Explain Functionality & Assess Risks",
    #                 "Trace Logic: Explain specific cell calculation",
    #                 "Audit Triage: Prioritize deterministic flags",
    #                 "Model Architecture Review"
    #             ]
    #         )
            

    #         prompt_text = ""
    #         system_role = "You are an expert Health Economist and advanced Excel Model Auditor."
            
    #         if "Full Workbook" in prompt_type:
    #             schema_definition = """
    #     class CellData(BaseModel):
    #         address: str
    #         value: Any
    #         formula: Optional[str]
    #         precedents: list[str]
    #         dependents: list[str]
    #         is_array_formula: bool
    #         array_range: Optional[str]
    #         parent_array_cell: Optional[str]
    #         data_type: Optional[str]

    #     class NamedRange(BaseModel):
    #         name: str
    #         refers_to: str
    #         scope: str

    #     class ExcelTable(BaseModel):
    #         name: str
    #         range_address: str
    #         columns: list[str]

    #     class WorksheetData(BaseModel):
    #         name: str
    #         visibility: str
    #         code_name: Optional[str]
    #         cells: dict[str, CellData]
    #         tables: list[ExcelTable]

    #     class VBAModule(BaseModel):
    #         filename: str
    #         content: str
    #         linked_worksheet: Optional[str]

    #     class WorkbookMetadata(BaseModel):
    #         file_name: str
    #         has_macros: bool
    #         has_array_formulas: bool
    #         has_data_tables: bool
    #         has_named_ranges: bool
    #         has_hidden_sheets: bool
    #         has_external_links: bool
    #         worksheet_count: int

    #     class ExcelModelData(BaseModel):
    #         metadata: WorkbookMetadata
    #         worksheets: dict[str, WorksheetData]
    #         named_ranges: list[NamedRange]
    #         vba_modules: list[VBAModule]
    #     """

    #             prompt_text = f"""{system_role}

    # I have attached a large JSON file representing the full structural extraction of a health economic Excel model. 

    # ### Task Instructions
    # Please use your Python / Advanced Data Analysis capabilities to load this `.json` file into memory. Do not attempt to read the entire file into your text context; write Python scripts to parse it and programmatically interrogate its structure.

    # **Primary Goal:** I need you to evaluate the OVERALL architecture of this model. Individual cell errors are easy to find, but I need you to identify systemic mistakes, architectural flaws, and macro-level risks that are hard to spot manually.

    # ### Pydantic Schema Definition
    # The JSON is serialized directly from these Pydantic models. Use this to write accurate Python parsing logic:
    # ```python
    # {schema_definition}
    # ```

    # ### Systemic Checks to Execute via Python:
    # 1. **Separation of Concerns:** Analyze cross-sheet dependencies (`precedents`). In health economics, Inputs, Engine (Markov/Decision Tree), and Results should be isolated. Flag "spaghetti logic" where calculation sheets pull directly from other calculation sheets rather than a centralized parameter dashboard.
    # 2. **Hidden Vulnerabilities:** Map out any dependencies where visible calculation or result cells rely on parameters hidden in "VeryHidden" sheets or obscured by overlapping Named Ranges.
    # 3. **Trace Inconsistencies:** In Markov trace sheets, formulas down a single column should be strictly identical relative to their row. Programmatically scan for "fudge factors" where a formula breaks the pattern mid-column.
    # 4. **Hardcoded Risks:** Scan for formulas combining dynamic variables with hardcoded "magic numbers" (e.g., `=A1 * 0.035`), ignoring structural numbers (0, 1, -1).

    # ### REQUIRED OUTPUT FORMAT:

    # ## 1. Architectural Overview
    # [Evaluate the model's structural integrity. Does it follow best practices for health economic modeling? Describe the overall data flow.]

    # ## 2. Systemic & Macro-Level Risks
    # [Provide a bulleted list of systemic flaws (e.g., circular data flows, poor parameter isolation, dangerous macro usage, or structural brittleness). Explain *why* these make the model hard to validate or prone to cascading errors.]

    # ## 3. High-Priority Isolated Risks
    # [Provide a Markdown table of the top 10 most critical individual formula or logic risks found across the workbook.]
    # | Sheet | Cell | Issue Type | Formula / Value | Systemic Impact |
    # |---|---|---|---|---|
    # | [Name] | [A1] | [e.g., Hidden Fudge Factor] | [The raw formula] | [Explanation] |

    # ## 4. Audit Recommendations
    # [Provide 2-3 concrete recommendations for rebuilding or refactoring the most fragile parts of the workbook.]"""
    #         # 2. Build the prompt dynamically based on selection
    #         elif "VBA" in prompt_type:
    #             vba_json = [m.model_dump() for m in data.vba_modules] if data.vba_modules else "No VBA modules found."
    #             prompt_text = f"""{system_role}

    # I am auditing a health economic Excel model. Please review the following extracted VBA modules.

    # Task:
    # Analyze the VBA code and populate the EXACT Markdown template provided below. Do not deviate from this structure.

    # ### REQUIRED OUTPUT FORMAT:

    # ## 1. Executive Summary
    # [Provide a 2-3 sentence summary of what the VBA in this model is primarily designed to do.]

    # ## 2. Basic Macros (Navigation & UI)
    # [Provide a Markdown table of simple macros used for sheet navigation, resetting views, or basic formatting. Exclude calculation macros.]
    # | Macro Name | Primary Purpose |
    # |---|---|
    # | [Name] | [e.g., Navigates to the 'Setup' tab] |

    # ## 3. Substantive Macros (Calculations & Simulations)
    # [Provide a Markdown table of complex macros that perform logic, run Probabilistic Sensitivity Analysis (PSA) loops, or manipulate health economic inputs/outputs.]
    # | Macro Name | Logic Summary | Key Variables / Sheets Affected |
    # |---|---|---|
    # | [Name] | [Brief explanation of the calculation/simulation steps] | [e.g., 'Results' sheet, rngInitialAge] |

    # ## 4. Risk & Issue Register
    # [Provide a Markdown table of potential structural or logical risks to the model's validity. Focus specifically on hardcoded parameters, magic numbers, hidden logic overrides, and unsafe sheet manipulations. Do NOT flag general code inefficiencies or suggest runtime optimizations.]
    # | Severity | Location (Module/Sub) | Issue Type | Description | Recommendation |
    # |---|---|---|---|---|
    # | [High/Med/Low] | [Name] | [e.g., Hardcoded Constant] | [What is wrong] | [How to fix it safely] |

    # ***

    # Data (JSON):
    # {vba_json}"""

    #         elif "Trace Logic" in prompt_type:
    #             trace_data = st.session_state.get("last_trace", None)
    #             if trace_data:
    #                 prompt_text = f"""{system_role}

    # I have traced the logic dependency tree for a specific cell in a health economic model.
    # Target Cell: {trace_data['target']}

    # Task:
    # 1. Explain the step-by-step logic of this calculation in plain English.
    # 2. What health economic parameter is this likely calculating (e.g., transition probability, discounted cost, state membership)?
    # 3. Are there any obvious structural risks in this logic chain (e.g., hardcoded constants mixed with dynamic variables)?

    # Trace Tree (JSON):
    # {trace_data['trace_text']}"""
    #             else:
    #                 prompt_text = "⚠️ Please go to the 'Audit Diagnostics' tab and run a Logic Trace first."

    #         elif "Audit Triage" in prompt_type:
    #             prompt_text = f"""{system_role}

    # I have run a deterministic python audit on a health economic model. Please review the findings.

    # Task:
    # 1. Review the provided JSON containing Magic Numbers, Broken References, and Highly Complex formulas.
    # 2. Prioritize the top 3 highest-risk issues that require immediate manual review.
    # 3. Explain WHY these 3 pose a risk to the model's validity.

    # Audit Data:
    # {audits.get('magic_numbers')[:5]} # Limiting to top 5 to save tokens
    # {audits.get('complex_logic')[:5]}"""

    #         elif "Architecture" in prompt_type:
    #             sheet_names = [ws.name for ws in data.worksheets.values()]
    #             named_ranges = [nr.name for nr in data.named_ranges]
    #             prompt_text = f"""{system_role}

    # I am auditing a new health economic Excel model. I have extracted its structural metadata.

    # Task:
    # 1. Based on the sheet names and named ranges, what type of model is this likely to be (e.g., Markov, Partitioned Survival, Decision Tree)?
    # 2. Does the structure adhere to standard health economic modeling best practices (e.g., separating inputs, engine, and outputs)?
    # 3. What key tabs or logic sections should I focus my manual audit on?

    # Sheets: {sheet_names}
    # Named Ranges (sample): {named_ranges[:20]}"""

    #         # 3. Display the prompt
    #         st.markdown("### Generated Prompt")
    #         st.info("Click the copy icon in the top right of the code block below, then paste it into your LLM.")
    #         st.code(prompt_text, language="markdown")