> **WARNING: This project is a work in progress.** Features are incomplete, the API is unstable, and the tool is not yet ready for production use.

# xl-ray

**Interrogate Excel-based health economic models and generate LLM-powered audit prompts.**

xl-ray extracts structural data from complex Excel workbooks (Markov models, decision trees, cost-effectiveness analyses), runs deterministic diagnostic checks, and produces structured prompts for AI-assisted auditing.

## Features

### Extraction

- Full workbook introspection — worksheets, cells, formulas, evaluated values
- Precedent and dependent tracking for every formula cell
- VBA macro extraction with worksheet linkage (for `.xlsm` files)
- Named range extraction at both workbook and worksheet scope
- Excel Data Table (ListObject) detection
- Array formula detection and spill propagation
- Hidden/VeryHidden sheet visibility tracking
- External link detection
- Dual-load approach: one open for formulas/structure, one for cached values

### Deterministic Audits

| Audit | What it finds |
|---|---|
| **Magic Numbers** | Hardcoded constants in formulas (ignores structural values like 0, 1, -1) |
| **Hidden Logic** | Visible cells pulling data from Hidden or VeryHidden sheets |
| **Broken References** | `#REF!` errors, missing cells, dead cross-sheet links, error-state precedents |
| **Inconsistent Columns** | Formulas that break an established pattern in a column (e.g. manual "fudge factors" in Markov traces) |
| **Complexity & Depth** | Highly nested formulas or deep dependency chains flagged by configurable thresholds |
| **Array Formulas** | Identification of all array formulas with their spill ranges |

### Dashboard interface

A Streamlit-based interface (`app.py`) provides:

**Tabs:**

- **📊 Model Summary** — Worksheet overview, metadata, cell counts
- **🛠️ Audit Diagnostics** — Filtered audit results with expandable views; adjustable complexity thresholds
- **📜 VBA Modules** — Browser for embedded VBA code per module with worksheet linkage
- **🔍 Raw JSON** — Download extracted data as compressed JSON
- **🤖 LLM Copilot Prompt Builder** — Generate structured prompts for a LLM to analyse the spreadsheet

### Logic Dependency Tracer

Step through any cell's computation tree. See formulas, constants, and reference chains in an interactive HTML tree.

## Installation

### Minimum Requirements

- Python >=3.11

### Install via Conda

```bash
git clone https://github.com/TomMonks/xl-ray.git
cd xl-ray
conda env create -f environment.yml
conda activate xl-ray
```

### Dependencies

| Package | Purpose |
|---|---|
| [openpyxl](https://openpyxl.readthedocs.io/) | Excel .xlsx/.xlsm parsing |
| [pydantic](https://docs.pydantic.dev/) | Data models & validation |
| [oletools](https://github.com/decalage2/oletools) | VBA macro extraction |
| [rich](https://rich.readthedocs.io/) | CLI output |
| [streamlit](https://streamlit.io/) | Web UI |
| [pandas](https://pandas.pydata.org/) | Data display in web UI |
| [pytest](https://docs.pytest.org/) | Testing |

## Usage

### Web UI

```bash
streamlit run app.py
```

Upload an Excel file, review the extraction and diagnostics, and generate LLM prompts.

### As a Library

```python
from xl_ray.schema import ExcelModelData
from xl_ray.audit import detect_magic_numbers, detect_broken_references, trace_cell_logic

# ...load workbook via extractor, build ExcelModelData, run audits...
```

## Project Structure

```
xl-ray/
├── src/xl_ray/
│   ├── __init__.py          # Package init, version info
│   ├── schema.py            # Pydantic data models
│   ├── extractor.py         # Workbook parsing logic
│   └── audit.py             # Deterministic audit functions
├── tests/                    # Unit tests
├── output/                   # Extracted JSON output
├── app.py                    # Streamlit web interface
├── main.py                   # CLI entrypoint
├── pyproject.toml            # Hatch package config
└── environment.yml           # Conda environment
```

## License

MIT — See [LICENSE](LICENSE) for details.

## Author

Tom Monks
