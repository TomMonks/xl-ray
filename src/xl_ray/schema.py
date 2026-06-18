"""
Defines schema for extracted Excel workbook
"""

from typing import Any, Optional
from pydantic import BaseModel, Field

class CellData(BaseModel):
    """Represents the contents and evaluated value of a single Excel cell."""
    address: str = Field(description="The A1 reference style address of the cell (e.g., 'A1').")
    value: Any = Field(default=None, description="The evaluated, cached value of the cell.")
    formula: Optional[str] = Field(default=None, description="The raw Excel formula, if present.")

    precedents: list[str] = Field(
        default_factory=list, 
        description="List of cell addresses or named ranges this cell references (its inputs)."
    )
    
    dependents: list[str] = Field(
        default_factory=list, 
        description="List of cell addresses that reference this cell (its outputs)."
    )
    
    # Updated fields for array formulas:
    is_array_formula: bool = Field(default=False, description="True if the cell is part of an array formula.")
    array_range: Optional[str] = Field(
        default=None, 
        description="The range this array formula spills over (e.g., 'A1:A5'). Populated on the top-left cell."
    )
    parent_array_cell: Optional[str] = Field(
        default=None, 
        description="If this cell is part of an array but not the top-left cell, this points to the cell with the formula."
    )
    
    data_type: Optional[str] = Field(default=None, description="The data type of the cell's evaluated value.")

class NamedRange(BaseModel):
    """Represents an Excel defined name (named range)."""
    name: str = Field(description="The defined name.")
    refers_to: str = Field(description="The range, constant, or formula the name refers to.")
    scope: str = Field(default="Workbook", description="The scope of the named range (Workbook or a specific Worksheet name).")

class ExcelTable(BaseModel):
    """Represents an Excel Data Table (ListObject)."""
    name: str = Field(description="The name of the table.")
    range_address: str = Field(description="The address range the table occupies.")
    columns: list[str] = Field(default_factory=list, description="A list of the table's column headers.")

class WorksheetData(BaseModel):
    """Represents a single tab/worksheet within the Excel workbook."""
    name: str = Field(description="The name of the worksheet.")
    visibility: str = Field(default="Visible", description="Visibility state: 'Visible', 'Hidden', or 'VeryHidden'.")
    cells: dict[str, CellData] = Field(default_factory=dict, description="A dictionary of cells keyed by their A1 address.")
    tables: list[ExcelTable] = Field(default_factory=list, description="A list of Excel tables present on this worksheet.")

class VBAModule(BaseModel):
    """Represents a VBA code module extracted from the workbook."""
    filename: str = Field(description="The name of the VBA module file/component.")
    content: str = Field(description="The raw text content of the VBA code.")
    
    # --- NEW FIELD ---
    linked_worksheet: Optional[str] = Field(
        default=None, 
        description="The user-facing name of the worksheet this module is attached to (if it is a Sheet object)."
    )

class WorkbookMetadata(BaseModel):
    """High-level metadata regarding the source Excel file to provide an immediate summary."""
    file_name: str = Field(description="The name of the source Excel file.")
    file_path: str = Field(description="The absolute or relative path to the source Excel file.")
    file_type: str = Field(description="The file extension/type (e.g., '.xlsx', '.xlsm').")
    
    # Feature flags for quick auditing triage
    has_macros: bool = Field(default=False, description="True if the workbook contains VBA macros.")
    has_array_formulas: bool = Field(default=False, description="True if the workbook utilizes array formulas.")
    has_data_tables: bool = Field(default=False, description="True if the workbook contains Excel Data Tables (ListObjects).")
    has_named_ranges: bool = Field(default=False, description="True if the workbook has defined named ranges.")
    has_hidden_sheets: bool = Field(default=False, description="True if the workbook contains Hidden or VeryHidden sheets.")
    has_external_links: bool = Field(default=False, description="True if the workbook links to external workbooks.")
    
    # Optional: Quick counts can also be helpful for the auditor to gauge complexity
    worksheet_count: int = Field(default=0, description="Total number of worksheets.")

class ExcelModelData(BaseModel):
    """
    The root schema representing the completely extracted contents of the Excel file.
    This acts as the intermediate JSON format for xl-ray.
    """
    metadata: WorkbookMetadata
    worksheets: dict[str, WorksheetData] = Field(default_factory=dict, description="A dictionary of worksheets keyed by sheet name.")
    named_ranges: list[NamedRange] = Field(default_factory=list, description="A list of all named ranges in the workbook.")
    vba_modules: list[VBAModule] = Field(default_factory=list, description="A list of extracted VBA modules.")
