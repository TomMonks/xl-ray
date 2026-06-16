from typing import Any, Optional
from pydantic import BaseModel, Field

class CellData(BaseModel):
    """Represents the contents and evaluated value of a single Excel cell."""
    address: str = Field(description="The A1 reference style address of the cell (e.g., 'A1').")
    value: Any = Field(default=None, description="The evaluated, cached value of the cell.")
    formula: Optional[str] = Field(default=None, description="The raw Excel formula, if present.")
    is_array_formula: bool = Field(default=False, description="True if the cell is part of an array formula.")
    data_type: Optional[str] = Field(default=None, description="The data type of the cell's evaluated value (e.g., 'float', 'str', 'error').")

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

class WorkbookMetadata(BaseModel):
    """High-level metadata regarding the source Excel file."""
    file_name: str = Field(description="The name of the source Excel file.")
    file_type: str = Field(description="The file extension/type (e.g., '.xlsx', '.xlsm').")
    has_macros: bool = Field(default=False, description="True if the workbook contains VBA macros.")

class ExcelModelData(BaseModel):
    """
    The root schema representing the completely extracted contents of the Excel file.
    This acts as the intermediate JSON format for xl-ray.
    """
    metadata: WorkbookMetadata
    worksheets: dict[str, WorksheetData] = Field(default_factory=dict, description="A dictionary of worksheets keyed by sheet name.")
    named_ranges: list[NamedRange] = Field(default_factory=list, description="A list of all named ranges in the workbook.")
    vba_modules: list[VBAModule] = Field(default_factory=list, description="A list of extracted VBA modules.")
