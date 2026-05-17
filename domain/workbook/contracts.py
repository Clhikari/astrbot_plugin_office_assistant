from __future__ import annotations

import re

from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator

from .models import WorkbookCellValue, WorkbookModel, validate_workbook_cell_value

DEFAULT_XLSX_FILENAME = "workbook.xlsx"
MAX_WORKBOOK_ROW_INDEX = 100_000
MAX_EXCEL_ROW = 1_048_576
MAX_EXCEL_COLUMN = 16_384
WINDOWS_DRIVE_PATTERN = re.compile(r"^[A-Za-z]:([\\/]|$)")
SHEET_NAME_FORBIDDEN_CHARS = set("[]:*?/\\")
_CELL_REF_PATTERN = re.compile(r"^([A-Z]{1,3})([1-9][0-9]*)$")
_COLUMN_LETTER_PATTERN = re.compile(r"^[A-Z]{1,3}$")


def _split_path_parts(value: str) -> list[str]:
    return [
        part
        for part in re.split(r"[\\/]+", value.strip())
        if part and part not in {".", ""}
    ]


def _looks_like_absolute_path(value: str) -> bool:
    candidate = value.strip()
    return (
        candidate.startswith(("/", "\\", "~"))
        or WINDOWS_DRIVE_PATTERN.match(candidate) is not None
    )


def _normalize_xlsx_filename(
    value: str,
    default: str = DEFAULT_XLSX_FILENAME,
) -> str:
    parts = _split_path_parts(value)
    candidate = parts[-1] if parts else default
    candidate = candidate or default
    if not candidate.lower().endswith(".xlsx"):
        candidate = f"{candidate}.xlsx"
    return candidate


def _normalize_xlsx_output_path(
    value: str,
    default: str = DEFAULT_XLSX_FILENAME,
) -> str:
    parts = _split_path_parts(value)
    if not parts:
        return default
    normalized_leaf = _normalize_xlsx_filename(parts[-1], default=default)
    if len(parts) == 1:
        return normalized_leaf
    return "/".join([*parts[:-1], normalized_leaf])


def _normalize_sheet_name(value: str) -> str:
    candidate = value.strip()
    if not candidate:
        raise ValueError("sheet name cannot be empty")
    if len(candidate) > 31:
        raise ValueError("sheet name cannot exceed 31 characters")
    if any(char in SHEET_NAME_FORBIDDEN_CHARS for char in candidate):
        raise ValueError("sheet name contains unsupported characters")
    if candidate[0] == "'" or candidate[-1] == "'":
        raise ValueError("sheet name cannot start or end with apostrophe")
    return candidate


def _column_letter_to_index(letter: str) -> int:
    result = 0
    for char in letter:
        result = result * 26 + (ord(char) - ord("A") + 1)
    return result


class WriteRowsOptions(BaseModel):
    model_config = ConfigDict(extra="forbid")

    freeze_panes: str | None = None
    column_widths: dict[str, float] | None = None
    number_formats: dict[str, str] | None = None
    autofilter: bool | None = None

    @field_validator("freeze_panes")
    @classmethod
    def validate_freeze_panes(cls, value: str | None) -> str | None:
        if value is None:
            return None
        candidate = value.strip().upper()
        if not candidate:
            return ""
        match = _CELL_REF_PATTERN.match(candidate)
        if not match:
            raise ValueError(
                "freeze_panes must be a valid cell reference (e.g. 'A2', 'B3')"
            )
        col_letter, row_str = match.group(1), match.group(2)
        if _column_letter_to_index(col_letter) > MAX_EXCEL_COLUMN:
            raise ValueError(
                f"freeze_panes column must not exceed XFD ({MAX_EXCEL_COLUMN} columns)"
            )
        if int(row_str) > MAX_EXCEL_ROW:
            raise ValueError(f"freeze_panes row must not exceed {MAX_EXCEL_ROW}")
        return candidate

    @field_validator("column_widths")
    @classmethod
    def validate_column_widths(
        cls, value: dict[str, float] | None
    ) -> dict[str, float] | None:
        if value is None:
            return None
        normalized: dict[str, float] = {}
        for key, width in value.items():
            col = str(key).strip().upper()
            if not _COLUMN_LETTER_PATTERN.match(col):
                raise ValueError(
                    f"column_widths key must be a column letter (A-XFD), got '{key}'"
                )
            if _column_letter_to_index(col) > MAX_EXCEL_COLUMN:
                raise ValueError(
                    f"column_widths key '{key}' exceeds maximum column XFD"
                )
            if not (1.0 <= width <= 255.0):
                raise ValueError(
                    f"column_widths value must be between 1.0 and 255.0, got {width} for column '{col}'"
                )
            normalized[col] = width
        return normalized

    @field_validator("number_formats")
    @classmethod
    def validate_number_formats(
        cls, value: dict[str, str] | None
    ) -> dict[str, str] | None:
        if value is None:
            return None
        normalized: dict[str, str] = {}
        for key, fmt in value.items():
            col = str(key).strip().upper()
            if not _COLUMN_LETTER_PATTERN.match(col):
                raise ValueError(
                    f"number_formats key must be a column letter (A-XFD), got '{key}'"
                )
            if _column_letter_to_index(col) > MAX_EXCEL_COLUMN:
                raise ValueError(
                    f"number_formats key '{key}' exceeds maximum column XFD"
                )
            if not isinstance(fmt, str) or not fmt.strip():
                raise ValueError(
                    f"number_formats value must be a non-empty format string, got '{fmt}' for column '{col}'"
                )
            normalized[col] = fmt
        return normalized


class CreateWorkbookRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    session_id: str = ""
    title: str = ""
    filename: str = DEFAULT_XLSX_FILENAME

    @field_validator("filename")
    @classmethod
    def validate_filename(cls, value: str) -> str:
        return _normalize_xlsx_filename(value)


class WriteRowsRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    workbook_id: str
    sheet: str
    rows: list[list[WorkbookCellValue]] = Field(min_length=1)
    start_row: int = Field(default=1, ge=1, le=MAX_WORKBOOK_ROW_INDEX)
    options: WriteRowsOptions | None = None

    @field_validator("workbook_id")
    @classmethod
    def validate_workbook_id(cls, value: str) -> str:
        candidate = str(value or "").strip()
        if not candidate:
            raise ValueError("workbook_id must not be empty")
        return candidate

    @field_validator("sheet")
    @classmethod
    def validate_sheet(cls, value: str) -> str:
        return _normalize_sheet_name(value)

    @field_validator("rows", mode="before")
    @classmethod
    def validate_rows(cls, value: object) -> list[list[WorkbookCellValue]]:
        if not isinstance(value, list):
            raise ValueError("rows must be a two-dimensional array")
        normalized_rows: list[list[WorkbookCellValue]] = []
        for row in value:
            if not isinstance(row, list):
                raise ValueError("rows must be a two-dimensional array")
            normalized_row: list[WorkbookCellValue] = []
            for cell in row:
                validated_cell = validate_workbook_cell_value(cell)
                if isinstance(validated_cell, str) and validated_cell.startswith("="):
                    raise ValueError(
                        "formula cell values are not supported in write_rows"
                    )
                normalized_row.append(validated_cell)
            normalized_rows.append(normalized_row)
        if not normalized_rows:
            raise ValueError("rows must contain at least one row")
        return normalized_rows

    @model_validator(mode="after")
    def validate_row_window(self) -> WriteRowsRequest:
        final_row = self.start_row + len(self.rows) - 1
        if final_row > MAX_WORKBOOK_ROW_INDEX:
            raise ValueError(
                f"final written row must not exceed {MAX_WORKBOOK_ROW_INDEX}"
            )
        return self


class ExportWorkbookRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    workbook_id: str
    output_name: str = ""

    @field_validator("workbook_id")
    @classmethod
    def validate_workbook_id(cls, value: str) -> str:
        candidate = str(value or "").strip()
        if not candidate:
            raise ValueError("workbook_id must not be empty")
        return candidate

    @field_validator("output_name")
    @classmethod
    def validate_output_name(cls, value: str) -> str:
        candidate = str(value or "").strip()
        if not candidate:
            return ""
        if _looks_like_absolute_path(candidate):
            raise ValueError("output_name must not be an absolute path")
        return _normalize_xlsx_output_path(candidate)


class WorkbookSummary(BaseModel):
    model_config = ConfigDict(extra="forbid")

    workbook_id: str
    session_id: str
    title: str
    format: str
    status: str
    sheet_names: list[str] = Field(default_factory=list)
    sheet_count: int
    latest_written_sheets: list[str] = Field(default_factory=list)
    output_path: str = ""
    preferred_filename: str


class ToolResult(BaseModel):
    model_config = ConfigDict(extra="forbid")

    success: bool
    message: str
    workbook: WorkbookSummary | None = None


class ExportWorkbookResult(ToolResult):
    model_config = ConfigDict(extra="forbid")

    file_path: str = ""


def _build_workbook_summary_payload(workbook_model: WorkbookModel) -> dict[str, object]:
    return {
        "workbook_id": workbook_model.workbook_id,
        "session_id": workbook_model.session_id,
        "title": workbook_model.metadata.title,
        "format": workbook_model.format,
        "status": workbook_model.status.value,
        "sheet_names": [worksheet.name for worksheet in workbook_model.worksheets],
        "sheet_count": len(workbook_model.worksheets),
        "latest_written_sheets": list(workbook_model.latest_written_sheets),
        "output_path": workbook_model.output_path,
        "preferred_filename": workbook_model.metadata.preferred_filename,
    }


def build_workbook_summary(workbook_model: WorkbookModel) -> WorkbookSummary:
    return WorkbookSummary(**_build_workbook_summary_payload(workbook_model))


def parse_write_rows_options(raw: object) -> WriteRowsOptions | None:
    if raw is None:
        return None
    if isinstance(raw, dict):
        return WriteRowsOptions(**raw)
    raise ValueError("options must be an object or null")


__all__ = [
    "CreateWorkbookRequest",
    "DEFAULT_XLSX_FILENAME",
    "ExportWorkbookRequest",
    "ExportWorkbookResult",
    "MAX_EXCEL_COLUMN",
    "MAX_EXCEL_ROW",
    "ToolResult",
    "WorkbookSummary",
    "WriteRowsOptions",
    "WriteRowsRequest",
    "build_workbook_summary",
    "parse_write_rows_options",
]
