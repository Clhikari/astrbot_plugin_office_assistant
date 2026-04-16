from __future__ import annotations

import re
from pathlib import Path

from pydantic import BaseModel, ConfigDict, Field, field_validator

from .models import WorkbookCellValue, WorkbookModel

DEFAULT_XLSX_FILENAME = "workbook.xlsx"
WINDOWS_DRIVE_PATTERN = re.compile(r"^[A-Za-z]:([\\/]|$)")
SHEET_NAME_FORBIDDEN_CHARS = set("[]:*?/\\")


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


class CreateWorkbookRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    session_id: str = ""
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
    start_row: int = Field(default=1, ge=1)

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

    @field_validator("rows")
    @classmethod
    def validate_rows(
        cls,
        value: list[list[WorkbookCellValue]],
    ) -> list[list[WorkbookCellValue]]:
        normalized_rows: list[list[WorkbookCellValue]] = []
        for row in value:
            if not isinstance(row, list):
                raise TypeError("rows must be a two-dimensional array")
            for cell in row:
                if isinstance(cell, str) and cell.startswith("="):
                    raise ValueError("formula cell values are not supported in write_rows")
            normalized_rows.append(list(row))
        if not normalized_rows:
            raise ValueError("rows must contain at least one row")
        return normalized_rows


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
        return _normalize_xlsx_filename(candidate)


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


def build_workbook_summary(workbook_model: WorkbookModel) -> WorkbookSummary:
    return WorkbookSummary(
        workbook_id=workbook_model.workbook_id,
        session_id=workbook_model.session_id,
        title=workbook_model.metadata.title,
        format=workbook_model.format,
        status=workbook_model.status.value,
        sheet_names=[worksheet.name for worksheet in workbook_model.worksheets],
        sheet_count=len(workbook_model.worksheets),
        latest_written_sheets=list(workbook_model.latest_written_sheets),
        output_path=workbook_model.output_path,
        preferred_filename=workbook_model.metadata.preferred_filename,
    )


__all__ = [
    "CreateWorkbookRequest",
    "DEFAULT_XLSX_FILENAME",
    "ExportWorkbookRequest",
    "ExportWorkbookResult",
    "ToolResult",
    "WorkbookSummary",
    "WriteRowsRequest",
    "build_workbook_summary",
]
