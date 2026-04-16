from __future__ import annotations

from datetime import datetime, timezone
from enum import Enum

from pydantic import BaseModel, ConfigDict, Field


def _utc_now() -> datetime:
    return datetime.now(timezone.utc)


WorkbookCellValue = str | int | float | bool | None


class WorkbookStatus(str, Enum):
    DRAFT = "draft"
    EXPORTED = "exported"


class WorksheetOptions(BaseModel):
    model_config = ConfigDict(extra="forbid")

    freeze_panes: str = ""
    column_widths: dict[str, float] = Field(default_factory=dict)
    autofilter: bool = False


class WorksheetModel(BaseModel):
    model_config = ConfigDict(extra="forbid")

    name: str
    rows: list[list[WorkbookCellValue]] = Field(default_factory=list)
    options: WorksheetOptions = Field(default_factory=WorksheetOptions)


class WorkbookMetadata(BaseModel):
    model_config = ConfigDict(extra="forbid")

    title: str = ""
    preferred_filename: str = "workbook.xlsx"
    created_at: datetime = Field(default_factory=_utc_now)
    updated_at: datetime = Field(default_factory=_utc_now)


class WorkbookModel(BaseModel):
    model_config = ConfigDict(extra="forbid")

    workbook_id: str
    session_id: str = ""
    format: str = "excel"
    status: WorkbookStatus = WorkbookStatus.DRAFT
    metadata: WorkbookMetadata = Field(default_factory=WorkbookMetadata)
    worksheets: list[WorksheetModel] = Field(default_factory=list)
    output_path: str = ""
    latest_written_sheets: list[str] = Field(default_factory=list)

    def touch(self) -> None:
        self.metadata.updated_at = _utc_now()

    def get_sheet(self, sheet_name: str) -> WorksheetModel | None:
        for worksheet in self.worksheets:
            if worksheet.name == sheet_name:
                return worksheet
        return None

    def get_worksheet(self, sheet_name: str) -> WorksheetModel | None:
        return self.get_sheet(sheet_name)

    def get_or_create_sheet(self, sheet_name: str) -> WorksheetModel:
        worksheet = self.get_sheet(sheet_name)
        if worksheet is not None:
            return worksheet
        worksheet = WorksheetModel(name=sheet_name)
        self.worksheets.append(worksheet)
        self.touch()
        return worksheet

    def remember_written_sheet(self, sheet_name: str) -> None:
        self.latest_written_sheets = [
            existing for existing in self.latest_written_sheets if existing != sheet_name
        ]
        self.latest_written_sheets.append(sheet_name)
        self.latest_written_sheets = self.latest_written_sheets[-3:]
        self.touch()
