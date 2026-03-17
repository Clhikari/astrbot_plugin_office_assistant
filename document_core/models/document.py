from __future__ import annotations

from datetime import datetime, timezone
from enum import StrEnum
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field, field_validator

from .blocks import DocumentBlock


def utc_now() -> datetime:
    return datetime.now(timezone.utc)


class DocumentStatus(StrEnum):
    DRAFT = "draft"
    FINALIZED = "finalized"
    EXPORTED = "exported"


class DocumentMetadata(BaseModel):
    model_config = ConfigDict(extra="forbid")

    title: str = ""
    preferred_filename: str = "document.docx"
    theme_name: Literal[
        "business_report",
        "project_review",
        "executive_brief",
    ] = "business_report"
    table_template: Literal[
        "report_grid",
        "metrics_compact",
        "minimal",
    ] = "report_grid"
    density: Literal["comfortable", "compact"] = "comfortable"
    accent_color: str = ""
    created_at: datetime = Field(default_factory=utc_now)
    updated_at: datetime = Field(default_factory=utc_now)

    @field_validator("accent_color")
    @classmethod
    def normalize_accent_color(cls, value: str) -> str:
        candidate = value.strip().lstrip("#").upper()
        if not candidate:
            return ""
        if len(candidate) != 6 or any(
            char not in "0123456789ABCDEF" for char in candidate
        ):
            return ""
        return candidate


class DocumentModel(BaseModel):
    model_config = ConfigDict(extra="forbid")

    document_id: str
    session_id: str = ""
    format: Literal["word"] = "word"
    status: DocumentStatus = DocumentStatus.DRAFT
    metadata: DocumentMetadata = Field(default_factory=DocumentMetadata)
    blocks: list[DocumentBlock] = Field(default_factory=list)
    output_path: str = ""

    def touch(self) -> None:
        self.metadata.updated_at = utc_now()

    def add_block(self, block: DocumentBlock) -> None:
        self.blocks.append(block)
        self.touch()
