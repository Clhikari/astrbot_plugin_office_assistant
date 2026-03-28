from __future__ import annotations

from datetime import datetime, timezone
from enum import StrEnum
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field, field_validator

from .blocks import (
    DocumentBlock,
    TableAlignment,
    TableBorderStyle,
    TableCaptionEmphasis,
    normalize_optional_hex_color,
)


def utc_now() -> datetime:
    return datetime.now(timezone.utc)


class DocumentStatus(StrEnum):
    DRAFT = "draft"
    FINALIZED = "finalized"
    EXPORTED = "exported"


class DocumentTableDefaults(BaseModel):
    model_config = ConfigDict(extra="forbid")

    preset: Literal["report_grid", "metrics_compact", "minimal"] | None = None
    header_fill: str | None = None
    header_text_color: str | None = None
    banded_rows: bool | None = None
    banded_row_fill: str | None = None
    first_column_bold: bool | None = None
    table_align: TableAlignment | None = None
    border_style: TableBorderStyle | None = None
    caption_emphasis: TableCaptionEmphasis | None = None
    cell_align: Literal["left", "center", "right"] | None = None

    @field_validator("header_fill", "header_text_color", "banded_row_fill")
    @classmethod
    def validate_optional_colors(cls, value: str | None) -> str | None:
        return normalize_optional_hex_color(value)


class DocumentSummaryCardDefaults(BaseModel):
    model_config = ConfigDict(extra="forbid")

    title_align: Literal["left", "center", "right", "justify"] | None = None
    title_emphasis: Literal["normal", "strong", "subtle"] | None = None
    title_font_scale: float | None = Field(default=None, ge=0.75, le=2.0)
    title_space_before: float | None = Field(default=None, ge=0, le=72)
    title_space_after: float | None = Field(default=None, ge=0, le=72)
    list_space_after: float | None = Field(default=None, ge=0, le=72)


class DocumentStyleConfig(BaseModel):
    model_config = ConfigDict(extra="forbid")

    brief: str = ""
    heading_color: str | None = None
    title_align: Literal["left", "center", "right", "justify"] | None = None
    body_font_size: float | None = Field(default=None, ge=9.0, le=16.0)
    body_line_spacing: float | None = Field(default=None, ge=1.0, le=2.5)
    paragraph_space_after: float | None = Field(default=None, ge=0, le=72)
    list_space_after: float | None = Field(default=None, ge=0, le=72)
    summary_card_defaults: DocumentSummaryCardDefaults = Field(
        default_factory=DocumentSummaryCardDefaults
    )
    table_defaults: DocumentTableDefaults = Field(default_factory=DocumentTableDefaults)

    @field_validator("heading_color")
    @classmethod
    def validate_heading_color(cls, value: str | None) -> str | None:
        return normalize_optional_hex_color(value)


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
    document_style: DocumentStyleConfig = Field(default_factory=DocumentStyleConfig)
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
