from __future__ import annotations

import re
from pathlib import Path
from typing import Annotated, Literal

from pydantic import BaseModel, ConfigDict, Field, field_validator

from ..document_core.models.document import DocumentModel

SUPPORTED_THEMES = {"business_report", "project_review", "executive_brief"}
SUPPORTED_TABLE_TEMPLATES = {"report_grid", "metrics_compact", "minimal"}
SUPPORTED_DENSITIES = {"comfortable", "compact"}
SUPPORTED_CARD_VARIANTS = {"summary", "conclusion"}
WINDOWS_DRIVE_PATTERN = re.compile(r"^[A-Za-z]:([\\/]|$)")


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


def _normalize_docx_filename(value: str, default: str = "document.docx") -> str:
    candidate = _split_path_parts(value)[-1] if value.strip() else default
    candidate = candidate or default
    if not candidate.lower().endswith(".docx"):
        candidate = f"{candidate}.docx"
    return candidate


class CreateDocumentRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    session_id: str = ""
    title: str = ""
    output_name: str = "document.docx"
    theme_name: str = "business_report"
    table_template: str = "report_grid"
    density: str = "comfortable"
    accent_color: str = ""

    @field_validator("output_name")
    @classmethod
    def validate_output_name(cls, value: str) -> str:
        return _normalize_docx_filename(value)

    @field_validator("theme_name")
    @classmethod
    def validate_theme_name(cls, value: str) -> str:
        candidate = value.strip() or "business_report"
        return candidate if candidate in SUPPORTED_THEMES else "business_report"

    @field_validator("table_template")
    @classmethod
    def validate_table_template(cls, value: str) -> str:
        candidate = value.strip() or "report_grid"
        return candidate if candidate in SUPPORTED_TABLE_TEMPLATES else "report_grid"

    @field_validator("density")
    @classmethod
    def validate_density(cls, value: str) -> str:
        candidate = value.strip() or "comfortable"
        return candidate if candidate in SUPPORTED_DENSITIES else "comfortable"

    @field_validator("accent_color")
    @classmethod
    def validate_accent_color(cls, value: str) -> str:
        candidate = value.strip().lstrip("#").upper()
        if not candidate:
            return ""
        if len(candidate) != 6 or any(
            char not in "0123456789ABCDEF" for char in candidate
        ):
            return ""
        return candidate


class AddHeadingRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    document_id: str
    text: str = Field(min_length=1)
    level: int = Field(default=1, ge=1, le=6)


class AddParagraphRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    document_id: str
    text: str = Field(min_length=1)


class SectionParagraphInput(BaseModel):
    model_config = ConfigDict(extra="forbid")

    type: Literal["paragraph"] = "paragraph"
    text: str = Field(min_length=1)


class AddTableRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    document_id: str
    headers: list[str] = Field(default_factory=list)
    rows: list[list[str]] = Field(default_factory=list)
    table_style: str = ""

    @field_validator("table_style")
    @classmethod
    def validate_table_style(cls, value: str) -> str:
        candidate = value.strip()
        if not candidate:
            return ""
        return candidate if candidate in SUPPORTED_TABLE_TEMPLATES else ""


class SectionTableInput(BaseModel):
    model_config = ConfigDict(extra="forbid")

    type: Literal["table"] = "table"
    headers: list[str] = Field(default_factory=list)
    rows: list[list[str]] = Field(default_factory=list)
    table_style: str = ""

    @field_validator("table_style")
    @classmethod
    def validate_table_style(cls, value: str) -> str:
        candidate = value.strip()
        if not candidate:
            return ""
        return candidate if candidate in SUPPORTED_TABLE_TEMPLATES else ""


class AddSummaryCardRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    document_id: str
    title: str = Field(min_length=1)
    items: list[str] = Field(min_length=1)
    variant: str = "summary"

    @field_validator("items")
    @classmethod
    def validate_items(cls, value: list[str]) -> list[str]:
        cleaned = [item.strip() for item in value if item and item.strip()]
        if not cleaned:
            raise ValueError("items must contain at least one non-empty item")
        return cleaned

    @field_validator("variant")
    @classmethod
    def validate_variant(cls, value: str) -> str:
        candidate = value.strip() or "summary"
        return candidate if candidate in SUPPORTED_CARD_VARIANTS else "summary"


class SectionCardInput(BaseModel):
    model_config = ConfigDict(extra="forbid")

    type: Literal["summary_card"] = "summary_card"
    title: str = Field(min_length=1)
    items: list[str] = Field(min_length=1)
    variant: str = "summary"

    @field_validator("items")
    @classmethod
    def validate_items(cls, value: list[str]) -> list[str]:
        cleaned = [item.strip() for item in value if item and item.strip()]
        if not cleaned:
            raise ValueError("items must contain at least one non-empty item")
        return cleaned

    @field_validator("variant")
    @classmethod
    def validate_variant(cls, value: str) -> str:
        candidate = value.strip() or "summary"
        return candidate if candidate in SUPPORTED_CARD_VARIANTS else "summary"


SectionBundleBlock = Annotated[
    SectionParagraphInput | SectionTableInput | SectionCardInput,
    Field(discriminator="type"),
]


class AddSectionBundleRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    document_id: str
    heading: str = Field(min_length=1)
    level: int = Field(default=1, ge=1, le=6)
    blocks: list[SectionBundleBlock] = Field(min_length=1)


class FinalizeDocumentRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    document_id: str


class ExportDocumentRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    document_id: str
    output_dir: str = ""
    output_name: str = ""

    @field_validator("output_dir")
    @classmethod
    def validate_output_dir(cls, value: str) -> str:
        candidate = value.strip()
        if not candidate:
            return ""

        if _looks_like_absolute_path(candidate):
            raise ValueError("output_dir must be relative to the document workspace")

        normalized_parts = _split_path_parts(candidate)
        if any(part == ".." for part in normalized_parts):
            raise ValueError("output_dir cannot escape the document workspace")

        return "" if not normalized_parts else str(Path(*normalized_parts))

    @field_validator("output_name")
    @classmethod
    def validate_output_name(cls, value: str) -> str:
        return _normalize_docx_filename(value) if value.strip() else ""


class DocumentSummary(BaseModel):
    model_config = ConfigDict(extra="forbid")

    document_id: str
    session_id: str
    title: str
    format: str
    status: str
    block_count: int
    output_path: str = ""
    preferred_filename: str
    theme_name: str = ""
    table_template: str = ""
    density: str = ""
    accent_color: str = ""


class ToolResult(BaseModel):
    model_config = ConfigDict(extra="forbid")

    success: bool
    message: str
    document: DocumentSummary | None = None


class ExportDocumentResult(ToolResult):
    file_path: str = ""


def build_document_summary(document_model: DocumentModel) -> DocumentSummary:
    return DocumentSummary(
        document_id=document_model.document_id,
        session_id=document_model.session_id,
        title=document_model.metadata.title,
        format=document_model.format,
        status=document_model.status.value,
        block_count=len(document_model.blocks),
        output_path=document_model.output_path,
        preferred_filename=document_model.metadata.preferred_filename,
        theme_name=document_model.metadata.theme_name,
        table_template=document_model.metadata.table_template,
        density=document_model.metadata.density,
        accent_color=document_model.metadata.accent_color,
    )
