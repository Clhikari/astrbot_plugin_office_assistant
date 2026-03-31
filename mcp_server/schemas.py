from __future__ import annotations

import copy
import re
from pathlib import Path
from typing import Annotated, Literal

from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator

from ..document_core.models.blocks import (
    BlockLayout,
    BlockStyle,
    HeaderFooterConfig,
    ParagraphRun,
    SectionMarginsConfig,
    SectionStartType,
    TableAlignment,
    TableBorderStyle,
    TableCaptionEmphasis,
    TableHeaderGroup,
    normalize_optional_hex_color,
    validate_section_page_numbering,
    validate_table_structure,
)
from ..document_core.models.document import DocumentModel, DocumentStyleConfig

SUPPORTED_THEMES = {"business_report", "project_review", "executive_brief"}
SUPPORTED_TABLE_TEMPLATES = {"report_grid", "metrics_compact", "minimal"}
SUPPORTED_DENSITIES = {"comfortable", "compact"}
SUPPORTED_CARD_VARIANTS = {"summary", "conclusion"}
WINDOWS_DRIVE_PATTERN = re.compile(r"^[A-Za-z]:([\\/]|$)")

_HEADER_FOOTER_SCHEMA_PROPERTIES = {
    "header_text": {
        "type": "string",
        "description": "Optional repeated header text for the document.",
    },
    "footer_text": {
        "type": "string",
        "description": "Optional repeated footer text for the document.",
    },
    "different_first_page": {
        "type": "boolean",
        "description": "Whether the first page should use different header and footer content.",
    },
    "first_page_header_text": {
        "type": "string",
        "description": "Optional first-page-only header text.",
    },
    "first_page_footer_text": {
        "type": "string",
        "description": "Optional first-page-only footer text.",
    },
    "first_page_show_page_number": {
        "type": "boolean",
        "description": "Optional override for whether the first page footer should include a page number.",
    },
    "different_odd_even": {
        "type": "boolean",
        "description": "Whether odd and even pages should use different headers and footers.",
    },
    "even_page_header_text": {
        "type": "string",
        "description": "Optional even-page-only header text.",
    },
    "even_page_footer_text": {
        "type": "string",
        "description": "Optional even-page-only footer text.",
    },
    "even_page_show_page_number": {
        "type": "boolean",
        "description": "Optional override for whether even-page footers should include a page number.",
    },
    "show_page_number": {
        "type": "boolean",
        "description": "Whether to append a PAGE field in the footer.",
    },
    "page_number_align": {
        "type": "string",
        "enum": ["left", "center", "right"],
        "description": "Paragraph alignment used for the footer page number field.",
    },
}


def build_header_footer_schema(*, description: str) -> dict:
    return {
        "type": "object",
        "description": description,
        "properties": copy.deepcopy(_HEADER_FOOTER_SCHEMA_PROPERTIES),
    }


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


def _normalize_table_style(value: str) -> str:
    candidate = value.strip()
    if not candidate:
        return ""
    return candidate if candidate in SUPPORTED_TABLE_TEMPLATES else ""


def _normalize_column_widths(value: list[float]) -> list[float]:
    return [width if width > 0 else 0 for width in value]


def _normalize_numeric_columns(value: list[int]) -> list[int]:
    return sorted({index for index in value if index >= 0})


def _normalize_table_title_text(value: str) -> str:
    return value.strip()


class CreateDocumentRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    session_id: str = ""
    title: str = ""
    output_name: str = "document.docx"
    theme_name: str = "business_report"
    table_template: str = "report_grid"
    density: str = "comfortable"
    accent_color: str = ""
    document_style: DocumentStyleConfig = Field(default_factory=DocumentStyleConfig)
    header_footer: HeaderFooterConfig = Field(default_factory=HeaderFooterConfig)

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
    style: BlockStyle = Field(default_factory=BlockStyle)
    layout: BlockLayout = Field(default_factory=BlockLayout)


class BlockHeadingInput(BaseModel):
    model_config = ConfigDict(extra="forbid")

    type: Literal["heading"] = "heading"
    text: str = Field(min_length=1)
    level: int = Field(default=1, ge=1, le=6)
    style: BlockStyle = Field(default_factory=BlockStyle)
    layout: BlockLayout = Field(default_factory=BlockLayout)


class ParagraphRunInput(ParagraphRun):
    model_config = ConfigDict(extra="forbid")


class AddParagraphRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    document_id: str
    text: str = ""
    variant: Literal["body", "summary_box", "key_takeaway"] = "body"
    title: str = ""
    runs: list[ParagraphRunInput] = Field(default_factory=list)
    style: BlockStyle = Field(default_factory=BlockStyle)
    layout: BlockLayout = Field(default_factory=BlockLayout)

    @model_validator(mode="after")
    def validate_content(self) -> AddParagraphRequest:
        if self.text.strip() or self.runs:
            return self
        raise ValueError("paragraph requires text or runs")


class SectionParagraphInput(BaseModel):
    model_config = ConfigDict(extra="forbid")

    type: Literal["paragraph"] = "paragraph"
    text: str = ""
    variant: Literal["body", "summary_box", "key_takeaway"] = "body"
    title: str = ""
    runs: list[ParagraphRunInput] = Field(default_factory=list)
    style: BlockStyle = Field(default_factory=BlockStyle)
    layout: BlockLayout = Field(default_factory=BlockLayout)

    @model_validator(mode="after")
    def validate_content(self) -> SectionParagraphInput:
        if self.text.strip() or self.runs:
            return self
        raise ValueError("paragraph requires text or runs")


class AddListRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    document_id: str
    items: list[str] = Field(min_length=1)
    ordered: bool = False
    style: BlockStyle = Field(default_factory=BlockStyle)
    layout: BlockLayout = Field(default_factory=BlockLayout)

    @field_validator("items")
    @classmethod
    def validate_items(cls, value: list[str]) -> list[str]:
        cleaned = [item.strip() for item in value if item and item.strip()]
        if not cleaned:
            raise ValueError("items must contain at least one non-empty item")
        return cleaned


class SectionListInput(BaseModel):
    model_config = ConfigDict(extra="forbid")

    type: Literal["list"] = "list"
    items: list[str] = Field(min_length=1)
    ordered: bool = False
    style: BlockStyle = Field(default_factory=BlockStyle)
    layout: BlockLayout = Field(default_factory=BlockLayout)

    @field_validator("items")
    @classmethod
    def validate_items(cls, value: list[str]) -> list[str]:
        cleaned = [item.strip() for item in value if item and item.strip()]
        if not cleaned:
            raise ValueError("items must contain at least one non-empty item")
        return cleaned


class AddTableRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    document_id: str
    headers: list[str] = Field(default_factory=list)
    rows: list[list[str]] = Field(default_factory=list)
    header_groups: list[TableHeaderGroup] = Field(default_factory=list)
    table_style: str = ""
    caption: str = ""
    title: str = ""
    column_widths: list[float] = Field(default_factory=list)
    numeric_columns: list[int] = Field(default_factory=list)
    header_fill: str | None = None
    header_text_color: str | None = None
    banded_rows: bool | None = None
    banded_row_fill: str | None = None
    first_column_bold: bool | None = None
    table_align: TableAlignment | None = None
    border_style: TableBorderStyle | None = None
    caption_emphasis: TableCaptionEmphasis | None = None
    style: BlockStyle = Field(default_factory=BlockStyle)
    layout: BlockLayout = Field(default_factory=BlockLayout)

    @field_validator("table_style")
    @classmethod
    def validate_table_style(cls, value: str) -> str:
        return _normalize_table_style(value)

    @field_validator("caption", "title")
    @classmethod
    def validate_table_title_text(cls, value: str) -> str:
        return _normalize_table_title_text(value)

    @field_validator("column_widths")
    @classmethod
    def validate_column_widths(cls, value: list[float]) -> list[float]:
        return _normalize_column_widths(value)

    @field_validator("numeric_columns")
    @classmethod
    def validate_numeric_columns(cls, value: list[int]) -> list[int]:
        return _normalize_numeric_columns(value)

    @field_validator("header_fill", "header_text_color", "banded_row_fill")
    @classmethod
    def validate_optional_colors(cls, value: str | None) -> str | None:
        return normalize_optional_hex_color(value)

    @model_validator(mode="after")
    def validate_table_shape(self) -> AddTableRequest:
        validate_table_structure(self.headers, self.rows, self.header_groups)
        return self


class SectionTableInput(BaseModel):
    model_config = ConfigDict(extra="forbid")

    type: Literal["table"] = "table"
    headers: list[str] = Field(default_factory=list)
    rows: list[list[str]] = Field(default_factory=list)
    header_groups: list[TableHeaderGroup] = Field(default_factory=list)
    table_style: str = ""
    caption: str = ""
    title: str = ""
    column_widths: list[float] = Field(default_factory=list)
    numeric_columns: list[int] = Field(default_factory=list)
    header_fill: str | None = None
    header_text_color: str | None = None
    banded_rows: bool | None = None
    banded_row_fill: str | None = None
    first_column_bold: bool | None = None
    table_align: TableAlignment | None = None
    border_style: TableBorderStyle | None = None
    caption_emphasis: TableCaptionEmphasis | None = None
    style: BlockStyle = Field(default_factory=BlockStyle)
    layout: BlockLayout = Field(default_factory=BlockLayout)

    @field_validator("table_style")
    @classmethod
    def validate_table_style(cls, value: str) -> str:
        return _normalize_table_style(value)

    @field_validator("caption", "title")
    @classmethod
    def validate_table_title_text(cls, value: str) -> str:
        return _normalize_table_title_text(value)

    @field_validator("column_widths")
    @classmethod
    def validate_column_widths(cls, value: list[float]) -> list[float]:
        return _normalize_column_widths(value)

    @field_validator("numeric_columns")
    @classmethod
    def validate_numeric_columns(cls, value: list[int]) -> list[int]:
        return _normalize_numeric_columns(value)

    @field_validator("header_fill", "header_text_color", "banded_row_fill")
    @classmethod
    def validate_optional_colors(cls, value: str | None) -> str | None:
        return normalize_optional_hex_color(value)

    @model_validator(mode="after")
    def validate_table_shape(self) -> SectionTableInput:
        validate_table_structure(self.headers, self.rows, self.header_groups)
        return self


class AddSummaryCardRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    document_id: str
    title: str = Field(min_length=1)
    items: list[str] = Field(min_length=1)
    variant: str = "summary"
    style: BlockStyle = Field(default_factory=BlockStyle)
    layout: BlockLayout = Field(default_factory=BlockLayout)

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
    style: BlockStyle = Field(default_factory=BlockStyle)
    layout: BlockLayout = Field(default_factory=BlockLayout)

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


class AddPageBreakRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    document_id: str


class SectionPageBreakInput(BaseModel):
    model_config = ConfigDict(extra="forbid")

    type: Literal["page_break"] = "page_break"


class SectionBreakInput(BaseModel):
    model_config = ConfigDict(extra="forbid")

    type: Literal["section_break"] = "section_break"
    start_type: SectionStartType = "new_page"
    inherit_header_footer: bool = True
    page_orientation: Literal["portrait", "landscape"] | None = None
    margins: SectionMarginsConfig = Field(default_factory=SectionMarginsConfig)
    restart_page_numbering: bool = False
    page_number_start: int | None = Field(default=None, ge=1, le=9999)
    header_footer: HeaderFooterConfig = Field(default_factory=HeaderFooterConfig)

    @model_validator(mode="after")
    def validate_page_numbering(self) -> SectionBreakInput:
        validate_section_page_numbering(
            self.restart_page_numbering,
            self.page_number_start,
        )
        return self


class TocInput(BaseModel):
    model_config = ConfigDict(extra="forbid")

    type: Literal["toc"] = "toc"
    title: str = "目录"
    levels: int = Field(default=3, ge=1, le=6)
    start_on_new_page: bool = False
    style: BlockStyle = Field(default_factory=BlockStyle)
    layout: BlockLayout = Field(default_factory=BlockLayout)


class BlockGroupInput(BaseModel):
    model_config = ConfigDict(extra="forbid")

    type: Literal["group"] = "group"
    blocks: list[BlockInput] = Field(min_length=1)
    style: BlockStyle = Field(default_factory=BlockStyle)
    layout: BlockLayout = Field(default_factory=BlockLayout)


class BlockColumnInput(BaseModel):
    model_config = ConfigDict(extra="forbid")

    blocks: list[BlockInput] = Field(min_length=1)


class BlockColumnsInput(BaseModel):
    model_config = ConfigDict(extra="forbid")

    type: Literal["columns"] = "columns"
    columns: list[BlockColumnInput] = Field(min_length=1, max_length=3)
    style: BlockStyle = Field(default_factory=BlockStyle)
    layout: BlockLayout = Field(default_factory=BlockLayout)


BlockInput = Annotated[
    BlockHeadingInput
    | SectionParagraphInput
    | SectionListInput
    | SectionTableInput
    | SectionCardInput
    | SectionPageBreakInput
    | SectionBreakInput
    | TocInput
    | BlockGroupInput
    | BlockColumnsInput,
    Field(discriminator="type"),
]

SectionBundleBlock = BlockInput


class AddBlocksRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    document_id: str
    blocks: list[BlockInput] = Field(min_length=1)


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
    document_style: dict = Field(default_factory=dict)
    header_footer: dict = Field(default_factory=dict)


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
        document_style=document_model.metadata.document_style.model_dump(
            mode="json",
            exclude_none=True,
            exclude_defaults=True,
        ),
        header_footer=document_model.metadata.header_footer.model_dump(
            mode="json",
            exclude_none=True,
            exclude_defaults=True,
        ),
    )


BlockGroupInput.model_rebuild()
BlockColumnInput.model_rebuild()
BlockColumnsInput.model_rebuild()
