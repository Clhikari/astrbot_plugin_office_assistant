from __future__ import annotations

from typing import Annotated, Literal
from urllib.parse import urlparse
from uuid import uuid4

from pydantic import (
    AnyHttpUrl,
    BaseModel,
    ConfigDict,
    Field,
    TypeAdapter,
    ValidationError as PydanticValidationError,
    field_validator,
    model_validator,
)

from ...shared_contracts import load_json_contract


HTTP_URL_ADAPTER = TypeAdapter(AnyHttpUrl)
_HYPERLINK_URL_CONTRACT = load_json_contract("hyperlink_url.json")
SUPPORTED_HYPERLINK_SCHEMES = tuple(_HYPERLINK_URL_CONTRACT["allowed_schemes"])
HYPERLINK_SCHEMES_REQUIRING_AUTHORITY = frozenset(
    _HYPERLINK_URL_CONTRACT["schemes_requiring_authority"]
)
HYPERLINK_SCHEMES_REQUIRING_PATH = frozenset(
    _HYPERLINK_URL_CONTRACT["schemes_requiring_path"]
)
HYPERLINK_URL_ERROR_MESSAGE = str(_HYPERLINK_URL_CONTRACT["error_message"])


class BlockLayout(BaseModel):
    model_config = ConfigDict(extra="forbid")

    spacing_before: float | None = Field(default=None, ge=0, le=72)
    spacing_after: float | None = Field(default=None, ge=0, le=72)
    padding_top_pt: float | None = Field(default=None, ge=0, le=72)
    padding_right_pt: float | None = Field(default=None, ge=0, le=72)
    padding_bottom_pt: float | None = Field(default=None, ge=0, le=72)
    padding_left_pt: float | None = Field(default=None, ge=0, le=72)


class BlockStyle(BaseModel):
    model_config = ConfigDict(extra="forbid")

    align: Literal["left", "center", "right", "justify"] | None = None
    emphasis: Literal["normal", "strong", "subtle"] | None = None
    font_scale: float | None = Field(default=None, ge=0.75, le=2.0)
    table_grid: Literal["report_grid", "metrics_compact", "minimal"] | None = None
    cell_align: Literal["left", "center", "right"] | None = None


PageNumberAlignment = Literal["left", "center", "right"]
PageNumberFormat = Literal[
    "decimal",
    "upperRoman",
    "lowerRoman",
    "upperLetter",
    "lowerLetter",
]
PageOrientation = Literal["portrait", "landscape"]


class SectionMarginsConfig(BaseModel):
    model_config = ConfigDict(extra="forbid")

    top_cm: float | None = Field(default=None, gt=0, le=10)
    bottom_cm: float | None = Field(default=None, gt=0, le=10)
    left_cm: float | None = Field(default=None, gt=0, le=10)
    right_cm: float | None = Field(default=None, gt=0, le=10)


def validate_section_page_numbering(
    restart_page_numbering: bool, page_number_start: int | None
) -> None:
    if page_number_start is not None and not restart_page_numbering:
        raise ValueError("page_number_start requires restart_page_numbering=True")


def _model_field_was_set(model: BaseModel, field_name: str) -> bool:
    return field_name in getattr(model, "model_fields_set", set())


class HeaderFooterConfig(BaseModel):
    model_config = ConfigDict(extra="forbid")

    header_text: str = ""
    footer_text: str = ""
    header_left: str = ""
    header_right: str = ""
    footer_left: str = ""
    footer_right: str = ""
    header_border_bottom: bool = False
    footer_border_top: bool = False
    header_border_color: str | None = None
    footer_border_color: str | None = None
    different_first_page: bool = False
    first_page_header_text: str = ""
    first_page_footer_text: str = ""
    first_page_show_page_number: bool | None = None
    different_odd_even: bool = False
    even_page_header_text: str = ""
    even_page_footer_text: str = ""
    even_page_show_page_number: bool | None = None
    show_page_number: bool | None = None
    page_number_align: PageNumberAlignment = "right"
    page_number_format: PageNumberFormat | None = None

    @field_validator("header_border_color", "footer_border_color")
    @classmethod
    def validate_optional_border_color(cls, value: str | None) -> str | None:
        return normalize_optional_hex_color(value)

    def has_explicit_overrides(self) -> bool:
        return any(
            _model_field_was_set(self, field_name)
            for field_name in type(self).model_fields
        )

    def merged_over(self, base_config: HeaderFooterConfig) -> HeaderFooterConfig:
        merged_config = base_config.model_copy(deep=True)
        for field_name in type(self).model_fields:
            if _model_field_was_set(self, field_name):
                setattr(merged_config, field_name, getattr(self, field_name))
        return merged_config


class BlockBase(BaseModel):
    model_config = ConfigDict(extra="forbid")

    block_id: str = Field(default_factory=lambda: uuid4().hex)
    type: str
    style: BlockStyle = Field(default_factory=BlockStyle)
    layout: BlockLayout = Field(default_factory=BlockLayout)


class HeadingBlock(BlockBase):
    type: Literal["heading"] = "heading"
    text: str = Field(min_length=1)
    level: int = Field(default=1, ge=1, le=6)
    bottom_border: bool | None = None
    bottom_border_color: str | None = None
    bottom_border_size_pt: float | None = Field(default=None, gt=0, le=6)

    @field_validator("bottom_border_color")
    @classmethod
    def validate_bottom_border_color(cls, value: str | None) -> str | None:
        return normalize_optional_hex_color(value)


class HeroBannerBlock(BlockBase):
    type: Literal["hero_banner"] = "hero_banner"
    title: str = Field(min_length=1)
    subtitle: str = ""
    theme_color: str | None = None
    text_color: str | None = None
    subtitle_color: str | None = None
    min_height_pt: float | None = Field(default=None, gt=0, le=240)
    full_width: bool = True

    @field_validator("theme_color", "text_color", "subtitle_color")
    @classmethod
    def validate_optional_colors(cls, value: str | None) -> str | None:
        return normalize_optional_hex_color(value)


PageTemplateName = Literal["business_review_cover", "technical_resume"]


class PageTemplateMetricItem(BaseModel):
    model_config = ConfigDict(extra="forbid")

    label: str = Field(min_length=1)
    value: str = Field(min_length=1)
    delta: str = ""
    delta_color: str | None = None
    note: str = ""

    @field_validator("delta_color")
    @classmethod
    def validate_optional_color(cls, value: str | None) -> str | None:
        return normalize_optional_hex_color(value)


class BusinessReviewCoverData(BaseModel):
    model_config = ConfigDict(extra="forbid")

    title: str = Field(min_length=1)
    subtitle: str = ""
    summary_title: str = "核心摘要"
    summary_text: str = Field(min_length=1)
    metrics: list[PageTemplateMetricItem] = Field(min_length=1, max_length=4)
    footer_note: str = ""
    auto_page_break: bool = False


class ResumeSectionEntry(BaseModel):
    model_config = ConfigDict(extra="forbid")

    heading: str = Field(min_length=1)
    date: str = ""
    subtitle: str = ""
    details: list[str | ListItem] = Field(default_factory=list)


class ResumeSection(BaseModel):
    model_config = ConfigDict(extra="forbid")

    title: str = Field(min_length=1)
    entries: list[ResumeSectionEntry] = Field(default_factory=list)
    lines: list[str | ListItem] = Field(default_factory=list)

    @model_validator(mode="after")
    def validate_content(self) -> ResumeSection:
        if self.entries or self.lines:
            return self
        raise ValueError("resume section requires entries or lines")


class TechnicalResumeData(BaseModel):
    model_config = ConfigDict(extra="forbid")

    name: str = Field(min_length=1)
    headline: str = ""
    contact_line: str = Field(min_length=1)
    sections: list[ResumeSection] = Field(min_length=1)


class PageTemplateBlock(BaseModel):
    model_config = ConfigDict(extra="forbid")

    block_id: str = Field(default_factory=lambda: uuid4().hex)
    type: Literal["page_template"] = "page_template"
    template: PageTemplateName
    data: BusinessReviewCoverData | TechnicalResumeData

    @model_validator(mode="after")
    def validate_template_data(self) -> PageTemplateBlock:
        expected_type = (
            BusinessReviewCoverData
            if self.template == "business_review_cover"
            else TechnicalResumeData
        )
        if isinstance(self.data, expected_type):
            return self
        raise ValueError(f"page_template data does not match template {self.template!r}")


class ParagraphRun(BaseModel):
    model_config = ConfigDict(extra="forbid")

    text: str = Field(min_length=1)
    bold: bool = False
    italic: bool = False
    underline: bool = False
    code: bool = False
    color: str | None = None
    url: str | None = None

    @field_validator("color")
    @classmethod
    def validate_color(cls, value: str | None) -> str | None:
        return normalize_optional_hex_color(value)

    @field_validator("url")
    @classmethod
    def validate_url(cls, value: str | None) -> str | None:
        return normalize_optional_hyperlink_url(value)


class ParagraphBlock(BlockBase):
    type: Literal["paragraph"] = "paragraph"
    text: str = ""
    variant: Literal["body", "summary_box", "key_takeaway"] = "body"
    title: str = ""
    runs: list[ParagraphRun] = Field(default_factory=list)

    @model_validator(mode="after")
    def validate_content(self) -> ParagraphBlock:
        if self.text.strip() or self.runs:
            return self
        raise ValueError("paragraph requires text or runs")


class ListItem(BaseModel):
    model_config = ConfigDict(extra="forbid")

    text: str = ""
    runs: list[ParagraphRun] = Field(default_factory=list)

    @model_validator(mode="after")
    def validate_content(self) -> ListItem:
        if self.text.strip() or self.runs:
            return self
        raise ValueError("list item requires text or runs")


class ListBlock(BlockBase):
    type: Literal["list"] = "list"
    items: list[str | ListItem] = Field(min_length=1)
    ordered: bool = False

    @field_validator("items")
    @classmethod
    def validate_items(cls, value: list[str | ListItem]) -> list[str | ListItem]:
        cleaned: list[str | ListItem] = []
        for item in value:
            if isinstance(item, str):
                text = item.strip()
                if text:
                    cleaned.append(text)
                continue
            if item.text.strip():
                item.text = item.text.strip()
            cleaned.append(item)
        if not cleaned:
            raise ValueError("items must contain at least one non-empty item")
        return cleaned


class TableHeaderGroup(BaseModel):
    model_config = ConfigDict(extra="forbid")

    title: str = Field(min_length=1)
    span: int = Field(ge=1)


TableAlignment = Literal["left", "center"]
TableBorderStyle = Literal["minimal", "standard", "strong"]
TableCaptionEmphasis = Literal["normal", "strong"]


def normalize_optional_hex_color(value: str | None) -> str | None:
    candidate = str(value or "").strip().lstrip("#").upper()
    if not candidate:
        return None
    if len(candidate) != 6 or any(char not in "0123456789ABCDEF" for char in candidate):
        raise ValueError("must be a 6-digit hex color")
    return candidate


def normalize_optional_hyperlink_url(value: str | None) -> str | None:
    candidate = str(value or "").strip()
    if not candidate:
        return None

    parsed = urlparse(candidate)
    scheme = parsed.scheme.lower()
    if scheme not in SUPPORTED_HYPERLINK_SCHEMES:
        raise ValueError(HYPERLINK_URL_ERROR_MESSAGE)
    if scheme in HYPERLINK_SCHEMES_REQUIRING_AUTHORITY:
        try:
            HTTP_URL_ADAPTER.validate_python(candidate)
        except PydanticValidationError as exc:
            raise ValueError(HYPERLINK_URL_ERROR_MESSAGE) from exc
    if scheme in HYPERLINK_SCHEMES_REQUIRING_PATH and not parsed.path:
        raise ValueError(HYPERLINK_URL_ERROR_MESSAGE)
    return candidate


class TableCell(BaseModel):
    model_config = ConfigDict(extra="forbid")

    text: str = ""
    row_span: int = Field(default=1, ge=1)
    col_span: int = Field(default=1, ge=1)
    fill: str | None = None
    text_color: str | None = None
    bold: bool | None = None
    align: Literal["left", "center", "right"] | None = None
    font_scale: float | None = Field(default=None, ge=0.5, le=3.0)

    @field_validator("fill", "text_color")
    @classmethod
    def validate_optional_colors(cls, value: str | None) -> str | None:
        return normalize_optional_hex_color(value)

    @model_validator(mode="after")
    def validate_merge_spans(self) -> TableCell:
        if self.row_span > 1 and self.col_span > 1:
            raise ValueError("table cell cannot combine row_span and col_span")
        return self


def resolve_table_cell_text(value: str | TableCell) -> str:
    return value if isinstance(value, str) else value.text


def resolve_table_cell_row_span(value: str | TableCell) -> int:
    return 1 if isinstance(value, str) else value.row_span


def resolve_table_cell_col_span(value: str | TableCell) -> int:
    return 1 if isinstance(value, str) else value.col_span


def is_empty_table_cell_placeholder(value: str | TableCell) -> bool:
    if isinstance(value, str):
        return value.strip() == ""
    return value.row_span == 1 and value.col_span == 1 and value.text.strip() == ""


def _resolve_table_body_column_count(
    rows: list[list[str | TableCell]],
    *,
    explicit_column_count: int | None = None,
) -> int:
    active_spans = [0] * explicit_column_count if explicit_column_count is not None else []
    max_columns = explicit_column_count or 0

    for row_index, row in enumerate(rows, start=1):
        next_active_spans = [max(span - 1, 0) for span in active_spans]
        column_index = 0
        for cell in row:
            consumed_placeholder = False
            while column_index < len(active_spans) and active_spans[column_index] > 0:
                if is_empty_table_cell_placeholder(cell):
                    column_index += 1
                    consumed_placeholder = True
                    break
                column_index += 1
            if consumed_placeholder:
                continue
            if explicit_column_count is not None and column_index >= explicit_column_count:
                raise ValueError(
                    f"table row {row_index} exceeds column count ({explicit_column_count})"
                )
            while len(active_spans) <= column_index:
                active_spans.append(0)
                next_active_spans.append(0)
            row_span = resolve_table_cell_row_span(cell)
            col_span = resolve_table_cell_col_span(cell)
            if explicit_column_count is not None and (
                column_index + col_span > explicit_column_count
            ):
                raise ValueError(
                    f"table row {row_index} exceeds column count ({explicit_column_count})"
                )
            while len(active_spans) < column_index + col_span:
                active_spans.append(0)
                next_active_spans.append(0)
            if any(
                active_spans[span_index] > 0
                for span_index in range(column_index, column_index + col_span)
            ):
                raise ValueError(f"table row {row_index} overlaps active row spans")
            if row_span > 1:
                for span_index in range(column_index, column_index + col_span):
                    next_active_spans[span_index] = max(
                        next_active_spans[span_index],
                        row_span - 1,
                    )
            column_index += col_span
        while column_index < len(active_spans) and active_spans[column_index] > 0:
            column_index += 1
        if explicit_column_count is not None and column_index < explicit_column_count:
            raise ValueError(
                f"table row {row_index} is missing cells "
                f"(expected {explicit_column_count}, got {column_index})"
            )
        max_columns = max(max_columns, len(active_spans), len(next_active_spans))
        active_spans = next_active_spans

    return max_columns


def resolve_table_column_count(headers: list[str], rows: list[list[str | TableCell]]) -> int:
    if headers:
        return len(headers)
    if rows:
        return _resolve_table_body_column_count(rows)
    return 0


def validate_table_structure(
    headers: list[str],
    rows: list[list[str | TableCell]],
    header_groups: list[TableHeaderGroup] | None = None,
) -> None:
    if headers:
        _resolve_table_body_column_count(rows, explicit_column_count=len(headers))
    column_count = resolve_table_column_count(headers, rows)
    if column_count <= 0:
        if header_groups:
            raise ValueError(
                "header_groups require at least one column from headers or rows "
                f"(column_count={column_count})"
            )
        return
    if not header_groups:
        return
    total_span = sum(group.span for group in header_groups)
    if total_span != column_count:
        raise ValueError(
            f"header_groups span total ({total_span}) must equal column count "
            f"({column_count})"
        )


class TableBlock(BlockBase):
    type: Literal["table"] = "table"
    headers: list[str] = Field(default_factory=list)
    rows: list[list[str | TableCell]] = Field(default_factory=list)
    header_groups: list[TableHeaderGroup] = Field(default_factory=list)
    table_style: Literal["report_grid", "metrics_compact", "minimal"] = "report_grid"
    caption: str = ""
    column_widths: list[float] = Field(default_factory=list)
    numeric_columns: list[int] = Field(default_factory=list)
    header_fill: str | None = None
    header_fill_enabled: bool | None = None
    header_text_color: str | None = None
    header_bold: bool | None = None
    banded_rows: bool | None = None
    banded_row_fill: str | None = None
    first_column_bold: bool | None = None
    table_align: TableAlignment | None = None
    border_style: TableBorderStyle | None = None
    caption_emphasis: TableCaptionEmphasis | None = None
    cell_padding_horizontal_pt: float | None = Field(default=None, ge=0, le=72)
    cell_padding_vertical_pt: float | None = Field(default=None, ge=0, le=72)
    header_font_scale: float | None = Field(default=None, ge=0.5, le=3.0)
    body_font_scale: float | None = Field(default=None, ge=0.5, le=3.0)

    @field_validator("column_widths")
    @classmethod
    def validate_column_widths(cls, value: list[float]) -> list[float]:
        return [width if width > 0 else 0 for width in value]

    @field_validator("numeric_columns")
    @classmethod
    def validate_numeric_columns(cls, value: list[int]) -> list[int]:
        cleaned = sorted({index for index in value if index >= 0})
        return cleaned

    @field_validator("header_fill", "header_text_color", "banded_row_fill")
    @classmethod
    def validate_optional_colors(cls, value: str | None) -> str | None:
        return normalize_optional_hex_color(value)

    @model_validator(mode="after")
    def validate_table_shape(self) -> TableBlock:
        validate_table_structure(self.headers, self.rows, self.header_groups)
        return self


class SummaryCardBlock(BlockBase):
    # Compatibility-only block. Writers may still emit it, but renderers should
    # expand it into standard primitives instead of treating it as a core
    # rendering branch.
    type: Literal["summary_card"] = "summary_card"
    title: str = Field(min_length=1)
    items: list[str] = Field(min_length=1)
    variant: Literal["summary", "conclusion"] = "summary"


class AccentBoxBlock(BlockBase):
    type: Literal["accent_box"] = "accent_box"
    title: str = ""
    text: str = ""
    runs: list[ParagraphRun] = Field(default_factory=list)
    items: list[str | ListItem] = Field(default_factory=list)
    accent_color: str | None = None
    fill_color: str | None = None
    title_color: str | None = None
    border_color: str | None = None
    border_width_pt: float | None = Field(default=None, gt=0, le=12)
    accent_border_width_pt: float | None = Field(default=None, gt=0, le=18)
    padding_pt: float | None = Field(default=None, ge=0, le=72)
    title_font_scale: float | None = Field(default=None, ge=0.5, le=3.0)
    body_font_scale: float | None = Field(default=None, ge=0.5, le=3.0)

    @field_validator("accent_color", "fill_color", "title_color", "border_color")
    @classmethod
    def validate_optional_colors(cls, value: str | None) -> str | None:
        return normalize_optional_hex_color(value)

    @model_validator(mode="after")
    def validate_content(self) -> AccentBoxBlock:
        if self.title.strip() or self.text.strip() or self.runs or self.items:
            return self
        raise ValueError("accent_box requires title, text, runs, or items")


class MetricCard(BaseModel):
    model_config = ConfigDict(extra="forbid")

    label: str = Field(min_length=1)
    value: str = Field(min_length=1)
    delta: str = ""
    note: str = ""
    value_color: str | None = None
    delta_color: str | None = None
    fill_color: str | None = None
    label_color: str | None = None
    note_color: str | None = None
    value_font_scale: float | None = Field(default=None, ge=0.5, le=3.0)
    delta_font_scale: float | None = Field(default=None, ge=0.5, le=3.0)

    @field_validator(
        "value_color",
        "delta_color",
        "fill_color",
        "label_color",
        "note_color",
    )
    @classmethod
    def validate_optional_colors(cls, value: str | None) -> str | None:
        return normalize_optional_hex_color(value)


class MetricCardsBlock(BlockBase):
    type: Literal["metric_cards"] = "metric_cards"
    metrics: list[MetricCard] = Field(min_length=1, max_length=4)
    accent_color: str | None = None
    fill_color: str | None = None
    label_color: str | None = None
    border_color: str | None = None
    border_width_pt: float | None = Field(default=None, gt=0, le=12)
    divider_color: str | None = None
    divider_width_pt: float | None = Field(default=None, gt=0, le=12)
    padding_pt: float | None = Field(default=None, ge=0, le=72)
    label_font_scale: float | None = Field(default=None, ge=0.5, le=3.0)
    value_font_scale: float | None = Field(default=None, ge=0.5, le=3.0)
    delta_font_scale: float | None = Field(default=None, ge=0.5, le=3.0)
    note_font_scale: float | None = Field(default=None, ge=0.5, le=3.0)

    @field_validator(
        "accent_color",
        "fill_color",
        "label_color",
        "border_color",
        "divider_color",
    )
    @classmethod
    def validate_optional_colors(cls, value: str | None) -> str | None:
        return normalize_optional_hex_color(value)


class ImageBlock(BlockBase):
    type: Literal["image"] = "image"
    path: str = Field(min_length=1)
    caption: str = ""
    width_px: int | None = Field(default=None, gt=0)


class GroupBlock(BlockBase):
    type: Literal["group"] = "group"
    blocks: list[DocumentBlock] = Field(min_length=1)


class ColumnBlock(BaseModel):
    model_config = ConfigDict(extra="forbid")

    blocks: list[DocumentBlock] = Field(min_length=1)


class ColumnsBlock(BlockBase):
    type: Literal["columns"] = "columns"
    columns: list[ColumnBlock] = Field(min_length=1, max_length=3)


class PageBreakBlock(BlockBase):
    type: Literal["page_break"] = "page_break"


SectionStartType = Literal[
    "new_page",
    "continuous",
    "odd_page",
    "even_page",
    "new_column",
]


class SectionBreakBlock(BlockBase):
    type: Literal["section_break"] = "section_break"
    start_type: SectionStartType = "new_page"
    inherit_header_footer: bool = True
    page_orientation: PageOrientation | None = None
    margins: SectionMarginsConfig = Field(default_factory=SectionMarginsConfig)
    restart_page_numbering: bool = False
    page_number_start: int | None = Field(default=None, ge=1, le=9999)
    header_footer: HeaderFooterConfig = Field(default_factory=HeaderFooterConfig)

    @model_validator(mode="after")
    def validate_page_numbering(self) -> SectionBreakBlock:
        validate_section_page_numbering(
            self.restart_page_numbering,
            self.page_number_start,
        )
        return self


class TocBlock(BlockBase):
    type: Literal["toc"] = "toc"
    title: str = "目录"
    levels: int = Field(default=3, ge=1, le=6)
    start_on_new_page: bool = False


DocumentBlock = Annotated[
    PageTemplateBlock
    | HeroBannerBlock
    | HeadingBlock
    | ParagraphBlock
    | ListBlock
    | TableBlock
    | SummaryCardBlock
    | AccentBoxBlock
    | MetricCardsBlock
    | ImageBlock
    | GroupBlock
    | ColumnsBlock
    | PageBreakBlock
    | SectionBreakBlock
    | TocBlock,
    Field(discriminator="type"),
]

GroupBlock.model_rebuild()
ColumnBlock.model_rebuild()
ColumnsBlock.model_rebuild()
