from __future__ import annotations

import copy
import re
from collections.abc import Mapping
from pathlib import Path
from typing import Annotated, Literal

from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator

from ...constants import (
    DOCUMENT_BLOCK_FONT_SCALE_MAX,
    DOCUMENT_BLOCK_FONT_SCALE_MIN,
    DOCUMENT_BLOCK_SPACING_MAX,
    DOCUMENT_BLOCK_SPACING_MIN,
)
from ...document_core.models.blocks import (
    BlockLayout,
    BlockStyle,
    BorderConfig,
    HeaderFooterConfig,
    ListItem,
    ParagraphRun,
    SectionMarginsConfig,
    SectionStartType,
    TableAlignment,
    TableBorderStyle,
    TableCaptionEmphasis,
    TableCell,
    TableHeaderGroup,
    normalize_optional_hex_color,
    validate_section_page_numbering,
    validate_table_structure,
)
from ...document_core.models.document import DocumentModel, DocumentStyleConfig

SUPPORTED_THEMES = {"business_report", "project_review", "executive_brief"}
SUPPORTED_TABLE_TEMPLATES = {"report_grid", "metrics_compact", "minimal"}
SUPPORTED_DENSITIES = {"comfortable", "compact"}
SUPPORTED_CARD_VARIANTS = {"summary", "conclusion"}
SUPPORTED_PAGE_TEMPLATES = {"business_review_cover", "technical_resume"}
WINDOWS_DRIVE_PATTERN = re.compile(r"^[A-Za-z]:([\\/]|$)")
DEFAULT_DOCX_FILENAME = "document.docx"
MAX_BLOCK_NORMALIZE_DEPTH = 32

_HEADER_FOOTER_SCHEMA_PROPERTIES = {
    "header_text": {
        "type": "string",
        "description": "Optional repeated header text for the document.",
    },
    "footer_text": {
        "type": "string",
        "description": "Optional repeated footer text for the document.",
    },
    "header_left": {
        "type": "string",
        "description": "Optional left-aligned header content.",
    },
    "header_right": {
        "type": "string",
        "description": "Optional right-aligned header content. Use {PAGE} if page number text should appear here.",
    },
    "footer_left": {
        "type": "string",
        "description": "Optional left-aligned footer content.",
    },
    "footer_right": {
        "type": "string",
        "description": "Optional right-aligned footer content. Use {PAGE} if page number text should appear here.",
    },
    "header_border_bottom": {
        "type": "boolean",
        "description": "Whether to draw a bottom divider line under the header.",
    },
    "footer_border_top": {
        "type": "boolean",
        "description": "Whether to draw a top divider line above the footer.",
    },
    "header_border_color": {
        "type": "string",
        "description": "Optional 6-digit hex color for the header divider line.",
    },
    "footer_border_color": {
        "type": "string",
        "description": "Optional 6-digit hex color for the footer divider line.",
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
    "page_number_format": {
        "type": "string",
        "enum": [
            "decimal",
            "upperRoman",
            "lowerRoman",
            "upperLetter",
            "lowerLetter",
        ],
        "description": "Optional page number format for this section, such as upperRoman.",
    },
}


def build_header_footer_schema(*, description: str) -> dict:
    return {
        "type": "object",
        "description": description,
        "properties": copy.deepcopy(_HEADER_FOOTER_SCHEMA_PROPERTIES),
    }


_SECTION_BREAK_COMPAT_KEYS = (
    "start_type",
    "inherit_header_footer",
    "page_orientation",
    "margins",
    "restart_page_numbering",
    "page_number_start",
    "header_footer",
)
_SECTION_BREAK_COMPAT_BLOCK_TYPES = {"heading", "paragraph", "table"}


def _copy_raw_block(block: object) -> object:
    return copy.deepcopy(block)


def _extract_text_from_column_alias(value: object) -> str:
    if isinstance(value, str):
        return value.strip()
    if not isinstance(value, Mapping):
        return ""
    if isinstance(value.get("text"), str) and value["text"].strip():
        return value["text"].strip()
    nested_blocks = value.get("blocks")
    if isinstance(nested_blocks, list):
        for child in nested_blocks:
            text = _extract_text_from_column_alias(child)
            if text:
                return text
    return ""


def _unwrap_singleton_list(value: object) -> object:
    current = value
    while isinstance(current, list) and len(current) == 1:
        current = current[0]
    return current


def _normalize_table_header_alias(block: dict) -> None:
    if block.get("type") != "table":
        return
    raw_headers = block.get("headers")
    if isinstance(raw_headers, list) and raw_headers:
        normalized_headers = [
            text
            for text in (_extract_text_from_column_alias(header) for header in raw_headers)
            if text
        ]
        if normalized_headers:
            block["headers"] = normalized_headers
            block.pop("columns", None)
            return
    raw_columns = block.get("columns")
    if not isinstance(raw_columns, list):
        return
    headers = [
        text
        for text in (_extract_text_from_column_alias(column) for column in raw_columns)
        if text
    ]
    if headers:
        block["headers"] = headers
        block.pop("columns", None)


def _normalize_table_row_alias(block: dict) -> None:
    if block.get("type") != "table":
        return
    raw_rows = block.get("rows")
    if isinstance(raw_rows, list):
        normalized_rows: list[object] = []
        changed = False
        for row in raw_rows:
            if isinstance(row, Mapping) and isinstance(row.get("cells"), list):
                normalized_rows.append(copy.deepcopy(row["cells"]))
                changed = True
                continue
            if isinstance(row, list):
                normalized_row = [_unwrap_singleton_list(cell) for cell in row]
                normalized_rows.append(normalized_row)
                changed = changed or normalized_row != row
                continue
            normalized_rows.append(row)
        if changed:
            block["rows"] = normalized_rows
    raw_items = block.pop("items", None)
    if raw_items is None or block.get("rows"):
        return
    if not isinstance(raw_items, list):
        return

    normalized_rows: list[list[str]] = []
    for item in raw_items:
        if isinstance(item, list):
            normalized_rows.append([str(cell).strip() for cell in item])
            continue
        if not isinstance(item, str):
            continue
        text = item.strip()
        if not text:
            continue
        if "|" in text:
            normalized_rows.append(
                [cell.strip() for cell in text.strip("|").split("|")]
            )
        else:
            normalized_rows.append([text])

    if normalized_rows:
        block["rows"] = normalized_rows


def _normalize_block_style_and_layout(block: dict) -> None:
    style = block.get("style")
    if isinstance(style, dict):
        font_scale = style.get("font_scale")
        if isinstance(font_scale, (int, float)):
            style["font_scale"] = min(
                max(float(font_scale), DOCUMENT_BLOCK_FONT_SCALE_MIN),
                DOCUMENT_BLOCK_FONT_SCALE_MAX,
            )

    layout = block.get("layout")
    if isinstance(layout, dict):
        for field_name in ("spacing_before", "spacing_after"):
            value = layout.get(field_name)
            if isinstance(value, (int, float)):
                layout[field_name] = min(
                    max(float(value), DOCUMENT_BLOCK_SPACING_MIN),
                    DOCUMENT_BLOCK_SPACING_MAX,
                )


_NESTED_BLOCK_ALIAS_KEYS = (
    "hero_banner",
    "accent_box",
    "metric_cards",
    "heading",
    "table",
    "paragraph",
)


def _merge_missing_mapping_fields(target: dict, key: str, value: object) -> None:
    if not isinstance(value, Mapping):
        if key not in target:
            target[key] = copy.deepcopy(value)
        return
    existing = target.get(key)
    if isinstance(existing, Mapping):
        merged = dict(value)
        merged.update(existing)
        target[key] = merged
        return
    if key not in target:
        target[key] = copy.deepcopy(value)


def _normalize_nested_block_payload_alias(block: dict) -> None:
    block_type = block.get("type")
    if not isinstance(block_type, str) or not block_type.strip():
        for alias_key in _NESTED_BLOCK_ALIAS_KEYS:
            if isinstance(block.get(alias_key), Mapping):
                block["type"] = alias_key
                block_type = alias_key
                break
    if not isinstance(block_type, str) or not block_type.strip():
        return

    nested_payload = block.get(block_type)
    if isinstance(nested_payload, Mapping):
        for key, value in nested_payload.items():
            if key == "type":
                continue
            _merge_missing_mapping_fields(block, key, value)
        block.pop(block_type, None)

    if block_type == "accent_box":
        content = block.pop("content", None)
        if isinstance(content, str) and content.strip() and not block.get("text"):
            block["text"] = content.strip()
    elif block_type == "hero_banner":
        title_color = block.pop("title_color", None)
        if isinstance(title_color, str) and title_color.strip() and not block.get("text_color"):
            block["text_color"] = title_color.strip()
        text = block.pop("text", None)
        if isinstance(text, str) and text.strip() and not block.get("title"):
            block["title"] = text.strip()
    elif block_type == "metric_cards":
        cards = block.pop("cards", None)
        if isinstance(cards, list) and not block.get("metrics"):
            block["metrics"] = copy.deepcopy(cards)
    elif block_type == "paragraph":
        layout = block.get("layout")
        if isinstance(layout, Mapping):
            normalized_layout = dict(layout)
            alignment = normalized_layout.pop("alignment", None)
            if isinstance(alignment, str) and alignment.strip():
                style = dict(block.get("style") or {})
                style.setdefault("align", alignment.strip())
                block["style"] = style
            block["layout"] = normalized_layout


def _drop_unsupported_block_aliases(block: dict) -> None:
    if block.get("type") == "heading":
        block.pop("heading_color", None)


_DOCUMENT_STYLE_COMPAT_KEYS = (
    "heading_color",
    "heading_level_1_color",
    "heading_level_2_color",
    "heading_bottom_border_color",
    "heading_bottom_border_size_pt",
    "title_align",
    "body_font_size",
    "body_line_spacing",
    "font_name",
    "heading_font_name",
    "table_font_name",
    "code_font_name",
    "paragraph_space_after",
    "list_space_after",
    "brief",
)


def _normalize_paragraph_items_alias(block: dict) -> None:
    if block.get("type") != "paragraph":
        return
    if block.get("text") or block.get("runs"):
        block.pop("items", None)
        return
    raw_items = block.get("items")
    if not isinstance(raw_items, list) or len(raw_items) != 1:
        return
    item = raw_items[0]
    if isinstance(item, str):
        text = item.strip()
        if text:
            block["text"] = text
            block.pop("items", None)
        return
    if not isinstance(item, Mapping):
        return
    item_runs = item.get("runs")
    if isinstance(item_runs, list) and item_runs:
        block["runs"] = item_runs
        block.pop("items", None)
        return
    item_text = item.get("text")
    if isinstance(item_text, str) and item_text.strip():
        block["text"] = item_text.strip()
        block.pop("items", None)


def _normalize_toc_title_alias(block: dict) -> None:
    if block.get("type") != "toc" or block.get("title"):
        return
    toc_text = block.pop("text", "")
    if isinstance(toc_text, str) and toc_text.strip():
        block["title"] = toc_text.strip()


def _normalize_heading_title_alias(block: dict) -> None:
    if block.get("type") != "heading":
        return
    heading_title = block.get("title")
    if block.get("text"):
        block.pop("title", None)
        return
    if isinstance(heading_title, str) and heading_title.strip():
        block["text"] = heading_title.strip()
    block.pop("title", None)


def _normalize_runs_color_aliases(runs: object) -> None:
    if not isinstance(runs, list):
        return
    for raw_run in runs:
        if not isinstance(raw_run, dict):
            continue
        text_color = raw_run.pop("text_color", None)
        if "color" not in raw_run and isinstance(text_color, str) and text_color.strip():
            raw_run["color"] = text_color.strip()


def _normalize_block_run_aliases(block: dict) -> None:
    _normalize_runs_color_aliases(block.get("runs"))
    items = block.get("items")
    if not isinstance(items, list):
        return
    normalized_items: list[object] = []
    changed = False
    for item in items:
        unwrapped_item = _unwrap_singleton_list(item)
        normalized_items.append(unwrapped_item)
        changed = changed or unwrapped_item != item
    if changed:
        block["items"] = normalized_items
        items = normalized_items
    for item in items:
        if not isinstance(item, dict):
            continue
        _normalize_runs_color_aliases(item.get("runs"))


def _normalize_table_cell_aliases(block: dict) -> None:
    if block.get("type") != "table":
        return
    rows = block.get("rows")
    if not isinstance(rows, list):
        return
    for row in rows:
        if not isinstance(row, list):
            continue
        for cell in row:
            if not isinstance(cell, dict):
                continue
            if cell.get("type") == "cell":
                cell.pop("type", None)
            _normalize_runs_color_aliases(cell.get("runs"))


def _normalize_legacy_paragraph_border_aliases(block: dict) -> None:
    if block.get("type") != "paragraph":
        return

    bottom_border = block.pop("bottom_border", None)
    bottom_border_color = block.pop("bottom_border_color", None)
    bottom_border_size_pt = block.pop("bottom_border_size_pt", None)
    bottom_border_style = block.pop("bottom_border_style", None)

    if bottom_border is False:
        return

    has_legacy_border = (
        bottom_border is True
        or bottom_border_color not in (None, "")
        or bottom_border_size_pt is not None
        or (
            isinstance(bottom_border_style, str)
            and bottom_border_style.strip() != ""
        )
    )
    if not has_legacy_border:
        return

    raw_border = block.get("border")
    border = dict(raw_border) if isinstance(raw_border, Mapping) else {}
    raw_bottom = border.get("bottom")
    bottom = dict(raw_bottom) if isinstance(raw_bottom, Mapping) else {}

    if (
        isinstance(bottom_border_style, str)
        and bottom_border_style.strip()
        and "style" not in bottom
    ):
        bottom["style"] = bottom_border_style.strip()
    if bottom_border_color not in (None, "") and "color" not in bottom:
        bottom["color"] = bottom_border_color
    if bottom_border_size_pt is not None and "width_pt" not in bottom:
        bottom["width_pt"] = bottom_border_size_pt

    border["bottom"] = bottom
    block["border"] = border


_BLOCK_SHAPE_NORMALIZERS = (
    _normalize_nested_block_payload_alias,
    _normalize_toc_title_alias,
)
_BLOCK_STYLE_NORMALIZERS = (
    _normalize_block_style_and_layout,
    _drop_unsupported_block_aliases,
)
_BLOCK_CONTENT_NORMALIZERS = (
    _normalize_heading_title_alias,
    _normalize_paragraph_items_alias,
    _normalize_block_run_aliases,
    _normalize_legacy_paragraph_border_aliases,
)
_TABLE_BLOCK_NORMALIZERS = (
    _normalize_table_header_alias,
    _normalize_table_row_alias,
    _normalize_table_cell_aliases,
)
_BLOCK_NORMALIZER_PIPELINE = (
    _BLOCK_SHAPE_NORMALIZERS,
    _BLOCK_STYLE_NORMALIZERS,
    _BLOCK_CONTENT_NORMALIZERS,
    _TABLE_BLOCK_NORMALIZERS,
)


def _apply_block_normalization_pipeline(block: dict) -> None:
    for normalizer_group in _BLOCK_NORMALIZER_PIPELINE:
        for normalizer in normalizer_group:
            normalizer(block)


def normalize_create_document_kwargs(kwargs: Mapping[str, object]) -> dict[str, object]:
    normalized = dict(kwargs)
    raw_document_style = normalized.get("document_style")
    document_style = dict(raw_document_style) if isinstance(raw_document_style, Mapping) else {}
    for key in _DOCUMENT_STYLE_COMPAT_KEYS:
        if key in normalized and key not in document_style:
            document_style[key] = normalized[key]
    if document_style:
        normalized["document_style"] = document_style
    return normalized


def _extract_compat_section_break(block: dict) -> dict | None:
    if block.get("type") not in _SECTION_BREAK_COMPAT_BLOCK_TYPES:
        return None

    section_break: dict[str, object] = {"type": "section_break"}
    has_section_fields = False

    if block.pop("start_on_new_page", False):
        section_break["start_type"] = "new_page"
        has_section_fields = True

    for key in _SECTION_BREAK_COMPAT_KEYS:
        if key not in block:
            continue
        value = block.pop(key)
        if value in (None, "", {}, []):
            continue
        section_break[key] = value
        has_section_fields = True

    return section_break if has_section_fields else None


def normalize_raw_block_payloads(
    blocks: list[object],
    *,
    _depth: int = 0,
) -> list[object]:
    if _depth > MAX_BLOCK_NORMALIZE_DEPTH:
        raise ValueError(
            f"block payload nesting exceeds limit ({MAX_BLOCK_NORMALIZE_DEPTH})"
        )
    normalized_blocks: list[object] = []
    for raw_block in blocks:
        block = _copy_raw_block(raw_block)
        if not isinstance(block, dict):
            normalized_blocks.append(block)
            continue

        block_type = block.get("type")
        block.pop("block_id", None)
        if block_type == "group" and isinstance(block.get("blocks"), list):
            block["blocks"] = normalize_raw_block_payloads(
                block["blocks"],
                _depth=_depth + 1,
            )
        elif block_type == "columns" and isinstance(block.get("columns"), list):
            normalized_columns: list[object] = []
            for column in block["columns"]:
                normalized_column = _copy_raw_block(column)
                if isinstance(normalized_column, dict) and isinstance(
                    normalized_column.get("blocks"), list
                ):
                    normalized_column["blocks"] = normalize_raw_block_payloads(
                        normalized_column["blocks"],
                        _depth=_depth + 1,
                    )
                normalized_columns.append(normalized_column)
            block["columns"] = normalized_columns

        _apply_block_normalization_pipeline(block)
        section_break = _extract_compat_section_break(block)
        if section_break is not None:
            normalized_blocks.append(section_break)
        normalized_blocks.append(block)

    return normalized_blocks


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


def _normalize_docx_filename(
    value: str,
    default: str = DEFAULT_DOCX_FILENAME,
) -> str:
    parts = _split_path_parts(value)
    candidate = parts[-1] if parts else default
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


def _coerce_table_input_rows_to_runtime(
    rows: list[list[str | TableCellInput]],
) -> list[list[str | TableCell]]:
    return [
        [
            cell
            if isinstance(cell, str)
            else TableCell(**cell.model_dump(mode="json", exclude_none=True))
            for cell in row
        ]
        for row in rows
    ]

class CreateDocumentRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    session_id: str = ""
    title: str = ""
    output_name: str = DEFAULT_DOCX_FILENAME
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
    bottom_border: bool | None = None
    bottom_border_color: str | None = None
    bottom_border_size_pt: float | None = Field(default=None, gt=0, le=6)
    style: BlockStyle = Field(default_factory=BlockStyle)
    layout: BlockLayout = Field(default_factory=BlockLayout)

    @field_validator("bottom_border_color")
    @classmethod
    def validate_bottom_border_color(cls, value: str | None) -> str | None:
        return normalize_optional_hex_color(value)


class BlockHeadingInput(BaseModel):
    model_config = ConfigDict(extra="forbid")

    type: Literal["heading"] = "heading"
    text: str = Field(min_length=1)
    level: int = Field(default=1, ge=1, le=6)
    bottom_border: bool | None = None
    bottom_border_color: str | None = None
    bottom_border_size_pt: float | None = Field(default=None, gt=0, le=6)
    style: BlockStyle = Field(default_factory=BlockStyle)
    layout: BlockLayout = Field(default_factory=BlockLayout)

    @field_validator("bottom_border_color")
    @classmethod
    def validate_bottom_border_color(cls, value: str | None) -> str | None:
        return normalize_optional_hex_color(value)


class ParagraphRunInput(ParagraphRun):
    model_config = ConfigDict(extra="forbid")


class ListItemInput(ListItem):
    model_config = ConfigDict(extra="forbid")


class TableCellInput(BaseModel):
    model_config = ConfigDict(extra="forbid")

    text: str = ""
    runs: list[ParagraphRunInput] = Field(default_factory=list)
    fill: str | None = None
    text_color: str | None = None
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    strikethrough: bool | None = None
    align: Literal["left", "center", "right"] | None = None
    font_name: str | None = None
    font_scale: float | None = Field(
        default=None,
        ge=DOCUMENT_BLOCK_FONT_SCALE_MIN,
        le=3.0,
    )
    border: BorderConfig | None = None

    @field_validator("fill", "text_color")
    @classmethod
    def validate_optional_colors(cls, value: str | None) -> str | None:
        return normalize_optional_hex_color(value)


class MetricCardInput(BaseModel):
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
    value_font_scale: float | None = Field(
        default=None,
        ge=DOCUMENT_BLOCK_FONT_SCALE_MIN,
        le=3.0,
    )
    delta_font_scale: float | None = Field(
        default=None,
        ge=DOCUMENT_BLOCK_FONT_SCALE_MIN,
        le=3.0,
    )

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


class AddParagraphRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    document_id: str
    text: str = ""
    variant: Literal["body", "summary_box", "key_takeaway"] = "body"
    title: str = ""
    runs: list[ParagraphRunInput] = Field(default_factory=list)
    border: BorderConfig | None = None
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
    border: BorderConfig | None = None
    style: BlockStyle = Field(default_factory=BlockStyle)
    layout: BlockLayout = Field(default_factory=BlockLayout)

    @model_validator(mode="after")
    def validate_content(self) -> SectionParagraphInput:
        if self.text.strip() or self.runs:
            return self
        raise ValueError("paragraph requires text or runs")


class AddAccentBoxRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    document_id: str
    title: str = ""
    text: str = ""
    runs: list[ParagraphRunInput] = Field(default_factory=list)
    items: list[str | ListItemInput] = Field(default_factory=list)
    accent_color: str | None = None
    fill_color: str | None = None
    title_color: str | None = None
    border_color: str | None = None
    border_width_pt: float | None = Field(default=None, gt=0, le=12)
    accent_border_width_pt: float | None = Field(default=None, gt=0, le=18)
    padding_pt: float | None = Field(default=None, ge=0, le=72)
    title_font_scale: float | None = Field(
        default=None,
        ge=DOCUMENT_BLOCK_FONT_SCALE_MIN,
        le=3.0,
    )
    body_font_scale: float | None = Field(
        default=None,
        ge=DOCUMENT_BLOCK_FONT_SCALE_MIN,
        le=3.0,
    )
    style: BlockStyle = Field(default_factory=BlockStyle)
    layout: BlockLayout = Field(default_factory=BlockLayout)

    @field_validator(
        "accent_color",
        "fill_color",
        "title_color",
        "border_color",
    )
    @classmethod
    def validate_optional_colors(cls, value: str | None) -> str | None:
        return normalize_optional_hex_color(value)

    @model_validator(mode="after")
    def validate_content(self) -> AddAccentBoxRequest:
        if self.title.strip() or self.text.strip() or self.runs or self.items:
            return self
        raise ValueError("accent_box requires title, text, runs, or items")


class SectionAccentBoxInput(BaseModel):
    model_config = ConfigDict(extra="forbid")

    type: Literal["accent_box"] = "accent_box"
    title: str = ""
    text: str = ""
    runs: list[ParagraphRunInput] = Field(default_factory=list)
    items: list[str | ListItemInput] = Field(default_factory=list)
    accent_color: str | None = None
    fill_color: str | None = None
    title_color: str | None = None
    border_color: str | None = None
    border_width_pt: float | None = Field(default=None, gt=0, le=12)
    accent_border_width_pt: float | None = Field(default=None, gt=0, le=18)
    padding_pt: float | None = Field(default=None, ge=0, le=72)
    title_font_scale: float | None = Field(
        default=None,
        ge=DOCUMENT_BLOCK_FONT_SCALE_MIN,
        le=3.0,
    )
    body_font_scale: float | None = Field(
        default=None,
        ge=DOCUMENT_BLOCK_FONT_SCALE_MIN,
        le=3.0,
    )
    style: BlockStyle = Field(default_factory=BlockStyle)
    layout: BlockLayout = Field(default_factory=BlockLayout)

    @field_validator(
        "accent_color",
        "fill_color",
        "title_color",
        "border_color",
    )
    @classmethod
    def validate_optional_colors(cls, value: str | None) -> str | None:
        return normalize_optional_hex_color(value)

    @model_validator(mode="after")
    def validate_content(self) -> SectionAccentBoxInput:
        if self.title.strip() or self.text.strip() or self.runs or self.items:
            return self
        raise ValueError("accent_box requires title, text, runs, or items")


class AddListRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    document_id: str
    items: list[str | ListItemInput] = Field(min_length=1)
    ordered: bool = False
    style: BlockStyle = Field(default_factory=BlockStyle)
    layout: BlockLayout = Field(default_factory=BlockLayout)

    @field_validator("items")
    @classmethod
    def validate_items(
        cls, value: list[str | ListItemInput]
    ) -> list[str | ListItemInput]:
        cleaned: list[str | ListItemInput] = []
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


class SectionListInput(BaseModel):
    model_config = ConfigDict(extra="forbid")

    type: Literal["list"] = "list"
    items: list[str | ListItemInput] = Field(min_length=1)
    ordered: bool = False
    style: BlockStyle = Field(default_factory=BlockStyle)
    layout: BlockLayout = Field(default_factory=BlockLayout)

    @field_validator("items")
    @classmethod
    def validate_items(
        cls, value: list[str | ListItemInput]
    ) -> list[str | ListItemInput]:
        cleaned: list[str | ListItemInput] = []
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


class AddTableRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    document_id: str
    headers: list[str] = Field(default_factory=list)
    rows: list[list[str | TableCellInput]] = Field(default_factory=list)
    header_groups: list[TableHeaderGroup] = Field(default_factory=list)
    table_style: str = ""
    caption: str = ""
    title: str = ""
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
    header_font_scale: float | None = Field(
        default=None,
        ge=DOCUMENT_BLOCK_FONT_SCALE_MIN,
        le=3.0,
    )
    body_font_scale: float | None = Field(
        default=None,
        ge=DOCUMENT_BLOCK_FONT_SCALE_MIN,
        le=3.0,
    )
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
        validate_table_structure(
            self.headers,
            _coerce_table_input_rows_to_runtime(self.rows),
            self.header_groups,
        )
        return self


class SectionTableInput(BaseModel):
    model_config = ConfigDict(extra="forbid")

    type: Literal["table"] = "table"
    headers: list[str] = Field(default_factory=list)
    rows: list[list[str | TableCellInput]] = Field(default_factory=list)
    header_groups: list[TableHeaderGroup] = Field(default_factory=list)
    table_style: str = ""
    caption: str = ""
    title: str = ""
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
    header_font_scale: float | None = Field(
        default=None,
        ge=DOCUMENT_BLOCK_FONT_SCALE_MIN,
        le=3.0,
    )
    body_font_scale: float | None = Field(
        default=None,
        ge=DOCUMENT_BLOCK_FONT_SCALE_MIN,
        le=3.0,
    )
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
        validate_table_structure(
            self.headers,
            _coerce_table_input_rows_to_runtime(self.rows),
            self.header_groups,
        )
        return self


class AddMetricCardsRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    document_id: str
    metrics: list[MetricCardInput] = Field(min_length=1, max_length=4)
    accent_color: str | None = None
    fill_color: str | None = None
    label_color: str | None = None
    border_color: str | None = None
    border_width_pt: float | None = Field(default=None, gt=0, le=12)
    divider_color: str | None = None
    divider_width_pt: float | None = Field(default=None, gt=0, le=12)
    padding_pt: float | None = Field(default=None, ge=0, le=72)
    label_font_scale: float | None = Field(
        default=None,
        ge=DOCUMENT_BLOCK_FONT_SCALE_MIN,
        le=3.0,
    )
    value_font_scale: float | None = Field(
        default=None,
        ge=DOCUMENT_BLOCK_FONT_SCALE_MIN,
        le=3.0,
    )
    delta_font_scale: float | None = Field(
        default=None,
        ge=DOCUMENT_BLOCK_FONT_SCALE_MIN,
        le=3.0,
    )
    note_font_scale: float | None = Field(
        default=None,
        ge=DOCUMENT_BLOCK_FONT_SCALE_MIN,
        le=3.0,
    )
    style: BlockStyle = Field(default_factory=BlockStyle)
    layout: BlockLayout = Field(default_factory=BlockLayout)

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


class SectionMetricCardsInput(BaseModel):
    model_config = ConfigDict(extra="forbid")

    type: Literal["metric_cards"] = "metric_cards"
    metrics: list[MetricCardInput] = Field(min_length=1, max_length=4)
    accent_color: str | None = None
    fill_color: str | None = None
    label_color: str | None = None
    border_color: str | None = None
    border_width_pt: float | None = Field(default=None, gt=0, le=12)
    divider_color: str | None = None
    divider_width_pt: float | None = Field(default=None, gt=0, le=12)
    padding_pt: float | None = Field(default=None, ge=0, le=72)
    label_font_scale: float | None = Field(
        default=None,
        ge=DOCUMENT_BLOCK_FONT_SCALE_MIN,
        le=3.0,
    )
    value_font_scale: float | None = Field(
        default=None,
        ge=DOCUMENT_BLOCK_FONT_SCALE_MIN,
        le=3.0,
    )
    delta_font_scale: float | None = Field(
        default=None,
        ge=DOCUMENT_BLOCK_FONT_SCALE_MIN,
        le=3.0,
    )
    note_font_scale: float | None = Field(
        default=None,
        ge=DOCUMENT_BLOCK_FONT_SCALE_MIN,
        le=3.0,
    )
    style: BlockStyle = Field(default_factory=BlockStyle)
    layout: BlockLayout = Field(default_factory=BlockLayout)

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


class SectionHeroBannerInput(BaseModel):
    model_config = ConfigDict(extra="forbid")

    type: Literal["hero_banner"] = "hero_banner"
    title: str = Field(min_length=1)
    subtitle: str = ""
    theme_color: str | None = None
    text_color: str | None = None
    subtitle_color: str | None = None
    min_height_pt: float | None = Field(default=None, gt=0, le=240)
    full_width: bool = True
    style: BlockStyle = Field(default_factory=BlockStyle)
    layout: BlockLayout = Field(default_factory=BlockLayout)

    @field_validator("theme_color", "text_color", "subtitle_color")
    @classmethod
    def validate_optional_colors(cls, value: str | None) -> str | None:
        return normalize_optional_hex_color(value)


class PageTemplateMetricInput(BaseModel):
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


class BusinessReviewCoverDataInput(BaseModel):
    model_config = ConfigDict(extra="forbid")

    title: str = Field(min_length=1)
    subtitle: str = ""
    summary_title: str = "核心摘要"
    summary_text: str = Field(min_length=1)
    metrics: list[PageTemplateMetricInput] = Field(min_length=1, max_length=4)
    footer_note: str = ""
    auto_page_break: bool = False


class ResumeSectionEntryInput(BaseModel):
    model_config = ConfigDict(extra="forbid")

    heading: str = Field(min_length=1)
    date: str = ""
    subtitle: str = ""
    details: list[str | ListItemInput] = Field(default_factory=list)


class ResumeSectionInput(BaseModel):
    model_config = ConfigDict(extra="forbid")

    title: str = Field(min_length=1)
    entries: list[ResumeSectionEntryInput] = Field(default_factory=list)
    lines: list[str | ListItemInput] = Field(default_factory=list)

    @model_validator(mode="after")
    def validate_content(self) -> ResumeSectionInput:
        if self.entries or self.lines:
            return self
        raise ValueError("resume section requires entries or lines")


class TechnicalResumeDataInput(BaseModel):
    model_config = ConfigDict(extra="forbid")

    name: str = Field(min_length=1)
    headline: str = ""
    contact_line: str = Field(min_length=1)
    sections: list[ResumeSectionInput] = Field(min_length=1)


class SectionPageTemplateInput(BaseModel):
    model_config = ConfigDict(extra="forbid")

    type: Literal["page_template"] = "page_template"
    template: Literal["business_review_cover", "technical_resume"] = (
        "business_review_cover"
    )
    data: BusinessReviewCoverDataInput | TechnicalResumeDataInput

    @field_validator("template")
    @classmethod
    def validate_template(cls, value: str) -> str:
        candidate = value.strip()
        if candidate not in SUPPORTED_PAGE_TEMPLATES:
            raise ValueError(
                f"unsupported page_template template: {candidate or value!r}"
            )
        return candidate

    @model_validator(mode="after")
    def validate_template_data(self) -> SectionPageTemplateInput:
        expected_type = (
            BusinessReviewCoverDataInput
            if self.template == "business_review_cover"
            else TechnicalResumeDataInput
        )
        if isinstance(self.data, expected_type):
            return self
        raise ValueError(f"page_template data does not match template {self.template!r}")


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
    SectionPageTemplateInput
    | SectionHeroBannerInput
    | BlockHeadingInput
    | SectionParagraphInput
    | SectionAccentBoxInput
    | SectionListInput
    | SectionTableInput
    | SectionMetricCardsInput
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
    model_config = ConfigDict(extra="forbid")
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
