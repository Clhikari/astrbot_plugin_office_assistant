from collections.abc import Awaitable, Callable
from typing import Any

from pydantic import ConfigDict, Field
from pydantic.dataclasses import dataclass

from astrbot import logger
from astrbot.core.agent.run_context import ContextWrapper
from astrbot.core.agent.tool import FunctionTool, ToolExecResult
from astrbot.core.astr_agent_context import AstrAgentContext

from ..domain.document.export_pipeline import export_document_via_pipeline
from ..domain.document.hooks import AfterExportHook, BeforeExportHook
from ..domain.document.render_backends import (
    build_document_render_backends,
    get_render_backend_config,
)
from ..domain.document.session_store import DocumentSessionStore
from ..domain.document.contracts import (
    AddBlocksRequest,
    CreateDocumentRequest,
    ExportDocumentRequest,
    ExportDocumentResult,
    FinalizeDocumentRequest,
    ToolResult,
    build_document_summary,
    build_header_footer_schema,
    normalize_create_document_kwargs,
    normalize_raw_block_payloads,
)


def _dump_result(result: ToolResult) -> str:
    return result.model_dump_json(exclude_none=True)


_CONTINUE_UNTIL_EXPORT = (
    "请继续调用文档工具，直到 export_document 成功。中途不要发自然语言回复。"
)
_FINALIZE_PROMPT = (
    "文档已定稿。下一步只能调用 export_document 导出文件，不要再调用 add_blocks、create_document 或 finalize_document，也不要发自然语言回复。"
)


_STYLE_SCHEMA = {
    "type": "object",
    "description": "Optional block style tokens such as align, emphasis, font_scale, table_grid, or cell_align.",
    "properties": {
        "align": {"type": "string"},
        "emphasis": {"type": "string"},
        "font_scale": {"type": "number"},
        "table_grid": {"type": "string"},
        "cell_align": {"type": "string"},
    },
}

_LAYOUT_SCHEMA = {
    "type": "object",
    "description": "Optional block layout tokens such as spacing_before, spacing_after, and container padding.",
    "properties": {
        "spacing_before": {"type": "number"},
        "spacing_after": {"type": "number"},
        "padding_top_pt": {"type": "number"},
        "padding_right_pt": {"type": "number"},
        "padding_bottom_pt": {"type": "number"},
        "padding_left_pt": {"type": "number"},
    },
}

_GENERIC_BLOCK_ITEM_SCHEMA = {
    "type": "object",
    "additionalProperties": True,
}

_COLUMN_ITEM_SCHEMA = {
    "type": "object",
    "additionalProperties": True,
    "properties": {
        "blocks": {
            "type": "array",
            "items": _GENERIC_BLOCK_ITEM_SCHEMA,
        }
    },
}


@dataclass(config=ConfigDict(arbitrary_types_allowed=True))
class DocumentToolBase(FunctionTool[AstrAgentContext]):
    store: DocumentSessionStore = Field(default_factory=DocumentSessionStore)


@dataclass(config=ConfigDict(arbitrary_types_allowed=True))
class CreateDocumentTool(DocumentToolBase):
    name: str = "create_document"
    description: str = (
        "Create a draft Word document session and return its document_id. "
        "Use this before adding hero banners, headings, paragraphs, accent boxes, metric cards, lists, tables, or summary cards. "
        "For whole-document styling, put document-wide defaults in document_style and keep "
        "table block fields for local overrides. document_style.table_defaults does not "
        "accept table-only flags such as header_fill_enabled or header_bold."
    )
    parameters: dict = Field(
        default_factory=lambda: {
            "type": "object",
            "properties": {
                "session_id": {
                    "type": "string",
                    "description": "Optional session identifier. Defaults to the current chat session.",
                },
                "title": {
                    "type": "string",
                    "description": "Optional document title.",
                },
                "output_name": {
                    "type": "string",
                    "description": "Preferred output filename. .docx will be appended if omitted.",
                },
                "theme_name": {
                    "type": "string",
                    "description": "Document theme preset, e.g. business_report, project_review, or executive_brief.",
                },
                "table_template": {
                    "type": "string",
                    "description": "Default table style preset, e.g. report_grid, metrics_compact, or minimal.",
                },
                "density": {
                    "type": "string",
                    "description": "Document density preset, use comfortable or compact.",
                },
                "accent_color": {
                    "type": "string",
                    "description": "Optional 6-digit hex accent override such as 1F4E79.",
                },
                "header_footer": build_header_footer_schema(
                    description="Optional whole-document header and footer settings."
                ),
                "document_style": {
                    "type": "object",
                    "description": "Optional whole-document style defaults, including high-level visual intent and table defaults.",
                    "properties": {
                        "brief": {
                            "type": "string",
                            "description": "Optional natural-language style brief stored as document metadata. If the user describes visual intent in plain language, convert it into explicit document_style fields when possible.",
                        },
                        "heading_color": {
                            "type": "string",
                            "description": "Optional 6-digit hex color for the document title and heading text. Default is 000000; natural-language requests like 黑色标题 or 深蓝标题 should be mapped here.",
                        },
                        "heading_level_1_color": {
                            "type": "string",
                            "description": "Optional 6-digit hex color override for level-1 headings.",
                        },
                        "heading_level_2_color": {
                            "type": "string",
                            "description": "Optional 6-digit hex color override for level-2 headings.",
                        },
                        "heading_bottom_border_color": {
                            "type": "string",
                            "description": "Optional 6-digit hex default color for heading divider lines.",
                        },
                        "heading_bottom_border_size_pt": {
                            "type": "number",
                            "description": "Optional default thickness in points for heading divider lines.",
                        },
                        "title_align": {
                            "type": "string",
                            "enum": ["left", "center", "right", "justify"],
                            "description": "Optional alignment for the document title.",
                        },
                        "body_font_size": {
                            "type": "number",
                            "description": "Optional base font size for body paragraphs and lists.",
                        },
                        "body_line_spacing": {
                            "type": "number",
                            "description": "Optional body paragraph line spacing multiplier.",
                        },
                        "font_name": {
                            "type": "string",
                            "description": "Optional default body font. Default is Microsoft YaHei.",
                        },
                        "heading_font_name": {
                            "type": "string",
                            "description": "Optional default heading font. Default follows font_name.",
                        },
                        "table_font_name": {
                            "type": "string",
                            "description": "Optional default table font. Default follows font_name.",
                        },
                        "code_font_name": {
                            "type": "string",
                            "description": "Optional default code run font. Default is Consolas.",
                        },
                        "paragraph_space_after": {
                            "type": "number",
                            "description": "Optional default spacing after body paragraphs.",
                        },
                        "list_space_after": {
                            "type": "number",
                            "description": "Optional default spacing after list items.",
                        },
                        "summary_card_defaults": {
                            "type": "object",
                            "description": "Optional defaults for summary_card, summary_box, and key_takeaway content.",
                            "properties": {
                                "title_align": {
                                    "type": "string",
                                    "enum": ["left", "center", "right", "justify"],
                                    "description": "Optional alignment default for summary card titles.",
                                },
                                "title_emphasis": {
                                    "type": "string",
                                    "enum": ["normal", "strong", "subtle"],
                                    "description": "Optional emphasis default for summary card titles.",
                                },
                                "title_font_scale": {
                                    "type": "number",
                                    "description": "Optional font scale default for summary card titles.",
                                },
                                "title_space_before": {
                                    "type": "number",
                                    "description": "Optional spacing before summary card titles.",
                                },
                                "title_space_after": {
                                    "type": "number",
                                    "description": "Optional spacing after summary card titles.",
                                },
                                "list_space_after": {
                                    "type": "number",
                                    "description": "Optional spacing after summary card list items.",
                                },
                            },
                        },
                        "table_defaults": {
                            "type": "object",
                            "description": "Optional default table styling applied unless a table block overrides it. Use this only for shared defaults; table-only flags like header_fill_enabled and header_bold belong on each table block.",
                            "properties": {
                                "preset": {
                                    "type": "string",
                                    "enum": [
                                        "report_grid",
                                        "metrics_compact",
                                        "minimal",
                                    ],
                                    "description": "Optional default table preset applied when a table block does not set table_style.",
                                },
                                "header_fill": {
                                    "type": "string",
                                    "description": "Optional 6-digit hex color for default table header backgrounds.",
                                },
                                "body_fill": {
                                    "type": "string",
                                    "description": "Optional 6-digit hex color for default table body cell backgrounds.",
                                },
                                "header_text_color": {
                                    "type": "string",
                                    "description": "Optional 6-digit hex color for default table header text.",
                                },
                                "banded_rows": {
                                    "type": "boolean",
                                    "description": "Optional flag to enable alternating row fills by default.",
                                },
                                "banded_row_fill": {
                                    "type": "string",
                                    "description": "Optional 6-digit hex color for default alternating table rows.",
                                },
                                "first_column_bold": {
                                    "type": "boolean",
                                    "description": "Optional flag to emphasize the first table column by default.",
                                },
                                "table_align": {
                                    "type": "string",
                                    "enum": ["left", "center"],
                                    "description": "Optional default alignment for whole tables.",
                                },
                                "border_style": {
                                    "type": "string",
                                    "enum": ["minimal", "standard", "strong"],
                                    "description": "Optional default border weight preset for tables.",
                                },
                                "caption_emphasis": {
                                    "type": "string",
                                    "enum": ["normal", "strong"],
                                    "description": "Optional default emphasis preset for merged table caption rows.",
                                },
                                "cell_align": {
                                    "type": "string",
                                    "enum": ["left", "center", "right"],
                                    "description": "Optional default paragraph alignment for table body cells.",
                                },
                            },
                        },
                    },
                },
            },
        }
    )

    async def call(
        self, context: ContextWrapper[AstrAgentContext], **kwargs
    ) -> ToolExecResult:
        normalized_kwargs = normalize_create_document_kwargs(kwargs)
        session_id = kwargs.get("session_id")
        if not session_id and context is not None:
            event = getattr(getattr(context, "context", None), "event", None)
            session_id = getattr(event, "unified_msg_origin", "")
        request = CreateDocumentRequest(
            session_id=str(session_id or ""),
            title=str(normalized_kwargs.get("title") or ""),
            output_name=str(
                normalized_kwargs.get("output_name")
                or normalized_kwargs.get("title")
                or "document.docx"
            ),
            theme_name=str(normalized_kwargs.get("theme_name") or "business_report"),
            table_template=str(
                normalized_kwargs.get("table_template") or "report_grid"
            ),
            density=str(normalized_kwargs.get("density") or "comfortable"),
            accent_color=str(normalized_kwargs.get("accent_color") or ""),
            header_footer=dict(normalized_kwargs.get("header_footer") or {}),
            document_style=dict(normalized_kwargs.get("document_style") or {}),
        )
        document = self.store.create_document(request)
        return _dump_result(
            ToolResult(
                success=True,
                message=(
                    "文档会话已创建。下一步只能调用 add_blocks 添加内容，"
                    "不要提前调用 finalize_document 或 export_document。"
                    f"{_CONTINUE_UNTIL_EXPORT}"
                ),
                document=build_document_summary(document),
            )
        )


@dataclass(config=ConfigDict(arbitrary_types_allowed=True))
class AddBlocksTool(DocumentToolBase):
    name: str = "add_blocks"
    description: str = (
        "Append one or more blocks in order. Use this for mixed content such as "
        "page_template, hero_banner, heading, paragraph, accent_box, metric_cards, list, table, image, summary_card, page_break, section_break, toc, group, or columns. "
        "For table blocks, if the user asks for a table title or 表格标题, put it in the table "
        "block's caption/title field so it renders as the first merged row inside the table, "
        "not as a separate heading block. For table styling, use table-specific fields like "
        "header_fill, header_text_color, header_bold, banded_rows, first_column_bold, table_align, "
        "border_style, caption_emphasis, and row_span when the user requests a custom visual style. "
        "header_fill_enabled and header_bold are table block fields, not document_style.table_defaults fields."
    )
    parameters: dict = Field(
        default_factory=lambda: {
            "type": "object",
            "properties": {
                "document_id": {
                    "type": "string",
                    "description": "Target document_id returned by create_document.",
                },
                "blocks": {
                    "type": "array",
                    "description": "Ordered block list. Supported block types: page_template, hero_banner, heading, paragraph, accent_box, metric_cards, list, table, image, summary_card, page_break, section_break, toc, group, columns.",
                    "items": {
                        "type": "object",
                        "properties": {
                            "type": {"type": "string"},
                            "template": {
                                "type": "string",
                                "description": "Template name for page_template blocks. Supported values include business_review_cover and technical_resume.",
                            },
                            "data": {
                                "type": "object",
                                "description": "Template payload for page_template blocks.",
                                "properties": {
                                    "title": {
                                        "type": "string",
                                        "description": "Primary title for the template page.",
                                    },
                                    "subtitle": {
                                        "type": "string",
                                        "description": "Optional subtitle for the template page.",
                                    },
                                    "summary_title": {
                                        "type": "string",
                                        "description": "Optional summary box title. Default is 核心摘要.",
                                    },
                                    "summary_text": {
                                        "type": "string",
                                        "description": "Main summary paragraph for the template page.",
                                    },
                                    "metrics": {
                                        "type": "array",
                                        "description": "Metric items rendered inside the template page. First version supports 1 to 4 items.",
                                        "items": {
                                            "type": "object",
                                            "properties": {
                                                "label": {"type": "string"},
                                                "value": {"type": "string"},
                                                "delta": {"type": "string"},
                                                "delta_color": {"type": "string"},
                                                "note": {"type": "string"},
                                            },
                                            "required": ["label", "value"],
                                        },
                                    },
                                    "footer_note": {
                                        "type": "string",
                                        "description": "Optional footer note on cover. DO NOT USE THIS for document '编制人/日期'. To add '编制人' information, you MUST append a right-aligned paragraph block with color 595959 as the VERY LAST block of the entire document.",
                                    },
                                    "auto_page_break": {
                                        "type": "boolean",
                                        "description": "Whether the template page should end with an automatic page break.",
                                    },
                                    "name": {
                                        "type": "string",
                                        "description": "Resume owner name for technical_resume.",
                                    },
                                    "headline": {
                                        "type": "string",
                                        "description": "Centered resume headline under the name.",
                                    },
                                    "contact_line": {
                                        "type": "string",
                                        "description": "Single centered contact line rendered above the first divider.",
                                    },
                                    "sections": {
                                        "type": "array",
                                        "description": "Resume sections for technical_resume. Each section needs a title plus entries or lines.",
                                        "items": {
                                            "type": "object",
                                            "properties": {
                                                "title": {"type": "string"},
                                                "entries": {
                                                    "type": "array",
                                                    "items": {
                                                        "type": "object",
                                                        "properties": {
                                                            "heading": {"type": "string"},
                                                            "date": {"type": "string"},
                                                            "subtitle": {"type": "string"},
                                                            "details": {
                                                                "type": "array",
                                                                "items": {
                                                                    "anyOf": [
                                                                        {"type": "string"},
                                                                        {
                                                                            "type": "object",
                                                                            "properties": {
                                                                                "text": {"type": "string"},
                                                                                "runs": {
                                                                                    "type": "array",
                                                                                    "items": {
                                                                                        "type": "object",
                                                                                        "properties": {
                                                                                            "text": {"type": "string"},
                                                                                            "bold": {"type": "boolean"},
                                                                                            "italic": {"type": "boolean"},
                                                                                            "underline": {"type": "boolean"},
                                                                                            "code": {"type": "boolean"},
                                                                                            "color": {"type": "string"},
                                                                                        },
                                                                                        "required": ["text"],
                                                                                    },
                                                                                },
                                                                            },
                                                                        },
                                                                    ]
                                                                },
                                                            },
                                                        },
                                                        "required": ["heading"],
                                                    },
                                                },
                                                "lines": {
                                                    "type": "array",
                                                    "items": {
                                                        "anyOf": [
                                                            {"type": "string"},
                                                            {
                                                                "type": "object",
                                                                "properties": {
                                                                    "text": {"type": "string"},
                                                                    "runs": {
                                                                        "type": "array",
                                                                        "items": {
                                                                            "type": "object",
                                                                            "properties": {
                                                                                "text": {"type": "string"},
                                                                                "bold": {"type": "boolean"},
                                                                                "italic": {"type": "boolean"},
                                                                                "underline": {"type": "boolean"},
                                                                                "code": {"type": "boolean"},
                                                                                "color": {"type": "string"},
                                                                            },
                                                                            "required": ["text"],
                                                                        },
                                                                    },
                                                                },
                                                            },
                                                        ]
                                                    },
                                                },
                                            },
                                        },
                                    },
                                },
                            },
                            "text": {"type": "string"},
                            "subtitle": {"type": "string"},
                            "runs": {
                                "type": "array",
                                "description": "Optional inline rich-text runs for paragraph or list item content.",
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "text": {"type": "string"},
                                        "bold": {"type": "boolean"},
                                        "italic": {"type": "boolean"},
                                        "underline": {"type": "boolean"},
                                        "code": {"type": "boolean"},
                                        "color": {
                                            "type": "string",
                                            "description": "Optional 6-digit hex color for this run, such as 666666.",
                                        },
                                    },
                                    "required": ["text"],
                                },
                            },
                            "level": {"type": "number"},
                            "bottom_border": {"type": "boolean"},
                            "bottom_border_color": {
                                "type": "string",
                                "description": "Optional 6-digit hex color for a heading bottom border, such as D0D7DE.",
                            },
                            "bottom_border_size_pt": {
                                "type": "number",
                                "description": "Optional heading bottom border thickness in points.",
                            },
                            "items": {
                                "type": "array",
                                "items": {
                                    "anyOf": [
                                        {"type": "string"},
                                        {
                                            "type": "object",
                                            "properties": {
                                                "text": {"type": "string"},
                                                "runs": {
                                                    "type": "array",
                                                    "items": {
                                                        "type": "object",
                                                        "properties": {
                                                            "text": {"type": "string"},
                                                            "bold": {"type": "boolean"},
                                                            "italic": {"type": "boolean"},
                                                            "underline": {"type": "boolean"},
                                                            "code": {"type": "boolean"},
                                                            "color": {"type": "string"},
                                                        },
                                                        "required": ["text"],
                                                    },
                                                },
                                            },
                                        },
                                    ]
                                },
                            },
                            "ordered": {"type": "boolean"},
                            "headers": {
                                "type": "array",
                                "items": {"type": "string"},
                            },
                            "rows": {
                                "type": "array",
                                "items": {
                                    "type": "array",
                                    "items": {
                                        "anyOf": [
                                            {"type": "string"},
                                            {
                                                "type": "object",
                                                "properties": {
                                                    "text": {"type": "string"},
                                                    "row_span": {
                                                        "type": "integer",
                                                        "minimum": 1,
                                                        "description": "Optional row span for vertically merged body cells.",
                                                    },
                                                    "fill": {
                                                        "type": "string",
                                                        "description": "Optional 6-digit hex fill color for this body cell.",
                                                    },
                                                    "text_color": {
                                                        "type": "string",
                                                        "description": "Optional 6-digit hex text color for this body cell.",
                                                    },
                                                    "bold": {
                                                        "type": "boolean",
                                                        "description": "Optional bold override for this body cell.",
                                                    },
                                                    "align": {
                                                        "type": "string",
                                                        "enum": ["left", "center", "right"],
                                                        "description": "Optional alignment override for this body cell.",
                                                    },
                                                    "font_scale": {
                                                        "type": "number",
                                                        "description": "Optional body cell font scale override.",
                                                    },
                                                },
                                            },
                                        ]
                                    },
                                },
                            },
                            "header_groups": {
                                "type": "array",
                                "description": "Optional grouped header definitions. Each item defines a title and horizontal span.",
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "title": {
                                            "type": "string",
                                            "minLength": 1,
                                        },
                                        "span": {
                                            "type": "integer",
                                            "minimum": 1,
                                        },
                                    },
                                    "required": ["title", "span"],
                                },
                            },
                            "table_style": {"type": "string"},
                            "header_fill": {
                                "type": "string",
                                "description": "Optional 6-digit hex color for the header row background, such as 1F4E79.",
                            },
                            "header_fill_enabled": {
                                "type": "boolean",
                                "description": "Set false to keep the header row white or unfilled.",
                            },
                            "header_text_color": {
                                "type": "string",
                                "description": "Optional 6-digit hex color for the header row text, such as FFFFFF.",
                            },
                            "header_bold": {
                                "type": "boolean",
                                "description": "Optional flag to control whether header text is bold.",
                            },
                            "banded_rows": {
                                "type": "boolean",
                                "description": "Optional flag to alternate row fills in the table body.",
                            },
                            "banded_row_fill": {
                                "type": "string",
                                "description": "Optional 6-digit hex color for alternating table body rows.",
                            },
                            "first_column_bold": {
                                "type": "boolean",
                                "description": "Optional flag to emphasize the first data column with bold text.",
                            },
                            "table_align": {
                                "type": "string",
                                "enum": ["left", "center"],
                                "description": "Optional table alignment override.",
                            },
                            "border_style": {
                                "type": "string",
                                "enum": ["minimal", "standard", "strong"],
                                "description": "Optional border weight preset for the table.",
                            },
                            "caption_emphasis": {
                                "type": "string",
                                "enum": ["normal", "strong"],
                                "description": "Optional emphasis preset for the merged caption row.",
                            },
                            "caption": {
                                "type": "string",
                                "description": "Table title rendered as the first merged row inside the table.",
                            },
                            "title": {
                                "type": "string",
                                "description": "Optional paragraph card title, or table title alias for table blocks.",
                            },
                            "theme_color": {
                                "type": "string",
                                "description": "Optional 6-digit hex background color for hero_banner blocks.",
                            },
                            "text_color": {
                                "type": "string",
                                "description": "Optional 6-digit hex title color for hero_banner blocks.",
                            },
                            "subtitle_color": {
                                "type": "string",
                                "description": "Optional 6-digit hex subtitle color for hero_banner blocks.",
                            },
                            "min_height_pt": {
                                "type": "number",
                                "description": "Optional minimum banner height in points.",
                            },
                            "full_width": {
                                "type": "boolean",
                                "description": "Whether hero_banner should stretch across the page width.",
                            },
                            "accent_color": {
                                "type": "string",
                                "description": "Optional 6-digit hex accent color for accent_box or metric_cards blocks.",
                            },
                            "fill_color": {
                                "type": "string",
                                "description": "Optional 6-digit hex fill color for accent_box or metric_cards blocks.",
                            },
                            "title_color": {
                                "type": "string",
                                "description": "Optional 6-digit hex title color for accent_box blocks.",
                            },
                            "border_color": {
                                "type": "string",
                                "description": "Optional 6-digit hex border color for accent_box or metric_cards blocks.",
                            },
                            "border_width_pt": {
                                "type": "number",
                                "description": "Optional border width in points for accent_box or metric_cards blocks.",
                            },
                            "accent_border_width_pt": {
                                "type": "number",
                                "description": "Optional left accent border width in points for accent_box blocks.",
                            },
                            "divider_color": {
                                "type": "string",
                                "description": "Optional 6-digit hex divider color for metric_cards blocks.",
                            },
                            "divider_width_pt": {
                                "type": "number",
                                "description": "Optional divider width in points for metric_cards blocks.",
                            },
                            "padding_pt": {
                                "type": "number",
                                "description": "Optional inner padding in points for accent_box or metric_cards blocks.",
                            },
                            "title_font_scale": {
                                "type": "number",
                                "description": "Optional title font scale for accent_box blocks.",
                            },
                            "body_font_scale": {
                                "type": "number",
                                "description": "Optional body font scale for accent_box or table blocks.",
                            },
                            "metrics": {
                                "type": "array",
                                "description": "Metric card definitions for metric_cards blocks.",
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "label": {"type": "string"},
                                        "value": {"type": "string"},
                                        "delta": {"type": "string"},
                                        "note": {"type": "string"},
                                        "value_color": {"type": "string"},
                                        "delta_color": {"type": "string"},
                                        "fill_color": {"type": "string"},
                                        "label_color": {"type": "string"},
                                        "note_color": {"type": "string"},
                                        "value_font_scale": {"type": "number"},
                                        "delta_font_scale": {"type": "number"},
                                    },
                                    "required": ["label", "value"],
                                },
                            },
                            "label_color": {
                                "type": "string",
                                "description": "Optional 6-digit hex label color for metric_cards blocks.",
                            },
                            "label_font_scale": {
                                "type": "number",
                                "description": "Optional default label font scale for metric_cards blocks.",
                            },
                            "value_font_scale": {
                                "type": "number",
                                "description": "Optional default value font scale for metric_cards blocks.",
                            },
                            "delta_font_scale": {
                                "type": "number",
                                "description": "Optional default delta font scale for metric_cards blocks.",
                            },
                            "note_font_scale": {
                                "type": "number",
                                "description": "Optional default note font scale for metric_cards blocks.",
                            },
                            "column_widths": {
                                "type": "array",
                                "description": "Optional table column widths in centimeters.",
                                "items": {"type": "number"},
                            },
                            "numeric_columns": {
                                "type": "array",
                                "description": "Optional zero-based column indexes that should be right-aligned for numeric values.",
                                "items": {"type": "integer"},
                            },
                            "cell_padding_horizontal_pt": {
                                "type": "number",
                                "description": "Optional horizontal cell padding for table blocks in points.",
                            },
                            "cell_padding_vertical_pt": {
                                "type": "number",
                                "description": "Optional vertical cell padding for table blocks in points.",
                            },
                            "header_font_scale": {
                                "type": "number",
                                "description": "Optional header font scale for table blocks.",
                            },
                            "path": {"type": "string"},
                            "width_px": {"type": "number"},
                            "variant": {"type": "string"},
                            "start_type": {
                                "type": "string",
                                "enum": [
                                    "new_page",
                                    "continuous",
                                    "odd_page",
                                    "even_page",
                                    "new_column",
                                ],
                            },
                            "inherit_header_footer": {"type": "boolean"},
                            "page_orientation": {
                                "type": "string",
                                "enum": ["portrait", "landscape"],
                            },
                            "margins": {
                                "type": "object",
                                "properties": {
                                    "top_cm": {"type": "number"},
                                    "bottom_cm": {"type": "number"},
                                    "left_cm": {"type": "number"},
                                    "right_cm": {"type": "number"},
                                },
                            },
                            "restart_page_numbering": {"type": "boolean"},
                            "page_number_start": {
                                "type": "integer",
                                "minimum": 1,
                                "description": "Optional starting page number. Use only when restart_page_numbering is true.",
                            },
                            "header_footer": build_header_footer_schema(
                                description="Optional section-level header and footer overrides."
                            ),
                            "levels": {
                                "type": "integer",
                                "minimum": 1,
                                "maximum": 6,
                            },
                            "start_on_new_page": {"type": "boolean"},
                            "blocks": {
                                "type": "array",
                                "items": _GENERIC_BLOCK_ITEM_SCHEMA,
                            },
                            "columns": {
                                "type": "array",
                                "items": _COLUMN_ITEM_SCHEMA,
                            },
                            "style": _STYLE_SCHEMA,
                            "layout": _LAYOUT_SCHEMA,
                        },
                        "required": ["type"],
                    },
                },
            },
            "required": ["document_id", "blocks"],
        }
    )

    async def call(
        self, context: ContextWrapper[AstrAgentContext], **kwargs
    ) -> ToolExecResult:
        try:
            raw_blocks = normalize_raw_block_payloads(list(kwargs.get("blocks") or []))
            request = AddBlocksRequest(
                document_id=str(kwargs.get("document_id") or ""),
                blocks=raw_blocks,
            )
            document = self.store.add_blocks(request)
        except Exception as exc:
            return _dump_result(
                ToolResult(
                    success=False,
                    message=(
                        "add_blocks 失败。继续使用同一个 document_id 再次调用 "
                        "add_blocks，只修正报错字段，不要改调 finalize_document "
                        f"或 export_document。原始错误：{exc}"
                    ),
                )
            )
        return _dump_result(
            ToolResult(
                success=True,
                message=(
                    "内容块已添加。如果还有任何章节、表格或补充信息没写完，继续调用 "
                    "add_blocks；只有确认全部内容写完，才调用 finalize_document。"
                    "定稿前不要调用 export_document。"
                    f"{_CONTINUE_UNTIL_EXPORT}"
                ),
                document=build_document_summary(document),
            )
        )


@dataclass(config=ConfigDict(arbitrary_types_allowed=True))
class FinalizeDocumentTool(DocumentToolBase):
    name: str = "finalize_document"
    description: str = "Mark a document draft as finalized before export."
    parameters: dict = Field(
        default_factory=lambda: {
            "type": "object",
            "properties": {
                "document_id": {
                    "type": "string",
                    "description": "Target document_id returned by create_document.",
                }
            },
            "required": ["document_id"],
        }
    )

    async def call(
        self, context: ContextWrapper[AstrAgentContext], **kwargs
    ) -> ToolExecResult:
        try:
            request = FinalizeDocumentRequest(
                document_id=str(kwargs.get("document_id") or "")
            )
            document = self.store.finalize_document(request)
        except Exception as exc:
            return _dump_result(ToolResult(success=False, message=str(exc)))
        return _dump_result(
            ToolResult(
                success=True,
                message=_FINALIZE_PROMPT,
                document=build_document_summary(document),
            )
        )


@dataclass(config=ConfigDict(arbitrary_types_allowed=True))
class ExportDocumentTool(DocumentToolBase):
    name: str = "export_document"
    description: str = (
        "Export the current Word draft to a .docx file and return the file path."
    )
    parameters: dict = Field(
        default_factory=lambda: {
            "type": "object",
            "properties": {
                "document_id": {
                    "type": "string",
                    "description": "Target document_id returned by create_document.",
                },
                "output_dir": {
                    "type": "string",
                    "description": "Optional output directory. Defaults to the plugin workspace.",
                },
                "output_name": {
                    "type": "string",
                    "description": "Optional output filename.",
                },
            },
            "required": ["document_id"],
        }
    )
    render_backends: list[Any] = Field(
        default_factory=lambda: build_document_render_backends("word")
    )
    render_backend_config: Any | None = None
    before_export_hooks: list[BeforeExportHook] = Field(default_factory=list)
    after_export_hooks: list[AfterExportHook] = Field(default_factory=list)
    after_export: (
        Callable[[ContextWrapper[AstrAgentContext], str], Awaitable[str | None]] | None
    ) = None

    async def call(
        self, context: ContextWrapper[AstrAgentContext], **kwargs
    ) -> ToolExecResult:
        try:
            request = ExportDocumentRequest(
                document_id=str(kwargs.get("document_id") or ""),
                output_dir=str(kwargs.get("output_dir") or ""),
                output_name=str(kwargs.get("output_name") or ""),
            )
            document_for_routing = self.store.require_document(request.document_id)
            resolved_render_backends = (
                build_document_render_backends(
                    document_for_routing.format,
                    self.render_backend_config
                    or get_render_backend_config(self.store),
                )
                if self.render_backend_config is not None
                or document_for_routing.format != "word"
                else self.render_backends
            )
            document, output_path = await export_document_via_pipeline(
                store=self.store,
                render_backends=resolved_render_backends,
                request=request,
                before_export_hooks=self.before_export_hooks,
                after_export_hooks=self.after_export_hooks,
                source="agent_tool",
            )
        except Exception as exc:
            return _dump_result(ToolResult(success=False, message=str(exc)))

        callback_message = ""
        delivery_handled = False
        if self.after_export is not None and context is not None:
            try:
                logger.debug(
                    "[office-assistant] invoking after_export callback for document=%s output=%s",
                    document.document_id,
                    output_path,
                )
                callback_message = (
                    await self.after_export(context, str(output_path)) or ""
                )
                delivery_handled = True
                logger.debug(
                    "[office-assistant] after_export callback completed for document=%s output=%s delivered=%s",
                    document.document_id,
                    output_path,
                    delivery_handled,
                )
            except Exception as exc:
                logger.warning(
                    "[office-assistant] after_export callback failed for %s: %s",
                    output_path,
                    exc,
                )
                callback_message = (
                    f"Document exported, but post-export delivery failed: {exc}"
                )
        if delivery_handled:
            # The exported file has already been delivered to the user by the callback.
            # Returning None makes the tool loop stop instead of prompting the model
            # to continue with send_message_to_user or other follow-up tools.
            return None
        return _dump_result(
            ExportDocumentResult(
                success=True,
                message=callback_message or "Document exported.",
                document=build_document_summary(document),
                file_path=str(output_path),
            )
        )


__all__ = [
    "AddBlocksTool",
    "CreateDocumentTool",
    "ExportDocumentTool",
    "FinalizeDocumentTool",
]
