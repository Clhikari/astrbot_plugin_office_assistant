from collections.abc import Awaitable, Callable

from pydantic import ConfigDict, Field
from pydantic.dataclasses import dataclass

from astrbot import logger
from astrbot.core.agent.run_context import ContextWrapper
from astrbot.core.agent.tool import FunctionTool, ToolExecResult
from astrbot.core.astr_agent_context import AstrAgentContext

from ..document_core.builders.word_builder import WordDocumentBuilder
from ..export_pipeline import export_document_via_pipeline
from ..internal_hooks import AfterExportHook, BeforeExportHook
from ..mcp_server.schemas import (
    AddBlocksRequest,
    CreateDocumentRequest,
    ExportDocumentRequest,
    ExportDocumentResult,
    FinalizeDocumentRequest,
    ToolResult,
    build_document_summary,
)
from ..mcp_server.session_store import DocumentSessionStore


def _dump_result(result: ToolResult) -> str:
    return result.model_dump_json(exclude_none=True)


_CONTINUE_UNTIL_EXPORT = (
    "请继续调用文档工具，直到 export_document 成功。中途不要发自然语言回复。"
)
_FINALIZE_PROMPT = (
    "文档已定稿。请立即调用 export_document 导出文件，不要发自然语言回复。"
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
    "description": "Optional block layout tokens such as spacing_before and spacing_after.",
    "properties": {
        "spacing_before": {"type": "number"},
        "spacing_after": {"type": "number"},
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
        "Use this before adding headings, paragraphs, lists, tables, or summary cards."
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
            },
        }
    )

    async def call(
        self, context: ContextWrapper[AstrAgentContext], **kwargs
    ) -> ToolExecResult:
        session_id = kwargs.get("session_id")
        if not session_id and context is not None:
            event = getattr(getattr(context, "context", None), "event", None)
            session_id = getattr(event, "unified_msg_origin", "")
        request = CreateDocumentRequest(
            session_id=str(session_id or ""),
            title=str(kwargs.get("title") or ""),
            output_name=str(
                kwargs.get("output_name") or kwargs.get("title") or "document.docx"
            ),
            theme_name=str(kwargs.get("theme_name") or "business_report"),
            table_template=str(kwargs.get("table_template") or "report_grid"),
            density=str(kwargs.get("density") or "comfortable"),
            accent_color=str(kwargs.get("accent_color") or ""),
        )
        document = self.store.create_document(request)
        return _dump_result(
            ToolResult(
                success=True,
                message=f"文档会话已创建。下一步请调用 add_blocks 添加内容。{_CONTINUE_UNTIL_EXPORT}",
                document=build_document_summary(document),
            )
        )


@dataclass(config=ConfigDict(arbitrary_types_allowed=True))
class AddBlocksTool(DocumentToolBase):
    name: str = "add_blocks"
    description: str = (
        "Append one or more blocks in order. Use this for mixed content such as "
        "heading, paragraph, list, table, image, summary_card, page_break, group, or columns. "
        "For table blocks, if the user asks for a table title or 表格标题, put it in the table "
        "block's caption/title field so it renders as the first merged row inside the table, "
        "not as a separate heading block."
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
                    "description": "Ordered block list. Supported block types: heading, paragraph, list, table, image, summary_card, page_break, group, columns.",
                    "items": {
                        "type": "object",
                        "properties": {
                            "type": {"type": "string"},
                            "text": {"type": "string"},
                            "runs": {
                                "type": "array",
                                "description": "Optional inline rich-text runs for paragraph blocks.",
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "text": {"type": "string"},
                                        "bold": {"type": "boolean"},
                                        "italic": {"type": "boolean"},
                                        "underline": {"type": "boolean"},
                                        "code": {"type": "boolean"},
                                    },
                                    "required": ["text"],
                                },
                            },
                            "level": {"type": "number"},
                            "items": {
                                "type": "array",
                                "items": {"type": "string"},
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
                                    "items": {"type": "string"},
                                },
                            },
                            "header_groups": {
                                "type": "array",
                                "description": "Optional grouped header row. Each item defines a title and horizontal span.",
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
                            "caption": {
                                "type": "string",
                                "description": "Table title rendered as the first merged row inside the table.",
                            },
                            "title": {
                                "type": "string",
                                "description": "Optional paragraph card title, or table title alias for table blocks.",
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
                            "path": {"type": "string"},
                            "width_px": {"type": "number"},
                            "variant": {"type": "string"},
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
            request = AddBlocksRequest(
                document_id=str(kwargs.get("document_id") or ""),
                blocks=list(kwargs.get("blocks") or []),
            )
            document = self.store.add_blocks(request)
        except Exception as exc:
            return _dump_result(ToolResult(success=False, message=str(exc)))
        return _dump_result(
            ToolResult(
                success=True,
                message=f"内容块已添加。如果还有更多内容，继续调用 add_blocks；否则调用 finalize_document。{_CONTINUE_UNTIL_EXPORT}",
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
    builder: WordDocumentBuilder = Field(default_factory=WordDocumentBuilder)
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
            document, output_path = await export_document_via_pipeline(
                store=self.store,
                builder=self.builder,
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
                callback_message = (
                    await self.after_export(context, str(output_path)) or ""
                )
                delivery_handled = True
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
