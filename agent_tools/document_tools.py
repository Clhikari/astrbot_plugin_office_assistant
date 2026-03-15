from collections.abc import Awaitable, Callable

from pydantic import ConfigDict, Field
from pydantic.dataclasses import dataclass

from astrbot.core.agent.run_context import ContextWrapper
from astrbot.core.agent.tool import FunctionTool, ToolExecResult
from astrbot.core.astr_agent_context import AstrAgentContext

from ..document_core.builders.word_builder import WordDocumentBuilder
from ..document_core.models.document import DocumentStatus
from ..mcp_server.schemas import (
    AddHeadingRequest,
    AddParagraphRequest,
    AddSectionBundleRequest,
    AddSummaryCardRequest,
    AddTableRequest,
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
    "Continue calling document tools until all requested sections are added and "
    "export_document succeeds. Do not send a final natural-language reply before export_document."
)
_EXPORT_REQUIRED_FLAG = "_office_doc_force_export_required"
_EXPORT_REQUIRED_BUDGET = "_office_doc_force_export_budget"
_EXPORT_REQUIRED_DOCUMENT_ID = "_office_doc_force_export_document_id"
_EXPORT_REQUIRED_RETRY_BUDGET = 2


def _set_export_required(
    context: ContextWrapper[AstrAgentContext] | None, document_id: str
) -> None:
    if context is None:
        return
    event = getattr(getattr(context, "context", None), "event", None)
    if event is None:
        return
    event.set_extra(_EXPORT_REQUIRED_FLAG, True)
    event.set_extra(_EXPORT_REQUIRED_BUDGET, _EXPORT_REQUIRED_RETRY_BUDGET)
    event.set_extra(_EXPORT_REQUIRED_DOCUMENT_ID, document_id)


def _clear_export_required(context: ContextWrapper[AstrAgentContext] | None) -> None:
    if context is None:
        return
    event = getattr(getattr(context, "context", None), "event", None)
    if event is None:
        return
    extras = event.get_extra()
    extras.pop(_EXPORT_REQUIRED_FLAG, None)
    extras.pop(_EXPORT_REQUIRED_BUDGET, None)
    extras.pop(_EXPORT_REQUIRED_DOCUMENT_ID, None)


@dataclass(config=ConfigDict(arbitrary_types_allowed=True))
class DocumentToolBase(FunctionTool[AstrAgentContext]):
    store: DocumentSessionStore = Field(default_factory=DocumentSessionStore)


@dataclass(config=ConfigDict(arbitrary_types_allowed=True))
class CreateDocumentTool(DocumentToolBase):
    name: str = "create_document"
    description: str = (
        "Create a draft Word document session and return its document_id. "
        "Use this before adding headings, paragraphs, tables, or summary cards."
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
        request = CreateDocumentRequest(
            session_id=str(
                kwargs.get("session_id") or context.context.event.unified_msg_origin
            ),
            title=str(kwargs.get("title") or ""),
            output_name=str(kwargs.get("output_name") or "document.docx"),
            theme_name=str(kwargs.get("theme_name") or "business_report"),
            table_template=str(kwargs.get("table_template") or "report_grid"),
            density=str(kwargs.get("density") or "comfortable"),
            accent_color=str(kwargs.get("accent_color") or ""),
        )
        document = self.store.create_document(request)
        _set_export_required(context, document.document_id)
        return _dump_result(
            ToolResult(
                success=True,
                message=f"Document session created. {_CONTINUE_UNTIL_EXPORT}",
                document=build_document_summary(document),
            )
        )


@dataclass(config=ConfigDict(arbitrary_types_allowed=True))
class AddHeadingTool(DocumentToolBase):
    name: str = "add_heading"
    description: str = "Append a heading block to the current draft document."
    parameters: dict = Field(
        default_factory=lambda: {
            "type": "object",
            "properties": {
                "document_id": {
                    "type": "string",
                    "description": "Target document_id returned by create_document.",
                },
                "text": {
                    "type": "string",
                    "description": "Heading text.",
                },
                "level": {
                    "type": "number",
                    "description": "Heading level from 1 to 6.",
                },
            },
            "required": ["document_id", "text"],
        }
    )

    async def call(
        self, context: ContextWrapper[AstrAgentContext], **kwargs
    ) -> ToolExecResult:
        try:
            request = AddHeadingRequest(
                document_id=str(kwargs.get("document_id") or ""),
                text=str(kwargs.get("text") or ""),
                level=int(kwargs.get("level", 1)),
            )
            document = self.store.add_heading(request)
        except Exception as exc:
            return _dump_result(ToolResult(success=False, message=str(exc)))
        _set_export_required(context, document.document_id)
        return _dump_result(
            ToolResult(
                success=True,
                message=f"Heading added. {_CONTINUE_UNTIL_EXPORT}",
                document=build_document_summary(document),
            )
        )


@dataclass(config=ConfigDict(arbitrary_types_allowed=True))
class AddParagraphTool(DocumentToolBase):
    name: str = "add_paragraph"
    description: str = "Append a paragraph block to the current draft document."
    parameters: dict = Field(
        default_factory=lambda: {
            "type": "object",
            "properties": {
                "document_id": {
                    "type": "string",
                    "description": "Target document_id returned by create_document.",
                },
                "text": {
                    "type": "string",
                    "description": "Paragraph text.",
                },
            },
            "required": ["document_id", "text"],
        }
    )

    async def call(
        self, context: ContextWrapper[AstrAgentContext], **kwargs
    ) -> ToolExecResult:
        try:
            request = AddParagraphRequest(
                document_id=str(kwargs.get("document_id") or ""),
                text=str(kwargs.get("text") or ""),
            )
            document = self.store.add_paragraph(request)
        except Exception as exc:
            return _dump_result(ToolResult(success=False, message=str(exc)))
        _set_export_required(context, document.document_id)
        return _dump_result(
            ToolResult(
                success=True,
                message=f"Paragraph added. {_CONTINUE_UNTIL_EXPORT}",
                document=build_document_summary(document),
            )
        )


@dataclass(config=ConfigDict(arbitrary_types_allowed=True))
class AddSectionBundleTool(DocumentToolBase):
    name: str = "add_section_bundle"
    description: str = (
        "Append one complete report section in a single tool call. "
        "Use this to reduce tool-call count for complex Word reports."
    )
    parameters: dict = Field(
        default_factory=lambda: {
            "type": "object",
            "properties": {
                "document_id": {
                    "type": "string",
                    "description": "Target document_id returned by create_document.",
                },
                "heading": {
                    "type": "string",
                    "description": "Section heading text.",
                },
                "level": {
                    "type": "number",
                    "description": "Heading level from 1 to 6. Usually use 1 for report sections.",
                },
                "blocks": {
                    "type": "array",
                    "description": "Ordered section blocks. Each block must declare a type: paragraph, table, or summary_card.",
                    "items": {
                        "type": "object",
                        "properties": {
                            "type": {"type": "string"},
                            "text": {"type": "string"},
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
                            "table_style": {"type": "string"},
                            "title": {"type": "string"},
                            "items": {
                                "type": "array",
                                "items": {"type": "string"},
                            },
                            "variant": {"type": "string"},
                        },
                        "required": ["type"],
                    },
                },
            },
            "required": ["document_id", "heading", "blocks"],
        }
    )

    async def call(
        self, context: ContextWrapper[AstrAgentContext], **kwargs
    ) -> ToolExecResult:
        try:
            request = AddSectionBundleRequest(
                document_id=str(kwargs.get("document_id") or ""),
                heading=str(kwargs.get("heading") or ""),
                level=int(kwargs.get("level", 1)),
                blocks=list(kwargs.get("blocks") or []),
            )
            document = self.store.add_section_bundle(request)
        except Exception as exc:
            return _dump_result(ToolResult(success=False, message=str(exc)))
        _set_export_required(context, document.document_id)
        return _dump_result(
            ToolResult(
                success=True,
                message=f"Section bundle added. {_CONTINUE_UNTIL_EXPORT}",
                document=build_document_summary(document),
            )
        )


@dataclass(config=ConfigDict(arbitrary_types_allowed=True))
class AddTableTool(DocumentToolBase):
    name: str = "add_table"
    description: str = "Append a table block to the current draft document with an optional style preset."
    parameters: dict = Field(
        default_factory=lambda: {
            "type": "object",
            "properties": {
                "document_id": {
                    "type": "string",
                    "description": "Target document_id returned by create_document.",
                },
                "headers": {
                    "type": "array",
                    "description": "Optional table headers.",
                    "items": {"type": "string"},
                },
                "rows": {
                    "type": "array",
                    "description": "Table rows, each item is an array of cell strings.",
                    "items": {
                        "type": "array",
                        "items": {"type": "string"},
                    },
                },
                "table_style": {
                    "type": "string",
                    "description": "Optional table style preset, e.g. report_grid, metrics_compact, or minimal.",
                },
            },
            "required": ["document_id"],
        }
    )

    async def call(
        self, context: ContextWrapper[AstrAgentContext], **kwargs
    ) -> ToolExecResult:
        try:
            request = AddTableRequest(
                document_id=str(kwargs.get("document_id") or ""),
                headers=[str(item) for item in kwargs.get("headers") or []],
                rows=[
                    [str(cell) for cell in row]
                    for row in (kwargs.get("rows") or [])
                    if isinstance(row, list)
                ],
                table_style=str(kwargs.get("table_style") or ""),
            )
            document = self.store.add_table(request)
        except Exception as exc:
            return _dump_result(ToolResult(success=False, message=str(exc)))
        _set_export_required(context, document.document_id)
        return _dump_result(
            ToolResult(
                success=True,
                message=f"Table added. {_CONTINUE_UNTIL_EXPORT}",
                document=build_document_summary(document),
            )
        )


@dataclass(config=ConfigDict(arbitrary_types_allowed=True))
class AddSummaryCardTool(DocumentToolBase):
    name: str = "add_summary_card"
    description: str = (
        "Append a summary or conclusion card block to the current draft document."
    )
    parameters: dict = Field(
        default_factory=lambda: {
            "type": "object",
            "properties": {
                "document_id": {
                    "type": "string",
                    "description": "Target document_id returned by create_document.",
                },
                "title": {
                    "type": "string",
                    "description": "Card title.",
                },
                "items": {
                    "type": "array",
                    "description": "Card bullet items.",
                    "items": {"type": "string"},
                },
                "variant": {
                    "type": "string",
                    "description": "Card variant, use summary or conclusion.",
                },
            },
            "required": ["document_id", "title", "items"],
        }
    )

    async def call(
        self, context: ContextWrapper[AstrAgentContext], **kwargs
    ) -> ToolExecResult:
        try:
            request = AddSummaryCardRequest(
                document_id=str(kwargs.get("document_id") or ""),
                title=str(kwargs.get("title") or ""),
                items=[str(item) for item in kwargs.get("items") or []],
                variant=str(kwargs.get("variant") or "summary"),
            )
            document = self.store.add_summary_card(request)
        except Exception as exc:
            return _dump_result(ToolResult(success=False, message=str(exc)))
        _set_export_required(context, document.document_id)
        return _dump_result(
            ToolResult(
                success=True,
                message=f"Summary card added. {_CONTINUE_UNTIL_EXPORT}",
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
        _set_export_required(context, document.document_id)
        return _dump_result(
            ToolResult(
                success=True,
                message=(
                    "Document finalized. Call export_document now. "
                    "Do not send a final natural-language reply before export_document."
                ),
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
            document, output_path = self.store.prepare_export_path(request)
            self.builder.build(document, output_path)
            document.status = DocumentStatus.EXPORTED
            document.touch()
            callback_message = ""
            if self.after_export is not None and context is not None:
                callback_message = (
                    await self.after_export(context, str(output_path)) or ""
                )
        except Exception as exc:
            return _dump_result(ToolResult(success=False, message=str(exc)))
        _clear_export_required(context)
        return _dump_result(
            ExportDocumentResult(
                success=True,
                message=callback_message or "Document exported.",
                document=build_document_summary(document),
                file_path=str(output_path),
            )
        )


__all__ = [
    "AddHeadingTool",
    "AddParagraphTool",
    "AddSectionBundleTool",
    "AddSummaryCardTool",
    "AddTableTool",
    "CreateDocumentTool",
    "ExportDocumentTool",
    "FinalizeDocumentTool",
]
