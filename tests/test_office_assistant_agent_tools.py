import json
import struct
import tempfile
import zlib
from collections.abc import Iterator
from datetime import datetime, timedelta, timezone
from pathlib import Path
from uuid import uuid4

import pytest

from astrbot.core.utils.astrbot_path import get_astrbot_plugin_data_path
from astrbot_plugin_office_assistant.agent_tools import (
    build_document_toolset,
)
from astrbot_plugin_office_assistant.agent_tools.document_tools import (
    CreateDocumentTool,
)
from astrbot_plugin_office_assistant.document_core.models.blocks import (
    GroupBlock,
)
from astrbot_plugin_office_assistant.document_core.builders.word_builder import (
    DOCX_TABLE_STYLES,
    WordDocumentBuilder,
)
from astrbot_plugin_office_assistant.mcp_server.schemas import (
    AddBlocksRequest,
    CreateDocumentRequest,
    ExportDocumentRequest,
)
from astrbot_plugin_office_assistant.mcp_server.server import (
    create_server,
)
from astrbot_plugin_office_assistant.mcp_server.session_store import (
    DocumentSessionStore,
)


def _cell_fill(cell) -> str | None:
    from docx.oxml.ns import qn

    tc_pr = cell._tc.tcPr
    if tc_pr is None:
        return None
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        return None
    return shd.get(qn("w:fill"))


@pytest.fixture
def workspace_root() -> Iterator[Path]:
    with tempfile.TemporaryDirectory(dir=Path.cwd()) as temp_dir:
        yield Path(temp_dir)


def _make_workspace(workspace_root: Path, name: str) -> Path:
    workspace_dir = workspace_root / f"{name}-{uuid4().hex}"
    workspace_dir.mkdir(parents=True, exist_ok=True)
    return workspace_dir


def _write_png(path: Path, *, width: int, height: int) -> None:
    def chunk(chunk_type: bytes, data: bytes) -> bytes:
        return (
            struct.pack(">I", len(data))
            + chunk_type
            + data
            + struct.pack(">I", zlib.crc32(chunk_type + data) & 0xFFFFFFFF)
        )

    scanline = b"\x00" + (b"\xFF\x66\x33" * width)
    raw = scanline * height
    png = (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
        + chunk(b"IDAT", zlib.compress(raw))
        + chunk(b"IEND", b"")
    )
    path.write_bytes(png)


def test_build_document_toolset_uses_shared_store_and_default_workspace():
    toolset = build_document_toolset()
    tool_names = [tool.name for tool in toolset.tools]

    assert tool_names == [
        "create_document",
        "add_blocks",
        "finalize_document",
        "export_document",
    ]

    stores = [tool.store for tool in toolset.tools if hasattr(tool, "store")]
    assert len(stores) == len(tool_names)
    assert len({id(store) for store in stores}) == 1

    expected_workspace = (
        Path(get_astrbot_plugin_data_path())
        / "astrbot_plugin_office_assistant"
        / "documents"
    )
    assert stores[0].workspace_dir == expected_workspace


def test_add_blocks_tool_schema_keeps_nested_array_items_for_gemini():
    toolset = build_document_toolset()
    add_blocks_tool = next(tool for tool in toolset.tools if tool.name == "add_blocks")

    block_properties = add_blocks_tool.parameters["properties"]["blocks"]["items"][
        "properties"
    ]
    assert block_properties["blocks"]["type"] == "array"
    assert block_properties["blocks"]["items"]["type"] == "object"
    assert block_properties["blocks"]["items"]["additionalProperties"] is True
    assert block_properties["columns"]["type"] == "array"
    assert block_properties["columns"]["items"]["type"] == "object"
    assert (
        block_properties["columns"]["items"]["properties"]["blocks"]["items"]["type"]
        == "object"
    )


@pytest.mark.asyncio
async def test_create_document_tool_does_not_stringify_missing_session():
    tool = CreateDocumentTool()

    created = json.loads(await tool.call(None, title="No Session"))

    assert created["success"] is True
    document = tool.store.require_document(created["document"]["document_id"])
    assert document.session_id == ""


@pytest.mark.asyncio
async def test_document_toolset_smoke_export(workspace_root: Path):
    docx = pytest.importorskip("docx")
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import RGBColor

    workspace_dir = _make_workspace(workspace_root, "pytest-agent-tools")
    toolset = build_document_toolset(workspace_dir=workspace_dir)
    tool_by_name = {tool.name: tool for tool in toolset.tools}

    created = json.loads(
        await tool_by_name["create_document"].call(
            None,
            session_id="pytest-session",
            title="Pytest Smoke",
            output_name="pytest-smoke.docx",
            theme_name="executive_brief",
            table_template="minimal",
            density="compact",
            accent_color="#AA5500",
        )
    )
    document_id = created["document"]["document_id"]
    assert created["document"]["theme_name"] == "executive_brief"
    assert created["document"]["table_template"] == "minimal"
    assert created["document"]["density"] == "compact"
    assert created["document"]["accent_color"] == "AA5500"

    await tool_by_name["add_blocks"].call(
        None,
        document_id=document_id,
        blocks=[
            {"type": "heading", "text": "Section 1", "level": 1},
            {
                "type": "paragraph",
                "text": "Hello from pytest.",
                "style": {
                    "align": "center",
                    "emphasis": "strong",
                    "font_scale": 1.1,
                },
                "layout": {"spacing_after": 9},
            },
            {
                "type": "list",
                "items": ["Point A", "Point B"],
                "ordered": True,
                "style": {"emphasis": "subtle"},
            },
            {
                "type": "table",
                "headers": ["Metric", "Jan", "Feb"],
                "rows": [["Users", "120", "140"]],
                "table_style": "minimal",
            },
            {
                "type": "summary_card",
                "title": "Conclusion",
                "items": ["The new layout should look more intentional."],
                "variant": "conclusion",
            },
            {"type": "page_break"},
            {"type": "heading", "text": "Appendix", "level": 1},
        ],
    )

    exported = json.loads(
        await tool_by_name["export_document"].call(
            None,
            document_id=document_id,
        )
    )

    assert exported["success"] is True
    assert Path(exported["file_path"]).exists()
    assert Path(exported["file_path"]).parent == workspace_dir
    loaded_doc = docx.Document(exported["file_path"])
    assert loaded_doc.paragraphs[0].text == "Pytest Smoke"
    assert loaded_doc.paragraphs[0].alignment == WD_ALIGN_PARAGRAPH.CENTER
    assert loaded_doc.paragraphs[0].runs[0].bold is True
    assert loaded_doc.paragraphs[0].runs[0].font.color.rgb == RGBColor.from_string(
        "AA5500"
    )
    assert loaded_doc.paragraphs[1].text == "Section 1"
    assert loaded_doc.paragraphs[1].runs[0].bold is True
    assert loaded_doc.paragraphs[2].text == "Hello from pytest."
    assert loaded_doc.paragraphs[2].alignment == WD_ALIGN_PARAGRAPH.CENTER
    assert loaded_doc.paragraphs[2].runs[0].bold is True
    assert loaded_doc.paragraphs[
        2
    ].paragraph_format.first_line_indent.pt == pytest.approx(18, abs=0.5)
    assert loaded_doc.paragraphs[2].paragraph_format.space_after.pt == pytest.approx(
        9, abs=0.5
    )
    assert loaded_doc.paragraphs[3].text == "1. Point A"
    assert loaded_doc.paragraphs[4].text == "2. Point B"
    assert loaded_doc.paragraphs[3].runs[0].font.color.rgb == RGBColor.from_string(
        "AA5500"
    )
    assert "Appendix" in [paragraph.text for paragraph in loaded_doc.paragraphs]
    assert 'w:type="page"' in loaded_doc.element.body.xml
    assert len(loaded_doc.tables) >= 1
    assert loaded_doc.tables[0].style.name == "Table Grid"
    assert loaded_doc.tables[0].rows[0].cells[0].paragraphs[0].runs[0].bold is True
    assert loaded_doc.tables[0].rows[1].cells[0].paragraphs[0].runs[0].bold is False
    assert loaded_doc.tables[0].rows[0].cells[0].text == "Metric"
    assert loaded_doc.tables[0].rows[1].cells[0].text == "Users"
    assert _cell_fill(loaded_doc.tables[0].rows[0].cells[0]) == "F1E4D6"
    assert loaded_doc.tables[0].rows[0].cells[0].paragraphs[0].runs[
        0
    ].font.color.rgb == RGBColor.from_string("AA5500")
    paragraph_texts = [paragraph.text for paragraph in loaded_doc.paragraphs]
    assert "Conclusion" in paragraph_texts
    assert "• The new layout should look more intentional." in paragraph_texts


@pytest.mark.asyncio
async def test_add_blocks_tool_supports_nested_primitives(workspace_root: Path):
    docx = pytest.importorskip("docx")

    workspace_dir = _make_workspace(workspace_root, "pytest-agent-add-blocks")
    toolset = build_document_toolset(workspace_dir=workspace_dir)
    tool_by_name = {tool.name: tool for tool in toolset.tools}

    created = json.loads(
        await tool_by_name["create_document"].call(
            None,
            title="Composite Blocks",
            output_name="composite-blocks.docx",
        )
    )
    document_id = created["document"]["document_id"]

    add_blocks_result = json.loads(
        await tool_by_name["add_blocks"].call(
            None,
            document_id=document_id,
            blocks=[
                {"type": "heading", "text": "Overview", "level": 1},
                {
                    "type": "group",
                    "blocks": [
                        {"type": "paragraph", "text": "Nested intro."},
                        {
                            "type": "list",
                            "items": ["Left detail", "Right detail"],
                            "ordered": False,
                        },
                    ],
                },
                {
                    "type": "columns",
                    "columns": [
                        {
                            "blocks": [
                                {"type": "paragraph", "text": "Column A body."},
                            ]
                        },
                        {
                            "blocks": [
                                {"type": "paragraph", "text": "Column B body."},
                            ]
                        },
                    ],
                },
                {"type": "page_break"},
                {"type": "heading", "text": "Appendix", "level": 1},
            ],
        )
    )

    assert add_blocks_result["success"] is True
    assert add_blocks_result["document"]["block_count"] == 5

    exported = json.loads(
        await tool_by_name["export_document"].call(
            None,
            document_id=document_id,
        )
    )

    loaded_doc = docx.Document(exported["file_path"])
    paragraph_texts = [paragraph.text for paragraph in loaded_doc.paragraphs]

    assert paragraph_texts[1] == "Overview"
    assert "Nested intro." in paragraph_texts
    assert "• Left detail" in paragraph_texts
    assert "• Right detail" in paragraph_texts
    assert "Column A body." in paragraph_texts
    assert "Column B body." in paragraph_texts
    assert "Appendix" in paragraph_texts
    assert 'w:type="page"' in loaded_doc.element.body.xml


@pytest.mark.asyncio
async def test_document_toolset_export_callback_runs(workspace_root: Path):
    pytest.importorskip("docx")

    workspace_dir = _make_workspace(workspace_root, "pytest-agent-tools-callback")
    callback_calls: list[str] = []

    async def after_export(_context, output_path: str) -> str:
        callback_calls.append(output_path)
        return "callback sent"

    toolset = build_document_toolset(
        workspace_dir=workspace_dir,
        after_export=after_export,
    )
    tool_by_name = {tool.name: tool for tool in toolset.tools}

    created = json.loads(
        await tool_by_name["create_document"].call(
            None,
            session_id="pytest-session",
            title="Pytest Callback",
            output_name="pytest-callback.docx",
        )
    )

    exported = json.loads(
        await tool_by_name["export_document"].call(
            object(),
            document_id=created["document"]["document_id"],
        )
    )

    assert exported["success"] is True
    assert exported["message"] == "callback sent"
    assert callback_calls == [exported["file_path"]]


@pytest.mark.asyncio
async def test_export_document_tool_keeps_success_when_callback_fails(
    workspace_root: Path,
):
    pytest.importorskip("docx")

    workspace_dir = _make_workspace(workspace_root, "pytest-agent-tools-callback-fail")

    async def after_export(_context, _output_path: str) -> str:
        raise RuntimeError("send failed")

    toolset = build_document_toolset(
        workspace_dir=workspace_dir,
        after_export=after_export,
    )
    tool_by_name = {tool.name: tool for tool in toolset.tools}

    created = json.loads(
        await tool_by_name["create_document"].call(
            None,
            session_id="pytest-session",
            title="Pytest Callback Failure",
            output_name="pytest-callback-failure.docx",
        )
    )

    exported = json.loads(
        await tool_by_name["export_document"].call(
            object(),
            document_id=created["document"]["document_id"],
        )
    )

    assert exported["success"] is True
    assert "post-export delivery failed" in exported["message"]
    assert Path(exported["file_path"]).exists()


@pytest.mark.asyncio
async def test_mcp_registers_only_core_document_tools():
    server = create_server()
    tools = await server.list_tools()
    tool_names = [tool.name for tool in tools]

    assert tool_names == [
        "create_document",
        "add_blocks",
        "finalize_document",
        "export_document",
    ]

    add_blocks_tool = next(tool for tool in tools if tool.name == "add_blocks")

    assert add_blocks_tool.inputSchema["required"] == ["document_id", "blocks"]
    add_blocks_items = add_blocks_tool.inputSchema["properties"]["blocks"]["items"]
    assert add_blocks_items["type"] == "object"
    assert add_blocks_items["additionalProperties"] is True


def test_document_session_store_keeps_exports_inside_workspace(workspace_root: Path):
    workspace_dir = _make_workspace(workspace_root, "pytest-agent-tools-paths")
    store = DocumentSessionStore(workspace_dir=workspace_dir)
    document = store.create_document(
        CreateDocumentRequest(
            session_id="pytest-session",
            output_name="../unsafe-name.docx",
        )
    )

    _, output_path = store.prepare_export_path(
        ExportDocumentRequest(
            document_id=document.document_id,
            output_dir="reports/q1",
            output_name="../final-report.docx",
        )
    )

    assert (
        output_path
        == (workspace_dir / "reports" / "q1" / "final-report.docx").resolve()
    )
    assert document.metadata.preferred_filename == "unsafe-name.docx"

    with pytest.raises(ValueError, match="relative to the document workspace"):
        ExportDocumentRequest(
            document_id=document.document_id,
            output_dir=str(workspace_dir.parent),
        )

    with pytest.raises(ValueError, match="cannot escape the document workspace"):
        ExportDocumentRequest(
            document_id=document.document_id,
            output_dir="../outside",
        )

    _, windows_style_output_path = store.prepare_export_path(
        ExportDocumentRequest(
            document_id=document.document_id,
            output_dir=r"reports\windows",
            output_name=r"nested\windows-report",
        )
    )

    assert (
        windows_style_output_path
        == (workspace_dir / "reports" / "windows" / "windows-report.docx").resolve()
    )

    with pytest.raises(ValueError, match="relative to the document workspace"):
        ExportDocumentRequest(
            document_id=document.document_id,
            output_dir=r"C:\temp",
        )

    with pytest.raises(ValueError, match="cannot escape the document workspace"):
        ExportDocumentRequest(
            document_id=document.document_id,
            output_dir=r"..\outside",
        )


def test_document_session_store_evicts_oldest_documents_when_capped():
    store = DocumentSessionStore(max_documents=2)
    first = store.create_document(CreateDocumentRequest(title="first"))
    second = store.create_document(CreateDocumentRequest(title="second"))
    third = store.create_document(CreateDocumentRequest(title="third"))

    assert store.get_document(first.document_id) is None
    assert store.get_document(second.document_id) is not None
    assert store.get_document(third.document_id) is not None


def test_document_session_store_evicts_expired_documents_by_ttl():
    store = DocumentSessionStore(ttl=timedelta(seconds=1))
    expired = store.create_document(CreateDocumentRequest(title="expired"))
    fresh = store.create_document(CreateDocumentRequest(title="fresh"))

    expired.metadata.updated_at = datetime.now(timezone.utc) - timedelta(seconds=5)
    fresh.metadata.updated_at = datetime.now(timezone.utc)

    assert store.get_document(expired.document_id) is None
    assert store.get_document(fresh.document_id) is not None


def test_word_document_builder_resolves_logical_table_styles():
    assert WordDocumentBuilder._resolve_docx_table_style("report_grid") == "Table Grid"
    assert (
        WordDocumentBuilder._resolve_docx_table_style("metrics_compact")
        == "Light List Accent 1"
    )
    assert WordDocumentBuilder._resolve_docx_table_style("minimal") == "Table Grid"
    assert WordDocumentBuilder._resolve_docx_table_style("") == "Table Grid"
    assert (
        WordDocumentBuilder._resolve_docx_table_style("custom_style") == "custom_style"
    )


def test_word_document_builder_uses_image_width_px_with_page_cap(
    workspace_root: Path,
):
    docx = pytest.importorskip("docx")
    from docx.shared import Inches

    workspace_dir = _make_workspace(workspace_root, "pytest-agent-tools-image-width")
    image_path = workspace_dir / "wide.png"
    output_path = workspace_dir / "image-width.docx"
    _write_png(image_path, width=2000, height=400)

    from astrbot_plugin_office_assistant.document_core.models.blocks import ImageBlock
    from astrbot_plugin_office_assistant.document_core.models.document import (
        DocumentMetadata,
        DocumentModel,
    )

    document = DocumentModel(
        document_id="image-width-test",
        metadata=DocumentMetadata(title="Image Width"),
        blocks=[
            ImageBlock(
                path=str(image_path),
                width_px=1600,
                caption="Image caption",
            )
        ],
    )

    WordDocumentBuilder().build(document, output_path)

    loaded_doc = docx.Document(output_path)
    assert len(loaded_doc.inline_shapes) == 1
    max_width = (
        loaded_doc.sections[0].page_width
        - loaded_doc.sections[0].left_margin
        - loaded_doc.sections[0].right_margin
    )
    assert loaded_doc.inline_shapes[0].width == min(Inches(1600 / 96.0), max_width)
    assert "Image caption" in [paragraph.text for paragraph in loaded_doc.paragraphs]


@pytest.mark.asyncio
async def test_document_toolset_falls_back_when_metrics_table_style_is_missing(
    workspace_root: Path, monkeypatch: pytest.MonkeyPatch
):
    docx = pytest.importorskip("docx")

    monkeypatch.setitem(DOCX_TABLE_STYLES, "metrics_compact", "Missing Docx Style")

    workspace_dir = _make_workspace(workspace_root, "pytest-agent-tools-missing-style")
    toolset = build_document_toolset(workspace_dir=workspace_dir)
    tool_by_name = {tool.name: tool for tool in toolset.tools}

    created = json.loads(
        await tool_by_name["create_document"].call(
            None,
            title="Missing Style Fallback",
            output_name="missing-style-fallback.docx",
            table_template="metrics_compact",
        )
    )

    await tool_by_name["add_blocks"].call(
        None,
        document_id=created["document"]["document_id"],
        blocks=[
            {
                "type": "table",
                "headers": ["Metric", "Value"],
                "rows": [["Users", "42"]],
                "table_style": "metrics_compact",
            }
        ],
    )

    exported = json.loads(
        await tool_by_name["export_document"].call(
            None,
            document_id=created["document"]["document_id"],
        )
    )

    assert exported["success"] is True
    loaded_doc = docx.Document(exported["file_path"])
    assert len(loaded_doc.tables) == 1
    assert loaded_doc.tables[0].style.name == "Table Grid"


def test_document_session_store_expands_summary_card_blocks():
    store = DocumentSessionStore()
    document = store.create_document(CreateDocumentRequest(title="Summary Test"))

    updated = store.add_blocks(
        AddBlocksRequest(
            document_id=document.document_id,
            blocks=[
                {
                    "type": "summary_card",
                    "title": "Conclusion",
                    "items": ["First takeaway"],
                    "variant": "conclusion",
                }
            ],
        )
    )

    assert len(updated.blocks) == 1
    assert isinstance(updated.blocks[0], GroupBlock)
    assert updated.blocks[0].blocks[0].text == "Conclusion"
    assert updated.blocks[0].blocks[1].items == ["First takeaway"]
