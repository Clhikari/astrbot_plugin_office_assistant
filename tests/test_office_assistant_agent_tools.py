import json
import struct
import zlib
from datetime import datetime, timedelta, timezone
from pathlib import Path
from uuid import uuid4

import pytest
from astrbot_plugin_office_assistant.agent_tools import (
    build_document_toolset,
)
from astrbot_plugin_office_assistant.agent_tools.document_tools import (
    CreateDocumentTool,
)
from astrbot_plugin_office_assistant.document_core.builders.table_renderer import (
    DOCX_TABLE_STYLES,
    TableRenderer,
)
from astrbot_plugin_office_assistant.document_core.builders.word_builder import (
    WordDocumentBuilder,
)
from astrbot_plugin_office_assistant.document_core.models.blocks import (
    GroupBlock,
    ParagraphBlock,
    ParagraphRun,
    TableBlock,
)
from astrbot_plugin_office_assistant.mcp_server.schemas import (
    AddBlocksRequest,
    AddTableRequest,
    BlockHeadingInput,
    CreateDocumentRequest,
    ExportDocumentRequest,
    SectionParagraphInput,
    SectionTableInput,
)
from astrbot_plugin_office_assistant.mcp_server.server import (
    create_server,
)
from astrbot_plugin_office_assistant.mcp_server.session_store import (
    DocumentSessionStore,
)
from pydantic import ValidationError

from astrbot.core.utils.astrbot_path import get_astrbot_plugin_data_path


def _cell_fill(cell) -> str | None:
    from docx.oxml.ns import qn

    tc_pr = cell._tc.tcPr
    if tc_pr is None:
        return None
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        return None
    return shd.get(qn("w:fill"))


def _grid_span(cell) -> int:
    from docx.oxml.ns import qn

    tc_pr = cell._tc.tcPr
    if tc_pr is None:
        return 1
    span = tc_pr.find(qn("w:gridSpan"))
    if span is None:
        return 1
    return int(span.get(qn("w:val"), "1"))


def _run_rgb(cell) -> str | None:
    runs = cell.paragraphs[0].runs
    if not runs:
        return None
    color = runs[0].font.color.rgb
    return str(color) if color is not None else None


def _run_bold(cell) -> bool | None:
    runs = cell.paragraphs[0].runs
    if not runs:
        return None
    return runs[0].bold


def _paragraph_run_rgb(paragraph) -> str | None:
    runs = paragraph.runs
    if not runs:
        return None
    color = runs[0].font.color.rgb
    return str(color) if color is not None else None


def _paragraph_run_size(paragraph) -> float | None:
    runs = paragraph.runs
    if not runs or runs[0].font.size is None:
        return None
    return runs[0].font.size.pt


def _find_paragraph(doc, text: str):
    return next(paragraph for paragraph in doc.paragraphs if paragraph.text == text)


def _table_border_size(table, edge_name: str) -> str | None:
    from docx.oxml.ns import qn

    tbl_pr = table._tbl.tblPr
    if tbl_pr is None:
        return None
    tbl_borders = tbl_pr.find(qn("w:tblBorders"))
    if tbl_borders is None:
        return None
    edge = tbl_borders.find(qn(f"w:{edge_name}"))
    if edge is None:
        return None
    return edge.get(qn("w:sz"))


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

    scanline = b"\x00" + (b"\xff\x66\x33" * width)
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


def test_create_document_tool_schema_exposes_document_style():
    toolset = build_document_toolset()
    create_document_tool = next(
        tool for tool in toolset.tools if tool.name == "create_document"
    )

    properties = create_document_tool.parameters["properties"]
    document_style = properties["document_style"]["properties"]
    table_defaults = document_style["table_defaults"]["properties"]

    assert document_style["brief"]["type"] == "string"
    assert document_style["heading_color"]["type"] == "string"
    assert document_style["title_align"]["enum"] == [
        "left",
        "center",
        "right",
        "justify",
    ]
    assert document_style["body_font_size"]["type"] == "number"
    assert document_style["body_line_spacing"]["type"] == "number"
    assert document_style["paragraph_space_after"]["type"] == "number"
    assert document_style["list_space_after"]["type"] == "number"
    assert document_style["summary_card_defaults"]["type"] == "object"
    assert document_style["summary_card_defaults"]["properties"]["title_align"][
        "enum"
    ] == ["left", "center", "right", "justify"]
    assert document_style["summary_card_defaults"]["properties"]["title_emphasis"][
        "enum"
    ] == ["normal", "strong", "subtle"]
    assert table_defaults["preset"]["enum"] == [
        "report_grid",
        "metrics_compact",
        "minimal",
    ]
    assert table_defaults["table_align"]["enum"] == ["left", "center"]
    assert table_defaults["border_style"]["enum"] == [
        "minimal",
        "standard",
        "strong",
    ]
    assert table_defaults["cell_align"]["enum"] == ["left", "center", "right"]


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
    assert block_properties["numeric_columns"]["items"]["type"] == "integer"
    assert (
        block_properties["runs"]["items"]["properties"]["italic"]["type"] == "boolean"
    )
    run_properties = block_properties["runs"]["items"]["properties"]
    assert run_properties["bold"]["type"] == "boolean"
    assert run_properties["underline"]["type"] == "boolean"
    assert run_properties["code"]["type"] == "boolean"
    assert block_properties["text"]["type"] == "string"
    assert block_properties["runs"]["type"] == "array"
    assert block_properties["runs"]["items"]["type"] == "object"
    assert block_properties["header_groups"]["type"] == "array"
    assert block_properties["header_groups"]["items"]["type"] == "object"
    assert (
        block_properties["header_groups"]["items"]["properties"]["title"]["type"]
        == "string"
    )
    assert (
        block_properties["header_groups"]["items"]["properties"]["title"]["minLength"]
        == 1
    )
    assert (
        block_properties["header_groups"]["items"]["properties"]["span"]["type"]
        == "integer"
    )
    assert (
        block_properties["header_groups"]["items"]["properties"]["span"]["minimum"] == 1
    )
    assert block_properties["header_groups"]["items"]["required"] == [
        "title",
        "span",
    ]
    assert block_properties["header_fill"]["type"] == "string"
    assert block_properties["header_text_color"]["type"] == "string"
    assert block_properties["banded_rows"]["type"] == "boolean"
    assert block_properties["banded_row_fill"]["type"] == "string"
    assert block_properties["first_column_bold"]["type"] == "boolean"
    assert block_properties["table_align"]["enum"] == ["left", "center"]
    assert block_properties["border_style"]["enum"] == [
        "minimal",
        "standard",
        "strong",
    ]
    assert block_properties["caption_emphasis"]["enum"] == ["normal", "strong"]


def test_paragraph_schema_requires_text_or_runs():
    with pytest.raises(ValidationError, match="paragraph requires text or runs"):
        SectionParagraphInput.model_validate(
            {
                "type": "paragraph",
                "text": "",
                "runs": [],
            }
        )


def test_word_document_builder_prefers_runs_when_both_text_and_runs_exist():
    block = ParagraphBlock(
        text="plain text",
        runs=[
            ParagraphRun(text="rich"),
            ParagraphRun(text=" content"),
        ],
    )

    assert WordDocumentBuilder._paragraph_text(block) == "rich content"


@pytest.mark.asyncio
async def test_create_document_tool_does_not_stringify_missing_session():
    tool = CreateDocumentTool()

    created = json.loads(await tool.call(None, title="No Session"))

    assert created["success"] is True
    document = tool.store.require_document(created["document"]["document_id"])
    assert document.session_id == ""


def test_create_document_request_normalizes_document_style():
    request = CreateDocumentRequest(
        title="Styled",
        document_style={
            "brief": "deep blue business report",
            "heading_color": "#1f4e79",
            "title_align": "left",
            "body_font_size": 12,
            "body_line_spacing": 1.25,
            "paragraph_space_after": 14,
            "list_space_after": 11,
            "summary_card_defaults": {
                "title_align": "center",
                "title_emphasis": "strong",
                "title_font_scale": 1.2,
                "title_space_before": 12,
                "title_space_after": 4,
                "list_space_after": 8,
            },
            "table_defaults": {
                "preset": "minimal",
                "header_fill": "#dce6f1",
                "header_text_color": "ffffff",
                "banded_rows": True,
                "banded_row_fill": "eef4fa",
                "first_column_bold": True,
                "table_align": "left",
                "border_style": "strong",
                "caption_emphasis": "strong",
                "cell_align": "center",
            },
        },
    )

    assert request.document_style.brief == "deep blue business report"
    assert request.document_style.heading_color == "1F4E79"
    assert request.document_style.title_align == "left"
    assert request.document_style.body_font_size == 12
    assert request.document_style.body_line_spacing == 1.25
    assert request.document_style.paragraph_space_after == 14
    assert request.document_style.list_space_after == 11
    assert request.document_style.summary_card_defaults.title_align == "center"
    assert request.document_style.summary_card_defaults.title_emphasis == "strong"
    assert request.document_style.summary_card_defaults.title_font_scale == 1.2
    assert request.document_style.summary_card_defaults.title_space_before == 12
    assert request.document_style.summary_card_defaults.title_space_after == 4
    assert request.document_style.summary_card_defaults.list_space_after == 8
    assert request.document_style.table_defaults.preset == "minimal"
    assert request.document_style.table_defaults.header_fill == "DCE6F1"
    assert request.document_style.table_defaults.header_text_color == "FFFFFF"
    assert request.document_style.table_defaults.banded_rows is True
    assert request.document_style.table_defaults.banded_row_fill == "EEF4FA"
    assert request.document_style.table_defaults.first_column_bold is True
    assert request.document_style.table_defaults.table_align == "left"
    assert request.document_style.table_defaults.border_style == "strong"
    assert request.document_style.table_defaults.caption_emphasis == "strong"
    assert request.document_style.table_defaults.cell_align == "center"


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


@pytest.mark.asyncio
async def test_create_document_tool_applies_document_style_defaults(
    workspace_root: Path,
):
    docx = pytest.importorskip("docx")
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    workspace_dir = _make_workspace(workspace_root, "pytest-document-style")
    toolset = build_document_toolset(workspace_dir=workspace_dir)
    tool_by_name = {tool.name: tool for tool in toolset.tools}

    created = json.loads(
        await tool_by_name["create_document"].call(
            None,
            title="Styled Report",
            output_name="styled-report.docx",
            theme_name="business_report",
            document_style={
                "brief": "deep blue business report",
                "heading_color": "0F4C81",
                "title_align": "left",
                "body_font_size": 12,
                "body_line_spacing": 1.25,
                "paragraph_space_after": 14,
                "list_space_after": 11,
                "summary_card_defaults": {
                    "title_align": "center",
                    "title_emphasis": "strong",
                    "title_font_scale": 1.2,
                    "title_space_before": 12,
                    "title_space_after": 4,
                    "list_space_after": 8,
                },
                "table_defaults": {
                    "preset": "minimal",
                    "header_fill": "DCE6F1",
                    "header_text_color": "123456",
                    "banded_rows": True,
                    "banded_row_fill": "EEF4FA",
                    "first_column_bold": True,
                    "table_align": "left",
                    "border_style": "standard",
                    "caption_emphasis": "strong",
                    "cell_align": "center",
                },
            },
        )
    )
    document_id = created["document"]["document_id"]
    assert created["document"]["document_style"]["brief"] == "deep blue business report"
    assert created["document"]["document_style"]["heading_color"] == "0F4C81"
    assert created["document"]["document_style"]["title_align"] == "left"

    await tool_by_name["add_blocks"].call(
        None,
        document_id=document_id,
        blocks=[
            {"type": "heading", "text": "Overview", "level": 1},
            {"type": "paragraph", "text": "Styled body paragraph."},
            {"type": "list", "items": ["Alpha", "Beta"]},
            {
                "type": "summary_card",
                "title": "Highlights",
                "items": ["Stable revenue", "Lower churn"],
                "variant": "conclusion",
            },
            {
                "type": "table",
                "caption": "Quarterly Summary",
                "headers": ["Region", "Score"],
                "rows": [["East", "92"], ["West", "88"]],
            },
        ],
    )

    exported = json.loads(
        await tool_by_name["export_document"].call(
            None,
            document_id=document_id,
        )
    )

    loaded_doc = docx.Document(exported["file_path"])
    heading_paragraph = _find_paragraph(loaded_doc, "Overview")
    title_paragraph = _find_paragraph(loaded_doc, "Styled Report")
    body_paragraph = _find_paragraph(loaded_doc, "Styled body paragraph.")
    list_paragraph = _find_paragraph(loaded_doc, "• Alpha")
    summary_title_paragraph = _find_paragraph(loaded_doc, "Highlights")
    summary_item_paragraph = _find_paragraph(loaded_doc, "• Stable revenue")
    table = loaded_doc.tables[0]

    assert title_paragraph.alignment == WD_ALIGN_PARAGRAPH.LEFT
    assert _paragraph_run_rgb(heading_paragraph) == "0F4C81"
    assert _paragraph_run_size(body_paragraph) == 12
    assert float(body_paragraph.paragraph_format.line_spacing) == pytest.approx(1.25)
    assert body_paragraph.paragraph_format.space_after.pt == pytest.approx(14, abs=0.5)
    assert list_paragraph.paragraph_format.space_after.pt == pytest.approx(11, abs=0.5)
    assert summary_title_paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER
    assert summary_title_paragraph.runs[0].bold is True
    assert _paragraph_run_size(summary_title_paragraph) == pytest.approx(14.0, abs=0.5)
    assert summary_title_paragraph.paragraph_format.space_before.pt == pytest.approx(
        12, abs=0.5
    )
    assert summary_title_paragraph.paragraph_format.space_after.pt == pytest.approx(
        4, abs=0.5
    )
    assert summary_item_paragraph.paragraph_format.space_after.pt == pytest.approx(
        8, abs=0.5
    )
    assert _cell_fill(table.rows[0].cells[0]) == "DCE6F1"
    assert _run_rgb(table.rows[0].cells[0]) == "123456"
    assert _cell_fill(table.rows[2].cells[0]) == "EEF4FA"
    assert _run_bold(table.rows[2].cells[0]) is True
    assert _run_bold(table.rows[2].cells[1]) is False
    assert _table_border_size(table, "top") == "8"


@pytest.mark.asyncio
async def test_add_blocks_tool_supports_rich_text_paragraph_runs(
    workspace_root: Path,
):
    docx = pytest.importorskip("docx")

    workspace_dir = _make_workspace(workspace_root, "pytest-agent-rich-paragraph")
    toolset = build_document_toolset(workspace_dir=workspace_dir)
    tool_by_name = {tool.name: tool for tool in toolset.tools}

    created = json.loads(
        await tool_by_name["create_document"].call(
            None,
            title="段落富文本",
            output_name="rich-paragraph.docx",
        )
    )
    document_id = created["document"]["document_id"]

    await tool_by_name["add_blocks"].call(
        None,
        document_id=document_id,
        blocks=[
            {
                "type": "paragraph",
                "runs": [
                    {"text": "粗体", "bold": True},
                    {"text": " / "},
                    {"text": "斜体", "italic": True},
                    {"text": " / "},
                    {"text": "下划线", "underline": True},
                    {"text": " / "},
                    {"text": "代码", "code": True},
                ],
            },
        ],
    )

    exported = json.loads(
        await tool_by_name["export_document"].call(
            None,
            document_id=document_id,
        )
    )

    loaded_doc = docx.Document(exported["file_path"])
    rich_paragraph = loaded_doc.paragraphs[1]

    assert rich_paragraph.runs[0].bold is True
    assert rich_paragraph.runs[2].italic is True
    assert rich_paragraph.runs[4].underline is True
    assert rich_paragraph.runs[6].font.name == "Consolas"
    assert rich_paragraph.runs[0].text == "粗体"
    assert rich_paragraph.runs[6].text == "代码"


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
async def test_add_blocks_tool_supports_enhanced_tables(workspace_root: Path):
    docx = pytest.importorskip("docx")
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Cm

    workspace_dir = _make_workspace(workspace_root, "pytest-agent-enhanced-table")
    toolset = build_document_toolset(workspace_dir=workspace_dir)
    tool_by_name = {tool.name: tool for tool in toolset.tools}

    created = json.loads(
        await tool_by_name["create_document"].call(
            None,
            title="增强表格",
            output_name="enhanced-table.docx",
        )
    )
    document_id = created["document"]["document_id"]

    await tool_by_name["add_blocks"].call(
        None,
        document_id=document_id,
        blocks=[
            {
                "type": "table",
                "caption": "季度经营指标",
                "headers": ["区域", "目标", "完成率"],
                "rows": [["华东", "120", "98%"], ["华南", "88", "103%"]],
                "column_widths": [4.2, 3.0, 3.0],
                "numeric_columns": [1, 2],
            }
        ],
    )

    exported = json.loads(
        await tool_by_name["export_document"].call(
            None,
            document_id=document_id,
        )
    )

    loaded_doc = docx.Document(exported["file_path"])
    assert len(loaded_doc.tables) == 1

    table = loaded_doc.tables[0]
    assert table.rows[0].cells[0].text == "季度经营指标"
    assert table.rows[1].cells[0].text == "区域"
    assert table.rows[2].cells[0].text == "华东"
    assert table.rows[2].cells[1].paragraphs[0].alignment == WD_ALIGN_PARAGRAPH.RIGHT
    assert table.rows[3].cells[2].paragraphs[0].alignment == WD_ALIGN_PARAGRAPH.RIGHT
    assert abs(table.rows[1].cells[0].width - Cm(4.2)) < 20000
    assert abs(table.rows[1].cells[1].width - Cm(3.0)) < 20000
    assert _cell_fill(table.rows[2].cells[0]) == "F7FBFF"
    assert _cell_fill(table.rows[3].cells[0]) is None


@pytest.mark.asyncio
@pytest.mark.parametrize("table_style", ["report_grid", "metrics_compact", "minimal"])
async def test_add_blocks_tool_supports_grouped_table_headers(
    workspace_root: Path,
    table_style: str,
):
    docx = pytest.importorskip("docx")
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Cm

    workspace_dir = _make_workspace(
        workspace_root, f"pytest-agent-grouped-table-{table_style}"
    )
    toolset = build_document_toolset(workspace_dir=workspace_dir)
    tool_by_name = {tool.name: tool for tool in toolset.tools}

    created = json.loads(
        await tool_by_name["create_document"].call(
            None,
            title="分组表头",
            output_name=f"grouped-table-{table_style}.docx",
            table_template=table_style,
        )
    )
    document_id = created["document"]["document_id"]

    await tool_by_name["add_blocks"].call(
        None,
        document_id=document_id,
        blocks=[
            {
                "type": "table",
                "caption": "季度经营指标",
                "header_groups": [
                    {"title": "经营数据", "span": 2},
                    {"title": "结果", "span": 2},
                ],
                "headers": ["区域", "目标", "完成值", "完成率"],
                "rows": [
                    ["华东", "120", "118", "98%"],
                    ["华南", "88", "91", "103%"],
                ],
                "column_widths": [3.2, 2.4, 2.4, 2.4],
                "numeric_columns": [1, 2, 3],
                "table_style": table_style,
            }
        ],
    )

    exported = json.loads(
        await tool_by_name["export_document"].call(
            None,
            document_id=document_id,
        )
    )

    loaded_doc = docx.Document(exported["file_path"])
    table = loaded_doc.tables[0]

    assert len(table.rows) == 5
    assert table.rows[0].cells[0].text == "季度经营指标"
    assert table.rows[1].cells[0].text == "经营数据"
    assert _grid_span(table.rows[1].cells[0]) == 2
    assert table.rows[1].cells[2].text == "结果"
    assert _grid_span(table.rows[1].cells[2]) == 2
    assert table.rows[2].cells[0].text == "区域"
    assert table.rows[3].cells[1].paragraphs[0].alignment == WD_ALIGN_PARAGRAPH.RIGHT
    assert table.rows[4].cells[3].paragraphs[0].alignment == WD_ALIGN_PARAGRAPH.RIGHT
    assert abs(table.rows[2].cells[0].width - Cm(3.2)) < 20000


@pytest.mark.asyncio
async def test_add_blocks_tool_applies_custom_table_style_overrides(
    workspace_root: Path,
):
    docx = pytest.importorskip("docx")
    from docx.enum.table import WD_TABLE_ALIGNMENT

    workspace_dir = _make_workspace(workspace_root, "pytest-agent-custom-table-style")
    toolset = build_document_toolset(workspace_dir=workspace_dir)
    tool_by_name = {tool.name: tool for tool in toolset.tools}

    created = json.loads(
        await tool_by_name["create_document"].call(
            None,
            title="样式覆盖",
            output_name="custom-table-style.docx",
        )
    )
    document_id = created["document"]["document_id"]

    await tool_by_name["add_blocks"].call(
        None,
        document_id=document_id,
        blocks=[
            {
                "type": "table",
                "caption": "季度经营指标",
                "caption_emphasis": "strong",
                "header_groups": [
                    {"title": "经营数据", "span": 2},
                    {"title": "结果", "span": 2},
                ],
                "headers": ["区域", "目标", "完成值", "完成率"],
                "rows": [
                    ["华东", "120", "118", "98%"],
                    ["华南", "88", "91", "103%"],
                ],
                "header_fill": "1F4E79",
                "header_text_color": "FFFFFF",
                "banded_rows": True,
                "banded_row_fill": "EEF4FA",
                "first_column_bold": True,
                "table_align": "left",
                "border_style": "strong",
            }
        ],
    )

    exported = json.loads(
        await tool_by_name["export_document"].call(
            None,
            document_id=document_id,
        )
    )

    loaded_doc = docx.Document(exported["file_path"])
    table = loaded_doc.tables[0]

    assert table.alignment == WD_TABLE_ALIGNMENT.LEFT
    assert _cell_fill(table.rows[0].cells[0]) == "1F4E79"
    assert _run_rgb(table.rows[0].cells[0]) == "FFFFFF"
    assert _cell_fill(table.rows[1].cells[0]) == "1F4E79"
    assert _run_rgb(table.rows[1].cells[0]) == "FFFFFF"
    assert _cell_fill(table.rows[3].cells[0]) == "EEF4FA"
    assert _cell_fill(table.rows[4].cells[0]) is None
    assert _run_bold(table.rows[3].cells[0]) is True
    assert _run_bold(table.rows[3].cells[1]) is False
    assert _table_border_size(table, "top") == "16"


@pytest.mark.asyncio
async def test_add_blocks_tool_treats_table_title_as_caption_alias(
    workspace_root: Path,
):
    docx = pytest.importorskip("docx")

    workspace_dir = _make_workspace(workspace_root, "pytest-agent-table-title-alias")
    toolset = build_document_toolset(workspace_dir=workspace_dir)
    tool_by_name = {tool.name: tool for tool in toolset.tools}

    created = json.loads(
        await tool_by_name["create_document"].call(
            None,
            title="表格标题别名",
            output_name="table-title-alias.docx",
        )
    )
    document_id = created["document"]["document_id"]

    await tool_by_name["add_blocks"].call(
        None,
        document_id=document_id,
        blocks=[
            {
                "type": "table",
                "title": "季度经营指标总览",
                "headers": ["区域", "营收（万元）"],
                "rows": [["华东", "1280"]],
            }
        ],
    )

    exported = json.loads(
        await tool_by_name["export_document"].call(
            None,
            document_id=document_id,
        )
    )

    loaded_doc = docx.Document(exported["file_path"])
    table = loaded_doc.tables[0]
    assert table.rows[0].cells[0].text == "季度经营指标总览"
    assert table.rows[1].cells[0].text == "区域"


@pytest.mark.asyncio
async def test_add_blocks_tool_absorbs_heading_before_table_into_table_title(
    workspace_root: Path,
):
    docx = pytest.importorskip("docx")

    workspace_dir = _make_workspace(workspace_root, "pytest-agent-table-heading-merge")
    toolset = build_document_toolset(workspace_dir=workspace_dir)
    tool_by_name = {tool.name: tool for tool in toolset.tools}

    created = json.loads(
        await tool_by_name["create_document"].call(
            None,
            title="标题吸收",
            output_name="table-heading-merge.docx",
        )
    )
    document_id = created["document"]["document_id"]

    await tool_by_name["add_blocks"].call(
        None,
        document_id=document_id,
        blocks=[
            {"type": "heading", "text": "季度经营指标总览", "level": 2},
            {
                "type": "table",
                "headers": ["区域", "营收（万元）"],
                "rows": [["华东", "1280"]],
            },
        ],
    )

    exported = json.loads(
        await tool_by_name["export_document"].call(
            None,
            document_id=document_id,
        )
    )

    loaded_doc = docx.Document(exported["file_path"])
    assert len(loaded_doc.tables) == 1
    table = loaded_doc.tables[0]
    assert table.rows[0].cells[0].text == "季度经营指标总览"
    assert table.rows[1].cells[0].text == "区域"
    assert "季度经营指标总览" not in [
        paragraph.text for paragraph in loaded_doc.paragraphs[1:]
    ]


@pytest.mark.asyncio
async def test_add_blocks_tool_drops_heading_that_duplicates_document_title(
    workspace_root: Path,
):
    docx = pytest.importorskip("docx")

    workspace_dir = _make_workspace(workspace_root, "pytest-agent-duplicate-title")
    toolset = build_document_toolset(workspace_dir=workspace_dir)
    tool_by_name = {tool.name: tool for tool in toolset.tools}

    created = json.loads(
        await tool_by_name["create_document"].call(
            None,
            title="项目阶段汇报",
            output_name="duplicate-title.docx",
        )
    )
    document_id = created["document"]["document_id"]

    await tool_by_name["add_blocks"].call(
        None,
        document_id=document_id,
        blocks=[
            {"type": "heading", "text": "项目阶段汇报", "level": 1},
            {"type": "paragraph", "text": "正文内容。"},
        ],
    )

    exported = json.loads(
        await tool_by_name["export_document"].call(
            None,
            document_id=document_id,
        )
    )

    loaded_doc = docx.Document(exported["file_path"])
    paragraph_texts = [paragraph.text for paragraph in loaded_doc.paragraphs]
    assert paragraph_texts.count("项目阶段汇报") == 1
    assert "正文内容。" in paragraph_texts


@pytest.mark.asyncio
async def test_add_blocks_tool_drops_duplicate_document_title_before_table_promotion(
    workspace_root: Path,
):
    docx = pytest.importorskip("docx")

    workspace_dir = _make_workspace(
        workspace_root, "pytest-agent-duplicate-title-before-table"
    )
    toolset = build_document_toolset(workspace_dir=workspace_dir)
    tool_by_name = {tool.name: tool for tool in toolset.tools}

    created = json.loads(
        await tool_by_name["create_document"].call(
            None,
            title="项目阶段汇报",
            output_name="duplicate-title-before-table.docx",
        )
    )
    document_id = created["document"]["document_id"]

    await tool_by_name["add_blocks"].call(
        None,
        document_id=document_id,
        blocks=[
            {"type": "heading", "text": "项目阶段汇报", "level": 1},
            {
                "type": "table",
                "headers": ["区域", "营收（万元）"],
                "rows": [["华东", "1280"]],
            },
        ],
    )

    exported = json.loads(
        await tool_by_name["export_document"].call(
            None,
            document_id=document_id,
        )
    )

    loaded_doc = docx.Document(exported["file_path"])
    paragraph_texts = [paragraph.text for paragraph in loaded_doc.paragraphs]
    table = loaded_doc.tables[0]

    assert paragraph_texts.count("项目阶段汇报") == 1
    assert table.rows[0].cells[0].text == "区域"
    assert "项目阶段汇报" not in table.rows[0].cells[0].text


@pytest.mark.asyncio
async def test_add_blocks_tool_does_not_absorb_long_heading_before_table_into_table_title(
    workspace_root: Path,
):
    docx = pytest.importorskip("docx")

    workspace_dir = _make_workspace(
        workspace_root, "pytest-agent-table-long-heading-no-merge"
    )
    toolset = build_document_toolset(workspace_dir=workspace_dir)
    tool_by_name = {tool.name: tool for tool in toolset.tools}
    long_heading = "季度经营指标总览" * 20

    created = json.loads(
        await tool_by_name["create_document"].call(
            None,
            title="长标题保留",
            output_name="table-long-heading-no-merge.docx",
        )
    )
    document_id = created["document"]["document_id"]

    await tool_by_name["add_blocks"].call(
        None,
        document_id=document_id,
        blocks=[
            {"type": "heading", "text": long_heading, "level": 2},
            {
                "type": "table",
                "headers": ["区域", "营收（万元）"],
                "rows": [["华东", "1280"]],
            },
        ],
    )

    exported = json.loads(
        await tool_by_name["export_document"].call(
            None,
            document_id=document_id,
        )
    )

    loaded_doc = docx.Document(exported["file_path"])
    paragraph_texts = [paragraph.text for paragraph in loaded_doc.paragraphs]
    assert long_heading in paragraph_texts[1:]

    table = loaded_doc.tables[0]
    assert table.rows[0].cells[0].text == "区域"
    assert long_heading not in table.rows[0].cells[0].text
    assert len(table.rows[0].cells) == len(table.rows[1].cells)


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

    exported = await tool_by_name["export_document"].call(
        object(),
        document_id=created["document"]["document_id"],
    )

    assert exported is None
    assert len(callback_calls) == 1
    assert Path(callback_calls[0]).exists()


@pytest.mark.asyncio
async def test_document_toolset_preserves_positional_after_export_callback(
    workspace_root: Path,
):
    pytest.importorskip("docx")

    workspace_dir = _make_workspace(
        workspace_root, "pytest-agent-tools-positional-callback"
    )
    callback_calls: list[str] = []

    async def after_export(_context, output_path: str) -> str:
        callback_calls.append(output_path)
        return "callback sent"

    toolset = build_document_toolset(workspace_dir, after_export)
    tool_by_name = {tool.name: tool for tool in toolset.tools}

    created = json.loads(
        await tool_by_name["create_document"].call(
            None,
            session_id="pytest-session",
            title="Positional Callback",
            output_name="positional-callback.docx",
        )
    )

    exported = await tool_by_name["export_document"].call(
        object(),
        document_id=created["document"]["document_id"],
    )

    assert exported is None
    assert len(callback_calls) == 1
    assert Path(callback_calls[0]).exists()
    assert Path(callback_calls[0]).name == "positional-callback.docx"


@pytest.mark.asyncio
async def test_document_toolset_runs_after_export_hooks_before_delivery_callback(
    workspace_root: Path,
):
    pytest.importorskip("docx")

    workspace_dir = _make_workspace(workspace_root, "pytest-agent-tools-after-export")
    hook_calls: list[tuple[str, str]] = []
    callback_calls: list[str] = []

    async def after_export_hook(context):
        hook_calls.append(
            (context.document.status.value, Path(context.output_path).name)
        )
        return context

    async def after_export(_context, output_path: str) -> str:
        assert hook_calls == [("exported", "after-export-hook.docx")]
        callback_calls.append(output_path)
        return "callback sent"

    toolset = build_document_toolset(
        workspace_dir=workspace_dir,
        after_export_hooks=[after_export_hook],
        after_export=after_export,
    )
    tool_by_name = {tool.name: tool for tool in toolset.tools}

    created = json.loads(
        await tool_by_name["create_document"].call(
            None,
            session_id="pytest-session",
            title="After Export Hook",
            output_name="after-export-hook.docx",
        )
    )

    exported = await tool_by_name["export_document"].call(
        object(),
        document_id=created["document"]["document_id"],
    )

    assert exported is None
    assert hook_calls == [("exported", "after-export-hook.docx")]
    assert len(callback_calls) == 1
    assert Path(callback_calls[0]).exists()


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
async def test_document_toolset_runs_before_export_hooks(workspace_root: Path):
    docx = pytest.importorskip("docx")

    workspace_dir = _make_workspace(workspace_root, "pytest-agent-tools-before-export")

    async def before_export(context):
        context.document.blocks.append(ParagraphBlock(text="Export hook note"))
        return context

    toolset = build_document_toolset(
        workspace_dir=workspace_dir,
        before_export_hooks=[before_export],
    )
    tool_by_name = {tool.name: tool for tool in toolset.tools}

    created = json.loads(
        await tool_by_name["create_document"].call(
            None,
            title="Before Export Hook",
            output_name="before-export-hook.docx",
        )
    )

    exported = json.loads(
        await tool_by_name["export_document"].call(
            None,
            document_id=created["document"]["document_id"],
        )
    )

    loaded_doc = docx.Document(exported["file_path"])
    assert "Export hook note" in [paragraph.text for paragraph in loaded_doc.paragraphs]


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


@pytest.mark.asyncio
async def test_mcp_export_document_runs_before_export_hooks(workspace_root: Path):
    docx = pytest.importorskip("docx")

    workspace_dir = _make_workspace(workspace_root, "pytest-mcp-before-export")

    async def before_export(context):
        context.document.blocks.append(ParagraphBlock(text="MCP export hook note"))
        return context

    server = create_server(
        workspace_dir=workspace_dir,
        before_export_hooks=[before_export],
    )
    _, created_payload = await server.call_tool(
        "create_document",
        {"title": "MCP Hook", "output_name": "mcp-hook.docx"},
    )
    _, exported_payload = await server.call_tool(
        "export_document",
        {"document_id": created_payload["document"]["document_id"]},
    )

    loaded_doc = docx.Document(exported_payload["file_path"])
    assert "MCP export hook note" in [
        paragraph.text for paragraph in loaded_doc.paragraphs
    ]


@pytest.mark.asyncio
async def test_mcp_export_document_runs_after_export_hooks(workspace_root: Path):
    pytest.importorskip("docx")

    workspace_dir = _make_workspace(workspace_root, "pytest-mcp-after-export")
    hook_calls: list[tuple[str, str]] = []

    async def after_export(context):
        hook_calls.append(
            (context.document.status.value, Path(context.output_path).name)
        )
        return context

    server = create_server(
        workspace_dir=workspace_dir,
        after_export_hooks=[after_export],
    )
    _, created_payload = await server.call_tool(
        "create_document",
        {"title": "MCP After Hook", "output_name": "mcp-after-hook.docx"},
    )
    _, exported_payload = await server.call_tool(
        "export_document",
        {"document_id": created_payload["document"]["document_id"]},
    )

    assert exported_payload["success"] is True
    assert hook_calls == [("exported", "mcp-after-hook.docx")]


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
    assert TableRenderer.resolve_docx_table_style("report_grid") == "Table Grid"
    assert (
        TableRenderer.resolve_docx_table_style("metrics_compact")
        == "Light List Accent 1"
    )
    assert TableRenderer.resolve_docx_table_style("minimal") == "Table Grid"
    assert TableRenderer.resolve_docx_table_style("") == "Table Grid"
    assert TableRenderer.resolve_docx_table_style("custom_style") == "custom_style"


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


def test_word_document_builder_preserves_workspace_for_summary_card_group(
    workspace_root: Path, monkeypatch: pytest.MonkeyPatch
):
    docx = pytest.importorskip("docx")

    workspace_dir = _make_workspace(
        workspace_root, "pytest-agent-tools-summary-card-workspace"
    )
    image_path = workspace_dir / "nested.png"
    output_path = workspace_dir / "summary-card-workspace.docx"
    _write_png(image_path, width=320, height=180)

    from astrbot_plugin_office_assistant.document_core.models.blocks import (
        GroupBlock,
        ImageBlock,
        ParagraphBlock,
    )
    from astrbot_plugin_office_assistant.document_core.models.document import (
        DocumentMetadata,
        DocumentModel,
    )

    paragraph = ParagraphBlock(text="Summary block body")
    object.__setattr__(paragraph, "variant", "summary_box")
    object.__setattr__(paragraph, "title", "Summary")

    monkeypatch.setattr(
        "astrbot_plugin_office_assistant.document_core.builders.word_builder.build_summary_card_group",
        lambda **_kwargs: GroupBlock(
            blocks=[
                ImageBlock(
                    path=image_path.name,
                    caption="Nested image caption",
                )
            ]
        ),
    )

    document = DocumentModel(
        document_id="summary-card-workspace-test",
        metadata=DocumentMetadata(title="Summary Card Workspace"),
        blocks=[paragraph],
    )

    WordDocumentBuilder().build(document, output_path)

    loaded_doc = docx.Document(output_path)
    assert len(loaded_doc.inline_shapes) == 1
    assert "Nested image caption" in [
        paragraph.text for paragraph in loaded_doc.paragraphs
    ]


def test_word_document_builder_skips_images_outside_workspace(
    workspace_root: Path,
):
    docx = pytest.importorskip("docx")

    workspace_dir = _make_workspace(workspace_root, "pytest-agent-tools-image-sandbox")
    external_dir = _make_workspace(workspace_root, "pytest-agent-tools-image-external")
    external_image_path = external_dir / "outside.png"
    output_path = workspace_dir / "image-sandbox.docx"
    _write_png(external_image_path, width=400, height=200)

    from astrbot_plugin_office_assistant.document_core.models.blocks import ImageBlock
    from astrbot_plugin_office_assistant.document_core.models.document import (
        DocumentMetadata,
        DocumentModel,
    )

    document = DocumentModel(
        document_id="image-sandbox-test",
        metadata=DocumentMetadata(title="Image Sandbox"),
        blocks=[
            ImageBlock(
                path=str(external_image_path),
                caption="Should be skipped",
            )
        ],
    )

    WordDocumentBuilder().build(document, output_path)

    loaded_doc = docx.Document(output_path)
    assert len(loaded_doc.inline_shapes) == 0
    assert "Should be skipped" not in [
        paragraph.text for paragraph in loaded_doc.paragraphs
    ]


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


def test_document_session_store_runs_internal_normalize_hooks():
    def _inject_heading(_context):
        return [
            BlockHeadingInput(text="Hook Title", level=2),
            SectionParagraphInput(text="正文"),
        ]

    store = DocumentSessionStore(normalize_block_hooks=[_inject_heading])
    document = store.create_document(CreateDocumentRequest(title="Hook Test"))

    updated = store.add_blocks(
        AddBlocksRequest(
            document_id=document.document_id,
            blocks=[{"type": "paragraph", "text": "原始正文"}],
        )
    )

    assert len(updated.blocks) == 2
    assert updated.blocks[0].text == "Hook Title"
    assert updated.blocks[1].text == "正文"


def test_table_schema_normalizers_are_shared():
    request = AddTableRequest(
        document_id="doc-1",
        headers=["区域", "目标"],
        rows=[["华东", "120"]],
        header_groups=[{"title": "经营数据", "span": 2}],
        table_style="invalid-style",
        column_widths=[4.2, 0, -1.0, 3.0],
        numeric_columns=[2, -1, 1, 2],
        header_fill="#1f4e79",
        header_text_color="ffffff",
        banded_rows=True,
        banded_row_fill="eef4fa",
        first_column_bold=True,
        table_align="center",
        border_style="strong",
        caption_emphasis="strong",
    )
    section = SectionTableInput(
        headers=["区域"],
        rows=[["华东"]],
        header_groups=[{"title": "经营概览", "span": 1}],
        table_style="invalid-style",
        column_widths=[4.2, 0, -1.0, 3.0],
        numeric_columns=[2, -1, 1, 2],
        header_fill="1f4e79",
        header_text_color="#ffffff",
        banded_rows=False,
        banded_row_fill="#eef4fa",
        first_column_bold=False,
        table_align="left",
        border_style="minimal",
        caption_emphasis="normal",
    )

    assert request.table_style == ""
    assert request.column_widths == [4.2, 0, 0, 3.0]
    assert request.numeric_columns == [1, 2]
    assert request.header_groups[0].span == 2
    assert request.header_fill == "1F4E79"
    assert request.header_text_color == "FFFFFF"
    assert request.banded_rows is True
    assert request.banded_row_fill == "EEF4FA"
    assert request.first_column_bold is True
    assert request.table_align == "center"
    assert request.border_style == "strong"
    assert request.caption_emphasis == "strong"
    assert section.table_style == ""
    assert section.column_widths == [4.2, 0, 0, 3.0]
    assert section.numeric_columns == [1, 2]
    assert section.header_groups[0].title == "经营概览"
    assert section.header_fill == "1F4E79"
    assert section.header_text_color == "FFFFFF"
    assert section.banded_rows is False
    assert section.banded_row_fill == "EEF4FA"
    assert section.first_column_bold is False
    assert section.table_align == "left"
    assert section.border_style == "minimal"
    assert section.caption_emphasis == "normal"


def test_table_schema_rejects_invalid_grouped_headers():
    with pytest.raises(
        ValidationError,
        match=r"header_groups span total \(1\) must equal column count \(2\)",
    ):
        AddTableRequest(
            document_id="doc-1",
            headers=["区域", "目标"],
            rows=[["华东", "120"]],
            header_groups=[{"title": "经营数据", "span": 1}],
        )

    with pytest.raises(ValidationError, match="greater than or equal to 1"):
        SectionTableInput(
            headers=["区域", "目标"],
            rows=[["华东", "120"]],
            header_groups=[{"title": "经营数据", "span": 0}],
        )

    with pytest.raises(ValidationError, match="6-digit hex color"):
        AddTableRequest(
            document_id="doc-1",
            headers=["区域"],
            rows=[["华东"]],
            header_fill="blue",
        )

    with pytest.raises(ValidationError):
        SectionTableInput(
            headers=["区域"],
            rows=[["华东"]],
            border_style="heavy",
        )

    with pytest.raises(
        ValidationError,
        match=r"header_groups require at least one column from headers or rows \(column_count=0\)",
    ):
        TableBlock(header_groups=[{"title": "经营数据", "span": 1}])


def test_table_schema_allows_empty_placeholder_tables():
    request = AddTableRequest(document_id="doc-1", headers=[], rows=[])
    section = SectionTableInput(headers=[], rows=[])
    block = TableBlock()

    assert request.headers == []
    assert request.rows == []
    assert section.headers == []
    assert section.rows == []
    assert block.headers == []
    assert block.rows == []


def test_document_session_store_preserves_grouped_headers_for_table_blocks():
    store = DocumentSessionStore()
    document = store.create_document(CreateDocumentRequest(title="Grouped Table"))

    updated = store.add_blocks(
        AddBlocksRequest(
            document_id=document.document_id,
            blocks=[
                {
                    "type": "table",
                    "header_groups": [
                        {"title": "经营数据", "span": 2},
                        {"title": "结果", "span": 1},
                    ],
                    "headers": ["区域", "目标", "完成率"],
                    "rows": [["华东", "120", "98%"]],
                    "header_fill": "1F4E79",
                    "header_text_color": "FFFFFF",
                    "banded_rows": True,
                    "banded_row_fill": "EEF4FA",
                    "first_column_bold": True,
                    "table_align": "center",
                    "border_style": "strong",
                    "caption_emphasis": "strong",
                }
            ],
        )
    )

    table = updated.blocks[0]
    assert table.header_groups[0].title == "经营数据"
    assert table.header_groups[0].span == 2
    assert table.header_groups[1].title == "结果"
    assert table.header_fill == "1F4E79"
    assert table.header_text_color == "FFFFFF"
    assert table.banded_rows is True
    assert table.banded_row_fill == "EEF4FA"
    assert table.first_column_bold is True
    assert table.table_align == "center"
    assert table.border_style == "strong"
    assert table.caption_emphasis == "strong"


def test_document_session_store_add_table_preserves_grouped_headers():
    store = DocumentSessionStore()
    document = store.create_document(CreateDocumentRequest(title="Legacy Table"))

    updated = store.add_table(
        AddTableRequest(
            document_id=document.document_id,
            headers=["区域", "目标", "完成率"],
            rows=[["华东", "120", "98%"]],
            header_groups=[
                {"title": "经营数据", "span": 2},
                {"title": "结果", "span": 1},
            ],
            header_fill="1F4E79",
            header_text_color="FFFFFF",
            banded_rows=True,
            banded_row_fill="EEF4FA",
            first_column_bold=True,
            table_align="center",
            border_style="strong",
            caption_emphasis="strong",
        )
    )

    table = updated.blocks[0]
    assert table.header_groups[0].title == "经营数据"
    assert table.header_groups[1].span == 1
    assert table.header_fill == "1F4E79"
    assert table.header_text_color == "FFFFFF"
    assert table.banded_rows is True
    assert table.banded_row_fill == "EEF4FA"
    assert table.first_column_bold is True
    assert table.table_align == "center"
    assert table.border_style == "strong"
    assert table.caption_emphasis == "strong"


@pytest.mark.asyncio
async def test_add_blocks_tool_ignores_blank_table_caption_when_absorbing_heading(
    workspace_root: Path,
):
    docx = pytest.importorskip("docx")

    workspace_dir = _make_workspace(workspace_root, "pytest-agent-table-blank-caption")
    toolset = build_document_toolset(workspace_dir=workspace_dir)
    tool_by_name = {tool.name: tool for tool in toolset.tools}

    created = json.loads(
        await tool_by_name["create_document"].call(
            None,
            title="空白表标题",
            output_name="blank-table-caption.docx",
        )
    )
    document_id = created["document"]["document_id"]

    await tool_by_name["add_blocks"].call(
        None,
        document_id=document_id,
        blocks=[
            {"type": "heading", "text": "季度经营指标总览", "level": 2},
            {
                "type": "table",
                "caption": "   ",
                "headers": ["区域", "营收（万元）"],
                "rows": [["华东", "1280"]],
            },
        ],
    )

    exported = json.loads(
        await tool_by_name["export_document"].call(
            None,
            document_id=document_id,
        )
    )

    loaded_doc = docx.Document(exported["file_path"])
    table = loaded_doc.tables[0]
    assert table.rows[0].cells[0].text == "季度经营指标总览"
