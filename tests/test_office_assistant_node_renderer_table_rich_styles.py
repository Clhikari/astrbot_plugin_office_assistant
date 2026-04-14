from pathlib import Path

from tests._docx_test_helpers import (
    _business_report_metadata,
    _cell_border_color,
    _cell_border_size,
    _raw_row_cell_vertical_merge,
    _render_structured_payload_with_node,
    _table_border_color,
    _table_border_size,
)


def _run_has_prop(run, prop_name: str) -> bool:
    from docx.oxml.ns import qn

    r_pr = run._r.rPr
    if r_pr is None:
        return False
    prop = r_pr.find(qn(f"w:{prop_name}"))
    if prop is None:
        return False
    value = prop.get(qn("w:val"))
    if value is None:
        return True
    return value not in {"0", "false", "off", "none"}


def _run_font_name(run) -> str | None:
    from docx.oxml.ns import qn

    r_pr = run._r.rPr
    if r_pr is None:
        return None
    fonts = r_pr.find(qn("w:rFonts"))
    if fonts is None:
        return None
    return fonts.get(qn("w:ascii"))


def _run_color(run) -> str | None:
    from docx.oxml.ns import qn

    r_pr = run._r.rPr
    if r_pr is None:
        return None
    color = r_pr.find(qn("w:color"))
    if color is None:
        return None
    return color.get(qn("w:val"))


def _run_size_pt(run) -> float | None:
    from docx.oxml.ns import qn

    r_pr = run._r.rPr
    if r_pr is None:
        return None
    size = r_pr.find(qn("w:sz"))
    if size is None:
        return None
    raw_value = size.get(qn("w:val"))
    if raw_value is None:
        return None
    return int(raw_value) / 2


def test_node_renderer_supports_table_cell_rich_style_controls(
    workspace_root: Path,
):
    loaded_doc, _ = _render_structured_payload_with_node(
        workspace_root,
        "pytest-node-renderer-table-rich-style-controls",
        {
            "document_id": "table-rich-style-controls",
            "metadata": _business_report_metadata(title="表格富样式"),
            "blocks": [
                {
                    "type": "table",
                    "headers": ["项目", "富文本", "备注"],
                    "border_style": "standard",
                    "body_font_scale": 0.9,
                    "rows": [
                        [
                            "旧式单元格",
                            {
                                "text": "占位文本",
                                "font_name": "SimSun",
                                "font_scale": 1.2,
                                "bold": True,
                                "italic": True,
                                "underline": True,
                                "strikethrough": True,
                                "text_color": "666666",
                                "border": {
                                    "top": {
                                        "style": "double",
                                        "color": "1F4E79",
                                        "width_pt": 1.0,
                                    },
                                    "right": {
                                        "style": "dashed",
                                        "color": "0F766E",
                                        "width_pt": 0.75,
                                    },
                                },
                                "runs": [
                                    {"text": "默认"},
                                    {
                                        "text": "覆盖",
                                        "font_name": "Consolas",
                                        "font_scale": 1.5,
                                        "bold": False,
                                        "italic": False,
                                        "underline": False,
                                        "strikethrough": False,
                                        "color": "B91C1C",
                                    },
                                ],
                            },
                            "普通文本",
                        ],
                        [
                            "兼容",
                            "文本",
                            "表格",
                        ],
                    ],
                }
            ],
        },
    )

    table = loaded_doc.tables[0]
    plain_cell = table.rows[1].cells[0]
    rich_cell = table.rows[1].cells[1]
    fallback_cell = table.rows[2].cells[2]

    assert plain_cell.text == "旧式单元格"
    assert rich_cell.text == "默认覆盖"
    assert fallback_cell.text == "表格"

    plain_run = plain_cell.paragraphs[0].runs[0]
    rich_first_run = rich_cell.paragraphs[0].runs[0]
    rich_second_run = rich_cell.paragraphs[0].runs[1]

    assert _run_font_name(rich_first_run) == "SimSun"
    assert _run_font_name(rich_second_run) == "Consolas"
    assert _run_color(rich_first_run) == "666666"
    assert _run_color(rich_second_run) == "B91C1C"
    assert _run_has_prop(rich_first_run, "b") is True
    assert _run_has_prop(rich_first_run, "i") is True
    assert _run_has_prop(rich_first_run, "u") is True
    assert _run_has_prop(rich_first_run, "strike") is True
    assert _run_has_prop(rich_second_run, "b") is False
    assert _run_has_prop(rich_second_run, "i") is False
    assert _run_has_prop(rich_second_run, "u") is False
    assert _run_has_prop(rich_second_run, "strike") is False
    assert _run_size_pt(rich_first_run) is not None
    assert _run_size_pt(rich_second_run) is not None
    assert _run_size_pt(rich_second_run) > _run_size_pt(rich_first_run)
    assert _run_size_pt(rich_first_run) > _run_size_pt(plain_run)

    assert _cell_border_color(rich_cell, "top") == "1F4E79"
    assert _cell_border_color(rich_cell, "right") == "0F766E"
    assert _cell_border_color(rich_cell, "left") == _table_border_color(table, "left")
    assert _cell_border_color(rich_cell, "bottom") == _table_border_color(table, "bottom")
    assert _cell_border_size(rich_cell, "top") == "8"
    assert _cell_border_size(rich_cell, "right") == "6"
    assert _cell_border_size(rich_cell, "left") == _table_border_size(table, "left")
    assert _cell_border_size(rich_cell, "bottom") == _table_border_size(table, "bottom")


def test_node_renderer_keeps_text_only_table_compatible(workspace_root: Path):
    loaded_doc, _ = _render_structured_payload_with_node(
        workspace_root,
        "pytest-node-renderer-table-text-only-compatibility",
        {
            "document_id": "table-text-only-compatibility",
            "metadata": _business_report_metadata(title="纯文本表格"),
            "blocks": [
                {
                    "type": "table",
                    "headers": ["A", "B"],
                    "border_style": "minimal",
                    "rows": [
                        ["第一行", "第二列"],
                        ["第二行", "第三列"],
                    ],
                }
            ],
        },
    )

    table = loaded_doc.tables[0]

    assert table.rows[1].cells[0].text == "第一行"
    assert table.rows[1].cells[1].text == "第二列"
    assert table.rows[2].cells[0].text == "第二行"
    assert table.rows[2].cells[1].text == "第三列"


def test_node_renderer_supports_runs_only_cells_beneath_vertical_merge(
    workspace_root: Path,
):
    loaded_doc, _ = _render_structured_payload_with_node(
        workspace_root,
        "pytest-node-renderer-table-runs-only-under-rowspan",
        {
            "document_id": "table-runs-only-under-rowspan",
            "metadata": _business_report_metadata(title="行合并下的富文本单元格"),
            "blocks": [
                {
                    "type": "table",
                    "headers": ["日期", "时间", "课程"],
                    "rows": [
                        [
                            {"text": "第一天", "row_span": 2},
                            {"runs": [{"text": "09:00"}]},
                            "课程 A",
                        ],
                        [
                            {"runs": [{"text": "13:00"}]},
                            "课程 B",
                        ],
                    ],
                }
            ],
        },
    )

    table = loaded_doc.tables[0]

    assert table.rows[1].cells[1].text == "09:00"
    assert table.rows[2].cells[1].text == "13:00"
    assert _raw_row_cell_vertical_merge(table.rows[2], 0) == "continue"


def test_node_renderer_honors_empty_table_cell_border_side_defaults(
    workspace_root: Path,
):
    loaded_doc, _ = _render_structured_payload_with_node(
        workspace_root,
        "pytest-node-renderer-default-table-cell-border-side",
        {
            "document_id": "default-table-cell-border-side",
            "metadata": _business_report_metadata(title="默认单元格边框"),
            "blocks": [
                {
                    "type": "table",
                    "headers": ["区域", "说明"],
                    "rows": [
                        [
                            "华东",
                            {
                                "text": "默认单元格边框",
                                "border": {
                                    "top": {},
                                },
                            },
                        ]
                    ],
                }
            ],
        },
    )

    table = loaded_doc.tables[0]
    cell = table.rows[1].cells[1]
    assert _cell_border_size(cell, "top") == "4"
    assert _cell_border_color(cell, "top") is None
