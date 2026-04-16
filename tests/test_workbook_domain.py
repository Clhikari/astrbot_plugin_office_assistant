from pathlib import Path

import pytest
from openpyxl import load_workbook
from pydantic import ValidationError

from astrbot_plugin_office_assistant.domain.workbook.contracts import (
    CreateWorkbookRequest,
    ExportWorkbookRequest,
    WriteRowsRequest,
    build_workbook_summary,
)
from astrbot_plugin_office_assistant.domain.workbook.session_store import (
    WorkbookSessionStore,
)


def test_create_workbook_returns_draft_summary(workspace_root: Path):
    store = WorkbookSessionStore(workspace_dir=workspace_root)
    workbook = store.create_workbook(
        CreateWorkbookRequest(
            session_id="pytest-session",
            title="季度销量汇总",
            filename="sales-report",
        )
    )

    summary = store.build_prompt_summary(workbook.workbook_id)

    assert workbook.workbook_id == "wb-1"
    assert workbook.metadata.title == "季度销量汇总"
    assert workbook.metadata.preferred_filename == "sales-report.xlsx"
    assert summary == {
        "workbook_id": workbook.workbook_id,
        "title": "季度销量汇总",
        "status": "draft",
        "sheet_names": [],
        "sheet_count": 0,
        "latest_written_sheets": [],
        "next_allowed_actions": ["write_rows", "export_workbook"],
    }


def test_write_rows_creates_sheet_and_overwrites_range(workspace_root: Path):
    store = WorkbookSessionStore(workspace_dir=workspace_root)
    workbook = store.create_workbook(CreateWorkbookRequest(filename="rows.xlsx"))

    store.write_rows(
        WriteRowsRequest(
            workbook_id=workbook.workbook_id,
            sheet="Summary",
            rows=[
                ["name", "amount"],
                ["A", 10],
            ],
            start_row=1,
        )
    )
    store.write_rows(
        WriteRowsRequest(
            workbook_id=workbook.workbook_id,
            sheet="Summary",
            rows=[
                ["A", 99],
                ["B", 35],
            ],
            start_row=2,
        )
    )

    loaded = store.require_workbook(workbook.workbook_id)
    worksheet = loaded.get_sheet("Summary")
    assert worksheet is not None
    assert worksheet.rows == [
        ["name", "amount"],
        ["A", 99],
        ["B", 35],
    ]

    prompt_summary = store.build_prompt_summary(workbook.workbook_id)
    assert prompt_summary["sheet_names"] == ["Summary"]
    assert prompt_summary["latest_written_sheets"] == ["Summary"]


def test_write_rows_rejects_non_primitive_cell_values():
    with pytest.raises(ValidationError):
        WriteRowsRequest(
            workbook_id="wb-1",
            sheet="Data",
            rows=[[{"invalid": True}]],  # type: ignore[list-item]
            start_row=1,
        )


def test_write_rows_rejects_formula_cells():
    with pytest.raises(ValidationError, match="not supported"):
        WriteRowsRequest(
            workbook_id="wb-1",
            sheet="Data",
            rows=[["=SUM(A1:A2)"]],
            start_row=1,
        )


def test_export_request_rejects_absolute_output_name():
    with pytest.raises(ValidationError, match="must not be an absolute path"):
        ExportWorkbookRequest(
            workbook_id="wb-1",
            output_name="C:/temp/final.xlsx",
        )


def test_export_rejects_output_path_outside_workspace(workspace_root: Path):
    store = WorkbookSessionStore(workspace_dir=workspace_root)
    workbook = store.create_workbook(CreateWorkbookRequest(filename="status.xlsx"))
    escaped_request = ExportWorkbookRequest.model_construct(
        workbook_id=workbook.workbook_id,
        output_name="../evil.xlsx",
    )

    with pytest.raises(ValueError, match="escape the workbook workspace"):
        store.export_workbook(escaped_request)


def test_exporter_writes_values_and_header_style(workspace_root: Path):
    store = WorkbookSessionStore(workspace_dir=workspace_root)
    workbook = store.create_workbook(CreateWorkbookRequest(filename="styled.xlsx"))
    store.write_rows(
        WriteRowsRequest(
            workbook_id=workbook.workbook_id,
            sheet="Sales",
            rows=[
                ["name", "amount"],
                ["A", 123.5],
            ],
        )
    )

    exported_workbook, output_path = store.export_workbook(
        ExportWorkbookRequest(
            workbook_id=workbook.workbook_id,
            output_name="styled-output.xlsx",
        )
    )

    loaded = load_workbook(output_path)
    sheet = loaded["Sales"]

    assert exported_workbook.status.value == "exported"
    assert output_path.name == "styled-output.xlsx"
    assert sheet["A1"].value == "name"
    assert sheet["B2"].value == 123.5
    assert sheet["A1"].font.bold is True
    assert sheet["A1"].fill.patternType == "solid"
    assert (sheet["A1"].fill.fgColor.rgb or "").endswith("D9D9D9")
    assert sheet["A1"].alignment.horizontal == "left"
    assert sheet["A1"].alignment.vertical == "center"
    assert sheet["A2"].alignment.horizontal == "left"
    assert sheet["A2"].font.bold is False


def test_exported_workbook_disallows_more_writes_or_reexport(workspace_root: Path):
    store = WorkbookSessionStore(workspace_dir=workspace_root)
    workbook = store.create_workbook(CreateWorkbookRequest(filename="status.xlsx"))
    store.write_rows(
        WriteRowsRequest(
            workbook_id=workbook.workbook_id,
            sheet="Data",
            rows=[["h1"], ["v1"]],
        )
    )
    store.export_workbook(ExportWorkbookRequest(workbook_id=workbook.workbook_id))

    with pytest.raises(ValueError, match="status is draft"):
        store.write_rows(
            WriteRowsRequest(
                workbook_id=workbook.workbook_id,
                sheet="Data",
                rows=[["v2"]],
                start_row=3,
            )
        )

    with pytest.raises(ValueError, match="status is draft"):
        store.export_workbook(
            ExportWorkbookRequest(workbook_id=workbook.workbook_id)
        )

    summary = store.build_prompt_summary(workbook.workbook_id)
    assert summary["status"] == "exported"
    assert summary["next_allowed_actions"] == []


def test_build_workbook_summary_includes_sheet_names(workspace_root: Path):
    store = WorkbookSessionStore(workspace_dir=workspace_root)
    workbook = store.create_workbook(CreateWorkbookRequest(filename="book.xlsx"))
    store.write_rows(
        WriteRowsRequest(
            workbook_id=workbook.workbook_id,
            sheet="One",
            rows=[["c1"], ["v1"]],
        )
    )
    store.write_rows(
        WriteRowsRequest(
            workbook_id=workbook.workbook_id,
            sheet="Two",
            rows=[["c1"], ["v2"]],
        )
    )

    summary = build_workbook_summary(store.require_workbook(workbook.workbook_id))

    assert summary.workbook_id == workbook.workbook_id
    assert summary.sheet_count == 2
    assert summary.sheet_names == ["One", "Two"]
    assert summary.latest_written_sheets == ["One", "Two"]
