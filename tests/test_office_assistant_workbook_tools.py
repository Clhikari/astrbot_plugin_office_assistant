import json
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

import pytest

from astrbot_plugin_office_assistant.agent_tools import (
    build_workbook_toolset,
)
from astrbot_plugin_office_assistant.agent_tools.workbook_tools import (
    CreateWorkbookTool,
    ExportWorkbookTool,
    WriteRowsTool,
)
from astrbot_plugin_office_assistant.constants import (
    EXCEL_SCRIPT_RETRY_EXHAUSTED_EVENT_KEY,
)
from astrbot_plugin_office_assistant.domain.workbook.contracts import (
    CreateWorkbookRequest,
    MAX_WORKBOOK_ROW_INDEX,
)
from astrbot_plugin_office_assistant.domain.workbook.session_store import (
    WorkbookSessionStore,
)
import astrbot_plugin_office_assistant.mcp_server.server as mcp_server_module
from astrbot_plugin_office_assistant.mcp_server.server import (
    create_server,
)

from astrbot.core.utils.astrbot_path import get_astrbot_plugin_data_path


def _build_agent_tool_context(*, excel_script_retry_exhausted: bool = False):
    event = MagicMock()
    event.get_extra.side_effect = lambda key, default=None: (
        excel_script_retry_exhausted
        if key == EXCEL_SCRIPT_RETRY_EXHAUSTED_EVENT_KEY
        else default
    )
    return SimpleNamespace(context=SimpleNamespace(event=event))


def test_build_workbook_toolset_uses_shared_store_and_default_workspace():
    toolset = build_workbook_toolset()
    tool_names = [tool.name for tool in toolset.tools]

    assert tool_names == [
        "create_workbook",
        "write_rows",
        "export_workbook",
    ]

    stores = [tool.store for tool in toolset.tools if hasattr(tool, "store")]
    assert len(stores) == len(tool_names)
    assert len({id(store) for store in stores}) == 1

    expected_workspace = (
        Path(get_astrbot_plugin_data_path())
        / "astrbot_plugin_office_assistant"
        / "workbooks"
    )
    assert stores[0].workspace_dir == expected_workspace


@pytest.mark.asyncio
async def test_write_rows_tool_returns_targeted_retry_message_for_validation_errors(
    workspace_root: Path,
):
    store = WorkbookSessionStore(workspace_dir=workspace_root)
    workbook = store.create_workbook(CreateWorkbookRequest(filename="rows.xlsx"))
    tool = WriteRowsTool(store=store)

    result = json.loads(
        await tool.call(
            None,
            workbook_id=workbook.workbook_id,
            sheet="Data",
            rows=[["=SUM(A1:A2)"]],
        )
    )

    assert result["success"] is False
    assert "only fix invalid fields" in result["message"]
    assert "Original error" in result["message"]


@pytest.mark.asyncio
async def test_create_workbook_tool_wraps_expected_store_errors(
    workspace_root: Path,
    monkeypatch: pytest.MonkeyPatch,
):
    store = WorkbookSessionStore(workspace_dir=workspace_root)
    tool = CreateWorkbookTool(store=store)
    monkeypatch.setattr(
        store, "create_workbook", MagicMock(side_effect=OSError("disk full"))
    )

    result = json.loads(await tool.call(None, filename="rows.xlsx"))

    assert result["success"] is False
    assert result["message"] == "create_workbook failed: disk full"


@pytest.mark.asyncio
async def test_create_workbook_tool_reraises_unexpected_store_errors(
    workspace_root: Path,
    monkeypatch: pytest.MonkeyPatch,
):
    store = WorkbookSessionStore(workspace_dir=workspace_root)
    tool = CreateWorkbookTool(store=store)
    monkeypatch.setattr(
        store, "create_workbook", MagicMock(side_effect=RuntimeError("boom"))
    )

    with pytest.raises(RuntimeError, match="boom"):
        await tool.call(None, filename="rows.xlsx")


@pytest.mark.asyncio
async def test_workbook_tools_refuse_after_excel_script_retry_exhaustion(
    workspace_root: Path,
):
    store = WorkbookSessionStore(workspace_dir=workspace_root)
    workbook = store.create_workbook(CreateWorkbookRequest(filename="rows.xlsx"))
    context = _build_agent_tool_context(excel_script_retry_exhausted=True)

    create_result = json.loads(
        await CreateWorkbookTool(store=store).call(context, filename="new.xlsx")
    )
    write_result = json.loads(
        await WriteRowsTool(store=store).call(
            context,
            workbook_id=workbook.workbook_id,
            sheet="Data",
            rows=[["value"]],
        )
    )
    export_result = json.loads(
        await ExportWorkbookTool(store=store).call(
            context,
            workbook_id=workbook.workbook_id,
        )
    )

    for result in (create_result, write_result, export_result):
        assert result["success"] is False
        assert "Excel 脚本重试次数已用尽" in result["message"]
        assert "请停止调用工具" in result["message"]


@pytest.mark.asyncio
async def test_write_rows_tool_preserves_invalid_zero_start_row(
    workspace_root: Path,
):
    store = WorkbookSessionStore(workspace_dir=workspace_root)
    workbook = store.create_workbook(CreateWorkbookRequest(filename="rows.xlsx"))
    tool = WriteRowsTool(store=store)

    result = json.loads(
        await tool.call(
            None,
            workbook_id=workbook.workbook_id,
            sheet="Data",
            rows=[["value"]],
            start_row=0,
        )
    )

    assert result["success"] is False
    assert "only fix invalid fields" in result["message"]
    assert "greater than or equal to 1" in result["message"]


@pytest.mark.asyncio
async def test_write_rows_tool_rejects_oversized_start_row(
    workspace_root: Path,
):
    store = WorkbookSessionStore(workspace_dir=workspace_root)
    workbook = store.create_workbook(CreateWorkbookRequest(filename="rows.xlsx"))
    tool = WriteRowsTool(store=store)

    result = json.loads(
        await tool.call(
            None,
            workbook_id=workbook.workbook_id,
            sheet="Data",
            rows=[["value"]],
            start_row=MAX_WORKBOOK_ROW_INDEX + 1,
        )
    )

    assert result["success"] is False
    assert "only fix invalid fields" in result["message"]
    assert "less than or equal to" in result["message"]


@pytest.mark.asyncio
async def test_write_rows_tool_returns_neutral_message_for_state_errors(
    workspace_root: Path,
):
    store = WorkbookSessionStore(workspace_dir=workspace_root)
    workbook = store.create_workbook(CreateWorkbookRequest(filename="rows.xlsx"))
    workbook.status = "exported"
    tool = WriteRowsTool(store=store)

    result = json.loads(
        await tool.call(
            None,
            workbook_id=workbook.workbook_id,
            sheet="Data",
            rows=[["value"]],
        )
    )

    assert result["success"] is False
    assert result["message"].startswith("write_rows failed:")
    assert "only fix invalid fields" not in result["message"]


@pytest.mark.asyncio
async def test_write_rows_tool_reraises_unexpected_store_errors(
    workspace_root: Path,
    monkeypatch: pytest.MonkeyPatch,
):
    store = WorkbookSessionStore(workspace_dir=workspace_root)
    workbook = store.create_workbook(CreateWorkbookRequest(filename="rows.xlsx"))
    tool = WriteRowsTool(store=store)
    monkeypatch.setattr(
        store, "write_rows", MagicMock(side_effect=RuntimeError("boom"))
    )

    with pytest.raises(RuntimeError, match="boom"):
        await tool.call(
            None,
            workbook_id=workbook.workbook_id,
            sheet="Data",
            rows=[["value"]],
        )


@pytest.mark.asyncio
async def test_export_workbook_tool_wraps_expected_validation_errors(
    workspace_root: Path,
):
    store = WorkbookSessionStore(workspace_dir=workspace_root)
    workbook = store.create_workbook(CreateWorkbookRequest(filename="rows.xlsx"))
    tool = ExportWorkbookTool(store=store)

    result = json.loads(
        await tool.call(
            None,
            workbook_id=workbook.workbook_id,
            output_name="C:/temp/final.xlsx",
        )
    )

    assert result["success"] is False
    assert result["message"].startswith("export_workbook failed:")
    assert "must not be an absolute path" in result["message"]


@pytest.mark.asyncio
async def test_export_workbook_tool_reraises_unexpected_store_errors(
    workspace_root: Path,
    monkeypatch: pytest.MonkeyPatch,
):
    store = WorkbookSessionStore(workspace_dir=workspace_root)
    workbook = store.create_workbook(CreateWorkbookRequest(filename="rows.xlsx"))
    tool = ExportWorkbookTool(store=store)
    monkeypatch.setattr(
        store, "export_workbook", MagicMock(side_effect=RuntimeError("boom"))
    )

    with pytest.raises(RuntimeError, match="boom"):
        await tool.call(
            None,
            workbook_id=workbook.workbook_id,
        )


@pytest.mark.asyncio
async def test_mcp_write_rows_returns_structured_failure_for_unknown_workbook():
    server = create_server()

    _, payload = await server.call_tool(
        "write_rows",
        {
            "workbook_id": "wb-missing",
            "sheet": "Data",
            "rows": [["value"]],
        },
    )

    assert payload["success"] is False
    assert payload["message"].startswith("write_rows failed:")


@pytest.mark.asyncio
async def test_mcp_write_rows_returns_structured_failure_for_validation_errors():
    server = create_server()
    _, created_payload = await server.call_tool(
        "create_workbook",
        {"filename": "validation.xlsx"},
    )

    _, payload = await server.call_tool(
        "write_rows",
        {
            "workbook_id": created_payload["workbook"]["workbook_id"],
            "sheet": "Data",
            "rows": [["=SUM(A1:A2)"]],
        },
    )

    assert payload["success"] is False
    assert "only fix invalid fields" in payload["message"]


@pytest.mark.asyncio
async def test_mcp_write_rows_rejects_oversized_row_window():
    server = create_server()
    _, created_payload = await server.call_tool(
        "create_workbook",
        {"filename": "validation.xlsx"},
    )

    _, payload = await server.call_tool(
        "write_rows",
        {
            "workbook_id": created_payload["workbook"]["workbook_id"],
            "sheet": "Data",
            "rows": [["value"], ["overflow"]],
            "start_row": MAX_WORKBOOK_ROW_INDEX,
        },
    )

    assert payload["success"] is False
    assert "only fix invalid fields" in payload["message"]
    assert "final written row must not exceed" in payload["message"]


def test_mcp_workbook_loader_logs_import_error_and_disables_support(
    monkeypatch: pytest.MonkeyPatch,
):
    def _raise_import_error(_module_name: str):
        raise ModuleNotFoundError("missing workbook deps")

    monkeypatch.setattr(
        mcp_server_module.importlib,
        "import_module",
        _raise_import_error,
    )

    with patch.object(mcp_server_module.logger, "warning") as warning_mock:
        result = mcp_server_module._load_workbook_session_store()

    assert result is None
    warning_mock.assert_called_once()


def test_mcp_workbook_loader_reraises_non_import_errors(
    monkeypatch: pytest.MonkeyPatch,
):
    def _raise_runtime_error(_module_name: str):
        raise NameError("boom")

    monkeypatch.setattr(
        mcp_server_module.importlib,
        "import_module",
        _raise_runtime_error,
    )

    with patch.object(mcp_server_module.logger, "warning") as warning_mock:
        with pytest.raises(NameError, match="boom"):
            mcp_server_module._load_workbook_session_store()

    warning_mock.assert_not_called()


def test_mcp_create_server_constructs_workbook_store_once(
    monkeypatch: pytest.MonkeyPatch,
):
    constructed_workspaces: list[Path | None] = []

    class StubWorkbookSessionStore:
        def __init__(self, workspace_dir=None):
            constructed_workspaces.append(workspace_dir)

    register_mock = MagicMock()
    monkeypatch.setattr(
        mcp_server_module,
        "_load_workbook_session_store",
        lambda: StubWorkbookSessionStore,
    )
    monkeypatch.setattr(mcp_server_module, "register_workbook_tools", register_mock)

    create_server()

    assert constructed_workspaces == [None]
    register_mock.assert_called_once()


# --- WriteRowsOptions tool-layer tests ---


@pytest.mark.asyncio
async def test_write_rows_tool_passes_options_to_store(workspace_root: Path):
    store = WorkbookSessionStore(workspace_dir=workspace_root)
    workbook = store.create_workbook(CreateWorkbookRequest(filename="opts.xlsx"))
    tool = WriteRowsTool(store=store)

    result = json.loads(
        await tool.call(
            None,
            workbook_id=workbook.workbook_id,
            sheet="Data",
            rows=[["h1", "h2"], ["v1", "v2"]],
            options={"freeze_panes": "A2", "column_widths": {"A": 20}, "autofilter": True},
        )
    )

    assert result["success"] is True
    worksheet = workbook.get_sheet("Data")
    assert worksheet.options.freeze_panes == "A2"
    assert worksheet.options.column_widths == {"A": 20.0}
    assert worksheet.options.autofilter is True


@pytest.mark.asyncio
async def test_write_rows_tool_returns_validation_error_for_bad_options(
    workspace_root: Path,
):
    store = WorkbookSessionStore(workspace_dir=workspace_root)
    workbook = store.create_workbook(CreateWorkbookRequest(filename="opts.xlsx"))
    tool = WriteRowsTool(store=store)

    result = json.loads(
        await tool.call(
            None,
            workbook_id=workbook.workbook_id,
            sheet="Data",
            rows=[["value"]],
            options={"freeze_panes": "INVALID"},
        )
    )

    assert result["success"] is False
    assert "only fix invalid fields" in result["message"]


@pytest.mark.asyncio
async def test_mcp_write_rows_passes_options():
    server = create_server()
    _, created_payload = await server.call_tool(
        "create_workbook",
        {"filename": "opts.xlsx"},
    )

    _, payload = await server.call_tool(
        "write_rows",
        {
            "workbook_id": created_payload["workbook"]["workbook_id"],
            "sheet": "Data",
            "rows": [["h1", "h2"], ["v1", "v2"]],
            "options": {"freeze_panes": "A2", "autofilter": True},
        },
    )

    assert payload["success"] is True


def test_write_rows_schema_options_column_widths_uses_number_type(
    workspace_root: Path,
):
    tool = WriteRowsTool(store=WorkbookSessionStore(workspace_dir=workspace_root))
    options_schema = tool.parameters["properties"]["options"]
    col_widths = options_schema["properties"]["column_widths"]

    assert col_widths["additionalProperties"] == {"type": "number"}

    def _assert_no_anyof_oneof(obj, path="options"):
        if isinstance(obj, dict):
            assert "anyOf" not in obj, f"{path} contains anyOf"
            assert "oneOf" not in obj, f"{path} contains oneOf"
            for key, value in obj.items():
                if isinstance(value, (dict, list)):
                    _assert_no_anyof_oneof(value, f"{path}.{key}")
        elif isinstance(obj, list):
            for i, item in enumerate(obj):
                _assert_no_anyof_oneof(item, f"{path}[{i}]")

    _assert_no_anyof_oneof(options_schema)
