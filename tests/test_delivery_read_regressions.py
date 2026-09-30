from types import SimpleNamespace
import json
from pathlib import Path
from unittest.mock import AsyncMock, MagicMock

import pytest
from openpyxl import Workbook

from astrbot_plugin_office_assistant.services.delivery_service import DeliveryService
from astrbot_plugin_office_assistant.services.post_export_hook_service import (
    PostExportHookService,
)
from astrbot_plugin_office_assistant.services.file_read_service import FileReadService
from astrbot_plugin_office_assistant.services.workspace_service import WorkspaceService
from astrbot_plugin_office_assistant.agent_tools import build_workbook_toolset
from astrbot_plugin_office_assistant.domain.export_artifacts import (
    write_owned_export_metadata,
)


def _delivery_service(service_type=DeliveryService, **overrides):
    options = dict(
        executor=None,
        preview_generator=None,
        enable_preview=False,
        auto_delete=True,
        reply_to_user=False,
    )
    if service_type is PostExportHookService:
        options["exported_message"] = "sent"
    return service_type(**(options | overrides))


@pytest.mark.asyncio
@pytest.mark.parametrize("service_type", [DeliveryService, PostExportHookService])
@pytest.mark.parametrize("failed_component", ["Image", "Plain", "File"])
async def test_delivery_success_tracks_file_and_ancillary_failures_do_not_block_it(
    workspace_root, service_type, failed_component
):
    output = workspace_root / "report.docx"
    output.write_bytes(b"generated document")
    preview = workspace_root / "preview.png"
    preview.write_bytes(b"transport is stubbed")
    attempted = []

    async def send(chain):
        kind = type(chain.chain[0]).__name__
        attempted.append(kind)
        if kind == failed_component:
            raise RuntimeError(f"{kind} failed")

    event = SimpleNamespace(send=send, get_sender_id=lambda: "alice")
    service = _delivery_service(
        service_type,
        preview_generator=MagicMock(generate_preview=MagicMock(return_value=preview)),
        enable_preview=True,
    )
    deliver = (
        service.send_exported_document
        if service_type is PostExportHookService
        else service.send_file_with_preview
    )
    if failed_component == "File":
        with pytest.raises(RuntimeError, match="File failed"):
            await deliver(event, output)
        assert attempted == ["File"]
        assert output.exists()
    else:
        await deliver(event, output)
        assert attempted == ["File", "Plain", "Image"]
        assert not output.exists()
        assert not preview.exists()


@pytest.mark.asyncio
@pytest.mark.parametrize(
    "limits,forbidden",
    [
        ({"max_excel_preview_rows": 2}, "row-6"),
        ({"max_excel_preview_chars": 10}, "row-6"),
        ({"max_excel_preview_sheets": 1}, "SECOND_SHEET_SECRET"),
    ],
)
async def test_generic_and_workbook_reads_share_configured_limits(
    workspace_root, limits, forbidden
):
    path = workspace_root / "source.xlsx"
    workbook = Workbook()
    for row in range(1, 7):
        workbook.active.append([f"row-{row}", row])
    workbook.create_sheet("Other").append(["SECOND_SHEET_SECRET"])
    workbook.save(path)
    workbook.close()
    workspace = WorkspaceService(
        plugin_data_path=workspace_root,
        executor=None,
        office_libs={"openpyxl": True},
        max_file_size=20_000_000,
    )
    service = FileReadService(
        workspace_service=workspace,
        word_read_service=None,
        allow_external_input_files=False,
        is_group_feature_enabled=lambda event: True,
        check_permission=lambda event: True,
        group_feature_disabled_error=lambda: "disabled",
        **limits,
    )
    generic = await service.read_file(None, path.name)
    dedicated = await service.read_workbook(None, path.name)
    assert forbidden not in dedicated
    assert forbidden not in generic
    assert generic == dedicated


@pytest.mark.asyncio
async def test_export_does_not_claim_delivery_when_output_disappeared(
    workspace_root, agent_tool_context
):
    service = _delivery_service(PostExportHookService, auto_delete=False)

    async def lose_file_then_deliver(tool_context, output_path):
        Path(output_path).unlink()
        return await service.handle_exported_document_tool(tool_context, output_path)

    tools = {
        tool.name: tool
        for tool in build_workbook_toolset(workspace_root, lose_file_then_deliver).tools
    }
    created = json.loads(await tools["create_workbook"].call(agent_tool_context))
    result = await tools["export_workbook"].call(
        agent_tool_context, workbook_id=created["workbook"]["workbook_id"]
    )
    assert result is not None
    assert "delivery failed" in json.loads(result)["message"]


@pytest.mark.asyncio
async def test_auto_delete_removes_private_export_metadata_and_directory(
    workspace_root,
):
    private_dir = workspace_root / ".office-export-test"
    private_dir.mkdir()
    output = private_dir / "report.xlsx"
    output.write_bytes(b"test output")
    write_owned_export_metadata(output, ("audit", "alice", "private"), workspace_root)

    await _delivery_service().send_file_with_preview(
        SimpleNamespace(send=AsyncMock()), output
    )
    assert not private_dir.exists()
