from types import SimpleNamespace
from unittest.mock import AsyncMock, MagicMock

import pytest

from astrbot_plugin_office_assistant.domain.document.contracts import (
    CreateDocumentRequest,
)
from astrbot_plugin_office_assistant.domain.document.session_store import (
    DocumentSessionStore,
)
from astrbot_plugin_office_assistant.domain.workbook.contracts import (
    CreateWorkbookRequest,
    WriteRowsRequest,
)
from astrbot_plugin_office_assistant.domain.workbook.session_store import (
    WorkbookSessionStore,
)
from astrbot_plugin_office_assistant.internal_hooks import NoticeBuildContext
from astrbot_plugin_office_assistant.services.request_hook_service import (
    RequestHookService,
)
from astrbot_plugin_office_assistant.services.runtime_builder import (
    _build_document_summary_lookup,
    _build_request_pipeline_services,
    _build_workbook_summary_lookup,
)


def _event(user="alice", origin="group", platform="audit"):
    return SimpleNamespace(
        get_sender_id=lambda: user,
        get_platform_id=lambda: platform,
        unified_msg_origin=origin,
    )


@pytest.fixture
def drafts(tmp_path):
    result = SimpleNamespace()
    for kind, store_type, request_type in (
        ("document", DocumentSessionStore, CreateDocumentRequest),
        ("workbook", WorkbookSessionStore, CreateWorkbookRequest),
    ):
        store = store_type(workspace_dir=tmp_path / kind)
        draft = getattr(store, f"create_{kind}")(
            request_type(title=f"PRIVATE_{kind.upper()}_TITLE", session_id="spoofed")
        )
        draft._owner_key = ("audit", "alice", "group")
        setattr(result, f"{kind}_toolset", SimpleNamespace(**{f"{kind}_store": store}))
        setattr(result, f"{kind}_id", getattr(draft, f"{kind}_id"))
    result.workbook_toolset.workbook_store.write_rows(
        WriteRowsRequest(
            workbook_id=result.workbook_id,
            sheet="PRIVATE_SHEET_NAME",
            rows=[[1]],
        )
    )
    return result


@pytest.mark.parametrize("kind", ["document", "workbook"])
@pytest.mark.parametrize(
    "other_event",
    [None, _event("bob"), _event(origin="other"), _event(platform="other"), object()],
)
def test_owned_summary_lookup_requires_trusted_matching_event(
    drafts, kind, other_event
):
    builder = (
        _build_document_summary_lookup
        if kind == "document"
        else _build_workbook_summary_lookup
    )
    lookup = builder(getattr(drafts, f"{kind}_toolset"), require_owner=True)
    identifier = getattr(drafts, f"{kind}_id")

    assert lookup(identifier, _event())["title"].startswith("PRIVATE_")
    assert lookup(identifier, other_event) is None
    assert lookup("not-found", _event()) is None


def _notice_context(event, kind, identifier):
    return NoticeBuildContext(
        event=event,
        request=SimpleNamespace(
            prompt=f"继续 {kind}_id={identifier}",
            func_tool=SimpleNamespace(
                names=lambda: {"create_workbook", "write_rows", "export_workbook"}
            ),
        ),
        should_expose=True,
        can_process_upload=False,
        explicit_tool_name=None,
    )


@pytest.mark.asyncio
@pytest.mark.parametrize("kind", ["document", "workbook"])
async def test_production_notice_pipeline_does_not_expose_another_users_summary(
    drafts, kind
):
    pipeline = _build_request_pipeline_services(
        astrbot_context=SimpleNamespace(),
        settings=SimpleNamespace(
            allow_external_input_files=False,
            auto_block_execution_tools=True,
            allow_local_excel_script=False,
            require_at_in_group=True,
        ),
        upload_session_service=SimpleNamespace(
            get_cached_upload_infos=lambda event: [],
            consume_session_notice_once=lambda event, key: True,
            get_attachment_session_key=lambda event: ("audit", "alice", "group"),
        ),
        access_policy_service=SimpleNamespace(
            is_group_feature_enabled=lambda event: True,
            check_permission=lambda event: True,
            is_bot_mentioned=lambda event: True,
        ),
        image_asset_service=SimpleNamespace(list_active_images=lambda key: []),
        document_toolset=drafts.document_toolset,
        workbook_toolset=drafts.workbook_toolset,
        extract_upload_source=AsyncMock(),
        store_uploaded_file=MagicMock(),
    )
    owner_context = _notice_context(_event(), kind, getattr(drafts, f"{kind}_id"))
    other_context = _notice_context(_event("bob"), kind, getattr(drafts, f"{kind}_id"))

    await pipeline.request_hook_service.append_office_tool_guide_notice(owner_context)
    await pipeline.request_hook_service.append_office_tool_guide_notice(other_context)

    owner_notice = "".join(owner_context.notices)
    other_notice = "".join(other_context.notices)
    assert "没有找到" not in owner_notice
    assert "没有找到" in other_notice
    assert "PRIVATE_SHEET_NAME" not in other_notice
    if kind == "workbook":
        assert "PRIVATE_SHEET_NAME" in owner_notice


@pytest.mark.asyncio
async def test_legacy_single_argument_summary_lookup_still_supported():
    lookup = MagicMock(return_value={"status": "draft", "block_count": 2})
    service = RequestHookService(
        auto_block_execution_tools=True,
        get_cached_upload_infos=lambda event: [],
        extract_upload_source=AsyncMock(),
        store_uploaded_file=MagicMock(),
        consume_session_notice_once=lambda event, key: True,
        allow_external_input_files=False,
        lookup_document_summary=lookup,
    )

    await service.append_office_tool_guide_notice(
        _notice_context(_event(), "document", "doc-local")
    )

    lookup.assert_called_once_with("doc-local")
