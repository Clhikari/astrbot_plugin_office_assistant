import asyncio
import json
from types import SimpleNamespace

import pytest
from astrbot.core.agent.run_context import ContextWrapper
from astrbot.core.astr_agent_context import AstrAgentContext
from astrbot.core.astr_agent_tool_exec import FunctionToolExecutor
from astrbot.core.platform.astr_message_event import AstrMessageEvent
from astrbot.core.platform.astrbot_message import AstrBotMessage, MessageMember
from astrbot.core.platform.message_type import MessageType
from astrbot.core.platform.platform_metadata import PlatformMetadata
from astrbot.core.provider import func_tool_manager
from astrbot.core.provider.entities import ProviderRequest
from astrbot.core.provider.func_tool_manager import FunctionToolManager
from astrbot.core.star.context import Context

from astrbot_plugin_office_assistant.agent_tools import (
    build_document_toolset,
    build_workbook_toolset,
)
from astrbot_plugin_office_assistant.services.llm_request_policy import LLMRequestPolicy
from astrbot_plugin_office_assistant.services.runtime_builder import (
    _register_structured_tools,
)


@pytest.fixture
def sdk_tools(workspace_root):
    manager = FunctionToolManager()
    context = Context(
        asyncio.Queue(),
        {},
        None,
        SimpleNamespace(llm_tools=manager),
        None,
        None,
        None,
        None,
        None,
        None,
        None,
    )
    message = AstrBotMessage()
    message.type = MessageType.FRIEND_MESSAGE
    message.self_id = "bot"
    message.message_id = "permission-probe"
    message.sender = MessageMember("alice", "Alice")
    message.message = []
    message.message_str = "创建文档"
    event = AstrMessageEvent(
        message.message_str,
        message,
        PlatformMetadata("audit", "permission regression", "audit-adapter"),
        "session-a",
    )
    document = build_document_toolset(workspace_dir=workspace_root)
    workbook = build_workbook_toolset(workspace_dir=workspace_root)
    return context, manager, event, document, workbook


def _policy(document, workbook, manager, *, permitted=True):
    return LLMRequestPolicy(
        document_toolset=document,
        workbook_toolset=workbook,
        tool_manager=manager,
        require_at_in_group=False,
        is_group_feature_enabled=lambda event: True,
        check_permission=lambda event: permitted,
        is_bot_mentioned=lambda event: False,
        notice_hooks=[],
        tool_exposure_hooks=[],
    )


@pytest.mark.asyncio
@pytest.mark.parametrize("permitted", [False, True])
async def test_registered_structured_tools_respect_exposure_and_active_state(
    sdk_tools, permitted
):
    context, manager, event, document, workbook = sdk_tools
    _register_structured_tools(context, document, workbook)
    manager.get_func("create_document").active = False
    manager.remove_func("write_rows")
    req = ProviderRequest(prompt="创建文档", func_tool=manager.get_full_tool_set())
    # A prior hook or provider request must not resurrect disabled/missing tools.
    req.func_tool.add_tool(workbook.get_tool("write_rows"))
    await _policy(document, workbook, manager, permitted=permitted).apply(event, req)
    assert req.func_tool.get_tool("create_document") is None
    assert req.func_tool.get_tool("write_rows") is None
    assert (req.func_tool.get_tool("create_workbook") is not None) is permitted
    if not permitted:
        assert req.func_tool.empty()


@pytest.mark.asyncio
@pytest.mark.skipif(
    not hasattr(func_tool_manager, "_PermissionGuardedTool"),
    reason="AstrBot before 4.28 has no per-tool permission guard",
)
@pytest.mark.parametrize(
    ("name", "arguments"),
    [
        ("create_document", {"session_id": "forged", "title": "private"}),
        ("create_workbook", {"session_id": "forged", "filename": "private.xlsx"}),
    ],
)
async def test_latest_sdk_permission_guard_survives_policy_and_actual_executor(
    sdk_tools, monkeypatch, name, arguments
):
    context, manager, event, document, workbook = sdk_tools
    _register_structured_tools(context, document, workbook)
    permissions = {"_default": {name: "admin"}}

    async def global_get(key, default=None):
        return permissions if key == "tool_permissions" else default

    # Only the preference storage is isolated; SDK guarding and dispatch stay real.
    monkeypatch.setattr(func_tool_manager, "sp", SimpleNamespace(global_get=global_get))
    req = ProviderRequest(prompt="创建文档", func_tool=manager.get_full_tool_set())
    await _policy(document, workbook, manager).apply(event, req)
    tool = req.func_tool.get_tool(name)
    assert type(tool).__name__ == "_PermissionGuardedTool"
    run_context = ContextWrapper(context=AstrAgentContext(context=context, event=event))

    # Reuse the exposed wrapper: permission must be checked anew during dispatch.
    for role in ("member", "admin"):
        event.role = role
        results = [
            result
            async for result in FunctionToolExecutor.execute(
                tool, run_context, **arguments
            )
        ]
        assert len(results) == 1
        if role == "member":
            assert "Permission denied" in results[0].content[0].text
        else:
            payload = json.loads(results[0].content[0].text)
            assert payload["success"] is True
            summary = payload["document" if name == "create_document" else "workbook"]
            assert summary["session_id"] == event.unified_msg_origin


def test_public_registration_resolves_plugin_module_and_keeps_tool_state(
    sdk_tools, monkeypatch
):
    context, manager, _, document, workbook = sdk_tools
    document.get_tool("create_document").active = False
    # Use the module prefix produced by AstrBot's normal plugin loader.
    monkeypatch.setattr(
        type(document.tools[0]),
        "__module__",
        "data.plugins.astrbot_plugin_office_assistant.agent_tools.document_tools",
    )

    _register_structured_tools(context, document, workbook)

    expected_path = "data.plugins.astrbot_plugin_office_assistant.main"
    for tool in [*document.tools, *workbook.tools]:
        assert manager.get_func(tool.name) is tool
        assert tool.handler_module_path == expected_path
    assert manager.get_func("create_document").active is False
