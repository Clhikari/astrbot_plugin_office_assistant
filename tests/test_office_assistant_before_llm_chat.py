import contextlib
import json
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import AsyncMock, MagicMock, patch

import pytest
from astrbot_plugin_office_assistant.constants import DOC_COMMAND_TRIGGER_EVENT_KEY
from astrbot_plugin_office_assistant.domain.document.contracts import (
    CreateDocumentRequest,
)
from astrbot_plugin_office_assistant.domain.workbook.contracts import (
    CreateWorkbookRequest,
)
from astrbot_plugin_office_assistant.internal_hooks import (
    NoticeBuildContext,
    ToolExposureContext,
)
from astrbot_plugin_office_assistant.main import FileOperationPlugin
from astrbot_plugin_office_assistant.message_buffer import BufferedMessage
from astrbot_plugin_office_assistant.services.llm_request_policy import (
    LLMRequestPolicy,
)
from astrbot_plugin_office_assistant.services.request_hook_service import (
    RequestHookService,
)
from astrbot_plugin_office_assistant.services.prompt_context_service import (
    SECTION_STATIC_DOCUMENT_TOOLS,
)
from astrbot_plugin_office_assistant.services.upload_session_service import (
    UploadSessionService,
)

import astrbot.api.message_components as Comp
from astrbot.core.agent.tool import FunctionTool, ToolSet
from astrbot.core.message.message_event_result import MessageEventResult
from astrbot.core.platform.message_type import MessageType
from astrbot.core.provider.entities import ProviderRequest
from conftest import build_notice_once_callback as _build_notice_once_callback

_REQUEST_FUNC_TOOL_UNSET = object()


def _build_config() -> dict:
    return {
        "file_settings": {
            "auto_delete_files": True,
            "max_file_size_mb": 20,
            "message_buffer_seconds": 4,
        },
        "trigger_settings": {
            "reply_to_user": True,
            "require_at_in_group": True,
            "enable_features_in_group": False,
            "auto_block_execution_tools": True,
        },
        "preview_settings": {
            "enable": False,
            "dpi": 150,
        },
        "path_settings": {
            "allow_external_input_files": False,
        },
        "permission_settings": {
            "whitelist_users": ["user-1"],
        },
    }


def _build_event(
    *,
    message_type=MessageType.FRIEND_MESSAGE,
    sender_id: str = "user-1",
    is_admin: bool = False,
):
    event = MagicMock()
    extras: dict[str, object] = {}
    event.message_obj = SimpleNamespace(type=message_type, message=[], self_id="bot-1")
    event.get_sender_id.return_value = sender_id
    event.get_platform_id.return_value = "platform-1"
    event.is_admin.return_value = is_admin
    event.unified_msg_origin = "session-1"
    event.message_str = ""
    event._buffer_reentry_count = 0
    event.set_extra.side_effect = lambda key, value: extras.__setitem__(key, value)
    event.get_extra.side_effect = lambda key=None, default=None: (
        dict(extras) if key is None else extras.get(key, default)
    )
    return event


def _tool(name: str) -> FunctionTool:
    return FunctionTool(
        name=name,
        description=f"tool {name}",
        parameters={"type": "object", "properties": {}},
        handler=None,
    )


def _build_provider_request(
    prompt: str,
    *,
    tool_names: list[str] | tuple[str, ...] | None = None,
    system_prompt: str = "base",
    func_tool=_REQUEST_FUNC_TOOL_UNSET,
):
    if func_tool is _REQUEST_FUNC_TOOL_UNSET:
        func_tool = ToolSet([_tool(name) for name in tool_names or []])
    return ProviderRequest(
        prompt=prompt,
        system_prompt=system_prompt,
        func_tool=func_tool,
    )


@contextlib.asynccontextmanager
async def _managed_plugin(*, context=None, config=None):
    effective_context = context or MagicMock()
    plugin = FileOperationPlugin(
        context=effective_context,
        config=config or _build_config(),
    )
    try:
        yield SimpleNamespace(context=effective_context, plugin=plugin)
    finally:
        await plugin.terminate()


def _build_request_hook_service(
    *,
    auto_block_execution_tools: bool = True,
    get_cached_upload_infos=None,
    extract_upload_source=None,
    store_uploaded_file=None,
    consume_session_notice_once=None,
    allow_external_input_files: bool = False,
):
    return RequestHookService(
        auto_block_execution_tools=auto_block_execution_tools,
        get_cached_upload_infos=get_cached_upload_infos or (lambda _event: []),
        extract_upload_source=extract_upload_source or AsyncMock(),
        store_uploaded_file=store_uploaded_file or MagicMock(),
        consume_session_notice_once=consume_session_notice_once
        or _build_notice_once_callback(),
        allow_external_input_files=allow_external_input_files,
    )


def _configure_uploaded_file(
    plugin: FileOperationPlugin,
    *,
    source_path: Path,
    original_name: str,
    stored_name: str,
):
    async def _fake_extract_upload_source(_component):
        return source_path, original_name

    plugin._extract_upload_source = _fake_extract_upload_source
    plugin._store_uploaded_file = lambda *_args, **_kwargs: Path(stored_name)


@pytest.mark.asyncio
async def test_before_llm_chat_injects_document_tools_per_request():
    async with _managed_plugin() as managed:
        event = _build_event(
            message_type=MessageType.FRIEND_MESSAGE, sender_id="user-1"
        )
        req = _build_provider_request(
            "请生成一份 Word 报告，并导出给我。",
            tool_names=["existing_tool", "astrbot_execute_shell"],
        )

        await managed.plugin.before_llm_chat(event, req)

        tool_names = set(req.func_tool.names())
        assert "existing_tool" in tool_names
        assert "astrbot_execute_shell" not in tool_names
        assert {
            "create_document",
            "add_blocks",
            "finalize_document",
            "export_document",
        }.issubset(tool_names)
        assert "generate_complex_word_document" not in tool_names
        assert "文件工具使用指南" in req.prompt
        assert "文件工具细节指南" not in req.prompt
        assert "MUST 持续调用直到 `export_document` 成功" in req.prompt
        assert "NEVER 调用网络搜索" in req.prompt
        assert "文件工具使用指南" not in req.system_prompt
        assert req.prompt.startswith("请生成一份 Word 报告，并导出给我。")


@pytest.mark.asyncio
async def test_before_llm_chat_keeps_document_id_follow_up_prompt_lightweight():
    async with _managed_plugin() as managed:
        event = _build_event(
            message_type=MessageType.FRIEND_MESSAGE, sender_id="user-1"
        )
        req = _build_provider_request(
            '继续完善 document_id="doc-1" 的内容',
            tool_names=["existing_tool", "create_document"],
        )
        document_store = managed.plugin._runtime.document_toolset.document_store
        document = document_store.create_document(
            CreateDocumentRequest(title="季度经营复盘")
        )

        req.prompt = f'继续完善 document_id="{document.document_id}" 的内容'

        await managed.plugin.before_llm_chat(event, req)

        assert "文件工具使用指南" not in req.system_prompt
        assert "当前文档状态摘要" not in req.system_prompt
        assert "文件工具使用指南" not in req.prompt
        assert "当前文档状态摘要" not in req.prompt
        assert "文件工具细节指南" not in req.prompt


@pytest.mark.asyncio
async def test_before_llm_chat_injects_workbook_tools_per_request():
    async with _managed_plugin() as managed:
        event = _build_event(
            message_type=MessageType.FRIEND_MESSAGE, sender_id="user-1"
        )
        req = _build_provider_request(
            "请生成一个 Excel 汇总表，分多 sheet 输出给我。",
            tool_names=["existing_tool", "astrbot_execute_shell"],
        )

        await managed.plugin.before_llm_chat(event, req)

        tool_names = set(req.func_tool.names())
        assert "existing_tool" in tool_names
        assert "astrbot_execute_shell" not in tool_names
        assert {
            "create_workbook",
            "write_rows",
            "export_workbook",
        }.issubset(tool_names)
        assert "Excel 原语工具使用指南" in req.prompt
        assert "create_workbook" in req.prompt
        assert "write_rows" in req.prompt
        assert "export_workbook" in req.prompt


@pytest.mark.asyncio
async def test_before_llm_chat_keeps_workbook_id_follow_up_prompt_lightweight():
    async with _managed_plugin() as managed:
        event = _build_event(
            message_type=MessageType.FRIEND_MESSAGE, sender_id="user-1"
        )
        req = _build_provider_request(
            '继续补充 workbook_id="wb-1" 的数据',
            tool_names=["existing_tool", "create_workbook"],
        )
        workbook_store = managed.plugin._runtime.workbook_toolset.workbook_store
        workbook = workbook_store.create_workbook(
            CreateWorkbookRequest(filename="sales-summary.xlsx")
        )

        req.prompt = f'继续补充 workbook_id="{workbook.workbook_id}" 的数据'

        await managed.plugin.before_llm_chat(event, req)

        assert "Excel 原语工具使用指南" not in req.system_prompt
        assert "Excel 原语工具使用指南" not in req.prompt
        assert "当前工作簿阶段" in req.prompt
        assert "`create_workbook` → `write_rows`" not in req.prompt


@pytest.mark.asyncio
async def test_before_llm_chat_hides_execute_excel_script_when_runtime_is_none():
    context = MagicMock()
    context.get_config.side_effect = lambda *args, **kwargs: {
        "provider_settings": {"computer_use_runtime": "none"}
    }
    async with _managed_plugin(context=context) as managed:
        event = _build_event(
            message_type=MessageType.FRIEND_MESSAGE, sender_id="user-1"
        )
        req = _build_provider_request(
            "请生成一个带公式和条件格式的 Excel 报表",
            tool_names=[
                "existing_tool",
                "execute_excel_script",
                "astrbot_execute_shell",
            ],
        )

        await managed.plugin.before_llm_chat(event, req)

        tool_names = set(req.func_tool.names())
        assert "existing_tool" in tool_names
        assert "astrbot_execute_shell" not in tool_names
        assert "execute_excel_script" not in tool_names
        assert "Excel 脚本工具当前不可用" not in req.prompt
        assert "无法完成新增公式或导出新版本" not in req.prompt
        assert "Excel 路径选择规则" in req.prompt


@pytest.mark.asyncio
async def test_before_llm_chat_hides_execute_excel_script_when_runtime_is_local_and_execution_tools_are_blocked():
    context = MagicMock()
    context.get_config.side_effect = lambda *args, **kwargs: {
        "provider_settings": {"computer_use_runtime": "local"}
    }
    async with _managed_plugin(context=context) as managed:
        event = _build_event(
            message_type=MessageType.FRIEND_MESSAGE, sender_id="user-1"
        )
        req = _build_provider_request(
            "请生成一个带公式和条件格式的 Excel 报表",
            tool_names=[
                "existing_tool",
                "execute_excel_script",
                "astrbot_execute_shell",
            ],
        )

        await managed.plugin.before_llm_chat(event, req)

        tool_names = set(req.func_tool.names())
        assert "existing_tool" in tool_names
        assert "astrbot_execute_shell" not in tool_names
        assert "execute_excel_script" not in tool_names
        assert "Excel 脚本工具当前不可用" not in req.prompt
        assert "无法完成新增公式或导出新版本" not in req.prompt
        assert "Excel 路径选择规则" in req.prompt


@pytest.mark.asyncio
async def test_before_llm_chat_skips_document_guide_for_generic_prompt():
    async with _managed_plugin() as managed:
        event = _build_event(
            message_type=MessageType.FRIEND_MESSAGE, sender_id="user-1"
        )
        req = _build_provider_request(
            "hello",
            tool_names=["existing_tool", "astrbot_execute_shell"],
        )

        await managed.plugin.before_llm_chat(event, req)

        tool_names = set(req.func_tool.names())
        assert "existing_tool" in tool_names
        assert "astrbot_execute_shell" not in tool_names
        assert {
            "create_document",
            "add_blocks",
            "finalize_document",
            "export_document",
        }.issubset(tool_names)
        assert "文件工具使用指南" not in req.system_prompt
        assert "文件工具使用指南" not in req.prompt


@pytest.mark.asyncio
async def test_before_llm_chat_injects_document_guide_for_buffered_word_instruction():
    async with _managed_plugin() as managed:
        event = _build_event(
            message_type=MessageType.FRIEND_MESSAGE,
            sender_id="user-1",
        )
        event._buffered = True
        req = _build_provider_request(
            (
                "[System Notice] 用户上传了 1 个文件\n\n"
                "[文件信息]\n"
                "- 原始文件名: source.docx\n"
                "  工作区文件名: source_1.docx\n\n"
                "[用户指令]\n"
                "请根据我刚上传的文档整理成正式汇报，并导出成 Word 发给我。\n\n"
                "[处理要求]\n"
                "1. 优先围绕这些上传文件完成用户请求。\n"
            ),
            tool_names=["existing_tool", "astrbot_execute_shell"],
        )

        await managed.plugin.before_llm_chat(event, req)

        tool_names = set(req.func_tool.names())
        assert "existing_tool" in tool_names
        assert "astrbot_execute_shell" not in tool_names
        assert {
            "create_document",
            "add_blocks",
            "finalize_document",
            "export_document",
        }.issubset(tool_names)
        assert "文件工具使用指南" in req.prompt
        assert "文件工具细节指南" not in req.prompt
        assert "文件工具使用指南" not in req.system_prompt


@pytest.mark.asyncio
async def test_before_llm_chat_removes_file_tools_without_permission():
    async with _managed_plugin() as managed:
        event = _build_event(
            message_type=MessageType.FRIEND_MESSAGE, sender_id="user-2"
        )
        req = _build_provider_request(
            "hello",
            tool_names=[
                "read_file",
                "create_document",
                "export_document",
                "existing_tool",
            ],
        )

        await managed.plugin.before_llm_chat(event, req)

        tool_names = set(req.func_tool.names())
        assert "existing_tool" in tool_names
        assert "read_file" not in tool_names
        assert "create_document" not in tool_names
        assert "export_document" not in tool_names
        assert "add_blocks" not in tool_names
        assert "generate_complex_word_document" not in tool_names
        assert "当前聊天不可使用文件/Office/PDF 相关功能" in req.system_prompt
        assert "`astrbot_execute_python`" in req.system_prompt


@pytest.mark.asyncio
async def test_before_llm_chat_warns_when_group_feature_disabled():
    async with _managed_plugin() as managed:
        event = _build_event(message_type=MessageType.GROUP_MESSAGE, sender_id="user-1")
        req = _build_provider_request(
            "请生成一份 Word 报告",
            tool_names=[
                "read_file",
                "create_document",
                "astrbot_execute_python",
                "existing_tool",
            ],
        )

        await managed.plugin.before_llm_chat(event, req)

        tool_names = set(req.func_tool.names())
        assert "existing_tool" in tool_names
        assert "astrbot_execute_python" not in tool_names
        assert "read_file" not in tool_names
        assert "create_document" not in tool_names
        assert "add_blocks" not in tool_names
        assert "generate_complex_word_document" not in tool_names
        assert "当前聊天不可使用文件/Office/PDF 相关功能" in req.system_prompt
        assert "`astrbot_execute_python`" in req.system_prompt


@pytest.mark.asyncio
async def test_before_llm_chat_requires_read_before_document_tools_for_uploaded_files():
    async with _managed_plugin() as managed:
        source_path = Path(__file__).resolve()
        event = _build_event(
            message_type=MessageType.FRIEND_MESSAGE,
            sender_id="user-1",
        )
        event.message_obj.message = [
            Comp.File(name="source.docx", file=str(source_path)),
        ]
        _configure_uploaded_file(
            managed.plugin,
            source_path=source_path,
            original_name="source.docx",
            stored_name="source_1.docx",
        )
        req = _build_provider_request(
            "根据上传文档整理成正式汇报", tool_names=["existing_tool"]
        )

        await managed.plugin.before_llm_chat(event, req)

        assert "read_file" in req.prompt
        assert "source.docx" in req.prompt
        assert "source_1.docx" in req.prompt
        assert "读取前不要创建新文档" in req.prompt
        assert "source_1.docx" not in req.system_prompt


@pytest.mark.asyncio
@pytest.mark.parametrize(
    ("prompt", "expected_tool"),
    [
        (
            "调用 create_office_file，filename=report，content=hello，file_type=word",
            "create_office_file",
        ),
        (
            "reviewpr请求create_office_file，filename=report，content=hello，file_type=word",
            "create_office_file",
        ),
        (
            "please call `create_office_file` with filename=report, content=hello, file_type=word",
            "create_office_file",
        ),
    ],
)
async def test_before_llm_chat_restricts_file_tools_for_explicit_tool_call(
    prompt: str,
    expected_tool: str,
):
    async with _managed_plugin() as managed:
        event = _build_event(
            message_type=MessageType.FRIEND_MESSAGE,
            sender_id="user-1",
        )
        req = _build_provider_request(
            prompt,
            tool_names=[
                "existing_tool",
                "create_office_file",
                "create_document",
                "add_blocks",
                "finalize_document",
                "export_document",
                "read_file",
            ],
        )

        await managed.plugin.before_llm_chat(event, req)

        tool_names = set(req.func_tool.names())
        assert "existing_tool" in tool_names
        assert expected_tool in tool_names
        assert "create_document" not in tool_names
        assert "add_blocks" not in tool_names
        assert "finalize_document" not in tool_names
        assert "export_document" not in tool_names
        assert "read_file" not in tool_names


@pytest.mark.asyncio
async def test_before_llm_chat_does_not_restrict_when_prompt_mentions_multiple_tools():
    async with _managed_plugin() as managed:
        event = _build_event(
            message_type=MessageType.FRIEND_MESSAGE,
            sender_id="user-1",
        )
        req = _build_provider_request(
            "先调用 read_file 再调用 create_document 处理文件",
            tool_names=[
                "existing_tool",
                "read_file",
                "create_document",
                "add_blocks",
                "export_document",
            ],
        )

        await managed.plugin.before_llm_chat(event, req)

        tool_names = set(req.func_tool.names())
        assert "existing_tool" in tool_names
        assert "read_file" in tool_names
        assert "create_document" in tool_names
        assert "add_blocks" in tool_names
        assert "export_document" in tool_names


@pytest.mark.asyncio
async def test_before_llm_chat_does_not_restrict_for_question_style_tool_mention():
    async with _managed_plugin() as managed:
        event = _build_event(
            message_type=MessageType.FRIEND_MESSAGE,
            sender_id="user-1",
        )
        req = _build_provider_request(
            "请问 create_office_file 怎么用？先告诉我可用工具。",
            tool_names=[
                "existing_tool",
                "create_office_file",
                "create_document",
                "read_file",
            ],
        )

        await managed.plugin.before_llm_chat(event, req)

        tool_names = set(req.func_tool.names())
        assert "existing_tool" in tool_names
        assert "create_office_file" in tool_names
        assert "create_document" in tool_names
        assert "read_file" in tool_names


@pytest.mark.asyncio
@pytest.mark.parametrize(
    "prompt",
    [
        "不要调用 create_office_file，先告诉我可用工具。",
        "请问 create_office_file 怎么用？先告诉我可用工具。",
    ],
)
async def test_before_llm_chat_does_not_restrict_for_non_explicit_tool_mentions(
    prompt: str,
):
    async with _managed_plugin() as managed:
        event = _build_event(
            message_type=MessageType.FRIEND_MESSAGE,
            sender_id="user-1",
        )
        req = _build_provider_request(
            prompt,
            tool_names=[
                "existing_tool",
                "create_office_file",
                "create_document",
                "read_file",
            ],
        )

        await managed.plugin.before_llm_chat(event, req)

        tool_names = set(req.func_tool.names())
        assert "existing_tool" in tool_names
        assert "create_office_file" in tool_names
        assert "create_document" in tool_names
        assert "read_file" in tool_names


@pytest.mark.asyncio
async def test_before_llm_chat_does_not_treat_system_notice_as_explicit_tool_call():
    async with _managed_plugin() as managed:
        event = _build_event(
            message_type=MessageType.FRIEND_MESSAGE,
            sender_id="user-1",
        )
        event.message_str = "[System Notice] 用户上传了文件，请先调用 `read_file` 读取内容，再继续处理。"
        req = _build_provider_request(
            "",
            tool_names=[
                "existing_tool",
                "read_file",
                "create_document",
                "add_blocks",
                "export_document",
            ],
        )

        await managed.plugin.before_llm_chat(event, req)

        tool_names = set(req.func_tool.names())
        assert "existing_tool" in tool_names
        assert "read_file" in tool_names
        assert "create_document" in tool_names
        assert "add_blocks" in tool_names
        assert "export_document" in tool_names


@pytest.mark.asyncio
async def test_before_llm_chat_falls_back_to_raw_prompt_when_system_notice_block_is_missing():
    async with _managed_plugin() as managed:
        event = _build_event(
            message_type=MessageType.FRIEND_MESSAGE,
            sender_id="user-1",
        )
        req = _build_provider_request(
            "[System Notice] 这是用户自己输入的字面量。调用 read_file，filename=report.txt",
            tool_names=[
                "existing_tool",
                "read_file",
                "create_document",
                "add_blocks",
            ],
        )

        await managed.plugin.before_llm_chat(event, req)

        tool_names = set(req.func_tool.names())
        assert "existing_tool" in tool_names
        assert "read_file" in tool_names
        assert "create_document" not in tool_names
        assert "add_blocks" not in tool_names


@pytest.mark.asyncio
async def test_before_llm_chat_uses_buffered_user_instruction_for_explicit_tool_detection():
    async with _managed_plugin() as managed:
        event = _build_event(
            message_type=MessageType.FRIEND_MESSAGE,
            sender_id="user-1",
        )
        event._buffered = True
        req = _build_provider_request(
            (
                "[System Notice] 用户上传了 1 个文件\n\n"
                "[文件信息]\n"
                "- 原始文件名: source.docx\n"
                "  工作区文件名: source_1.docx\n\n"
                "[用户指令]\n"
                "请根据我刚上传的文档整理成正式汇报，标题叫《项目进展汇总》，最后导出成 Word 并发给我。\n\n"
                "[处理要求]\n"
                "1. 优先围绕这些上传文件完成用户请求。\n"
            ),
            tool_names=[
                "existing_tool",
                "read_file",
                "create_document",
                "add_blocks",
                "finalize_document",
                "export_document",
            ],
        )

        await managed.plugin.before_llm_chat(event, req)

        tool_names = set(req.func_tool.names())
        assert "existing_tool" in tool_names
        assert "read_file" in tool_names
        assert "create_document" in tool_names
        assert "add_blocks" in tool_names
        assert "finalize_document" in tool_names
        assert "export_document" in tool_names


@pytest.mark.asyncio
async def test_execute_excel_script_returns_direct_message_after_retry_exhaustion():
    async with _managed_plugin() as managed:
        event = _build_event()
        managed.plugin._runtime.file_tool_service.execute_excel_script = AsyncMock(
            return_value=json.dumps(
                {
                    "success": False,
                    "error": "错误：生成的 Excel 文件存在质量警告",
                    "traceback": "",
                    "script": "bad script",
                    "retry_count": 3,
                    "max_retries": 3,
                    "retry_exhausted": True,
                    "user_message": "Excel 脚本已经达到最多 3 次重试，本次没有生成合格文件。",
                },
                ensure_ascii=False,
            )
        )

        result = await managed.plugin.execute_excel_script(
            event,
            script="bad script",
            output_name="out.xlsx",
        )

    assert isinstance(result, MessageEventResult)
    assert result.is_stopped()
    assert "最多 3 次重试" in result.get_plain_text()


@pytest.mark.asyncio
async def test_before_llm_chat_can_still_restrict_tool_from_buffered_user_instruction():
    async with _managed_plugin() as managed:
        event = _build_event(
            message_type=MessageType.FRIEND_MESSAGE,
            sender_id="user-1",
        )
        event._buffered = True
        req = _build_provider_request(
            (
                "[System Notice] 用户上传了 1 个文件\n\n"
                "[文件信息]\n"
                "- 原始文件名: table.csv\n"
                "  工作区文件名: table_1.csv\n\n"
                "[用户指令]\n"
                "调用 read_file，filename=table_1.csv\n\n"
                "[处理要求]\n"
                "1. 优先围绕这些上传文件完成用户请求。\n"
            ),
            tool_names=[
                "existing_tool",
                "read_file",
                "create_document",
                "add_blocks",
            ],
        )

        await managed.plugin.before_llm_chat(event, req)

        tool_names = set(req.func_tool.names())
        assert "existing_tool" in tool_names
        assert "read_file" in tool_names
        assert "create_document" not in tool_names
        assert "add_blocks" not in tool_names


@pytest.mark.asyncio
async def test_llm_request_policy_runs_internal_notice_and_tool_hooks():
    async def _custom_notice_hook(context: NoticeBuildContext):
        context.notices.append("\n[custom notice]")
        return context

    async def _custom_tool_hook(context: ToolExposureContext):
        if context.request.func_tool:
            context.request.func_tool.remove_tool("create_document")
        return context

    policy = LLMRequestPolicy(
        document_toolset=SimpleNamespace(tools=[_tool("create_document")]),
        require_at_in_group=True,
        is_group_feature_enabled=lambda _event: True,
        check_permission=lambda _event: True,
        is_bot_mentioned=lambda _event: True,
        notice_hooks=[_custom_notice_hook],
        tool_exposure_hooks=[_custom_tool_hook],
    )
    event = _build_event(message_type=MessageType.FRIEND_MESSAGE, sender_id="user-1")
    req = _build_provider_request(
        "hello",
        tool_names=["create_document", "existing_tool"],
    )

    await policy.apply(event, req)

    assert "[custom notice]" in req.prompt
    assert "create_document" not in set(req.func_tool.names())
    assert "existing_tool" in set(req.func_tool.names())


@pytest.mark.asyncio
async def test_llm_request_policy_does_not_build_request_hook_service_when_hooks_are_provided():
    with patch(
        "astrbot_plugin_office_assistant.services.llm_request_policy.RequestHookService"
    ) as request_hook_service_cls:
        policy = LLMRequestPolicy(
            document_toolset=SimpleNamespace(tools=[_tool("create_document")]),
            require_at_in_group=True,
            is_group_feature_enabled=lambda _event: True,
            check_permission=lambda _event: True,
            is_bot_mentioned=lambda _event: True,
            notice_hooks=[],
            tool_exposure_hooks=[],
        )

    assert policy is not None
    request_hook_service_cls.assert_not_called()


@pytest.mark.asyncio
async def test_llm_request_policy_uses_injected_request_hook_service_for_default_hooks():
    request_hook_service = _build_request_hook_service()
    policy = LLMRequestPolicy(
        document_toolset=SimpleNamespace(tools=[_tool("create_document")]),
        require_at_in_group=True,
        is_group_feature_enabled=lambda _event: True,
        check_permission=lambda _event: True,
        is_bot_mentioned=lambda _event: True,
        request_hook_service=request_hook_service,
    )
    event = _build_event(message_type=MessageType.FRIEND_MESSAGE, sender_id="user-1")
    req = _build_provider_request(
        "请生成一份 Word 报告，并导出给我。",
        tool_names=[
            "create_document",
            "existing_tool",
            "astrbot_execute_shell",
        ],
    )

    with patch(
        "astrbot_plugin_office_assistant.services.llm_request_policy.logger.debug"
    ) as logger_debug:
        await policy.apply(event, req)

    tool_names = set(req.func_tool.names())
    assert "create_document" in tool_names
    assert "existing_tool" in tool_names
    assert "astrbot_execute_shell" not in tool_names
    assert "文件工具使用指南" in req.prompt
    assert req.prompt.startswith("请生成一份 Word 报告，并导出给我。")
    assert "\n\n[System Notice] 文件工具使用指南" in req.prompt
    assert any(
        call.args
        and call.args[0] == "[文件管理] Prompt sections(%s): %s"
        and call.args[1] == "prompt_suffix"
        and SECTION_STATIC_DOCUMENT_TOOLS in str(call.args[2])
        for call in logger_debug.call_args_list
    )


@pytest.mark.asyncio
async def test_llm_request_policy_appends_notice_to_prompt_suffix_without_extra_prefix():
    def _notice_hook(context):
        context.section_names.append("scene_test")
        context.notices.append("[notice]")
        return context

    policy = LLMRequestPolicy(
        document_toolset=SimpleNamespace(tools=[]),
        require_at_in_group=True,
        is_group_feature_enabled=lambda _event: True,
        check_permission=lambda _event: True,
        is_bot_mentioned=lambda _event: True,
        notice_hooks=[_notice_hook],
        tool_exposure_hooks=[],
    )
    event = _build_event()
    req = _build_provider_request("hello", tool_names=["existing_tool"])

    await policy.apply(event, req)

    assert req.prompt == "hello\n\n[notice]"


@pytest.mark.asyncio
async def test_llm_request_policy_uses_notice_directly_when_prompt_is_empty():
    def _notice_hook(context):
        context.section_names.append("scene_test")
        context.notices.append("[notice]")
        return context

    policy = LLMRequestPolicy(
        document_toolset=SimpleNamespace(tools=[]),
        require_at_in_group=True,
        is_group_feature_enabled=lambda _event: True,
        check_permission=lambda _event: True,
        is_bot_mentioned=lambda _event: True,
        notice_hooks=[_notice_hook],
        tool_exposure_hooks=[],
    )
    event = _build_event()
    req = _build_provider_request("", tool_names=["existing_tool"])

    await policy.apply(event, req)

    assert req.prompt == "[notice]"


@pytest.mark.asyncio
async def test_llm_request_policy_checks_permission_once_per_request():
    permission_calls = 0

    def _check_permission(_event):
        nonlocal permission_calls
        permission_calls += 1
        return False

    policy = LLMRequestPolicy(
        document_toolset=SimpleNamespace(tools=[_tool("create_document")]),
        require_at_in_group=True,
        is_group_feature_enabled=lambda _event: True,
        check_permission=_check_permission,
        is_bot_mentioned=lambda _event: False,
        notice_hooks=[],
        tool_exposure_hooks=[],
    )
    event = _build_event(message_type=MessageType.FRIEND_MESSAGE, sender_id="user-2")
    req = _build_provider_request(
        "hello",
        tool_names=["create_document", "existing_tool"],
    )

    await policy.apply(event, req)

    assert permission_calls == 1
    assert "create_document" not in set(req.func_tool.names())
    assert "当前聊天不可使用文件/Office/PDF 相关功能" in req.system_prompt


def test_llm_request_policy_requires_hook_pairs():
    with pytest.raises(ValueError, match="must be provided together"):
        LLMRequestPolicy(
            document_toolset=SimpleNamespace(tools=[]),
            require_at_in_group=True,
            is_group_feature_enabled=lambda _event: True,
            check_permission=lambda _event: True,
            is_bot_mentioned=lambda _event: True,
            notice_hooks=[],
            tool_exposure_hooks=None,
        )


@pytest.mark.asyncio
async def test_llm_request_policy_returns_after_tools_denied_notice():
    request_hook_service = _build_request_hook_service(
        get_cached_upload_infos=lambda _event: [
            {
                "original_name": "report.docx",
                "file_suffix": ".docx",
                "stored_name": "report_1.docx",
                "source_path": "/tmp/report.docx",
                "is_supported": True,
            }
        ],
    )
    policy = LLMRequestPolicy(
        document_toolset=SimpleNamespace(tools=[_tool("create_document")]),
        require_at_in_group=True,
        is_group_feature_enabled=lambda _event: True,
        check_permission=lambda _event: False,
        is_bot_mentioned=lambda _event: False,
        request_hook_service=request_hook_service,
    )
    event = _build_event(message_type=MessageType.FRIEND_MESSAGE, sender_id="user-2")
    event.message_obj.message = [
        Comp.File(name="report.docx", file="/tmp/report.docx"),
    ]
    req = _build_provider_request(
        "根据上传文件整理一下",
        tool_names=[
            "read_file",
            "existing_tool",
            "astrbot_execute_shell",
        ],
    )

    await policy.apply(event, req)

    assert "当前聊天不可使用文件/Office/PDF 相关功能" in req.system_prompt
    assert "已收到上传文件" not in req.system_prompt
    assert "existing_tool" in set(req.func_tool.names())
    assert "read_file" not in set(req.func_tool.names())
    assert "astrbot_execute_shell" not in set(req.func_tool.names())


@pytest.mark.asyncio
async def test_llm_request_policy_group_switch_off_hides_execution_tools():
    request_hook_service = _build_request_hook_service()
    policy = LLMRequestPolicy(
        document_toolset=SimpleNamespace(tools=[_tool("create_document")]),
        require_at_in_group=True,
        is_group_feature_enabled=lambda _event: False,
        check_permission=lambda _event: True,
        is_bot_mentioned=lambda _event: False,
        request_hook_service=request_hook_service,
    )
    event = _build_event(message_type=MessageType.GROUP_MESSAGE, sender_id="user-2")
    req = _build_provider_request(
        "根据上传文件整理一下",
        tool_names=[
            "read_file",
            "existing_tool",
            "astrbot_execute_shell",
        ],
    )

    await policy.apply(event, req)

    assert "当前聊天不可使用文件/Office/PDF 相关功能" in req.system_prompt
    assert "read_file" not in set(req.func_tool.names())
    assert "astrbot_execute_shell" not in set(req.func_tool.names())
    assert "existing_tool" in set(req.func_tool.names())


@pytest.mark.asyncio
async def test_runtime_bundle_does_not_expose_recent_text_cache():
    async with _managed_plugin() as managed:
        assert hasattr(managed.plugin._runtime, "upload_session_service") is True
        assert hasattr(managed.plugin._runtime, "recent_text_by_session") is False


@pytest.mark.asyncio
async def test_plugin_terminate_allows_missing_runtime():
    plugin = FileOperationPlugin.__new__(FileOperationPlugin)
    await plugin.terminate()


@pytest.mark.asyncio
async def test_plugin_terminate_cleans_runtime_resources():
    plugin = FileOperationPlugin.__new__(FileOperationPlugin)
    plugin._runtime = SimpleNamespace(
        message_buffer=MagicMock(set_complete_callback=MagicMock()),
        office_gen=MagicMock(cleanup=MagicMock()),
        pdf_converter=MagicMock(cleanup=MagicMock()),
        executor=MagicMock(shutdown=MagicMock()),
        temp_dir=MagicMock(cleanup=MagicMock()),
    )
    runtime = plugin._runtime

    await plugin.terminate()

    runtime.message_buffer.set_complete_callback.assert_called_once_with(None)
    runtime.office_gen.cleanup.assert_called_once_with()
    runtime.pdf_converter.cleanup.assert_called_once_with()
    runtime.executor.shutdown.assert_called_once_with(wait=False)
    runtime.temp_dir.cleanup.assert_called_once_with()
    assert plugin._runtime is None


@pytest.mark.asyncio
async def test_plugin_terminate_tolerates_temp_dir_permission_error():
    plugin = FileOperationPlugin.__new__(FileOperationPlugin)
    plugin._runtime = SimpleNamespace(
        message_buffer=MagicMock(set_complete_callback=MagicMock()),
        office_gen=MagicMock(cleanup=MagicMock()),
        pdf_converter=MagicMock(cleanup=MagicMock()),
        executor=MagicMock(shutdown=MagicMock()),
        temp_dir=MagicMock(cleanup=MagicMock(side_effect=PermissionError("locked"))),
    )
    runtime = plugin._runtime

    await plugin.terminate()

    runtime.message_buffer.set_complete_callback.assert_called_once_with(None)
    runtime.office_gen.cleanup.assert_called_once_with()
    runtime.pdf_converter.cleanup.assert_called_once_with()
    runtime.executor.shutdown.assert_called_once_with(wait=False)
    runtime.temp_dir.cleanup.assert_called_once_with()
    assert plugin._runtime is None


@pytest.mark.asyncio
async def test_on_buffer_complete_ignores_released_runtime():
    plugin = FileOperationPlugin.__new__(FileOperationPlugin)
    plugin._runtime = None
    buf = MagicMock()

    with patch("astrbot_plugin_office_assistant.main.logger.warning") as warning:
        await plugin._on_buffer_complete(buf)

    warning.assert_called_once()


@pytest.mark.asyncio
async def test_handle_exported_document_tool_uses_bound_service_after_runtime_release():
    plugin = FileOperationPlugin.__new__(FileOperationPlugin)
    service = MagicMock()
    service.handle_exported_document_tool = AsyncMock(return_value="sent")
    plugin._post_export_hook_service = service
    plugin._runtime = None
    context = MagicMock()

    result = await plugin._handle_exported_document_tool(
        context,
        "/tmp/exported.docx",
    )

    service.handle_exported_document_tool.assert_awaited_once_with(
        context,
        "/tmp/exported.docx",
    )
    assert result == "sent"


@pytest.mark.asyncio
async def test_doc_list_command_stops_event_after_sending():
    config = _build_config()
    config["trigger_settings"]["enable_features_in_group"] = True
    async with _managed_plugin(config=config) as managed:
        event = _build_event(message_type=MessageType.GROUP_MESSAGE)
        event.set_result = MagicMock()

        await managed.plugin.doc_list(event)

        event.set_result.assert_called_once()
        result = event.set_result.call_args.args[0]
        assert result.get_plain_text() == "当前没有可处理的上传文件。"
        assert result.is_stopped() is True


@pytest.mark.asyncio
async def test_doc_use_command_stops_event_after_requeue():
    async with _managed_plugin() as managed:
        event = _build_event(message_type=MessageType.GROUP_MESSAGE)
        event.stop_event = MagicMock()
        managed.plugin._runtime.command_service.doc_use = AsyncMock(return_value=None)

        await managed.plugin.doc_use(event, "f1 根据这份文件整理")

        managed.plugin._runtime.command_service.doc_use.assert_awaited_once_with(
            event,
            "f1 根据这份文件整理",
        )
        event.stop_event.assert_called_once_with()


@pytest.mark.asyncio
async def test_doc_clear_command_sets_stopped_result():
    config = _build_config()
    config["trigger_settings"]["enable_features_in_group"] = True
    async with _managed_plugin(config=config) as managed:
        event = _build_event(message_type=MessageType.GROUP_MESSAGE)
        event.set_result = MagicMock()

        await managed.plugin.doc_clear(event, "")

        event.set_result.assert_called_once()
        result = event.set_result.call_args.args[0]
        assert result.get_plain_text() == "❌ 当前没有可处理的上传文件。"
        assert result.is_stopped() is True


@pytest.mark.asyncio
async def test_buffered_upload_without_prompt_requeues_in_friend_chat():
    context = MagicMock()
    event_queue = AsyncMock()
    context.get_event_queue.return_value = event_queue
    async with _managed_plugin(context=context) as managed:
        source_path = Path(__file__).resolve()
        event = _build_event()
        upload = Comp.File(name="report.docx", file="report.docx")
        buf = BufferedMessage(event=event, files=[upload], texts=[])
        _configure_uploaded_file(
            managed.plugin,
            source_path=source_path,
            original_name="report.docx",
            stored_name="report_1.docx",
        )

        await managed.plugin._on_buffer_complete(buf)

        queued_event = event_queue.put.await_args.args[0]
        prompt_text = queued_event.message_obj.message[0].text
        upload_infos = (
            managed.plugin._runtime.upload_session_service.list_session_upload_infos(
                event
            )
        )
        assert len(upload_infos) == 1
        assert upload_infos[0]["original_name"] == "report.docx"
        assert upload_infos[0]["stored_name"] == "report_1.docx"
        assert upload_infos[0]["file_id"] == "f1"
        assert "[用户指令]" not in prompt_text
        assert "用户意图尚不明确时，再用中文询问用户想要如何处理" in prompt_text


@pytest.mark.asyncio
async def test_buffered_upload_without_prompt_only_caches_upload_infos_in_group_chat():
    context = MagicMock()
    event_queue = AsyncMock()
    context.get_event_queue.return_value = event_queue
    async with _managed_plugin(context=context) as managed:
        source_path = Path(__file__).resolve()
        event = _build_event(message_type=MessageType.GROUP_MESSAGE)
        upload = Comp.File(name="report.docx", file="report.docx")
        buf = BufferedMessage(event=event, files=[upload], texts=[])
        _configure_uploaded_file(
            managed.plugin,
            source_path=source_path,
            original_name="report.docx",
            stored_name="report_1.docx",
        )

        await managed.plugin._on_buffer_complete(buf)

        event_queue.put.assert_not_awaited()
        upload_infos = (
            managed.plugin._runtime.upload_session_service.list_session_upload_infos(
                event
            )
        )
        assert len(upload_infos) == 1
        assert upload_infos[0]["original_name"] == "report.docx"
        assert upload_infos[0]["stored_name"] == "report_1.docx"
        assert upload_infos[0]["file_id"] == "f1"


@pytest.mark.asyncio
async def test_buffered_upload_with_prompt_uses_structured_notice_and_follow_through_guidance():
    context = MagicMock()
    event_queue = AsyncMock()
    context.get_event_queue.return_value = event_queue
    service = UploadSessionService(
        context=context,
        recent_text_ttl_seconds=30,
        upload_session_ttl_seconds=300,
        recent_text_max_entries=32,
        recent_text_cleanup_interval_seconds=10,
        upload_session_cleanup_interval_seconds=30,
        extract_upload_source=AsyncMock(
            return_value=(Path(__file__).resolve(), "report.docx")
        ),
        store_uploaded_file=MagicMock(return_value=Path("report_1.docx")),
        allow_external_input_files=True,
    )
    event = _build_event()
    upload = Comp.File(name="report.docx", file="report.docx")
    buf = BufferedMessage(
        event=event,
        files=[upload],
        texts=["请根据上传文档整理成正式汇报"],
    )

    await service.on_buffer_complete(buf)

    queued_event = event_queue.put.await_args.args[0]
    assert queued_event is not event
    assert isinstance(queued_event.message_obj.message[0], Comp.Plain)
    prompt_text = queued_event.message_obj.message[0].text
    assert "[文件信息]" in prompt_text
    assert "[用户指令]" in prompt_text
    assert "请根据上传文档整理成正式汇报" in prompt_text
    assert prompt_text.index("[用户指令]") < prompt_text.index("[文件信息]")
    assert "[处理要求]" in prompt_text
    assert "优先围绕这些上传文件完成用户请求" in prompt_text
    assert "工作区文件名: report_1.docx" in prompt_text
    assert "外部绝对路径:" not in prompt_text
    assert "先调用 `read_file` 读取文件" in prompt_text
    assert "不要猜文件名，不要列目录，不要调用 shell" in prompt_text
    assert "读取后如果用户已明确说明具体改动，再继续调用工具" in prompt_text
    assert not queued_event.message_str.startswith("/")
    assert queued_event.message_str.endswith(prompt_text.strip())
    event_queue.put.assert_awaited_once()


@pytest.mark.asyncio
async def test_before_llm_chat_exposes_file_tools_for_buffered_group_upload_when_mentioned():
    context = MagicMock()
    event_queue = AsyncMock()
    context.get_event_queue.return_value = event_queue
    config = _build_config()
    config["trigger_settings"]["enable_features_in_group"] = True
    async with _managed_plugin(context=context, config=config) as managed:
        source_path = Path(__file__).resolve()
        event = _build_event(
            message_type=MessageType.GROUP_MESSAGE,
            sender_id="user-1",
        )
        raw_message = SimpleNamespace(mentions=[SimpleNamespace(id="bot-1")])
        event.message_obj.raw_message = raw_message
        event.is_mentioned.side_effect = lambda: hasattr(
            event.message_obj.raw_message, "mentions"
        ) and any(
            str(mention.id) == str(event.message_obj.self_id)
            for mention in event.message_obj.raw_message.mentions
        )
        upload = Comp.File(name="source.docx", file=str(source_path))
        buf = BufferedMessage(
            event=event,
            files=[upload],
            texts=["请根据上传文档整理成正式汇报"],
        )
        _configure_uploaded_file(
            managed.plugin,
            source_path=source_path,
            original_name="source.docx",
            stored_name="source_1.docx",
        )

        await managed.plugin._on_buffer_complete(buf)
        queued_event = event_queue.put.await_args.args[0]

        req = _build_provider_request(
            queued_event.message_str,
            tool_names=["existing_tool"],
        )

        await managed.plugin.before_llm_chat(queued_event, req)

        tool_names = set(req.func_tool.names())
        assert queued_event.message_obj.raw_message is raw_message
        assert queued_event.is_mentioned() is True
        assert "create_document" in tool_names
        assert "add_blocks" in tool_names
        assert "export_document" in tool_names
        assert "当前聊天不可使用文件/Office/PDF 相关功能" not in req.system_prompt
        assert "source_1.docx" in req.prompt


@pytest.mark.asyncio
async def test_before_llm_chat_hides_file_tools_for_buffered_group_upload_when_not_mentioned():
    context = MagicMock()
    event_queue = AsyncMock()
    context.get_event_queue.return_value = event_queue
    config = _build_config()
    config["trigger_settings"]["enable_features_in_group"] = True
    async with _managed_plugin(context=context, config=config) as managed:
        source_path = Path(__file__).resolve()
        event = _build_event(
            message_type=MessageType.GROUP_MESSAGE,
            sender_id="user-1",
        )
        raw_message = SimpleNamespace(mentions=[])
        event.message_obj.raw_message = raw_message
        event.is_mentioned.side_effect = lambda: hasattr(
            event.message_obj.raw_message, "mentions"
        ) and any(
            str(mention.id) == str(event.message_obj.self_id)
            for mention in event.message_obj.raw_message.mentions
        )
        upload = Comp.File(name="source.docx", file=str(source_path))
        buf = BufferedMessage(
            event=event,
            files=[upload],
            texts=["请根据上传文档整理成正式汇报"],
        )
        _configure_uploaded_file(
            managed.plugin,
            source_path=source_path,
            original_name="source.docx",
            stored_name="source_1.docx",
        )

        await managed.plugin._on_buffer_complete(buf)

        req = _build_provider_request(
            event.message_str,
            tool_names=["existing_tool"],
        )

        await managed.plugin.before_llm_chat(event, req)

        tool_names = set(req.func_tool.names())
        assert event.message_obj.raw_message is raw_message
        assert event.is_mentioned() is False
        assert "create_document" not in tool_names
        assert "add_blocks" not in tool_names
        assert "export_document" not in tool_names
        assert "当前聊天不可使用文件/Office/PDF 相关功能" in req.system_prompt
        assert "工作区文件名：source_1.docx" not in req.system_prompt


@pytest.mark.asyncio
async def test_before_llm_chat_exposes_file_tools_for_group_doc_command_without_mention():
    config = _build_config()
    config["trigger_settings"]["enable_features_in_group"] = True
    async with _managed_plugin(config=config) as managed:
        event = _build_event(
            message_type=MessageType.GROUP_MESSAGE,
            sender_id="user-1",
        )
        event.is_mentioned.return_value = False
        event._buffered = True
        event.set_extra(DOC_COMMAND_TRIGGER_EVENT_KEY, True)
        event.message_str = (
            "[System Notice] 用户上传了 1 个文件\n\n"
            "[文件信息]\n"
            "- 原始文件名: B.xlsx\n"
            "  工作区文件名: B_1.xlsx\n\n"
            "[用户指令]\n"
            "根据这份文件整理成正式汇报\n\n"
            "[处理要求]\n"
            "1. 优先围绕这些上传文件完成用户请求。\n"
        )
        req = _build_provider_request(event.message_str, tool_names=["existing_tool"])

        await managed.plugin.before_llm_chat(event, req)

        tool_names = set(req.func_tool.names())
        assert "create_document" in tool_names
        assert "add_blocks" in tool_names
        assert "export_document" in tool_names
        assert "当前聊天不可使用文件/Office/PDF 相关功能" not in req.system_prompt


@pytest.mark.asyncio
async def test_before_llm_chat_does_not_inject_upload_notice_when_file_tools_hidden():
    config = _build_config()
    config["trigger_settings"]["enable_features_in_group"] = True
    async with _managed_plugin(config=config) as managed:
        source_path = Path(__file__).resolve()
        event = _build_event(
            message_type=MessageType.GROUP_MESSAGE,
            sender_id="user-1",
        )
        event.is_mentioned.return_value = False
        event.message_obj.message = [
            Comp.File(name="source.docx", file=str(source_path)),
        ]
        _configure_uploaded_file(
            managed.plugin,
            source_path=source_path,
            original_name="source.docx",
            stored_name="source_1.docx",
        )

        req = _build_provider_request(
            "根据上传文档整理成正式汇报",
            tool_names=["existing_tool"],
        )

        await managed.plugin.before_llm_chat(event, req)

        assert "当前聊天不可使用文件/Office/PDF 相关功能" in req.system_prompt
        assert "工作区文件名：source_1.docx" not in req.system_prompt
        assert "MUST 先调用 `read_file` 读取此文件" not in req.system_prompt


@pytest.mark.asyncio
async def test_before_llm_chat_skips_upload_notices_when_func_tool_missing():
    async with _managed_plugin() as managed:
        source_path = Path(__file__).resolve()
        event = _build_event(
            message_type=MessageType.FRIEND_MESSAGE,
            sender_id="user-1",
        )
        event.message_obj.message = [
            Comp.File(name="source.docx", file=str(source_path)),
        ]
        _configure_uploaded_file(
            managed.plugin,
            source_path=source_path,
            original_name="source.docx",
            stored_name="source_1.docx",
        )

        req = _build_provider_request(
            "根据上传文档整理成正式汇报",
            func_tool=None,
        )

        await managed.plugin.before_llm_chat(event, req)

        assert "文件工具使用指南" not in req.system_prompt
        assert "工作区文件名：source_1.docx" not in req.system_prompt
        assert "MUST 先调用 `read_file` 读取此文件" not in req.system_prompt
