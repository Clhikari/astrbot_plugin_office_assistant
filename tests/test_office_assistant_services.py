import base64
import shutil
import zipfile
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import AsyncMock, MagicMock, patch
from uuid import uuid4

import mcp
import pytest
import astrbot_plugin_office_assistant.utils as office_utils
from astrbot_plugin_office_assistant.constants import (
    EXPLICIT_FILE_TOOL_EVENT_KEY,
    OfficeType,
)
from astrbot_plugin_office_assistant.message_buffer import BufferedMessage
from astrbot_plugin_office_assistant.services import (
    AccessPolicyService,
    CommandService,
    DeliveryService,
    ErrorHookService,
    ExportHookService,
    FileToolService,
    GeneratedFileDeliveryService,
    IncomingMessageService,
    PostExportHookService,
    RequestHookService,
    UploadPromptService,
    UploadSessionService,
    WorkspaceService,
    WordReadService,
    build_plugin_runtime,
)
from astrbot_plugin_office_assistant.utils import (
    ExtractedWordContent,
    ExtractedWordItem,
    extract_word_text,
    format_extracted_word_content,
)

import astrbot.api.message_components as Comp
from astrbot.core.agent.tool import FunctionTool, ToolSet
from astrbot.core.platform.message_type import MessageType
from astrbot.core.provider.entities import ProviderRequest


def _build_event(
    *,
    sender_id: str = "user-1",
    message_type=MessageType.FRIEND_MESSAGE,
):
    event = MagicMock()
    event.message_obj = SimpleNamespace(type=message_type, message=[], self_id="bot-1")
    event.get_sender_id.return_value = sender_id
    event.get_platform_id.return_value = "platform-1"
    event.unified_msg_origin = "session-1"
    event.message_str = ""
    event.is_admin.return_value = False
    event._buffer_reentry_count = 0
    event._buffered = False
    return event


def _tool(name: str) -> FunctionTool:
    return FunctionTool(
        name=name,
        description=f"tool {name}",
        parameters={"type": "object", "properties": {}},
        handler=None,
    )


def _make_workspace(name: str) -> Path:
    workspace_base = Path(__file__).resolve().parent / ".tmp_services"
    workspace_base.mkdir(parents=True, exist_ok=True)
    workspace_dir = workspace_base / f"{name}-{uuid4().hex}"
    workspace_dir.mkdir(parents=True, exist_ok=True)
    return workspace_dir


def _write_png(path: Path) -> None:
    png_bytes = base64.b64decode(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aF9kAAAAASUVORK5CYII="
    )
    path.write_bytes(png_bytes)


def _import_docx():
    return pytest.importorskip("docx")


def _rewrite_docx_document_xml(path: Path, transform) -> None:
    with zipfile.ZipFile(path, "r") as source_zip:
        file_map = {
            info.filename: source_zip.read(info.filename)
            for info in source_zip.infolist()
        }

    document_xml = file_map["word/document.xml"].decode("utf-8")
    file_map["word/document.xml"] = transform(document_xml).encode("utf-8")

    with zipfile.ZipFile(path, "w") as target_zip:
        for filename, content in file_map.items():
            target_zip.writestr(filename, content)


def test_access_policy_service_handles_whitelist_and_group_flags():
    service = AccessPolicyService(
        whitelist_users=["user-1"],
        enable_features_in_group=False,
    )
    friend_event = _build_event(message_type=MessageType.FRIEND_MESSAGE)
    group_event = _build_event(message_type=MessageType.GROUP_MESSAGE)

    assert service.check_permission(friend_event) is True
    assert service.is_group_feature_enabled(friend_event) is True
    assert service.is_group_feature_enabled(group_event) is False


def test_access_policy_service_detects_bot_mention():
    service = AccessPolicyService(
        whitelist_users=["user-1"],
        enable_features_in_group=True,
    )
    event = _build_event()
    event.message_obj.message = [Comp.At(qq="bot-1")]

    assert service.is_bot_mentioned(event) is True


def test_build_plugin_runtime_returns_temp_workspace_and_services():
    context = MagicMock()
    config = {
        "file_settings": {
            "auto_delete_files": True,
            "max_file_size_mb": 8,
            "enable_docx_image_review": False,
            "max_inline_docx_image_mb": 3,
            "max_inline_docx_image_count": 4,
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
    runtime = build_plugin_runtime(
        context=context,
        config=config,
        plugin_name="astrbot_plugin_office_assistant",
        handle_exported_document_tool=AsyncMock(),
        extract_upload_source=AsyncMock(),
        store_uploaded_file=MagicMock(),
    )

    try:
        assert runtime.settings.auto_delete is True
        assert runtime.plugin_data_path.exists()
        assert runtime.temp_dir is not None
        assert runtime.settings.enable_docx_image_review is False
        assert runtime.settings.max_inline_docx_image_bytes == 3 * 1024 * 1024
        assert runtime.settings.max_inline_docx_image_count == 4
        assert runtime.workspace_service.plugin_data_path == runtime.plugin_data_path
        assert runtime.post_export_hook_service is not None
        assert runtime.message_buffer is not None
        assert runtime.incoming_message_service is not None
    finally:
        runtime.executor.shutdown(wait=False)
        runtime.office_gen.cleanup()
        runtime.pdf_converter.cleanup()
        if runtime.temp_dir is not None:
            try:
                runtime.temp_dir.cleanup()
            except PermissionError:
                pass


def test_build_plugin_runtime_uses_persistent_workspace_when_auto_delete_disabled(
    monkeypatch: pytest.MonkeyPatch,
):
    data_root = _make_workspace("runtime-builder-data-root")
    called: dict[str, str | None] = {}
    context = MagicMock()
    config = {
        "file_settings": {
            "auto_delete_files": False,
            "max_file_size_mb": 16,
            "max_inline_docx_image_mb": 5,
            "max_inline_docx_image_count": 6,
            "message_buffer_seconds": 7,
        },
        "trigger_settings": {
            "reply_to_user": False,
            "require_at_in_group": False,
            "enable_features_in_group": True,
            "auto_block_execution_tools": False,
        },
        "preview_settings": {
            "enable": True,
            "dpi": 180,
        },
        "path_settings": {
            "allow_external_input_files": True,
        },
        "permission_settings": {
            "whitelist_users": ["user-2"],
        },
        "feature_settings": {
            "enable_office_files": True,
        },
    }

    def fake_get_data_dir(plugin_name=None):
        called["plugin_name"] = plugin_name
        return data_root

    monkeypatch.setattr(
        "astrbot_plugin_office_assistant.services.runtime_builder.StarTools.get_data_dir",
        fake_get_data_dir,
    )

    runtime = build_plugin_runtime(
        context=context,
        config=config,
        plugin_name="astrbot_plugin_office_assistant",
        handle_exported_document_tool=AsyncMock(),
        extract_upload_source=AsyncMock(),
        store_uploaded_file=MagicMock(),
    )

    try:
        assert runtime.temp_dir is None
        assert runtime.plugin_data_path == data_root / "files"
        assert runtime.plugin_data_path.exists()
        assert runtime.settings.auto_delete is False
        assert runtime.settings.max_file_size == 16 * 1024 * 1024
        assert runtime.settings.max_inline_docx_image_bytes == 5 * 1024 * 1024
        assert runtime.settings.max_inline_docx_image_count == 6
        assert runtime.settings.reply_to_user is False
        assert runtime.settings.require_at_in_group is False
        assert runtime.settings.enable_features_in_group is True
        assert runtime.settings.auto_block_execution_tools is False
        assert runtime.settings.enable_preview is True
        assert runtime.settings.preview_dpi == 180
        assert runtime.settings.allow_external_input_files is True
        assert runtime.settings.recent_text_ttl_seconds == 20
        assert runtime.settings.recent_text_cleanup_interval_seconds == 20
        assert runtime.command_service._plugin_data_path == data_root / "files"
        assert runtime.workspace_service.plugin_data_path == data_root / "files"
        assert runtime.post_export_hook_service is not None
        assert called["plugin_name"] == "astrbot_plugin_office_assistant"
    finally:
        runtime.executor.shutdown(wait=False)
        runtime.office_gen.cleanup()
        runtime.pdf_converter.cleanup()
        shutil.rmtree(data_root, ignore_errors=True)


@pytest.mark.asyncio
async def test_post_export_hook_service_handles_exported_document_tool():
    event = _build_event()
    event.send = AsyncMock()
    context = SimpleNamespace(context=SimpleNamespace(event=event))
    service = PostExportHookService(
        executor=ThreadPoolExecutor(max_workers=1),
        preview_generator=MagicMock(),
        enable_preview=False,
        auto_delete=False,
        reply_to_user=False,
        exported_message="✅ 文档已导出",
    )
    file_path = Path(__file__).resolve()

    try:
        result = await service.handle_exported_document_tool(
            context,
            str(file_path),
        )
    finally:
        service._executor.shutdown(wait=False)

    assert result == f"Document exported and sent to the user: {file_path.name}"
    assert event.send.await_count == 2


@pytest.mark.asyncio
async def test_export_hook_service_aliases_post_export_hook_service():
    assert ExportHookService is PostExportHookService


@pytest.mark.asyncio
async def test_request_hook_service_builds_default_tool_hooks():
    request = ProviderRequest(
        prompt="请用 read_file 看一下 report.docx",
        system_prompt="base",
        func_tool=ToolSet(
            [
                _tool("read_file"),
                _tool("create_document"),
                _tool("astrbot_execute_shell"),
                _tool("astrbot_execute_python"),
            ]
        ),
    )
    service = RequestHookService(
        auto_block_execution_tools=True,
        get_cached_upload_infos=lambda _event: [],
        extract_upload_source=AsyncMock(),
        store_uploaded_file=MagicMock(),
        allow_external_input_files=False,
    )
    context = SimpleNamespace(
        event=_build_event(),
        request=request,
        should_expose=True,
        can_process_upload=True,
        explicit_tool_name="read_file",
    )

    for hook in service.build_tool_exposure_hooks():
        context = await hook(context)

    tool_names = set(request.func_tool.names())
    assert "read_file" in tool_names
    assert "create_document" not in tool_names
    assert "astrbot_execute_shell" not in tool_names
    assert "astrbot_execute_python" not in tool_names


def test_upload_prompt_service_builds_instructional_notice_for_readable_files():
    service = UploadPromptService(allow_external_input_files=True)

    prompt_text = service.build_prompt(
        upload_infos=[
            {
                "original_name": "report.docx",
                "file_suffix": ".docx",
                "stored_name": "report_1.docx",
                "source_path": "/AstrBot/data/temp/report.docx",
                "is_supported": True,
            }
        ],
        user_instruction="看看里面的内容",
    )

    assert "[用户指令]" in prompt_text
    assert "看看里面的内容" in prompt_text
    assert "工作区文件名: report_1.docx" in prompt_text
    assert "外部绝对路径: /AstrBot/data/temp/report.docx" in prompt_text
    assert "先调用 `read_file` 读取文件" in prompt_text


def test_upload_prompt_service_builds_generic_notice_for_unreadable_files():
    service = UploadPromptService(allow_external_input_files=False)

    prompt_text = service.build_prompt(
        upload_infos=[
            {
                "original_name": "archive.bin",
                "file_suffix": ".bin",
                "stored_name": "",
                "source_path": "",
                "is_supported": False,
            }
        ],
        user_instruction="",
    )

    assert "[操作要求]" in prompt_text
    assert "使用中文与用户沟通" in prompt_text
    assert "[用户指令]" not in prompt_text
    assert "工作区文件名" not in prompt_text


@pytest.mark.asyncio
async def test_word_read_service_returns_disabled_image_notice_for_image_only_docx():
    workspace_dir = _make_workspace("word-read-service-image-only")
    executor = ThreadPoolExecutor(max_workers=1)
    docx_path = workspace_dir / "image-only.docx"
    image_path = workspace_dir / "embedded.png"

    try:
        docx = _import_docx()

        _write_png(image_path)
        document = docx.Document()
        document.add_picture(str(image_path), width=docx.shared.Inches(1))
        document.save(docx_path)

        workspace_service = WorkspaceService(
            plugin_data_path=workspace_dir,
            executor=executor,
            office_libs={"docx": object()},
            max_file_size=1024 * 1024,
            feature_settings={},
        )
        service = WordReadService(
            workspace_service=workspace_service,
            enable_docx_image_review=False,
        )

        results = [
            result
            async for result in service.iter_word_results(
                docx_path,
                docx_path.name,
                docx_path.suffix.lower(),
                docx_path.stat().st_size,
            )
        ]
    finally:
        executor.shutdown(wait=False)
        shutil.rmtree(workspace_dir, ignore_errors=True)

    assert len(results) == 1
    assert isinstance(results[0], str)
    assert "该 Word 文档仅包含图片内容，当前未启用图片理解。" in results[0]


def test_workspace_service_pre_check_rejects_outside_workspace():
    workspace_dir = Path(__file__).resolve().parent / "workspace-root"
    outside_file = Path(__file__).resolve()

    executor = ThreadPoolExecutor(max_workers=1)
    try:
        service = WorkspaceService(
            plugin_data_path=workspace_dir,
            executor=executor,
            office_libs={},
            max_file_size=1024,
            feature_settings={},
        )

        ok, resolved_path, err = service.pre_check(
            _build_event(),
            str(outside_file),
            require_exists=True,
            is_group_feature_enabled=lambda _event: True,
            check_permission_fn=lambda _event: True,
            group_feature_disabled_error=lambda: "group disabled",
        )

        assert not ok
        assert resolved_path is None
        assert err == "错误：非法路径：禁止访问工作区外的文件"
    finally:
        executor.shutdown(wait=False)


def test_upload_session_service_skips_system_notice_messages():
    service = UploadSessionService(
        context=MagicMock(),
        recent_text_ttl_seconds=30,
        recent_text_max_entries=32,
        recent_text_cleanup_interval_seconds=10,
        extract_upload_source=AsyncMock(),
        store_uploaded_file=MagicMock(),
        allow_external_input_files=False,
    )
    event = _build_event()
    session_key = service.get_attachment_session_key(event)

    event.message_str = "[System Notice] internal guidance"
    service.remember_recent_text(event)
    assert session_key not in service.recent_text_by_session

    event.message_str = "整理成正式汇报"
    service.remember_recent_text(event)
    assert service.recent_text_by_session[session_key][0] == "整理成正式汇报"


@pytest.mark.asyncio
async def test_upload_session_service_builds_read_first_prompt_for_buffered_upload():
    context = MagicMock()
    event_queue = AsyncMock()
    context.get_event_queue.return_value = event_queue
    source_path = Path(__file__).resolve()
    service = UploadSessionService(
        context=context,
        recent_text_ttl_seconds=30,
        recent_text_max_entries=32,
        recent_text_cleanup_interval_seconds=10,
        extract_upload_source=AsyncMock(return_value=(source_path, "report.docx")),
        store_uploaded_file=MagicMock(return_value=Path("report_1.docx")),
        allow_external_input_files=True,
    )
    event = _build_event()
    upload = Comp.File(name="report.docx", file="report.docx")
    buf = BufferedMessage(event=event, files=[upload], texts=[])

    await service.on_buffer_complete(buf)

    assert isinstance(event.message_obj.message[0], Comp.Plain)
    prompt_text = event.message_obj.message[0].text
    assert "用户上传了可读取文件，后续应优先围绕这些文件处理。" in prompt_text
    assert "工作区文件名: report_1.docx" in prompt_text
    assert "外部绝对路径:" in prompt_text
    assert "不要自行猜测文件名，也不要列目录或调用 shell。" in prompt_text
    assert "若使用相对路径，请使用上面的工作区文件名" in prompt_text
    assert "NEVER 创建新文档" not in prompt_text
    assert event.message_str == prompt_text.strip()
    event_queue.put.assert_awaited_once_with(event)


@pytest.mark.asyncio
async def test_upload_session_service_restores_recent_text_before_resolving_upload_infos():
    context = MagicMock()
    event_queue = AsyncMock()
    context.get_event_queue.return_value = event_queue
    service = UploadSessionService(
        context=context,
        recent_text_ttl_seconds=30,
        recent_text_max_entries=32,
        recent_text_cleanup_interval_seconds=10,
        extract_upload_source=AsyncMock(),
        store_uploaded_file=MagicMock(),
        allow_external_input_files=False,
    )
    event = _build_event()
    event.message_str = "看看里面的内容"
    service.remember_recent_text(event)
    session_key = service.get_attachment_session_key(event)
    upload = Comp.File(name="report.docx", file="report.docx")
    buf = BufferedMessage(event=event, files=[upload], texts=[])

    async def fake_ensure_upload_infos(_event, _files):
        assert session_key not in service.recent_text_by_session
        return [
            {
                "original_name": "report.docx",
                "file_suffix": ".docx",
                "type_desc": "Office文档 (Word/Excel/PPT)",
                "is_supported": True,
                "stored_name": "report_1.docx",
                "source_path": "",
            }
        ]

    service._ensure_upload_infos = AsyncMock(side_effect=fake_ensure_upload_infos)

    await service.on_buffer_complete(buf)

    prompt_text = event.message_obj.message[0].text
    assert "[用户指令]" in prompt_text
    assert "看看里面的内容" in prompt_text
    event_queue.put.assert_awaited_once_with(event)


@pytest.mark.asyncio
async def test_upload_session_service_omits_external_path_when_disabled():
    context = MagicMock()
    event_queue = AsyncMock()
    context.get_event_queue.return_value = event_queue
    source_path = Path(__file__).resolve()
    service = UploadSessionService(
        context=context,
        recent_text_ttl_seconds=30,
        recent_text_max_entries=32,
        recent_text_cleanup_interval_seconds=10,
        extract_upload_source=AsyncMock(return_value=(source_path, "report.docx")),
        store_uploaded_file=MagicMock(return_value=Path("report_1.docx")),
        allow_external_input_files=False,
    )
    event = _build_event()
    upload = Comp.File(name="report.docx", file="report.docx")
    buf = BufferedMessage(event=event, files=[upload], texts=[])

    await service.on_buffer_complete(buf)

    assert isinstance(event.message_obj.message[0], Comp.Plain)
    prompt_text = event.message_obj.message[0].text
    assert "工作区文件名: report_1.docx" in prompt_text
    assert "外部绝对路径:" not in prompt_text
    assert "如果需要使用工作区外路径" not in prompt_text
    assert "若使用相对路径，请使用上面的工作区文件名。" in prompt_text
    assert event.message_str == prompt_text.strip()
    event_queue.put.assert_awaited_once_with(event)


@pytest.mark.asyncio
async def test_upload_session_service_uses_extracted_filename_for_type_detection():
    context = MagicMock()
    source_path = Path(__file__).resolve()
    service = UploadSessionService(
        context=context,
        recent_text_ttl_seconds=30,
        recent_text_max_entries=32,
        recent_text_cleanup_interval_seconds=10,
        extract_upload_source=AsyncMock(return_value=(source_path, "report.docx")),
        store_uploaded_file=MagicMock(return_value=Path("report_1.docx")),
        allow_external_input_files=False,
    )
    event = _build_event()
    upload = Comp.File(name="upload-token", file="upload-token")

    infos = await service._ensure_upload_infos(event, [upload])

    assert len(infos) == 1
    info = infos[0]
    assert info["original_name"] == "report.docx"
    assert info["file_suffix"] == ".docx"
    assert info["type_desc"] == "Office文档 (Word/Excel/PPT)"
    assert info["is_supported"] is True
    assert info["stored_name"] == "report_1.docx"
    assert info["source_path"] == str(source_path.resolve())


@pytest.mark.asyncio
async def test_delivery_service_sends_message_and_file_for_existing_path():
    event = _build_event()
    event.send = AsyncMock()
    service = DeliveryService(
        executor=ThreadPoolExecutor(max_workers=1),
        preview_generator=MagicMock(),
        enable_preview=False,
        auto_delete=False,
        reply_to_user=True,
    )
    file_path = Path(__file__).resolve()
    try:
        await service.send_file_with_preview(event, file_path, "✅ 已发送")
    finally:
        service._executor.shutdown(wait=False)

    assert event.send.await_count == 2


@pytest.mark.asyncio
async def test_delivery_service_returns_missing_message_for_absent_export():
    event = _build_event()
    service = DeliveryService(
        executor=ThreadPoolExecutor(max_workers=1),
        preview_generator=MagicMock(),
        enable_preview=False,
        auto_delete=False,
        reply_to_user=False,
    )
    missing_path = Path(__file__).resolve().parent / "missing-output.docx"
    try:
        result = await service.send_exported_document(
            event,
            missing_path,
            "✅ exported",
        )
    finally:
        service._executor.shutdown(wait=False)

    assert "but the file does not exist" in result


@pytest.mark.asyncio
async def test_delivery_service_handles_exported_document_tool_via_context_event():
    event = _build_event()
    event.send = AsyncMock()
    context = SimpleNamespace(context=SimpleNamespace(event=event))
    service = DeliveryService(
        executor=ThreadPoolExecutor(max_workers=1),
        preview_generator=MagicMock(),
        enable_preview=False,
        auto_delete=False,
        reply_to_user=False,
    )
    file_path = Path(__file__).resolve()
    try:
        result = await service.handle_exported_document_tool(
            context,
            str(file_path),
            "✅ exported",
        )
    finally:
        service._executor.shutdown(wait=False)

    assert result == f"Document exported and sent to the user: {file_path.name}"
    assert event.send.await_count == 2


@pytest.mark.asyncio
async def test_generated_file_delivery_service_rejects_oversized_output():
    workspace_dir = _make_workspace("generated-file-delivery")
    event = _build_event()
    delivery_service = MagicMock()
    delivery_service.send_file_with_preview = AsyncMock()
    executor = ThreadPoolExecutor(max_workers=1)
    output_path = workspace_dir / "oversized.pdf"
    output_path.write_bytes(b"x" * 32)

    try:
        workspace_service = WorkspaceService(
            plugin_data_path=workspace_dir,
            executor=executor,
            office_libs={},
            max_file_size=8,
            feature_settings={},
        )
        service = GeneratedFileDeliveryService(
            workspace_service=workspace_service,
            delivery_service=delivery_service,
        )

        result = await service.deliver_generated_file(event, output_path)
    finally:
        executor.shutdown(wait=False)
        shutil.rmtree(workspace_dir, ignore_errors=True)

    assert result.status == "oversized"
    assert result.file_size == 32
    assert result.max_size == 8
    delivery_service.send_file_with_preview.assert_not_called()


@pytest.mark.asyncio
async def test_delivery_service_logs_preview_generation_failure():
    event = _build_event()
    event.send = AsyncMock()
    preview_generator = MagicMock()
    preview_generator.generate_preview.side_effect = RuntimeError("preview boom")
    service = DeliveryService(
        executor=ThreadPoolExecutor(max_workers=1),
        preview_generator=preview_generator,
        enable_preview=True,
        auto_delete=False,
        reply_to_user=False,
    )
    file_path = Path(__file__).resolve()
    try:
        with patch(
            "astrbot_plugin_office_assistant.services.delivery_service.logger.warning"
        ) as warning_mock:
            await service.send_file_with_preview(event, file_path, "✅ 已发送")
    finally:
        service._executor.shutdown(wait=False)

    warning_mock.assert_called_once()
    assert "生成预览图失败" in warning_mock.call_args.args[0]
    assert event.send.await_count == 2


@pytest.mark.asyncio
async def test_delivery_service_logs_file_cleanup_failure():
    event = _build_event()
    event.send = AsyncMock()
    workspace_dir = _make_workspace("delivery-cleanup")
    preview_path = workspace_dir / "preview.png"
    preview_path.write_text("preview", encoding="utf-8")
    file_path = workspace_dir / "result.docx"
    file_path.write_text("result", encoding="utf-8")

    preview_generator = MagicMock()
    preview_generator.generate_preview.return_value = preview_path

    service = DeliveryService(
        executor=ThreadPoolExecutor(max_workers=1),
        preview_generator=preview_generator,
        enable_preview=True,
        auto_delete=True,
        reply_to_user=False,
    )

    with patch.object(Path, "unlink", side_effect=PermissionError("locked")):
        with patch(
            "astrbot_plugin_office_assistant.services.delivery_service.logger.warning"
        ) as warning_mock:
            try:
                await service.send_file_with_preview(event, file_path, "✅ 已发送")
            finally:
                service._executor.shutdown(wait=False)
                shutil.rmtree(workspace_dir, ignore_errors=True)

    assert warning_mock.call_count == 2
    logged_messages = [call.args[0] for call in warning_mock.call_args_list]
    assert any("自动删除预览文件失败" in message for message in logged_messages)
    assert any("自动删除文件失败" in message for message in logged_messages)


@pytest.mark.asyncio
async def test_incoming_message_service_buffers_supported_file_and_stops_event():
    message_buffer = MagicMock()
    message_buffer.add_message = AsyncMock(return_value=True)
    message_buffer.is_buffering.return_value = False
    remember_recent_text = MagicMock()
    service = IncomingMessageService(
        message_buffer=message_buffer,
        remember_recent_text=remember_recent_text,
        is_group_feature_enabled=lambda _event: True,
    )
    event = _build_event()
    event.stop_event = MagicMock()
    event.message_obj.message = [Comp.File(name="report.docx", file="report.docx")]

    await service.handle_file_message(event)

    remember_recent_text.assert_called_once_with(event)
    message_buffer.add_message.assert_awaited_once_with(event)
    event.stop_event.assert_called_once()


@pytest.mark.asyncio
async def test_incoming_message_service_keeps_existing_buffer_for_unsupported_file():
    message_buffer = MagicMock()
    message_buffer.add_message = AsyncMock(return_value=True)
    message_buffer.is_buffering.return_value = True
    remember_recent_text = MagicMock()
    service = IncomingMessageService(
        message_buffer=message_buffer,
        remember_recent_text=remember_recent_text,
        is_group_feature_enabled=lambda _event: True,
    )
    event = _build_event()
    event.stop_event = MagicMock()
    event.message_obj.message = [Comp.File(name="avatar.png", file="avatar.png")]

    await service.handle_file_message(event)

    remember_recent_text.assert_called_once_with(event)
    message_buffer.is_buffering.assert_called_once_with(event)
    message_buffer.add_message.assert_awaited_once_with(event)
    event.stop_event.assert_called_once()


@pytest.mark.asyncio
async def test_error_hook_service_uses_event_session_when_target_not_configured():
    context = MagicMock()
    context.send_message = AsyncMock(return_value=True)
    service = ErrorHookService(
        context=context,
        config={},
        plugin_name="astrbot_plugin_office_assistant",
    )
    event = _build_event()
    event.stop_event = MagicMock()

    await service.handle_plugin_error(
        event,
        "astrbot_plugin_office_assistant",
        "handler_name",
        RuntimeError("boom"),
        "traceback line 1\ntraceback line 2",
    )

    context.send_message.assert_awaited_once()
    args, _kwargs = context.send_message.await_args
    assert args[0] == "session-1"
    assert "handler=handler_name" in args[1].chain[0].text
    event.stop_event.assert_called_once()


@pytest.mark.asyncio
async def test_file_tool_service_reads_text_from_workspace_file():
    workspace_dir = Path(__file__).resolve().parent
    executor = ThreadPoolExecutor(max_workers=1)
    event = _build_event()

    try:
        workspace_service = WorkspaceService(
            plugin_data_path=workspace_dir,
            executor=executor,
            office_libs={},
            max_file_size=1024 * 1024,
            feature_settings={},
        )
        service = FileToolService(
            workspace_service=workspace_service,
            office_generator=MagicMock(),
            pdf_converter=MagicMock(),
            delivery_service=MagicMock(),
            office_libs={},
            allow_external_input_files=False,
            is_group_feature_enabled=lambda _event: True,
            check_permission=lambda _event: True,
            group_feature_disabled_error=lambda: "group disabled",
        )

        result = await service.read_file(event, Path(__file__).resolve().name)
    finally:
        executor.shutdown(wait=False)

    assert result is not None
    assert "[文件:" in result
    assert "test_file_tool_service_reads_text_from_workspace_file" in result


@pytest.mark.asyncio
async def test_file_tool_service_reads_docx_and_extracts_embedded_images():
    workspace_dir = _make_workspace("read-docx-images")
    executor = ThreadPoolExecutor(max_workers=1)
    event = _build_event()
    docx_path = workspace_dir / "image-report.docx"
    image_path = workspace_dir / "embedded.png"

    try:
        docx = _import_docx()

        _write_png(image_path)
        document = docx.Document()
        document.add_paragraph("文档正文")
        document.add_picture(str(image_path), width=docx.shared.Inches(1))
        document.add_paragraph("图片后的说明")
        document.save(docx_path)

        workspace_service = WorkspaceService(
            plugin_data_path=workspace_dir,
            executor=executor,
            office_libs={"docx": object()},
            max_file_size=1024 * 1024,
            feature_settings={},
        )
        service = FileToolService(
            workspace_service=workspace_service,
            office_generator=MagicMock(),
            pdf_converter=MagicMock(),
            delivery_service=MagicMock(),
            office_libs={"docx": object()},
            allow_external_input_files=False,
            is_group_feature_enabled=lambda _event: True,
            check_permission=lambda _event: True,
            group_feature_disabled_error=lambda: "group disabled",
        )

        result = await service.read_file(event, docx_path.name)
    finally:
        executor.shutdown(wait=False)
        shutil.rmtree(workspace_dir, ignore_errors=True)

    assert result is not None
    assert "文档正文" in result
    assert "[插图1]" in result
    assert "图片后的说明" in result


def test_extract_word_text_ignores_deleted_and_field_code_runs():
    workspace_dir = _make_workspace("extract-docx-hidden-runs")
    docx_path = workspace_dir / "tracked.docx"

    try:
        docx = _import_docx()

        document = docx.Document()
        document.add_paragraph("保留文本")
        document.save(docx_path)

        _rewrite_docx_document_xml(
            docx_path,
            lambda xml: xml.replace(
                "<w:t>保留文本</w:t>",
                (
                    "<w:t>保留文本</w:t>"
                    "<w:delText>删除内容</w:delText>"
                    "<w:instrText> MERGEFIELD secret </w:instrText>"
                ),
            ),
        )

        extracted = extract_word_text(docx_path, workspace_dir)
    finally:
        shutil.rmtree(workspace_dir, ignore_errors=True)

    assert extracted is not None
    assert "保留文本" in extracted
    assert "删除内容" not in extracted
    assert "MERGEFIELD secret" not in extracted


def test_format_extracted_word_content_uses_basename_outside_workspace():
    workspace_dir = _make_workspace("format-word-content-paths")
    outside_image = Path(__file__).resolve()

    try:
        content = ExtractedWordContent(
            items=[
                ExtractedWordItem(type="text", text="正文"),
                ExtractedWordItem(
                    type="image",
                    image_path=outside_image,
                    image_index=1,
                ),
            ]
        )

        formatted = format_extracted_word_content(
            content,
            workspace_root=workspace_dir,
            include_image_paths=True,
        )
    finally:
        shutil.rmtree(workspace_dir, ignore_errors=True)

    assert formatted is not None
    assert "[插图1]" in formatted
    assert outside_image.name in formatted
    assert outside_image.as_posix() not in formatted


def test_extract_word_content_returns_structured_result_for_legacy_doc(
    monkeypatch: pytest.MonkeyPatch,
):
    workspace_dir = _make_workspace("extract-legacy-doc")
    doc_path = workspace_dir / "legacy.doc"
    doc_path.write_bytes(b"legacy")

    monkeypatch.setattr(office_utils, "_ANTIWORD_AVAILABLE", True)
    monkeypatch.setattr(office_utils, "_WIN32COM_AVAILABLE", False)
    monkeypatch.setattr(
        office_utils,
        "_extract_doc_text_antiword",
        lambda _path: "Legacy Word content",
    )

    try:
        extracted = office_utils.extract_word_content(doc_path, workspace_dir)
    finally:
        shutil.rmtree(workspace_dir, ignore_errors=True)

    assert extracted == ExtractedWordContent(
        text="Legacy Word content",
        image_paths=[],
        items=[],
    )


def test_workspace_service_extract_office_text_reads_legacy_doc(
    monkeypatch: pytest.MonkeyPatch,
):
    workspace_dir = _make_workspace("workspace-legacy-doc")
    executor = ThreadPoolExecutor(max_workers=1)
    doc_path = workspace_dir / "legacy.doc"
    doc_path.write_bytes(b"legacy")

    monkeypatch.setattr(office_utils, "_ANTIWORD_AVAILABLE", True)
    monkeypatch.setattr(office_utils, "_WIN32COM_AVAILABLE", False)
    monkeypatch.setattr(
        office_utils,
        "_extract_doc_text_antiword",
        lambda _path: "Legacy document text",
    )

    try:
        workspace_service = WorkspaceService(
            plugin_data_path=workspace_dir,
            executor=executor,
            office_libs={},
            max_file_size=1024 * 1024,
            feature_settings={},
        )
        extracted = workspace_service.extract_office_text(doc_path, OfficeType.WORD)
    finally:
        executor.shutdown(wait=False)
        shutil.rmtree(workspace_dir, ignore_errors=True)

    assert extracted == "Legacy document text"


def test_workspace_service_extract_word_content_skips_docx_image_materialization_when_disabled():
    workspace_dir = _make_workspace("workspace-docx-no-image-materialize")
    executor = ThreadPoolExecutor(max_workers=1)
    docx_path = workspace_dir / "image-report.docx"
    image_path = workspace_dir / "embedded.png"

    try:
        docx = _import_docx()

        _write_png(image_path)
        document = docx.Document()
        document.add_paragraph("图前说明")
        document.add_picture(str(image_path), width=docx.shared.Inches(1))
        document.add_paragraph("图后说明")
        document.save(docx_path)

        workspace_service = WorkspaceService(
            plugin_data_path=workspace_dir,
            executor=executor,
            office_libs={"docx": object()},
            max_file_size=1024 * 1024,
            feature_settings={},
        )

        extracted = workspace_service.extract_word_content(
            docx_path,
            include_images=False,
        )
    finally:
        executor.shutdown(wait=False)
        shutil.rmtree(workspace_dir, ignore_errors=True)

    assert extracted is not None
    assert extracted.image_count == 1
    assert extracted.image_paths == []
    assert all(item.type == "text" for item in extracted.items)


def test_workspace_service_extract_word_content_formats_relative_image_paths_when_enabled():
    workspace_dir = _make_workspace("workspace-docx-image-materialize")
    executor = ThreadPoolExecutor(max_workers=1)
    docx_path = workspace_dir / "image-report.docx"
    image_path = workspace_dir / "embedded.png"

    try:
        docx = _import_docx()

        _write_png(image_path)
        document = docx.Document()
        document.add_paragraph("图前说明")
        document.add_picture(str(image_path), width=docx.shared.Inches(1))
        document.add_paragraph("图后说明")
        document.save(docx_path)

        workspace_service = WorkspaceService(
            plugin_data_path=workspace_dir,
            executor=executor,
            office_libs={"docx": object()},
            max_file_size=1024 * 1024,
            feature_settings={},
        )

        extracted = workspace_service.extract_word_content(
            docx_path,
            include_images=True,
        )
        formatted = workspace_service.format_word_content(
            extracted,
            include_image_paths=True,
        )
    finally:
        executor.shutdown(wait=False)
        shutil.rmtree(workspace_dir, ignore_errors=True)

    assert extracted is not None
    assert extracted.image_count == 1
    assert len(extracted.image_paths) == 1
    assert formatted is not None
    assert "图前说明" in formatted
    assert "图后说明" in formatted
    assert "[插图1]" in formatted
    assert ".read_assets/" in formatted
    assert str(workspace_dir) not in formatted
    assert str(workspace_dir.resolve()) not in formatted


@pytest.mark.asyncio
async def test_file_tool_service_streams_docx_images_as_tool_results():
    workspace_dir = _make_workspace("stream-docx-images")
    executor = ThreadPoolExecutor(max_workers=1)
    event = _build_event()
    docx_path = workspace_dir / "image-report.docx"
    image_path = workspace_dir / "embedded.png"

    try:
        docx = _import_docx()

        _write_png(image_path)
        document = docx.Document()
        document.add_paragraph("文档正文")
        document.add_picture(str(image_path), width=docx.shared.Inches(1))
        document.add_paragraph("图片后的说明")
        document.save(docx_path)

        workspace_service = WorkspaceService(
            plugin_data_path=workspace_dir,
            executor=executor,
            office_libs={"docx": object()},
            max_file_size=1024 * 1024,
            feature_settings={},
        )
        service = FileToolService(
            workspace_service=workspace_service,
            office_generator=MagicMock(),
            pdf_converter=MagicMock(),
            delivery_service=MagicMock(),
            office_libs={"docx": object()},
            allow_external_input_files=False,
            is_group_feature_enabled=lambda _event: True,
            check_permission=lambda _event: True,
            group_feature_disabled_error=lambda: "group disabled",
        )

        results = [
            result
            async for result in service.iter_read_file_tool_results(
                event, docx_path.name
            )
        ]
    finally:
        executor.shutdown(wait=False)
        shutil.rmtree(workspace_dir, ignore_errors=True)

    assert isinstance(results[0], str)
    assert "文档正文" in results[0]
    assert "[插图1]" in results[0]
    assert len(results) == 2
    assert isinstance(results[1], mcp.types.CallToolResult)
    assert isinstance(results[1].content[0], mcp.types.ImageContent)
    assert results[1].content[0].mimeType == "image/png"
    assert results[1].content[0].data
    assert "图片后的说明" in results[0]


@pytest.mark.asyncio
async def test_file_tool_service_returns_error_when_docx_library_missing():
    workspace_dir = _make_workspace("missing-docx-lib")
    executor = ThreadPoolExecutor(max_workers=1)
    event = _build_event()
    docx_path = workspace_dir / "missing-lib.docx"
    docx_path.write_bytes(b"not-a-real-docx")

    try:
        workspace_service = WorkspaceService(
            plugin_data_path=workspace_dir,
            executor=executor,
            office_libs={},
            max_file_size=1024 * 1024,
            feature_settings={},
        )
        service = FileToolService(
            workspace_service=workspace_service,
            office_generator=MagicMock(),
            pdf_converter=MagicMock(),
            delivery_service=MagicMock(),
            office_libs={},
            allow_external_input_files=False,
            is_group_feature_enabled=lambda _event: True,
            check_permission=lambda _event: True,
            group_feature_disabled_error=lambda: "group disabled",
        )

        results = [
            result
            async for result in service.iter_read_file_tool_results(
                event, docx_path.name
            )
        ]
    finally:
        executor.shutdown(wait=False)
        shutil.rmtree(workspace_dir, ignore_errors=True)

    assert results == ["错误：文件 'missing-lib.docx' 无法读取，可能未安装对应解析库"]


@pytest.mark.asyncio
async def test_file_tool_service_skips_unreadable_docx_image_bytes(
    monkeypatch: pytest.MonkeyPatch,
):
    workspace_dir = _make_workspace("broken-docx-image")
    executor = ThreadPoolExecutor(max_workers=1)
    event = _build_event()
    docx_path = workspace_dir / "broken-image.docx"
    docx_path.write_bytes(b"placeholder")
    broken_image_path = workspace_dir / "broken.png"
    healthy_image_path = workspace_dir / "healthy.png"
    broken_image_path.write_bytes(b"broken")
    _write_png(healthy_image_path)

    extracted = ExtractedWordContent(
        items=[
            ExtractedWordItem(type="text", text="文档正文"),
            ExtractedWordItem(
                type="image",
                image_path=broken_image_path,
                image_index=1,
            ),
            ExtractedWordItem(type="text", text="收尾说明"),
            ExtractedWordItem(
                type="image",
                image_path=healthy_image_path,
                image_index=2,
            ),
        ],
        image_paths=[broken_image_path, healthy_image_path],
    )

    original_read_bytes = Path.read_bytes

    def fake_read_bytes(path: Path) -> bytes:
        if path == broken_image_path:
            raise OSError("broken image")
        return original_read_bytes(path)

    monkeypatch.setattr(Path, "read_bytes", fake_read_bytes)

    try:
        workspace_service = WorkspaceService(
            plugin_data_path=workspace_dir,
            executor=executor,
            office_libs={"docx": object()},
            max_file_size=1024 * 1024,
            feature_settings={},
        )
        monkeypatch.setattr(
            workspace_service,
            "extract_word_content",
            lambda _path, include_images=True: extracted,
        )
        service = FileToolService(
            workspace_service=workspace_service,
            office_generator=MagicMock(),
            pdf_converter=MagicMock(),
            delivery_service=MagicMock(),
            office_libs={"docx": object()},
            allow_external_input_files=False,
            is_group_feature_enabled=lambda _event: True,
            check_permission=lambda _event: True,
            group_feature_disabled_error=lambda: "group disabled",
        )

        results = [
            result
            async for result in service.iter_read_file_tool_results(
                event, docx_path.name
            )
        ]
    finally:
        executor.shutdown(wait=False)
        shutil.rmtree(workspace_dir, ignore_errors=True)

    assert len(results) == 2
    assert isinstance(results[0], str)
    assert "文档正文" in results[0]
    assert "[插图1]" in results[0]
    assert "[插图2]" in results[0]
    assert "收尾说明" in results[0]
    assert isinstance(results[1], mcp.types.CallToolResult)
    assert len(results[1].content) == 1
    assert isinstance(results[1].content[0], mcp.types.ImageContent)


@pytest.mark.asyncio
async def test_file_tool_service_uses_item_image_index_for_skip_reasons():
    workspace_dir = _make_workspace("stream-docx-image-index")
    executor = ThreadPoolExecutor(max_workers=1)
    event = _build_event()
    docx_path = workspace_dir / "image-report.docx"
    large_image = workspace_dir / "embedded-large.png"
    small_image = workspace_dir / "embedded-small.png"

    try:
        docx_path.write_bytes(b"docx")
        large_image.write_bytes(b"a" * 40)
        small_image.write_bytes(b"b" * 10)

        workspace_service = WorkspaceService(
            plugin_data_path=workspace_dir,
            executor=executor,
            office_libs={"docx": object()},
            max_file_size=1024 * 1024,
            feature_settings={},
        )
        workspace_service.extract_word_content = MagicMock(
            return_value=SimpleNamespace(
                text="文档正文",
                image_paths=[large_image, small_image],
                image_count=2,
                items=[
                    SimpleNamespace(type="text", text="文档正文"),
                    SimpleNamespace(
                        type="image",
                        image_path=small_image,
                        image_index=2,
                    ),
                    SimpleNamespace(
                        type="image",
                        image_path=large_image,
                        image_index=1,
                    ),
                    SimpleNamespace(type="text", text="收尾说明"),
                ],
            )
        )

        service = FileToolService(
            workspace_service=workspace_service,
            office_generator=MagicMock(),
            pdf_converter=MagicMock(),
            delivery_service=MagicMock(),
            office_libs={"docx": object()},
            allow_external_input_files=False,
            max_inline_docx_image_bytes=20,
            max_inline_docx_image_count=2,
            is_group_feature_enabled=lambda _event: True,
            check_permission=lambda _event: True,
            group_feature_disabled_error=lambda: "group disabled",
        )

        results = [
            result
            async for result in service.iter_read_file_tool_results(
                event, docx_path.name
            )
        ]
    finally:
        executor.shutdown(wait=False)
        shutil.rmtree(workspace_dir, ignore_errors=True)

    assert isinstance(results[0], str)
    assert "[插图2]" in results[0]
    assert "[插图1]（未注入模型上下文" in results[0]
    assert "超过 20.00 B 限制" in results[0]
    assert isinstance(results[1], mcp.types.CallToolResult)
    assert len(results[1].content) == 1


@pytest.mark.asyncio
async def test_file_tool_service_limits_inline_docx_images():
    workspace_dir = _make_workspace("stream-docx-image-limits")
    executor = ThreadPoolExecutor(max_workers=1)
    event = _build_event()
    docx_path = workspace_dir / "image-report.docx"
    small_image_1 = workspace_dir / "embedded-1.png"
    large_image = workspace_dir / "embedded-large.png"
    small_image_2 = workspace_dir / "embedded-2.png"
    small_image_3 = workspace_dir / "embedded-3.png"

    try:
        docx_path.write_bytes(b"docx")
        small_image_1.write_bytes(b"a" * 10)
        large_image.write_bytes(b"b" * 40)
        small_image_2.write_bytes(b"c" * 10)
        small_image_3.write_bytes(b"d" * 10)

        workspace_service = WorkspaceService(
            plugin_data_path=workspace_dir,
            executor=executor,
            office_libs={"docx": object()},
            max_file_size=1024 * 1024,
            feature_settings={},
        )
        workspace_service.extract_word_content = MagicMock(
            return_value=SimpleNamespace(
                text="文档正文",
                image_paths=[
                    small_image_1,
                    large_image,
                    small_image_2,
                    small_image_3,
                ],
                image_count=4,
                items=[
                    SimpleNamespace(type="text", text="文档正文"),
                    SimpleNamespace(type="image", image_path=small_image_1),
                    SimpleNamespace(type="image", image_path=large_image),
                    SimpleNamespace(type="image", image_path=small_image_2),
                    SimpleNamespace(type="image", image_path=small_image_3),
                    SimpleNamespace(type="text", text="收尾说明"),
                ],
            )
        )

        service = FileToolService(
            workspace_service=workspace_service,
            office_generator=MagicMock(),
            pdf_converter=MagicMock(),
            delivery_service=MagicMock(),
            office_libs={"docx": object()},
            allow_external_input_files=False,
            max_inline_docx_image_bytes=20,
            max_inline_docx_image_count=2,
            is_group_feature_enabled=lambda _event: True,
            check_permission=lambda _event: True,
            group_feature_disabled_error=lambda: "group disabled",
        )

        results = [
            result
            async for result in service.iter_read_file_tool_results(
                event, docx_path.name
            )
        ]
    finally:
        executor.shutdown(wait=False)
        shutil.rmtree(workspace_dir, ignore_errors=True)

    assert isinstance(results[0], str)
    assert "文档正文" in results[0]
    assert "[插图1]" in results[0]
    assert "插图2" in results[0]
    assert "超过 20.00 B 限制" in results[0]
    assert "插图4" in results[0]
    assert "超过单文档最多 2 张限制" in results[0]
    assert "收尾说明" in results[0]
    assert isinstance(results[1], mcp.types.CallToolResult)
    assert len(results[1].content) == 2
    assert all(isinstance(item, mcp.types.ImageContent) for item in results[1].content)
    assert len(results) == 2


@pytest.mark.asyncio
async def test_file_tool_service_skips_docx_image_review_when_disabled():
    workspace_dir = _make_workspace("stream-docx-image-review-disabled")
    executor = ThreadPoolExecutor(max_workers=1)
    event = _build_event()
    docx_path = workspace_dir / "image-report.docx"
    image_path = workspace_dir / "embedded.png"

    try:
        _write_png(image_path)
        docx_path.write_bytes(b"docx")

        workspace_service = WorkspaceService(
            plugin_data_path=workspace_dir,
            executor=executor,
            office_libs={"docx": object()},
            max_file_size=1024 * 1024,
            feature_settings={},
        )
        workspace_service.extract_word_content = MagicMock(
            return_value=SimpleNamespace(
                text="文档正文",
                image_paths=[image_path],
                image_count=1,
                items=[
                    SimpleNamespace(type="text", text="图前说明"),
                    SimpleNamespace(
                        type="image",
                        image_path=image_path,
                        image_index=1,
                    ),
                    SimpleNamespace(type="text", text="图后说明"),
                ],
            )
        )

        service = FileToolService(
            workspace_service=workspace_service,
            office_generator=MagicMock(),
            pdf_converter=MagicMock(),
            delivery_service=MagicMock(),
            office_libs={"docx": object()},
            allow_external_input_files=False,
            enable_docx_image_review=False,
            is_group_feature_enabled=lambda _event: True,
            check_permission=lambda _event: True,
            group_feature_disabled_error=lambda: "group disabled",
        )

        results = [
            result
            async for result in service.iter_read_file_tool_results(
                event, docx_path.name
            )
        ]
    finally:
        executor.shutdown(wait=False)
        shutil.rmtree(workspace_dir, ignore_errors=True)

    assert len(results) == 1
    assert isinstance(results[0], str)
    assert "图前说明" in results[0]
    assert "图后说明" in results[0]
    assert "[插图1]" not in results[0]
    workspace_service.extract_word_content.assert_called_once_with(
        docx_path,
        include_images=False,
    )


@pytest.mark.asyncio
async def test_file_tool_service_returns_message_for_image_only_docx_when_review_disabled():
    workspace_dir = _make_workspace("stream-docx-image-only-disabled")
    executor = ThreadPoolExecutor(max_workers=1)
    event = _build_event()
    docx_path = workspace_dir / "image-only.docx"
    image_path = workspace_dir / "embedded.png"

    try:
        _write_png(image_path)
        docx_path.write_bytes(b"docx")

        workspace_service = WorkspaceService(
            plugin_data_path=workspace_dir,
            executor=executor,
            office_libs={"docx": object()},
            max_file_size=1024 * 1024,
            feature_settings={},
        )
        workspace_service.extract_word_content = MagicMock(
            return_value=SimpleNamespace(
                text=None,
                image_paths=[image_path],
                image_count=1,
                items=[
                    SimpleNamespace(
                        type="image",
                        image_path=image_path,
                        image_index=1,
                    )
                ],
            )
        )

        service = FileToolService(
            workspace_service=workspace_service,
            office_generator=MagicMock(),
            pdf_converter=MagicMock(),
            delivery_service=MagicMock(),
            office_libs={"docx": object()},
            allow_external_input_files=False,
            enable_docx_image_review=False,
            is_group_feature_enabled=lambda _event: True,
            check_permission=lambda _event: True,
            group_feature_disabled_error=lambda: "group disabled",
        )

        results = [
            result
            async for result in service.iter_read_file_tool_results(
                event, docx_path.name
            )
        ]
    finally:
        executor.shutdown(wait=False)
        shutil.rmtree(workspace_dir, ignore_errors=True)

    assert len(results) == 1
    assert isinstance(results[0], str)
    assert "仅包含图片内容" in results[0]
    assert "[插图1]" not in results[0]
    workspace_service.extract_word_content.assert_called_once_with(
        docx_path,
        include_images=False,
    )


@pytest.mark.asyncio
async def test_file_tool_service_returns_local_guidance_for_missing_file():
    workspace_dir = Path(__file__).resolve().parent
    executor = ThreadPoolExecutor(max_workers=1)
    event = _build_event()

    try:
        workspace_service = WorkspaceService(
            plugin_data_path=workspace_dir,
            executor=executor,
            office_libs={},
            max_file_size=1024 * 1024,
            feature_settings={},
        )
        service = FileToolService(
            workspace_service=workspace_service,
            office_generator=MagicMock(),
            pdf_converter=MagicMock(),
            delivery_service=MagicMock(),
            office_libs={},
            allow_external_input_files=False,
            is_group_feature_enabled=lambda _event: True,
            check_permission=lambda _event: True,
            group_feature_disabled_error=lambda: "group disabled",
        )

        result = await service.read_file(event, "CLAUDE.md")
    finally:
        executor.shutdown(wait=False)

    assert result is not None
    assert "错误：文件 'CLAUDE.md' 不存在。" in result
    assert "不要联网搜索" in result
    assert "重新上传文件或提供正确的本地路径" in result


@pytest.mark.asyncio
async def test_file_tool_service_creates_office_file_via_generator_and_delivery():
    workspace_dir = Path(__file__).resolve().parent
    executor = ThreadPoolExecutor(max_workers=1)
    event = _build_event()
    event.send = AsyncMock()
    office_generator = MagicMock()
    office_generator.generate = AsyncMock(return_value=Path(__file__).resolve())
    delivery_service = MagicMock()
    delivery_service.send_file_with_preview = AsyncMock()

    try:
        workspace_service = WorkspaceService(
            plugin_data_path=workspace_dir,
            executor=executor,
            office_libs={"docx": object()},
            max_file_size=1024 * 1024,
            feature_settings={"enable_office_files": True},
        )
        service = FileToolService(
            workspace_service=workspace_service,
            office_generator=office_generator,
            pdf_converter=MagicMock(),
            delivery_service=delivery_service,
            office_libs={"docx": object()},
            allow_external_input_files=False,
            is_group_feature_enabled=lambda _event: True,
            check_permission=lambda _event: True,
            group_feature_disabled_error=lambda: "group disabled",
        )

        result = await service.create_office_file(
            event,
            filename="report.docx",
            content="hello world",
            file_type="word",
        )
    finally:
        executor.shutdown(wait=False)

    office_generator.generate.assert_awaited_once()
    delivery_service.send_file_with_preview.assert_awaited_once()
    assert result is None


@pytest.mark.asyncio
async def test_file_tool_service_create_office_file_returns_error_without_sending():
    workspace_dir = Path(__file__).resolve().parent
    executor = ThreadPoolExecutor(max_workers=1)
    event = _build_event()
    event.send = AsyncMock()
    office_generator = MagicMock()
    delivery_service = MagicMock()

    try:
        workspace_service = WorkspaceService(
            plugin_data_path=workspace_dir,
            executor=executor,
            office_libs={},
            max_file_size=1024 * 1024,
            feature_settings={"enable_office_files": True},
        )
        service = FileToolService(
            workspace_service=workspace_service,
            office_generator=office_generator,
            pdf_converter=MagicMock(),
            delivery_service=delivery_service,
            office_libs={},
            allow_external_input_files=False,
            is_group_feature_enabled=lambda _event: True,
            check_permission=lambda _event: True,
            group_feature_disabled_error=lambda: "group disabled",
        )

        result = await service.create_office_file(
            event,
            filename="report.unsupported",
            content="hello world",
            file_type="unknown",
        )
    finally:
        executor.shutdown(wait=False)

    assert result == "错误：不支持的文件类型 'unknown'。允许值：excel/powerpoint"
    event.send.assert_not_called()
    office_generator.generate.assert_not_called()
    delivery_service.send_file_with_preview.assert_not_called()


@pytest.mark.asyncio
async def test_file_tool_service_create_office_file_requires_explicit_type_without_suffix():
    workspace_dir = Path(__file__).resolve().parent
    executor = ThreadPoolExecutor(max_workers=1)
    event = _build_event()
    office_generator = MagicMock()
    delivery_service = MagicMock()

    try:
        workspace_service = WorkspaceService(
            plugin_data_path=workspace_dir,
            executor=executor,
            office_libs={"openpyxl": object()},
            max_file_size=1024 * 1024,
            feature_settings={"enable_office_files": True},
        )
        service = FileToolService(
            workspace_service=workspace_service,
            office_generator=office_generator,
            pdf_converter=MagicMock(),
            delivery_service=delivery_service,
            office_libs={"openpyxl": object()},
            allow_external_input_files=False,
            is_group_feature_enabled=lambda _event: True,
            check_permission=lambda _event: True,
            group_feature_disabled_error=lambda: "group disabled",
        )

        result = await service.create_office_file(
            event,
            filename="report",
            content="hello world",
            file_type="",
        )
    finally:
        executor.shutdown(wait=False)

    assert (
        result
        == "错误：未指定文件类型。请提供带后缀的文件名，或显式传入 file_type（excel/powerpoint）。"
    )
    office_generator.generate.assert_not_called()
    delivery_service.send_file_with_preview.assert_not_called()


@pytest.mark.asyncio
async def test_file_tool_service_create_office_file_rejects_word_fallback_without_suffix():
    workspace_dir = Path(__file__).resolve().parent
    executor = ThreadPoolExecutor(max_workers=1)
    event = _build_event()
    office_generator = MagicMock()
    delivery_service = MagicMock()

    try:
        workspace_service = WorkspaceService(
            plugin_data_path=workspace_dir,
            executor=executor,
            office_libs={"docx": object()},
            max_file_size=1024 * 1024,
            feature_settings={"enable_office_files": True},
        )
        service = FileToolService(
            workspace_service=workspace_service,
            office_generator=office_generator,
            pdf_converter=MagicMock(),
            delivery_service=delivery_service,
            office_libs={"docx": object()},
            allow_external_input_files=False,
            is_group_feature_enabled=lambda _event: True,
            check_permission=lambda _event: True,
            group_feature_disabled_error=lambda: "group disabled",
        )

        result = await service.create_office_file(
            event,
            filename="report",
            content="hello world",
            file_type="word",
        )
    finally:
        executor.shutdown(wait=False)

    assert (
        result
        == "错误：Word 文档请直接提供 .docx/.doc 文件名，或改用 create_document → "
        "add_blocks → finalize_document → export_document。"
    )
    office_generator.generate.assert_not_called()
    delivery_service.send_file_with_preview.assert_not_called()


@pytest.mark.asyncio
async def test_file_tool_service_create_office_file_returns_direct_result_for_explicit_tool_error():
    workspace_dir = Path(__file__).resolve().parent
    executor = ThreadPoolExecutor(max_workers=1)
    event = _build_event()
    event.get_extra.side_effect = lambda key, default=None: {
        EXPLICIT_FILE_TOOL_EVENT_KEY: "create_office_file"
    }.get(key, default)
    event.plain_result.side_effect = lambda text: f"DIRECT::{text}"
    office_generator = MagicMock()
    delivery_service = MagicMock()

    try:
        workspace_service = WorkspaceService(
            plugin_data_path=workspace_dir,
            executor=executor,
            office_libs={"docx": object()},
            max_file_size=1024 * 1024,
            feature_settings={"enable_office_files": True},
        )
        service = FileToolService(
            workspace_service=workspace_service,
            office_generator=office_generator,
            pdf_converter=MagicMock(),
            delivery_service=delivery_service,
            office_libs={"docx": object()},
            allow_external_input_files=False,
            is_group_feature_enabled=lambda _event: True,
            check_permission=lambda _event: True,
            group_feature_disabled_error=lambda: "group disabled",
        )

        result = await service.create_office_file(
            event,
            filename="report",
            content="hello world",
            file_type="word",
        )
    finally:
        executor.shutdown(wait=False)

    assert (
        result
        == "DIRECT::错误：Word 文档请直接提供 .docx/.doc 文件名，或改用 create_document → "
        "add_blocks → finalize_document → export_document。"
    )
    event.plain_result.assert_called_once()
    office_generator.generate.assert_not_called()
    delivery_service.send_file_with_preview.assert_not_called()


@pytest.mark.asyncio
async def test_file_tool_service_convert_to_pdf_returns_none_after_delivery():
    workspace_dir = Path(__file__).resolve().parent
    executor = ThreadPoolExecutor(max_workers=1)
    event = _build_event()
    source_path = workspace_dir / "convert-source.docx"
    source_path.write_text("demo", encoding="utf-8")
    output_path = workspace_dir / "convert-source.pdf"
    pdf_converter = MagicMock()
    pdf_converter.is_available.return_value = True
    pdf_converter.office_to_pdf = AsyncMock(return_value=output_path)
    delivery_service = MagicMock()
    delivery_service.send_file_with_preview = AsyncMock()

    try:
        workspace_service = WorkspaceService(
            plugin_data_path=workspace_dir,
            executor=executor,
            office_libs={"docx": object()},
            max_file_size=1024 * 1024,
            feature_settings={"enable_pdf_conversion": True},
        )
        service = FileToolService(
            workspace_service=workspace_service,
            office_generator=MagicMock(),
            pdf_converter=pdf_converter,
            delivery_service=delivery_service,
            office_libs={"docx": object()},
            allow_external_input_files=False,
            is_group_feature_enabled=lambda _event: True,
            check_permission=lambda _event: True,
            group_feature_disabled_error=lambda: "group disabled",
        )
        output_path.write_text("pdf", encoding="utf-8")

        result = await service.convert_to_pdf(
            event,
            filename=source_path.name,
        )
    finally:
        source_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)
        executor.shutdown(wait=False)

    pdf_converter.office_to_pdf.assert_awaited_once_with(source_path)
    delivery_service.send_file_with_preview.assert_awaited_once()
    assert result is None


def test_command_service_lists_office_files_in_workspace():
    workspace_dir = Path(__file__).resolve().parent / "workspace-command-list"
    workspace_dir.mkdir(exist_ok=True)
    sample_file = workspace_dir / "report.docx"
    sample_file.write_text("demo", encoding="utf-8")

    executor = ThreadPoolExecutor(max_workers=1)
    try:
        workspace_service = WorkspaceService(
            plugin_data_path=workspace_dir,
            executor=executor,
            office_libs={},
            max_file_size=1024 * 1024,
            feature_settings={},
        )
        pdf_converter = MagicMock()
        pdf_converter.capabilities = {
            "office_to_pdf": False,
            "pdf_to_word": False,
            "pdf_to_excel": False,
        }
        service = CommandService(
            workspace_service=workspace_service,
            pdf_converter=pdf_converter,
            plugin_data_path=workspace_dir,
            auto_delete=False,
            allow_external_input_files=False,
            enable_features_in_group=True,
            auto_block_execution_tools=True,
            reply_to_user=True,
            is_group_feature_enabled=lambda _event: True,
            check_permission=lambda _event: True,
            group_feature_disabled_error=lambda: "group disabled",
        )

        result = service.list_files(_build_event())
    finally:
        sample_file.unlink(missing_ok=True)
        workspace_dir.rmdir()
        executor.shutdown(wait=False)

    assert "机器人工作区 Office 文件列表" in result
    assert "report.docx" in result


def test_command_service_builds_pdf_status_summary():
    executor = ThreadPoolExecutor(max_workers=1)
    try:
        workspace_service = WorkspaceService(
            plugin_data_path=Path(__file__).resolve().parent,
            executor=executor,
            office_libs={},
            max_file_size=1024 * 1024,
            feature_settings={},
        )
        pdf_converter = MagicMock()
        pdf_converter.get_detailed_status.return_value = {
            "capabilities": {
                "office_to_pdf": True,
                "pdf_to_word": False,
                "pdf_to_excel": True,
            },
            "office_to_pdf_backend": "libreoffice",
            "word_backend": None,
            "excel_backend": "tabula",
            "is_windows": False,
            "java_available": True,
            "libreoffice_path": "/usr/bin/libreoffice",
            "libs": {"pdf2docx": False, "tabula-py": True},
        }
        pdf_converter.get_missing_dependencies.return_value = ["pdf2docx"]
        service = CommandService(
            workspace_service=workspace_service,
            pdf_converter=pdf_converter,
            plugin_data_path=Path(__file__).resolve().parent,
            auto_delete=False,
            allow_external_input_files=False,
            enable_features_in_group=True,
            auto_block_execution_tools=True,
            reply_to_user=True,
            is_group_feature_enabled=lambda _event: True,
            check_permission=lambda _event: True,
            group_feature_disabled_error=lambda: "group disabled",
        )

        result = service.pdf_status(_build_event())
    finally:
        executor.shutdown(wait=False)

    assert "Office→PDF: ✅ 可用 (libreoffice)" in result
    assert "PDF→Excel:  ✅ 可用 (tabula)" in result
    assert "缺失依赖" in result
    assert "pdf2docx" in result
