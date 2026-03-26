from collections.abc import AsyncGenerator
from pathlib import Path

import astrbot.api.message_components as Comp
import mcp
from astrbot.api import AstrBotConfig, llm_tool, logger
from astrbot.api.event import AstrMessageEvent, filter
from astrbot.api.star import Context, Star
from astrbot.core.agent.run_context import ContextWrapper
from astrbot.core.astr_agent_context import AstrAgentContext
from astrbot.core.message.message_event_result import MessageChain
from astrbot.core.provider.entities import ProviderRequest
from astrbot.core.star.star import star_map

from .constants import DEFAULT_CHUNK_SIZE, MSG_DOCUMENT_EXPORTED, OfficeType
from .message_buffer import BufferedMessage
from .services import PluginRuntimeBundle, build_plugin_runtime

# 向后兼容性：旧版本的 AstrBot 未公开此钩子.
_on_plugin_error_decorator = getattr(filter, "on_plugin_error", None)
if _on_plugin_error_decorator is None:
    logger.debug(
        "[plugin-error-hook] on_plugin_error is unavailable in this AstrBot version; fallback to default error handling."
    )

    def on_plugin_error_filter(*args, **kwargs):
        def decorator(func):
            return func

        return decorator

else:
    on_plugin_error_filter = _on_plugin_error_decorator


class FileOperationPlugin(Star):
    """基于工具调用的智能文件管理插件"""

    def __init__(self, context: Context, config: AstrBotConfig):
        super().__init__(context)
        self.config = config
        plugin_name = self._resolve_plugin_name()
        runtime = build_plugin_runtime(
            context=self.context,
            config=self.config,
            plugin_name=plugin_name,
            handle_exported_document_tool=self._handle_exported_document_tool,
            extract_upload_source=lambda component: self._extract_upload_source(
                component
            ),
            store_uploaded_file=lambda src_path, original_name: (
                self._store_uploaded_file(src_path, original_name)
            ),
        )
        self._apply_runtime(runtime)

        mode = "临时目录(自动删除)" if self._auto_delete else "持久化存储"
        logger.info(
            f"[文件管理] 插件加载完成。模式: {mode}, 数据目录: {self.plugin_data_path}"
        )

    def _resolve_plugin_name(self) -> str:
        plugin_name = getattr(self, "name", None)
        if isinstance(plugin_name, str) and plugin_name.strip():
            return plugin_name

        metadata = star_map.get(self.__class__.__module__)
        if metadata and metadata.name:
            return metadata.name

        module_parts = self.__class__.__module__.split(".")
        if len(module_parts) >= 2:
            return module_parts[-2]
        return self.__class__.__module__

    def _apply_runtime(self, runtime: PluginRuntimeBundle) -> None:
        settings = runtime.settings
        self._auto_delete = settings.auto_delete
        self._max_file_size = settings.max_file_size
        self._buffer_wait = settings.buffer_wait
        self._reply_to_user = settings.reply_to_user
        self._require_at_in_group = settings.require_at_in_group
        self._enable_features_in_group = settings.enable_features_in_group
        self._auto_block_execution_tools = settings.auto_block_execution_tools
        self._enable_preview = settings.enable_preview
        self._preview_dpi = settings.preview_dpi
        self._allow_external_input_files = settings.allow_external_input_files
        self._feature_settings = settings.feature_settings
        self._recent_text_ttl_seconds = settings.recent_text_ttl_seconds
        self._recent_text_max_entries = settings.recent_text_max_entries
        self._recent_text_cleanup_interval_seconds = (
            settings.recent_text_cleanup_interval_seconds
        )

        self._temp_dir = runtime.temp_dir
        self.plugin_data_path = runtime.plugin_data_path
        self._executor = runtime.executor
        self.office_gen = runtime.office_gen
        self.pdf_converter = runtime.pdf_converter
        self.preview_gen = runtime.preview_gen
        self._office_libs = runtime.office_libs
        self._workspace_service = runtime.workspace_service
        self._access_policy_service = runtime.access_policy_service
        self._upload_session_service = runtime.upload_session_service
        self._recent_text_by_session = runtime.recent_text_by_session
        self._document_toolset = runtime.document_toolset
        self._llm_request_policy = runtime.llm_request_policy
        self._delivery_service = runtime.delivery_service
        self._post_export_hook_service = runtime.post_export_hook_service
        self._file_tool_service = runtime.file_tool_service
        self._command_service = runtime.command_service
        self._error_hook_service = runtime.error_hook_service
        self._message_buffer = runtime.message_buffer
        self._message_buffer.set_complete_callback(self._on_buffer_complete)
        self._incoming_message_service = runtime.incoming_message_service

    async def terminate(self):
        """插件卸载时释放资源"""
        # 清理 Office 生成器资源
        if hasattr(self, "office_gen") and self.office_gen:
            self.office_gen.cleanup()
            logger.debug("[文件管理] Office生成器已清理")

        # 清理 PDF 转换器资源
        if hasattr(self, "pdf_converter") and self.pdf_converter:
            self.pdf_converter.cleanup()
            logger.debug("[文件管理] PDF转换器已清理")

        # 关闭主线程池（子模块使用共享线程池，不会自己关闭）
        if hasattr(self, "_executor") and self._executor:
            self._executor.shutdown(wait=False)
            logger.debug("[文件管理] 主线程池已关闭")

        # 清理临时目录
        if hasattr(self, "_temp_dir") and self._temp_dir:
            try:
                self._temp_dir.cleanup()
                logger.debug("[文件管理] 临时目录已清理")
            except Exception as e:
                logger.warning(f"[文件管理] 清理临时目录失败: {e}")

    async def _on_buffer_complete(self, buf: BufferedMessage):
        await self._upload_session_service.on_buffer_complete(buf)

    def _check_permission(self, event: AstrMessageEvent) -> bool:
        return self._access_policy_service.check_permission(event)

    def _is_group_message(self, event: AstrMessageEvent) -> bool:
        return self._access_policy_service.is_group_message(event)

    def _is_group_feature_enabled(self, event: AstrMessageEvent) -> bool:
        return self._access_policy_service.is_group_feature_enabled(event)

    def _group_feature_disabled_error(self) -> str:
        return self._access_policy_service.group_feature_disabled_error()

    def _is_bot_mentioned(self, event: AstrMessageEvent) -> bool:
        return self._access_policy_service.is_bot_mentioned(event)

    def _validate_path(
        self, filename: str, *, allow_external: bool = False
    ) -> tuple[bool, Path, str]:
        return self._workspace_service.validate_path(
            filename, allow_external=allow_external
        )

    def _display_name(self, filename: str | Path) -> str:
        return self._workspace_service.display_name(filename)

    def _get_attachment_session_key(
        self, event: AstrMessageEvent
    ) -> tuple[str, str, str]:
        return self._upload_session_service.get_attachment_session_key(event)

    def _cleanup_recent_text_cache(self, now: float, *, force: bool = False) -> None:
        self._upload_session_service.cleanup_recent_text_cache(now, force=force)

    def _store_uploaded_file(self, src_path: Path, preferred_name: str) -> Path:
        return self._workspace_service.store_uploaded_file(src_path, preferred_name)

    def _try_copy_uploaded_file(self, src_path: Path, dst_path: Path) -> bool:
        return self._workspace_service.try_copy_uploaded_file(src_path, dst_path)

    def _remember_recent_text(self, event: AstrMessageEvent) -> None:
        self._upload_session_service.remember_recent_text(event)

    def _pop_recent_text(self, event: AstrMessageEvent) -> str:
        return self._upload_session_service.pop_recent_text(event)

    async def _extract_upload_source(
        self, component: Comp.File
    ) -> tuple[Path | None, str]:
        """Extract local source path and display name from upload component."""
        file_path = await component.get_file()
        if not file_path:
            return None, component.name or "unknown_file"
        return Path(file_path), component.name or Path(file_path).name

    def _pre_check(
        self,
        event: AstrMessageEvent,
        filename: str | None = None,
        *,
        check_permission: bool = True,
        feature_key: str | None = None,
        require_exists: bool = False,
        allowed_suffixes: frozenset | set | None = None,
        required_suffix: str | None = None,
        allow_external_path: bool = False,
    ) -> tuple[bool, Path | None, str | None]:
        return self._workspace_service.pre_check(
            event,
            filename,
            check_permission=check_permission,
            feature_key=feature_key,
            require_exists=require_exists,
            allowed_suffixes=allowed_suffixes,
            required_suffix=required_suffix,
            allow_external_path=allow_external_path,
            is_group_feature_enabled=self._is_group_feature_enabled,
            check_permission_fn=self._check_permission,
            group_feature_disabled_error=self._group_feature_disabled_error,
        )

    def _get_max_file_size(self) -> int:
        return self._workspace_service.get_max_file_size()

    async def _send_file_with_preview(
        self,
        event: AstrMessageEvent,
        file_path: Path,
        success_message: str = "✅ 文件已处理成功",
    ) -> None:
        await self._delivery_service.send_file_with_preview(
            event,
            file_path,
            success_message,
        )

    async def _handle_exported_document_tool(
        self, context: ContextWrapper[AstrAgentContext], output_path: str
    ) -> str | None:
        return await self._post_export_hook_service.handle_exported_document_tool(
            context, output_path
        )

    async def _send_exported_document(
        self, event: AstrMessageEvent, file_path: Path
    ) -> str:
        return await self._delivery_service.send_exported_document(
            event,
            file_path,
            MSG_DOCUMENT_EXPORTED,
        )

    async def _read_text_file(
        self, file_path: Path, max_size: int, chunk_size: int = DEFAULT_CHUNK_SIZE
    ) -> str:
        return await self._workspace_service.read_text_file(
            file_path, max_size, chunk_size
        )

    def _read_text_file_sync(
        self, file_path: Path, max_size: int, chunk_size: int
    ) -> str:
        return self._workspace_service.read_text_file_sync(
            file_path, max_size, chunk_size
        )

    def _extract_office_text(
        self, file_path: Path, office_type: OfficeType
    ) -> str | None:
        return self._workspace_service.extract_office_text(file_path, office_type)

    def _format_file_result(
        self, filename: str, suffix: str, file_size: int, content: str
    ) -> str:
        return self._workspace_service.format_file_result(
            filename, suffix, file_size, content
        )

    def _extract_pdf_text(self, file_path: Path) -> str | None:
        return self._workspace_service.extract_pdf_text(file_path)

    @filter.event_message_type(filter.EventMessageType.ALL, priority=100)
    async def on_file_message(self, event: AstrMessageEvent):
        """
        拦截包含文件的消息，使用缓冲器聚合文件和后续文本消息
        仅处理支持的文件格式（Office、文本、PDF），其他格式直接放行
        """
        await self._incoming_message_service.handle_file_message(event)

    @filter.on_llm_request()
    async def before_llm_chat(self, event: AstrMessageEvent, req: ProviderRequest):
        await self._llm_request_policy.apply(event, req)

    @llm_tool(name="read_file")
    async def read_file(
        self,
        event: AstrMessageEvent,
        filename: str = "",
    ) -> AsyncGenerator[str | mcp.types.CallToolResult, None]:
        """读取文本、Office 或 PDF 文件内容。

        支持格式：
        - 文本：.txt、.md、.log、.py、.js、.ts、.json、.yaml、.xml、.csv、.html、.css、.sql 等
        - Office：.docx、.xlsx、.pptx、.doc、.xls、.ppt
        - PDF：.pdf

        注意：
        - 不支持图片、视频、音频等二进制文件
        - 如果文件不存在或路径非法，直接告知用户并请其重新上传，NEVER 调用网络搜索

        Args:
            filename(string): 要读取的文件名。
        """
        async for result in self._file_tool_service.iter_read_file_tool_results(
            event,
            filename,
        ):
            yield result

    @llm_tool(name="create_office_file")
    async def create_office_file(
        self,
        event: AstrMessageEvent,
        filename: str = "",
        content: str = "",
        file_type: str = "",
    ):
        """[DEPRECATED] 创建简单的 Office 文件（Excel/PPT）并发送给用户。

        ⚠️ 对于 Word 文档，MUST 改用文档工具链：
        create_document → add_blocks → finalize_document → export_document
        此工具仅建议用于简单的一次性 Excel/PPT 输出。
        如果用户显式点名 `create_office_file` 并给出参数，MUST 先调用此工具；
        即使预期会报错，也不要擅自改成 `create_document` 或其他工具。

        内容格式：
        - Excel：用 `|` 分隔单元格，换行分隔行，如 `Name|Age\\nAlice|25`
        - PowerPoint：用 `[Slide 1]` 标记幻灯片页

        Args:
            filename(string): 输出文件名（.docx/.xlsx/.pptx）。
            content(string): 按上述格式提供的文件内容。
            file_type(string): 当文件名没有后缀时必须显式指定，支持 `excel` / `powerpoint`。
        """
        return await self._file_tool_service.create_office_file(
            event,
            filename=filename,
            content=content,
            file_type=file_type,
        )

    @llm_tool(name="convert_to_pdf")
    async def convert_to_pdf(
        self,
        event: AstrMessageEvent,
        filename: str = "",
        file_path: str = "",  # 别名，兼容 LLM 可能使用的参数名
    ) -> str | None:
        """把 Office 文件转换为 PDF。支持 .docx/.doc、.xlsx/.xls、.pptx/.ppt。
        直接调用即可，不需要先调用 read_file。

        Args:
            filename(string): 要转换的 Office 文件名，例如 report.docx。
        """
        return await self._file_tool_service.convert_to_pdf(
            event,
            filename=filename,
            file_path=file_path,
        )

    @llm_tool(name="convert_from_pdf")
    async def convert_from_pdf(
        self,
        event: AstrMessageEvent,
        filename: str = "",
        target_format: str = "word",
        file_id: str = "",  # 别名，兼容 LLM 可能使用的参数名
    ) -> str | None:
        """把 PDF 文件转换为 Word 或 Excel 格式。直接调用即可，不需要先调用 read_file。

        Args:
            filename(string): 要转换的 PDF 文件名，例如 document.pdf。
            target_format(string): 目标格式，`word` 或 `excel`，默认 `word`。
        """
        return await self._file_tool_service.convert_from_pdf(
            event,
            filename=filename,
            target_format=target_format,
            file_id=file_id,
        )

    @filter.command("delete_file", alias={"删除文件", "file_rm"})
    async def delete_file(self, event: AstrMessageEvent):
        """从工作区中永久删除指定文件。用法: /delete_file 文件名"""
        result = self._command_service.delete_file(event, event.message_str)
        await event.send(MessageChain().message(result))

    @filter.command("fileinfo")
    async def fileinfo(self, event: AstrMessageEvent):
        """显示文件管理工具的运行信息"""
        yield event.plain_result(self._command_service.fileinfo(event))

    @filter.command("list_files", alias={"文件列表", "file_ls"})
    async def list_files(self, event: AstrMessageEvent):
        """列出机器人工作区中的 Office 文件。"""
        await event.send(
            MessageChain().message(self._command_service.list_files(event))
        )

    @filter.command("pdf_status", alias={"pdf状态"})
    async def pdf_status(self, event: AstrMessageEvent):
        """显示 PDF 转换功能的状态和依赖信息"""
        yield event.plain_result(self._command_service.pdf_status(event))

    @on_plugin_error_filter()
    async def on_plugin_error(
        self,
        event: AstrMessageEvent,
        plugin_name: str,
        handler_name: str,
        error: Exception,
        traceback_text: str,
    ) -> None:
        """Intercept plugin errors and forward to target session."""
        await self._error_hook_service.handle_plugin_error(
            event,
            plugin_name,
            handler_name,
            error,
            traceback_text,
        )
