from collections.abc import AsyncGenerator
from pathlib import Path

import astrbot.api.message_components as Comp
import mcp
from astrbot.api import AstrBotConfig, llm_tool, logger
from astrbot.api.event import AstrMessageEvent, filter
from astrbot.api.star import Context, Star
from astrbot.core.agent.run_context import ContextWrapper
from astrbot.core.astr_agent_context import AstrAgentContext
from astrbot.core.message.message_event_result import MessageChain, MessageEventResult
from astrbot.core.provider.entities import ProviderRequest
from astrbot.core.star.filter.command import GreedyStr
from astrbot.core.star.star import star_map

from .app.runtime import PluginRuntimeBundle
from .message_buffer import BufferedMessage
from .services import build_plugin_runtime

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
        self._runtime = runtime
        self._post_export_hook_service = runtime.post_export_hook_service
        # 构造函数日志需要的快捷字段
        self._auto_delete = runtime.settings.auto_delete
        self.plugin_data_path = runtime.plugin_data_path
        runtime.message_buffer.set_complete_callback(self._on_buffer_complete)

    async def terminate(self):
        """插件卸载时释放资源"""
        rt = getattr(self, "_runtime", None)
        if rt is None:
            return

        if getattr(rt, "message_buffer", None):
            rt.message_buffer.set_complete_callback(None)

        if rt.office_gen:
            rt.office_gen.cleanup()
            logger.debug("[文件管理] Office生成器已清理")

        if rt.pdf_converter:
            rt.pdf_converter.cleanup()
            logger.debug("[文件管理] PDF转换器已清理")

        if rt.executor:
            rt.executor.shutdown(wait=False)
            logger.debug("[文件管理] 主线程池已关闭")

        if rt.temp_dir:
            try:
                rt.temp_dir.cleanup()
                logger.debug("[文件管理] 临时目录已清理")
            except Exception as e:
                logger.warning(f"[文件管理] 清理临时目录失败: {e}")

        self._runtime = None

    async def _on_buffer_complete(self, buf: BufferedMessage):
        rt = getattr(self, "_runtime", None)
        if rt is None:
            logger.warning("[文件管理] 缓冲区完成回调触发时运行时已释放，忽略此次处理")
            return

        await rt.upload_session_service.on_buffer_complete(buf)

    async def _extract_upload_source(
        self, component: Comp.File
    ) -> tuple[Path | None, str]:
        """Extract local source path and display name from upload component."""
        file_path = await component.get_file()
        if not file_path:
            return None, component.name or "unknown_file"
        return Path(file_path), component.name or Path(file_path).name

    def _store_uploaded_file(self, src_path: Path, preferred_name: str) -> Path:
        return self._runtime.workspace_service.store_uploaded_file(
            src_path, preferred_name
        )

    async def _handle_exported_document_tool(
        self, context: ContextWrapper[AstrAgentContext], output_path: str
    ) -> str | None:
        service = getattr(self, "_post_export_hook_service", None)
        if service is None:
            logger.warning("[文件管理] 导出回调触发时发送服务不可用，跳过文件回传")
            return None

        return await service.handle_exported_document_tool(context, output_path)

    @filter.event_message_type(filter.EventMessageType.ALL, priority=100)
    async def on_file_message(self, event: AstrMessageEvent):
        """
        拦截包含文件的消息，使用缓冲器聚合文件和后续文本消息
        仅处理支持的文件格式（Office、文本、PDF），其他格式直接放行
        """
        await self._runtime.incoming_message_service.handle_file_message(event)

    @filter.on_llm_request()
    async def before_llm_chat(self, event: AstrMessageEvent, req: ProviderRequest):
        await self._runtime.llm_request_policy.apply(event, req)

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
        async for result in self._runtime.file_tool_service.iter_read_file_tool_results(
            event,
            filename,
        ):
            yield result

    @llm_tool(name="read_workbook")
    async def read_workbook(
        self,
        event: AstrMessageEvent,
        filename: str = "",
    ) -> AsyncGenerator[str | mcp.types.CallToolResult, None]:
        """读取已有 Excel 工作簿内容。

        适用场景：
        - 读取 `.xlsx` / `.xls`
        - 查看 Sheet 内容、统计、汇总、解释

        注意：
        - 这是 Excel 专用读取入口
        - 不创建 workbook session，不导出新文件

        Args:
            filename(string): 要读取的 Excel 文件名。
        """
        async for result in self._runtime.file_tool_service.iter_read_workbook_tool_results(
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
        return await self._runtime.file_tool_service.create_office_file(
            event,
            filename=filename,
            content=content,
            file_type=file_type,
        )

    @llm_tool(name="execute_excel_script")
    async def execute_excel_script(
        self,
        event: AstrMessageEvent,
        script: str = "",
        input_files: list[str] | None = None,
        output_name: str = "",
    ) -> str:
        """执行 Excel 脚本以生成或修改工作簿。

        适用场景：
        - 新建包含公式、图表、条件格式、数据验证的复杂 Excel
        - 修改已有 `.xlsx` / `.xls`

        Args:
            script(string): 要执行的 Python 脚本。
            input_files(array): 作为输入的 Excel 文件列表。
            output_name(string): 需要返回文件时的输出文件名。
        """
        return await self._runtime.file_tool_service.execute_excel_script(
            event,
            script=script,
            input_files=input_files,
            output_name=output_name or None,
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
        return await self._runtime.file_tool_service.convert_to_pdf(
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
        return await self._runtime.file_tool_service.convert_from_pdf(
            event,
            filename=filename,
            target_format=target_format,
            file_id=file_id,
        )

    @filter.command("delete_file", alias={"删除文件", "file_rm"})
    async def delete_file(self, event: AstrMessageEvent):
        """从工作区中永久删除指定文件。用法: /delete_file 文件名"""
        result = self._runtime.command_service.delete_file(event, event.message_str)
        await event.send(MessageChain().message(result))

    @filter.command("fileinfo")
    async def fileinfo(self, event: AstrMessageEvent):
        """显示文件管理工具的运行信息"""
        yield event.plain_result(self._runtime.command_service.fileinfo(event))

    @filter.command("list_files", alias={"文件列表", "file_ls"})
    async def list_files(self, event: AstrMessageEvent):
        """列出机器人工作区中的 Office 文件。"""
        await event.send(
            MessageChain().message(self._runtime.command_service.list_files(event))
        )

    @filter.command("pdf_status", alias={"pdf状态"})
    async def pdf_status(self, event: AstrMessageEvent):
        """显示 PDF 转换功能的状态和依赖信息"""
        yield event.plain_result(self._runtime.command_service.pdf_status(event))

    @filter.command_group("doc")
    def doc(self):
        """处理当前会话中的上传文件。"""

    @doc.command("list")
    async def doc_list(self, event: AstrMessageEvent):
        """查看当前会话可用的上传文件。"""
        result = self._runtime.command_service.doc_list(event)
        event.set_result(MessageEventResult().message(result).stop_event())

    @doc.command("clear")
    async def doc_clear(self, event: AstrMessageEvent, file_id: str = ""):
        """清理当前会话中的上传文件。"""
        result = self._runtime.command_service.doc_clear(event, file_id)
        event.set_result(MessageEventResult().message(result).stop_event())

    @doc.command("use")
    async def doc_use(
        self,
        event: AstrMessageEvent,
        selection: GreedyStr,
    ):
        """使用指定文件继续处理请求。"""
        result = await self._runtime.command_service.doc_use(
            event,
            str(selection),
        )
        if result:
            event.set_result(MessageEventResult().message(result).stop_event())
            return
        event.stop_event()

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
        await self._runtime.error_hook_service.handle_plugin_error(
            event,
            plugin_name,
            handler_name,
            error,
            traceback_text,
        )
