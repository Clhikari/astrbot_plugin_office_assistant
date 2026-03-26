from collections.abc import Awaitable, Callable
from pathlib import Path

import astrbot.api.message_components as Comp
from astrbot.api import logger
from astrbot.api.event import AstrMessageEvent

from ..constants import (
    ALL_OFFICE_SUFFIXES,
    EXECUTION_TOOLS,
    FILE_TOOLS,
    NOTICE_DOCUMENT_TOOLS_GUIDE,
    NOTICE_UPLOADED_FILE_TEMPLATE,
    PDF_SUFFIX,
    TEXT_SUFFIXES,
)
from ..internal_hooks import (
    NoticeBuildContext,
    NoticeBuildHook,
    ToolExposureContext,
    ToolExposureHook,
)


class RequestHookService:
    def __init__(
        self,
        *,
        auto_block_execution_tools: bool,
        get_cached_upload_infos: Callable[[AstrMessageEvent], list[dict]],
        extract_upload_source: Callable[
            [Comp.File], Awaitable[tuple[Path | None, str]]
        ],
        store_uploaded_file: Callable[[Path, str], Path],
        allow_external_input_files: bool,
    ) -> None:
        self._auto_block_execution_tools = auto_block_execution_tools
        self._get_cached_upload_infos = get_cached_upload_infos
        self._extract_upload_source = extract_upload_source
        self._store_uploaded_file = store_uploaded_file
        self._allow_external_input_files = allow_external_input_files

    def build_notice_hooks(self) -> list[NoticeBuildHook]:
        return [
            self.append_document_tool_guide_notice,
            self.append_uploaded_file_notices,
        ]

    def build_tool_exposure_hooks(self) -> list[ToolExposureHook]:
        return [
            self.apply_execution_tool_block,
            self.apply_explicit_file_tool_restriction,
        ]

    async def append_document_tool_guide_notice(
        self,
        context: NoticeBuildContext,
    ) -> NoticeBuildContext:
        if context.should_expose and context.request.func_tool:
            context.notices.append(NOTICE_DOCUMENT_TOOLS_GUIDE)
        return context

    async def append_uploaded_file_notices(
        self,
        context: NoticeBuildContext,
    ) -> NoticeBuildContext:
        if not context.can_process_upload:
            return context

        event = context.event
        req = context.request
        cached_upload_infos = iter(self._get_cached_upload_infos(event))
        for component in getattr(event.message_obj, "message", None) or []:
            if not isinstance(component, Comp.File):
                continue

            try:
                cached_info = next(cached_upload_infos, None)
                original_name = component.name or ""
                stored_name = ""
                source_path_text = ""
                file_suffix = Path(original_name).suffix.lower()
                type_desc = ""
                is_supported = False

                if cached_info:
                    original_name = cached_info.get("original_name", original_name)
                    stored_name = cached_info.get("stored_name", "")
                    source_path_text = cached_info.get("source_path", "")
                    file_suffix = cached_info.get("file_suffix", file_suffix)
                    type_desc = cached_info.get("type_desc", "")
                    is_supported = bool(cached_info.get("is_supported", False))

                if not cached_info or (is_supported and not stored_name):
                    src_path, original_name = await self._extract_upload_source(
                        component
                    )
                    if not src_path or not src_path.exists():
                        continue

                    source_path_text = str(src_path.resolve())
                    stored_path = self._store_uploaded_file(src_path, original_name)
                    stored_name = stored_path.name
                    file_suffix = stored_path.suffix.lower()

                    if file_suffix in ALL_OFFICE_SUFFIXES:
                        type_desc = "Office文档 (Word/Excel/PPT)"
                        is_supported = True
                    elif file_suffix in TEXT_SUFFIXES:
                        type_desc = "文本/代码文件"
                        is_supported = True
                    elif file_suffix == PDF_SUFFIX:
                        type_desc = "PDF文档"
                        is_supported = True

                if not is_supported:
                    logger.info(
                        "[文件管理] 文件 %s 格式不支持 (%s)，跳过处理",
                        original_name,
                        file_suffix,
                    )
                    continue

                if not context.should_expose or not req.func_tool:
                    logger.info(
                        "[文件管理] 文件 %s 已保存为 %s，但当前轮未暴露文件工具或未附加函数工具，跳过上传文件提示注入",
                        original_name,
                        stored_name or original_name,
                    )
                    continue

                prompt = NOTICE_UPLOADED_FILE_TEMPLATE.format(
                    type_desc=type_desc,
                    original_name=original_name,
                    file_suffix=file_suffix,
                    stored_name=stored_name,
                    external_path_line=(
                        f"- 外部绝对路径：{source_path_text}\n"
                        if self._allow_external_input_files and source_path_text
                        else ""
                    ),
                    external_path_rule=(
                        f"如果需要使用工作区外路径，也可以直接使用绝对路径 `{source_path_text}`。"
                        if self._allow_external_input_files and source_path_text
                        else "当前未启用外部绝对路径，不要使用工作区外路径。"
                    ),
                )
                context.notices.append(prompt)
                logger.info(
                    "[文件管理] 收到文件 %s，已保存为 %s。",
                    original_name,
                    stored_name,
                )
            except Exception as exc:
                logger.error(f"[文件管理] 处理上传文件失败: {exc}")
        return context

    async def apply_execution_tool_block(
        self,
        context: ToolExposureContext,
    ) -> ToolExposureContext:
        if (
            context.should_expose
            and context.request.func_tool
            and self._auto_block_execution_tools
        ):
            for tool_name in EXECUTION_TOOLS:
                context.request.func_tool.remove_tool(tool_name)
            logger.debug("[文件管理] 已自动屏蔽 shell/python 执行类工具")
        return context

    async def apply_explicit_file_tool_restriction(
        self,
        context: ToolExposureContext,
    ) -> ToolExposureContext:
        if (
            context.should_expose
            and context.request.func_tool
            and context.explicit_tool_name
        ):
            for tool_name in FILE_TOOLS:
                if tool_name != context.explicit_tool_name:
                    context.request.func_tool.remove_tool(tool_name)
            logger.info(
                "[文件管理] 检测到用户显式指定工具 %s，本轮仅保留该文件工具",
                context.explicit_tool_name,
            )
        return context
