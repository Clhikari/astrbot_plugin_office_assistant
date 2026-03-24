import re
from collections.abc import Awaitable, Callable
from pathlib import Path

import astrbot.api.message_components as Comp
from astrbot.api import logger
from astrbot.api.event import AstrMessageEvent
from astrbot.core.platform.message_type import MessageType
from astrbot.core.provider.entities import ProviderRequest

from ..constants import (
    ALL_OFFICE_SUFFIXES,
    EXPLICIT_FILE_TOOL_EVENT_KEY,
    EXECUTION_TOOLS,
    FILE_TOOLS,
    NOTICE_DOCUMENT_TOOLS_GUIDE,
    NOTICE_TOOLS_DENIED,
    NOTICE_UPLOADED_FILE_TEMPLATE,
    PDF_SUFFIX,
    TEXT_SUFFIXES,
)
from ..internal_hooks import (
    NoticeBuildContext,
    NoticeBuildHook,
    ToolExposureContext,
    ToolExposureHook,
    run_notice_hooks,
    run_tool_exposure_hooks,
)


class LLMRequestPolicy:
    _NEGATIVE_TOOL_PREFIX_RE = re.compile(
        r"(?:不要|别|勿|不用|无需|do\s+not|don't|not)\s*(?:调用|使用|call|use|invoke)?\s*$",
        flags=re.IGNORECASE,
    )
    _BUFFERED_USER_INSTRUCTION_RE = re.compile(
        r"\[用户指令\]\s*(?P<instruction>.*?)(?:\n\s*\[|\Z)",
        flags=re.DOTALL,
    )

    def __init__(
        self,
        *,
        document_toolset,
        auto_block_execution_tools: bool,
        require_at_in_group: bool,
        is_group_feature_enabled: Callable[[AstrMessageEvent], bool],
        check_permission: Callable[[AstrMessageEvent], bool],
        is_bot_mentioned: Callable[[AstrMessageEvent], bool],
        get_cached_upload_infos: Callable[[AstrMessageEvent], list[dict]],
        extract_upload_source: Callable[
            [Comp.File], Awaitable[tuple[Path | None, str]]
        ],
        store_uploaded_file: Callable[[Path, str], Path],
        allow_external_input_files: bool,
        notice_hooks: list[NoticeBuildHook] | None = None,
        tool_exposure_hooks: list[ToolExposureHook] | None = None,
    ) -> None:
        self._document_toolset = document_toolset
        self._auto_block_execution_tools = auto_block_execution_tools
        self._require_at_in_group = require_at_in_group
        self._is_group_feature_enabled = is_group_feature_enabled
        self._check_permission = check_permission
        self._is_bot_mentioned = is_bot_mentioned
        self._get_cached_upload_infos = get_cached_upload_infos
        self._extract_upload_source = extract_upload_source
        self._store_uploaded_file = store_uploaded_file
        self._allow_external_input_files = allow_external_input_files
        self._notice_hooks = notice_hooks or [
            self._append_document_tool_guide_notice,
            self._append_uploaded_file_notices,
        ]
        self._tool_exposure_hooks = tool_exposure_hooks or [
            self._apply_execution_tool_block,
            self._apply_explicit_file_tool_restriction,
        ]

    def _detect_explicit_file_tool(self, text: str) -> str | None:
        if not text:
            return None

        explicit_matches: set[str] = set()
        for tool_name in sorted(FILE_TOOLS, key=len, reverse=True):
            patterns = (
                rf"(?P<tool>(?:调用|使用|call|use|invoke)\s*`?{re.escape(tool_name)}`?)",
                rf"(?P<tool>`{re.escape(tool_name)}`)",
                rf"(?P<tool>{re.escape(tool_name)}\s*\()",
                rf"(?P<tool>{re.escape(tool_name)}\s*[,，]\s*[a-zA-Z_]\w*\s*=)",
            )
            for pattern in patterns:
                for match in re.finditer(pattern, text, flags=re.IGNORECASE):
                    tool_start = match.start("tool")
                    prefix = text[max(0, tool_start - 20) : tool_start]
                    if self._NEGATIVE_TOOL_PREFIX_RE.search(prefix):
                        continue
                    explicit_matches.add(tool_name)
                    break

        if len(explicit_matches) == 1:
            return next(iter(explicit_matches))
        return None

    def _extract_explicit_tool_text(
        self,
        event: AstrMessageEvent,
        req: ProviderRequest,
    ) -> str:
        prompt_text = str(req.prompt or "").strip()
        if not prompt_text:
            return ""

        is_buffered_prompt = getattr(event, "_buffered", False) is True
        if is_buffered_prompt or "[System Notice]" in prompt_text:
            match = self._BUFFERED_USER_INSTRUCTION_RE.search(prompt_text)
            if match:
                return match.group("instruction").strip()
            return ""

        return prompt_text

    def _restrict_to_explicit_file_tool(
        self, req: ProviderRequest, explicit_tool_name: str
    ) -> None:
        if not req.func_tool:
            return

        for tool_name in FILE_TOOLS:
            if tool_name != explicit_tool_name:
                req.func_tool.remove_tool(tool_name)

    async def _run_before_expose_tools(
        self, context: ToolExposureContext
    ) -> ToolExposureContext:
        return await run_tool_exposure_hooks(self._tool_exposure_hooks, context)

    async def _run_before_build_notices(
        self, context: NoticeBuildContext
    ) -> NoticeBuildContext:
        return await run_notice_hooks(self._notice_hooks, context)

    async def _append_document_tool_guide_notice(
        self, context: NoticeBuildContext
    ) -> NoticeBuildContext:
        if context.should_expose and context.request.func_tool:
            context.notices.append(NOTICE_DOCUMENT_TOOLS_GUIDE)
        return context

    async def _append_uploaded_file_notices(
        self, context: NoticeBuildContext
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
                        f"[文件管理] 文件 {original_name} 格式不支持 ({file_suffix})，跳过处理"
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
                    f"[文件管理] 收到文件 {original_name}，已保存为 {stored_name}。"
                )
            except Exception as exc:
                logger.error(f"[文件管理] 处理上传文件失败: {exc}")
        return context

    async def _apply_execution_tool_block(
        self, context: ToolExposureContext
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

    async def _apply_explicit_file_tool_restriction(
        self, context: ToolExposureContext
    ) -> ToolExposureContext:
        if (
            context.should_expose
            and context.request.func_tool
            and context.explicit_tool_name
        ):
            self._restrict_to_explicit_file_tool(
                context.request, context.explicit_tool_name
            )
            logger.info(
                "[文件管理] 检测到用户显式指定工具 %s，本轮仅保留该文件工具",
                context.explicit_tool_name,
            )
        return context

    async def apply(self, event: AstrMessageEvent, req: ProviderRequest) -> None:
        is_group = event.message_obj.type == MessageType.GROUP_MESSAGE
        is_friend = event.message_obj.type == MessageType.FRIEND_MESSAGE
        explicit_tool_name = self._detect_explicit_file_tool(
            self._extract_explicit_tool_text(event, req)
        )
        event.set_extra(EXPLICIT_FILE_TOOL_EVENT_KEY, explicit_tool_name)

        if not self._is_group_feature_enabled(event):
            if req.func_tool:
                for tool_name in FILE_TOOLS:
                    req.func_tool.remove_tool(tool_name)
            req.system_prompt = (req.system_prompt or "") + NOTICE_TOOLS_DENIED
            logger.debug("[文件管理] 群聊总开关关闭，已隐藏全部文件工具")
            return

        has_permission = self._check_permission(event)
        can_process_upload = has_permission or event.is_admin()
        should_expose = (is_friend and event.is_admin()) or (
            has_permission
            and (
                not is_group
                or not self._require_at_in_group
                or self._is_bot_mentioned(event)
            )
        )

        if not should_expose:
            logger.debug(
                f"[文件管理] 用户 {event.get_sender_id()} 权限不足，已隐藏文件工具"
            )
            if req.func_tool:
                for tool_name in FILE_TOOLS:
                    req.func_tool.remove_tool(tool_name)
            req.system_prompt = (req.system_prompt or "") + NOTICE_TOOLS_DENIED

        if should_expose and req.func_tool:
            for tool in self._document_toolset.tools:
                req.func_tool.add_tool(tool)

        await self._run_before_expose_tools(
            ToolExposureContext(
                event=event,
                request=req,
                should_expose=should_expose,
                can_process_upload=can_process_upload,
                explicit_tool_name=explicit_tool_name,
            )
        )

        notice_context = await self._run_before_build_notices(
            NoticeBuildContext(
                event=event,
                request=req,
                should_expose=should_expose,
                can_process_upload=can_process_upload,
                explicit_tool_name=explicit_tool_name,
            )
        )
        if notice_context.notices:
            req.system_prompt = (req.system_prompt or "") + "".join(
                notice_context.notices
            )
