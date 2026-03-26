import re
from collections.abc import Callable

from astrbot.api import logger
from astrbot.api.event import AstrMessageEvent
from astrbot.core.platform.message_type import MessageType
from astrbot.core.provider.entities import ProviderRequest

from ..constants import (
    EXPLICIT_FILE_TOOL_EVENT_KEY,
    FILE_TOOLS,
    NOTICE_TOOLS_DENIED,
)
from ..internal_hooks import (
    NoticeBuildContext,
    NoticeBuildHook,
    ToolExposureContext,
    ToolExposureHook,
    run_notice_hooks,
    run_tool_exposure_hooks,
)
from .request_hook_service import RequestHookService


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
        get_cached_upload_infos,
        extract_upload_source,
        store_uploaded_file,
        allow_external_input_files: bool,
        notice_hooks: list[NoticeBuildHook] | None = None,
        tool_exposure_hooks: list[ToolExposureHook] | None = None,
    ) -> None:
        self._document_toolset = document_toolset
        self._require_at_in_group = require_at_in_group
        self._is_group_feature_enabled = is_group_feature_enabled
        self._check_permission = check_permission
        self._is_bot_mentioned = is_bot_mentioned
        request_hook_service = RequestHookService(
            auto_block_execution_tools=auto_block_execution_tools,
            get_cached_upload_infos=get_cached_upload_infos,
            extract_upload_source=extract_upload_source,
            store_uploaded_file=store_uploaded_file,
            allow_external_input_files=allow_external_input_files,
        )
        self._notice_hooks = notice_hooks or request_hook_service.build_notice_hooks()
        self._tool_exposure_hooks = (
            tool_exposure_hooks or request_hook_service.build_tool_exposure_hooks()
        )

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
            return prompt_text

        return prompt_text

    async def _run_before_expose_tools(
        self, context: ToolExposureContext
    ) -> ToolExposureContext:
        return await run_tool_exposure_hooks(self._tool_exposure_hooks, context)

    async def _run_before_build_notices(
        self, context: NoticeBuildContext
    ) -> NoticeBuildContext:
        return await run_notice_hooks(self._notice_hooks, context)

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
