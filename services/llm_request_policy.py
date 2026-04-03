import re
from collections.abc import Callable
from dataclasses import dataclass

from astrbot.api import logger
from astrbot.api.event import AstrMessageEvent
from astrbot.core.platform.message_type import MessageType
from astrbot.core.provider.entities import ProviderRequest

from ..constants import (
    DOC_COMMAND_TRIGGER_EVENT_KEY,
    EXPLICIT_FILE_TOOL_EVENT_KEY,
    FILE_TOOLS,
)
from ..internal_hooks import (
    NoticeBuildContext,
    NoticeBuildHook,
    ToolExposureContext,
    ToolExposureHook,
    run_notice_hooks,
    run_tool_exposure_hooks,
)
from .prompt_context_service import PromptContextService
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

    @dataclass(slots=True)
    class ExposureDecision:
        explicit_tool_name: str | None
        can_process_upload: bool
        should_expose: bool

    def __init__(
        self,
        *,
        document_toolset,
        require_at_in_group: bool,
        is_group_feature_enabled: Callable[[AstrMessageEvent], bool],
        check_permission: Callable[[AstrMessageEvent], bool],
        is_bot_mentioned: Callable[[AstrMessageEvent], bool],
        request_hook_service: RequestHookService | None = None,
        prompt_context_service: PromptContextService | None = None,
        notice_hooks: list[NoticeBuildHook] | None = None,
        tool_exposure_hooks: list[ToolExposureHook] | None = None,
    ) -> None:
        self._document_toolset = document_toolset
        self._require_at_in_group = require_at_in_group
        self._is_group_feature_enabled = is_group_feature_enabled
        self._check_permission = check_permission
        self._is_bot_mentioned = is_bot_mentioned
        self._prompt_context_service = (
            prompt_context_service
            or getattr(request_hook_service, "prompt_context_service", None)
            or PromptContextService(allow_external_input_files=False)
        )
        explicit_hooks_provided = (
            notice_hooks is not None or tool_exposure_hooks is not None
        )
        if explicit_hooks_provided and (
            notice_hooks is None or tool_exposure_hooks is None
        ):
            raise ValueError(
                "notice_hooks and tool_exposure_hooks must be provided together"
            )
        if not explicit_hooks_provided and request_hook_service is None:
            raise ValueError(
                "request_hook_service is required when hooks are not provided"
            )
        if explicit_hooks_provided:
            self._notice_hooks = notice_hooks
            self._tool_exposure_hooks = tool_exposure_hooks
            self._request_hook_service = None
        else:
            self._notice_hooks = None
            self._tool_exposure_hooks = None
            self._request_hook_service = request_hook_service

    def _detect_explicit_file_tool(self, text: str) -> str | None:
        if not text:
            return None

        explicit_matches: set[str] = set()
        tool_invocation_prefix = (
            r"(?:调用|使用|请求(?:调用|使用)?|请(?!问)(?:调用|使用)?|"
            r"\b(?:call|use|invoke)\b|\bplease\s+(?:call|use|invoke)\b)"
        )
        for tool_name in sorted(FILE_TOOLS, key=len, reverse=True):
            patterns = (
                rf"(?P<tool>{tool_invocation_prefix}\s*`?{re.escape(tool_name)}`?)",
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
        if self._request_hook_service is not None:
            return await self._request_hook_service.apply_tool_exposure_hooks(context)
        return await run_tool_exposure_hooks(self._tool_exposure_hooks, context)

    async def _run_before_build_notices(
        self, context: NoticeBuildContext
    ) -> NoticeBuildContext:
        if self._request_hook_service is not None:
            return await self._request_hook_service.apply_notice_hooks(context)
        return await run_notice_hooks(self._notice_hooks, context)

    @staticmethod
    def _set_event_extra(
        event: AstrMessageEvent,
        key: str,
        value,
    ) -> None:
        set_extra = getattr(event, "set_extra", None)
        if callable(set_extra):
            set_extra(key, value)

    @staticmethod
    def _remove_file_tools(req: ProviderRequest) -> None:
        if req.func_tool:
            for tool_name in FILE_TOOLS:
                req.func_tool.remove_tool(tool_name)

    def _append_tools_denied_notice(self, req: ProviderRequest) -> None:
        denied_section = self._prompt_context_service.build_tools_denied_section()
        ordered_names, ordered_notices = (
            self._prompt_context_service.order_notice_sections(
                section_names=[denied_section.name],
                notices=[denied_section.content],
            )
        )
        req.system_prompt = (req.system_prompt or "") + "".join(ordered_notices)
        logger.debug(
            "[文件管理] Prompt sections: %s",
            self._prompt_context_service.build_section_trace(
                section_names=ordered_names,
                notices=ordered_notices,
            ),
        )

    def _resolve_exposure_decision(
        self,
        event: AstrMessageEvent,
        req: ProviderRequest,
        *,
        is_group: bool,
        is_friend: bool,
    ) -> ExposureDecision:
        explicit_tool_name = self._detect_explicit_file_tool(
            self._extract_explicit_tool_text(event, req)
        )
        has_permission = self._check_permission(event)
        can_process_upload = has_permission or event.is_admin()
        get_extra = getattr(event, "get_extra", None)
        doc_command_triggered = bool(
            get_extra(DOC_COMMAND_TRIGGER_EVENT_KEY, False)
            if callable(get_extra)
            else False
        )
        meets_group_trigger = (
            not is_group
            or not self._require_at_in_group
            or self._is_bot_mentioned(event)
            or doc_command_triggered
        )
        should_expose = (is_friend and event.is_admin()) or (
            has_permission and meets_group_trigger
        )
        return self.ExposureDecision(
            explicit_tool_name=explicit_tool_name,
            can_process_upload=can_process_upload,
            should_expose=should_expose,
        )

    async def apply(self, event: AstrMessageEvent, req: ProviderRequest) -> None:
        is_group = event.message_obj.type == MessageType.GROUP_MESSAGE
        is_friend = event.message_obj.type == MessageType.FRIEND_MESSAGE
        decision = self._resolve_exposure_decision(
            event,
            req,
            is_group=is_group,
            is_friend=is_friend,
        )
        self._set_event_extra(
            event,
            EXPLICIT_FILE_TOOL_EVENT_KEY,
            decision.explicit_tool_name,
        )

        if not self._is_group_feature_enabled(event):
            self._remove_file_tools(req)
            self._append_tools_denied_notice(req)
            logger.debug("[文件管理] 群聊总开关关闭，已隐藏全部文件工具")
            return

        if not decision.should_expose:
            if not self._check_permission(event):
                logger.debug(
                    f"[文件管理] 用户 {event.get_sender_id()} 无文件权限，已隐藏文件工具"
                )
            else:
                logger.debug(
                    f"[文件管理] 用户 {event.get_sender_id()} 未满足群聊触发条件，已隐藏文件工具"
                )
            self._remove_file_tools(req)
            self._append_tools_denied_notice(req)

        if decision.should_expose and req.func_tool:
            for tool in self._document_toolset.tools:
                req.func_tool.add_tool(tool)

        await self._run_before_expose_tools(
            ToolExposureContext(
                event=event,
                request=req,
                should_expose=decision.should_expose,
                can_process_upload=decision.can_process_upload,
                explicit_tool_name=decision.explicit_tool_name,
            )
        )

        notice_context = await self._run_before_build_notices(
            NoticeBuildContext(
                event=event,
                request=req,
                should_expose=decision.should_expose,
                can_process_upload=decision.can_process_upload,
                explicit_tool_name=decision.explicit_tool_name,
            )
        )
        if notice_context.notices:
            ordered_names, ordered_notices = (
                self._prompt_context_service.order_notice_sections(
                    section_names=notice_context.section_names,
                    notices=notice_context.notices,
                )
            )
            req.system_prompt = (req.system_prompt or "") + "".join(ordered_notices)
            logger.debug(
                "[文件管理] Prompt sections: %s",
                self._prompt_context_service.build_section_trace(
                    section_names=ordered_names,
                    notices=ordered_notices,
                ),
            )
