import re
from collections.abc import Callable
from dataclasses import dataclass

import astrbot.api.message_components as Comp
from astrbot.api import logger
from astrbot.api.event import AstrMessageEvent
from astrbot.core.platform.message_type import MessageType
from astrbot.core.provider.entities import ProviderRequest

from ..constants import (
    DOC_COMMAND_TRIGGER_EVENT_KEY,
    DOCUMENT_FULL_TOOLS,
    EXECUTION_TOOLS,
    EXPLICIT_FILE_TOOL_EVENT_KEY,
    ExposureLevel,
    FILE_TOOLS,
    FILE_ONLY_TOOLS,
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
    _DOCUMENT_ID_RE = re.compile(
        r'document_id["`\']?\s*[:=]\s*["`\']?(?P<document_id>[A-Za-z0-9_-]+)',
        flags=re.IGNORECASE,
    )
    _DOCUMENT_INTENT_RE = re.compile(
        r"(create_document|add_blocks|finalize_document|export_document|"
        r"document_id\b|整理成\s*(?:文档|报告|汇报|word|docx)|"
        r"导出成\s*word|导出为\s*word|生成\s*(?:文档|报告|汇报|word|docx)|"
        r"\bword\b|\bdocx\b|文档|报告|汇报)",
        flags=re.IGNORECASE,
    )
    _FILE_INTENT_RE = re.compile(
        r"(read_file|convert_to_pdf|convert_from_pdf|"
        r"读取.*文件|查看.*文件|看看.*文件|读取内容|读取这个|"
        r"\bpdf\b|转成\s*pdf|转换成\s*pdf|导出成\s*pdf|导出为\s*pdf|"
        r"pdf转word|pdf转excel)",
        flags=re.IGNORECASE,
    )
    _DOCUMENT_FOLLOWUP_RE = re.compile(
        r"(继续|接着|再加|加一章|加一节|补充|完善|导出|发给我)",
        flags=re.IGNORECASE,
    )
    _DOCUMENT_FOLLOWUP_ACTION_RE = re.compile(
        r"(导出|发给我|加一章|加一节|补充|再加)",
        flags=re.IGNORECASE,
    )
    _BUFFERED_USER_INSTRUCTION_RE = re.compile(
        r"\[用户指令\]\s*(?P<instruction>.*?)(?:\n\s*\[|\Z)",
        flags=re.DOTALL,
    )

    @dataclass(slots=True)
    class ExposureDecision:
        explicit_tool_name: str | None
        has_permission: bool
        can_process_upload: bool
        should_expose: bool
        exposure_level: ExposureLevel
        allowed_tool_names: tuple[str, ...]
        active_document_summary: dict[str, object] | None
        denied_reason: str | None

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

    @staticmethod
    def _remove_execution_tools(req: ProviderRequest) -> None:
        if req.func_tool:
            for tool_name in EXECUTION_TOOLS:
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

    @staticmethod
    def _extract_document_id(text: str) -> str | None:
        if not text:
            return None
        match = LLMRequestPolicy._DOCUMENT_ID_RE.search(text)
        if not match:
            return None
        return match.group("document_id")

    @staticmethod
    def _has_uploaded_file_component(event: AstrMessageEvent) -> bool:
        return any(
            isinstance(component, Comp.File)
            for component in (getattr(event.message_obj, "message", None) or [])
        )

    def _get_cached_upload_infos(self, event: AstrMessageEvent) -> list[dict[str, object]]:
        if self._request_hook_service is None:
            return []
        return self._request_hook_service.get_cached_upload_infos(event)

    def _get_active_document_summary(
        self, event: AstrMessageEvent
    ) -> dict[str, object] | None:
        if self._request_hook_service is None:
            return None
        return self._request_hook_service.get_active_document_prompt_summary(event)

    @classmethod
    def _looks_like_document_followup(cls, text: str) -> bool:
        normalized_text = str(text or "").strip()
        if not normalized_text:
            return False
        if not cls._DOCUMENT_FOLLOWUP_RE.search(normalized_text):
            return False
        if len(normalized_text) <= 24:
            return True
        return bool(cls._DOCUMENT_FOLLOWUP_ACTION_RE.search(normalized_text))

    @staticmethod
    def _allowed_tool_names_for_level(
        exposure_level: ExposureLevel,
    ) -> tuple[str, ...]:
        if exposure_level == ExposureLevel.FILE_ONLY:
            return tuple(FILE_ONLY_TOOLS)
        if exposure_level == ExposureLevel.DOCUMENT_FULL:
            return tuple(DOCUMENT_FULL_TOOLS)
        return ()

    def _apply_allowed_file_tools(
        self,
        req: ProviderRequest,
        *,
        exposure_level: ExposureLevel,
    ) -> tuple[str, ...]:
        if not req.func_tool:
            return ()

        allowed_tool_names = self._allowed_tool_names_for_level(exposure_level)
        if exposure_level == ExposureLevel.DOCUMENT_FULL:
            for tool in self._document_toolset.tools:
                if tool.name in allowed_tool_names:
                    req.func_tool.add_tool(tool)

        allowed_tool_set = set(allowed_tool_names)
        for tool_name in FILE_TOOLS:
            if tool_name not in allowed_tool_set:
                req.func_tool.remove_tool(tool_name)
        return allowed_tool_names

    def _resolve_exposure_decision(
        self,
        event: AstrMessageEvent,
        req: ProviderRequest,
        *,
        is_group: bool,
        is_friend: bool,
    ) -> ExposureDecision:
        request_text = self._extract_explicit_tool_text(event, req)
        explicit_tool_name = self._detect_explicit_file_tool(request_text)
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
        if not self._is_group_feature_enabled(event):
            return self.ExposureDecision(
                explicit_tool_name=explicit_tool_name,
                has_permission=has_permission,
                can_process_upload=can_process_upload,
                should_expose=False,
                exposure_level=ExposureLevel.NONE,
                allowed_tool_names=(),
                active_document_summary=None,
                denied_reason="group_feature_disabled",
            )
        if not ((is_friend and event.is_admin()) or has_permission):
            return self.ExposureDecision(
                explicit_tool_name=explicit_tool_name,
                has_permission=has_permission,
                can_process_upload=can_process_upload,
                should_expose=False,
                exposure_level=ExposureLevel.NONE,
                allowed_tool_names=(),
                active_document_summary=None,
                denied_reason="missing_permission",
            )
        if not ((is_friend and event.is_admin()) or meets_group_trigger):
            return self.ExposureDecision(
                explicit_tool_name=explicit_tool_name,
                has_permission=has_permission,
                can_process_upload=can_process_upload,
                should_expose=False,
                exposure_level=ExposureLevel.NONE,
                allowed_tool_names=(),
                active_document_summary=None,
                denied_reason="missing_group_trigger",
            )

        has_uploaded_files = self._has_uploaded_file_component(event)
        has_cached_uploads = bool(self._get_cached_upload_infos(event))
        active_document_summary = self._get_active_document_summary(event)
        has_active_document = bool(active_document_summary)
        has_document_id = bool(self._extract_document_id(request_text))
        has_document_intent = bool(
            has_document_id or self._DOCUMENT_INTENT_RE.search(request_text)
        )
        has_file_intent = bool(self._FILE_INTENT_RE.search(request_text))

        if has_document_id:
            exposure_level = ExposureLevel.DOCUMENT_FULL
        elif has_active_document and self._looks_like_document_followup(request_text):
            exposure_level = ExposureLevel.DOCUMENT_FULL
        elif (has_uploaded_files or has_cached_uploads) and has_document_intent:
            exposure_level = ExposureLevel.DOCUMENT_FULL
        elif has_document_intent:
            exposure_level = ExposureLevel.DOCUMENT_FULL
        elif has_uploaded_files or has_cached_uploads or has_file_intent:
            exposure_level = ExposureLevel.FILE_ONLY
        else:
            exposure_level = ExposureLevel.NONE

        allowed_tool_names = self._allowed_tool_names_for_level(exposure_level)
        should_expose = exposure_level != ExposureLevel.NONE
        return self.ExposureDecision(
            explicit_tool_name=explicit_tool_name,
            has_permission=has_permission,
            can_process_upload=can_process_upload,
            should_expose=should_expose,
            exposure_level=exposure_level,
            allowed_tool_names=allowed_tool_names,
            active_document_summary=active_document_summary,
            denied_reason=None if should_expose else "no_relevant_intent",
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

        if not decision.should_expose:
            if decision.denied_reason == "group_feature_disabled":
                logger.debug("[文件管理] 群聊总开关关闭，已隐藏全部文件工具")
            elif decision.denied_reason == "missing_permission":
                logger.debug(
                    f"[文件管理] 用户 {event.get_sender_id()} 无文件权限，已隐藏文件工具"
                )
            elif decision.denied_reason == "no_relevant_intent":
                logger.debug("[文件管理] 当前请求无文件相关意图，已隐藏文件工具")
            else:
                logger.debug(
                    f"[文件管理] 用户 {event.get_sender_id()} 未满足群聊触发条件，已隐藏文件工具"
                )
            self._remove_file_tools(req)
            self._remove_execution_tools(req)
            self._append_tools_denied_notice(req)
            return

        allowed_tool_names = self._apply_allowed_file_tools(
            req,
            exposure_level=decision.exposure_level,
        )
        logger.debug(
            "[文件管理] exposure_level=%s | allowed_tools=%s",
            decision.exposure_level.value,
            ",".join(allowed_tool_names) or "none",
        )

        await self._run_before_expose_tools(
            ToolExposureContext(
                event=event,
                request=req,
                should_expose=decision.should_expose,
                can_process_upload=decision.can_process_upload,
                explicit_tool_name=decision.explicit_tool_name,
                exposure_level=decision.exposure_level,
                allowed_tool_names=allowed_tool_names,
            )
        )

        notice_context = await self._run_before_build_notices(
            NoticeBuildContext(
                event=event,
                request=req,
                should_expose=decision.should_expose,
                can_process_upload=decision.can_process_upload,
                explicit_tool_name=decision.explicit_tool_name,
                exposure_level=decision.exposure_level,
                allowed_tool_names=allowed_tool_names,
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
                (
                    self._prompt_context_service.build_section_trace(
                        section_names=ordered_names,
                        notices=ordered_notices,
                    )
                    + f" | exposure_level={decision.exposure_level.value}"
                    + f" | allowed_tools={','.join(allowed_tool_names) or 'none'}"
                ),
            )
