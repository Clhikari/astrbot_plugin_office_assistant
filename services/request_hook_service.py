import re
from collections.abc import Awaitable, Callable
from pathlib import Path
from typing import Any

import astrbot.api.message_components as Comp
from astrbot.api import logger
from astrbot.api.event import AstrMessageEvent

from ..constants import (
    ALL_OFFICE_SUFFIXES,
    EXECUTION_TOOLS,
    FILE_TOOLS,
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
from .prompt_context_service import PromptContextService, PromptSection
from .excel_intent_router import ExcelIntentRouter
from .request_follow_up import (
    FollowUpNoticeStrategy,
    IdentifierDescriptor,
    build_identifier_descriptor,
    extract_identifier_from_text,
)
from .request_hook_notice_helpers import (
    FollowUpNoticeHelper,
    FollowUpNoticeRule,
)
from .upload_types import UploadInfo
from .runtime_config import (
    SUPPORTED_COMPUTER_RUNTIME_MODES,
    resolve_computer_runtime_mode,
)


def _normalize_string_list(value: object) -> list[str]:
    if not isinstance(value, list):
        return []
    return [
        normalized_value
        for item in value
        if item is not None and (normalized_value := str(item).strip())
    ]


class RequestHookService:
    _WORKBOOK_TOOL_NAMES = frozenset(
        {"create_workbook", "write_rows", "export_workbook"}
    )
    _DOCUMENT_ID_HEX_RE = re.compile(r"[0-9a-f]{32}", flags=re.IGNORECASE)
    _DOCUMENT_ID_DOC_PREFIX_RE = re.compile(
        r"doc-[A-Za-z0-9_-]+",
        flags=re.IGNORECASE,
    )
    _BUFFERED_USER_INSTRUCTION_RE = re.compile(
        r"\[用户指令\]\s*(?P<instruction>.*?)(?:\n\s*\[|\Z)",
        flags=re.DOTALL,
    )
    _DOCUMENT_WORKFLOW_HINT_RE = re.compile(
        r"(create_document|add_blocks|finalize_document|export_document|"
        r"正式汇报|正式报告|导出成\s*word|导出为\s*word|"
        r"\bword\b|\bdocx\b|汇报|报告|"
        r"生成\s*(?:word|docx|报告|汇报)|整理成\s*(?:word|docx|报告|汇报))",
        flags=re.IGNORECASE,
    )
    _DOCUMENT_DETAIL_HINT_RE = re.compile(
        r"(business_report|project_review|executive_brief|accent_color|document_style)",
        flags=re.IGNORECASE,
    )
    _WORKBOOK_ID_HEX_RE = re.compile(r"[0-9a-f]{32}", flags=re.IGNORECASE)
    _WORKBOOK_ID_PREFIX_RE = re.compile(
        r"(?:wb|workbook)-[A-Za-z0-9_-]+",
        flags=re.IGNORECASE,
    )
    _DOCUMENT_IDENTIFIER = build_identifier_descriptor(
        token_name="document_id",
        bare_validator_name="_looks_like_bare_document_id",
    )
    _WORKBOOK_IDENTIFIER = build_identifier_descriptor(
        token_name="workbook_id",
        bare_validator_name="_looks_like_bare_workbook_id",
    )
    _DOCUMENT_FOLLOW_UP_STRATEGY = FollowUpNoticeStrategy(
        identifier=_DOCUMENT_IDENTIFIER,
        lookup_attr_name="_lookup_document_summary",
        section_builder_name="build_document_follow_up_section",
        missing_section_builder_name="build_document_follow_up_missing_section",
        payload_builder_name="_build_document_follow_up_payload",
        log_label="文档",
    )
    _WORKBOOK_FOLLOW_UP_STRATEGY = FollowUpNoticeStrategy(
        identifier=_WORKBOOK_IDENTIFIER,
        lookup_attr_name="_lookup_workbook_summary",
        section_builder_name="build_workbook_follow_up_section",
        missing_section_builder_name="build_workbook_follow_up_missing_section",
        payload_builder_name="_build_workbook_follow_up_payload",
        log_label="工作簿",
    )
    _DOCUMENT_CORE_NOTICE_KEY = "document_core_guide"
    _DOCUMENT_DETAIL_NOTICE_KEY = "document_detail_guide"
    _EXCEL_ROUTING_NOTICE_KEY = "excel_routing_guide"
    _EXCEL_READ_NOTICE_KEY = "excel_read_guide"
    _EXCEL_SCRIPT_NOTICE_KEY = "excel_script_guide"
    _WORKBOOK_CORE_NOTICE_KEY = "workbook_core_guide"

    def __init__(
        self,
        *,
        astrbot_context=None,
        auto_block_execution_tools: bool,
        get_cached_upload_infos: Callable[[AstrMessageEvent], list[UploadInfo]],
        extract_upload_source: Callable[
            [Comp.File], Awaitable[tuple[Path | None, str]]
        ],
        store_uploaded_file: Callable[[Path, str], Path],
        consume_session_notice_once: Callable[[AstrMessageEvent, str], bool],
        allow_external_input_files: bool,
        prompt_context_service: PromptContextService | None = None,
        lookup_document_summary: Callable[[str], dict[str, object] | None] | None = None,
        lookup_workbook_summary: Callable[[str], dict[str, object] | None] | None = None,
    ) -> None:
        self._astrbot_context = astrbot_context
        self._auto_block_execution_tools = auto_block_execution_tools
        self._get_cached_upload_infos = get_cached_upload_infos
        self._extract_upload_source = extract_upload_source
        self._store_uploaded_file = store_uploaded_file
        self._consume_session_notice_once = consume_session_notice_once
        self._lookup_document_summary = lookup_document_summary
        self._lookup_workbook_summary = lookup_workbook_summary
        self.prompt_context_service = prompt_context_service or PromptContextService(
            allow_external_input_files=allow_external_input_files
        )
        self._follow_up_notice_helper = FollowUpNoticeHelper(
            is_follow_up_request=self._is_follow_up_request,
            build_follow_up_section=self._build_follow_up_section,
            append_notice_section=self._append_notice_section,
        )
        self._notice_hooks = [
            self.append_office_tool_guide_notice,
            self.append_uploaded_file_notices,
        ]
        self._tool_exposure_hooks = [
            self.apply_execution_tool_block,
            self.apply_excel_script_runtime_restriction,
            self.apply_explicit_file_tool_restriction,
        ]

    def build_notice_hooks(self) -> list[NoticeBuildHook]:
        return list(self._notice_hooks)

    def build_tool_exposure_hooks(self) -> list[ToolExposureHook]:
        return list(self._tool_exposure_hooks)

    async def apply_tool_exposure_hooks(
        self,
        context: ToolExposureContext,
    ) -> ToolExposureContext:
        return await run_tool_exposure_hooks(self._tool_exposure_hooks, context)

    async def apply_notice_hooks(
        self,
        context: NoticeBuildContext,
    ) -> NoticeBuildContext:
        return await run_notice_hooks(self._notice_hooks, context)

    async def append_office_tool_guide_notice(
        self,
        context: NoticeBuildContext,
    ) -> NoticeBuildContext:
        if not (context.should_expose and context.request.func_tool):
            return context

        request_text = self._extract_prompt_text(str(context.request.prompt or ""))
        exposed_tool_names = self._get_exposed_tool_names(context.request.func_tool)
        if self._append_follow_up_notice_if_needed(
            context,
            request_text=request_text,
            exposed_tool_names=exposed_tool_names,
        ):
            return context

        self._append_document_tool_guides(
            context,
            request_text=request_text,
        )
        self._append_excel_tool_guides(
            context,
            request_text=request_text,
            exposed_tool_names=exposed_tool_names,
        )
        return context

    async def append_document_tool_guide_notice(
        self,
        context: NoticeBuildContext,
    ) -> NoticeBuildContext:
        return await self.append_office_tool_guide_notice(context)

    @classmethod
    def _extract_prompt_text(cls, prompt: str) -> str:
        stripped_prompt = prompt.strip()
        if not stripped_prompt:
            return ""
        match = cls._BUFFERED_USER_INSTRUCTION_RE.search(stripped_prompt)
        if match:
            return match.group("instruction").strip()
        return stripped_prompt

    @classmethod
    def _get_exposed_tool_names(cls, func_tool) -> set[str]:
        names = getattr(func_tool, "names", None)
        if not callable(names):
            return set()
        return {str(name) for name in names() if name}

    def _resolve_excel_runtime_mode(self, event: AstrMessageEvent) -> str:
        return resolve_computer_runtime_mode(self._astrbot_context, event)

    @classmethod
    def _has_workbook_tools_available(cls, *, exposed_tool_names: set[str]) -> bool:
        return bool(cls._WORKBOOK_TOOL_NAMES.intersection(exposed_tool_names))

    @classmethod
    def _has_full_workbook_toolset_available(
        cls,
        *,
        exposed_tool_names: set[str],
    ) -> bool:
        return cls._WORKBOOK_TOOL_NAMES.issubset(exposed_tool_names)

    def _append_follow_up_notice_if_needed(
        self,
        context: NoticeBuildContext,
        *,
        request_text: str,
        exposed_tool_names: set[str],
    ) -> bool:
        workbook_availability_checker = (
            self._has_workbook_tools_available
            if context.explicit_tool_name is None
            else self._has_full_workbook_toolset_available
        )
        follow_up_rules = (
            FollowUpNoticeRule(
                strategy=self._WORKBOOK_FOLLOW_UP_STRATEGY,
                availability_checker=lambda names: workbook_availability_checker(
                    exposed_tool_names=names
                ),
            ),
            FollowUpNoticeRule(strategy=self._DOCUMENT_FOLLOW_UP_STRATEGY),
        )
        return self._follow_up_notice_helper.append_first_matching_notice(
            context,
            request_text=request_text,
            exposed_tool_names=exposed_tool_names,
            rules=follow_up_rules,
        )

    def _append_document_tool_guides(
        self,
        context: NoticeBuildContext,
        *,
        request_text: str,
    ) -> None:
        should_inject_core = self._should_inject_document_tool_guide(
            request_text=request_text
        )
        should_inject_detail = self._should_inject_document_tool_detail(
            request_text=request_text
        )

        if should_inject_core and self._consume_session_notice_once(
            context.event, self._DOCUMENT_CORE_NOTICE_KEY
        ):
            self._append_notice_section(
                context,
                self.prompt_context_service.build_document_tool_guide_section(),
            )
        if should_inject_detail and self._consume_session_notice_once(
            context.event, self._DOCUMENT_DETAIL_NOTICE_KEY
        ):
            self._append_notice_section(
                context,
                self.prompt_context_service.build_document_tool_detail_section(),
            )

    def _append_excel_tool_guides(
        self,
        context: NoticeBuildContext,
        *,
        request_text: str,
        exposed_tool_names: set[str],
    ) -> None:
        excel_route = ExcelIntentRouter.decide(
            request_text=request_text,
            upload_infos=self._get_cached_upload_infos(context.event),
            explicit_tool_name=context.explicit_tool_name,
            exposed_tool_names=exposed_tool_names,
        )
        if excel_route is None:
            return

        if not excel_route.should_inject_guide:
            return

        if self._consume_session_notice_once(
            context.event, self._EXCEL_ROUTING_NOTICE_KEY
        ):
            self._append_notice_section(
                context,
                self.prompt_context_service.build_excel_routing_section(),
            )

        read_tool_available = (
            "read_workbook" in exposed_tool_names
            and context.explicit_tool_name in (None, "read_workbook")
        )
        script_tool_available = (
            "execute_excel_script" in exposed_tool_names
            and context.explicit_tool_name in (None, "execute_excel_script")
        )
        workbook_tools_available = (
            self._has_full_workbook_toolset_available(
                exposed_tool_names=exposed_tool_names
            )
            and (
                context.explicit_tool_name is None
                or context.explicit_tool_name in self._WORKBOOK_TOOL_NAMES
            )
        )

        if read_tool_available:
            if self._consume_session_notice_once(
                context.event, self._EXCEL_READ_NOTICE_KEY
            ):
                self._append_notice_section(
                    context,
                    self.prompt_context_service.build_excel_read_section(),
                )

        if script_tool_available:
            if self._consume_session_notice_once(
                context.event, self._EXCEL_SCRIPT_NOTICE_KEY
            ):
                self._append_notice_section(
                    context,
                    self.prompt_context_service.build_excel_script_section(),
                )
        if workbook_tools_available and self._consume_session_notice_once(
            context.event, self._WORKBOOK_CORE_NOTICE_KEY
        ):
            self._append_notice_section(
                context,
                self.prompt_context_service.build_workbook_tool_guide_section(),
            )

    @classmethod
    def _should_inject_document_tool_guide(cls, *, request_text: str) -> bool:
        if not request_text:
            return False
        return bool(cls._DOCUMENT_WORKFLOW_HINT_RE.search(request_text))

    @classmethod
    def _should_inject_document_tool_detail(
        cls,
        *,
        request_text: str,
    ) -> bool:
        if not request_text:
            return False
        return bool(cls._DOCUMENT_DETAIL_HINT_RE.search(request_text))

    @classmethod
    def _is_document_follow_up(cls, *, request_text: str) -> bool:
        return cls._is_follow_up_request(
            request_text=request_text,
            descriptor=cls._DOCUMENT_IDENTIFIER,
        )

    @classmethod
    def _is_workbook_follow_up(cls, *, request_text: str) -> bool:
        if not request_text:
            return False
        return cls._is_follow_up_request(
            request_text=request_text,
            descriptor=cls._WORKBOOK_IDENTIFIER,
        )

    @classmethod
    def _is_follow_up_request(
        cls,
        *,
        request_text: str,
        descriptor: IdentifierDescriptor,
    ) -> bool:
        if not request_text:
            return False
        return bool(descriptor.follow_up_re.search(request_text))

    @classmethod
    def _extract_document_id(cls, *, request_text: str) -> str:
        return cls._extract_identifier(
            request_text=request_text,
            descriptor=cls._DOCUMENT_IDENTIFIER,
        )

    @classmethod
    def _extract_workbook_id(cls, *, request_text: str) -> str:
        return cls._extract_identifier(
            request_text=request_text,
            descriptor=cls._WORKBOOK_IDENTIFIER,
        )

    @classmethod
    def _extract_identifier(
        cls,
        *,
        request_text: str,
        descriptor: IdentifierDescriptor,
    ) -> str:
        return extract_identifier_from_text(
            request_text=request_text,
            explicit_capture_re=descriptor.explicit_capture_re,
            bare_capture_re=descriptor.bare_capture_re,
            group_name=descriptor.group_name,
            is_valid_bare_id=getattr(cls, descriptor.bare_validator_name),
        )

    @classmethod
    def _looks_like_bare_document_id(cls, candidate: str) -> bool:
        if not candidate:
            return False
        return bool(
            cls._DOCUMENT_ID_HEX_RE.fullmatch(candidate)
            or cls._DOCUMENT_ID_DOC_PREFIX_RE.fullmatch(candidate)
        )

    @classmethod
    def _looks_like_bare_workbook_id(cls, candidate: str) -> bool:
        if not candidate:
            return False
        return bool(
            cls._WORKBOOK_ID_HEX_RE.fullmatch(candidate)
            or cls._WORKBOOK_ID_PREFIX_RE.fullmatch(candidate)
        )

    def _build_document_follow_up_section(
        self,
        *,
        request_text: str,
    ) -> PromptSection | None:
        return self._build_follow_up_section(
            request_text=request_text,
            strategy=self._DOCUMENT_FOLLOW_UP_STRATEGY,
        )

    def _build_workbook_follow_up_section(
        self,
        *,
        request_text: str,
    ) -> PromptSection | None:
        return self._build_follow_up_section(
            request_text=request_text,
            strategy=self._WORKBOOK_FOLLOW_UP_STRATEGY,
        )

    def _build_follow_up_section(
        self,
        *,
        request_text: str,
        strategy: FollowUpNoticeStrategy,
    ) -> PromptSection | None:
        identifier_value = self._extract_identifier(
            request_text=request_text,
            descriptor=strategy.identifier,
        )
        if not identifier_value:
            return None

        lookup_summary = getattr(self, strategy.lookup_attr_name)
        if lookup_summary is None:
            return None

        try:
            summary = lookup_summary(identifier_value)
        except KeyError:
            summary = None
        except Exception as exc:
            logger.exception(
                "[文件管理] 查询%s会话摘要失败 %s=%s: %s"
                % (
                    strategy.log_label,
                    strategy.identifier.token_name,
                    identifier_value,
                    exc,
                )
            )
            raise

        if not isinstance(summary, dict):
            missing_builder = getattr(
                self.prompt_context_service,
                strategy.missing_section_builder_name,
            )
            return missing_builder(**{strategy.identifier.token_name: identifier_value})

        section_builder = getattr(
            self.prompt_context_service,
            strategy.section_builder_name,
        )
        payload_builder = getattr(self, strategy.payload_builder_name)
        return section_builder(
            **payload_builder(identifier_value=identifier_value, summary=summary)
        )

    @staticmethod
    def _build_document_follow_up_payload(
        *,
        identifier_value: str,
        summary: dict[str, object],
    ) -> dict[str, Any]:
        return {
            "document_id": identifier_value,
            "status": str(summary.get("status") or ""),
            "block_count": int(summary.get("block_count") or 0),
        }

    @staticmethod
    def _build_workbook_follow_up_payload(
        *,
        identifier_value: str,
        summary: dict[str, object],
    ) -> dict[str, Any]:
        return {
            "workbook_id": identifier_value,
            "status": str(summary.get("status") or ""),
            "sheet_names": _normalize_string_list(summary.get("sheet_names")),
            "sheet_count": int(summary.get("sheet_count") or 0),
            "latest_written_sheets": _normalize_string_list(
                summary.get("latest_written_sheets")
            ),
            "next_allowed_actions": _normalize_string_list(
                summary.get("next_allowed_actions")
            ),
        }

    @staticmethod
    def _append_notice_section(
        context: NoticeBuildContext,
        section: PromptSection | None,
    ) -> None:
        if not section or not section.content:
            return
        if section.target == "system":
            notices = getattr(context, "system_notices", None)
            section_names = getattr(context, "system_section_names", None)
        else:
            notices = getattr(context, "notices", None)
            section_names = getattr(context, "section_names", None)
        if isinstance(notices, list):
            notices.append(section.content)
        if isinstance(section_names, list):
            section_names.append(section.name)

    async def append_uploaded_file_notices(
        self,
        context: NoticeBuildContext,
    ) -> NoticeBuildContext:
        if not context.can_process_upload:
            return context
        if getattr(context.event, "_buffered", False) is True:
            return context

        event = context.event
        req = context.request
        cached_upload_infos = iter(self._get_cached_upload_infos(event))
        readable_upload_infos: list[UploadInfo] = []
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

                readable_upload_infos.append(
                    {
                        "original_name": original_name,
                        "file_suffix": file_suffix,
                        "type_desc": type_desc,
                        "is_supported": is_supported,
                        "stored_name": stored_name,
                        "source_path": source_path_text,
                    }
                )
                logger.info(
                    "[文件管理] 收到文件 %s，已保存为 %s。",
                    original_name,
                    stored_name,
                )
            except Exception as exc:
                logger.error(f"[文件管理] 处理上传文件失败: {exc}")

        if not readable_upload_infos:
            return context

        self._append_notice_section(
            context,
            self.prompt_context_service.build_uploaded_file_context_section(
                upload_infos=readable_upload_infos
            ),
        )
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

    async def apply_excel_script_runtime_restriction(
        self,
        context: ToolExposureContext,
    ) -> ToolExposureContext:
        if not (context.should_expose and context.request.func_tool):
            return context
        try:
            runtime_mode = self._resolve_excel_runtime_mode(context.event)
        except Exception as exc:
            logger.warning(
                f"[文件管理] 读取 Excel runtime 配置失败，跳过 execute_excel_script 显隐控制: {exc}"
            )
            return context
        if self._auto_block_execution_tools:
            if runtime_mode == "sandbox":
                return context
            context.request.func_tool.remove_tool("execute_excel_script")
            if runtime_mode == "local":
                logger.info(
                    "[文件管理] 已启用执行类工具自动屏蔽，当前 computer runtime 为 local，已隐藏 execute_excel_script"
                )
                return context
        elif runtime_mode in {"local", "sandbox"}:
            return context
        context.request.func_tool.remove_tool("execute_excel_script")
        if runtime_mode == "none":
            logger.info("[文件管理] 当前 computer runtime 为 none，已隐藏 execute_excel_script")
            return context
        if runtime_mode not in SUPPORTED_COMPUTER_RUNTIME_MODES:
            logger.warning(
                f"[文件管理] 当前 computer runtime 配置不受支持：{runtime_mode}，已隐藏 execute_excel_script"
            )
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
            exposed_file_tool_names = self._get_exposed_tool_names(
                context.request.func_tool
            ).intersection(FILE_TOOLS)
            if context.explicit_tool_name not in exposed_file_tool_names:
                logger.info(
                    "[文件管理] 检测到用户显式指定工具 %s，但该工具当前未暴露，跳过工具收窄",
                    context.explicit_tool_name,
                )
                return context
            for tool_name in exposed_file_tool_names:
                if tool_name != context.explicit_tool_name:
                    context.request.func_tool.remove_tool(tool_name)
            logger.info(
                "[文件管理] 检测到用户显式指定工具 %s，本轮仅保留该文件工具",
                context.explicit_tool_name,
            )
        return context
