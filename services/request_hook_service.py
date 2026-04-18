import re
from collections.abc import Awaitable, Callable
from dataclasses import dataclass
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
from .upload_types import UploadInfo


@dataclass(frozen=True, slots=True)
class _IdentifierDescriptor:
    token_name: str
    group_name: str
    follow_up_re: re.Pattern[str]
    explicit_capture_re: re.Pattern[str]
    bare_capture_re: re.Pattern[str]
    bare_validator_name: str


@dataclass(frozen=True, slots=True)
class _FollowUpNoticeStrategy:
    identifier: _IdentifierDescriptor
    lookup_attr_name: str
    section_builder_name: str
    missing_section_builder_name: str
    payload_builder_name: str
    log_label: str


def _build_identifier_token_pattern(token_name: str) -> str:
    return rf"(?<![A-Za-z0-9_]){re.escape(token_name)}(?![A-Za-z0-9_])"


def _compile_identifier_token_regex(token_name: str) -> re.Pattern[str]:
    return re.compile(_build_identifier_token_pattern(token_name), flags=re.IGNORECASE)


def _compile_identifier_explicit_capture_regex(
    token_name: str,
    group_name: str,
) -> re.Pattern[str]:
    return re.compile(
        _build_identifier_token_pattern(token_name)
        + rf"(?:\s*[:=：]\s*|\s*(?:为|是)\s*|\s+is\s+)[`\"']?(?P<{group_name}>[A-Za-z0-9_-]+)[`\"']?",
        flags=re.IGNORECASE,
    )


def _compile_identifier_bare_capture_regex(
    token_name: str,
    group_name: str,
) -> re.Pattern[str]:
    return re.compile(
        _build_identifier_token_pattern(token_name)
        + rf"\s+[`\"']?(?P<{group_name}>[A-Za-z0-9_-]*[\d_-][A-Za-z0-9_-]*)[`\"']?",
        flags=re.IGNORECASE,
    )


def _extract_identifier_from_text(
    *,
    request_text: str,
    explicit_capture_re: re.Pattern[str],
    bare_capture_re: re.Pattern[str],
    group_name: str,
    is_valid_bare_id: Callable[[str], bool],
) -> str:
    if not request_text:
        return ""

    explicit_match = explicit_capture_re.search(request_text)
    if explicit_match:
        return str(explicit_match.group(group_name) or "").strip()

    bare_match = bare_capture_re.search(request_text)
    if not bare_match:
        return ""

    candidate = str(bare_match.group(group_name) or "").strip()
    if is_valid_bare_id(candidate):
        return candidate
    return ""


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
    _DOCUMENT_ID_TOKEN_RE = _build_identifier_token_pattern("document_id")
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
    _DOCUMENT_ID_FOLLOW_UP_RE = _compile_identifier_token_regex("document_id")
    _DOCUMENT_ID_EXPLICIT_CAPTURE_RE = _compile_identifier_explicit_capture_regex(
        "document_id",
        "document_id",
    )
    _DOCUMENT_ID_BARE_CAPTURE_RE = _compile_identifier_bare_capture_regex(
        "document_id",
        "document_id",
    )
    _WORKBOOK_ID_TOKEN_RE = _build_identifier_token_pattern("workbook_id")
    _WORKBOOK_ID_HEX_RE = re.compile(r"[0-9a-f]{32}", flags=re.IGNORECASE)
    _WORKBOOK_ID_PREFIX_RE = re.compile(
        r"(?:wb|workbook)-[A-Za-z0-9_-]+",
        flags=re.IGNORECASE,
    )
    _WORKBOOK_TOOL_CALL_HINT_RE = re.compile(
        r"(create_workbook|write_rows|export_workbook)",
        flags=re.IGNORECASE,
    )
    _WORKBOOK_SUBJECT_HINT_RE = re.compile(
        r"(\bexcel\b|\bxlsx\b|报表|汇总表|工作簿|多\s*sheet)",
        flags=re.IGNORECASE,
    )
    _WORKBOOK_GENERATION_HINT_RE = re.compile(
        r"(生成|创建|新建|制作|整理成|整理为|写入|填入|输出|导出(?:成|为)?|返回|做(?:成|个|一份)?)",
        flags=re.IGNORECASE,
    )
    _WORKBOOK_NON_GENERATION_HINT_RE = re.compile(
        r"(读取|阅读|查看|打开|解析|提取|分析|总结|"
        r"导出(?:成|为)\s*pdf|"
        r"(?:转换|转成|转为|转到).*(?:pdf|word|docx|ppt|pptx)|"
        r"(?:pdf|word|docx|ppt|pptx).*(?:转换|转成|转为|转到)|"
        r"\bread\b|\bopen\b|\bparse\b|\bextract\b|\banaly[sz]e\b|\bconvert\b)",
        flags=re.IGNORECASE,
    )
    _WORKBOOK_DETAIL_HINT_RE = re.compile(
        r"(create_workbook|write_rows|export_workbook|start_row|多\s*sheet)",
        flags=re.IGNORECASE,
    )
    _WORKBOOK_ID_FOLLOW_UP_RE = _compile_identifier_token_regex("workbook_id")
    _WORKBOOK_ID_EXPLICIT_CAPTURE_RE = _compile_identifier_explicit_capture_regex(
        "workbook_id",
        "workbook_id",
    )
    _WORKBOOK_ID_BARE_CAPTURE_RE = _compile_identifier_bare_capture_regex(
        "workbook_id",
        "workbook_id",
    )
    _DOCUMENT_IDENTIFIER = _IdentifierDescriptor(
        token_name="document_id",
        group_name="document_id",
        follow_up_re=_DOCUMENT_ID_FOLLOW_UP_RE,
        explicit_capture_re=_DOCUMENT_ID_EXPLICIT_CAPTURE_RE,
        bare_capture_re=_DOCUMENT_ID_BARE_CAPTURE_RE,
        bare_validator_name="_looks_like_bare_document_id",
    )
    _WORKBOOK_IDENTIFIER = _IdentifierDescriptor(
        token_name="workbook_id",
        group_name="workbook_id",
        follow_up_re=_WORKBOOK_ID_FOLLOW_UP_RE,
        explicit_capture_re=_WORKBOOK_ID_EXPLICIT_CAPTURE_RE,
        bare_capture_re=_WORKBOOK_ID_BARE_CAPTURE_RE,
        bare_validator_name="_looks_like_bare_workbook_id",
    )
    _DOCUMENT_FOLLOW_UP_STRATEGY = _FollowUpNoticeStrategy(
        identifier=_DOCUMENT_IDENTIFIER,
        lookup_attr_name="_lookup_document_summary",
        section_builder_name="build_document_follow_up_section",
        missing_section_builder_name="build_document_follow_up_missing_section",
        payload_builder_name="_build_document_follow_up_payload",
        log_label="文档",
    )
    _WORKBOOK_FOLLOW_UP_STRATEGY = _FollowUpNoticeStrategy(
        identifier=_WORKBOOK_IDENTIFIER,
        lookup_attr_name="_lookup_workbook_summary",
        section_builder_name="build_workbook_follow_up_section",
        missing_section_builder_name="build_workbook_follow_up_missing_section",
        payload_builder_name="_build_workbook_follow_up_payload",
        log_label="工作簿",
    )
    _DOCUMENT_CORE_NOTICE_KEY = "document_core_guide"
    _DOCUMENT_DETAIL_NOTICE_KEY = "document_detail_guide"
    _WORKBOOK_CORE_NOTICE_KEY = "workbook_core_guide"
    _WORKBOOK_DETAIL_NOTICE_KEY = "workbook_detail_guide"

    def __init__(
        self,
        *,
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
        self._notice_hooks = [
            self.append_office_tool_guide_notice,
            self.append_uploaded_file_notices,
        ]
        self._tool_exposure_hooks = [
            self.apply_execution_tool_block,
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
        self._append_workbook_tool_guides(
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
            (
                self._WORKBOOK_FOLLOW_UP_STRATEGY,
                workbook_availability_checker,
            ),
            (
                self._DOCUMENT_FOLLOW_UP_STRATEGY,
                None,
            ),
        )
        for strategy, availability_checker in follow_up_rules:
            if self._append_follow_up_notice(
                context,
                request_text=request_text,
                exposed_tool_names=exposed_tool_names,
                strategy=strategy,
                availability_checker=availability_checker,
            ):
                return True
        return False

    def _append_follow_up_notice(
        self,
        context: NoticeBuildContext,
        *,
        request_text: str,
        exposed_tool_names: set[str],
        strategy: _FollowUpNoticeStrategy,
        availability_checker: Callable[..., bool] | None,
    ) -> bool:
        if availability_checker is not None and not availability_checker(
            exposed_tool_names=exposed_tool_names
        ):
            return False
        if not self._is_follow_up_request(
            request_text=request_text,
            descriptor=strategy.identifier,
        ):
            return False

        section = self._build_follow_up_section(
            request_text=request_text,
            strategy=strategy,
        )
        if section is None:
            return False

        self._append_notice_section(context, section)
        return True

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

    def _append_workbook_tool_guides(
        self,
        context: NoticeBuildContext,
        *,
        request_text: str,
        exposed_tool_names: set[str],
    ) -> None:
        if not self._has_full_workbook_toolset_available(
            exposed_tool_names=exposed_tool_names
        ):
            return

        should_inject_core = self._should_inject_workbook_tool_guide(
            request_text=request_text
        )
        should_inject_detail = self._should_inject_workbook_tool_detail(
            request_text=request_text
        )
        if not should_inject_core and not should_inject_detail:
            return

        if should_inject_core and self._consume_session_notice_once(
            context.event, self._WORKBOOK_CORE_NOTICE_KEY
        ):
            self._append_notice_section(
                context,
                self.prompt_context_service.build_workbook_tool_guide_section(),
            )
        if should_inject_detail and self._consume_session_notice_once(
            context.event, self._WORKBOOK_DETAIL_NOTICE_KEY
        ):
            self._append_notice_section(
                context,
                self.prompt_context_service.build_workbook_tool_detail_section(),
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
    def _should_inject_workbook_tool_guide(cls, *, request_text: str) -> bool:
        if not request_text:
            return False
        if cls._WORKBOOK_TOOL_CALL_HINT_RE.search(request_text):
            return True
        if cls._WORKBOOK_ID_FOLLOW_UP_RE.search(request_text):
            return True
        if not cls._WORKBOOK_SUBJECT_HINT_RE.search(request_text):
            return False
        if not cls._WORKBOOK_GENERATION_HINT_RE.search(request_text):
            return False
        return not cls._WORKBOOK_NON_GENERATION_HINT_RE.search(request_text)

    @classmethod
    def _should_inject_workbook_tool_detail(
        cls,
        *,
        request_text: str,
    ) -> bool:
        if not request_text:
            return False
        return bool(cls._WORKBOOK_DETAIL_HINT_RE.search(request_text))

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
        descriptor: _IdentifierDescriptor,
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
        descriptor: _IdentifierDescriptor,
    ) -> str:
        return _extract_identifier_from_text(
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
        strategy: _FollowUpNoticeStrategy,
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
