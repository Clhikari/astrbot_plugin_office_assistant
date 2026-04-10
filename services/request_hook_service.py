import re
from collections.abc import Awaitable, Callable
from pathlib import Path

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


class RequestHookService:
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
    _DOCUMENT_ID_FOLLOW_UP_RE = re.compile(r"\bdocument_id\b", flags=re.IGNORECASE)
    _DOCUMENT_ID_CAPTURE_RE = re.compile(
        r"\bdocument_id\b(?:\s*(?:[:=]|为|是)\s*|\s+)[\"']?(?P<document_id>[A-Za-z0-9_-]+)[\"']?",
        flags=re.IGNORECASE,
    )
    _DOCUMENT_CORE_NOTICE_KEY = "document_core_guide"
    _DOCUMENT_DETAIL_NOTICE_KEY = "document_detail_guide"

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
    ) -> None:
        self._auto_block_execution_tools = auto_block_execution_tools
        self._get_cached_upload_infos = get_cached_upload_infos
        self._extract_upload_source = extract_upload_source
        self._store_uploaded_file = store_uploaded_file
        self._consume_session_notice_once = consume_session_notice_once
        self._lookup_document_summary = lookup_document_summary
        self.prompt_context_service = prompt_context_service or PromptContextService(
            allow_external_input_files=allow_external_input_files
        )
        self._notice_hooks = [
            self.append_document_tool_guide_notice,
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

    async def append_document_tool_guide_notice(
        self,
        context: NoticeBuildContext,
    ) -> NoticeBuildContext:
        if not (context.should_expose and context.request.func_tool):
            return context

        request_text = self._extract_prompt_text(str(context.request.prompt or ""))
        if self._is_document_follow_up(request_text=request_text):
            section = self._build_document_follow_up_section(request_text=request_text)
            if section is not None:
                self._append_notice_section(context, section)
                return context
        should_inject_core = self._should_inject_document_tool_guide(
            request_text=request_text
        )
        should_inject_detail = self._should_inject_document_tool_detail(
            request_text=request_text
        )
        if not should_inject_core and not should_inject_detail:
            return context

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
        return context

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
        if not request_text:
            return False
        return bool(cls._DOCUMENT_ID_FOLLOW_UP_RE.search(request_text))

    @classmethod
    def _extract_document_id(cls, *, request_text: str) -> str:
        if not request_text:
            return ""
        match = cls._DOCUMENT_ID_CAPTURE_RE.search(request_text)
        if not match:
            return ""
        return str(match.group("document_id") or "").strip()

    def _build_document_follow_up_section(
        self,
        *,
        request_text: str,
    ) -> PromptSection | None:
        document_id = self._extract_document_id(request_text=request_text)
        if not document_id:
            return None
        if self._lookup_document_summary is None:
            return self.prompt_context_service.build_document_follow_up_missing_section(
                document_id=document_id
            )
        try:
            summary = self._lookup_document_summary(document_id)
        except KeyError:
            summary = None
        except Exception as exc:
            logger.exception(
                f"[文件管理] 查询文档会话摘要失败 document_id={document_id}: {exc}"
            )
            raise
        if not isinstance(summary, dict):
            return self.prompt_context_service.build_document_follow_up_missing_section(
                document_id=document_id
            )
        return self.prompt_context_service.build_document_follow_up_section(
            document_id=document_id,
            status=str(summary.get("status") or ""),
            block_count=int(summary.get("block_count") or 0),
        )

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
            for tool_name in FILE_TOOLS:
                if tool_name != context.explicit_tool_name:
                    context.request.func_tool.remove_tool(tool_name)
            logger.info(
                "[文件管理] 检测到用户显式指定工具 %s，本轮仅保留该文件工具",
                context.explicit_tool_name,
            )
        return context
