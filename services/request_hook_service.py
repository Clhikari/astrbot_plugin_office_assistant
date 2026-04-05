import re
from collections.abc import Awaitable, Callable
from pathlib import Path

import astrbot.api.message_components as Comp
from astrbot.api import logger
from astrbot.api.event import AstrMessageEvent

from ..constants import (
    ALL_OFFICE_SUFFIXES,
    EXECUTION_TOOLS,
    ExposureLevel,
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
from .intent_patterns import (
    extract_document_id,
    should_inject_document_tool_detail,
    should_use_active_document_summary,
)
from .prompt_context_service import PromptContextService
from .upload_types import UploadInfo


class RequestHookService:
    _BUFFERED_USER_INSTRUCTION_RE = re.compile(
        r"\[用户指令\]\s*(?P<instruction>.*?)(?:\n\s*\[|\Z)",
        flags=re.DOTALL,
    )

    def __init__(
        self,
        *,
        auto_block_execution_tools: bool,
        get_cached_upload_infos: Callable[[AstrMessageEvent], list[UploadInfo]],
        extract_upload_source: Callable[
            [Comp.File], Awaitable[tuple[Path | None, str]]
        ],
        store_uploaded_file: Callable[[Path, str], Path],
        allow_external_input_files: bool,
        get_document_prompt_summary: (
            Callable[[str], dict[str, object] | None] | None
        ) = None,
        get_active_document_prompt_summary: (
            Callable[[str], dict[str, object] | None] | None
        ) = None,
        prompt_context_service: PromptContextService | None = None,
    ) -> None:
        self._auto_block_execution_tools = auto_block_execution_tools
        self._get_cached_upload_infos = get_cached_upload_infos
        self._extract_upload_source = extract_upload_source
        self._store_uploaded_file = store_uploaded_file
        self._get_document_prompt_summary = get_document_prompt_summary
        self._get_active_document_prompt_summary = (
            get_active_document_prompt_summary
        )
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

        exposure_level = getattr(context, "exposure_level", ExposureLevel.NONE)
        request_text = self._extract_prompt_text(str(context.request.prompt or ""))
        document_summary = self._resolve_document_summary(
            event=context.event,
            request_text=request_text,
        )
        document_id = str((document_summary or {}).get("document_id") or "").strip()

        if exposure_level == ExposureLevel.NONE:
            return context

        if exposure_level == ExposureLevel.FILE_ONLY:
            self._append_prompt_section(
                context,
                self.prompt_context_service.build_file_only_notice_section(),
            )
            return context

        self._append_prompt_section(
            context,
            self.prompt_context_service.build_document_tool_guide_section(),
        )
        if should_inject_document_tool_detail(
            request_text=request_text,
            document_id=document_id or None,
        ):
            self._append_prompt_section(
                context,
                self.prompt_context_service.build_document_tool_detail_section(),
            )
        if document_summary:
            self._append_prompt_section(
                context,
                self.prompt_context_service.build_document_summary_section(
                    summary=document_summary
                ),
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

    def get_cached_upload_infos(
        self, event: AstrMessageEvent
    ) -> list[UploadInfo]:
        return list(self._get_cached_upload_infos(event))

    def get_active_document_prompt_summary(
        self, event: AstrMessageEvent
    ) -> dict[str, object] | None:
        if self._get_active_document_prompt_summary is None:
            return None
        session_id = str(getattr(event, "unified_msg_origin", "") or "").strip()
        if not session_id:
            return None
        return self._get_active_document_prompt_summary(session_id)

    def _resolve_document_summary(
        self,
        *,
        event: AstrMessageEvent,
        request_text: str,
    ) -> dict[str, object] | None:
        document_id = extract_document_id(request_text)
        if document_id and self._get_document_prompt_summary:
            summary = self._get_document_prompt_summary(document_id)
            if summary:
                return summary
            return None
        if not should_use_active_document_summary(request_text):
            return None
        return self.get_active_document_prompt_summary(event)

    @staticmethod
    def _append_prompt_section(
        context: NoticeBuildContext,
        section,
    ) -> None:
        if not section or not section.content:
            return
        context.notices.append(section.content)
        section_names = getattr(context, "section_names", None)
        if isinstance(section_names, list):
            section_names.append(section.name)

    async def append_uploaded_file_notices(
        self,
        context: NoticeBuildContext,
    ) -> NoticeBuildContext:
        if not context.can_process_upload:
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

        if self._should_append_uploaded_file_scene_notice(
            event=event,
            prompt=str(req.prompt or ""),
        ):
            self._append_prompt_section(
                context,
                self.prompt_context_service.build_uploaded_file_scene_section(
                    file_count=len(readable_upload_infos),
                    document_workflow=(
                        getattr(context, "exposure_level", ExposureLevel.NONE)
                        == ExposureLevel.DOCUMENT_FULL
                    ),
                ),
            )

        if len(readable_upload_infos) == 1:
            info = readable_upload_infos[0]
            self._append_prompt_section(
                context,
                self.prompt_context_service.build_uploaded_file_notice_section(
                    type_desc=info["type_desc"],
                    original_name=info["original_name"],
                    file_suffix=info["file_suffix"],
                    stored_name=info["stored_name"],
                    source_path=info["source_path"],
                ),
            )
            return context

        self._append_prompt_section(
            context,
            self.prompt_context_service.build_uploaded_file_summary_section(
                upload_infos=readable_upload_infos
            ),
        )
        return context

    @classmethod
    def _should_append_uploaded_file_scene_notice(
        cls,
        *,
        event: AstrMessageEvent,
        prompt: str,
    ) -> bool:
        if not getattr(event, "_buffered", False):
            return True
        return cls._BUFFERED_USER_INSTRUCTION_RE.search(prompt.strip()) is not None

    async def apply_execution_tool_block(
        self,
        context: ToolExposureContext,
    ) -> ToolExposureContext:
        if context.request.func_tool and self._auto_block_execution_tools:
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
            and context.explicit_tool_name
            in set(getattr(context, "allowed_tool_names", ()))
        ):
            for tool_name in tuple(getattr(context, "allowed_tool_names", ())):
                if tool_name != context.explicit_tool_name:
                    context.request.func_tool.remove_tool(tool_name)
            logger.info(
                "[文件管理] 检测到用户显式指定工具 %s，本轮仅保留该文件工具",
                context.explicit_tool_name,
            )
        return context
