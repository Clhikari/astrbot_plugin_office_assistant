import hashlib
from dataclasses import dataclass

from astrbot.api import logger

from ..prompts.scenes.uploaded_file import (
    build_buffered_upload_prompt,
    build_uploaded_file_notice,
    build_uploaded_file_scene_notice,
    build_uploaded_file_summary_notice,
)
from ..prompts.static import (
    build_document_tools_core_notice,
    build_document_tools_detail_notice,
    build_tools_denied_notice,
)
from .upload_types import UploadInfo

SECTION_STATIC_ACCESS = "static_access"
SECTION_STATIC_DOCUMENT_TOOLS = "static_document_tools"
SECTION_STATIC_DOCUMENT_TOOLS_DETAIL = "static_document_tools_detail"
SECTION_SCENE_UPLOADED_FILE = "scene_uploaded_file"
SECTION_DYNAMIC_UPLOAD_SUMMARY = "dynamic_upload_summary"
SECTION_DYNAMIC_DOCUMENT_SUMMARY = "dynamic_document_summary"


@dataclass(frozen=True, slots=True)
class PromptSection:
    name: str
    content: str


class PromptContextService:
    _SECTION_GROUP_ORDER = {
        "static": 0,
        "scene": 1,
        "dynamic": 2,
    }

    def __init__(self, *, allow_external_input_files: bool) -> None:
        self._allow_external_input_files = allow_external_input_files

    @staticmethod
    def render_sections(*sections: PromptSection | None) -> str:
        return "".join(
            section.content for section in sections if section and section.content
        )

    @classmethod
    def order_notice_sections(
        cls,
        *,
        section_names: list[str],
        notices: list[str],
    ) -> tuple[list[str], list[str]]:
        ordered_section_names = list(section_names)
        ordered_notices = list(notices)
        if len(ordered_section_names) != len(ordered_notices):
            logger.debug(
                "[文件管理] Prompt section mismatch: sections=%s notices=%s",
                len(ordered_section_names),
                len(ordered_notices),
            )
            return ordered_section_names, ordered_notices

        indexed_sections = list(
            enumerate(zip(ordered_section_names, ordered_notices, strict=True))
        )
        indexed_sections.sort(
            key=lambda item: (
                cls._section_group_rank(item[1][0]),
                item[0],
            )
        )
        return (
            [section_name for _, (section_name, _) in indexed_sections],
            [notice for _, (_, notice) in indexed_sections],
        )

    @classmethod
    def build_section_trace(
        cls,
        *,
        section_names: list[str],
        notices: list[str],
    ) -> str:
        ordered_section_names, ordered_notices = cls.order_notice_sections(
            section_names=section_names,
            notices=notices,
        )
        if not ordered_section_names:
            return "none"
        group_totals: dict[str, int] = {}
        length_trace = ", ".join(
            f"{section_name}:{len(notice)}"
            for section_name, notice in zip(
                ordered_section_names,
                ordered_notices,
                strict=True,
            )
        )
        for section_name, notice in zip(
            ordered_section_names,
            ordered_notices,
            strict=True,
        ):
            group_name, _, _ = section_name.partition("_")
            group_totals[group_name] = group_totals.get(group_name, 0) + len(notice)
        group_trace = ", ".join(
            f"{group_name}:{group_totals[group_name]}"
            for group_name in cls._SECTION_GROUP_ORDER
            if group_name in group_totals
        )
        total_length = sum(len(notice) for notice in ordered_notices)
        digest_source = "|".join(
            f"{section_name}:{len(notice)}"
            for section_name, notice in zip(
                ordered_section_names,
                ordered_notices,
                strict=True,
            )
        )
        digest = hashlib.sha1(digest_source.encode("utf-8")).hexdigest()[:8]
        return (
            f"{', '.join(ordered_section_names)} "
            f"[len={length_trace}] [groups={group_trace}] "
            f"[total={total_length}] [sig={digest}]"
        )

    @classmethod
    def _section_group_rank(cls, section_name: str) -> int:
        group_name, _, _ = section_name.partition("_")
        return cls._SECTION_GROUP_ORDER.get(group_name, len(cls._SECTION_GROUP_ORDER))

    def build_tools_denied_section(self) -> PromptSection:
        return PromptSection(
            name=SECTION_STATIC_ACCESS,
            content=build_tools_denied_notice(),
        )

    def build_tools_denied_notice(self) -> str:
        return self.render_sections(self.build_tools_denied_section())

    def build_document_tool_guide_section(self) -> PromptSection:
        return PromptSection(
            name=SECTION_STATIC_DOCUMENT_TOOLS,
            content=build_document_tools_core_notice(),
        )

    def build_document_tool_detail_section(self) -> PromptSection:
        return PromptSection(
            name=SECTION_STATIC_DOCUMENT_TOOLS_DETAIL,
            content=build_document_tools_detail_notice(),
        )

    def build_document_tool_guide_notice(self) -> str:
        return self.render_sections(
            self.build_document_tool_guide_section(),
            self.build_document_tool_detail_section(),
        )

    def build_document_summary_section(
        self,
        *,
        summary: dict[str, object],
    ) -> PromptSection:
        next_allowed_actions = (
            ", ".join(
                str(action)
                for action in (summary.get("next_allowed_actions") or [])
                if action
            )
            or "暂无"
        )
        status = str(summary.get("status") or "unknown")
        block_count = int(summary.get("block_count") or 0)
        document_id = str(summary.get("document_id") or "").strip()
        return PromptSection(
            name=SECTION_DYNAMIC_DOCUMENT_SUMMARY,
            content=(
                "\n[System Notice] 当前文档状态摘要\n"
                f"- document_id: {document_id}\n"
                f"- 状态: {status}\n"
                f"- 块数: {block_count}\n"
                f"- 下一步: {next_allowed_actions}\n"
            ),
        )

    def build_document_summary_notice(
        self,
        *,
        summary: dict[str, object],
    ) -> str:
        return self.render_sections(
            self.build_document_summary_section(summary=summary)
        )

    def build_uploaded_file_notice_section(
        self,
        *,
        type_desc: str,
        original_name: str,
        file_suffix: str,
        stored_name: str,
        source_path: str,
    ) -> PromptSection:
        return PromptSection(
            name=SECTION_DYNAMIC_UPLOAD_SUMMARY,
            content=build_uploaded_file_notice(
                type_desc=type_desc,
                original_name=original_name,
                file_suffix=file_suffix,
                stored_name=stored_name,
                source_path=source_path,
                allow_external_input_files=self._allow_external_input_files,
            ),
        )

    def build_uploaded_file_scene_section(
        self,
        *,
        file_count: int,
    ) -> PromptSection:
        return PromptSection(
            name=SECTION_SCENE_UPLOADED_FILE,
            content=build_uploaded_file_scene_notice(
                file_count=file_count,
                allow_external_input_files=self._allow_external_input_files,
            ),
        )

    def build_uploaded_file_notice(
        self,
        *,
        type_desc: str,
        original_name: str,
        file_suffix: str,
        stored_name: str,
        source_path: str,
    ) -> str:
        return self.render_sections(
            self.build_uploaded_file_notice_section(
                type_desc=type_desc,
                original_name=original_name,
                file_suffix=file_suffix,
                stored_name=stored_name,
                source_path=source_path,
            )
        )

    def build_uploaded_file_summary_section(
        self,
        *,
        upload_infos: list[UploadInfo],
    ) -> PromptSection:
        return PromptSection(
            name=SECTION_DYNAMIC_UPLOAD_SUMMARY,
            content=build_uploaded_file_summary_notice(
                upload_infos=upload_infos,
                allow_external_input_files=self._allow_external_input_files,
            ),
        )

    def build_uploaded_file_summary_notice(
        self,
        *,
        upload_infos: list[UploadInfo],
    ) -> str:
        return self.render_sections(
            self.build_uploaded_file_summary_section(upload_infos=upload_infos)
        )

    def build_buffered_upload_prompt(
        self,
        *,
        upload_infos: list[UploadInfo],
        user_instruction: str,
    ) -> str:
        return build_buffered_upload_prompt(
            upload_infos=upload_infos,
            user_instruction=user_instruction,
            allow_external_input_files=self._allow_external_input_files,
        )
