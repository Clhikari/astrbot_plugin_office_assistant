import hashlib
from dataclasses import dataclass

from astrbot.api import logger

from ..prompts.scenes.uploaded_file import (
    build_buffered_upload_prompt,
    build_uploaded_file_context_notice,
)
from ..prompts.static import (
    build_document_follow_up_missing_notice,
    build_document_follow_up_notice,
    build_document_tools_core_notice,
    build_document_tools_detail_notice,
    build_excel_domain_hints,
    build_excel_read_notice,
    build_excel_routing_notice,
    build_excel_script_notice,
    build_excel_script_unavailable_notice,
    build_tools_denied_notice,
    build_workbook_follow_up_missing_notice,
    build_workbook_follow_up_notice,
    build_workbook_tools_core_notice,
    build_workbook_tools_detail_notice,
)
from .upload_types import UploadInfo

SECTION_STATIC_ACCESS = "static_access"
SECTION_STATIC_DOCUMENT_TOOLS = "static_document_tools"
SECTION_STATIC_DOCUMENT_TOOLS_DETAIL = "static_document_tools_detail"
SECTION_STATIC_EXCEL_ROUTING = "static_excel_routing"
SECTION_STATIC_EXCEL_READ = "static_excel_read"
SECTION_STATIC_EXCEL_SCRIPT = "static_excel_script"
SECTION_STATIC_EXCEL_DOMAIN = "static_excel_domain"
SECTION_STATIC_EXCEL_SCRIPT_UNAVAILABLE = "static_excel_script_unavailable"
SECTION_STATIC_WORKBOOK_TOOLS = "static_workbook_tools"
SECTION_STATIC_WORKBOOK_TOOLS_DETAIL = "static_workbook_tools_detail"
SECTION_SCENE_UPLOADED_CONTEXT = "scene_uploaded_context"
SECTION_DYNAMIC_DOCUMENT_FOLLOW_UP = "dynamic_document_follow_up"
SECTION_DYNAMIC_WORKBOOK_FOLLOW_UP = "dynamic_workbook_follow_up"


@dataclass(frozen=True, slots=True)
class PromptSection:
    name: str
    content: str
    target: str = "prompt_suffix"


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
        paired_sections = list(zip(ordered_section_names, ordered_notices))
        if not paired_sections:
            return ", ".join(ordered_section_names) or "none"
        group_totals: dict[str, int] = {}
        length_trace = ", ".join(
            f"{section_name}:{len(notice)}" for section_name, notice in paired_sections
        )
        for section_name, notice in paired_sections:
            group_name, _, _ = section_name.partition("_")
            group_totals[group_name] = group_totals.get(group_name, 0) + len(notice)
        group_trace = ", ".join(
            f"{group_name}:{group_totals[group_name]}"
            for group_name in cls._SECTION_GROUP_ORDER
            if group_name in group_totals
        )
        total_length = sum(len(notice) for notice in ordered_notices)
        digest_source = "|".join(
            f"{section_name}:{len(notice)}" for section_name, notice in paired_sections
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
            target="system",
        )

    def build_tools_denied_notice(self) -> str:
        return self.render_sections(self.build_tools_denied_section())

    def build_document_tool_guide_section(self) -> PromptSection:
        return PromptSection(
            name=SECTION_STATIC_DOCUMENT_TOOLS,
            content=build_document_tools_core_notice(),
            target="prompt_suffix",
        )

    def build_document_tool_detail_section(self) -> PromptSection:
        return PromptSection(
            name=SECTION_STATIC_DOCUMENT_TOOLS_DETAIL,
            content=build_document_tools_detail_notice(),
            target="prompt_suffix",
        )

    def build_excel_routing_section(self) -> PromptSection:
        return PromptSection(
            name=SECTION_STATIC_EXCEL_ROUTING,
            content=build_excel_routing_notice(),
            target="prompt_suffix",
        )

    def build_excel_read_section(self) -> PromptSection:
        return PromptSection(
            name=SECTION_STATIC_EXCEL_READ,
            content=build_excel_read_notice(),
            target="prompt_suffix",
        )

    def build_excel_script_section(self) -> PromptSection:
        return PromptSection(
            name=SECTION_STATIC_EXCEL_SCRIPT,
            content=build_excel_script_notice(),
            target="prompt_suffix",
        )

    def build_excel_domain_section(self, scenario: str) -> PromptSection:
        return PromptSection(
            name=SECTION_STATIC_EXCEL_DOMAIN,
            content=build_excel_domain_hints(scenario),
            target="prompt_suffix",
        )

    def build_excel_script_unavailable_section(self) -> PromptSection:
        return PromptSection(
            name=SECTION_STATIC_EXCEL_SCRIPT_UNAVAILABLE,
            content=build_excel_script_unavailable_notice(),
            target="prompt_suffix",
        )

    def build_workbook_tool_guide_section(self) -> PromptSection:
        return PromptSection(
            name=SECTION_STATIC_WORKBOOK_TOOLS,
            content=build_workbook_tools_core_notice(),
            target="prompt_suffix",
        )

    def build_workbook_tool_detail_section(self) -> PromptSection:
        return PromptSection(
            name=SECTION_STATIC_WORKBOOK_TOOLS_DETAIL,
            content=build_workbook_tools_detail_notice(),
            target="prompt_suffix",
        )

    def build_uploaded_file_context_section(
        self,
        *,
        upload_infos: list[UploadInfo],
    ) -> PromptSection:
        return PromptSection(
            name=SECTION_SCENE_UPLOADED_CONTEXT,
            content=build_uploaded_file_context_notice(upload_infos=upload_infos),
            target="prompt_suffix",
        )

    def build_document_follow_up_section(
        self,
        *,
        document_id: str,
        status: str,
        block_count: int,
    ) -> PromptSection:
        return PromptSection(
            name=SECTION_DYNAMIC_DOCUMENT_FOLLOW_UP,
            content=build_document_follow_up_notice(
                document_id=document_id,
                status=status,
                block_count=block_count,
            ),
            target="prompt_suffix",
        )

    def build_document_follow_up_missing_section(
        self,
        *,
        document_id: str,
    ) -> PromptSection:
        return PromptSection(
            name=SECTION_DYNAMIC_DOCUMENT_FOLLOW_UP,
            content=build_document_follow_up_missing_notice(document_id=document_id),
            target="prompt_suffix",
        )

    def build_workbook_follow_up_section(
        self,
        *,
        workbook_id: str,
        status: str,
        sheet_names: list[str],
        sheet_count: int,
        latest_written_sheets: list[str],
        next_allowed_actions: list[str],
    ) -> PromptSection:
        return PromptSection(
            name=SECTION_DYNAMIC_WORKBOOK_FOLLOW_UP,
            content=build_workbook_follow_up_notice(
                workbook_id=workbook_id,
                status=status,
                sheet_names=sheet_names,
                sheet_count=sheet_count,
                latest_written_sheets=latest_written_sheets,
                next_allowed_actions=next_allowed_actions,
            ),
            target="prompt_suffix",
        )

    def build_workbook_follow_up_missing_section(
        self,
        *,
        workbook_id: str,
    ) -> PromptSection:
        return PromptSection(
            name=SECTION_DYNAMIC_WORKBOOK_FOLLOW_UP,
            content=build_workbook_follow_up_missing_notice(workbook_id=workbook_id),
            target="prompt_suffix",
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
        )
