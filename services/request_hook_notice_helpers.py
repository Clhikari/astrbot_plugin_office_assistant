from __future__ import annotations

import re
from collections.abc import Callable, Sequence
from dataclasses import dataclass
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from ..internal_hooks import NoticeBuildContext
    from .prompt_context_service import PromptSection
    from .request_follow_up import FollowUpNoticeStrategy
else:
    NoticeBuildContext = object
    PromptSection = object
    FollowUpNoticeStrategy = object

NoticeAvailabilityChecker = Callable[[set[str]], bool]
BuildFollowUpSection = Callable[..., PromptSection | None]
AppendNoticeSection = Callable[[NoticeBuildContext, PromptSection | None], None]
IsFollowUpRequest = Callable[..., bool]


@dataclass(frozen=True)
class FollowUpNoticeRule:
    strategy: FollowUpNoticeStrategy
    availability_checker: NoticeAvailabilityChecker | None = None


class FollowUpNoticeHelper:
    def __init__(
        self,
        *,
        is_follow_up_request: IsFollowUpRequest,
        build_follow_up_section: BuildFollowUpSection,
        append_notice_section: AppendNoticeSection,
    ) -> None:
        self._is_follow_up_request = is_follow_up_request
        self._build_follow_up_section = build_follow_up_section
        self._append_notice_section = append_notice_section

    def append_first_matching_notice(
        self,
        context: NoticeBuildContext,
        *,
        request_text: str,
        exposed_tool_names: set[str],
        rules: Sequence[FollowUpNoticeRule],
    ) -> bool:
        for rule in rules:
            if not self._matches_rule(
                request_text=request_text,
                exposed_tool_names=exposed_tool_names,
                rule=rule,
            ):
                continue
            section = self._build_follow_up_section(
                request_text=request_text,
                strategy=rule.strategy,
            )
            if section is None:
                continue
            self._append_notice_section(context, section)
            return True
        return False

    def _matches_rule(
        self,
        *,
        request_text: str,
        exposed_tool_names: set[str],
        rule: FollowUpNoticeRule,
    ) -> bool:
        if rule.availability_checker is not None and not rule.availability_checker(
            exposed_tool_names
        ):
            return False
        return bool(
            self._is_follow_up_request(
                request_text=request_text,
                descriptor=rule.strategy.identifier,
            )
        )


@dataclass(frozen=True)
class WorkbookGuideDecision:
    inject_core: bool
    inject_detail: bool


class WorkbookGuideMatcher:
    _TOOL_CALL_RE = re.compile(
        r"(create_workbook|write_rows|export_workbook)",
        flags=re.IGNORECASE,
    )
    _SUBJECT_RE = re.compile(
        r"(\bexcel\b|\bxlsx\b|报表|汇总表|工作簿|多\s*sheet)",
        flags=re.IGNORECASE,
    )
    _GENERATION_RE = re.compile(
        r"(生成|创建|新建|制作|整理成|整理为|写入|填入|输出|导出(?:成|为)?|返回|做(?:成|个|一份)?)",
        flags=re.IGNORECASE,
    )
    _READ_RE = re.compile(
        r"(读取|阅读|查看|打开|解析|提取|\bread\b|\bopen\b|\bparse\b|\bextract\b)",
        flags=re.IGNORECASE,
    )
    _ANALYSIS_RE = re.compile(
        r"(分析|总结|\banaly[sz]e\b)",
        flags=re.IGNORECASE,
    )
    _CONVERSION_RE = re.compile(
        r"(导出(?:成|为)?\s*pdf|"
        r"(?:转换|转成|转为|转到).*(?:pdf|word|docx|ppt|pptx)|"
        r"(?:pdf|word|docx|ppt|pptx).*(?:转换|转成|转为|转到)|"
        r"\bconvert\b)",
        flags=re.IGNORECASE,
    )
    _DETAIL_RE = re.compile(
        r"(create_workbook|write_rows|export_workbook|start_row|多\s*sheet)",
        flags=re.IGNORECASE,
    )

    @classmethod
    def detect(
        cls,
        *,
        request_text: str,
        workbook_follow_up_re: re.Pattern[str],
    ) -> WorkbookGuideDecision:
        if not request_text:
            return WorkbookGuideDecision(inject_core=False, inject_detail=False)

        mentions_tool_call = bool(cls._TOOL_CALL_RE.search(request_text))
        mentions_follow_up_id = bool(workbook_follow_up_re.search(request_text))
        mentions_subject = bool(cls._SUBJECT_RE.search(request_text))
        requests_generation = bool(cls._GENERATION_RE.search(request_text))
        requests_read = bool(cls._READ_RE.search(request_text))
        requests_analysis = bool(cls._ANALYSIS_RE.search(request_text))
        requests_conversion = bool(cls._CONVERSION_RE.search(request_text))
        inject_core = (
            mentions_tool_call
            or mentions_follow_up_id
            or (
                mentions_subject
                and requests_generation
                and not (requests_read or requests_analysis or requests_conversion)
            )
        )
        inject_detail = bool(cls._DETAIL_RE.search(request_text))
        return WorkbookGuideDecision(
            inject_core=inject_core,
            inject_detail=inject_detail,
        )


__all__ = [
    "FollowUpNoticeHelper",
    "FollowUpNoticeRule",
    "WorkbookGuideDecision",
    "WorkbookGuideMatcher",
]
