from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Literal

from ..constants import EXCEL_SUFFIXES
from .upload_types import UploadInfo

ExcelIntentRoute = Literal[
    "excel_context",
]


@dataclass(frozen=True, slots=True)
class ExcelRouteDecision:
    route: ExcelIntentRoute
    matched_files: tuple[str, ...]
    requires_script: bool
    should_inject_guide: bool
    should_inject_detail: bool = False


class ExcelIntentRouter:
    _MATCH_CONTEXT_WINDOW = 24
    _ENGLISH_ADD_INSERT_TARGET_RE = (
        r"(?:column|columns|row|rows|sheet|sheets|worksheet|worksheets|"
        r"formula|formulas|chart|charts|style|styles|cell|cells|"
        r"conditional formatting|data validation)"
    )
    _WORKBOOK_TOOL_NAMES = frozenset(
        {"create_workbook", "write_rows", "export_workbook"}
    )
    _EXCEL_FILE_TOOL_NAMES = frozenset({"read_workbook", "execute_excel_script"})
    _EXCEL_SUFFIXES = EXCEL_SUFFIXES
    _EXCEL_SUBJECT_RE = re.compile(
        r"(\bexcel\b|\bxlsx\b|\bxls\b|工作簿|sheet|报表|汇总表)",
        flags=re.IGNORECASE,
    )
    _SCRIPT_RE = re.compile(
        r"(图表|公式|条件格式|数据验证|样式|格式|格式刷|透视表|下拉框|"
        r"冻结|列宽|联动|引用|清洗|去重|重复表头|异常|对账|匹配|"
        r"调整项|库存|问卷|考试|得分|"
        r"\bchart\b|\bformula\b|\bconditional\b|\bvalidation\b)",
        flags=re.IGNORECASE,
    )
    _WORKBOOK_ACTION_RE = re.compile(
        r"(生成|创建|新建|制作|整理成|整理为|写入|填入|输出|导出(?:成|为)?|"
        r"做(?:个|一份|一个)?|帮我做|帮我生成|workbook_id|"
        r"create_workbook|write_rows|export_workbook|"
        r"\bcreate\b|\bgenerate\b|\bbuild\b|\bexport\b)",
        flags=re.IGNORECASE,
    )
    _CONVERSION_RE = re.compile(
        r"(导出(?:成|为)?\s*pdf|"
        r"(?:转换|转成|转为|转到).*(?:pdf|word|docx|ppt|pptx)|"
        r"(?:pdf|word|docx|ppt|pptx).*(?:转换|转成|转为|转到)|"
        r"\bconvert\b)",
        flags=re.IGNORECASE,
    )
    _XLS_TO_XLSX_RE = re.compile(
        r"((?:转换|转成|转为|另存为|保存为|导出(?:成|为)?|升级(?:成|为)?)\s*\.?xlsx|"
        r"\.xls\b.{0,32}\.xlsx\b|"
        r"\bxls\b.{0,32}\bxlsx\b)",
        flags=re.IGNORECASE,
    )
    _DETAIL_RE = re.compile(
        r"(create_workbook|write_rows|export_workbook|start_row|多\s*sheet)",
        flags=re.IGNORECASE,
    )
    _FILENAME_RE = re.compile(
        r"([^\s'\"`「」《》“”‘’，,]+?\.(?:xlsx|xls))",
        flags=re.IGNORECASE,
    )
    _OUTPUT_FILENAME_PREFIX_RE = re.compile(
        r"(文件名(?:叫|为)?|命名(?:为)?|保存(?:为|到)|另存为|输出(?:成|为|到)|"
        r"导出(?:成|为)|存为|"
        r"(?:生成|创建|新建|制作)(?:一个|一份|一张|这个|这份)?"
        r"(?:\s*(?:Excel|xlsx|工作簿|表格))?\s*[「《“\"'`(（]*)$",
        flags=re.IGNORECASE,
    )
    _OUTPUT_FILENAME_SUFFIX_RE = re.compile(
        r"^(?:\s*(?:作为|当作)\s*输出(?:文件)?|"
        r"\s*(?:是|为)\s*输出(?:文件)?)",
        flags=re.IGNORECASE,
    )

    @classmethod
    def decide(
        cls,
        *,
        request_text: str,
        upload_infos: list[UploadInfo],
        explicit_tool_name: str | None,
        exposed_tool_names: set[str],
    ) -> ExcelRouteDecision | None:
        normalized_text = (request_text or "").strip()
        matched_files = cls._match_excel_files(
            request_text=normalized_text,
            upload_infos=upload_infos,
        )
        has_excel_file = bool(matched_files)
        mentions_excel = bool(cls._EXCEL_SUBJECT_RE.search(normalized_text))
        mentions_excel_tool = bool(
            cls._WORKBOOK_TOOL_NAMES.union(cls._EXCEL_FILE_TOOL_NAMES).intersection(
                exposed_tool_names
            )
        )
        if not (has_excel_file or mentions_excel or mentions_excel_tool):
            return None
        if cls._CONVERSION_RE.search(normalized_text):
            return None

        return ExcelRouteDecision(
            route="excel_context",
            matched_files=matched_files,
            requires_script=False,
            should_inject_guide=cls._can_inject_any_excel_guide(
                request_text=normalized_text,
                explicit_tool_name=explicit_tool_name,
                exposed_tool_names=exposed_tool_names,
            ),
            should_inject_detail=bool(cls._DETAIL_RE.search(normalized_text)),
        )

    @classmethod
    def _match_excel_files(
        cls,
        *,
        request_text: str,
        upload_infos: list[UploadInfo],
    ) -> tuple[str, ...]:
        matched_names: list[str] = []

        for info in upload_infos:
            suffix = str(info.get("file_suffix", "")).lower()
            if suffix not in cls._EXCEL_SUFFIXES:
                continue
            stored_name = str(info.get("stored_name", "")).strip()
            original_name = str(info.get("original_name", "")).strip()
            chosen_name = stored_name or original_name
            if chosen_name and chosen_name not in matched_names:
                matched_names.append(chosen_name)

        for match in cls._FILENAME_RE.finditer(request_text):
            normalized_match = str(match.group(1)).strip()
            if cls._looks_like_output_filename_reference(
                request_text=request_text,
                start=match.start(1),
                end=match.end(1),
            ):
                continue
            if normalized_match and normalized_match not in matched_names:
                matched_names.append(normalized_match)

        return tuple(matched_names)

    @classmethod
    def _looks_like_output_filename_reference(
        cls,
        *,
        request_text: str,
        start: int,
        end: int,
    ) -> bool:
        context_start = max(0, start - cls._MATCH_CONTEXT_WINDOW)
        context_end = min(len(request_text), end + cls._MATCH_CONTEXT_WINDOW)
        prefix = request_text[context_start:start].rstrip()
        suffix = request_text[end:context_end].lstrip()
        return bool(
            cls._OUTPUT_FILENAME_PREFIX_RE.search(prefix)
            or cls._OUTPUT_FILENAME_SUFFIX_RE.search(suffix)
        )

    @classmethod
    def mentions_script_feature(cls, request_text: str) -> bool:
        return bool(cls._SCRIPT_RE.search(request_text or ""))

    @classmethod
    def mentions_workbook_action(cls, request_text: str) -> bool:
        return bool(cls._WORKBOOK_ACTION_RE.search(request_text or ""))

    @classmethod
    def _can_inject_any_excel_guide(
        cls,
        *,
        request_text: str,
        explicit_tool_name: str | None,
        exposed_tool_names: set[str],
    ) -> bool:
        if explicit_tool_name:
            if explicit_tool_name in cls._EXCEL_FILE_TOOL_NAMES:
                return explicit_tool_name in exposed_tool_names
            if explicit_tool_name in cls._WORKBOOK_TOOL_NAMES:
                return cls._WORKBOOK_TOOL_NAMES.issubset(exposed_tool_names)
            return False
        return bool(
            cls._EXCEL_FILE_TOOL_NAMES.intersection(exposed_tool_names)
            or (
                cls._WORKBOOK_TOOL_NAMES.issubset(exposed_tool_names)
                and cls.mentions_workbook_action(request_text)
            )
        )

    @classmethod
    def _can_use_workbook_primitives(
        cls,
        *,
        explicit_tool_name: str | None,
        exposed_tool_names: set[str],
    ) -> bool:
        if explicit_tool_name and explicit_tool_name not in cls._WORKBOOK_TOOL_NAMES:
            return False
        return cls._WORKBOOK_TOOL_NAMES.issubset(exposed_tool_names)

    @classmethod
    def _can_read_existing(
        cls,
        *,
        explicit_tool_name: str | None,
        exposed_tool_names: set[str],
    ) -> bool:
        if explicit_tool_name and explicit_tool_name != "read_workbook":
            return False
        return "read_workbook" in exposed_tool_names

    @classmethod
    def _can_run_script(
        cls,
        *,
        explicit_tool_name: str | None,
        exposed_tool_names: set[str],
    ) -> bool:
        if explicit_tool_name and explicit_tool_name != "execute_excel_script":
            return False
        return "execute_excel_script" in exposed_tool_names


__all__ = ["ExcelIntentRoute", "ExcelIntentRouter", "ExcelRouteDecision"]
