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
    should_inject_guide: bool


class ExcelIntentRouter:
    _EXCEL_SUBJECT_TERMS = (
        "excel",
        "xlsx",
        "xls",
        "工作簿",
        "sheet",
        "表格",
    )
    _CONVERSION_ACTION_TERMS = (
        "导出成",
        "导出为",
        "转换",
        "转成",
        "转为",
        "转到",
        "convert",
    )
    _NON_EXCEL_CONVERSION_TARGET_TERMS = ("pdf", "word", "docx", "ppt", "pptx")
    _WORKBOOK_TOOL_NAMES = frozenset(
        {"create_workbook", "write_rows", "export_workbook"}
    )
    _EXCEL_FILE_TOOL_NAMES = frozenset({"read_workbook", "execute_excel_script"})
    _EXCEL_SUFFIXES = EXCEL_SUFFIXES
    _FILENAME_RE = re.compile(
        r"([^\s'\"`「」《》“”‘’，,]+?\.(?:xlsx|xls))",
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
        has_uploaded_excel = cls._has_uploaded_excel(upload_infos)
        mentions_excel = cls._mentions_excel_context(normalized_text)
        uses_explicit_excel_tool = explicit_tool_name in (
            cls._WORKBOOK_TOOL_NAMES | cls._EXCEL_FILE_TOOL_NAMES
        )
        if not (has_uploaded_excel or mentions_excel or uses_explicit_excel_tool):
            return None
        if cls._mentions_non_excel_conversion(normalized_text):
            return None

        return ExcelRouteDecision(
            route="excel_context",
            should_inject_guide=cls._can_inject_any_excel_guide(
                explicit_tool_name=explicit_tool_name,
                exposed_tool_names=exposed_tool_names,
            ),
        )

    @classmethod
    def _has_uploaded_excel(cls, upload_infos: list[UploadInfo]) -> bool:
        for info in upload_infos:
            suffix = str(info.get("file_suffix", "")).lower()
            if suffix in cls._EXCEL_SUFFIXES:
                return True
        return False

    @classmethod
    def _mentions_excel_context(cls, request_text: str) -> bool:
        return cls._contains_any(request_text, cls._EXCEL_SUBJECT_TERMS) or bool(
            cls._FILENAME_RE.search(request_text)
        )

    @classmethod
    def _contains_any(cls, request_text: str, terms: tuple[str, ...]) -> bool:
        normalized_text = (request_text or "").lower()
        return any(term.lower() in normalized_text for term in terms)

    @classmethod
    def _mentions_non_excel_conversion(cls, request_text: str) -> bool:
        normalized_text = (request_text or "").lower()
        return cls._contains_any(
            normalized_text,
            cls._CONVERSION_ACTION_TERMS,
        ) and cls._contains_any(
            normalized_text,
            cls._NON_EXCEL_CONVERSION_TARGET_TERMS,
        )

    @classmethod
    def _can_inject_any_excel_guide(
        cls,
        *,
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
            or cls._WORKBOOK_TOOL_NAMES.issubset(exposed_tool_names)
        )

__all__ = ["ExcelIntentRoute", "ExcelIntentRouter", "ExcelRouteDecision"]
