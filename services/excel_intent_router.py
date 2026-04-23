from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Literal

from ..constants import EXCEL_SUFFIXES
from .upload_types import UploadInfo

ExcelIntentRoute = Literal[
    "new_primitive",
    "new_script",
    "read_existing",
    "modify_existing",
]


@dataclass(frozen=True, slots=True)
class ExcelRouteDecision:
    route: ExcelIntentRoute
    matched_files: tuple[str, ...]
    requires_script: bool
    should_inject_guide: bool
    should_inject_detail: bool = False


def _compile_pattern(*parts: str) -> re.Pattern[str]:
    return re.compile("".join(parts), flags=re.IGNORECASE)


@dataclass(frozen=True, slots=True)
class _FilenameMatch:
    filename: str
    start: int
    end: int


@dataclass(frozen=True, slots=True)
class _ExcelIntentFeatures:
    matched_files: tuple[str, ...]
    mentions_excel_subject: bool
    mentions_excel_tool: bool
    is_conversion_request: bool
    references_current_upload_workbook: bool
    has_read_intent: bool
    has_modify_intent: bool
    has_explicit_modify_intent: bool
    has_new_intent: bool
    has_script_intent: bool
    should_inject_detail: bool


class ExcelIntentRouter:
    _MATCH_CONTEXT_WINDOW = 24
    _SCRIPT_EDIT_SUFFIXES = frozenset({".xlsx"})
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
    _EXCEL_SUBJECT_TERMS = (
        r"\bexcel\b",
        r"\bxlsx\b",
        r"\bxls\b",
        r"工作簿",
        r"sheet",
        r"报表",
        r"汇总表",
    )
    _EXCEL_SUBJECT_RE = _compile_pattern(r"(?:", "|".join(_EXCEL_SUBJECT_TERMS), r")")
    _READ_TERMS = (
        r"读取",
        r"阅读",
        r"查看",
        r"打开",
        r"解析",
        r"提取",
        r"统计",
        r"汇总",
        r"解释",
        r"查询",
        r"看看",
        r"内容",
        r"数据",
        r"\bread\b",
        r"\bopen\b",
        r"\bparse\b",
        r"\bextract\b",
        r"\bsummar(?:y|ize)\b",
    )
    _READ_RE = _compile_pattern(r"(?:", "|".join(_READ_TERMS), r")")
    _MODIFY_TERMS = (
        r"修改",
        r"补写",
        r"更新",
        r"重排",
        r"改写",
        r"调整",
        r"删除",
        r"新增",
        r"追加",
        r"插入",
        r"替换",
        r"生成新版本",
        r"写入.{0,32}(?:\.xlsx|\.xls|工作簿|sheet)",
        r"填入.{0,32}(?:\.xlsx|\.xls|工作簿|sheet)",
        r"加公式",
        r"改样式",
        r"加样式",
        r"加图表",
        r"加条件格式",
        r"加数据验证",
        r"\bmodify\b",
        r"\bedit\b",
        r"\bupdate\b",
        r"\brewrite\b",
        rf"\b(?:add|insert)\s+(?:a|an|the|new)?\s*{_ENGLISH_ADD_INSERT_TARGET_RE}\b",
    )
    _MODIFY_RE = _compile_pattern(r"(?:", "|".join(_MODIFY_TERMS), r")")
    _EXPLICIT_MODIFY_TERMS = (
        r"修改",
        r"补写",
        r"重排",
        r"改写",
        r"调整",
        r"删除",
        r"追加",
        r"插入",
        r"替换",
        r"生成新版本",
        r"加公式",
        r"改样式",
        r"加样式",
        r"加图表",
        r"加条件格式",
        r"加数据验证",
        r"(?:新增|增加|添加)\s*(?:一列|列|一行|行|sheet|工作表|公式|图表|条件格式|数据验证)",
        r"(?:写入|填入).{0,32}(?:\.xlsx|\.xls|工作簿|sheet)",
        r"更新\s*(?:这个|该|当前|现有|已有|原有|文件|工作簿|表格|sheet|xlsx|xls)",
        r"更新.{0,32}(?:数据|内容|单元格|公式|样式)",
        r"\bupdate\s+(?:this|the|current|existing|workbook|file|sheet|xlsx|xls)\b",
        r"\bupdate\b.{0,32}\b(?:data|content|cell|cells|formula|formulas|style|styles|chart|charts|sheet|sheets|worksheet|worksheets)\b",
        rf"\b(?:add|insert)\s+(?:a|an|the|new)?\s*{_ENGLISH_ADD_INSERT_TARGET_RE}\b",
        r"\bmodify\b",
        r"\bedit\b",
        r"\brewrite\b",
    )
    _EXPLICIT_MODIFY_RE = _compile_pattern(
        r"(?:", "|".join(_EXPLICIT_MODIFY_TERMS), r")"
    )
    _NEW_TERMS = (
        r"生成",
        r"创建",
        r"新建",
        r"制作",
        r"整理成",
        r"整理为",
        r"写入",
        r"填入",
        r"输出",
        r"导出(?:成|为)?",
        r"做(?:个|一份|一个)?",
        r"帮我做",
        r"帮我生成",
        r"\bcreate\b",
        r"\bgenerate\b",
        r"\bbuild\b",
    )
    _NEW_RE = _compile_pattern(r"(?:", "|".join(_NEW_TERMS), r")")
    _SCRIPT_TERMS = (
        r"图表",
        r"公式",
        r"条件格式",
        r"数据验证",
        r"样式",
        r"格式刷",
        r"透视表",
        r"下拉框",
        r"\bchart\b",
        r"\bformula\b",
        r"\bconditional\b",
        r"\bvalidation\b",
    )
    _SCRIPT_RE = _compile_pattern(r"(?:", "|".join(_SCRIPT_TERMS), r")")
    _CONVERSION_TERMS = (
        r"导出(?:成|为)?\s*pdf",
        r"(?:转换|转成|转为|转到).*(?:pdf|word|docx|ppt|pptx)",
        r"(?:pdf|word|docx|ppt|pptx).*(?:转换|转成|转为|转到)",
        r"\bconvert\b.{0,64}\b(?:pdf|word|docx|ppt|pptx)\b",
        r"\b(?:pdf|word|docx|ppt|pptx)\b.{0,32}\bconvert\b",
    )
    _CONVERSION_RE = _compile_pattern(r"(?:", "|".join(_CONVERSION_TERMS), r")")
    _DETAIL_TERMS = (
        r"create_workbook",
        r"write_rows",
        r"export_workbook",
        r"start_row",
        r"多\s*sheet",
        r"sheet",
    )
    _DETAIL_RE = _compile_pattern(r"(?:", "|".join(_DETAIL_TERMS), r")")
    _OUTPUT_FILENAME_PREFIX_TERMS = (
        r"文件名(?:叫|为)?",
        r"命名(?:为)?",
        r"保存(?:为|到)",
        r"另存为",
        r"输出(?:成|为|到)",
        r"导出(?:成|为)",
        r"存为",
        r"save(?:d)?\s+as",
        r"save\s+to",
        r"output\s+as",
        r"output\s+to",
        r"export\s+as",
        r"export\s+to",
        r"name(?:d)?(?:\s+as)?",
        r"call(?:ed)?(?:\s+it)?\s+as",
    )
    _OUTPUT_FILENAME_PREFIX_RE = _compile_pattern(
        r"(?:", "|".join(_OUTPUT_FILENAME_PREFIX_TERMS), r")$"
    )
    _OUTPUT_FILENAME_SUFFIX_TERMS = (
        r"\s*(?:作为|当作)\s*输出(?:文件)?",
        r"\s*(?:是|为)\s*输出(?:文件)?",
    )
    _OUTPUT_FILENAME_SUFFIX_RE = _compile_pattern(
        r"^(?:", "|".join(_OUTPUT_FILENAME_SUFFIX_TERMS), r")"
    )
    _ANY_FILENAME_RE = re.compile(r"([^\s'\"`]+?\.[A-Za-z0-9]+)", flags=re.IGNORECASE)
    _CURRENT_UPLOAD_WORKBOOK_TERMS = (
        r"这个表",
        r"该表",
        r"当前表",
        r"这个工作簿",
        r"该工作簿",
        r"当前工作簿",
        r"这个xlsx",
        r"该xlsx",
        r"当前xlsx",
        r"这个xls",
        r"该xls",
        r"当前xls",
        r"这个excel",
        r"该excel",
        r"当前excel",
        r"这个sheet",
        r"该sheet",
        r"当前sheet",
        r"\b(?:this|that|current)\s+(?:sheet|worksheet|workbook|spreadsheet|excel|xlsx|xls)\b",
    )
    _CURRENT_UPLOAD_WORKBOOK_RE = _compile_pattern(
        r"(?:", "|".join(_CURRENT_UPLOAD_WORKBOOK_TERMS), r")"
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
        features = cls._extract_features(
            request_text=request_text,
            upload_infos=upload_infos,
            exposed_tool_names=exposed_tool_names,
        )
        has_excel_file = bool(features.matched_files)
        if not (
            has_excel_file
            or features.mentions_excel_subject
            or features.mentions_excel_tool
        ):
            return None
        if features.is_conversion_request:
            return None

        if has_excel_file and features.has_modify_intent and (
            not features.has_read_intent or features.has_explicit_modify_intent
        ):
            if not cls._all_script_editable_files(features.matched_files):
                return ExcelRouteDecision(
                    route="read_existing",
                    matched_files=features.matched_files,
                    requires_script=False,
                    should_inject_guide=cls._can_read_existing(
                        explicit_tool_name=explicit_tool_name,
                        exposed_tool_names=exposed_tool_names,
                    ),
                )
            return ExcelRouteDecision(
                route="modify_existing",
                matched_files=features.matched_files,
                requires_script=True,
                should_inject_guide=cls._can_run_script(
                    explicit_tool_name=explicit_tool_name,
                    exposed_tool_names=exposed_tool_names,
                ),
            )

        if has_excel_file and features.has_read_intent:
            return ExcelRouteDecision(
                route="read_existing",
                matched_files=features.matched_files,
                requires_script=False,
                should_inject_guide=cls._can_read_existing(
                    explicit_tool_name=explicit_tool_name,
                    exposed_tool_names=exposed_tool_names,
                ),
            )

        if not (features.mentions_excel_subject or has_excel_file):
            return None

        if not features.has_new_intent:
            return None

        if features.has_script_intent:
            return ExcelRouteDecision(
                route="new_script",
                matched_files=features.matched_files,
                requires_script=True,
                should_inject_guide=cls._can_run_script(
                    explicit_tool_name=explicit_tool_name,
                    exposed_tool_names=exposed_tool_names,
                ),
            )

        return ExcelRouteDecision(
            route="new_primitive",
            matched_files=features.matched_files,
            requires_script=False,
            should_inject_guide=cls._can_use_workbook_primitives(
                explicit_tool_name=explicit_tool_name,
                exposed_tool_names=exposed_tool_names,
            ),
            should_inject_detail=features.should_inject_detail,
        )

    @classmethod
    def _match_excel_files(
        cls,
        *,
        request_text: str,
        upload_infos: list[UploadInfo],
        filename_matches: tuple[_FilenameMatch, ...],
        request_mentions_filename: bool,
        references_current_upload_workbook: bool,
    ) -> tuple[str, ...]:
        matched_names: list[str] = []
        excel_upload_infos = [
            info
            for info in upload_infos
            if str(info.get("file_suffix", "")).lower() in cls._EXCEL_SUFFIXES
        ]

        for info in excel_upload_infos:
            stored_name = str(info.get("stored_name", "")).strip()
            original_name = str(info.get("original_name", "")).strip()
            if request_mentions_filename:
                if cls._request_references_upload_name(
                    request_text,
                    stored_name=stored_name,
                    original_name=original_name,
                ):
                    pass
                elif references_current_upload_workbook:
                    pass
                else:
                    continue
            chosen_name = stored_name or original_name
            if chosen_name and chosen_name not in matched_names:
                matched_names.append(chosen_name)

        for match in filename_matches:
            if not str(match.filename).lower().endswith(tuple(cls._EXCEL_SUFFIXES)):
                continue
            if match.filename not in matched_names:
                matched_names.append(match.filename)

        return tuple(matched_names)

    @classmethod
    def _extract_features(
        cls,
        *,
        request_text: str,
        upload_infos: list[UploadInfo],
        exposed_tool_names: set[str],
    ) -> _ExcelIntentFeatures:
        normalized_text = (request_text or "").strip()
        filename_matches = cls._iter_effective_filename_matches(
            request_text=normalized_text,
            pattern=cls._ANY_FILENAME_RE,
        )
        excel_upload_infos = [
            info
            for info in upload_infos
            if str(info.get("file_suffix", "")).lower() in cls._EXCEL_SUFFIXES
        ]
        references_current_upload_workbook = cls._references_current_upload_workbook(
            request_text=normalized_text,
            excel_upload_count=len(excel_upload_infos),
        )
        matched_files = cls._match_excel_files(
            request_text=normalized_text,
            upload_infos=upload_infos,
            filename_matches=filename_matches,
            request_mentions_filename=cls._request_mentions_filename(
                filename_matches=filename_matches
            ),
            references_current_upload_workbook=references_current_upload_workbook,
        )
        return _ExcelIntentFeatures(
            matched_files=matched_files,
            mentions_excel_subject=bool(cls._EXCEL_SUBJECT_RE.search(normalized_text)),
            mentions_excel_tool=bool(
                cls._WORKBOOK_TOOL_NAMES.union(cls._EXCEL_FILE_TOOL_NAMES).intersection(
                    exposed_tool_names
                )
            ),
            is_conversion_request=bool(cls._CONVERSION_RE.search(normalized_text)),
            references_current_upload_workbook=references_current_upload_workbook,
            has_read_intent=bool(cls._READ_RE.search(normalized_text)),
            has_modify_intent=bool(cls._MODIFY_RE.search(normalized_text)),
            has_explicit_modify_intent=cls._has_explicit_modify_intent(normalized_text),
            has_new_intent=bool(cls._NEW_RE.search(normalized_text)),
            has_script_intent=bool(cls._SCRIPT_RE.search(normalized_text)),
            should_inject_detail=bool(cls._DETAIL_RE.search(normalized_text)),
        )

    @classmethod
    def _iter_effective_filename_matches(
        cls,
        *,
        request_text: str,
        pattern: re.Pattern[str],
    ) -> tuple[_FilenameMatch, ...]:
        matches: list[_FilenameMatch] = []
        for match in pattern.finditer(request_text):
            filename = str(match.group(1)).strip()
            if not filename:
                continue
            if cls._looks_like_output_filename_reference(
                request_text=request_text,
                start=match.start(1),
                end=match.end(1),
            ):
                continue
            matches.append(
                _FilenameMatch(
                    filename=filename,
                    start=match.start(1),
                    end=match.end(1),
                )
            )
        return tuple(matches)

    @classmethod
    def _request_mentions_filename(
        cls, *, filename_matches: tuple[_FilenameMatch, ...]
    ) -> bool:
        return bool(filename_matches)

    @classmethod
    def _request_references_upload_name(
        cls,
        request_text: str,
        *,
        stored_name: str,
        original_name: str,
    ) -> bool:
        for candidate in (stored_name, original_name):
            normalized_candidate = candidate.strip()
            if not normalized_candidate:
                continue
            pattern = re.compile(
                rf"(?<![A-Za-z0-9_.-]){re.escape(normalized_candidate)}(?![A-Za-z0-9_.-])",
                flags=re.IGNORECASE,
            )
            if pattern.search(request_text):
                return True
        return False

    @classmethod
    def _references_current_upload_workbook(
        cls,
        *,
        request_text: str,
        excel_upload_count: int,
    ) -> bool:
        if excel_upload_count != 1:
            return False
        return bool(cls._CURRENT_UPLOAD_WORKBOOK_RE.search(request_text))

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
    def _has_explicit_modify_intent(cls, request_text: str) -> bool:
        return bool(cls._EXPLICIT_MODIFY_RE.search(request_text))

    @classmethod
    def _all_script_editable_files(cls, matched_files: tuple[str, ...]) -> bool:
        if not matched_files:
            return True
        return all(
            str(file_name).lower().endswith(tuple(cls._SCRIPT_EDIT_SUFFIXES))
            for file_name in matched_files
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
