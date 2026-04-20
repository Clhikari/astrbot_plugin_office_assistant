from __future__ import annotations

import asyncio
from collections.abc import AsyncGenerator, Callable
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING

import mcp

from astrbot.api import logger
from astrbot.api.event import AstrMessageEvent

from ..constants import PDF_SUFFIX, SUFFIX_TO_OFFICE_TYPE, TEXT_SUFFIXES, OfficeType
from ..utils import extract_excel_sheets, format_file_size, safe_error_message

if TYPE_CHECKING:
    from .word_read_service import WordReadService
    from .workspace_service import WorkspaceService


class FileReadService:
    _EXCEL_SUFFIXES = frozenset({".xlsx", ".xls"})

    @dataclass(frozen=True, slots=True)
    class ReadTarget:
        resolved_path: Path
        display_name: str
        file_size: int
        max_size: int
        suffix: str

    def __init__(
        self,
        *,
        workspace_service: WorkspaceService,
        word_read_service: WordReadService,
        allow_external_input_files: bool,
        is_group_feature_enabled: Callable[[AstrMessageEvent], bool],
        check_permission: Callable[[AstrMessageEvent], bool],
        group_feature_disabled_error: Callable[[], str],
    ) -> None:
        self._workspace_service = workspace_service
        self._word_read_service = word_read_service
        self._allow_external_input_files = allow_external_input_files
        self._is_group_feature_enabled = is_group_feature_enabled
        self._check_permission = check_permission
        self._group_feature_disabled_error = group_feature_disabled_error

    async def _prepare_read_target(
        self,
        event: AstrMessageEvent,
        filename: str,
        *,
        allowed_suffixes: frozenset[str] | None = None,
    ) -> tuple["FileReadService.ReadTarget | None", str | None]:
        if not filename:
            return None, "错误：请提供要读取的文件名"

        ok, resolved_path, err = self._workspace_service.pre_check(
            event,
            filename,
            require_exists=True,
            allowed_suffixes=allowed_suffixes,
            allow_external_path=self._allow_external_input_files,
            is_group_feature_enabled=self._is_group_feature_enabled,
            check_permission_fn=self._check_permission,
            group_feature_disabled_error=self._group_feature_disabled_error,
        )
        if not ok:
            return None, err or "错误：未知错误"

        if resolved_path is None:
            return None, "错误：文件路径解析失败"

        display_name = self._workspace_service.display_name(resolved_path)
        try:
            file_size = (await asyncio.to_thread(resolved_path.stat)).st_size
        except OSError as exc:
            logger.error(f"获取文件状态失败: {exc}")
            return None, f"错误：无法读取文件信息 ({display_name})"

        max_size = self._workspace_service.get_max_file_size()
        if file_size > max_size:
            size_str = format_file_size(file_size)
            max_str = format_file_size(max_size)
            return None, f"错误：文件大小 {size_str} 超过限制 {max_str}"

        return (
            self.ReadTarget(
                resolved_path=resolved_path,
                display_name=display_name,
                file_size=file_size,
                max_size=max_size,
                suffix=resolved_path.suffix.lower(),
            ),
            None,
        )

    async def iter_read_file_tool_results(
        self,
        event: AstrMessageEvent,
        filename: str = "",
    ) -> AsyncGenerator[str | mcp.types.CallToolResult, None]:
        target, err = await self._prepare_read_target(event, filename)
        if err:
            yield err
            return

        try:
            if target is None:
                yield "错误：文件路径解析失败"
                return

            resolved_path = target.resolved_path
            suffix = target.suffix
            if suffix in TEXT_SUFFIXES:
                try:
                    content = await self._workspace_service.read_text_file(
                        resolved_path, target.max_size
                    )
                    yield (
                        f"[文件: {target.display_name}, 大小: {format_file_size(target.file_size)}]\n"
                        f"{content}"
                    )
                    return
                except Exception as exc:
                    logger.error(f"读取文件失败: {exc}")
                    yield f"错误：{safe_error_message(exc, '读取文件失败')}"
                    return

            office_type = SUFFIX_TO_OFFICE_TYPE.get(suffix)
            if office_type:
                if office_type is OfficeType.WORD:
                    async for result in self._word_read_service.iter_word_results(
                        resolved_path,
                        target.display_name,
                        suffix,
                        target.file_size,
                    ):
                        yield result
                    return
                extracted = await asyncio.to_thread(
                    self._workspace_service.extract_office_text,
                    resolved_path,
                    office_type,
                )
                if extracted:
                    yield self._workspace_service.format_file_result(
                        target.display_name,
                        suffix,
                        target.file_size,
                        extracted,
                    )
                    return
                yield (
                    f"错误：文件 '{target.display_name}' 无法读取，可能未安装对应解析库"
                )
                return

            if suffix == PDF_SUFFIX:
                extracted = await asyncio.to_thread(
                    self._workspace_service.extract_pdf_text,
                    resolved_path,
                )
                if extracted:
                    yield self._workspace_service.format_file_result(
                        target.display_name,
                        suffix,
                        target.file_size,
                        extracted,
                    )
                    return
                yield (
                    f"错误：无法从 PDF 文件 '{target.display_name}' 中提取文本内容，"
                    "文件可能为空、已损坏或只包含图片。"
                )
                return

            yield f"错误：不支持读取 '{suffix}' 格式的文件"
        except Exception as exc:
            logger.error(f"读取文件失败: {exc}")
            yield f"错误：{safe_error_message(exc, '读取文件失败')}"

    async def iter_read_workbook_tool_results(
        self,
        event: AstrMessageEvent,
        filename: str = "",
    ) -> AsyncGenerator[str | mcp.types.CallToolResult, None]:
        target, err = await self._prepare_read_target(
            event,
            filename,
            allowed_suffixes=self._EXCEL_SUFFIXES,
        )
        if err:
            yield err
            return
        if target is None:
            yield "错误：文件路径解析失败"
            return

        try:
            extracted_sheets = await asyncio.to_thread(
                extract_excel_sheets,
                target.resolved_path,
            )
            if not extracted_sheets:
                yield (
                    f"错误：文件 '{target.display_name}' 无法读取，可能未安装对应解析库"
                )
                return

            sheet_names = ", ".join(sheet.name for sheet in extracted_sheets) or "无"
            sheet_sections = [
                f"[Sheet: {sheet.name}]\n{sheet.text or '[空表]'}"
                for sheet in extracted_sheets
            ]
            yield (
                f"[文件信息] 文件名: {target.display_name}, 类型: {target.suffix}, "
                f"大小: {format_file_size(target.file_size)}\n"
                f"[Sheet 列表] {sheet_names}\n"
                + "\n\n".join(sheet_sections)
            )
        except Exception as exc:
            logger.error(f"读取工作簿失败: {exc}")
            yield f"错误：{safe_error_message(exc, '读取工作簿失败')}"

    async def read_file(
        self,
        event: AstrMessageEvent,
        filename: str = "",
    ) -> str | None:
        text_parts: list[str] = []
        async for result in self.iter_read_file_tool_results(event, filename):
            if isinstance(result, str):
                text_parts.append(result)
        if not text_parts:
            return None
        return "\n".join(text_parts)

    async def read_workbook(
        self,
        event: AstrMessageEvent,
        filename: str = "",
    ) -> str | None:
        text_parts: list[str] = []
        async for result in self.iter_read_workbook_tool_results(event, filename):
            if isinstance(result, str):
                text_parts.append(result)
        if not text_parts:
            return None
        return "\n".join(text_parts)
