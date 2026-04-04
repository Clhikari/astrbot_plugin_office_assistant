from collections.abc import AsyncGenerator

import mcp

from astrbot.api import logger
from astrbot.api.event import AstrMessageEvent

from ..constants import PDF_SUFFIX, SUFFIX_TO_OFFICE_TYPE, TEXT_SUFFIXES, OfficeType
from ..utils import format_file_size, safe_error_message


class FileReadService:
    def __init__(
        self,
        *,
        workspace_service,
        word_read_service,
        allow_external_input_files: bool,
        is_group_feature_enabled,
        check_permission,
        group_feature_disabled_error,
    ) -> None:
        self._workspace_service = workspace_service
        self._word_read_service = word_read_service
        self._allow_external_input_files = allow_external_input_files
        self._is_group_feature_enabled = is_group_feature_enabled
        self._check_permission = check_permission
        self._group_feature_disabled_error = group_feature_disabled_error

    async def iter_read_file_tool_results(
        self,
        event: AstrMessageEvent,
        filename: str = "",
    ) -> AsyncGenerator[str | mcp.types.CallToolResult, None]:
        if not filename:
            yield "错误：请提供要读取的文件名"
            return

        ok, resolved_path, err = self._workspace_service.pre_check(
            event,
            filename,
            require_exists=True,
            allow_external_path=self._allow_external_input_files,
            is_group_feature_enabled=self._is_group_feature_enabled,
            check_permission_fn=self._check_permission,
            group_feature_disabled_error=self._group_feature_disabled_error,
        )
        if not ok:
            yield err or "错误：未知错误"
            return

        if resolved_path is None:
            yield "错误：文件路径解析失败"
            return
        display_name = self._workspace_service.display_name(resolved_path)
        file_size = resolved_path.stat().st_size
        max_size = self._workspace_service.get_max_file_size()
        if file_size > max_size:
            size_str = format_file_size(file_size)
            max_str = format_file_size(max_size)
            yield f"错误：文件大小 {size_str} 超过限制 {max_str}"
            return

        try:
            suffix = resolved_path.suffix.lower()
            if suffix in TEXT_SUFFIXES:
                try:
                    content = await self._workspace_service.read_text_file(
                        resolved_path, max_size
                    )
                    yield (
                        f"[文件: {display_name}, 大小: {format_file_size(file_size)}]\n"
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
                        display_name,
                        suffix,
                        file_size,
                    ):
                        yield result
                    return
                extracted = self._workspace_service.extract_office_text(
                    resolved_path, office_type
                )
                if extracted:
                    yield self._workspace_service.format_file_result(
                        display_name, suffix, file_size, extracted
                    )
                    return
                yield f"错误：文件 '{display_name}' 无法读取，可能未安装对应解析库"
                return

            if suffix == PDF_SUFFIX:
                extracted = self._workspace_service.extract_pdf_text(resolved_path)
                if extracted:
                    yield self._workspace_service.format_file_result(
                        display_name, suffix, file_size, extracted
                    )
                    return
                yield (
                    f"错误：无法从 PDF 文件 '{display_name}' 中提取文本内容，"
                    "文件可能为空、已损坏或只包含图片。"
                )
                return

            yield f"错误：不支持读取 '{suffix}' 格式的文件"
        except Exception as exc:
            logger.error(f"读取文件失败: {exc}")
            yield f"错误：{safe_error_message(exc, '读取文件失败')}"

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
