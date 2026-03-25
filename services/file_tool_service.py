import base64
import mimetypes
import warnings
from collections.abc import AsyncGenerator
from pathlib import Path

import mcp

from astrbot.api import logger
from astrbot.api.event import AstrMessageEvent

from ..constants import (
    CONVERTIBLE_TO_PDF,
    DEFAULT_MAX_INLINE_DOCX_IMAGE_COUNT,
    DEFAULT_MAX_INLINE_DOCX_IMAGE_MB,
    EXPLICIT_FILE_TOOL_EVENT_KEY,
    OFFICE_LIBS,
    OFFICE_TYPE_MAP,
    PDF_SUFFIX,
    PDF_TARGET_FORMATS,
    SUFFIX_TO_OFFICE_TYPE,
    TEXT_SUFFIXES,
    OfficeType,
)
from ..utils import (
    WORD_ITEM_IMAGE,
    WORD_ITEM_TEXT,
    format_file_size,
    safe_error_message,
)


class FileToolService:
    def __init__(
        self,
        *,
        workspace_service,
        office_generator,
        pdf_converter,
        delivery_service,
        office_libs: dict,
        allow_external_input_files: bool,
        enable_docx_image_review: bool = True,
        max_inline_docx_image_bytes: int = DEFAULT_MAX_INLINE_DOCX_IMAGE_MB
        * 1024
        * 1024,
        max_inline_docx_image_count: int = DEFAULT_MAX_INLINE_DOCX_IMAGE_COUNT,
        is_group_feature_enabled,
        check_permission,
        group_feature_disabled_error,
    ) -> None:
        self._workspace_service = workspace_service
        self._office_generator = office_generator
        self._pdf_converter = pdf_converter
        self._delivery_service = delivery_service
        self._office_libs = office_libs
        self._allow_external_input_files = allow_external_input_files
        self._enable_docx_image_review = bool(enable_docx_image_review)
        self._max_inline_docx_image_bytes = max(0, int(max_inline_docx_image_bytes))
        self._max_inline_docx_image_count = max(0, int(max_inline_docx_image_count))
        self._is_group_feature_enabled = is_group_feature_enabled
        self._check_permission = check_permission
        self._group_feature_disabled_error = group_feature_disabled_error

    def _is_explicit_tool_locked(
        self,
        event: AstrMessageEvent,
        tool_name: str,
    ) -> bool:
        get_extra = getattr(event, "get_extra", None)
        if not callable(get_extra):
            return False
        try:
            explicit_tool_name = get_extra(EXPLICIT_FILE_TOOL_EVENT_KEY)
        except TypeError:
            return False
        return isinstance(explicit_tool_name, str) and explicit_tool_name == tool_name

    def _finalize_create_office_file_error(
        self,
        event: AstrMessageEvent,
        message: str,
    ) -> str | None:
        if self._is_explicit_tool_locked(event, "create_office_file"):
            return event.plain_result(message)
        return message

    def _build_image_tool_result(
        self,
        image_paths: list[Path],
    ) -> mcp.types.CallToolResult | None:
        content: list[mcp.types.ImageContent] = []
        for image_path in image_paths:
            try:
                base64_data = base64.b64encode(image_path.read_bytes()).decode("utf-8")
            except Exception as exc:
                logger.warning(
                    "[文件管理] 读取嵌入图片失败: %s",
                    safe_error_message(exc, "读取嵌入图片失败"),
                )
                continue

            mime_type = mimetypes.guess_type(image_path.name)[0] or "image/png"
            content.append(
                mcp.types.ImageContent(
                    type="image",
                    data=base64_data,
                    mimeType=mime_type,
                )
            )

        if not content:
            return None

        return mcp.types.CallToolResult(content=content)

    def _plan_inline_word_images(
        self,
        image_paths: list[Path],
    ) -> dict[int, str]:
        skipped_reasons: dict[int, str] = {}
        selected_count = 0

        logger.info("[文件管理] Word 嵌入图片提取数量: %d", len(image_paths))

        for index, image_path in enumerate(image_paths, start=1):
            try:
                image_size = image_path.stat().st_size
            except OSError as exc:
                skipped_reasons[index] = (
                    f"未注入模型上下文（读取失败：{safe_error_message(exc)}）。"
                )
                continue

            logger.info(
                "[文件管理] Word 嵌入图片%d 实际大小: %s",
                index,
                format_file_size(image_size),
            )

            if selected_count >= self._max_inline_docx_image_count:
                skipped_reasons[index] = (
                    f"未注入模型上下文（超过单文档最多 {self._max_inline_docx_image_count} 张限制）。"
                )
                continue

            if image_size > self._max_inline_docx_image_bytes:
                skipped_reasons[index] = (
                    "未注入模型上下文（文件大小 "
                    f"{format_file_size(image_size)} 超过 "
                    f"{format_file_size(self._max_inline_docx_image_bytes)} 限制）。"
                )
                continue

            selected_count += 1

        return skipped_reasons

    async def _iter_word_file_results(
        self,
        resolved_path: Path,
        display_name: str,
        suffix: str,
        file_size: int,
    ) -> AsyncGenerator[str | mcp.types.CallToolResult, None]:
        extracted = self._workspace_service.extract_word_content(
            resolved_path,
            include_images=self._enable_docx_image_review,
        )
        if extracted is None:
            yield f"错误：文件 '{display_name}' 无法读取，可能未安装对应解析库"
            return

        image_count = getattr(extracted, "image_count", 0) or len(
            getattr(extracted, "image_paths", [])
        )

        if not self._enable_docx_image_review:
            text_chunks: list[str] = []
            for item in getattr(extracted, "items", []):
                if item.type != WORD_ITEM_TEXT:
                    continue
                text = (item.text or "").strip()
                if text:
                    text_chunks.append(text)
            formatted = "\n".join(text_chunks) if text_chunks else None
            if not formatted and extracted.text:
                formatted = extracted.text.strip() or None
            if formatted:
                yield self._workspace_service.format_file_result(
                    display_name, suffix, file_size, formatted
                )
            elif image_count > 0:
                yield self._workspace_service.format_file_result(
                    display_name,
                    suffix,
                    file_size,
                    "该 Word 文档仅包含图片内容，当前未启用图片理解。",
                )
            return

        skipped_image_reasons = self._plan_inline_word_images(extracted.image_paths)
        fallback_image_index = 0
        text_chunks: list[str] = []
        selected_image_paths: list[Path] = []

        if extracted.items:
            for item in extracted.items:
                if item.type == WORD_ITEM_TEXT:
                    text = (item.text or "").strip()
                    if text:
                        text_chunks.append(text)
                    continue

                if item.type != WORD_ITEM_IMAGE or item.image_path is None:
                    continue

                image_index = getattr(item, "image_index", None)
                if image_index is None:
                    fallback_image_index += 1
                    image_index = fallback_image_index

                reason = skipped_image_reasons.get(image_index)
                if reason:
                    text_chunks.append(f"[插图{image_index}]（{reason}）")
                    continue

                text_chunks.append(f"[插图{image_index}]")
                selected_image_paths.append(item.image_path)

            final_text = "\n".join(part.strip() for part in text_chunks if part.strip())
            if final_text:
                yield self._workspace_service.format_file_result(
                    display_name, suffix, file_size, final_text
                )
            image_result = self._build_image_tool_result(selected_image_paths)
            if image_result is not None:
                yield image_result
            return

        formatted = self._workspace_service.format_word_content(extracted)
        if formatted:
            yield self._workspace_service.format_file_result(
                display_name, suffix, file_size, formatted
            )

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

        assert resolved_path is not None
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
                    async for result in self._iter_word_file_results(
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

    async def create_office_file(
        self,
        event: AstrMessageEvent,
        filename: str = "",
        content: str = "",
        file_type: str = "",
    ) -> str | None:
        warnings.warn(
            "create_office_file is deprecated for Word documents. "
            "Use create_document -> add_blocks -> finalize_document -> export_document instead.",
            DeprecationWarning,
            stacklevel=2,
        )

        ok, _, err = self._workspace_service.pre_check(
            event,
            feature_key="enable_office_files",
            is_group_feature_enabled=self._is_group_feature_enabled,
            check_permission_fn=self._check_permission,
            group_feature_disabled_error=self._group_feature_disabled_error,
        )
        if not ok:
            return self._finalize_create_office_file_error(
                event,
                err or "错误：未知错误",
            )

        if not content:
            return self._finalize_create_office_file_error(
                event,
                "错误：请提供 content（文件内容）",
            )

        filename = Path(filename).name if filename else ""
        if not filename:
            return self._finalize_create_office_file_error(
                event,
                "错误：请提供 filename（文件名）",
            )

        allowed_fallback_types = "/".join(
            office_name for office_name in OFFICE_TYPE_MAP if office_name != "word"
        )
        normalized_file_type = str(file_type or "").strip().lower()
        suffix = Path(filename).suffix.lower()
        if suffix in SUFFIX_TO_OFFICE_TYPE:
            office_type = SUFFIX_TO_OFFICE_TYPE[suffix]
        else:
            if not normalized_file_type:
                return self._finalize_create_office_file_error(
                    event,
                    "错误：未指定文件类型。请提供带后缀的文件名，"
                    f"或显式传入 file_type（{allowed_fallback_types}）。",
                )
            if normalized_file_type == "word":
                return self._finalize_create_office_file_error(
                    event,
                    "错误：Word 文档请直接提供 .docx/.doc 文件名，"
                    "或改用 create_document → add_blocks → finalize_document → "
                    "export_document。",
                )
            office_type = OFFICE_TYPE_MAP.get(normalized_file_type)

        if not office_type:
            return self._finalize_create_office_file_error(
                event,
                f"错误：不支持的文件类型 '{normalized_file_type}'。"
                f"允许值：{allowed_fallback_types}",
            )

        module_name = OFFICE_LIBS[office_type][0]
        if not self._office_libs.get(module_name):
            package_name = OFFICE_LIBS[office_type][1]
            return self._finalize_create_office_file_error(
                event,
                f"错误：需要安装 {package_name}",
            )

        file_info = {"type": office_type, "filename": filename, "content": content}
        try:
            output_path = await self._office_generator.generate(
                event, file_info["type"], filename, file_info
            )
            if output_path and output_path.exists():
                file_size = output_path.stat().st_size
                max_size = self._workspace_service.get_max_file_size()
                if file_size > max_size:
                    output_path.unlink()
                    size_str = format_file_size(file_size)
                    max_str = format_file_size(max_size)
                    return self._finalize_create_office_file_error(
                        event,
                        f"错误：文件过大 ({size_str})，超过限制 {max_str}",
                    )

                await self._delivery_service.send_file_with_preview(event, output_path)
                return None
        except Exception as exc:
            return self._finalize_create_office_file_error(
                event,
                f"错误：文件操作异常: {exc}",
            )

        return self._finalize_create_office_file_error(event, "错误：文件生成失败")

    async def convert_to_pdf(
        self,
        event: AstrMessageEvent,
        filename: str = "",
        file_path: str = "",
    ) -> str | None:
        if not filename and file_path:
            filename = file_path

        if not filename:
            return "错误：请提供要转换的 Office 文件名"

        logger.debug(
            "[PDF转换] convert_to_pdf 被调用，filename=%s",
            self._workspace_service.display_name(filename),
        )
        ok, resolved_path, err = self._workspace_service.pre_check(
            event,
            filename,
            feature_key="enable_pdf_conversion",
            require_exists=True,
            allowed_suffixes=CONVERTIBLE_TO_PDF,
            allow_external_path=self._allow_external_input_files,
            is_group_feature_enabled=self._is_group_feature_enabled,
            check_permission_fn=self._check_permission,
            group_feature_disabled_error=self._group_feature_disabled_error,
        )
        if not ok:
            logger.warning(f"[PDF转换] 前置检查失败: {err}")
            return err or "错误：未知错误"

        assert resolved_path is not None
        display_name = self._workspace_service.display_name(resolved_path)
        if not self._pdf_converter.is_available("office_to_pdf"):
            return "错误：Office→PDF 转换不可用，需要安装 LibreOffice"

        try:
            logger.info(f"[PDF转换] 开始转换: {display_name} → PDF")
            output_path = await self._pdf_converter.office_to_pdf(resolved_path)
            if output_path and output_path.exists():
                file_size = output_path.stat().st_size
                max_size = self._workspace_service.get_max_file_size()
                if file_size > max_size:
                    output_path.unlink()
                    return f"错误：生成的 PDF 文件过大 ({format_file_size(file_size)})"

                await self._delivery_service.send_file_with_preview(
                    event, output_path, f"✅ 已将 {display_name} 转换为 PDF"
                )
                return None

            return "错误：PDF 转换失败，请检查文件格式是否正确"
        except Exception as exc:
            logger.error(f"[PDF转换] 转换失败: {exc}", exc_info=True)
            return f"错误：{safe_error_message(exc, '转换失败')}"

    async def convert_from_pdf(
        self,
        event: AstrMessageEvent,
        filename: str = "",
        target_format: str = "word",
        file_id: str = "",
    ) -> str | None:
        if not filename and file_id:
            filename = file_id

        if not filename:
            return "错误：请提供要转换的 PDF 文件名"

        ok, source_path, err = self._workspace_service.pre_check(
            event,
            filename,
            feature_key="enable_pdf_conversion",
            require_exists=True,
            required_suffix=PDF_SUFFIX,
            allow_external_path=self._allow_external_input_files,
            is_group_feature_enabled=self._is_group_feature_enabled,
            check_permission_fn=self._check_permission,
            group_feature_disabled_error=self._group_feature_disabled_error,
        )
        if not ok:
            return err or "错误：未知错误"

        assert source_path is not None
        display_name = self._workspace_service.display_name(source_path)
        target = target_format.lower().strip()
        if target not in PDF_TARGET_FORMATS:
            supported = ", ".join(PDF_TARGET_FORMATS.keys())
            return f"错误：不支持的目标格式 '{target_format}'，可选: {supported}"

        _, target_desc = PDF_TARGET_FORMATS[target]
        conversion_type = f"pdf_to_{target}"
        if not self._pdf_converter.is_available(conversion_type):
            missing = self._pdf_converter.get_missing_dependencies()
            return f"错误：PDF→{target_desc} 转换不可用，缺少依赖: {', '.join(missing)}"

        try:
            logger.info(f"[PDF转换] 开始转换: {display_name} → {target_desc}")
            if target == "word":
                output_path = await self._pdf_converter.pdf_to_word(source_path)
            elif target == "excel":
                output_path = await self._pdf_converter.pdf_to_excel(source_path)
            else:
                return f"错误：未实现的转换类型: {target}"

            if output_path and output_path.exists():
                file_size = output_path.stat().st_size
                max_size = self._workspace_service.get_max_file_size()
                if file_size > max_size:
                    output_path.unlink()
                    return f"错误：生成的文件过大 ({format_file_size(file_size)})"

                await self._delivery_service.send_file_with_preview(
                    event,
                    output_path,
                    f"✅ 已将 {display_name} 转换为 {target_desc}",
                )
                return None

            return f"错误：PDF→{target_desc} 转换失败"
        except Exception as exc:
            logger.error(f"[PDF转换] 转换失败: {exc}", exc_info=True)
            return f"错误：{safe_error_message(exc, '转换失败')}"
