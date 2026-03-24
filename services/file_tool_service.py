import warnings
from pathlib import Path

from astrbot.api import logger
from astrbot.api.event import AstrMessageEvent

from ..constants import (
    CONVERTIBLE_TO_PDF,
    OFFICE_LIBS,
    OFFICE_TYPE_MAP,
    PDF_SUFFIX,
    PDF_TARGET_FORMATS,
    SUFFIX_TO_OFFICE_TYPE,
    TEXT_SUFFIXES,
)
from ..utils import format_file_size, safe_error_message


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
        self._is_group_feature_enabled = is_group_feature_enabled
        self._check_permission = check_permission
        self._group_feature_disabled_error = group_feature_disabled_error

    async def read_file(
        self,
        event: AstrMessageEvent,
        filename: str = "",
    ) -> str | None:
        if not filename:
            return "错误：请提供要读取的文件名"

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
            return err or "错误：未知错误"

        assert resolved_path is not None
        display_name = self._workspace_service.display_name(resolved_path)
        file_size = resolved_path.stat().st_size
        max_size = self._workspace_service.get_max_file_size()
        if file_size > max_size:
            size_str = format_file_size(file_size)
            max_str = format_file_size(max_size)
            return f"错误：文件大小 {size_str} 超过限制 {max_str}"

        try:
            suffix = resolved_path.suffix.lower()
            if suffix in TEXT_SUFFIXES:
                try:
                    content = await self._workspace_service.read_text_file(
                        resolved_path, max_size
                    )
                    return (
                        f"[文件: {display_name}, 大小: {format_file_size(file_size)}]\n"
                        f"{content}"
                    )
                except Exception as exc:
                    logger.error(f"读取文件失败: {exc}")
                    return f"错误：{safe_error_message(exc, '读取文件失败')}"

            office_type = SUFFIX_TO_OFFICE_TYPE.get(suffix)
            if office_type:
                extracted = self._workspace_service.extract_office_text(
                    resolved_path, office_type
                )
                if extracted:
                    return self._workspace_service.format_file_result(
                        display_name, suffix, file_size, extracted
                    )
                return f"错误：文件 '{display_name}' 无法读取，可能未安装对应解析库"

            if suffix == PDF_SUFFIX:
                extracted = self._workspace_service.extract_pdf_text(resolved_path)
                if extracted:
                    return self._workspace_service.format_file_result(
                        display_name, suffix, file_size, extracted
                    )
                return (
                    f"错误：无法从 PDF 文件 '{display_name}' 中提取文本内容，"
                    "文件可能为空、已损坏或只包含图片。"
                )

            return f"错误：不支持读取 '{suffix}' 格式的文件"
        except Exception as exc:
            logger.error(f"读取文件失败: {exc}")
            return f"错误：{safe_error_message(exc, '读取文件失败')}"

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
            return err or "错误：未知错误"

        if not content:
            return "错误：请提供 content（文件内容）"

        filename = Path(filename).name if filename else ""
        if not filename:
            return "错误：请提供 filename（文件名）"

        allowed_fallback_types = "/".join(
            office_name for office_name in OFFICE_TYPE_MAP if office_name != "word"
        )
        normalized_file_type = str(file_type or "").strip().lower()
        suffix = Path(filename).suffix.lower()
        if suffix in SUFFIX_TO_OFFICE_TYPE:
            office_type = SUFFIX_TO_OFFICE_TYPE[suffix]
        else:
            if not normalized_file_type:
                return (
                    "错误：未指定文件类型。请提供带后缀的文件名，"
                    f"或显式传入 file_type（{allowed_fallback_types}）。"
                )
            if normalized_file_type == "word":
                return (
                    "错误：Word 文档请直接提供 .docx/.doc 文件名，"
                    "或改用 create_document → add_blocks → finalize_document → "
                    "export_document。"
                )
            office_type = OFFICE_TYPE_MAP.get(normalized_file_type)

        if not office_type:
            return (
                f"错误：不支持的文件类型 '{normalized_file_type}'。"
                f"允许值：{allowed_fallback_types}"
            )

        module_name = OFFICE_LIBS[office_type][0]
        if not self._office_libs.get(module_name):
            package_name = OFFICE_LIBS[office_type][1]
            return f"错误：需要安装 {package_name}"

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
                    return f"错误：文件过大 ({size_str})，超过限制 {max_str}"

                await self._delivery_service.send_file_with_preview(event, output_path)
                return None
        except Exception as exc:
            return f"错误：文件操作异常: {exc}"

        return "错误：文件生成失败"

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
