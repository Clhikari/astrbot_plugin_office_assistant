from astrbot.api import logger
from astrbot.api.event import AstrMessageEvent

from ..constants import CONVERTIBLE_TO_PDF, PDF_SUFFIX, PDF_TARGET_FORMATS
from ..utils import safe_error_message


class PdfConvertService:
    def __init__(
        self,
        *,
        workspace_service,
        pdf_converter,
        file_delivery_service,
        allow_external_input_files: bool,
        is_group_feature_enabled,
        check_permission,
        group_feature_disabled_error,
    ) -> None:
        self._workspace_service = workspace_service
        self._pdf_converter = pdf_converter
        self._file_delivery_service = file_delivery_service
        self._allow_external_input_files = allow_external_input_files
        self._is_group_feature_enabled = is_group_feature_enabled
        self._check_permission = check_permission
        self._group_feature_disabled_error = group_feature_disabled_error

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
            delivery_error = await self._file_delivery_service.deliver_generated_file(
                event,
                output_path,
                success_message=f"✅ 已将 {display_name} 转换为 PDF",
                missing_message="错误：PDF 转换失败，未找到生成的 PDF 文件",
                oversized_template="错误：生成的 PDF 文件过大 ({file_size})",
            )
            if delivery_error:
                return delivery_error
            return None
        except Exception as exc:
            logger.exception(f"[PDF转换] 转换失败: {exc}")
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

            delivery_error = await self._file_delivery_service.deliver_generated_file(
                event,
                output_path,
                success_message=f"✅ 已将 {display_name} 转换为 {target_desc}",
                missing_message=f"错误：PDF→{target_desc} 转换失败，未找到生成的文件",
                oversized_template="错误：生成的文件过大 ({file_size})",
            )
            if delivery_error:
                return delivery_error
            return None
        except Exception as exc:
            logger.exception(f"[PDF转换] 转换失败: {exc}")
            return f"错误：{safe_error_message(exc, '转换失败')}"
