from astrbot.api.event import AstrMessageEvent

from .generated_file_delivery_service import GeneratedFileDeliveryResult
from ..utils import format_file_size


class FileDeliveryService:
    def __init__(self, *, generated_file_delivery_service) -> None:
        self._generated_file_delivery_service = generated_file_delivery_service

    @staticmethod
    def _format_generated_file_delivery_error(
        delivery_result,
        *,
        missing_message: str,
        oversized_template: str,
    ) -> str | None:
        if delivery_result.status == "missing":
            return missing_message
        if delivery_result.status == "oversized":
            return oversized_template.format(
                file_size=format_file_size(delivery_result.file_size),
                max_size=format_file_size(delivery_result.max_size),
            )
        if delivery_result.status == "invalid":
            errors = delivery_result.validation_errors or []
            details = "；".join(errors)
            return f"错误：生成的 Excel 文件存在明显公式风险：{details}"
        return None

    async def deliver_generated_file(
        self,
        event: AstrMessageEvent,
        output_path,
        *,
        missing_message: str,
        oversized_template: str,
        success_message: str | None = None,
    ) -> str | None:
        delivery_error, _ = await self.deliver_generated_file_with_result(
            event,
            output_path,
            missing_message=missing_message,
            oversized_template=oversized_template,
            success_message=success_message,
        )
        return delivery_error

    async def deliver_generated_file_with_result(
        self,
        event: AstrMessageEvent,
        output_path,
        *,
        missing_message: str,
        oversized_template: str,
        success_message: str | None = None,
    ) -> tuple[str | None, GeneratedFileDeliveryResult]:
        delivery_result = (
            await self._generated_file_delivery_service.deliver_generated_file(
                event,
                output_path,
                success_message=success_message,
            )
        )
        delivery_error = self._format_generated_file_delivery_error(
            delivery_result,
            missing_message=missing_message,
            oversized_template=oversized_template,
        )
        return delivery_error, delivery_result
