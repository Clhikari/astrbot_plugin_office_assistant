from astrbot.api.event import AstrMessageEvent

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
        delivery_result = (
            await self._generated_file_delivery_service.deliver_generated_file(
                event,
                output_path,
                success_message=success_message,
            )
        )
        return self._format_generated_file_delivery_error(
            delivery_result,
            missing_message=missing_message,
            oversized_template=oversized_template,
        )
