from dataclasses import dataclass
from pathlib import Path
from typing import Literal

from astrbot.api import logger
from astrbot.api.event import AstrMessageEvent


@dataclass(slots=True)
class GeneratedFileDeliveryResult:
    status: Literal["sent", "missing", "oversized"]
    file_size: int = 0
    max_size: int = 0


class GeneratedFileDeliveryService:
    def __init__(self, *, workspace_service, delivery_service) -> None:
        self._workspace_service = workspace_service
        self._delivery_service = delivery_service

    async def deliver_generated_file(
        self,
        event: AstrMessageEvent,
        output_path: Path | None,
        *,
        success_message: str | None = None,
    ) -> GeneratedFileDeliveryResult:
        if output_path is None or not output_path.exists():
            logger.info(
                "[文件管理] 生成文件不存在，跳过发送: %s",
                str(output_path) if output_path is not None else "<none>",
            )
            return GeneratedFileDeliveryResult(status="missing")

        file_size = output_path.stat().st_size
        max_size = self._workspace_service.get_max_file_size()
        if file_size > max_size:
            try:
                output_path.unlink(missing_ok=True)
            except Exception as exc:
                logger.warning(f"[文件管理] 删除超限生成文件失败: {exc}")
            return GeneratedFileDeliveryResult(
                status="oversized",
                file_size=file_size,
                max_size=max_size,
            )

        if success_message is None:
            await self._delivery_service.send_file_with_preview(event, output_path)
        else:
            await self._delivery_service.send_file_with_preview(
                event,
                output_path,
                success_message,
            )
        return GeneratedFileDeliveryResult(
            status="sent",
            file_size=file_size,
            max_size=max_size,
        )
