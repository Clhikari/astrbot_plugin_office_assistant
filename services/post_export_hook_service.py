from pathlib import Path

from astrbot.api import logger  # noqa: F401 - compatibility for existing log hooks
from astrbot.api.event import AstrMessageEvent
from astrbot.core.agent.run_context import ContextWrapper
from astrbot.core.astr_agent_context import AstrAgentContext

from .delivery_service import DeliveryService


class PostExportHookService(DeliveryService):
    """Adapt structured exports to the shared file delivery implementation."""

    def __init__(
        self,
        *,
        executor,
        preview_generator,
        enable_preview: bool,
        auto_delete: bool,
        reply_to_user: bool,
        exported_message: str,
    ) -> None:
        super().__init__(
            executor=executor,
            preview_generator=preview_generator,
            enable_preview=enable_preview,
            auto_delete=auto_delete,
            reply_to_user=reply_to_user,
        )
        self._exported_message = exported_message

    @staticmethod
    def _missing_export_message(file_path: Path) -> str:
        return f"文档已导出，但文件“{file_path.name}”不存在。"

    async def send_exported_document(
        self,
        event: AstrMessageEvent,
        file_path: Path,
    ) -> str:
        if not file_path.exists():
            return self._missing_export_message(file_path)
        await self.send_file_with_preview(event, file_path, self._exported_message)
        return f"文档已导出并发送给用户：{file_path.name}"

    async def handle_exported_document_tool(
        self,
        context: ContextWrapper[AstrAgentContext],
        output_path: str,
    ) -> str | None:
        file_path = Path(output_path)
        if not file_path.is_file():
            raise FileNotFoundError(self._missing_export_message(file_path))
        await self.send_file_with_preview(
            context.context.event, file_path, self._exported_message
        )
        return f"文档已导出并发送给用户：{file_path.name}"
