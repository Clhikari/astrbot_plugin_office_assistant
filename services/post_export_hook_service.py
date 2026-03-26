import asyncio
from pathlib import Path

import astrbot.api.message_components as Comp
from astrbot.api import logger
from astrbot.api.event import AstrMessageEvent
from astrbot.core.agent.run_context import ContextWrapper
from astrbot.core.astr_agent_context import AstrAgentContext
from astrbot.core.message.message_event_result import MessageChain


class PostExportHookService:
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
        self._executor = executor
        self._preview_generator = preview_generator
        self._enable_preview = enable_preview
        self._auto_delete = auto_delete
        self._reply_to_user = reply_to_user
        self._exported_message = exported_message

    async def _generate_preview(self, file_path: Path) -> Path | None:
        if not self._enable_preview:
            return None

        try:
            loop = asyncio.get_running_loop()
            return await loop.run_in_executor(
                self._executor,
                self._preview_generator.generate_preview,
                file_path,
                None,
            )
        except Exception as exc:
            logger.warning(f"[预览生成] 生成预览图失败: {exc}")
            return None

    async def _send_success_message(
        self,
        event: AstrMessageEvent,
        file_path: Path,
    ) -> None:
        text_chain = MessageChain()
        text_chain.message(f"{self._exported_message}：{file_path.name}")
        if self._reply_to_user:
            text_chain.chain.append(Comp.At(qq=event.get_sender_id()))
        await event.send(text_chain)

    async def _send_preview_if_available(
        self,
        event: AstrMessageEvent,
        preview_path: Path | None,
    ) -> None:
        if not preview_path or not preview_path.exists():
            return

        await event.send(MessageChain([Comp.Image(file=str(preview_path.resolve()))]))
        await self._delete_file_if_needed(preview_path, "预览文件")

    async def _send_output_file(
        self,
        event: AstrMessageEvent,
        file_path: Path,
    ) -> None:
        await event.send(
            MessageChain(
                [Comp.File(file=str(file_path.resolve()), name=file_path.name)]
            )
        )

    async def _delete_file_if_needed(self, file_path: Path, label: str) -> None:
        if not self._auto_delete or not file_path.exists():
            return

        try:
            file_path.unlink()
        except Exception as exc:
            logger.warning(f"[文件管理] 自动删除{label}失败: {exc}")

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

        preview_path = await self._generate_preview(file_path)
        await self._send_success_message(event, file_path)
        await self._send_preview_if_available(event, preview_path)
        await self._send_output_file(event, file_path)
        await self._delete_file_if_needed(file_path, "文件")
        return f"文档已导出并发送给用户：{file_path.name}"

    async def handle_exported_document_tool(
        self,
        context: ContextWrapper[AstrAgentContext],
        output_path: str,
    ) -> str | None:
        return await self.send_exported_document(
            context.context.event,
            Path(output_path),
        )
