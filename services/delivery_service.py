import asyncio
from pathlib import Path

import astrbot.api.message_components as Comp
from astrbot.api import logger
from astrbot.api.event import AstrMessageEvent
from astrbot.core.agent.run_context import ContextWrapper
from astrbot.core.astr_agent_context import AstrAgentContext
from astrbot.core.message.message_event_result import MessageChain


class DeliveryService:
    def __init__(
        self,
        *,
        executor,
        preview_generator,
        enable_preview: bool,
        auto_delete: bool,
        reply_to_user: bool,
    ) -> None:
        self._executor = executor
        self._preview_generator = preview_generator
        self._enable_preview = enable_preview
        self._auto_delete = auto_delete
        self._reply_to_user = reply_to_user

    async def send_file_with_preview(
        self,
        event: AstrMessageEvent,
        file_path: Path,
        success_message: str = "✅ 文件已处理成功",
    ) -> None:
        preview_path = None

        if self._enable_preview:
            try:
                loop = asyncio.get_running_loop()
                preview_path = await loop.run_in_executor(
                    self._executor,
                    self._preview_generator.generate_preview,
                    file_path,
                    None,
                )
            except Exception as exc:
                logger.warning(f"[预览生成] 生成预览图失败: {exc}")
                preview_path = None

        text_chain = MessageChain()
        text_chain.message(f"{success_message}：{file_path.name}")
        if self._reply_to_user:
            text_chain.chain.append(Comp.At(qq=event.get_sender_id()))
        await event.send(text_chain)

        if preview_path and preview_path.exists():
            await event.send(
                MessageChain([Comp.Image(file=str(preview_path.resolve()))])
            )
            if self._auto_delete:
                try:
                    preview_path.unlink()
                except Exception as exc:
                    logger.warning(f"[文件管理] 自动删除预览文件失败: {exc}")

        await event.send(
            MessageChain([Comp.File(file=str(file_path.resolve()), name=file_path.name)])
        )

        if self._auto_delete and file_path.exists():
            try:
                file_path.unlink()
            except Exception as exc:
                logger.warning(f"[文件管理] 自动删除文件失败: {exc}")

    async def send_exported_document(
        self,
        event: AstrMessageEvent,
        file_path: Path,
        exported_message: str,
    ) -> str:
        if not file_path.exists():
            return f"Document exported to {file_path}, but the file does not exist."

        await self.send_file_with_preview(event, file_path, exported_message)
        return f"Document exported and sent to the user: {file_path.name}"

    async def handle_exported_document_tool(
        self,
        context: ContextWrapper[AstrAgentContext],
        output_path: str,
        exported_message: str,
    ) -> str | None:
        file_path = Path(output_path)
        if not file_path.exists():
            return f"Document exported to {output_path}, but the file does not exist."

        return await self.send_exported_document(
            context.context.event,
            file_path,
            exported_message,
        )
