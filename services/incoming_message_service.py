from pathlib import Path

import astrbot.api.message_components as Comp
from astrbot.api import logger

from ..constants import ALL_OFFICE_SUFFIXES, PDF_SUFFIX, TEXT_SUFFIXES


class IncomingMessageService:
    def __init__(
        self,
        *,
        message_buffer,
        remember_recent_text,
        is_group_feature_enabled,
    ) -> None:
        self._message_buffer = message_buffer
        self._remember_recent_text = remember_recent_text
        self._is_group_feature_enabled = is_group_feature_enabled

    def _has_follow_up_text(self, event) -> bool:
        has_plain_text = False
        for component in event.message_obj.message:
            if isinstance(component, Comp.File):
                return False
            if isinstance(component, Comp.Plain):
                text = component.text.strip()
                if not text:
                    continue
                if text.startswith("/"):
                    return False
                has_plain_text = True
        return has_plain_text

    async def handle_file_message(self, event) -> None:
        if getattr(event, "_buffered", False):
            return

        if not self._is_group_feature_enabled(event):
            return

        if not event.message_obj.message:
            return

        has_supported_file = False
        for component in event.message_obj.message:
            if isinstance(component, Comp.File):
                name = component.name or ""
                suffix = Path(name).suffix.lower() if name else ""
                if (
                    suffix in ALL_OFFICE_SUFFIXES
                    or suffix in TEXT_SUFFIXES
                    or suffix == PDF_SUFFIX
                ):
                    has_supported_file = True
                    break

        if not has_supported_file:
            if self._message_buffer.is_buffering(event):
                if not self._has_follow_up_text(event):
                    return
                buffered = await self._message_buffer.add_message(event)
                if buffered:
                    event.stop_event()
                    logger.debug("[文件管理] 已将后续文本追加到文件缓冲区...")
                return
            self._remember_recent_text(event)
            return

        buffered = await self._message_buffer.add_message(event)
        if buffered:
            event.stop_event()
            logger.debug("[文件管理] 支持的文件已缓冲，等待聚合...")
