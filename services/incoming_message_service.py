from pathlib import Path

import astrbot.api.message_components as Comp
from astrbot.api import logger

from ..constants import ALL_OFFICE_SUFFIXES, PDF_SUFFIX, TEXT_SUFFIXES
from .image_file_utils import (
    is_supported_image_mime_type,
    is_supported_image_reference,
)


class IncomingMessageService:
    def __init__(
        self,
        *,
        message_buffer,
        remember_recent_text,
        is_group_feature_enabled,
        cache_pending_image_resource,
    ) -> None:
        self._message_buffer = message_buffer
        self._remember_recent_text = remember_recent_text
        self._is_group_feature_enabled = is_group_feature_enabled
        self._cache_pending_image_resource = cache_pending_image_resource

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

    def _is_image_message(self, event) -> bool:
        """检查消息是否只包含图片（Comp.Image 或图片后缀的 Comp.File），无命令文本"""
        has_image = False
        for component in event.message_obj.message:
            if isinstance(component, Comp.Image):
                has_image = True
                continue
            if isinstance(component, Comp.File):
                if self._is_image_file_component(component):
                    has_image = True
                    continue
                return False
            if isinstance(component, Comp.Plain):
                text = component.text.strip()
                if text.startswith("/"):
                    return False
        return has_image

    @staticmethod
    def _is_image_file_component(component: Comp.File) -> bool:
        for attr_name in ("name", "file", "file_", "url"):
            value = getattr(component, attr_name, "") or ""
            if value and is_supported_image_reference(str(value)):
                return True
        for attr_name in ("mime_type", "mimetype", "mime", "content_type"):
            value = getattr(component, attr_name, "") or ""
            if value and is_supported_image_mime_type(str(value)):
                return True
        return False

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
            if self._is_image_message(event):
                pending_resources = []
                for component in event.message_obj.message:
                    if isinstance(component, (Comp.Image, Comp.File)):
                        self._cache_pending_image_resource(event, component)
                        pending_resources.append(component)
                setattr(event, "_has_pending_images", True)
                setattr(event, "_pending_image_resources", pending_resources)
                logger.debug("[文件管理] 图片已缓存为待注册资源，消息正常流向 LLM")
                return
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
