"""
消息缓冲器模块
"""

from __future__ import annotations

import asyncio
from dataclasses import dataclass, field
from typing import TYPE_CHECKING

import astrbot.api.message_components as Comp
from astrbot.api import logger

if TYPE_CHECKING:
    from astrbot.api.event import AstrMessageEvent


@dataclass
class BufferedMessage:
    """缓冲的消息数据"""

    event: AstrMessageEvent  # 原始事件（用于发送响应等）
    files: list[Comp.File] = field(default_factory=list)  # 文件列表
    texts: list[str] = field(default_factory=list)  # 文本内容列表
    timer_task: asyncio.Task | None = None  # 定时器任务


class MessageBuffer:
    """
    消息缓冲器

    按 (平台ID, 用户ID, 会话ID) 分组缓冲消息，
    收到文件消息后等待一段时间以聚合后续消息。
    """

    def __init__(self, wait_seconds: float = 4):
        """
        Args:
            wait_seconds: 缓冲等待时间（秒）
        """
        self._wait_seconds = wait_seconds
        # 缓冲区: key = (platform_id, sender_id, session_id)
        self._buffers: dict[tuple[str, str, str], BufferedMessage] = {}
        # 回调函数，当消息聚合完成时调用
        self._on_complete_callback = None
        # 锁，保证线程安全
        self._lock = asyncio.Lock()

    def set_complete_callback(self, callback):
        """设置消息聚合完成后的回调函数（有文件时）"""
        self._on_complete_callback = callback

    def _get_buffer_key(self, event: AstrMessageEvent) -> tuple[str, str, str]:
        """获取缓冲区的 key"""
        platform_id = event.get_platform_id() or "unknown"
        sender_id = str(event.get_sender_id() or "unknown")
        # 使用 unified_msg_origin 作为会话标识
        session_id = event.unified_msg_origin or "unknown"
        return (platform_id, sender_id, session_id)

    def _extract_components(
        self, event: AstrMessageEvent
    ) -> tuple[list[Comp.File], list[str]]:
        """从消息中提取文件和文本组件"""
        files = []
        texts = []

        for component in event.message_obj.message:
            if isinstance(component, Comp.File):
                files.append(component)
            elif isinstance(component, Comp.Plain):
                text = component.text.strip()
                if text:
                    texts.append(text)
            # 忽略 At、Reply 等其他组件

        return files, texts

    async def add_message(self, event: AstrMessageEvent) -> bool:
        """
        添加消息到缓冲区

        Args:
            event: 消息事件

        Returns:
            True: 消息已被缓冲，调用方应停止事件传播
            False: 消息不需要缓冲（缓冲器已禁用）
        """
        # 如果等待时间为 0，则禁用缓冲
        if self._wait_seconds <= 0:
            return False

        files, texts = self._extract_components(event)
        key = self._get_buffer_key(event)

        async with self._lock:
            # 检查是否已有缓冲
            if key in self._buffers:
                buf = self._buffers[key]
                buf.files.extend(files)
                buf.texts.extend(texts)
                logger.debug(
                    f"[消息缓冲] 追加消息到缓冲区: {key}, "
                    f"文件数: {len(files)}, 文本数: {len(texts)}"
                )
                return True

            # 没有缓冲，开始新的缓冲
            buf = BufferedMessage(
                event=event,
                files=files,
                texts=texts,
            )

            logger.info(f"[消息缓冲] 开始缓冲: {key}, 等待 {self._wait_seconds} 秒")

            buf.timer_task = asyncio.create_task(
                self._wait_and_process(key, self._wait_seconds)
            )
            self._buffers[key] = buf
            return True

    async def _wait_and_process(self, key: tuple[str, str, str], wait_time: float):
        """等待超时后处理缓冲的消息"""
        try:
            await asyncio.sleep(wait_time)
        except asyncio.CancelledError:
            # 定时器被取消，正常退出
            return

        async with self._lock:
            if key not in self._buffers:
                return

            buf = self._buffers.pop(key)

        # 只有文件时才调用回调
        if buf.files and self._on_complete_callback:
            try:
                logger.info(
                    f"[消息缓冲] 缓冲完成，"
                    f"文件数: {len(buf.files)}, 文本数: {len(buf.texts)}"
                )
                await self._on_complete_callback(buf)
            except Exception as e:
                logger.error(f"[消息缓冲] 处理回调时出错: {e}")

    def is_buffering(self, event: AstrMessageEvent) -> bool:
        """检查指定用户是否正在缓冲状态"""
        key = self._get_buffer_key(event)
        return key in self._buffers

    async def cancel_buffer(self, event: AstrMessageEvent):
        """取消指定用户的缓冲"""
        key = self._get_buffer_key(event)
        async with self._lock:
            if key in self._buffers:
                buf = self._buffers.pop(key)
                if buf.timer_task and not buf.timer_task.done():
                    buf.timer_task.cancel()
                logger.debug(f"[消息缓冲] 取消缓冲: {key}")
