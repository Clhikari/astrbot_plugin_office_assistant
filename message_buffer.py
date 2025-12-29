"""
消息缓冲器模块

解决 QQ 平台文件和文本消息分离的问题：
- PC 端：文件和消息可能连发
- 手机端：只能单发文件，不能连带消息

核心思路（短观察期 + 文件触发延长缓冲）：
1. 收到任何消息时，都开始一个短的"观察期"（默认 0.8 秒）
2. 如果在观察期内收到文件，则延长等待时间到完整缓冲期（默认 2.5 秒）
3. 如果观察期结束时没有文件，则立即处理（不额外延迟）
4. 这样可以同时处理"先文件后文本"和"先文本后文件"两种场景
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
    has_file: bool = False  # 是否包含文件（决定等待时间）
    extended: bool = False  # 是否已延长到完整缓冲期


class MessageBuffer:
    """
    消息缓冲器

    按 (平台ID, 用户ID, 会话ID) 分组缓冲消息，
    使用两阶段等待策略：短观察期 + 文件触发延长。
    """

    def __init__(
        self,
        wait_seconds: float = 2.5,
        observe_seconds: float = 0.8,
    ):
        """
        Args:
            wait_seconds: 完整等待时间（秒），有文件时使用
            observe_seconds: 观察期时间（秒），无文件时的短等待
        """
        self._wait_seconds = wait_seconds
        self._observe_seconds = observe_seconds
        # 缓冲区: key = (platform_id, sender_id, session_id)
        self._buffers: dict[tuple[str, str, str], BufferedMessage] = {}
        # 回调函数，当消息聚合完成时调用
        self._on_complete_callback = None
        # 无文件时的回调（直接放行，不触发文件处理逻辑）
        self._on_passthrough_callback = None
        # 锁，保证线程安全
        self._lock = asyncio.Lock()

    def set_complete_callback(self, callback):
        """设置消息聚合完成后的回调函数（有文件时）"""
        self._on_complete_callback = callback

    def set_passthrough_callback(self, callback):
        """设置无文件消息放行的回调函数"""
        self._on_passthrough_callback = callback

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
        # 如果观察期和等待期都为 0，则禁用缓冲
        if self._observe_seconds <= 0 and self._wait_seconds <= 0:
            return False

        files, texts = self._extract_components(event)
        key = self._get_buffer_key(event)
        has_file = len(files) > 0

        async with self._lock:
            # 检查是否已有缓冲
            if key in self._buffers:
                buf = self._buffers[key]
                buf.files.extend(files)
                buf.texts.extend(texts)

                # 如果新消息包含文件，且还没延长过，则延长等待时间
                if has_file and not buf.extended:
                    buf.has_file = True
                    buf.extended = True
                    # 取消旧的定时器
                    if buf.timer_task and not buf.timer_task.done():
                        buf.timer_task.cancel()
                    # 使用完整等待时间重新启动定时器
                    buf.timer_task = asyncio.create_task(
                        self._wait_and_process(key, self._wait_seconds)
                    )
                    logger.info(
                        f"[消息缓冲] 检测到文件，延长等待至 {self._wait_seconds} 秒: {key}"
                    )
                else:
                    # 没有新文件，不改变定时器
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
                has_file=has_file,
                extended=has_file,  # 如果一开始就有文件，直接标记为已延长
            )

            # 决定等待时间
            if has_file:
                wait_time = self._wait_seconds
                logger.info(
                    f"[消息缓冲] 收到文件，开始完整缓冲: {key}, 等待 {wait_time} 秒"
                )
            else:
                wait_time = self._observe_seconds
                logger.info(f"[消息缓冲] 开始观察期: {key}, 等待 {wait_time} 秒")

            buf.timer_task = asyncio.create_task(self._wait_and_process(key, wait_time))
            self._buffers[key] = buf
            return True

    async def _wait_and_process(self, key: tuple[str, str, str], wait_time: float):
        """等待超时后处理缓冲的消息"""
        try:
            await asyncio.sleep(wait_time)
        except asyncio.CancelledError:
            # 定时器被取消（等待时间被延长），正常退出
            return

        async with self._lock:
            if key not in self._buffers:
                return

            buf = self._buffers.pop(key)

        # 根据是否有文件决定调用哪个回调
        if buf.has_file:
            # 有文件，调用文件处理回调
            if self._on_complete_callback:
                try:
                    logger.info(
                        f"[消息缓冲] 缓冲完成（有文件），"
                        f"文件数: {len(buf.files)}, 文本数: {len(buf.texts)}"
                    )
                    await self._on_complete_callback(buf)
                except Exception as e:
                    logger.error(f"[消息缓冲] 处理回调时出错: {e}")
        else:
            # 无文件，调用放行回调
            if self._on_passthrough_callback:
                try:
                    logger.debug("[消息缓冲] 观察期结束（无文件），放行消息")
                    await self._on_passthrough_callback(buf)
                except Exception as e:
                    logger.error(f"[消息缓冲] 放行回调时出错: {e}")

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
