import asyncio
import importlib
import shutil
import tempfile
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import pdfplumber

import astrbot.api.message_components as Comp
from astrbot.api import AstrBotConfig, llm_tool, logger
from astrbot.api.event import AstrMessageEvent, filter
from astrbot.api.star import Context, Star, StarTools
from astrbot.core.message.components import At, Reply
from astrbot.core.message.message_event_result import MessageChain
from astrbot.core.platform.message_type import MessageType
from astrbot.core.provider.entities import ProviderRequest

from .constants import (
    ALL_OFFICE_SUFFIXES,
    CONVERTIBLE_TO_PDF,
    DEFAULT_CHUNK_SIZE,
    DEFAULT_MAX_FILE_SIZE_MB,
    FILE_TOOLS,
    OFFICE_LIBS,
    OFFICE_TYPE_MAP,
    EXECUTION_TOOLS,
    PDF_SUFFIX,
    PDF_TARGET_FORMATS,
    SUFFIX_TO_OFFICE_TYPE,
    TEXT_SUFFIXES,
    OfficeType,
)
from .message_buffer import BufferedMessage, MessageBuffer
from .office_generator import OfficeGenerator
from .pdf_converter import PDFConverter
from .preview_generator import PreviewGenerator
from .utils import (
    extract_excel_text,
    extract_ppt_text,
    extract_word_text,
    format_file_size,
    safe_error_message,
)

# 向后兼容性：旧版本的 AstrBot 未公开此钩子.
_on_plugin_error_decorator = getattr(filter, "on_plugin_error", None)
if _on_plugin_error_decorator is None:
    logger.debug(
        "[plugin-error-hook] on_plugin_error is unavailable in this AstrBot version; fallback to default error handling."
    )

    def on_plugin_error_filter(*args, **kwargs):
        def decorator(func):
            return func

        return decorator

else:
    on_plugin_error_filter = _on_plugin_error_decorator


class FileOperationPlugin(Star):
    """基于工具调用的智能文件管理插件"""

    def __init__(self, context: Context, config: AstrBotConfig):
        super().__init__(context)
        self.config = config

        # 预加载常用配置
        file_settings = self.config.get("file_settings", {})
        trigger_settings = self.config.get("trigger_settings", {})
        preview_settings = self.config.get("preview_settings", {})
        path_settings = self.config.get("path_settings", {})

        self._auto_delete = file_settings.get("auto_delete_files", True)
        self._max_file_size = (
            file_settings.get("max_file_size_mb", DEFAULT_MAX_FILE_SIZE_MB)
            * 1024
            * 1024
        )
        self._buffer_wait = file_settings.get("message_buffer_seconds", 4)
        self._reply_to_user = trigger_settings.get("reply_to_user", True)
        self._require_at_in_group = trigger_settings.get("require_at_in_group", True)
        self._enable_features_in_group = trigger_settings.get(
            "enable_features_in_group", False
        )
        self._auto_block_execution_tools = trigger_settings.get(
            "auto_block_execution_tools", True
        )
        self._enable_preview = preview_settings.get("enable", True)
        self._preview_dpi = preview_settings.get("dpi", 150)
        self._allow_external_input_files = path_settings.get(
            "allow_external_input_files", False
        )
        self._recent_text_ttl_seconds = max(10, int(self._buffer_wait) + 10)
        self._recent_text_max_entries = 512
        self._recent_text_cleanup_interval_seconds = max(
            5, min(60, self._recent_text_ttl_seconds)
        )

        # 根据配置决定使用临时目录还是持久化目录
        if self._auto_delete:
            # 使用临时目录，发送后自动删除
            self._temp_dir = tempfile.TemporaryDirectory(prefix="astrbot_file_")
            self.plugin_data_path = Path(self._temp_dir.name)
        else:
            # 持久化存储到标准插件数据目录
            self._temp_dir = None
            self.plugin_data_path = StarTools.get_data_dir() / "files"
            self.plugin_data_path.mkdir(parents=True, exist_ok=True)

        # 统一的线程池，供所有子模块共享
        self._executor = ThreadPoolExecutor(max_workers=4)

        self.office_gen = OfficeGenerator(
            self.plugin_data_path, executor=self._executor
        )
        self.pdf_converter = PDFConverter(
            self.plugin_data_path, executor=self._executor
        )

        # 初始化预览图生成器
        self.preview_gen = PreviewGenerator(dpi=self._preview_dpi)

        self._office_libs = self._check_office_libs()
        self._recent_text_by_session: dict[tuple[str, str, str], tuple[str, float]] = {}
        self._recent_text_last_cleanup_ts = 0.0

        # 初始化消息缓冲器
        self._message_buffer = MessageBuffer(wait_seconds=self._buffer_wait)
        self._message_buffer.set_complete_callback(self._on_buffer_complete)

        mode = "临时目录(自动删除)" if self._auto_delete else "持久化存储"
        logger.info(
            f"[文件管理] 插件加载完成。模式: {mode}, 数据目录: {self.plugin_data_path}"
        )

    async def terminate(self):
        """插件卸载时释放资源"""
        # 清理 Office 生成器资源
        if hasattr(self, "office_gen") and self.office_gen:
            self.office_gen.cleanup()
            logger.debug("[文件管理] Office生成器已清理")

        # 清理 PDF 转换器资源
        if hasattr(self, "pdf_converter") and self.pdf_converter:
            self.pdf_converter.cleanup()
            logger.debug("[文件管理] PDF转换器已清理")

        # 关闭主线程池（子模块使用共享线程池，不会自己关闭）
        if hasattr(self, "_executor") and self._executor:
            self._executor.shutdown(wait=False)
            logger.debug("[文件管理] 主线程池已关闭")

        # 清理临时目录
        if hasattr(self, "_temp_dir") and self._temp_dir:
            try:
                self._temp_dir.cleanup()
                logger.debug("[文件管理] 临时目录已清理")
            except Exception as e:
                logger.warning(f"[文件管理] 清理临时目录失败: {e}")

    async def _on_buffer_complete(self, buf: BufferedMessage):
        """
        消息缓冲完成后的回调（有文件时）

        将聚合后的文件和文本消息合并，重新构造消息链并触发处理。
        """
        event = buf.event
        files = buf.files
        texts = buf.texts

        logger.info(f"[消息缓冲] 缓冲完成，文件数: {len(files)}, 文本数: {len(texts)}")

        # 检查重入次数，防止无限循环
        reentry_count = getattr(event, "_buffer_reentry_count", 0)
        if reentry_count >= 3:
            logger.warning("[消息缓冲] 事件重入次数过多，停止处理")
            return

        # 构建文件信息列表
        file_info_list = []
        has_readable_file = False
        for f in files:
            name = ""
            if isinstance(f, Comp.File):
                name = f.name or ""
            name = name or "未命名文件"
            suffix = Path(name).suffix.lower() if name else ""
            file_info_list.append(f"文件名: {name} (类型: {suffix})")
            if (
                suffix in ALL_OFFICE_SUFFIXES
                or suffix in TEXT_SUFFIXES
                or suffix == PDF_SUFFIX
            ):
                has_readable_file = True

        # 合并用户的文本指令
        user_instruction = " ".join(texts) if texts else ""
        if not user_instruction:
            user_instruction = self._pop_recent_text(event)

        # 构建给 LLM 的提示文本
        if has_readable_file and user_instruction:
            prompt_text = (
                f"\n[系统通知] 用户上传了 {len(file_info_list)} 个文件:\n"
                + "\n".join(file_info_list)
                + f"\n\n用户指令: {user_instruction}"
                + "\n\n请使用 `read_file` 工具读取上述文件内容，然后根据用户指令进行处理。"
            )
        elif has_readable_file:
            prompt_text = (
                f"\n[系统通知] 用户上传了 {len(file_info_list)} 个文件:\n"
                + "\n".join(file_info_list)
                + "\n\n请立即使用 `read_file` 工具读取上述文件内容。"
                "\n(注意：用户未提供具体指令，请读取文件后询问用户需要什么帮助)"
            )
        else:
            prompt_text = (
                f"\n[系统通知] 用户上传了 {len(file_info_list)} 个文件:\n"
                + "\n".join(file_info_list)
                + "\n\n请根据用户需求处理这些文件。"
            )

        # 重构消息链
        # 注意：不要把 At 放在开头，会影响 WakingCheckStage 的检查逻辑
        new_chain = []
        new_chain.append(Comp.Plain(prompt_text))

        # 保留原始文件组件（用于 before_llm_chat 处理）
        for f in files:
            new_chain.append(f)

        # 修改事件对象
        event.message_obj.message = new_chain
        if hasattr(event.message_obj, "raw_message"):
            event.message_obj.raw_message = prompt_text
        # 更新 message_str（唤醒检查会用到）
        event.message_str = prompt_text.strip()

        logger.info(f"[消息缓冲] 已合并消息，提示: {prompt_text[:50]}...")

        # 重新触发事件处理
        # 通过 context 的 event_queue 重新将事件放入队列
        try:
            # 标记事件已经过缓冲处理，避免重复缓冲
            setattr(event, "_buffered", True)
            setattr(event, "_buffer_reentry_count", reentry_count + 1)

            # 重置事件状态，让它可以继续传播
            event._result = None
            # 预设唤醒状态，跳过 WakingCheckStage 的唤醒检查
            event.is_wake = True
            event.is_at_or_wake_command = True

            # 使用 context 的 event_queue 重新分发事件
            event_queue = self.context.get_event_queue()
            await event_queue.put(event)
            logger.debug("[消息缓冲] 事件已重新放入队列")
        except Exception as e:
            logger.error(f"[消息缓冲] 重新分发事件失败: {e}")

    def _check_permission(self, event: AstrMessageEvent) -> bool:
        """检查用户权限"""
        logger.debug("正在检查用户权限")

        # 管理员始终有权限
        if event.is_admin():
            return True

        # 白名单检查（空白名单 = 仅管理员可用）
        whitelist = self.config.get("permission_settings", {}).get(
            "whitelist_users", []
        )
        if not whitelist:
            return False

        user_id = str(event.get_sender_id())
        return user_id in [str(u) for u in whitelist]

    def _is_group_message(self, event: AstrMessageEvent) -> bool:
        """Return True if event is from group chat."""
        return event.message_obj.type == MessageType.GROUP_MESSAGE

    def _is_group_feature_enabled(self, event: AstrMessageEvent) -> bool:
        """Return True when plugin features are allowed in current chat type."""
        if not self._is_group_message(event):
            return True
        return bool(self._enable_features_in_group)

    def _group_feature_disabled_error(self) -> str:
        """Unified message when group features are disabled by configuration."""
        return (
            "错误：群聊中已禁用本插件功能，请私聊使用，或在配置中开启“群聊启用插件功能”"
        )

    def _is_bot_mentioned(self, event: AstrMessageEvent) -> bool:
        """检查是否被@/回复"""
        try:
            bot_id = str(event.message_obj.self_id)
            for segment in event.message_obj.message:
                if isinstance(segment, At) or isinstance(segment, Reply):
                    target_id = getattr(segment, "qq", None) or getattr(
                        segment, "target", None
                    )
                    if target_id and str(target_id) == bot_id:
                        return True
            return False
        except Exception as e:
            logger.error(f"未知错误{e}")
            return False

    def _validate_path(
        self, filename: str, *, allow_external: bool = False
    ) -> tuple[bool, Path, str]:
        """
        验证文件路径安全性
        返回: (是否有效, 文件路径, 错误信息)
        """
        input_path = Path(filename).expanduser()
        try:
            if input_path.is_absolute():
                resolved = input_path.resolve()
            else:
                resolved = (self.plugin_data_path / input_path).resolve()

            base = self.plugin_data_path.resolve()
            if resolved.is_relative_to(base):
                return True, resolved, ""

            if allow_external and input_path.is_absolute():
                return True, resolved, ""

            return False, resolved, "非法路径：禁止访问工作区外的文件"
        except Exception as e:
            fallback = (
                input_path
                if input_path.is_absolute()
                else (self.plugin_data_path / input_path)
            )
            return False, fallback, f"路径解析失败: {e}"

    def _display_name(self, filename: str | Path) -> str:
        """Return a safe display name without directory path."""
        value = str(filename).strip()
        if not value:
            return ""
        name = Path(value).name
        return name or value

    def _get_attachment_session_key(
        self, event: AstrMessageEvent
    ) -> tuple[str, str, str]:
        """Build a stable key for per-session image attachments."""
        platform_id = event.get_platform_id() or ""
        sender_id_obj = event.get_sender_id()
        sender_id = str(sender_id_obj) if sender_id_obj is not None else ""
        origin = event.unified_msg_origin or ""

        # 在字段缺失时做更细粒度回退，避免多个真实会话共享固定 "unknown" key。
        message_obj = getattr(event, "message_obj", None)
        if message_obj is not None:
            if not platform_id:
                platform_id = str(
                    getattr(message_obj, "platform", "")
                    or getattr(message_obj, "platform_id", "")
                    or getattr(message_obj, "self_id", "")
                )
            if not sender_id:
                sender_id = str(
                    getattr(message_obj, "sender_id", "")
                    or getattr(message_obj, "user_id", "")
                    or getattr(message_obj, "qq", "")
                )
            if not origin:
                origin = str(
                    getattr(message_obj, "session_id", "")
                    or getattr(message_obj, "conversation_id", "")
                    or getattr(message_obj, "group_id", "")
                    or getattr(message_obj, "channel_id", "")
                    or getattr(message_obj, "target_id", "")
                )

        return (
            platform_id or f"unknown_platform:{id(self)}",
            sender_id
            or f"unknown_sender:{id(message_obj) if message_obj else id(event)}",
            origin or f"unknown_origin:{id(event)}",
        )

    def _cleanup_recent_text_cache(self, now: float, *, force: bool = False) -> None:
        """Cleanup expired entries and enforce hard capacity cap."""
        if (
            not force
            and now - self._recent_text_last_cleanup_ts
            < self._recent_text_cleanup_interval_seconds
            and len(self._recent_text_by_session) <= self._recent_text_max_entries
        ):
            return

        expire_before = now - self._recent_text_ttl_seconds
        expired_keys = [
            key
            for key, (_, ts) in self._recent_text_by_session.items()
            if ts <= expire_before
        ]
        for key in expired_keys:
            self._recent_text_by_session.pop(key, None)

        overflow = len(self._recent_text_by_session) - self._recent_text_max_entries
        if overflow > 0:
            oldest_keys = sorted(
                self._recent_text_by_session.items(), key=lambda item: item[1][1]
            )[:overflow]
            for key, _ in oldest_keys:
                self._recent_text_by_session.pop(key, None)

        self._recent_text_last_cleanup_ts = now

    def _store_uploaded_file(self, src_path: Path, preferred_name: str) -> Path:
        """Store uploaded file in workspace with collision-safe naming."""
        safe_name = Path(preferred_name).name or "uploaded_file"
        valid, dst_path, error = self._validate_path(safe_name)
        if not valid:
            raise ValueError(error)

        if self._try_copy_uploaded_file(src_path, dst_path):
            return dst_path

        stem = dst_path.stem or "file"
        suffix = dst_path.suffix
        index = 1
        while True:
            candidate_name = f"{stem}_{index}{suffix}"
            valid, candidate_path, error = self._validate_path(candidate_name)
            if not valid:
                raise ValueError(error)
            if self._try_copy_uploaded_file(src_path, candidate_path):
                return candidate_path
            index += 1

    def _try_copy_uploaded_file(self, src_path: Path, dst_path: Path) -> bool:
        """Try to copy upload file with exclusive-create semantics.

        Returns:
            True: copy succeeded and target was created.
            False: target already exists.
        """
        try:
            with src_path.open("rb") as src, dst_path.open("xb") as dst:
                shutil.copyfileobj(src, dst)
            try:
                shutil.copystat(src_path, dst_path)
            except OSError:
                pass
            return True
        except FileExistsError:
            return False

    def _remember_recent_text(self, event: AstrMessageEvent) -> None:
        """Remember latest user instruction text for current session."""
        text = str(event.message_str or "").strip()
        if not text:
            return
        if text.startswith("[系统通知]"):
            return
        now = time.time()
        self._cleanup_recent_text_cache(now)
        session_key = self._get_attachment_session_key(event)
        self._recent_text_by_session[session_key] = (text, now)
        if len(self._recent_text_by_session) > self._recent_text_max_entries:
            self._cleanup_recent_text_cache(now, force=True)

    def _pop_recent_text(self, event: AstrMessageEvent) -> str:
        """Pop recent instruction text if still fresh for current session."""
        now = time.time()
        self._cleanup_recent_text_cache(now)
        session_key = self._get_attachment_session_key(event)
        item = self._recent_text_by_session.get(session_key)
        if not item:
            return ""

        text, ts = item
        if now - ts > self._recent_text_ttl_seconds:
            self._recent_text_by_session.pop(session_key, None)
            return ""

        self._recent_text_by_session.pop(session_key, None)
        return text

    async def _extract_upload_source(
        self, component: Comp.File
    ) -> tuple[Path | None, str]:
        """Extract local source path and display name from upload component."""
        file_path = await component.get_file()
        if not file_path:
            return None, component.name or "unknown_file"
        return Path(file_path), component.name or Path(file_path).name

    def _pre_check(
        self,
        event: AstrMessageEvent,
        filename: str | None = None,
        *,
        check_permission: bool = True,
        feature_key: str | None = None,
        require_exists: bool = False,
        allowed_suffixes: frozenset | set | None = None,
        required_suffix: str | None = None,
        allow_external_path: bool = False,
    ) -> tuple[bool, Path | None, str | None]:
        """
        统一的前置检查方法

        Args:
            event: 消息事件
            filename: 要检查的文件名（可选）
            check_permission: 是否检查用户权限
            feature_key: 功能配置键名（如 "enable_pdf_conversion"）
            require_exists: 是否要求文件存在
            allowed_suffixes: 允许的文件后缀集合
            required_suffix: 必须的文件后缀（如 ".pdf"）
            allow_external_path: 是否允许工作区外绝对路径

        Returns:
            (通过检查, 文件路径, 错误信息)
            - 通过时: (True, Path, None)
            - 失败时: (False, None, "错误信息")
        """
        if not self._is_group_feature_enabled(event):
            return False, None, self._group_feature_disabled_error()

        # 权限检查
        if check_permission and not self._check_permission(event):
            return False, None, "错误：权限不足"

        # 功能开关检查
        if feature_key:
            if not self.config.get("feature_settings", {}).get(feature_key, True):
                return False, None, "错误：该功能已被禁用"

        # 如果不需要检查文件，直接返回成功
        if filename is None:
            return True, None, None

        display_name = self._display_name(filename)

        # 路径验证
        valid, file_path, error = self._validate_path(
            filename, allow_external=allow_external_path
        )
        if not valid:
            return False, None, f"错误：{error}"

        # 文件存在性检查
        if require_exists and not file_path.exists():
            return False, None, f"错误：文件 '{display_name}' 不存在"

        # 文件后缀检查
        suffix = file_path.suffix.lower()
        if required_suffix and suffix != required_suffix:
            return (
                False,
                None,
                f"错误：仅支持 {required_suffix} 文件，当前格式: {suffix}",
            )

        if allowed_suffixes and suffix not in allowed_suffixes:
            supported = ", ".join(allowed_suffixes)
            return (
                False,
                None,
                f"错误：不支持的文件格式 '{suffix}'，仅支持: {supported}",
            )

        return True, file_path, None

    def _check_office_libs(self) -> dict:
        """检查并缓存 Office 库的可用性"""
        libs = {}
        for office_type in OFFICE_LIBS:
            try:
                module_name, package_name = OFFICE_LIBS[office_type]
                libs[module_name] = importlib.import_module(module_name)
                logger.debug(f"[文件管理] {package_name} 已加载")
            except ImportError:
                libs[module_name] = None
                logger.warning(f"[文件管理] {package_name} 未安装")
        return libs

    def _get_max_file_size(self) -> int:
        """获取最大文件大小（字节）"""
        return self._max_file_size

    async def _send_file_with_preview(
        self,
        event: AstrMessageEvent,
        file_path: Path,
        success_message: str = "✅ 文件已处理成功",
    ) -> None:
        """发送文件，并根据配置生成预览图

        Args:
            event: 消息事件
            file_path: 要发送的文件路径
            success_message: 成功消息前缀
        """
        preview_path = None

        # 生成预览图
        if self._enable_preview:
            try:
                loop = asyncio.get_running_loop()
                preview_path = await loop.run_in_executor(
                    self._executor,
                    self.preview_gen.generate_preview,
                    file_path,
                    None,
                )
            except Exception as e:
                logger.warning(f"[文件管理] 生成预览图失败: {e}")
                preview_path = None

        # 构建并发送消息
        text_chain = MessageChain()
        text_chain.message(f"{success_message}：{file_path.name}")
        if self._reply_to_user:
            text_chain.chain.append(Comp.At(qq=event.get_sender_id()))
        await event.send(text_chain)

        # 先发送预览图（如果有）
        if preview_path and preview_path.exists():
            await event.send(
                MessageChain([Comp.Image(file=str(preview_path.resolve()))])
            )
            # 清理预览图
            if self._auto_delete:
                try:
                    preview_path.unlink()
                except Exception:
                    pass

        # 发送文件
        await event.send(
            MessageChain(
                [Comp.File(file=str(file_path.resolve()), name=file_path.name)]
            )
        )

        # 根据配置决定是否删除文件
        if self._auto_delete and file_path.exists():
            try:
                file_path.unlink()
                logger.debug(f"[文件管理] 已自动删除文件: {file_path.name}")
            except Exception as del_e:
                logger.warning(f"[文件管理] 自动删除文件失败: {del_e}")

    async def _read_text_file(
        self, file_path: Path, max_size: int, chunk_size: int = DEFAULT_CHUNK_SIZE
    ) -> str:
        """异步分块读取文本文件"""
        loop = asyncio.get_running_loop()
        return await loop.run_in_executor(
            self._executor, self._read_text_file_sync, file_path, max_size, chunk_size
        )

    def _read_text_file_sync(
        self, file_path: Path, max_size: int, chunk_size: int
    ) -> str:
        """同步分块读取文本文件"""
        if chunk_size <= 0:
            chunk_size = DEFAULT_CHUNK_SIZE

        chunks = []
        bytes_read = 0
        with open(file_path, encoding="utf-8", errors="replace") as f:
            while bytes_read < max_size:
                chunk = f.read(chunk_size)
                if not chunk:
                    break
                chunks.append(chunk)
                bytes_read += len(chunk.encode("utf-8"))

        content = "".join(chunks)
        if bytes_read >= max_size:
            content += (
                f"\n\n[警告: 文件内容已截断，仅显示前 {format_file_size(max_size)}]"
            )
        return content

    def _extract_office_text(
        self, file_path: Path, office_type: OfficeType
    ) -> str | None:
        """根据 Office 类型提取文本内容"""
        extractors = {
            OfficeType.WORD: ("docx", extract_word_text),
            OfficeType.EXCEL: ("openpyxl", extract_excel_text),
            OfficeType.POWERPOINT: ("pptx", extract_ppt_text),
        }
        lib_key, extractor = extractors.get(office_type, (None, None))

        # 检查库是否可用/已加载
        if not lib_key or not self._office_libs.get(lib_key):
            logger.debug(
                f"[文件管理] Office 类型 '{office_type.name}' 对应的库未加载或类型不支持。"
            )
            return None

        # 确保提取器是可调用的
        if not callable(extractor):
            logger.error(
                f"[文件管理] 针对 Office 类型 '{office_type.name}' 的文本提取器不可调用。"
            )
            return None

        return extractor(file_path)

    def _format_file_result(
        self, filename: str, suffix: str, file_size: int, content: str
    ) -> str:
        """格式化文件读取结果"""
        return (
            f"[文件信息] 文件名: {filename}, 类型: {suffix}, 大小: {format_file_size(file_size)}\n"
            f"[文件内容]\n{content}"
        )

    def _extract_pdf_text(self, file_path: Path) -> str | None:
        """使用 pdfplumber 提取 PDF 文本内容"""
        try:
            text_parts = []
            with pdfplumber.open(file_path) as pdf:
                for i, page in enumerate(pdf.pages, 1):
                    page_text = page.extract_text()
                    if page_text:
                        text_parts.append(f"--- 第 {i} 页 ---\n{page_text}")
            if text_parts:
                return "\n\n".join(text_parts)
            logger.warning(f"[文件管理] PDF 文件 {file_path.name} 未提取到文本")
            return None
        except Exception as e:
            logger.error(f"[文件管理] 提取 PDF 文本失败: {e}")
            return None

    @filter.event_message_type(filter.EventMessageType.ALL, priority=100)
    async def on_file_message(self, event: AstrMessageEvent):
        """
        拦截包含文件的消息，使用缓冲器聚合文件和后续文本消息
        仅处理支持的文件格式（Office、文本、PDF），其他格式直接放行
        """
        # 检查是否已经过缓冲处理，避免重复缓冲
        if getattr(event, "_buffered", False):
            return

        if not self._is_group_feature_enabled(event):
            return

        # 过滤空消息（如"正在输入..."状态消息）
        if not event.message_obj.message:
            return

        self._remember_recent_text(event)

        # 检查消息是否包含支持的文件格式
        has_supported_file = False
        for component in event.message_obj.message:
            if isinstance(component, Comp.File):
                name = component.name or ""
                suffix = Path(name).suffix.lower() if name else ""
                # 只有支持的格式才进入缓冲流程
                if (
                    suffix in ALL_OFFICE_SUFFIXES
                    or suffix in TEXT_SUFFIXES
                    or suffix == PDF_SUFFIX
                ):
                    has_supported_file = True
                    break

        # 只有包含支持格式的文件才需要缓冲
        # 不支持的文件（如视频）直接放行
        if not has_supported_file:
            # 检查是否有正在等待的缓冲（用户可能先发文件再发文本）
            if self._message_buffer.is_buffering(event):
                # 有缓冲正在等待，将此消息加入缓冲
                await self._message_buffer.add_message(event)
                event.stop_event()
                logger.debug("[文件管理] 消息已加入现有缓冲")
            # 不支持的文件格式直接放行，不做任何处理
            return

        # 消息包含支持的文件格式，进行缓冲
        buffered = await self._message_buffer.add_message(event)

        if buffered:
            # 消息已被缓冲，停止事件传播
            # 等待缓冲完成后会通过回调重新触发处理
            event.stop_event()
            logger.debug("[文件管理] 支持的文件已缓冲，等待聚合...")
            return

    @filter.on_llm_request()
    async def before_llm_chat(self, event: AstrMessageEvent, req: ProviderRequest):
        """动态控制工具可见性"""
        is_group = event.message_obj.type == MessageType.GROUP_MESSAGE
        is_friend = event.message_obj.type == MessageType.FRIEND_MESSAGE

        if not self._is_group_feature_enabled(event):
            if req.func_tool:
                for tool_name in FILE_TOOLS:
                    req.func_tool.remove_tool(tool_name)
            logger.debug("[文件管理] 群聊总开关关闭，已隐藏全部文件工具")
            return

        has_permission = self._check_permission(event)
        can_process_upload = has_permission or event.is_admin()

        # 判断是否暴露文件工具
        should_expose = (
            # 管理员私聊始终可用
            (is_friend and event.is_admin())
            # 有权限且满足群聊条件
            or (
                has_permission
                and (
                    not is_group
                    or not self._require_at_in_group
                    or self._is_bot_mentioned(event)
                )
            )
        )

        if not should_expose:
            logger.debug(
                f"[文件管理] 用户 {event.get_sender_id()} 权限不足，已隐藏文件工具"
            )
            if req.func_tool:
                for tool_name in FILE_TOOLS:
                    req.func_tool.remove_tool(tool_name)

        # 仅在插件功能实际生效时，按配置屏蔽执行类工具，避免误伤无权限会话。
        if should_expose and req.func_tool and self._auto_block_execution_tools:
            for tool_name in EXECUTION_TOOLS:
                req.func_tool.remove_tool(tool_name)
            logger.debug("[文件管理] 已自动屏蔽 shell/python 执行类工具")

        # 文件入库不依赖“是否@机器人”，只依赖权限，避免群聊先传文件后触发工具时文件丢失
        if not can_process_upload:
            return

        # 处理文件消息
        for component in event.message_obj.message:
            if not isinstance(component, Comp.File):
                continue

            try:
                src_path, original_name = await self._extract_upload_source(component)
                if not src_path or not src_path.exists():
                    continue

                stored_path = self._store_uploaded_file(src_path, original_name)
                file_suffix = stored_path.suffix.lower()
                type_desc = ""
                is_supported = False

                if file_suffix in ALL_OFFICE_SUFFIXES:
                    type_desc = "Office文档 (Word/Excel/PPT)"
                    is_supported = True
                elif file_suffix in TEXT_SUFFIXES:
                    type_desc = "文本/代码文件"
                    is_supported = True
                elif file_suffix == PDF_SUFFIX:
                    type_desc = "PDF文档"
                    is_supported = True

                if not is_supported:
                    logger.info(
                        f"[文件管理] 文件 {original_name} 格式不支持 ({file_suffix})，跳过处理"
                    )
                    continue

                prompt = (
                    f"\n[系统通知] 收到用户上传的 {type_desc}: {original_name} (后缀: {file_suffix})。"
                    f"文件已存入工作区。如果用户需要读取或分析该文件，可使用 `read_file` 工具。"
                    "请先询问用户想对文件做什么，不要主动调用工具。"
                )

                req.system_prompt += prompt
                logger.info(f"[文件管理] 收到文件 {original_name}，已保存。")
            except Exception as e:
                logger.error(f"[文件管理] 处理上传文件失败: {e}")

    @llm_tool(name="read_file")
    async def read_file(
        self,
        event: AstrMessageEvent,
        filename: str = "",
    ) -> str | None:
        """Read a text, Office, or PDF file and return its content to the LLM.

        Supported formats:
        - Text: .txt, .md, .log, .py, .js, .ts, .json, .yaml, .xml, .csv, .html, .css, .sql, etc.
        - Office: .docx, .xlsx, .pptx, .doc, .xls, .ppt
        - PDF: .pdf (requires pdfplumber or pdf2docx)

        Unsupported:
        - Binary files such as images, videos, and audio.

        Args:
            filename(string): File name to read.
        """
        if not filename:
            return "错误：请提供要读取的文件名"

        # 统一前置检查
        ok, resolved_path, err = self._pre_check(
            event,
            filename,
            require_exists=True,
            allow_external_path=self._allow_external_input_files,
        )
        if not ok:
            return err or "错误：未知错误"

        assert resolved_path is not None  # 类型断言：ok=True 时 resolved_path 必定存在
        display_name = self._display_name(resolved_path)
        file_size = resolved_path.stat().st_size
        max_size = self._get_max_file_size()
        if file_size > max_size:
            size_str = format_file_size(file_size)
            max_str = format_file_size(max_size)
            return f"错误：文件大小 {size_str} 超过限制 {max_str}"
        try:
            suffix = resolved_path.suffix.lower()
            # 文本文件：使用流式读取并限制最大读取量以防止内存耗尽
            if suffix in TEXT_SUFFIXES:
                try:
                    content = await self._read_text_file(resolved_path, max_size)
                    return f"[文件: {display_name}, 大小: {format_file_size(file_size)}]\n{content}"
                except Exception as e:
                    logger.error(f"读取文件失败: {e}")
                    return f"错误：{safe_error_message(e, '读取文件失败')}"
            office_type = SUFFIX_TO_OFFICE_TYPE.get(suffix)
            # Office 文件：尝试提取文本（若未安装对应解析库，则提示为二进制）
            if office_type:
                extracted = self._extract_office_text(resolved_path, office_type)
                if extracted:
                    return self._format_file_result(
                        display_name, suffix, file_size, extracted
                    )
                return f"错误：文件 '{display_name}' 无法读取，可能未安装对应解析库"
            # PDF 文件：使用 pdfplumber 提取文本
            if suffix == PDF_SUFFIX:
                extracted = self._extract_pdf_text(resolved_path)
                if extracted:
                    return self._format_file_result(
                        display_name, suffix, file_size, extracted
                    )
                return f"错误：无法从 PDF 文件 '{display_name}' 中提取文本内容，文件可能为空、已损坏或只包含图片。"
            return f"错误：不支持读取 '{suffix}' 格式的文件"
        except Exception as e:
            logger.error(f"读取文件失败: {e}")
            return f"错误：{safe_error_message(e, '读取文件失败')}"

    @llm_tool(name="create_office_file")
    async def create_office_file(
        self,
        event: AstrMessageEvent,
        filename: str = "",
        content: str = "",
        file_type: str = "word",
    ):
        """Create an Office file (Excel/Word/PowerPoint) and send it to the user.

        Only basic content is supported. Advanced styles, charts, and complex layouts are not supported.

        Content format:
        - Excel: use `|` to split cells and newline to split rows, e.g. `Name|Age\\nAlice|25`
        - Word: plain text, with blank lines as paragraph separators.
        - PowerPoint: mark slides with `[Slide 1]`, or let blank lines split slides automatically

        Args:
            filename(string): Output file name (prefer .docx/.xlsx/.pptx).
            content(string): File content in the format above.
            file_type(string): Fallback type (word/excel/powerpoint) when filename has no extension.
        """
        # 统一前置检查（仅检查权限和功能开关，不检查文件）
        ok, _, err = self._pre_check(event, feature_key="enable_office_files")
        if not ok:
            await event.send(MessageChain().message(f"❌ {err}"))
            return err or "错误：未知错误"

        if not content:
            return "错误：请提供 content（文件内容）"

        filename = Path(filename).name if filename else ""
        if not filename:
            return "错误：请提供 filename（文件名）"

        # 优先根据文件名扩展名自动推断文件类型
        suffix = Path(filename).suffix.lower()
        if suffix in SUFFIX_TO_OFFICE_TYPE:
            office_type = SUFFIX_TO_OFFICE_TYPE[suffix]
        else:
            # 扩展名不匹配，使用传入的 file_type 参数
            file_type_lower = file_type.lower()
            office_type = OFFICE_TYPE_MAP.get(file_type_lower)
        if not office_type:
            await event.send(
                MessageChain().message(
                    f"❌ 不支持的类型，可选：{', '.join(OFFICE_TYPE_MAP.keys())}"
                )
            )
            return f"错误：不支持的文件类型 '{file_type}'"

        module_name = OFFICE_LIBS[office_type][0]
        if not self._office_libs.get(module_name):
            package_name = OFFICE_LIBS[office_type][1]
            await event.send(
                MessageChain().message(f"❌ 需要安装 {package_name} 才能生成此类型文件")
            )
            return f"错误：需要安装 {package_name}"

        file_info = {
            "type": office_type,
            "filename": filename,
            "content": content,
        }
        try:
            file_path = await self.office_gen.generate(
                event, file_info["type"], filename, file_info
            )
            if file_path and file_path.exists():
                file_size = file_path.stat().st_size
                max_size = self._get_max_file_size()

                if file_size > max_size:
                    # 删除过大的文件
                    file_path.unlink()
                    size_str = format_file_size(file_size)
                    max_str = format_file_size(max_size)
                    await event.send(
                        MessageChain().message(
                            f"❌ 生成的文件过大 ({size_str})，超过限制 {max_str}"
                        )
                    )
                    return f"错误：文件过大 ({size_str})，超过限制 {max_str}"

                # 发送文件（带预览图）
                await self._send_file_with_preview(event, file_path)
                return f"已将文件{file_path.name}发送给用户"
        except Exception as e:
            await event.send(MessageChain().message(f"文件操作异常: {e}"))
            return f"错误：文件操作异常: {e}"

    @llm_tool(name="convert_to_pdf")
    async def convert_to_pdf(
        self,
        event: AstrMessageEvent,
        filename: str = "",
        file_path: str = "",  # 别名，兼容 LLM 可能使用的参数名
    ) -> str:
        """Convert an Office file (Word/Excel/PowerPoint) to PDF.

        Supported extensions: .docx/.doc, .xlsx/.xls, .pptx/.ppt.
        Call this tool directly; you do not need to call `read_file` first.

        Args:
            filename(string): Office file name to convert (for example: report.docx, data.xlsx, report.xls).
        """
        # 兼容 file_path 参数名
        if not filename and file_path:
            filename = file_path

        if not filename:
            return "错误：请提供要转换的 Office 文件名"

        logger.debug(
            f"[PDF转换] convert_to_pdf 被调用，filename={self._display_name(filename)}"
        )
        # 统一前置检查
        ok, resolved_path, err = self._pre_check(
            event,
            filename,
            feature_key="enable_pdf_conversion",
            require_exists=True,
            allowed_suffixes=CONVERTIBLE_TO_PDF,
            allow_external_path=self._allow_external_input_files,
        )
        if not ok:
            logger.warning(f"[PDF转换] 前置检查失败: {err}")
            return err or "错误：未知错误"

        assert resolved_path is not None  # 类型断言：ok=True 时 resolved_path 必定存在
        display_name = self._display_name(resolved_path)

        # 检查转换器是否可用
        if not self.pdf_converter.is_available("office_to_pdf"):
            return "错误：Office→PDF 转换不可用，需要安装 LibreOffice"

        try:
            logger.info(f"[PDF转换] 开始转换: {display_name} → PDF")
            output_path = await self.pdf_converter.office_to_pdf(resolved_path)

            if output_path and output_path.exists():
                file_size = output_path.stat().st_size
                max_size = self._get_max_file_size()

                if file_size > max_size:
                    output_path.unlink()
                    return f"错误：生成的 PDF 文件过大 ({format_file_size(file_size)})"

                # 发送文件（带预览图）
                await self._send_file_with_preview(
                    event, output_path, f"✅ 已将 {display_name} 转换为 PDF"
                )

                return f"已将 {display_name} 转换为 {output_path.name} 并发送给用户"

            return "错误：PDF 转换失败，请检查文件格式是否正确"

        except Exception as e:
            logger.error(f"[PDF转换] 转换失败: {e}", exc_info=True)
            return f"错误：{safe_error_message(e, '转换失败')}"

    @llm_tool(name="convert_from_pdf")
    async def convert_from_pdf(
        self,
        event: AstrMessageEvent,
        filename: str = "",
        target_format: str = "word",
        file_id: str = "",  # 别名，兼容 LLM 可能使用的参数名
    ) -> str:
        """Convert a PDF file to an Office format (Word or Excel).

        Call this tool directly; you do not need to call `read_file` first.

        Notes:
        - PDF -> Word works best for text-heavy PDFs; complex layouts may shift.
        - PDF -> Excel extracts table data only; non-table content may be lost.

        Args:
            filename(string): PDF file name to convert (for example: document.pdf).
            target_format(string): Target format, `word` or `excel`. Default is `word`.
        """
        # 兼容 file_id 参数名
        if not filename and file_id:
            filename = file_id

        if not filename:
            return "错误：请提供要转换的 PDF 文件名"
        # 统一前置检查：权限、功能开关、文件存在性、PDF 后缀
        ok, file_path, err = self._pre_check(
            event,
            filename,
            feature_key="enable_pdf_conversion",
            require_exists=True,
            required_suffix=PDF_SUFFIX,
            allow_external_path=self._allow_external_input_files,
        )
        if not ok:
            return err or "错误：未知错误"

        assert file_path is not None  # 类型断言：ok=True 时 file_path 必定存在
        display_name = self._display_name(file_path)

        # 验证目标格式
        target = target_format.lower().strip()
        if target not in PDF_TARGET_FORMATS:
            supported = ", ".join(PDF_TARGET_FORMATS.keys())
            return f"错误：不支持的目标格式 '{target_format}'，可选: {supported}"

        _, target_desc = PDF_TARGET_FORMATS[target]

        # 检查转换器是否可用
        conversion_type = f"pdf_to_{target}"
        if not self.pdf_converter.is_available(conversion_type):
            missing = self.pdf_converter.get_missing_dependencies()
            return f"错误：PDF→{target_desc} 转换不可用，缺少依赖: {', '.join(missing)}"

        try:
            logger.info(f"[PDF转换] 开始转换: {display_name} → {target_desc}")

            if target == "word":
                output_path = await self.pdf_converter.pdf_to_word(file_path)
            elif target == "excel":
                output_path = await self.pdf_converter.pdf_to_excel(file_path)
            else:
                return f"错误：未实现的转换类型: {target}"

            if output_path and output_path.exists():
                file_size = output_path.stat().st_size
                max_size = self._get_max_file_size()

                if file_size > max_size:
                    output_path.unlink()
                    return f"错误：生成的文件过大 ({format_file_size(file_size)})"

                # 发送文件（带预览图）
                await self._send_file_with_preview(
                    event, output_path, f"✅ 已将 {display_name} 转换为 {target_desc}"
                )

                return f"已将 {display_name} 转换为 {output_path.name} 并发送给用户"

            return f"错误：PDF→{target_desc} 转换失败"

        except Exception as e:
            logger.error(f"[PDF转换] 转换失败: {e}", exc_info=True)
            return f"错误：{safe_error_message(e, '转换失败')}"

    @filter.command("delete_file", alias={"删除文件", "file_rm"})
    async def delete_file(self, event: AstrMessageEvent):
        """从工作区中永久删除指定文件。用法: /delete_file 文件名"""
        if not self._is_group_feature_enabled(event):
            await event.send(
                MessageChain().message("❌ " + self._group_feature_disabled_error())
            )
            return

        if not self._check_permission(event):
            await event.send(MessageChain().message("❌ 权限不足"))
            return

        # 从消息中获取文件名参数
        text = event.message_str.strip()
        parts = text.split(maxsplit=1)
        if len(parts) < 2:
            await event.send(MessageChain().message("❌ 用法: /delete_file 文件名"))
            return
        filename = parts[1].strip()
        display_name = self._display_name(filename)

        valid, file_path, error = self._validate_path(filename)
        if not valid:
            await event.send(MessageChain().message(f"❌ {error}"))
            return

        if file_path.exists():
            try:
                file_path.unlink(missing_ok=True)
                await event.send(
                    MessageChain().message(f"成功：文件 '{display_name}' 已删除。")
                )
                return
            except IsADirectoryError:
                await event.send(
                    MessageChain().message(f"'{display_name}'是目录,拒绝删除")
                )
                return
            except PermissionError:
                await event.send(MessageChain().message("❌ 权限不足，无法删除文件"))
                return
            except Exception as e:
                logger.error(f"删除文件时发生错误{e}")
                await event.send(MessageChain().message(f"删除文件时发生错误{e}"))
                return
        await event.send(MessageChain().message(f"错误：找不到文件 '{display_name}'"))
        return

    @filter.command("fileinfo")
    async def fileinfo(self, event: AstrMessageEvent):
        """显示文件管理工具的运行信息"""
        if not self._is_group_feature_enabled(event):
            yield event.plain_result("❌ " + self._group_feature_disabled_error())
            return

        storage_mode = "临时目录(自动删除)" if self._auto_delete else "持久化存储"

        # 获取 PDF 转换器状态
        pdf_caps = self.pdf_converter.capabilities
        pdf_status = []
        if pdf_caps.get("office_to_pdf"):
            pdf_status.append("Office→PDF ✓")
        else:
            pdf_status.append("Office→PDF ✗ (需要LibreOffice)")
        if pdf_caps.get("pdf_to_word"):
            pdf_status.append("PDF→Word ✓")
        else:
            pdf_status.append("PDF→Word ✗ (需要pdf2docx)")
        if pdf_caps.get("pdf_to_excel"):
            pdf_status.append("PDF→Excel ✓")
        else:
            pdf_status.append("PDF→Excel ✗ (需要tabula-py)")

        yield event.plain_result(
            "📂 AstrBot 文件操作工具\n"
            f"存储模式: {storage_mode}\n"
            f"工作目录: {self.plugin_data_path}\n"
            f"外部路径读取: {'开启' if self._allow_external_input_files else '关闭'}\n"
            f"群聊启用插件功能: {'开启' if self._enable_features_in_group else '关闭'}\n"
            f"自动屏蔽 shell/python 工具: {'开启' if self._auto_block_execution_tools else '关闭'}\n"
            f"回复模式: {'开启' if self._reply_to_user else '关闭'}\n"
            f"PDF转换: {', '.join(pdf_status)}"
        )

    @filter.command("list_files", alias={"文件列表", "file_ls"})
    async def list_files(self, event: AstrMessageEvent):
        """列出机器人文件库中的所有文件。"""
        if not self._is_group_feature_enabled(event):
            await event.send(
                MessageChain().message("❌ " + self._group_feature_disabled_error())
            )
            return

        if not self._check_permission(event):
            await event.send(MessageChain().message("❌ 权限不足"))
            return

        try:
            files = [
                f
                for f in self.plugin_data_path.glob("*")
                if f.is_file() and f.suffix.lower() in ALL_OFFICE_SUFFIXES
            ]
            if not files:
                msg = "文件库当前没有 Office 文件"
                if self._auto_delete:
                    msg += "（自动删除模式已开启，文件发送后会自动清理）"
                await event.send(MessageChain().message(msg))
                return

            files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
            res = ["📂 机器人工作区 Office 文件列表："]
            if self._auto_delete:
                res.append("⚠️ 自动删除模式已开启")
            for f in files:
                res.append(f"- {f.name} ({format_file_size(f.stat().st_size)})")

            result = "\n".join(res)
            await event.send(MessageChain().message(result))
        except Exception as e:
            logger.error(f"获取列表失败: {e}")
            await event.send(MessageChain().message(f"获取列表失败: {e}"))

    @filter.command("pdf_status", alias={"pdf状态"})
    async def pdf_status(self, event: AstrMessageEvent):
        """显示 PDF 转换功能的状态和依赖信息"""
        if not self._is_group_feature_enabled(event):
            await event.send(
                MessageChain().message("❌ " + self._group_feature_disabled_error())
            )
            return

        status = self.pdf_converter.get_detailed_status()
        caps = status["capabilities"]
        missing = self.pdf_converter.get_missing_dependencies()

        lines = ["📄 PDF 转换功能状态\n"]

        # 功能状态
        lines.append("【功能可用性】")
        office_status = "✅ 可用" if caps["office_to_pdf"] else "❌ 不可用"
        if status["office_to_pdf_backend"]:
            office_status += f" ({status['office_to_pdf_backend']})"
        lines.append(f"  Office→PDF: {office_status}")
        word_status = "✅ 可用" if caps["pdf_to_word"] else "❌ 不可用"
        if status["word_backend"]:
            word_status += f" ({status['word_backend']})"
        lines.append(f"  PDF→Word:   {word_status}")
        excel_status = "✅ 可用" if caps["pdf_to_excel"] else "❌ 不可用"
        if status["excel_backend"]:
            excel_status += f" ({status['excel_backend']})"
        lines.append(f"  PDF→Excel:  {excel_status}")

        # 环境信息
        lines.append("\n【环境检测】")
        lines.append(f"  平台: {'Windows' if status['is_windows'] else 'Linux/macOS'}")
        lines.append(
            f"  Java: {'✅ 可用' if status['java_available'] else '❌ 不可用'}"
        )
        if status["libreoffice_path"]:
            lines.append(f"  LibreOffice: {status['libreoffice_path']}")

        # 已安装的库
        libs = status["libs"]
        installed = [k for k, v in libs.items() if v]
        if installed:
            lines.append(f"\n【已安装库】 {', '.join(installed)}")

        # 缺失依赖
        if missing:
            lines.append("\n【缺失依赖】")
            for dep in missing:
                lines.append(f"  • {dep}")
        else:
            lines.append("\n✅ 所有依赖已安装")

        yield event.plain_result("\n".join(lines))

    @on_plugin_error_filter()
    async def on_plugin_error(
        self,
        event: AstrMessageEvent,
        plugin_name: str,
        handler_name: str,
        error: Exception,
        traceback_text: str,
    ) -> None:
        """Intercept plugin errors and forward to target session."""
        if plugin_name != "astrbot_plugin_office_assistant":
            return

        debug_settings = self.config.get("debug_settings", {})
        target_session = debug_settings.get(
            "debug_error_hook_target_session",
            self.config.get("debug_error_hook_target_session"),
        )
        if target_session is None:
            target_session = ""
        else:
            target_session = str(target_session).strip()
        if not target_session:
            target_session = event.unified_msg_origin

        trace_lines = traceback_text.splitlines() if traceback_text else []
        trace_tail = "\n".join(trace_lines[-3:]) if trace_lines else ""
        if len(trace_tail) > 800:
            trace_tail = trace_tail[-800:]

        sent = await self.context.send_message(
            target_session,
            MessageChain().message(
                "[plugin-error-hook]\n"
                f"plugin={plugin_name}\n"
                f"handler={handler_name}\n"
                f"error={error}\n"
                f"trace_tail={trace_tail}",
            ),
        )
        if not sent:
            logger.warning(
                f"[plugin-error-hook] target session not found: {target_session}"
            )

        event.stop_event()
