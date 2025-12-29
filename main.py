import asyncio
import base64
import importlib
import shutil
import tempfile
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import astrbot.api.message_components as Comp
from astrbot.api import AstrBotConfig, llm_tool, logger
from astrbot.api.event import AstrMessageEvent, filter
from astrbot.api.provider import ProviderRequest
from astrbot.api.star import Context, Star
from astrbot.core.message.components import At, Reply
from astrbot.core.message.message_event_result import MessageChain
from astrbot.core.platform.message_type import MessageType
from astrbot.core.utils.astrbot_path import get_astrbot_data_path

# 导入代码执行器
from .code_executor import (
    CODE_GENERATION_PROMPT,
    CodeExecutionError,
    RestrictedCodeExecutor,
    SecurityViolationError,
    extract_code_from_response,
)
from .constants import (
    DEFAULT_CHUNK_SIZE,
    DEFAULT_MAX_FILE_SIZE_MB,
    FILE_TOOLS,
    OFFICE_LIBS,
    OFFICE_SUFFIXES,
    OFFICE_TYPE_MAP,
    PLUGIN_NAME,
    SUFFIX_TO_OFFICE_TYPE,
    TEXT_SUFFIXES,
    OfficeType,
)

# 导入消息缓冲器
from .message_buffer import BufferedMessage, MessageBuffer

# 导入底层生成器
from .office_generator import OfficeGenerator
from .utils import (
    extract_excel_text,
    extract_ppt_text,
    extract_word_text,
    format_file_size,
)


class FileOperationPlugin(Star):
    """基于工具调用的智能文件管理插件"""

    def __init__(self, context: Context, config: AstrBotConfig):
        super().__init__(context)
        self.config = config

        # 根据配置决定使用临时目录还是持久化目录
        self._auto_delete = self.config.get("file_settings", {}).get(
            "auto_delete_files", True
        )

        if self._auto_delete:
            # 使用临时目录，发送后自动删除
            self._temp_dir = tempfile.TemporaryDirectory(prefix="astrbot_file_")
            self.plugin_data_path = Path(self._temp_dir.name)
        else:
            # 持久化存储到 data/plugin_data/插件名/
            self._temp_dir = None
            self.plugin_data_path = (
                Path(get_astrbot_data_path()) / "plugin_data" / PLUGIN_NAME / "files"
            )
            self.plugin_data_path.mkdir(parents=True, exist_ok=True)

        self.office_gen = OfficeGenerator(self.plugin_data_path)

        self._office_libs = self._check_office_libs()
        self._executor = ThreadPoolExecutor(max_workers=2)

        # 初始化消息缓冲器
        file_settings = self.config.get("file_settings", {})
        buffer_wait = file_settings.get("message_buffer_seconds", 2.5)
        observe_wait = file_settings.get("message_observe_seconds", 0.8)
        self._message_buffer = MessageBuffer(
            wait_seconds=buffer_wait,
            observe_seconds=observe_wait,
        )
        self._message_buffer.set_complete_callback(self._on_buffer_complete)
        self._message_buffer.set_passthrough_callback(self._on_buffer_passthrough)

        mode = "临时目录(自动删除)" if self._auto_delete else "持久化存储"
        logger.info(
            f"[文件管理] 插件加载完成。模式: {mode}, 数据目录: {self.plugin_data_path}"
        )

    async def _on_buffer_passthrough(self, buf: BufferedMessage):
        """
        无文件消息的放行回调

        观察期结束后没有收到文件，直接将原始事件重新放入队列处理。
        """
        event = buf.event
        logger.debug(
            f"[消息缓冲] 放行无文件消息，文本: {buf.texts[:2] if buf.texts else '(空)'}..."
        )

        try:
            # 标记事件已经过缓冲处理，避免重复缓冲
            setattr(event, "_buffered", True)

            # 重置事件状态，让它可以继续传播
            event._result = None

            # 使用 context 的 event_queue 重新分发事件
            event_queue = self.context.get_event_queue()
            await event_queue.put(event)
            logger.debug("[消息缓冲] 无文件事件已重新放入队列")
        except Exception as e:
            logger.error(f"[消息缓冲] 重新分发无文件事件失败: {e}")

    async def _on_buffer_complete(self, buf: BufferedMessage):
        """
        消息缓冲完成后的回调（有文件时）

        将聚合后的文件和文本消息合并，重新构造消息链并触发处理。
        """
        event = buf.event
        files = buf.files
        texts = buf.texts

        logger.info(f"[消息缓冲] 缓冲完成，文件数: {len(files)}, 文本数: {len(texts)}")

        # 构建文件信息列表
        file_info_list = []
        for f in files:
            name = f.name or "未命名文件"
            suffix = Path(name).suffix.lower() if name else ""
            file_info_list.append(f"文件名: {name} (类型: {suffix})")

        # 合并用户的文本指令
        user_instruction = " ".join(texts) if texts else ""

        # 构建给 LLM 的提示文本
        if user_instruction:
            prompt_text = (
                f"\n[系统通知] 用户上传了 {len(file_info_list)} 个文件:\n"
                + "\n".join(file_info_list)
                + f"\n\n用户指令: {user_instruction}"
                + "\n\n请使用 `read_file` 工具读取上述文件内容，然后根据用户指令进行处理。"
            )
        else:
            prompt_text = (
                f"\n[系统通知] 用户上传了 {len(file_info_list)} 个文件:\n"
                + "\n".join(file_info_list)
                + "\n\n请立即使用 `read_file` 工具读取上述文件内容。"
                "\n(注意：用户未提供具体指令，请读取文件后询问用户需要什么帮助)"
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
        logger.info("正在检查用户权限")
        perm_cfg = self.config.get("permission_settings", {})

        # 管理员检查
        if perm_cfg.get("require_admin", False) and not event.is_admin():
            return False

        # 白名单检查
        whitelist = perm_cfg.get("whitelist_users", [])
        if whitelist:
            user_id = str(event.get_sender_id())
            if user_id not in [str(u) for u in whitelist]:
                return False

        return True

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

    def _validate_path(self, filename: str) -> tuple[bool, Path, str]:
        """
        验证文件路径安全性
        返回: (是否有效, 文件路径, 错误信息)
        """
        file_path = self.plugin_data_path / filename
        try:
            resolved = file_path.resolve()
            base = self.plugin_data_path.resolve()
            if not resolved.is_relative_to(base):
                return False, file_path, "非法路径：禁止访问工作区外的文件"
            return True, file_path, ""
        except Exception as e:
            return False, file_path, f"路径解析失败: {e}"

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

    async def _read_file_as_base64(
        self, file_path: Path, chunk_size: int = 64 * 1024
    ) -> str:
        """
        异步分块读取文件并转为 Base64

        Args:
            file_path: 文件路径
            chunk_size: 每次读取的块大小，默认 64KB
                        (Base64 编码要求输入是 3 的倍数，64KB = 65536 是 3 的倍数)
        """
        loop = asyncio.get_event_loop()
        return await loop.run_in_executor(
            self._executor, self._read_file_as_base64_sync, file_path, chunk_size
        )

    def _read_file_as_base64_sync(self, file_path: Path, chunk_size: int) -> str:
        """同步分块读取文件并转为 Base64"""
        # 确保 chunk_size 是 3 的倍数
        chunk_size = (chunk_size // 3) * 3

        # 防御性检查：确保chunk_size有效
        if chunk_size <= 0:
            chunk_size = DEFAULT_CHUNK_SIZE

        encoded_chunks = []
        with open(file_path, "rb") as f:
            while True:
                chunk = f.read(chunk_size)
                if not chunk:
                    break
                encoded_chunks.append(base64.b64encode(chunk).decode("utf-8"))

        return "".join(encoded_chunks)

    def _get_max_file_size(self) -> int:
        """获取最大文件大小（字节）"""
        mb = self.config.get("file_settings", {}).get(
            "max_file_size_mb", DEFAULT_MAX_FILE_SIZE_MB
        )
        return mb * 1024 * 1024

    async def _read_text_file(
        self, file_path: Path, max_size: int, chunk_size: int = DEFAULT_CHUNK_SIZE
    ) -> str:
        """异步分块读取文本文件"""
        loop = asyncio.get_event_loop()
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

    @filter.event_message_type(filter.EventMessageType.ALL, priority=0)
    async def on_file_message(self, event: AstrMessageEvent):
        """
        拦截包含文件的消息，使用缓冲器聚合文件和后续文本消息
        """
        # 检查是否已经过缓冲处理，避免重复缓冲
        if getattr(event, "_buffered", False):
            return

        # 过滤空消息（如"正在输入..."状态消息）
        if not event.message_obj.message:
            return

        # 检查消息是否包含文件
        has_file = False
        for component in event.message_obj.message:
            if isinstance(component, Comp.File):
                has_file = True
                break

        # 只有包含文件的消息才需要缓冲
        # 纯文本消息（包括命令）直接放行，不进行缓冲
        if not has_file:
            # 检查是否有正在等待的缓冲（用户可能先发文件再发文本）
            if self._message_buffer.is_buffering(event):
                # 有缓冲正在等待，将此文本消息加入缓冲
                await self._message_buffer.add_message(event)
                event.stop_event()
                logger.debug("[文件管理] 文本消息已加入现有缓冲")
            return

        # 消息包含文件，进行缓冲
        buffered = await self._message_buffer.add_message(event)

        if buffered:
            # 消息已被缓冲，停止事件传播
            # 等待缓冲完成后会通过回调重新触发处理
            event.stop_event()
            logger.debug("[文件管理] 文件消息已缓冲，等待聚合...")
            return

    @filter.on_llm_request()
    async def before_llm_chat(self, event: AstrMessageEvent, req: ProviderRequest):
        """动态控制工具可见性"""
        trigger_cfg = self.config.get("trigger_settings", {})
        should_expose = True
        is_group = event.message_obj.type == MessageType.GROUP_MESSAGE
        is_friend = event.message_obj.type == MessageType.FRIEND_MESSAGE
        # 私聊判断
        if is_friend and event.is_admin():
            pass  # keep should_expose True
        # 权限拦截
        elif not self._check_permission(event):
            should_expose = False
        # 群聊@/回复拦截
        elif (
            is_group
            and trigger_cfg.get("require_at_in_group", True)
            and not self._is_bot_mentioned(event)
        ):
            should_expose = False

        if not should_expose and req.func_tool:
            for tool_name in FILE_TOOLS:
                req.func_tool.remove_tool(tool_name)
            return  # 如果工具被移除，通常意味着权限不足或未触发，不处理文件

        # 处理文件消息
        for component in event.message_obj.message:
            if isinstance(component, Comp.File):
                try:
                    # 获取文件路径
                    file_path = await component.get_file()
                    if file_path and Path(file_path).exists():
                        src_path = Path(file_path)
                        dst_path = self.plugin_data_path / component.name
                        # 复制文件到工作区
                        shutil.copy2(src_path, dst_path)
                        file_suffix = dst_path.suffix.lower()
                        type_desc = "未知格式文件"

                        if file_suffix in OFFICE_SUFFIXES:
                            type_desc = "Office文档 (Word/Excel/PPT)"
                        elif file_suffix in TEXT_SUFFIXES:
                            type_desc = "文本/代码文件"

                        # 构建更Prompt
                        prompt = (
                            f"\n[系统通知] 收到用户上传的 {type_desc}: {component.name} (后缀: {file_suffix})。"
                            f"文件已存入工作区。请使用 `read_file` 工具读取其内容进行分析。"
                        )
                        req.system_prompt += prompt
                        logger.info(f"[文件管理] 收到文件 {component.name}，已保存。")
                except Exception as e:
                    logger.error(f"[文件管理] 处理上传文件失败: {e}")

    @llm_tool(name="list_files")
    async def list_files(self, event: AstrMessageEvent):
        """列出机器人文件库中的所有文件。"""

        if not self._check_permission(event):
            await event.send(MessageChain().message("❌ 拒绝访问：权限不足"))
            return "拒绝访问：权限不足"

        # 自动删除模式下，文件发送后会被删除，列表通常为空
        if self._auto_delete:
            msg = "当前为自动删除模式，文件发送后会自动清理，文件库为空。"
            return msg

        try:
            files = [
                f
                for f in self.plugin_data_path.glob("*")
                if f.is_file() and f.suffix.lower() in OFFICE_SUFFIXES
            ]
            if not files:
                msg = "文件库当前没有 Office 文件,无需重复调用"
                return msg

            files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
            res = ["📂 机器人工作区 Office 文件列表："]
            for f in files:
                res.append(f"- {f.name} ({format_file_size(f.stat().st_size)})")

            result = "\n".join(res)
            await event.send(MessageChain().message(result))
            return result
        except Exception as e:
            logger.error(f"获取列表失败: {e}")
            await event.send(MessageChain().message("获取列表失败"))
            return f"获取列表失败: {e}"

    @llm_tool(name="read_file")
    async def read_file(self, event: AstrMessageEvent, filename: str) -> str | None:
        """读取文件内容并返回给 LLM 处理。LLM 会根据用户的请求（如总结、分析、提取信息等）对文件内容进行相应处理。

        Args:
            filename(string): 要读取的文件名
        """
        if not self._check_permission(event):
            return "错误：拒绝访问，权限不足"
        valid, file_path, error = self._validate_path(filename)
        if not valid:
            return f"错误：{error}"
        if not file_path.exists():
            return f"错误：文件 '{filename}' 不存在"

        file_size = file_path.stat().st_size
        max_size = self._get_max_file_size()
        if file_size > max_size:
            size_str = format_file_size(file_size)
            max_str = format_file_size(max_size)
            return f"错误：文件大小 {size_str} 超过限制 {max_str}"
        try:
            suffix = file_path.suffix.lower()
            # 文本文件：使用流式读取并限制最大读取量以防止内存耗尽
            if suffix in TEXT_SUFFIXES:
                try:
                    content = await self._read_text_file(file_path, max_size)
                    return f"[文件: {filename}, 大小: {format_file_size(file_size)}]\n{content}"
                except Exception as e:
                    logger.error(f"读取文件失败: {e}")
                    return f"错误：读取文件失败 - {e}"
            office_type = SUFFIX_TO_OFFICE_TYPE.get(suffix)
            # Office 文件：尝试提取文本（若未安装对应解析库，则提示为二进制）
            if office_type:
                extracted = self._extract_office_text(file_path, office_type)
                if extracted:
                    return self._format_file_result(
                        filename, suffix, file_size, extracted
                    )
                return f"错误：文件 '{filename}' 无法读取，可能未安装对应解析库"
            return f"错误：不支持读取 '{suffix}' 格式的文件"
        except Exception as e:
            logger.error(f"读取文件失败: {e}")
            return f"错误：读取文件失败 - {e}"

    @llm_tool(name="write_file")
    async def write_file(
        self,
        event: AstrMessageEvent,
        filename: str,
        content: str,
        file_type: str = "word",
    ):
        """在机器人工作区中创建 Office 文件并发送给用户。

        【修改文件】：先用 read_file 读取原文件，修改内容后用此工具创建（相同文件名会覆盖）。

        【Excel 格式】严格要求：
        - 用 | 分隔每个单元格，用换行分隔每一行
        - 第一行是表头，后续是数据行
        - 每行的列数必须一致
        - 示例（3列5行的表格）：
          姓名|年龄|城市
          张三|25|北京
          李四|30|上海
          王五|28|广州
          赵六|35|深圳

        【Word 格式】：
        - 纯文本书写，不要用 Markdown（禁止 #、**、- 等符号）
        - 用空行分隔段落，标题直接写文字即可

        【PPT 格式】：
        - 用 [幻灯片 1]、[幻灯片 2] 标记每页
        - 标题写在标记后第一行，内容写在后续行

        Args:
            filename(string): 文件名（含扩展名如 .xlsx/.docx/.pptx，系统会自动识别类型）
            content(string): 文件内容（严格按上述格式，不要用Markdown）
            file_type(string): 仅当文件名无扩展名时使用，可选：word、excel、powerpoint
        """
        filename = Path(filename).name
        if not self._check_permission(event):
            await event.send(MessageChain().message("❌ 拒绝访问：权限不足"))
            return

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
            return

        if not self.config.get("feature_settings", {}).get("enable_office_files", True):
            await event.send(
                MessageChain().message("错误：当前配置禁用了 Office 文件生成功能。")
            )
            return

        module_name = OFFICE_LIBS[office_type][0]
        if not self._office_libs.get(module_name):
            package_name = OFFICE_LIBS[office_type][1]
            await event.send(
                MessageChain().message(f"❌ 需要安装 {package_name} 才能生成此类型文件")
            )
            return
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
                    return
                use_reply = self.config.get("trigger_settings", {}).get(
                    "reply_to_user", True
                )

                # 先发送文本消息
                text_chain = [Comp.Plain(f"✅ 文件已处理成功：{file_path.name}")]
                if use_reply:
                    text_chain.append(Comp.At(qq=event.get_sender_id()))
                await event.send(MessageChain(text_chain))
                await event.send(
                    MessageChain(
                        [Comp.File(file=str(file_path.resolve()), name=file_path.name)]
                    )
                )
                # 发送后根据配置决定是否删除文件
                if self._auto_delete and file_path.exists():
                    try:
                        file_path.unlink()
                        logger.debug(f"[文件管理] 已自动删除文件: {file_path.name}")
                    except Exception as del_e:
                        logger.warning(f"[文件管理] 自动删除文件失败: {del_e}")
                return f"已将文件{file_path.name}发送给用户"
        except Exception as e:
            await event.send(MessageChain().message(f"文件操作异常: {e}"))

    @llm_tool(name="delete_file")
    async def delete_file(self, event: AstrMessageEvent, filename: str) -> str:
        """从工作区中永久删除指定文件。

        Args:
            filename(string): 要删除的文件名
        """

        if not self._check_permission(event):
            await event.send(MessageChain().message("❌ 拒绝访问：权限不足"))
            return ""
        valid, file_path, error = self._validate_path(filename)
        if not valid:
            return f"❌ {error}"

        if file_path.exists():
            try:
                file_path.unlink(missing_ok=True)
                await event.send(
                    MessageChain().message(f"成功：文件 '{filename}' 已删除。")
                )
                return ""
            except IsADirectoryError:
                await event.send(MessageChain().message(f"'{filename}'是目录,拒绝删除"))
                return ""
            except PermissionError:
                await event.send(MessageChain().message("❌ 权限不足，无法删除文件"))
                return ""
            except Exception as e:
                logger.error(f"删除文件时发生错误{e}")
                await event.send(MessageChain().message(f"删除文件时发生错误{e}"))
                return ""
        await event.send(MessageChain().message(f"错误：找不到文件 '{filename}'"))
        return ""

    @llm_tool(name="create_office_file")
    async def create_office_file(
        self,
        event: AstrMessageEvent,
        description: str,
        filename: str = "",
    ) -> str:
        """通过 AI 生成代码来创建高质量的 Office 文件（Excel/Word/PPT）。

        适用于需要复杂格式的文件，如：
        - 带样式的 Excel 表格（合并单元格、公式、图表）
        - 格式丰富的 Word 文档（多级标题、目录、图片）
        - 专业的 PPT 演示文稿（多页幻灯片、布局）

        Args:
            description(string): 详细描述你需要的文件内容和格式要求
            filename(string): 可选的文件名（如 report.xlsx），不填则自动生成
        """
        if not self._check_permission(event):
            await event.send(MessageChain().message("❌ 拒绝访问：权限不足"))
            return "拒绝访问：权限不足"

        if not self.config.get("feature_settings", {}).get("enable_office_files", True):
            return "错误：当前配置禁用了 Office 文件生成功能。"

        # 构建代码生成提示
        user_prompt = f"用户需求：{description}"
        if filename:
            user_prompt += f"\n文件名：{filename}"

        try:
            # 获取 LLM provider 生成代码
            provider = self.context.get_using_provider()
            if not provider:
                return "错误：没有可用的 LLM 提供商"

            # 调用 LLM 生成代码
            from astrbot.api.provider import ProviderRequest

            req = ProviderRequest(
                prompt=user_prompt, system_prompt=CODE_GENERATION_PROMPT
            )
            resp = await provider.text_chat(**req.__dict__)

            if not resp or not resp.completion_text:
                return "错误：LLM 未返回有效代码"

            # 提取代码
            code = extract_code_from_response(resp.completion_text)
            if not code:
                return "错误：无法从 LLM 响应中提取代码"

            logger.debug(f"[文件生成] 生成的代码:\n{code[:500]}...")

            # 记录生成前的文件列表
            existing_files = set(self.plugin_data_path.glob("*"))

            # 执行代码
            executor = RestrictedCodeExecutor(
                work_dir=self.plugin_data_path,
                timeout=30.0,
            )

            try:
                result = await executor.execute(code)
                logger.info(f"[文件生成] 代码执行成功: {result}")
            except SecurityViolationError as e:
                logger.warning(f"[文件生成] 安全违规: {e}")
                return f"错误：代码安全检查失败 - {e}"
            except CodeExecutionError as e:
                logger.error(f"[文件生成] 代码执行失败: {e}")
                return f"错误：代码执行失败 - {e}"
            finally:
                executor.cleanup()

            # 查找新生成的文件
            current_files = set(self.plugin_data_path.glob("*"))
            new_files = [f for f in current_files - existing_files if f.is_file()]

            if not new_files:
                return "代码执行成功，但未生成任何文件。请检查代码是否正确调用了 save_file() 函数。"

            # 发送所有生成的文件
            max_size = self._get_max_file_size()
            sent_files = []

            for file_path in new_files:
                file_size = file_path.stat().st_size
                if file_size > max_size:
                    file_path.unlink()
                    await event.send(
                        MessageChain().message(
                            f"❌ 文件 {file_path.name} 过大 ({format_file_size(file_size)})，已删除"
                        )
                    )
                    continue

                # 发送文件
                await event.send(
                    MessageChain().message(f"✅ 已生成文件：{file_path.name}")
                )
                await event.send(
                    MessageChain(
                        [Comp.File(file=str(file_path.resolve()), name=file_path.name)]
                    )
                )
                sent_files.append(file_path.name)

                # 自动删除模式
                if self._auto_delete and file_path.exists():
                    try:
                        file_path.unlink()
                        logger.debug(f"[文件生成] 已自动删除: {file_path.name}")
                    except Exception as del_e:
                        logger.warning(f"[文件生成] 自动删除失败: {del_e}")

            if sent_files:
                return f"成功生成并发送文件: {', '.join(sent_files)}"
            return "未能成功发送任何文件"

        except Exception as e:
            logger.error(f"[文件生成] 异常: {e}", exc_info=True)
            return f"错误：文件生成失败 - {e}"

    @filter.command("fileinfo")
    async def fileinfo(self, event: AstrMessageEvent):
        """显示文件管理工具的运行信息"""
        storage_mode = "临时目录(自动删除)" if self._auto_delete else "持久化存储"
        yield event.plain_result(
            "📂 AstrBot 文件操作工具\n"
            f"存储模式: {storage_mode}\n"
            f"工作目录: {self.plugin_data_path}\n"
            f"回复模式: {'开启' if self.config.get('trigger_settings', {}).get('reply_to_user') else '关闭'}"
        )
