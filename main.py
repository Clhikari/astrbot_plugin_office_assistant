import base64
import importlib
import asyncio
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from typing import Optional
from astrbot.api.event import filter, AstrMessageEvent
from astrbot.api.star import Context, Star
from astrbot.api import logger, llm_tool, AstrBotConfig
from astrbot.api.provider import ProviderRequest
from astrbot.core.message.message_event_result import MessageChain
from astrbot.core.platform.message_type import MessageType
from astrbot.core.message.components import At, Reply
import astrbot.api.message_components as Comp

# 导入底层生成器
from .office_generator import OfficeGenerator
from .utils import (
    format_file_size,
    extract_word_text,
    extract_excel_text,
    extract_ppt_text,
)
from .constants import (
    DEFAULT_MAX_FILE_SIZE_MB,
    FILE_TOOLS,
    OFFICE_LIBS,
    OFFICE_SUFFIXES,
    OFFICE_TYPE_MAP,
    SUFFIX_TO_OFFICE_TYPE,
    OfficeType,
    TEXT_SUFFIXES,
)


class FileOperationPlugin(Star):
    """基于工具调用的智能文件管理插件"""

    def __init__(self, context: Context, config: AstrBotConfig):
        super().__init__(context)
        self.config = config

        self.plugin_data_path = Path(__file__).parent / "files"
        self.plugin_data_path.mkdir(parents=True, exist_ok=True)

        self.office_gen = OfficeGenerator(self.plugin_data_path)

        self._office_libs = self._check_office_libs()
        self._executor = ThreadPoolExecutor(max_workers=2)
        logger.info(f"[文件管理] 插件加载完成。数据目录: {self.plugin_data_path}")

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
            chunk_size = 64 * 1024

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

    def _extract_office_text(self, file_path: Path, office_type: OfficeType) -> Optional[str]:
        """根据 Office 类型提取文本内容"""
        extractors = {
            OfficeType.WORD: ("docx", extract_word_text),
            OfficeType.EXCEL: ("openpyxl", extract_excel_text),
            OfficeType.POWERPOINT: ("pptx", extract_ppt_text),
        }
        lib_key, extractor = extractors.get(office_type, (None, None))
        if not lib_key or not self._office_libs.get(lib_key):
            return None
        # 检查库是否可用/已加载，并确保提取器是可调用的
        if not lib_key or not self._office_libs.get(lib_key) or not callable(extractor):
            # 记录更具体的错误日志，帮助调试
            if lib_key and self._office_libs.get(lib_key) and not callable(extractor):
                logger.error(
                    f"[文件管理] 针对 Office 类型 '{office_type.name}' 的文本提取器不可调用。"
                )
            else:
                logger.debug(
                    f"[文件管理] Office 类型 '{office_type.name}' 对应的库未加载或类型不支持。"
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

    @filter.on_llm_request()
    async def before_llm_chat(self, event: AstrMessageEvent, req: ProviderRequest):
        """动态控制工具可见性"""
        trigger_cfg = self.config.get("trigger_settings", {})
        should_expose = True
        is_group = event.message_obj.type == MessageType.GROUP_MESSAGE
        is_friend = MessageType.FRIEND_MESSAGE
        # 私聊判断
        if is_friend and event.is_admin():
            return
        # 权限拦截
        if not self._check_permission(event):
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

    @llm_tool(name="list_files")
    async def list_files(self, event: AstrMessageEvent):
        """列出机器人文件库中的所有文件。"""

        if not self._check_permission(event):
            await event.send(MessageChain().message("❌ 拒绝访问：权限不足"))
            return "拒绝访问：权限不足"
        try:
            files = [
                f
                for f in self.plugin_data_path.glob("*")
                if f.is_file() and f.suffix.lower() in OFFICE_SUFFIXES
            ]
            if not files:
                msg = "文件库当前没有 Office 文件"
                await event.send(MessageChain().message(msg))
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
        """读取文件内容并返回给 LLM 处理。LLM 会根据用户的请求（如总结、分析、提取信息等）对文件内容进行相应处理。"""
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
            file_size = file_path.stat().st_size
            # 文本文件：使用流式读取并限制最大读取量以防止内存耗尽
            if suffix in TEXT_SUFFIXES:
                try:
                    content = file_path.read_text(encoding="utf-8", errors="replace")
                    return f"[文件: {filename}, 大小: {format_file_size(file_size)}]\n{content}"

                except Exception as e:
                    logger.error(f"读取文件失败: {e}")
                    return f"错误：读取文件失败 - {e}"
            office_type = SUFFIX_TO_OFFICE_TYPE.get(suffix)
            # Office 文件：尝试提取文本（若未安装对应解析库，则提示为二进制）
            if office_type:
                extracted = self._extract_office_text(file_path, office_type)
                if extracted:
                    return self._format_file_result(filename, suffix, file_size, extracted)
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
        """在机器人工作区中创建或更新文件（仅支持 Office 文件）。"""
        filename = Path(filename).name
        if not self._check_permission(event):
            await event.send(MessageChain().message("❌ 拒绝访问：权限不足"))
            return
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

        def format_file_size(size_bytes: int) -> str:
            """格式化文件大小为可读格式"""
            if size_bytes < 1024:
                return f"{size_bytes} B"
            elif size_bytes < 1024 * 1024:
                return f"{size_bytes / 1024:.1f} KB"
            elif size_bytes < 1024 * 1024 * 1024:
                return f"{size_bytes / (1024 * 1024):.1f} MB"
            else:
                return f"{size_bytes / (1024 * 1024 * 1024):.2f} GB"

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
                b64_str = await self._read_file_as_base64(file_path)

                use_reply = self.config.get("trigger_settings", {}).get(
                    "reply_to_user", True
                )
                chain = [
                    Comp.Plain(f"✅ 文件已处理成功：{file_path.name}"),
                    Comp.File(file=f"base64://{b64_str}", name=file_path.name),
                ]
                if use_reply:
                    chain.append(Comp.At(qq=event.get_sender_id()))
                yield event.chain_result(chain)
                await event.send(
                    MessageChain().message(f"✅ 文件已处理成功：{file_path.name}")
                )
        except Exception as e:
            await event.send(MessageChain().message(f"文件操作异常: {e}"))

    @llm_tool(name="delete_file")
    async def delete_file(self, event: AstrMessageEvent, filename: str) -> str:
        """从工作区中永久删除指定文件。"""

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

    @filter.command("fileinfo")
    async def fileinfo(self, event: AstrMessageEvent):
        """显示文件管理工具的运行信息"""
        yield event.plain_result(
            "📂 AstrBot 文件操作工具\n"
            f"工作目录: {self.plugin_data_path}\n"
            f"回复模式: {'开启' if self.config.get('trigger_settings', {}).get('reply_to_user') else '关闭'}"
        )
