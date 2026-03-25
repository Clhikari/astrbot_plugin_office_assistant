import asyncio
import shutil
from collections.abc import Callable, Mapping
from pathlib import Path

import pdfplumber

from astrbot.api import logger
from astrbot.api.event import AstrMessageEvent

from ..constants import DEFAULT_CHUNK_SIZE, OfficeType
from ..utils import (
    ExtractedWordContent,
    extract_excel_text,
    extract_ppt_text,
    extract_word_content,
    format_extracted_word_content,
    format_file_size,
)


class WorkspaceService:
    def __init__(
        self,
        *,
        plugin_data_path: Path,
        executor,
        office_libs: dict,
        max_file_size: int,
        feature_settings: Mapping[str, bool] | None = None,
    ) -> None:
        self.plugin_data_path = plugin_data_path
        self._executor = executor
        self._office_libs = office_libs
        self._max_file_size = max_file_size
        self._feature_settings = dict(feature_settings or {})

    def get_max_file_size(self) -> int:
        return self._max_file_size

    def validate_path(
        self, filename: str, *, allow_external: bool = False
    ) -> tuple[bool, Path, str]:
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
        except Exception as exc:
            fallback = (
                input_path
                if input_path.is_absolute()
                else (self.plugin_data_path / input_path)
            )
            return False, fallback, f"路径解析失败: {exc}"

    def display_name(self, filename: str | Path) -> str:
        value = str(filename).strip()
        if not value:
            return ""
        name = Path(value).name
        return name or value

    def try_copy_uploaded_file(self, src_path: Path, dst_path: Path) -> bool:
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

    def store_uploaded_file(self, src_path: Path, preferred_name: str) -> Path:
        safe_name = Path(preferred_name).name or "uploaded_file"
        valid, dst_path, error = self.validate_path(safe_name)
        if not valid:
            raise ValueError(error)

        if self.try_copy_uploaded_file(src_path, dst_path):
            return dst_path

        stem = dst_path.stem or "file"
        suffix = dst_path.suffix
        index = 1
        while True:
            candidate_name = f"{stem}_{index}{suffix}"
            valid, candidate_path, error = self.validate_path(candidate_name)
            if not valid:
                raise ValueError(error)
            if self.try_copy_uploaded_file(src_path, candidate_path):
                return candidate_path
            index += 1

    async def read_text_file(
        self, file_path: Path, max_size: int, chunk_size: int = DEFAULT_CHUNK_SIZE
    ) -> str:
        loop = asyncio.get_running_loop()
        return await loop.run_in_executor(
            self._executor, self.read_text_file_sync, file_path, max_size, chunk_size
        )

    def read_text_file_sync(
        self, file_path: Path, max_size: int, chunk_size: int
    ) -> str:
        if chunk_size <= 0:
            chunk_size = DEFAULT_CHUNK_SIZE

        chunks = []
        bytes_read = 0
        with open(file_path, encoding="utf-8", errors="replace") as file:
            while bytes_read < max_size:
                chunk = file.read(chunk_size)
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

    def extract_office_text(
        self, file_path: Path, office_type: OfficeType
    ) -> str | None:
        if office_type is OfficeType.WORD:
            extracted = self.extract_word_content(file_path)
            return self.format_word_content(extracted)

        extractors = {
            OfficeType.EXCEL: ("openpyxl", extract_excel_text),
            OfficeType.POWERPOINT: ("pptx", extract_ppt_text),
        }
        lib_key, extractor = extractors.get(office_type, (None, None))
        if not lib_key or not self._office_libs.get(lib_key):
            logger.debug(
                f"[文件管理] Office 类型 '{office_type.name}' 对应的库未加载或类型不支持。"
            )
            return None

        if not callable(extractor):
            logger.error(
                f"[文件管理] 针对 Office 类型 '{office_type.name}' 的文本提取器不可调用。"
            )
            return None

        return extractor(file_path)

    def extract_word_content(
        self,
        file_path: Path,
        *,
        include_images: bool = True,
    ) -> ExtractedWordContent | None:
        if file_path.suffix.lower() != ".doc" and not self._office_libs.get("docx"):
            logger.debug("[文件管理] Word 解析库未加载，无法提取结构化内容。")
            return None
        return extract_word_content(
            file_path,
            self.plugin_data_path,
            include_images=include_images,
        )

    def format_word_content(self, content: ExtractedWordContent | None) -> str | None:
        return format_extracted_word_content(content)

    def format_file_result(
        self, filename: str, suffix: str, file_size: int, content: str
    ) -> str:
        return (
            f"[文件信息] 文件名: {filename}, 类型: {suffix}, 大小: {format_file_size(file_size)}\n"
            f"[文件内容]\n{content}"
        )

    def extract_pdf_text(self, file_path: Path) -> str | None:
        try:
            text_parts = []
            with pdfplumber.open(file_path) as pdf:
                for index, page in enumerate(pdf.pages, 1):
                    page_text = page.extract_text()
                    if page_text:
                        text_parts.append(f"--- 第 {index} 页 ---\n{page_text}")
            if text_parts:
                return "\n\n".join(text_parts)
            logger.warning(f"[文件管理] PDF 文件 {file_path.name} 未提取到文本")
            return None
        except Exception as exc:
            logger.error(f"[文件管理] 提取 PDF 文本失败: {exc}")
            return None

    def pre_check(
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
        is_group_feature_enabled: Callable[[AstrMessageEvent], bool],
        check_permission_fn: Callable[[AstrMessageEvent], bool],
        group_feature_disabled_error: Callable[[], str],
    ) -> tuple[bool, Path | None, str | None]:
        if not is_group_feature_enabled(event):
            return False, None, group_feature_disabled_error()

        if check_permission and not check_permission_fn(event):
            return False, None, "错误：权限不足"

        if feature_key and not self._feature_settings.get(feature_key, True):
            return False, None, "错误：该功能已被禁用"

        if filename is None:
            return True, None, None

        display_name = self.display_name(filename)
        valid, file_path, error = self.validate_path(
            filename, allow_external=allow_external_path
        )
        if not valid:
            return False, None, f"错误：{error}"

        if require_exists and not file_path.exists():
            return (
                False,
                None,
                f"错误：文件 '{display_name}' 不存在。"
                " 这里只能读取工作区内文件，或在开启外部路径后读取绝对路径。"
                " 不要联网搜索；请让用户重新上传文件或提供正确的本地路径。",
            )

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
