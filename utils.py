import platform
import re
import shutil
import subprocess
from collections.abc import Generator
from contextlib import contextmanager, suppress
from pathlib import Path

from astrbot.api import logger

# 检测是否为 Windows 平台，旧格式需要 win32com
_IS_WINDOWS = platform.system() == "Windows"
_WIN32COM_AVAILABLE = False
_ANTIWORD_AVAILABLE = shutil.which("antiword") is not None

if _IS_WINDOWS:
    try:
        import pythoncom
        import win32com.client

        _WIN32COM_AVAILABLE = True
    except ImportError:
        pass


@contextmanager
def com_application(app_name: str) -> Generator:
    """Windows COM 应用上下文管理器

    自动处理 COM 初始化/反初始化和应用退出，避免资源泄漏。

    Args:
        app_name: COM 应用名称，如 "Word.Application"

    Yields:
        COM 应用对象

    Example:
        with com_application("Word.Application") as app:
            doc = app.Documents.Open(path)
            # ...
    """
    if not _WIN32COM_AVAILABLE:
        raise RuntimeError("win32com 不可用，请安装 pywin32: pip install pywin32")

    app = None
    try:
        pythoncom.CoInitialize()
        app = win32com.client.Dispatch(app_name)
        app.Visible = False
        if hasattr(app, "DisplayAlerts"):
            app.DisplayAlerts = False
        yield app
    finally:
        if app:
            with suppress(Exception):
                app.Quit()
        with suppress(Exception):
            pythoncom.CoUninitialize()


def format_file_size(size: int | float) -> str:
    """格式化文件大小"""
    if size < 0:
        return "0 B"
    if size == 0:
        return "0 B"
    for unit in ["B", "KB", "MB", "GB"]:
        if size < 1024:
            return f"{size:.2f} {unit}"
        size /= 1024
    return f"{size:.2f} TB"


def safe_error_message(error: Exception, context: str = "") -> str:
    """
    生成安全的错误消息，隐藏敏感路径信息

    Args:
        error: 异常对象
        context: 错误上下文描述

    Returns:
        不包含敏感信息的错误消息
    """
    error_str = str(error)

    # 移除可能包含路径的信息
    # 常见模式: 'D:\\path\\to\\file' 或 '/path/to/file'

    # Windows 路径模式
    error_str = re.sub(r"[A-Za-z]:\\[^\s'\"]+", "[路径已隐藏]", error_str)
    # Unix 路径模式
    error_str = re.sub(
        r"/(?:home|tmp|var|usr|opt|data)[^\s'\"]*", "[路径已隐藏]", error_str
    )

    if context:
        return f"{context}: {error_str}"
    return error_str


def extract_word_text(file_path: Path) -> str | None:
    """提取 Word 文档文本（支持 .docx 和 .doc）"""
    suffix = file_path.suffix.lower()

    # 旧格式 .doc 优先用 antiword（跨平台），其次 win32com（Windows）
    if suffix == ".doc":
        if _ANTIWORD_AVAILABLE:
            return _extract_doc_text_antiword(file_path)
        elif _WIN32COM_AVAILABLE:
            return _extract_doc_text_win32com(file_path)
        else:
            if _IS_WINDOWS:
                logger.warning(
                    "读取 .doc 文件需要 pywin32 和 Microsoft Word，请安装: pip install pywin32"
                )
            else:
                logger.warning(
                    "读取 .doc 文件需要 antiword，请安装: apt install antiword (Linux) 或 brew install antiword (macOS)"
                )
            return None

    # 新格式 .docx 使用 python-docx
    try:
        from docx import Document

        doc = Document(file_path)
        paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
        return "\n".join(paragraphs)
    except ImportError:
        return None
    except Exception as e:
        logger.warning(f"Word 文本提取失败: {e}", exc_info=True)
        return None


def _extract_doc_text_antiword(file_path: Path) -> str | None:
    """使用 antiword 提取 .doc 文件文本（Linux/macOS）"""
    try:
        result = subprocess.run(
            ["antiword", str(file_path.resolve())],
            capture_output=True,
            text=True,
            errors="replace",
            timeout=30,
        )
        if result.returncode == 0:
            return result.stdout.strip() if result.stdout else None
        else:
            logger.warning(f"antiword 执行失败: {result.stderr}")
            return None
    except subprocess.TimeoutExpired:
        logger.warning("antiword 执行超时")
        return None
    except Exception as e:
        logger.warning(f"antiword 执行出错: {e}")
        return None


def _extract_doc_text_win32com(file_path: Path) -> str | None:
    """使用 win32com 提取 .doc 文件文本（仅 Windows）"""
    try:
        with com_application("Word.Application") as app:
            doc = app.Documents.Open(str(file_path.resolve()))
            try:
                text = doc.Content.Text
                return text.strip() if text else None
            finally:
                with suppress(Exception):
                    doc.Close()
    except Exception as e:
        logger.warning(f".doc 文本提取失败: {e}", exc_info=True)
        return None


def extract_excel_text(file_path: Path) -> str | None:
    """提取 Excel 表格文本（支持 .xlsx 和 .xls）"""
    suffix = file_path.suffix.lower()

    # 旧格式 .xls 优先用 xlrd（跨平台），其次 win32com（Windows）
    if suffix == ".xls":
        result = _extract_xls_text_xlrd(file_path)
        if result is not None:
            return result
        if _WIN32COM_AVAILABLE:
            return _extract_xls_text_win32com(file_path)
        if not _IS_WINDOWS:
            logger.warning("读取 .xls 文件需要 xlrd，请安装: pip install xlrd")
        else:
            logger.warning(
                "读取 .xls 文件失败。请确保 `xlrd` 已安装 (`pip install xlrd`)，或在 Windows 上安装 `pywin32` (`pip install pywin32`) 及 Microsoft Office。"
            )
        return None

    # 新格式 .xlsx 使用 openpyxl
    try:
        from openpyxl import load_workbook

        wb = load_workbook(file_path, read_only=True, data_only=True)
        lines = []
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                lines.append("\t".join("" if v is None else str(v) for v in row))
        return "\n".join(lines)
    except ImportError:
        return None
    except Exception as e:
        logger.warning(f"Excel 文本提取失败: {e}", exc_info=True)
        return None


def _extract_xls_text_xlrd(file_path: Path) -> str | None:
    """使用 xlrd 提取 .xls 文件文本（跨平台）"""
    try:
        import xlrd

        wb = xlrd.open_workbook(str(file_path))
        lines = []
        for sheet in wb.sheets():
            for row_idx in range(sheet.nrows):
                row_values = [str(cell.value) for cell in sheet.row(row_idx)]
                lines.append("\t".join(row_values))
        return "\n".join(lines) if lines else None
    except ImportError:
        return None
    except Exception as e:
        logger.warning(f"xlrd 读取 .xls 失败: {e}")
        return None


def _extract_xls_text_win32com(file_path: Path) -> str | None:
    """使用 win32com 提取 .xls 文件文本（仅 Windows）"""
    try:
        with com_application("Excel.Application") as app:
            wb = app.Workbooks.Open(str(file_path.resolve()))
            try:
                lines = []
                for ws in wb.Worksheets:
                    used_range = ws.UsedRange
                    if used_range:
                        for row in used_range.Rows:
                            row_values = [
                                "" if cell.Value is None else str(cell.Value)
                                for cell in row.Cells
                            ]
                            lines.append("\t".join(row_values))
                return "\n".join(lines) if lines else None
            finally:
                with suppress(Exception):
                    wb.Close(SaveChanges=False)
    except Exception as e:
        logger.warning(f".xls 文本提取失败: {e}", exc_info=True)
        return None


def extract_ppt_text(file_path: Path) -> str | None:
    """提取 PPT 幻灯片文本（支持 .pptx 和 .ppt）"""
    suffix = file_path.suffix.lower()

    # 旧格式 .ppt 需要 win32com
    if suffix == ".ppt":
        return _extract_ppt_text_win32com(file_path)

    # 新格式 .pptx 使用 python-pptx
    try:
        from pptx import Presentation

        prs = Presentation(file_path)
        texts = []
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text") and shape.text.strip():
                    texts.append(shape.text)
        return "\n".join(texts)
    except ImportError:
        return None
    except Exception as e:
        logger.warning(f"PPT 文本提取失败: {e}", exc_info=True)
        return None


def _extract_ppt_text_win32com(file_path: Path) -> str | None:
    """使用 win32com 提取 .ppt 文件文本（仅 Windows）"""
    if not _WIN32COM_AVAILABLE:
        if _IS_WINDOWS:
            logger.warning(
                "读取 .ppt 文件需要 pywin32 和 Microsoft PowerPoint，请安装: pip install pywin32"
            )
        else:
            logger.warning(
                "读取 .ppt 旧格式文件仅支持 Windows 环境，请将文件另存为 .pptx 格式"
            )
        return None

    try:
        with com_application("PowerPoint.Application") as app:
            ppt = app.Presentations.Open(str(file_path.resolve()), WithWindow=False)
            try:
                texts = []
                for slide in ppt.Slides:
                    for shape in slide.Shapes:
                        if shape.HasTextFrame:
                            text_frame = shape.TextFrame
                            if text_frame.HasText:
                                texts.append(text_frame.TextRange.Text)
                return "\n".join(texts) if texts else None
            finally:
                with suppress(Exception):
                    ppt.Close()
    except Exception as e:
        logger.warning(f".ppt 文本提取失败: {e}", exc_info=True)
        return None
