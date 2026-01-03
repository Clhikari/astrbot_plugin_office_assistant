from pathlib import Path
import platform

from astrbot.api import logger

# 检测是否为 Windows 平台，旧格式需要 win32com
_IS_WINDOWS = platform.system() == "Windows"
_WIN32COM_AVAILABLE = False

if _IS_WINDOWS:
    try:
        import pythoncom
        import win32com.client

        _WIN32COM_AVAILABLE = True
    except ImportError:
        pass


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
    import re

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

    # 旧格式 .doc 需要 win32com
    if suffix == ".doc":
        return _extract_doc_text_win32com(file_path)

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


def _extract_doc_text_win32com(file_path: Path) -> str | None:
    """使用 win32com 提取 .doc 文件文本（仅 Windows）"""
    if not _WIN32COM_AVAILABLE:
        logger.warning("读取 .doc 文件需要 pywin32，请安装: pip install pywin32")
        return None

    app = None
    doc = None
    try:
        pythoncom.CoInitialize()
        app = win32com.client.Dispatch("Word.Application")
        app.Visible = False
        app.DisplayAlerts = False

        doc = app.Documents.Open(str(file_path.resolve()))
        text = doc.Content.Text
        return text.strip() if text else None

    except Exception as e:
        logger.warning(f".doc 文本提取失败: {e}", exc_info=True)
        return None

    finally:
        if doc:
            try:
                doc.Close()
            except Exception:
                pass
        if app:
            try:
                app.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def extract_excel_text(file_path: Path) -> str | None:
    """提取 Excel 表格文本（支持 .xlsx 和 .xls）"""
    suffix = file_path.suffix.lower()

    # 旧格式 .xls 需要 win32com
    if suffix == ".xls":
        return _extract_xls_text_win32com(file_path)

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


def _extract_xls_text_win32com(file_path: Path) -> str | None:
    """使用 win32com 提取 .xls 文件文本（仅 Windows）"""
    if not _WIN32COM_AVAILABLE:
        logger.warning("读取 .xls 文件需要 pywin32，请安装: pip install pywin32")
        return None

    app = None
    wb = None
    try:
        pythoncom.CoInitialize()
        app = win32com.client.Dispatch("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False

        wb = app.Workbooks.Open(str(file_path.resolve()))
        lines = []

        for ws in wb.Worksheets:
            used_range = ws.UsedRange
            if used_range:
                for row in used_range.Rows:
                    row_values = []
                    for cell in row.Cells:
                        val = cell.Value
                        row_values.append("" if val is None else str(val))
                    lines.append("\t".join(row_values))

        return "\n".join(lines) if lines else None

    except Exception as e:
        logger.warning(f".xls 文本提取失败: {e}", exc_info=True)
        return None

    finally:
        if wb:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if app:
            try:
                app.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


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
        logger.warning("读取 .ppt 文件需要 pywin32，请安装: pip install pywin32")
        return None

    app = None
    ppt = None
    try:
        pythoncom.CoInitialize()
        app = win32com.client.Dispatch("PowerPoint.Application")
        app.DisplayAlerts = False

        ppt = app.Presentations.Open(str(file_path.resolve()), WithWindow=False)
        texts = []

        for slide in ppt.Slides:
            for shape in slide.Shapes:
                if shape.HasTextFrame:
                    text_frame = shape.TextFrame
                    if text_frame.HasText:
                        texts.append(text_frame.TextRange.Text)

        return "\n".join(texts) if texts else None

    except Exception as e:
        logger.warning(f".ppt 文本提取失败: {e}", exc_info=True)
        return None

    finally:
        if ppt:
            try:
                ppt.Close()
            except Exception:
                pass
        if app:
            try:
                app.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
