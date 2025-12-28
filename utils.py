from pathlib import Path
from typing import Optional
from astrbot.api import logger


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


def extract_word_text(file_path: Path) -> Optional[str]:
    """提取 Word 文档文本"""
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


def extract_excel_text(file_path: Path) -> Optional[str]:
    """提取 Excel 表格文本"""
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


def extract_ppt_text(file_path: Path) -> Optional[str]:
    """提取 PPT 幻灯片文本"""
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
