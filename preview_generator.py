"""
预览图生成器：为 Office/PDF 文件生成第一页预览图
"""

import sys
import tempfile
from pathlib import Path

import fitz  # PyMuPDF

from astrbot.api import logger

from .constants import ALL_OFFICE_SUFFIXES, PDF_SUFFIX


class PreviewGenerator:
    """文件预览图生成器"""

    def __init__(self, dpi: int = 150):
        """
        Args:
            dpi: 预览图分辨率，默认 150
        """
        self.dpi = dpi
        self._zoom = dpi / 72  # PDF 默认 72 DPI

    def generate_preview(self, file_path: Path, output_path: Path | None = None) -> Path | None:
        """
        生成文件第一页预览图

        Args:
            file_path: 源文件路径
            output_path: 输出图片路径，默认为源文件同目录下的 {filename}_preview.png

        Returns:
            生成的预览图路径，失败返回 None
        """
        suffix = file_path.suffix.lower()

        if output_path is None:
            output_path = file_path.with_name(f"{file_path.stem}_preview.png")

        try:
            if suffix == PDF_SUFFIX:
                return self._generate_from_pdf(file_path, output_path)

            if suffix in ALL_OFFICE_SUFFIXES:
                return self._generate_from_office(file_path, output_path)

            logger.warning(f"[预览生成] 不支持的文件格式: {suffix}")
            return None

        except Exception as e:
            logger.error(f"[预览生成] 生成预览图失败: {e}")
            return None

    def _generate_from_pdf(self, pdf_path: Path, output_path: Path) -> Path | None:
        """从 PDF 生成预览图"""
        try:
            doc = fitz.open(str(pdf_path))
            if doc.page_count == 0:
                logger.warning(f"[预览生成] PDF 文件为空: {pdf_path.name}")
                doc.close()
                return None

            page = doc[0]
            mat = fitz.Matrix(self._zoom, self._zoom)
            pix = page.get_pixmap(matrix=mat)
            pix.save(str(output_path))
            doc.close()

            logger.info(f"[预览生成] 已生成预览图: {output_path.name}")
            return output_path

        except Exception as e:
            logger.error(f"[预览生成] PDF 预览生成失败: {e}")
            return None

    def _generate_from_office(self, office_path: Path, output_path: Path) -> Path | None:
        """从 Office 文件生成预览图（先转 PDF）

        注意：docx2pdf 仅支持 Word 文件，Excel/PPT 暂不支持预览图生成
        """
        suffix = office_path.suffix.lower()

        # docx2pdf 仅支持 Word 格式
        if suffix not in (".doc", ".docx"):
            logger.debug(f"[预览生成] 暂不支持 {suffix} 格式预览图（仅支持 Word）")
            return None

        if sys.platform != "win32":
            logger.warning("[预览生成] Office 预览仅支持 Windows 系统")
            return None

        try:
            from docx2pdf import convert
        except ImportError:
            logger.warning("[预览生成] docx2pdf 未安装，无法生成 Office 预览")
            return None

        # 创建临时 PDF
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            tmp_pdf = Path(tmp.name)

        try:
            # Windows COM 需要在当前线程初始化
            import pythoncom

            pythoncom.CoInitialize()
            try:
                convert(str(office_path), str(tmp_pdf))
                result = self._generate_from_pdf(tmp_pdf, output_path)
                return result
            finally:
                pythoncom.CoUninitialize()
        except Exception as e:
            logger.error(f"[预览生成] Office 转 PDF 失败: {e}")
            return None
        finally:
            # 清理临时文件
            if tmp_pdf.exists():
                tmp_pdf.unlink()
