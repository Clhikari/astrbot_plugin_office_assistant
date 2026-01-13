"""
PDF 转换器模块

支持:
- Office → PDF:
  - Windows: docx2pdf (需要 MS Office) 或 win32com
  - 跨平台: LibreOffice
- PDF → Word (依赖 pdf2docx)
- PDF → Excel (依赖 tabula-py 或 pdfplumber)
"""

from __future__ import annotations

import asyncio
import platform
import shutil
import subprocess
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from pathlib import Path

from astrbot.api import logger

# 系统类型检测
_IS_WINDOWS = platform.system() == "Windows"

# 可选依赖的导入状态
_PDF2DOCX_AVAILABLE = False
_TABULA_AVAILABLE = False
_PDFPLUMBER_AVAILABLE = False
_DOCX2PDF_AVAILABLE = False
_WIN32COM_AVAILABLE = False

try:
    from pdf2docx import Converter

    _PDF2DOCX_AVAILABLE = True
except ImportError:
    pass

try:
    import tabula

    _TABULA_AVAILABLE = True
except ImportError:
    pass

try:
    import pdfplumber

    _PDFPLUMBER_AVAILABLE = True
except ImportError:
    pass

# Windows 专用：docx2pdf（仅支持 Word）
if _IS_WINDOWS:
    try:
        import docx2pdf  # noqa: F401

        _DOCX2PDF_AVAILABLE = True
    except ImportError:
        pass

    # win32com 支持 Word/Excel/PPT（需要 pythoncom 配合）
    try:
        import pythoncom
        import win32com.client

        _WIN32COM_AVAILABLE = True
    except ImportError:
        pass


class PDFConverter:
    """PDF 转换器"""

    # LibreOffice 可执行文件的可能路径
    LIBREOFFICE_PATHS = [
        "soffice",  # Linux/macOS (在 PATH 中)
        "libreoffice",  # Linux 备选
        r"C:\Program Files\LibreOffice\program\soffice.exe",  # Windows 默认
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",  # Windows 32位
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",  # macOS
        "/usr/bin/libreoffice",  # Linux
        "/usr/bin/soffice",  # Linux
        "/opt/libreoffice/program/soffice",  # Docker/自定义安装
    ]

    def __init__(self, data_path: Path, executor: ThreadPoolExecutor | None = None):
        self.data_path = data_path
        self._executor = executor  # 使用外部传入的线程池
        self._owns_executor = executor is None  # 标记是否自己管理线程池
        if self._owns_executor:
            self._executor = ThreadPoolExecutor(max_workers=2)
        self._libreoffice_path = self._find_libreoffice()
        self._java_available = self._check_java()

        # tabula 需要 Java，如果 Java 不可用则降级到 pdfplumber
        tabula_usable = _TABULA_AVAILABLE and self._java_available

        # Office→PDF 转换后端选择（优先级：docx2pdf > win32com > LibreOffice）
        if _IS_WINDOWS and (_DOCX2PDF_AVAILABLE or _WIN32COM_AVAILABLE):
            self._office_to_pdf_backend = (
                "docx2pdf" if _DOCX2PDF_AVAILABLE else "win32com"
            )
            office_to_pdf_available = True
        elif self._libreoffice_path:
            self._office_to_pdf_backend = "libreoffice"
            office_to_pdf_available = True
        else:
            self._office_to_pdf_backend = None
            office_to_pdf_available = False

        # 记录可用功能
        self._capabilities = {
            "office_to_pdf": office_to_pdf_available,
            "pdf_to_word": _PDF2DOCX_AVAILABLE,
            "pdf_to_excel": tabula_usable or _PDFPLUMBER_AVAILABLE,
        }

        # 记录实际使用的 Word 转换后端
        self._word_backend = "pdf2docx" if _PDF2DOCX_AVAILABLE else None

        # 记录实际使用的 Excel 转换后端
        if tabula_usable:
            self._excel_backend = "tabula"
        elif _PDFPLUMBER_AVAILABLE:
            self._excel_backend = "pdfplumber"
        else:
            self._excel_backend = None

        logger.info(
            f"[PDF转换器] 初始化完成，功能状态: {self._capabilities}, "
            f"Office→PDF 后端: {self._office_to_pdf_backend}"
        )
        logger.debug(
            f"[PDF转换器] 库检测: docx2pdf={_DOCX2PDF_AVAILABLE}, "
            f"win32com={_WIN32COM_AVAILABLE}, is_windows={_IS_WINDOWS}"
        )
        if _TABULA_AVAILABLE and not self._java_available:
            logger.warning(
                "[PDF转换器] tabula-py 已安装但 Java 不可用，将使用 pdfplumber"
            )

    def _check_java(self) -> bool:
        """检查 Java 是否可用（tabula-py 依赖 Java）"""
        # 只检查 java 是否在 PATH 中，不执行验证
        # 因为多版本 Java 环境下 subprocess 可能有环境变量问题
        java_path = shutil.which("java")
        if java_path:
            logger.debug(f"[PDF转换器] Java 可用: {java_path}")
            return True
        logger.debug("[PDF转换器] Java 不在 PATH 中")
        return False

    def _find_libreoffice(self) -> str | None:
        """查找 LibreOffice 可执行文件"""
        for path in self.LIBREOFFICE_PATHS:
            if shutil.which(path):
                logger.debug(f"[PDF转换器] 找到 LibreOffice: {path}")
                return path
            # 对于完整路径，直接检查文件是否存在
            if Path(path).exists():
                logger.debug(f"[PDF转换器] 找到 LibreOffice: {path}")
                return path
        # Windows 平台有其他后端可用，不需要警告
        if _IS_WINDOWS and (_DOCX2PDF_AVAILABLE or _WIN32COM_AVAILABLE):
            logger.debug(
                "[PDF转换器] 未找到 LibreOffice，但 Windows 可使用 docx2pdf/win32com"
            )
        else:
            logger.warning("[PDF转换器] 未找到 LibreOffice，Office→PDF 功能不可用")
        return None

    @property
    def capabilities(self) -> dict[str, bool]:
        """获取当前可用的转换功能"""
        return self._capabilities.copy()

    def is_available(self, conversion_type: str) -> bool:
        """检查指定转换类型是否可用"""
        return self._capabilities.get(conversion_type, False)

    def get_missing_dependencies(self) -> list[str]:
        """获取缺失的依赖列表"""
        missing = []
        if not self._capabilities["office_to_pdf"]:
            if _IS_WINDOWS:
                missing.append(
                    "docx2pdf (pip install docx2pdf，需要 MS Office) 或 LibreOffice"
                )
            else:
                missing.append(
                    "LibreOffice (Docker: apt-get install -y libreoffice-writer libreoffice-calc libreoffice-impress)"
                )
        if not self._capabilities["pdf_to_word"]:
            missing.append("pdf2docx (pip install pdf2docx)")
        if not self._capabilities["pdf_to_excel"]:
            missing.append(
                "pdfplumber (pip install pdfplumber，推荐) 或 tabula-py (需要 Java)"
            )
        return missing

    def get_detailed_status(self) -> dict:
        """获取详细状态信息（用于调试/状态显示）"""
        return {
            "capabilities": self._capabilities.copy(),
            "office_to_pdf_backend": self._office_to_pdf_backend,
            "libreoffice_path": self._libreoffice_path,
            "java_available": self._java_available,
            "word_backend": self._word_backend,
            "excel_backend": self._excel_backend,
            "is_windows": _IS_WINDOWS,
            "libs": {
                "pdf2docx": _PDF2DOCX_AVAILABLE,
                "tabula": _TABULA_AVAILABLE,
                "pdfplumber": _PDFPLUMBER_AVAILABLE,
                "docx2pdf": _DOCX2PDF_AVAILABLE,
                "win32com": _WIN32COM_AVAILABLE,
            },
        }

    async def office_to_pdf(self, input_path: Path, timeout: int = 120) -> Path | None:
        """
        将 Office 文件转换为 PDF

        Args:
            input_path: 输入的 Office 文件路径
            timeout: 转换超时时间（秒）

        Returns:
            生成的 PDF 文件路径，失败返回 None
        """
        if not self._office_to_pdf_backend:
            logger.error("[PDF转换器] 没有可用的 Office→PDF 转换后端")
            return None

        if not input_path.exists():
            logger.error(f"[PDF转换器] 输入文件不存在: {input_path}")
            return None

        loop = asyncio.get_running_loop()
        try:
            # 根据后端选择转换方法
            if self._office_to_pdf_backend == "docx2pdf":
                result = await loop.run_in_executor(
                    self._executor,
                    self._office_to_pdf_docx2pdf,
                    input_path,
                )
            elif self._office_to_pdf_backend == "win32com":
                result = await loop.run_in_executor(
                    self._executor,
                    self._office_to_pdf_win32com,
                    input_path,
                )
            else:  # libreoffice
                result = await loop.run_in_executor(
                    self._executor,
                    self._office_to_pdf_libreoffice,
                    input_path,
                    timeout,
                )
            return result
        except Exception as e:
            logger.error(f"[PDF转换器] Office→PDF 转换失败: {e}", exc_info=True)
            return None

    def _office_to_pdf_docx2pdf(self, input_path: Path) -> Path | None:
        """使用 docx2pdf 转换（仅支持 Word，需要 MS Office）"""
        import docx2pdf

        suffix = input_path.suffix.lower()
        if suffix not in (".doc", ".docx"):
            logger.warning(f"[PDF转换器] docx2pdf 仅支持 Word 文件，当前: {suffix}")
            # 尝试用 win32com 作为备选
            if _WIN32COM_AVAILABLE:
                return self._office_to_pdf_win32com(input_path)
            # 没有 win32com，抛出明确错误
            raise ValueError(
                "docx2pdf 仅支持 Word 文件，Excel/PPT 需要安装 pywin32: pip install pywin32"
            )

        output_path = self.data_path / f"{input_path.stem}.pdf"
        try:
            # Windows COM 需要在当前线程初始化
            pythoncom.CoInitialize()
            try:
                docx2pdf.convert(str(input_path), str(output_path))
            finally:
                pythoncom.CoUninitialize()

            if output_path.exists():
                logger.info(f"[PDF转换器] docx2pdf 转换成功: {output_path}")
                return output_path
            return None
        except Exception as e:
            logger.error(f"[PDF转换器] docx2pdf 转换失败: {e}")
            return None

    def _office_to_pdf_win32com(self, input_path: Path) -> Path | None:
        """使用 win32com 转换（支持 Word/Excel/PPT，需要 MS Office）"""
        suffix = input_path.suffix.lower()
        output_path = self.data_path / f"{input_path.stem}.pdf"
        input_abs = str(input_path.resolve())
        output_abs = str(output_path.resolve())

        app = None
        doc = None  # doc/wb/ppt

        try:
            pythoncom.CoInitialize()

            if suffix in (".doc", ".docx"):
                app = win32com.client.Dispatch("Word.Application")
                app.Visible = False
                app.DisplayAlerts = False  # 禁止弹窗
                doc = app.Documents.Open(input_abs)
                doc.SaveAs(output_abs, FileFormat=17)  # 17 = PDF

            elif suffix in (".xls", ".xlsx"):
                app = win32com.client.Dispatch("Excel.Application")
                app.Visible = False
                app.DisplayAlerts = False  # 禁止弹窗
                doc = app.Workbooks.Open(input_abs)
                doc.ExportAsFixedFormat(0, output_abs)  # 0 = PDF

            elif suffix in (".ppt", ".pptx"):
                app = win32com.client.Dispatch("PowerPoint.Application")
                app.DisplayAlerts = False  # 禁止弹窗 (ppAlertsNone = 0 在某些版本)
                doc = app.Presentations.Open(input_abs, WithWindow=False)
                doc.SaveAs(output_abs, 32)  # 32 = PDF

            else:
                logger.error(f"[PDF转换器] win32com 不支持的格式: {suffix}")
                return None

            if output_path.exists():
                logger.info(f"[PDF转换器] win32com 转换成功: {output_path}")
                return output_path
            return None

        except Exception as e:
            logger.error(f"[PDF转换器] win32com 转换失败: {e}")
            return None

        finally:
            # 确保 COM 对象正确释放
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

    def _office_to_pdf_libreoffice(self, input_path: Path, timeout: int) -> Path | None:
        """使用 LibreOffice 转换（跨平台）"""
        output_dir = self.data_path

        cmd = [
            self._libreoffice_path,
            "--headless",
            "--invisible",
            "--nologo",
            "--nofirststartwizard",
            "--convert-to",
            "pdf",
            "--outdir",
            str(output_dir),
            str(input_path),
        ]

        logger.debug(f"[PDF转换器] 执行命令: {' '.join(cmd)}")

        try:
            result = subprocess.run(
                cmd,
                capture_output=True,
                timeout=timeout,
                text=True,
            )

            if result.returncode != 0:
                logger.error(f"[PDF转换器] LibreOffice 返回错误: {result.stderr}")
                return None

            # 构建输出文件路径
            output_path = output_dir / f"{input_path.stem}.pdf"
            if output_path.exists():
                logger.info(f"[PDF转换器] LibreOffice 转换成功: {output_path}")
                return output_path

            logger.error(f"[PDF转换器] 输出文件未生成: {output_path}")
            return None

        except subprocess.TimeoutExpired:
            logger.error(f"[PDF转换器] 转换超时 ({timeout}s)")
            return None
        except Exception as e:
            logger.error(f"[PDF转换器] 执行转换命令失败: {e}")
            return None

    async def pdf_to_word(self, input_path: Path) -> Path | None:
        """
        将 PDF 转换为 Word 文档

        Args:
            input_path: 输入的 PDF 文件路径

        Returns:
            生成的 Word 文件路径，失败返回 None
        """
        if not _PDF2DOCX_AVAILABLE:
            logger.error("[PDF转换器] pdf2docx 未安装")
            return None

        if not input_path.exists():
            logger.error(f"[PDF转换器] 输入文件不存在: {input_path}")
            return None

        loop = asyncio.get_running_loop()
        try:
            result = await loop.run_in_executor(
                self._executor,
                self._pdf_to_word_sync,
                input_path,
            )
            return result
        except Exception as e:
            logger.error(f"[PDF转换器] PDF→Word 转换失败: {e}", exc_info=True)
            return None

    def _pdf_to_word_sync(self, input_path: Path) -> Path | None:
        """同步执行 PDF→Word 转换"""
        output_path = self.data_path / f"{input_path.stem}.docx"

        try:
            cv = Converter(str(input_path))
            cv.convert(str(output_path))
            cv.close()

            if output_path.exists():
                logger.info(f"[PDF转换器] PDF→Word 成功: {output_path}")
                return output_path

            return None
        except Exception as e:
            logger.error(f"[PDF转换器] pdf2docx 转换错误: {e}")
            return None

    async def pdf_to_excel(self, input_path: Path) -> Path | None:
        """
        将 PDF 中的表格转换为 Excel

        Args:
            input_path: 输入的 PDF 文件路径

        Returns:
            生成的 Excel 文件路径，失败返回 None
        """
        if not input_path.exists():
            logger.error(f"[PDF转换器] 输入文件不存在: {input_path}")
            return None

        loop = asyncio.get_running_loop()

        # 根据初始化时检测的后端选择转换方式
        if self._excel_backend == "tabula":
            try:
                result = await loop.run_in_executor(
                    self._executor,
                    self._pdf_to_excel_tabula,
                    input_path,
                )
                if result:
                    return result
            except Exception as e:
                logger.warning(f"[PDF转换器] tabula 转换失败，尝试 pdfplumber: {e}")
                # 降级到 pdfplumber
                if _PDFPLUMBER_AVAILABLE:
                    try:
                        return await loop.run_in_executor(
                            self._executor,
                            self._pdf_to_excel_pdfplumber,
                            input_path,
                        )
                    except Exception as e2:
                        logger.error(f"[PDF转换器] pdfplumber 转换也失败: {e2}")
                return None

        if self._excel_backend == "pdfplumber":
            try:
                result = await loop.run_in_executor(
                    self._executor,
                    self._pdf_to_excel_pdfplumber,
                    input_path,
                )
                return result
            except Exception as e:
                logger.error(f"[PDF转换器] pdfplumber 转换失败: {e}")

        logger.error("[PDF转换器] 没有可用的 PDF→Excel 转换库")
        return None

    def _pdf_to_excel_tabula(self, input_path: Path) -> Path | None:
        """使用 tabula 提取 PDF 表格到 Excel"""
        import pandas as pd

        output_path = self.data_path / f"{input_path.stem}.xlsx"

        try:
            # 读取所有页面的表格
            tables = tabula.read_pdf(
                str(input_path),
                pages="all",
                multiple_tables=True,
                silent=True,
            )

            if not tables:
                logger.warning("[PDF转换器] PDF 中未检测到表格")
                # 创建一个空表格提示
                tables = [pd.DataFrame({"提示": ["PDF 中未检测到表格数据"]})]

            # 写入 Excel，每个表格一个工作表
            with pd.ExcelWriter(str(output_path), engine="openpyxl") as writer:
                for i, table in enumerate(tables):
                    sheet_name = f"表格_{i + 1}" if len(tables) > 1 else "Sheet1"
                    # 限制工作表名称长度
                    sheet_name = sheet_name[:31]
                    table.to_excel(writer, sheet_name=sheet_name, index=False)

            if output_path.exists():
                logger.info(f"[PDF转换器] PDF→Excel 成功 (tabula): {output_path}")
                return output_path

            return None
        except Exception as e:
            logger.error(f"[PDF转换器] tabula 转换错误: {e}")
            raise

    def _pdf_to_excel_pdfplumber(self, input_path: Path) -> Path | None:
        """使用 pdfplumber 提取 PDF 表格到 Excel"""
        import pandas as pd

        output_path = self.data_path / f"{input_path.stem}.xlsx"

        try:
            all_tables = []

            with pdfplumber.open(str(input_path)) as pdf:
                for page_num, page in enumerate(pdf.pages):
                    tables = page.extract_tables()
                    for table in tables:
                        if table:
                            # 第一行作为表头
                            df = pd.DataFrame(table[1:], columns=table[0])
                            all_tables.append((page_num + 1, df))

            if not all_tables:
                logger.warning("[PDF转换器] PDF 中未检测到表格")
                all_tables = [(1, pd.DataFrame({"提示": ["PDF 中未检测到表格数据"]}))]

            # 写入 Excel
            with pd.ExcelWriter(str(output_path), engine="openpyxl") as writer:
                for i, (page_num, table) in enumerate(all_tables):
                    sheet_name = f"页{page_num}_表{i + 1}"[:31]
                    table.to_excel(writer, sheet_name=sheet_name, index=False)

            if output_path.exists():
                logger.info(f"[PDF转换器] PDF→Excel 成功 (pdfplumber): {output_path}")
                return output_path

            return None
        except Exception as e:
            logger.error(f"[PDF转换器] pdfplumber 转换错误: {e}")
            raise

    def get_unique_filename(self, base_name: str, extension: str) -> Path:
        """生成唯一文件名，避免覆盖"""
        output_path = self.data_path / f"{base_name}{extension}"
        if not output_path.exists():
            return output_path

        # 添加时间戳
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return self.data_path / f"{base_name}_{timestamp}{extension}"

    def cleanup(self):
        """清理资源"""
        # 只有自己创建的线程池才需要关闭
        if self._owns_executor and hasattr(self, "_executor") and self._executor:
            self._executor.shutdown(wait=False)
            logger.debug("[PDF转换器] 线程池已关闭")
