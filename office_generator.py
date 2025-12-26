import json
from pathlib import Path
import asyncio
from datetime import datetime
from typing import Optional
from docx import Document
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from astrbot.api.event import AstrMessageEvent
from astrbot.core.message.message_event_result import MessageChain
from pptx import Presentation
from concurrent.futures import ThreadPoolExecutor
import importlib.util

from astrbot.api import logger


class OfficeGenerator:
    """Office文件生成器"""

    def __init__(self, data_path: Path):
        self.data_path = data_path
        self.support = self._check_support()
        self._executor = ThreadPoolExecutor(max_workers=2)

    async def _generate_word(self, file_path: Path, content: dict):
        """生成Word文档"""
        loop = asyncio.get_event_loop()
        await loop.run_in_executor(
            self._executor, self._generate_word_sync, file_path, content
        )

    def _check_support(self) -> dict[str, bool]:
        """检查Office库支持"""
        support = {
            "word": False,
            "excel": False,
            "powerpoint": False,
        }

        modules = {
            "word": ("docx", "python-docx"),
            "excel": ("openpyxl", "openpyxl"),
            "powerpoint": ("pptx", "python-pptx"),
        }

        for file_type, (module_name, package_name) in modules.items():
            if importlib.util.find_spec(module_name) is not None:
                support[file_type] = True
            else:
                logger.warning(
                    f"[文件生成器] {package_name}未安装，{file_type.capitalize()}文件生成不可用"
                )

        return support

    async def generate(self, event: AstrMessageEvent, file_type: str, file_info: dict) -> Optional[Path]:
        """生成Office文件"""
        if not self.support.get(file_type, False):
            await event.send(
                MessageChain().message(
                    f"[文件生成器] {file_type}文件生成不支持，缺少相关库"
                )
            )
            return None

        try:
            filename = file_info.get("filename", f"{file_type}_file")
            content = file_info.get("content", {})

            # 解析content
            if isinstance(content, str):
                try:
                    content = json.loads(content)
                except:
                    content = self._create_default_content(file_type, content)

            # 清理文件名并添加扩展名
            filename = self._sanitize_filename(filename)
            extensions = {"word": ".docx", "excel": ".xlsx", "powerpoint": ".pptx"}
            extension = extensions[file_type]

            if not filename.endswith(extension):
                filename = filename + extension

            file_path = self._get_unique_filepath(filename)

            # 根据类型生成文件
            if file_type == "word":
                await self._generate_word(file_path, content)
            elif file_type == "excel":
                await self._generate_excel(file_path, content)
            elif file_type == "powerpoint":
                await self._generate_powerpoint(file_path, content)

            logger.info(f"[文件生成器] Office文件已生成: {file_path}")
            return file_path

        except Exception as e:
            logger.error(f"[文件生成器] 生成Office文件失败: {e}", exc_info=True)
            return None

    def _create_default_content(self, file_type: str, text: str) -> dict:
        """创建默认的内容结构"""
        if file_type == "word":
            return {"paragraphs": [text]}
        elif file_type == "excel":
            return {"sheets": [{"name": "Sheet1", "data": [[text]]}]}
        elif file_type == "powerpoint":
            return {"slides": [{"title": "内容", "content": [text]}]}
        return {}

    async def _generate_word_sync(self, file_path: Path, content: dict):
        """生成Word文档"""
        doc = Document()

        # 添加标题
        if "title" in content:
            title = doc.add_heading(content["title"], 0)
            title.alignment = 1  # 居中

        # 添加段落
        paragraphs = content.get("paragraphs", [])
        if isinstance(paragraphs, str):
            paragraphs = [paragraphs]

        for para_text in paragraphs:
            if para_text.strip():
                doc.add_paragraph(para_text)

        # 添加表格（如果有）
        if "table" in content:
            table_data = content["table"]
            if table_data and len(table_data) > 0:
                table = doc.add_table(rows=len(table_data), cols=len(table_data[0]))
                table.style = "Light Grid Accent 1"

                for i, row_data in enumerate(table_data):
                    for j, cell_value in enumerate(row_data):
                        table.rows[i].cells[j].text = str(cell_value)

        doc.save(str(file_path))

    async def _generate_excel(self, file_path: Path, content: dict):
        """生成Excel表格"""
        wb = Workbook()

        # 删除默认sheet
        if "Sheet" in wb.sheetnames:
            wb.remove(wb["Sheet"])

        # 添加工作表
        sheets = content.get("sheets", [])
        if not sheets:
            sheets = [{"name": "Sheet1", "data": [["无数据"]]}]

        for sheet_info in sheets:
            sheet_name = sheet_info.get("name", "Sheet1")
            sheet_data = sheet_info.get("data", [])

            ws = wb.create_sheet(title=sheet_name)

            # 写入数据
            for row_idx, row_data in enumerate(sheet_data, 1):
                for col_idx, cell_value in enumerate(row_data, 1):
                    cell = ws.cell(row=row_idx, column=col_idx, value=cell_value)

                    # 第一行加粗
                    if row_idx == 1:
                        cell.font = Font(bold=True)
                        cell.fill = PatternFill(
                            start_color="DDDDDD", end_color="DDDDDD", fill_type="solid"
                        )

                    cell.alignment = Alignment(horizontal="left", vertical="center")

        wb.save(str(file_path))

    async def _generate_powerpoint(self, file_path: Path, content: dict):
        """生成PowerPoint演示文稿"""
        prs = Presentation()

        # 添加幻灯片
        slides_data = content.get("slides", [])
        if not slides_data:
            slides_data = [{"title": "标题", "content": ["内容"]}]

        for slide_info in slides_data:
            # 使用标题和内容布局
            slide_layout = prs.slide_layouts[1]  # Title and Content
            slide = prs.slides.add_slide(slide_layout)

            # 设置标题
            title = slide_info.get("title", "")
            if title:
                slide.shapes.title.text = title

            # 添加内容
            content_list = slide_info.get("content", [])
            if isinstance(content_list, str):
                content_list = [content_list]

            if content_list and len(slide.shapes) > 1:
                text_frame = slide.shapes[1].text_frame
                text_frame.clear()

                for item in content_list:
                    p = text_frame.add_paragraph()
                    p.text = item
                    p.level = 0

        prs.save(str(file_path))

    def _sanitize_filename(self, filename: str) -> str:
        """清理文件名"""
        filename = "".join(
            c for c in filename if c.isalnum() or c in (" ", "-", "_", ".")
        ).strip()

        if not filename:
            filename = f"office_file_{datetime.now().strftime('%Y%m%d_%H%M%S')}"

        return filename

    def _get_unique_filepath(self, filename: str) -> Path:
        """获取唯一的文件路径"""
        file_path = self.data_path / filename

        if file_path.exists():
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            name_parts = filename.rsplit(".", 1)

            if len(name_parts) == 2:
                filename = f"{name_parts[0]}_{timestamp}.{name_parts[1]}"
            else:
                filename = f"{filename}_{timestamp}"

            file_path = self.data_path / filename

        return file_path
