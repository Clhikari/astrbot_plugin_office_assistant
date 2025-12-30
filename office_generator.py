import asyncio
import importlib.util
import json
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from pathlib import Path

from astrbot.api import logger
from astrbot.api.event import AstrMessageEvent
from astrbot.core.message.message_event_result import MessageChain

from .constants import OFFICE_EXTENSIONS, OFFICE_LIBS, OfficeType


class OfficeGenerator:
    """Office文件生成器"""

    def __init__(self, data_path: Path):
        self.data_path = data_path
        self.support = self._check_support()
        self._executor = ThreadPoolExecutor(max_workers=2)

    # 定义映射表
    _GENERATORS = {
        OfficeType.WORD: "_generate_word",
        OfficeType.EXCEL: "_generate_excel",
        OfficeType.POWERPOINT: "_generate_powerpoint",
    }

    async def _generate_word(self, file_path: Path, content: dict):
        """异步生成 Word"""
        loop = asyncio.get_running_loop()
        await loop.run_in_executor(
            self._executor, self._generate_word_sync, file_path, content
        )

    async def _generate_excel(self, file_path: Path, content: dict):
        """异步生成 Excel"""
        loop = asyncio.get_running_loop()
        await loop.run_in_executor(
            self._executor, self._generate_excel_sync, file_path, content
        )

    async def _generate_powerpoint(self, file_path: Path, content: dict):
        """异步生成 PPT"""
        loop = asyncio.get_running_loop()
        await loop.run_in_executor(
            self._executor, self._generate_ppt_sync, file_path, content
        )

    def _check_support(self) -> dict[OfficeType, bool]:
        """检查Office库支持"""
        support = {}

        for office_type in OfficeType:
            module_name, package_name = OFFICE_LIBS[office_type]
            available = importlib.util.find_spec(module_name) is not None
            support[office_type] = available

            if not available:
                logger.warning(f"[文件生成器] {package_name} 未安装")

        return support

    async def generate(
        self,
        event: AstrMessageEvent,
        office_type: OfficeType,
        filename: str,
        content: dict,
    ):
        """生成Office文件"""
        if not self.support.get(office_type, False):
            await event.send(
                MessageChain().message(
                    f"[文件生成器] {office_type}文件生成不支持，缺少相关库"
                )
            )
            return ""

        try:
            filename = content.get("filename", f"{office_type}_file")
            content = content.get("content", {})

            # 解析content
            if isinstance(content, str):
                try:
                    content = json.loads(content)
                except json.JSONDecodeError:
                    content = self._create_default_content(office_type, content)

            # 清理文件名并添加扩展名
            filename = self._sanitize_filename(filename)
            extension = OFFICE_EXTENSIONS[office_type]

            if not filename.endswith(extension):
                filename = filename + extension

            file_path = self._get_unique_filepath(filename)
            generator = getattr(self, self._GENERATORS[office_type])
            await generator(file_path, content)

            logger.info(f"[文件生成器] Office文件已生成: {file_path}")
            return file_path

        except Exception as e:
            logger.error(f"[文件生成器] 生成Office文件失败: {e}", exc_info=True)
            return ""

    def _create_default_content(self, file_type: OfficeType, text: str) -> dict:
        """创建默认的内容结构，智能解析文本格式"""
        if file_type == OfficeType.WORD:
            return self._parse_word_content(text)
        elif file_type == OfficeType.EXCEL:
            return self._parse_excel_content(text)
        elif file_type == OfficeType.POWERPOINT:
            return self._parse_ppt_content(text)
        return {}

    def _parse_word_content(self, text: str) -> dict:
        """解析 Word 内容格式

        支持：
        - 空行分隔段落
        - 第一行作为标题（如果后面有空行）
        """
        text = text.strip()
        if not text:
            return {"paragraphs": ["（空文档）"]}

        # 按空行分割段落
        paragraphs = []
        current_para = []

        for line in text.split("\n"):
            if line.strip():
                current_para.append(line.strip())
            else:
                if current_para:
                    paragraphs.append(" ".join(current_para))
                    current_para = []

        if current_para:
            paragraphs.append(" ".join(current_para))

        if not paragraphs:
            return {"paragraphs": [text]}

        # 如果只有一个段落，直接返回
        if len(paragraphs) == 1:
            return {"paragraphs": paragraphs}

        # 第一段作为标题
        title = paragraphs[0]
        body_paragraphs = paragraphs[1:]

        return {"title": title, "paragraphs": body_paragraphs}

    def _parse_excel_content(self, text: str) -> dict:
        """解析 Excel 内容格式

        支持：
        - 用 | 分隔每个单元格
        - 换行分隔每一行
        """
        lines = [line.strip() for line in text.strip().split("\n") if line.strip()]
        if not lines:
            return {"sheets": [{"name": "Sheet1", "data": [["（空表格）"]]}]}

        data = []
        for line in lines:
            if "|" in line:
                cells = [cell.strip() for cell in line.split("|")]
            else:
                cells = [line]
            data.append(cells)

        return {"sheets": [{"name": "Sheet1", "data": data}]}

    def _parse_ppt_content(self, text: str) -> dict:
        """解析 PPT 内容格式

        支持：
        - 用 [幻灯片 N] 或 [Slide N] 标记每页
        - 标记后第一行作为标题
        - 后续行作为内容要点
        """
        import re

        text = text.strip()
        if not text:
            return {"slides": [{"title": "空白幻灯片", "content": []}]}

        # 匹配 [幻灯片 N]、[幻灯片N]、[Slide N]、[slide N] 等格式
        slide_pattern = r"\[(?:幻灯片|Slide|slide)\s*\d+\]"

        # 检查是否有幻灯片标记
        if not re.search(slide_pattern, text):
            # 没有标记，尝试按空行分割成多页
            return self._parse_ppt_by_blank_lines(text)

        # 按幻灯片标记分割
        slides = []
        parts = re.split(f"({slide_pattern})", text)

        current_slide = None
        for part in parts:
            part = part.strip()
            if not part:
                continue

            if re.match(slide_pattern, part):
                # 这是一个幻灯片标记，准备新幻灯片
                if current_slide:
                    slides.append(current_slide)
                current_slide = {"title": "", "content": []}
            elif current_slide is not None:
                # 解析幻灯片内容
                lines = [ln.strip() for ln in part.split("\n") if ln.strip()]
                if lines:
                    current_slide["title"] = lines[0]
                    current_slide["content"] = lines[1:] if len(lines) > 1 else []

        if current_slide:
            slides.append(current_slide)

        if not slides:
            return {"slides": [{"title": "内容", "content": [text]}]}

        return {"slides": slides}

    def _parse_ppt_by_blank_lines(self, text: str) -> dict:
        """按空行分割 PPT 内容为多页"""
        slides = []
        current_lines = []

        for line in text.split("\n"):
            if line.strip():
                current_lines.append(line.strip())
            else:
                if current_lines:
                    title = current_lines[0]
                    content = current_lines[1:] if len(current_lines) > 1 else []
                    slides.append({"title": title, "content": content})
                    current_lines = []

        if current_lines:
            title = current_lines[0]
            content = current_lines[1:] if len(current_lines) > 1 else []
            slides.append({"title": title, "content": content})

        if not slides:
            return {"slides": [{"title": "内容", "content": [text]}]}

        return {"slides": slides}

    def _generate_word_sync(self, file_path: Path, content: dict):
        """生成Word文档"""
        from docx import Document

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

    def _generate_excel_sync(self, file_path: Path, content: dict):
        """生成Excel表格"""
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Font, PatternFill

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

    def _generate_ppt_sync(self, file_path: Path, content: dict):
        """生成PowerPoint演示文稿"""
        from pptx import Presentation

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
        """获取文件路径（覆盖已存在的同名文件）"""
        return self.data_path / filename
