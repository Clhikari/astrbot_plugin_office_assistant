import json
from typing import Optional

from astrbot.api.event import AstrMessageEvent
from astrbot.api import logger


class FileAnalyzer:
    """文件生成需求分析器"""

    def __init__(self, context, config_mgr):
        self.context = context
        self.config_mgr = config_mgr

    async def analyze_message(
        self, event: AstrMessageEvent, message: str
    ) -> Optional[dict]:
        """使用 AI 分析消息是否需要生成文件"""
        try:
            umo = event.unified_msg_origin
            provider_id = await self.context.get_current_chat_provider_id(umo=umo)

            if not provider_id:
                logger.warning("[文件生成器] 无法获取聊天模型ID")
                return None

            # 构造分析提示词
            office_hint = ""
            if self.config_mgr.get("enable_office_files", True):
                office_hint = """
            - word: Word文档 (.docx)
            - excel: Excel表格 (.xlsx)
            - powerpoint: PowerPoint演示文稿 (.pptx)"""

            analysis_prompt = f"""请分析以下用户消息，判断用户是否需要生成文件。

            用户消息：
            {message}

            请按照以下JSON格式回复（只返回JSON，不要有其他内容）：
            {{
            "needs_file": true/false,
            "file_info": {{
                "type": "文件类型",
                "filename": "建议的文件名（不含扩展名）",
                "content": "文件内容或结构化数据",
                "description": "简短描述"
            }}
            }}

            支持的文件类型：
            - python, javascript, java, cpp, html, css, json, xml, yaml, markdown, text, csv, sql{office_hint}

            判断标准：
            1. 用户明确要求生成、创建、保存、导出文件
            2. 用户要求输出代码/文档并保存
            3. 用户要求创建Word/Excel/PPT文件
            4. 用户描述了需要写入文件的内容

            对于Office文件，content字段应该包含结构化数据：
            - Word: {{"paragraphs": ["段落1", "段落2"], "title": "标题"}}
            - Excel: {{"sheets": [{{"name": "Sheet1", "data": [["A1", "B1"], ["A2", "B2"]]}}]}}
            - PowerPoint: {{"slides": [{{"title": "标题", "content": ["要点1", "要点2"]}}]}}

            如果不需要生成文件，返回 {{"needs_file": false}}"""

            # 调用 AI
            llm_resp = await self.context.llm_generate(
                chat_provider_id=provider_id,
                prompt=analysis_prompt,
            )

            response_text = llm_resp.completion_text.strip()
            logger.info(f"[文件生成器] AI分析结果: {response_text[:200]}...")

            # 清理markdown代码块标记
            response_text = (
                response_text.replace("```json", "").replace("```", "").strip()
            )

            # 解析 JSON
            try:
                result = json.loads(response_text)
                return result
            except json.JSONDecodeError as e:
                logger.error(f"[文件生成器] JSON解析失败: {e}")
                return None

        except Exception as e:
            logger.error(f"[文件生成器] AI分析失败: {e}", exc_info=True)
            return None

    async def generate_content(
        self, event: AstrMessageEvent, file_type: str, user_description: str
    ) -> str:
        """使用AI生成完整的文件内容"""
        try:
            umo = event.unified_msg_origin
            provider_id = await self.context.get_current_chat_provider_id(umo=umo)

            if not provider_id:
                logger.warning("[文件生成器] 无法获取聊天模型ID，使用原始内容")
                return user_description

            # 根据文件类型构造不同的提示词
            if file_type in ["word", "excel", "powerpoint"]:
                prompt = self._build_office_prompt(file_type, user_description)
            else:
                prompt = f"""请根据用户的描述，生成一个完整的{file_type}文件内容。

            用户描述：{user_description}

            请直接输出文件内容，不要有任何解释或markdown代码块标记。"""

            llm_resp = await self.context.llm_generate(
                chat_provider_id=provider_id,
                prompt=prompt,
            )

            full_content = llm_resp.completion_text.strip()
            # 清理markdown代码块
            full_content = (
                full_content.replace("```json", "").replace("```", "").strip()
            )

            return full_content

        except Exception as e:
            logger.error(f"[文件生成器] AI生成内容失败: {e}")
            return user_description

    def _build_office_prompt(self, file_type: str, description: str) -> str:
        """构建Office文件的AI提示词"""
        prompt = f"""请根据用户描述，生成{file_type}文件所需的结构化数据。

        用户描述：{description}

        请以JSON格式输出："""

        if file_type == "word":
            prompt += """
        {
        "title": "文档标题",
        "paragraphs": ["段落1", "段落2", "..."]
        }"""
        elif file_type == "excel":
            prompt += """
        {
        "sheets": [
            {
            "name": "Sheet1",
            "data": [
                ["列1", "列2"],
                ["数据1", "数据2"]
            ]
            }
        ]
        }"""
        elif file_type == "powerpoint":
            prompt += """
        {
        "slides": [
            {
            "title": "幻灯片标题",
            "content": ["要点1", "要点2"]
            }
        ]
        }"""

        return prompt
