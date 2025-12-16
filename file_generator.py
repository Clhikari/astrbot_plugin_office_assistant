from pathlib import Path
from datetime import datetime
from typing import Optional

from astrbot.api import logger


class FileGenerator:
    """普通文件生成器"""

    FILE_EXTENSIONS = {
        "python": ".py",
        "javascript": ".js",
        "typescript": ".ts",
        "java": ".java",
        "cpp": ".cpp",
        "c": ".c",
        "html": ".html",
        "css": ".css",
        "json": ".json",
        "xml": ".xml",
        "yaml": ".yaml",
        "markdown": ".md",
        "text": ".txt",
        "csv": ".csv",
        "sql": ".sql",
        "shell": ".sh",
        "batch": ".bat",
    }

    def __init__(self, data_path: Path):
        self.data_path = data_path

    async def generate(self, file_info: dict) -> Optional[Path]:
        """生成普通文件"""
        try:
            file_type = file_info.get("type", "text").lower()
            filename = file_info.get("filename", "generated_file")
            content = file_info.get("content", "")

            # 清理文件名
            filename = self._sanitize_filename(filename)

            # 获取文件扩展名
            extension = self.FILE_EXTENSIONS.get(file_type, ".txt")
            if not filename.endswith(extension):
                filename = filename + extension

            # 获取唯一文件路径
            file_path = self._get_unique_filepath(filename)

            # 写入文件
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(content)

            logger.info(f"[文件生成器] 文件已生成: {file_path}")
            return file_path

        except Exception as e:
            logger.error(f"[文件生成器] 生成文件失败: {e}", exc_info=True)
            return None

    def _sanitize_filename(self, filename: str) -> str:
        """清理文件名中的非法字符"""
        filename = "".join(
            c for c in filename if c.isalnum() or c in (" ", "-", "_", ".")
        ).strip()

        if not filename:
            filename = f"file_{datetime.now().strftime('%Y%m%d_%H%M%S')}"

        return filename

    def _get_unique_filepath(self, filename: str) -> Path:
        """获取唯一的文件路径（如果文件存在则添加时间戳）"""
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
