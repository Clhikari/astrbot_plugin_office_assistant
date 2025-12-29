"""
静态常量
"""

from enum import Enum, auto


class OfficeType(Enum):
    WORD = auto()
    EXCEL = auto()
    POWERPOINT = auto()


OFFICE_EXTENSIONS = {
    OfficeType.WORD: ".docx",
    OfficeType.EXCEL: ".xlsx",
    OfficeType.POWERPOINT: ".pptx",
}

OFFICE_LIBS = {
    OfficeType.WORD: ("docx", "python-docx"),
    OfficeType.EXCEL: ("openpyxl", "openpyxl"),
    OfficeType.POWERPOINT: ("pptx", "python-pptx"),
}

# 字符串映射
OFFICE_TYPE_MAP = {
    "word": OfficeType.WORD,
    "excel": OfficeType.EXCEL,
    "powerpoint": OfficeType.POWERPOINT,
}

TEXT_SUFFIXES = frozenset(
    {
        ".txt",
        ".md",
        ".log",
        ".rst",
        ".py",
        ".js",
        ".ts",
        ".jsx",
        ".tsx",
        ".c",
        ".cpp",
        ".h",
        ".java",
        ".go",
        ".rs",
        ".json",
        ".yaml",
        ".yml",
        ".toml",
        ".xml",
        ".csv",
        ".html",
        ".css",
        ".sql",
        ".sh",
        ".bat",
    }
)

# 默认值
OFFICE_SUFFIXES = frozenset(OFFICE_EXTENSIONS.values())
SUFFIX_TO_OFFICE_TYPE = {v: k for k, v in OFFICE_EXTENSIONS.items()}
DEFAULT_MAX_FILE_SIZE_MB = 50
DEFAULT_CHUNK_SIZE = 64 * 1024  # 64 KB
FILE_TOOLS = [
    "list_files",
    "read_file",
    "write_file",
    "delete_file",
    "create_office_file",
]
PLUGIN_NAME = "astrbot_plugin_file_operation_tool"
