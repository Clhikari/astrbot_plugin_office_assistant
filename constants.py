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

# 旧版 Office 格式也映射到对应类型（用于读取）
LEGACY_OFFICE_SUFFIXES = {
    ".doc": OfficeType.WORD,
    ".xls": OfficeType.EXCEL,
    ".ppt": OfficeType.POWERPOINT,
}
# 合并新旧格式映射
SUFFIX_TO_OFFICE_TYPE.update(LEGACY_OFFICE_SUFFIXES)

# 所有可读取的 Office 格式（新旧）
ALL_OFFICE_SUFFIXES = frozenset(SUFFIX_TO_OFFICE_TYPE.keys())

DEFAULT_MAX_FILE_SIZE_MB = 20
DEFAULT_CHUNK_SIZE = 64 * 1024  # 64 KB
# 所有文件类工具名称，用于 before_llm_chat 中的可见性控制。
# 同步点：@llm_tool(name=...) 定义在 main.py，
#         document tool 名称定义在 agent_tools/document_tools.py 的 name 字段。
FILE_TOOLS = [
    "read_file",
    "create_office_file",
    "create_document",
    "add_blocks",
    "finalize_document",
    "export_document",
    "convert_to_pdf",
    "convert_from_pdf",
]

EXECUTION_TOOLS = [
    "astrbot_execute_shell",
    "astrbot_execute_python",
    "astrbot_execute_ipython",
]

NOTICE_TOOLS_DENIED = (
    "\n[System Notice] 当前聊天不可使用文件/Office/PDF 相关功能。"
    " 请用中文说明，并建议用户改为私聊或让管理员开启功能。"
    " 不要调用文件工具，也不要使用 `astrbot_execute_python`、`astrbot_execute_shell`"
    " 或 `astrbot_execute_ipython` 绕过限制。"
)

NOTICE_DOCUMENT_TOOLS_GUIDE = (
    "\n[System Notice] 处理 Word 文档时，请使用有状态文档工具链："
    " `create_document(theme_name=..., table_template=..., density=..., accent_color=...)`"
    " -> `add_blocks(blocks=[...])`"
    " -> `finalize_document` -> `export_document`。"
    " 最好按章节或逻辑块调用 `add_blocks`。"
    " 主题优先 `business_report`、`project_review`、`executive_brief`；"
    " 表格优先 `report_grid`、`metrics_compact`、`minimal`。"
    " 紧凑版式用 `density=compact`，品牌色用 `accent_color=RRGGBB`。"
    " 预设不够时，再用 `style={align, emphasis, font_scale, table_grid, cell_align}`"
    " 和 `layout={spacing_before, spacing_after}`。"
    " 一旦开始使用文档工具，在导出前不要停在普通助手回复上。"
    " 继续调用文档工具，直到 `export_document` 成功。"
    " `export_document` 会直接发送文件。"
    " 如果用户请求依赖上传的可读文件，先调用 `read_file`，再调用 `create_document`"
    " 或 `create_office_file`。"
    " 如果 `read_file` 返回文件不存在或路径非法，不要调用网络搜索；"
    " 直接请用户重新上传文件或提供正确的本地路径。"
    " `create_office_file` 只用于简单的一次性输出（Excel/PPT）。"
    " 如果在调用工具前需要先给用户一句过渡说明，也请使用中文，不要先输出英文说明。"
)

NOTICE_UPLOADED_FILE_TEMPLATE = (
    "\n[System Notice] 已收到上传的{type_desc}：{original_name}（后缀：{file_suffix}），"
    " 文件已保存到工作区。"
    " 如果用户请求依赖这个上传文件，先调用 `read_file`，再调用 `create_document`"
    " 或 `create_office_file`。"
    " 在至少读取一次上传源文件之前，不要先创建新文档。"
    " 如果 `read_file` 返回文件不存在或路径非法，不要调用网络搜索；"
    " 直接请用户重新上传文件或提供正确的本地路径。"
    " 仅在任务仍不清楚时，才先追问用户。"
    " 如果要先给用户一句过渡说明，也请使用中文，不要先输出英文说明。"
    " 对于复杂 Word 输出，优先使用有状态文档工具，而不是 `create_office_file`。"
)

MSG_DOCUMENT_EXPORTED = "✅ 文档已导出"

# PDF 相关常量
PDF_SUFFIX = ".pdf"
PDF_LIBS = {
    "pdf2docx": ("pdf2docx", "pdf2docx"),
    "tabula": ("tabula", "tabula-py"),
    "pdfplumber": ("pdfplumber", "pdfplumber"),
}

# 支持转换为 PDF 的格式（包括旧版 Office 格式）
CONVERTIBLE_TO_PDF = frozenset(
    {
        ".docx",
        ".doc",  # Word
        ".xlsx",
        ".xls",  # Excel
        ".pptx",
        ".ppt",  # PowerPoint
    }
)

# 支持从 PDF 转换的目标格式
PDF_TARGET_FORMATS = {
    "word": (".docx", "Word 文档"),
    "excel": (".xlsx", "Excel 表格"),
}
