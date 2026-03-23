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
    "\n[System Notice] 当前聊天不可使用文件/Office/PDF 相关功能。\n"
    "1. 用中文告知用户当前聊天无法使用文件功能，建议私聊或让管理员开启。\n"
    "2. NEVER 调用任何文件工具。\n"
    "3. NEVER 使用 `astrbot_execute_python`、`astrbot_execute_shell`"
    " 或 `astrbot_execute_ipython` 绕过限制。\n"
    "4. NEVER 尝试任何变通方案来绕过以上限制。"
)

NOTICE_DOCUMENT_TOOLS_GUIDE = (
    "\n[System Notice] 文件工具使用指南\n"
    "\n"
    "[核心工作流]\n"
    "生成 Word 文档 MUST 按以下顺序调用工具链：\n"
    "  `create_document` → `add_blocks`(可多次) → `finalize_document` → `export_document`\n"
    "- `create_document` 参数：theme_name / table_template / density / accent_color\n"
    "- 按章节或逻辑块分批调用 `add_blocks`\n"
    "- `export_document` 会直接发送文件给用户\n"
    "\n"
    "[工具选择]\n"
    "- 复杂 Word 文档 → 使用上述工具链\n"
    "- 简单的一次性 Excel/PPT → 使用 `create_office_file`\n"
    "- 主题：优先 `business_report`、`project_review`、`executive_brief`\n"
    "- 表格样式：优先 `report_grid`、`metrics_compact`、`minimal`\n"
    "- 紧凑版式用 `density=compact`，品牌色用 `accent_color=RRGGBB`\n"
    "- 预设不够时再用 `style={align, emphasis, font_scale, table_grid, cell_align}`"
    " 和 `layout={spacing_before, spacing_after}`\n"
    "\n"
    "[约束规则]\n"
    "1. 如果用户请求依赖上传文件，MUST 先调用 `read_file` 读取内容，再创建文档。\n"
    "2. 如果 `read_file` 返回文件不存在或路径非法，NEVER 调用网络搜索；直接请用户重新上传。\n"
    "3. 一旦开始使用文档工具链，MUST 持续调用直到 `export_document` 成功，中途不要停下来发自然语言回复。\n"
    "4. 所有面向用户的回复和过渡说明 MUST 使用中文。"
)

NOTICE_UPLOADED_FILE_TEMPLATE = (
    "\n[System Notice] [ACTION REQUIRED] 已收到上传文件\n"
    "- 文件类型：{type_desc}\n"
    "- 原始文件名：{original_name}（后缀：{file_suffix}）\n"
    "- 工作区文件名：{stored_name}\n"
    "- 状态：已保存到工作区\n"
    "\n"
    "[操作要求]\n"
    "1. MUST 先调用 `read_file` 读取此文件，读取时优先使用工作区文件名 `{stored_name}`。在读取前 NEVER 创建新文档。\n"
    "2. 如果用户意图明确，读取后按需处理；如果意图不清楚，读取后用中文追问用户。\n"
    "3. 所有面向用户的回复 MUST 使用中文。"
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
