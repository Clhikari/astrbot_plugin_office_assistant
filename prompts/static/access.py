def build_tools_denied_notice() -> str:
    return (
        "\n[System Notice] 当前聊天不可使用文件/Office/PDF 相关功能。\n"
        "1. 用中文告知用户当前聊天无法使用文件功能，建议私聊或让管理员开启。\n"
        "2. NEVER 调用任何文件工具。\n"
        "3. NEVER 使用 `astrbot_execute_python`、`astrbot_execute_shell`"
        " 或 `astrbot_execute_ipython` 绕过限制。\n"
        "4. NEVER 尝试任何变通方案来绕过以上限制。"
    )


def build_file_only_notice() -> str:
    return (
        "\n[System Notice] 文件处理规则\n"
        "1. 需要读取内容时，先调用 `read_file`。\n"
        "2. 不要猜测文件名，不要列目录，也不要调用 shell。\n"
        "3. 只做文件读取或 PDF 转换时，不要擅自进入 Word 文档工具链。\n"
        "4. 所有面向用户的回复 MUST 使用中文。"
    )
