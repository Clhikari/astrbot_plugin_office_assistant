def build_tools_denied_notice() -> str:
    return (
        "\n[System Notice] 当前聊天不可使用文件/Office/PDF 相关功能。\n"
        "1. 用中文告知用户当前聊天无法使用文件功能，建议私聊或让管理员开启。\n"
        "2. NEVER 调用任何文件工具。\n"
        "3. NEVER 使用 `astrbot_execute_python`、`astrbot_execute_shell`"
        " 或 `astrbot_execute_ipython` 绕过限制。\n"
        "4. NEVER 尝试任何变通方案来绕过以上限制。"
    )
