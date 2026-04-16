def build_workbook_tools_guide_notice() -> str:
    return build_workbook_tools_core_notice() + build_workbook_tools_detail_notice()


def build_workbook_tools_core_notice() -> str:
    return (
        "\n[System Notice] Excel 原语工具使用指南\n"
        "\n"
        "[核心工作流]\n"
        "生成结构化 Excel MUST 按以下顺序调用工具链：\n"
        "  `create_workbook` → `write_rows`(可多次) → `export_workbook`\n"
        "- 同一份工作簿在拿到 `workbook_id` 后，不要再次调用 `create_workbook`\n"
        "- `write_rows` 可按 Sheet 分批调用；多 Sheet 场景继续重复调用 `write_rows`\n"
        "- `sheet` 不存在时自动创建，`start_row` 使用 1-based 行号\n"
        "- 不要把整本工作簿一次性塞进 `create_workbook`\n"
        "- 只有确认数据写完才调用 `export_workbook`\n"
        "- `export_workbook` 会直接发送 `.xlsx` 文件给用户\n"
        "\n"
        "[边界说明]\n"
        "- 这条路径只覆盖结构化表格写入，不覆盖公式、图表、条件格式、数据验证\n"
        "- `create_office_file` 仍是旧的简单一次性入口，不等价于上述结构化工作流\n"
    )


def build_workbook_follow_up_notice(
    *,
    workbook_id: str,
    status: str,
    sheet_names: list[str],
    sheet_count: int,
    latest_written_sheets: list[str],
    next_allowed_actions: list[str],
) -> str:
    normalized_sheet_count = max(sheet_count, 0)
    normalized_sheet_names = ", ".join(sheet_names) if sheet_names else "无"
    normalized_latest_sheets = (
        ", ".join(latest_written_sheets) if latest_written_sheets else "无"
    )
    normalized_actions = ", ".join(next_allowed_actions) if next_allowed_actions else "无"

    if status == "draft":
        return (
            "\n[System Notice] 当前工作簿阶段\n"
            f"- 当前 `workbook_id={workbook_id}` 仍是 draft\n"
            f"- Sheet 数量：{normalized_sheet_count}，当前 Sheet：{normalized_sheet_names}\n"
            f"- 最近写入 Sheet：{normalized_latest_sheets}\n"
            f"- 下一步允许动作：{normalized_actions}\n"
            "- 如果还有数据没写完，继续调用 `write_rows`\n"
            "- 只有确认内容写完，才调用 `export_workbook`\n"
        )
    if status == "exported":
        return (
            "\n[System Notice] 当前工作簿阶段\n"
            f"- 当前 `workbook_id={workbook_id}` 已导出\n"
            "- 不要继续对这份工作簿调用 `write_rows` 或 `export_workbook`\n"
        )
    return (
        "\n[System Notice] 当前工作簿阶段\n"
        f"- 当前 `workbook_id={workbook_id}` 状态未知\n"
        f"- 当前 Sheet：{normalized_sheet_names}\n"
        f"- 下一步允许动作：{normalized_actions}\n"
        "- 先核对工作簿状态，不要切换到别的 Excel 工具\n"
    )


def build_workbook_follow_up_missing_notice(*, workbook_id: str) -> str:
    return (
        "\n[System Notice] 当前工作簿阶段\n"
        f"- 没有找到 `workbook_id={workbook_id}` 对应的工作簿会话\n"
        "- 先核对 `workbook_id`，不要改调其他 Excel 工具\n"
    )


def build_workbook_tools_detail_notice() -> str:
    return (
        "\n[System Notice] Excel 原语工具细节指南\n"
        "\n"
        "[工具选择]\n"
        "- 结构化 Excel（多 Sheet、分批写入）→ 使用 `create_workbook` / `write_rows` / `export_workbook`\n"
        "- 简单一次性 Excel/PPT → 使用 `create_office_file`\n"
        "- `write_rows` 的 `rows` 传二维数组，避免超大单次 JSON\n"
        "- 多 Sheet 场景推荐每个 Sheet 分批写入，按需设置 `start_row`\n"
        "- 需要公式、图表、条件格式、数据验证时，不要继续扩展这条原语路径\n"
    )
