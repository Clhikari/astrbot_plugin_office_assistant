from ...constants import (
    DOCUMENT_BLOCK_FONT_SCALE_MAX,
    DOCUMENT_BLOCK_FONT_SCALE_MIN,
    DOCUMENT_BLOCK_SPACING_MAX,
    DOCUMENT_BLOCK_SPACING_MIN,
)


def build_document_tools_guide_notice() -> str:
    return build_document_tools_core_notice() + build_document_tools_detail_notice()


def build_document_tools_core_notice() -> str:
    return (
        "\n[System Notice] 文件工具使用指南\n"
        "\n"
        "[核心工作流]\n"
        "生成 Word 文档 MUST 按以下顺序调用工具链：\n"
        "  `create_document` → `add_blocks`(可多次) → `finalize_document` → `export_document`\n"
        "- 同一份文档在拿到 `document_id` 后，不要再次调用 `create_document`，继续对该 "
        "`document_id` 调用 `add_blocks`\n"
        "- `create_document` 参数：theme_name / table_template / density / accent_color\n"
        "- 按章节或逻辑块分批调用 `add_blocks`\n"
        "- 只要还有任何章节、表格或补充信息没写完，继续调用 `add_blocks`，不要提前调用 "
        "`finalize_document`\n"
        "- 只有确认整份文档内容已经写完，才调用 `finalize_document`\n"
        "- 一旦 `finalize_document` 成功，下一步只能调用 `export_document`，不要再回到 "
        "`add_blocks` 或 `create_document`\n"
        "- `export_document` 会直接发送文件给用户\n"
        "\n"
        "[约束规则]\n"
        "1. 如果用户请求依赖上传文件，MUST 先调用 `read_file` 读取内容，再创建文档。\n"
        "2. 如果 `read_file` 返回文件不存在或路径非法，NEVER 调用网络搜索；直接请用户重新上传。\n"
        "3. 如果用户显式指定了某个工具名和参数，MUST 先按该工具调用；即使预期会报错，也不要擅自改调其他工具，也不要自行修改参数后重试。\n"
        "4. 一旦开始使用文档工具链，MUST 持续调用直到 `export_document` 成功，中途不要停下来发自然语言回复。\n"
        "5. 所有面向用户的回复和过渡说明 MUST 使用中文。"
    )


def build_document_follow_up_notice(
    *,
    document_id: str,
    status: str,
    block_count: int,
) -> str:
    if status == "draft":
        return (
            "\n[System Notice] 当前文档阶段\n"
            f"- 当前 `document_id={document_id}` 仍是 draft，已写入 {block_count} 个内容块\n"
            "- 如果还有任何内容没写完，继续调用 `add_blocks`\n"
            "- 只有确认内容全部写完，才调用 `finalize_document`\n"
        )
    if status == "finalized":
        return (
            "\n[System Notice] 当前文档阶段\n"
            f"- 当前 `document_id={document_id}` 已定稿\n"
            "- 下一步只能调用 `export_document`\n"
            "- 不要再调用 `add_blocks`、`finalize_document` 或 `create_document`\n"
        )
    if status == "exported":
        return (
            "\n[System Notice] 当前文档阶段\n"
            f"- 当前 `document_id={document_id}` 已导出\n"
            "- 不要继续对这份文档调用 `add_blocks`、`finalize_document` 或 `export_document`\n"
        )
    return (
        "\n[System Notice] 当前文档阶段\n"
        f"- 当前 `document_id={document_id}` 状态未知\n"
        "- 先核对文档状态，不要切换到别的文档工具\n"
    )


def build_document_follow_up_missing_notice(*, document_id: str) -> str:
    return (
        "\n[System Notice] 当前文档阶段\n"
        f"- 没有找到 `document_id={document_id}` 对应的文档会话\n"
        "- 先核对 `document_id`，不要改调其他文档工具\n"
    )


def build_document_tools_detail_notice() -> str:
    return (
        "\n[System Notice] 文件工具细节指南\n"
        "\n"
        "[工具选择]\n"
        "- 复杂 Word 文档 → 使用上述工具链\n"
        "- 简单的一次性 Excel/PPT → 使用 `create_office_file`\n"
        "- 主题：优先 `business_report`、`project_review`、`executive_brief`\n"
        "- 表格样式：优先 `report_grid`、`metrics_compact`、`minimal`\n"
        "- 紧凑版式用 `density=compact`，品牌色用 `accent_color=RRGGBB`\n"
        "- 如果用户描述的是整份文档气质，例如深蓝商务风、浅灰极简风、留白更克制，可在 "
        "`create_document` 里使用 `document_style={brief, heading_color, title_align, "
        "body_font_size, body_line_spacing, paragraph_space_after, list_space_after, "
        "summary_card_defaults, table_defaults}`\n"
        "- 如果用户要求深蓝商务风、浅灰极简风、首列强调、浅色斑马纹等表格效果，可在 table 块上使用 "
        "`header_fill`、`header_text_color`、`banded_rows`、`banded_row_fill`、"
        "`first_column_bold`、`table_align`、`border_style`、`caption_emphasis`\n"
        "- 预设不够时再用 `style={align, emphasis, font_scale, table_grid, cell_align}`"
        " 和 `layout={spacing_before, spacing_after}`\n"
        f"- `style.font_scale` 建议保持在 {DOCUMENT_BLOCK_FONT_SCALE_MIN} 到 "
        f"{DOCUMENT_BLOCK_FONT_SCALE_MAX} 之间，`layout.spacing_before` 和 "
        f"`layout.spacing_after` 保持在 {int(DOCUMENT_BLOCK_SPACING_MIN)} 到 "
        f"{int(DOCUMENT_BLOCK_SPACING_MAX)} 之间\n"
        "- 标题颜色属于整份文档风格，请在 `create_document.document_style.heading_color` "
        "里设置；不要把 `heading_color` 写到单个 `heading` block 上\n"
        "- 横向页面、节级页眉页脚、页码重置 MUST 使用独立的 `section_break` block；"
        "不要把 `page_orientation`、`start_type`、`restart_page_numbering`、"
        "`page_number_start`、`header_footer` 写到 `heading`、`paragraph` 或 `table` 上\n"
        "- 宽表推荐步骤：先单独插入一个 `section_break` block，设置 "
        "`start_type=new_page` 和 `page_orientation=landscape`，下一条再写 `table`；"
        "如果后面要恢复纵向，再单独插入一个 `section_break` block，设置 "
        "`start_type=new_page` 和 `page_orientation=portrait`\n"
        "- `toc` 只使用 `title`、`levels`、`start_on_new_page`；不要给 `toc` 传 `text`\n"
        "- 表格列标题使用 `headers`；不要给 `table` 传 `columns`\n"
    )
