from ...services.upload_types import UploadInfo

MAX_DETAILED_UPLOAD_INFOS = 3


def format_upload_file_info(
    info: UploadInfo,
    *,
    allow_external_input_files: bool,
) -> str:
    file_lines = [f"原始文件名: {info['original_name']} (类型: {info['file_suffix']})"]
    if info["stored_name"]:
        file_lines.append(f"  工作区文件名: {info['stored_name']}")
    if allow_external_input_files and info["source_path"]:
        file_lines.append(f"  外部绝对路径: {info['source_path']}")
    return "\n".join(file_lines)


def build_uploaded_file_notice(
    *,
    type_desc: str,
    original_name: str,
    file_suffix: str,
    stored_name: str,
    source_path: str,
    allow_external_input_files: bool,
) -> str:
    return (
        "\n[System Notice] [ACTION REQUIRED] 已收到上传文件\n"
        f"- 文件类型：{type_desc}\n"
        f"- 原始文件名：{original_name}（后缀：{file_suffix}）\n"
        f"- 工作区文件名：{stored_name}\n"
        f"{_build_external_path_line(source_path=source_path, allow_external_input_files=allow_external_input_files)}"
        "- 状态：已保存到工作区\n"
        "\n"
        "[路径提示]\n"
        f"- 必须使用工作区文件名 `{stored_name}`\n"
        f"- 读取时请使用工作区文件名：`{stored_name}`\n"
        f"- 不要使用原始文件名：`{original_name}`\n"
        f"{_build_external_path_hint(source_path=source_path, allow_external_input_files=allow_external_input_files)}"
    )


def build_uploaded_file_scene_notice(
    *,
    file_count: int,
    allow_external_input_files: bool,
) -> str:
    file_requirement = (
        "MUST 先调用 `read_file` 读取此文件"
        if file_count == 1
        else "MUST 先调用 `read_file` 依次读取这些文件"
    )
    return (
        "\n[操作要求]\n"
        f"1. {file_requirement}，不要自行猜测文件名，也不要列目录或调用 shell。"
        f"{_build_relative_path_rule(allow_external_input_files=allow_external_input_files)}"
        "在读取前 NEVER 创建新文档。\n"
        "2. 如果用户意图明确，读取后按需处理；如果意图不清楚，读取后用中文追问用户。\n"
        "3. 所有面向用户的回复 MUST 使用中文。\n"
    )


def build_uploaded_file_summary_notice(
    *,
    upload_infos: list[UploadInfo],
    allow_external_input_files: bool,
) -> str:
    readable_infos = [info for info in upload_infos if info["is_supported"]]
    if not readable_infos:
        return ""

    file_info_list, omitted_infos = _build_limited_file_info_list(
        readable_infos,
        allow_external_input_files=allow_external_input_files,
    )

    return (
        "\n[System Notice] [ACTION REQUIRED] 已收到上传文件\n"
        + f"- 文件数量：{len(readable_infos)}\n"
        + "- 状态：已保存到工作区\n"
        + "\n"
        + "[文件信息]\n"
        + "\n".join(f"- {info}" for info in file_info_list)
        + _build_omitted_upload_info_line(
            omitted_infos,
            allow_external_input_files=allow_external_input_files,
        )
        + "\n\n"
        + "[路径提示]\n"
        + "- 若使用相对路径，请使用上面的工作区文件名。\n"
        + _build_multi_file_external_path_hint(
            allow_external_input_files=allow_external_input_files
        )
    )


def build_buffered_upload_prompt(
    *,
    upload_infos: list[UploadInfo],
    user_instruction: str,
    allow_external_input_files: bool,
) -> str:
    file_info_list, omitted_infos = _build_limited_file_info_list(
        upload_infos,
        allow_external_input_files=allow_external_input_files,
    )
    has_readable_file = any(info["is_supported"] for info in upload_infos)
    relative_path_guidance = _build_relative_path_guidance(
        allow_external_input_files=allow_external_input_files
    )
    omitted_info_line = _build_omitted_upload_info_line(
        omitted_infos,
        allow_external_input_files=allow_external_input_files,
    )

    if has_readable_file and user_instruction:
        return (
            f"\n[System Notice] 用户上传了 {len(upload_infos)} 个文件\n"
            + "\n"
            + "[文件信息]\n"
            + "\n".join(f"- {info}" for info in file_info_list)
            + omitted_info_line
            + "\n"
            + "\n"
            + "[用户指令]\n"
            + f"{user_instruction}\n"
            + "\n"
            + "[处理建议]\n"
            + "1. 优先围绕这些上传文件完成用户请求。\n"
            + "2. 先调用 `read_file` 读取文件，不要自行猜测文件名，也不要列目录或调用 shell。\n"
            + relative_path_guidance
            + "4. 如果用户已经明确要求整理成正式汇报、报告、文档或 Word 文件，读取后继续调用相应工具完成结果，不要停下来只回复过渡说明。\n"
            + "5. 所有面向用户的回复 MUST 使用中文。"
        )

    if has_readable_file:
        return (
            f"\n[System Notice] 用户上传了 {len(upload_infos)} 个文件\n"
            + "\n"
            + "[文件信息]\n"
            + "\n".join(f"- {info}" for info in file_info_list)
            + omitted_info_line
            + "\n"
            + "\n"
            + "[处理建议]\n"
            + "1. 用户上传了可读取文件，后续应优先围绕这些文件处理。\n"
            + "2. 如果要读取文件，不要自行猜测文件名，也不要列目录或调用 shell。\n"
            + relative_path_guidance
            + "4. 用户意图尚不明确时，再用中文询问用户想要如何处理。"
        )

    return (
        f"\n[System Notice] 用户上传了 {len(upload_infos)} 个文件\n"
        "\n"
        "[文件信息]\n"
        + "\n".join(f"- {info}" for info in file_info_list)
        + omitted_info_line
        + "\n"
        "\n"
        "[操作要求]\n"
        "请根据用户要求处理这些文件，使用中文与用户沟通。"
    )


def _build_relative_path_guidance(*, allow_external_input_files: bool) -> str:
    if allow_external_input_files:
        return (
            "3. 若使用相对路径，请使用上面的工作区文件名；"
            "如果已提供外部绝对路径，则可直接使用该绝对路径。\n"
        )
    return "3. 若使用相对路径，请使用上面的工作区文件名。\n"


def _build_relative_path_rule(*, allow_external_input_files: bool) -> str:
    if allow_external_input_files:
        return "若使用相对路径，请使用上面给出的工作区文件名；如果已提供外部绝对路径，也可以直接使用。"
    return "若使用相对路径，请使用上面给出的工作区文件名。"


def _build_external_path_line(
    *,
    source_path: str,
    allow_external_input_files: bool,
) -> str:
    if allow_external_input_files and source_path:
        return f"- 外部绝对路径：{source_path}\n"
    return ""


def _build_external_path_hint(
    *,
    source_path: str,
    allow_external_input_files: bool,
) -> str:
    if allow_external_input_files and source_path:
        return f"- 也可以直接使用外部绝对路径：`{source_path}`\n"
    return "- 当前未启用外部绝对路径，不要使用工作区外路径。\n"


def _build_multi_file_external_path_hint(*, allow_external_input_files: bool) -> str:
    if allow_external_input_files:
        return "- 已提供外部绝对路径时，也可以直接使用上面给出的绝对路径。\n"
    return "- 当前未启用外部绝对路径，不要使用工作区外路径。\n"


def _build_limited_file_info_list(
    upload_infos: list[UploadInfo],
    *,
    allow_external_input_files: bool,
) -> tuple[list[str], list[UploadInfo]]:
    limited_infos = upload_infos[:MAX_DETAILED_UPLOAD_INFOS]
    omitted_infos = upload_infos[MAX_DETAILED_UPLOAD_INFOS:]
    file_info_list = [
        format_upload_file_info(
            info,
            allow_external_input_files=allow_external_input_files,
        )
        for info in limited_infos
    ]
    return file_info_list, omitted_infos


def _build_omitted_upload_info_line(
    omitted_infos: list[UploadInfo],
    *,
    allow_external_input_files: bool,
) -> str:
    if not omitted_infos:
        return ""
    if allow_external_input_files:
        rendered_items = [
            rendered_item
            for info in omitted_infos
            if (rendered_item := _build_compact_upload_path_item(info))
        ]
        if rendered_items:
            return (
                f"\n- 其余 {len(rendered_items)} 个文件："
                + "；".join(rendered_items)
                + "（未展开详细信息）"
            )
    display_names = [
        display_name
        for info in omitted_infos
        if (display_name := _display_upload_name(info))
    ]
    if display_names:
        return (
            f"\n- 其余 {len(display_names)} 个文件："
            + "、".join(display_names)
            + "（未展开详细信息）"
        )
    return f"\n- 其余 {len(omitted_infos)} 个文件未展开详细信息"


def _display_upload_name(info: UploadInfo) -> str:
    return info.get("stored_name") or info.get("original_name") or ""


def _build_compact_upload_path_item(info: UploadInfo) -> str:
    display_name = _display_upload_name(info)
    source_path = info.get("source_path") or ""
    if display_name and source_path:
        return f"{display_name} -> {source_path}"
    return display_name
