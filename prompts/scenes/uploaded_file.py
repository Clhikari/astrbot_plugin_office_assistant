from ...constants import EXCEL_SUFFIXES
from ...services.excel_intent_router import ExcelIntentRouter
from ...services.upload_types import UploadInfo

MAX_DETAILED_UPLOAD_INFOS = 3
_SCRIPT_EDIT_SUFFIXES = frozenset({".xlsx"})
_DEFAULT_EXCEL_TOOL_NAMES = {
    "read_workbook",
    "execute_excel_script",
    "create_workbook",
    "write_rows",
    "export_workbook",
}


def format_minimal_upload_file_info(info: UploadInfo) -> str:
    file_lines = [f"原始文件名: {info['original_name']}"]
    if info["stored_name"]:
        file_lines.append(f"  工作区文件名: {info['stored_name']}")
    return "\n".join(file_lines)


def build_uploaded_file_context_notice(*, upload_infos: list[UploadInfo]) -> str:
    if not upload_infos:
        return ""

    file_info_list, omitted_infos = _build_limited_file_info_list(upload_infos)
    preferred_read_tool = _preferred_read_tool(upload_infos)
    file_requirement = (
        f"先调用 `{preferred_read_tool}` 读取此文件。"
        if len(upload_infos) == 1
        else f"先调用 `{preferred_read_tool}` 依次读取这些文件。"
    )
    return (
        "\n[System Notice] [ACTION REQUIRED] 已收到上传文件\n"
        + f"- 文件数量：{len(upload_infos)}\n\n"
        + "[文件信息]\n"
        + "\n".join(f"- {info}" for info in file_info_list)
        + _build_omitted_upload_info_line(omitted_infos)
        + "\n\n"
        + "[处理要求]\n"
        + f"1. {file_requirement}\n"
        + "2. 不要猜文件名，不要列目录，不要调用 shell。\n"
        + "3. 读取前不要创建新文档。"
    )


def build_buffered_upload_prompt(
    *,
    upload_infos: list[UploadInfo],
    user_instruction: str,
    exposed_tool_names: set[str] | None = None,
) -> str:
    file_info_list, omitted_infos = _build_limited_file_info_list(upload_infos)
    has_readable_file = any(info["is_supported"] for info in upload_infos)
    preferred_tool = _preferred_read_tool(
        upload_infos,
        user_instruction=user_instruction,
        exposed_tool_names=exposed_tool_names,
    )
    preferred_tool_action = _preferred_tool_action(
        preferred_tool,
        multiple=len(upload_infos) > 1,
    )
    file_info_block = (
        f"\n[System Notice] 用户上传了 {len(upload_infos)} 个文件\n\n"
        "[文件信息]\n"
        + "\n".join(f"- {info}" for info in file_info_list)
        + _build_omitted_upload_info_line(omitted_infos)
    )

    if has_readable_file and user_instruction:
        return (
            "\n[用户指令]\n"
            + f"{user_instruction}\n\n"
            + file_info_block
            + "\n\n"
            + "[处理要求]\n"
            + "1. 优先围绕这些上传文件完成用户请求。\n"
            + f"2. {preferred_tool_action}\n"
            + "3. 不要猜文件名，不要列目录，不要调用 shell。\n"
            + "4. 读取后按用户指令继续调用工具，不要只回复过渡说明。"
        )

    if has_readable_file:
        return (
            file_info_block
            + "\n\n"
            + "[处理要求]\n"
            + "1. 用户上传了可读取文件，后续应优先围绕这些文件处理。\n"
            + f"2. 如需读取文件，先调用 `{preferred_tool}`。\n"
            + "3. 不要猜文件名，不要列目录，不要调用 shell。\n"
            + "4. 用户意图尚不明确时，再用中文询问用户想要如何处理。"
        )

    return (
        file_info_block
        + "\n\n"
        + "[操作要求]\n"
        + "请根据用户要求处理这些文件，使用中文与用户沟通。"
    )


def _build_limited_file_info_list(
    upload_infos: list[UploadInfo],
) -> tuple[list[str], list[UploadInfo]]:
    limited_infos = upload_infos[:MAX_DETAILED_UPLOAD_INFOS]
    omitted_infos = upload_infos[MAX_DETAILED_UPLOAD_INFOS:]
    file_info_list = [format_minimal_upload_file_info(info) for info in limited_infos]
    return file_info_list, omitted_infos


def _build_omitted_upload_info_line(omitted_infos: list[UploadInfo]) -> str:
    if not omitted_infos:
        return ""
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


def _preferred_read_tool(
    upload_infos: list[UploadInfo],
    *,
    user_instruction: str = "",
    exposed_tool_names: set[str] | None = None,
) -> str:
    readable_infos = [info for info in upload_infos if info.get("is_supported")]
    if not readable_infos:
        return "read_file"

    if all(
        str(info.get("file_suffix", "")).lower() in _SCRIPT_EDIT_SUFFIXES
        for info in readable_infos
    ):
        available_tool_names = (
            exposed_tool_names if exposed_tool_names is not None else _DEFAULT_EXCEL_TOOL_NAMES
        )
        decision = ExcelIntentRouter.decide(
            request_text=user_instruction,
            upload_infos=readable_infos,
            explicit_tool_name=None,
            exposed_tool_names=available_tool_names,
        )
        if (
            decision is not None
            and decision.requires_script
            and decision.should_inject_guide
        ):
            return "execute_excel_script"
    if all(
        str(info.get("file_suffix", "")).lower() in EXCEL_SUFFIXES
        for info in readable_infos
    ):
        return "read_workbook"
    return "read_file"


def _preferred_tool_action(tool_name: str, *, multiple: bool) -> str:
    if tool_name == "read_file":
        return (
            "先调用 `read_file` 依次读取这些文件。"
            if multiple
            else "先调用 `read_file` 读取文件。"
        )
    if tool_name == "read_workbook":
        return (
            "先调用 `read_workbook` 依次读取这些文件。"
            if multiple
            else "先调用 `read_workbook` 读取文件。"
        )
    return (
        "先调用 `execute_excel_script` 依次处理这些文件。"
        if multiple
        else "先调用 `execute_excel_script` 处理文件。"
    )
