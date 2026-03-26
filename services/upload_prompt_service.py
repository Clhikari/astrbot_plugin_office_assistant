class UploadPromptService:
    def __init__(self, *, allow_external_input_files: bool) -> None:
        self._allow_external_input_files = allow_external_input_files

    def build_prompt(
        self,
        *,
        upload_infos: list[dict],
        user_instruction: str,
    ) -> str:
        file_info_list = [self._format_file_info(info) for info in upload_infos]
        has_readable_file = any(info["is_supported"] for info in upload_infos)
        relative_path_guidance = self._build_relative_path_guidance()

        if has_readable_file and user_instruction:
            return (
                f"\n[System Notice] 用户上传了 {len(file_info_list)} 个文件\n"
                + "\n"
                + "[文件信息]\n"
                + "\n".join(f"- {info}" for info in file_info_list)
                + "\n"
                + "\n"
                + "[用户指令]\n"
                + f"{user_instruction}\n"
                + "\n"
                + "[处理建议]\n"
                + "1. 优先围绕这些上传文件完成用户请求。\n"
                + "2. 先调用 `read_file` 读取文件，不要自行猜测文件名，也不要列目录或调用 shell。\n"
                + relative_path_guidance
                + "4. 所有面向用户的回复 MUST 使用中文。"
            )

        if has_readable_file:
            return (
                f"\n[System Notice] 用户上传了 {len(file_info_list)} 个文件\n"
                + "\n"
                + "[文件信息]\n"
                + "\n".join(f"- {info}" for info in file_info_list)
                + "\n"
                + "\n"
                + "[处理建议]\n"
                + "1. 用户上传了可读取文件，后续应优先围绕这些文件处理。\n"
                + "2. 如果要读取文件，不要自行猜测文件名，也不要列目录或调用 shell。\n"
                + relative_path_guidance
                + "4. 用户意图尚不明确时，再用中文询问用户想要如何处理。"
            )

        return (
            f"\n[System Notice] 用户上传了 {len(file_info_list)} 个文件\n"
            "\n"
            "[文件信息]\n" + "\n".join(f"- {info}" for info in file_info_list) + "\n"
            "\n"
            "[操作要求]\n"
            "请根据用户要求处理这些文件，使用中文与用户沟通。"
        )

    def _format_file_info(self, info: dict) -> str:
        file_lines = [
            f"原始文件名: {info['original_name']} (类型: {info['file_suffix']})"
        ]
        if info["stored_name"]:
            file_lines.append(f"  工作区文件名: {info['stored_name']}")
        if self._allow_external_input_files and info["source_path"]:
            file_lines.append(f"  外部绝对路径: {info['source_path']}")
        return "\n".join(file_lines)

    def _build_relative_path_guidance(self) -> str:
        if self._allow_external_input_files:
            return (
                "3. 若使用相对路径，请使用上面的工作区文件名；"
                "如果已提供外部绝对路径，则可直接使用该绝对路径。\n"
            )
        return "3. 若使用相对路径，请使用上面的工作区文件名。\n"
