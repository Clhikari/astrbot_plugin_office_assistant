from .prompt_context_service import PromptContextService
from .runtime_config import resolve_computer_runtime_mode
from .upload_types import UploadInfo


class UploadPromptService:
    def __init__(
        self,
        *,
        allow_external_input_files: bool,
        astrbot_context=None,
        auto_block_execution_tools: bool = False,
        prompt_context_service: PromptContextService | None = None,
    ) -> None:
        self._astrbot_context = astrbot_context
        self._auto_block_execution_tools = auto_block_execution_tools
        self._prompt_context_service = prompt_context_service or PromptContextService(
            allow_external_input_files=allow_external_input_files
        )

    def build_prompt(
        self,
        *,
        upload_infos: list[UploadInfo],
        user_instruction: str,
        event=None,
    ) -> str:
        exposed_tool_names = self._get_exposed_tool_names(event)
        return self._prompt_context_service.build_buffered_upload_prompt(
            upload_infos=upload_infos,
            user_instruction=user_instruction,
            exposed_tool_names=exposed_tool_names,
        )

    def _get_exposed_tool_names(self, event) -> set[str] | None:
        if event is None:
            return None

        exposed_tool_names = {
            "read_workbook",
            "execute_excel_script",
            "create_workbook",
            "write_rows",
            "export_workbook",
        }
        try:
            runtime_mode = resolve_computer_runtime_mode(self._astrbot_context, event)
        except Exception:
            return exposed_tool_names
        if runtime_mode not in {"local", "sandbox"}:
            exposed_tool_names.discard("execute_excel_script")
            return exposed_tool_names
        if self._auto_block_execution_tools and runtime_mode != "sandbox":
            exposed_tool_names.discard("execute_excel_script")
        return exposed_tool_names
