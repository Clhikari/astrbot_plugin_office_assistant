from .prompt_context_service import PromptContextService
from .upload_types import UploadInfo


class UploadPromptService:
    def __init__(
        self,
        *,
        allow_external_input_files: bool,
        prompt_context_service: PromptContextService | None = None,
    ) -> None:
        self._prompt_context_service = prompt_context_service or PromptContextService(
            allow_external_input_files=allow_external_input_files
        )

    def build_prompt(
        self,
        *,
        upload_infos: list[UploadInfo],
        user_instruction: str,
    ) -> str:
        return self._prompt_context_service.build_buffered_upload_prompt(
            upload_infos=upload_infos,
            user_instruction=user_instruction,
        )
