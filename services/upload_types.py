from typing import NotRequired, TypedDict


class UploadInfo(TypedDict):
    original_name: str
    file_suffix: str
    stored_name: str
    source_path: str
    is_supported: bool
    file_id: NotRequired[str]
    type_desc: NotRequired[str]
