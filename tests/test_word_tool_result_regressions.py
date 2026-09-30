import base64
from types import SimpleNamespace

import mcp
import pytest
from tests._docx_test_helpers import _write_png

from astrbot_plugin_office_assistant.services.word_read_service import WordReadService
from astrbot_plugin_office_assistant.services.workspace_service import WorkspaceService


@pytest.mark.asyncio
@pytest.mark.parametrize("image_mode", ["inline", "absent", "disabled", "limited"])
async def test_real_sdk_permission_guard_preserves_word_text_and_images(
    tmp_path, image_mode
):
    docx = pytest.importorskip("docx")
    from astrbot.core.agent.run_context import ContextWrapper
    from astrbot.core.agent.tool import FunctionTool
    from astrbot.core.provider import func_tool_manager

    if not hasattr(func_tool_manager, "_PermissionGuardedTool"):
        pytest.skip("This SDK version does not wrap tools with a permission guard")

    image = tmp_path / "pixel.png"
    _write_png(image, width=1, height=1)
    image_bytes = image.read_bytes()
    document = docx.Document()
    document.add_paragraph("WORD-BODY-BEFORE-IMAGE")
    if image_mode != "absent":
        document.add_picture(str(image))
    document.add_paragraph("WORD-BODY-AFTER-IMAGE")
    path = tmp_path / "guarded-word.docx"
    document.save(path)

    workspace = WorkspaceService(
        plugin_data_path=tmp_path,
        executor=None,
        office_libs={"docx": docx},
        max_file_size=1024 * 1024,
    )
    reader = WordReadService(
        workspace_service=workspace,
        enable_docx_image_review=image_mode != "disabled",
        max_inline_docx_image_count=0 if image_mode == "limited" else 1,
    )

    async def read_word(_event):
        async for result in reader.iter_word_results(
            path, path.name, path.suffix, path.stat().st_size
        ):
            yield result

    manager = func_tool_manager.FunctionToolManager()
    name = "office_word_reader_retention_regression"
    manager.func_list.append(
        FunctionTool(
            name=name, description="Read fixture Word", parameters={}, handler=read_word
        )
    )
    guarded_tool = manager.get_full_tool_set().get_tool(name)
    assert isinstance(guarded_tool, func_tool_manager._PermissionGuardedTool)
    assert guarded_tool.handler is None
    result = await guarded_tool.call(
        ContextWrapper(
            context=SimpleNamespace(event=SimpleNamespace(is_admin=lambda: True))
        )
    )

    if image_mode == "inline":
        assert isinstance(result, mcp.types.CallToolResult)
        assert [item.type for item in result.content] == ["text", "image"]
        text = result.content[0].text
        assert base64.b64decode(result.content[1].data) == image_bytes
        assert result.content[1].mimeType == "image/png"
        assert "[插图1]" in text
    else:
        assert isinstance(result, str)
        text = result
        if image_mode == "limited":
            assert "超过单文档最多 0 张限制" in text
    assert "WORD-BODY-BEFORE-IMAGE" in text
    assert "WORD-BODY-AFTER-IMAGE" in text
    assert path.name in text
