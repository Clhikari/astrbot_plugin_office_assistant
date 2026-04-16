from __future__ import annotations

from types import SimpleNamespace

from astrbot_plugin_office_assistant.agent_tools import build_workbook_toolset
from astrbot_plugin_office_assistant.tools.astrbot_adapter import (
    build_workbook_toolset_from_registry,
)
from astrbot_plugin_office_assistant.tools.mcp_adapter import (
    register_workbook_tools_from_registry,
)
from astrbot_plugin_office_assistant.tools.registry import (
    WorkbookToolSpec,
    get_document_tool_specs,
    get_workbook_tool_specs,
)


def test_workbook_tool_registry_keeps_workbook_tool_order():
    assert [spec.name for spec in get_workbook_tool_specs()] == [
        "create_workbook",
        "write_rows",
        "export_workbook",
    ]


def test_document_tool_registry_keeps_document_tool_order_after_workbook_extension():
    assert [spec.name for spec in get_document_tool_specs()] == [
        "create_document",
        "add_blocks",
        "finalize_document",
        "export_document",
    ]


def test_astrbot_workbook_toolset_preserves_registry_order(
    monkeypatch,
):
    fake_store = object()
    captured_after_export = {}

    def _factory(name: str):
        def _build(store, after_export):
            captured_after_export[name] = after_export
            return SimpleNamespace(name=name, store=store)

        return _build

    specs = (
        WorkbookToolSpec(
            name="create_workbook",
            astrbot_factory=_factory("create_workbook"),
            mcp_registrar=lambda *_args, **_kwargs: None,
        ),
        WorkbookToolSpec(
            name="write_rows",
            astrbot_factory=_factory("write_rows"),
            mcp_registrar=lambda *_args, **_kwargs: None,
        ),
        WorkbookToolSpec(
            name="export_workbook",
            astrbot_factory=_factory("export_workbook"),
            mcp_registrar=lambda *_args, **_kwargs: None,
        ),
    )

    monkeypatch.setattr(
        "astrbot_plugin_office_assistant.tools.astrbot_adapter.build_workbook_store",
        lambda workspace_dir=None: fake_store,
    )
    monkeypatch.setattr(
        "astrbot_plugin_office_assistant.tools.astrbot_adapter.get_workbook_tool_specs",
        lambda: specs,
    )
    monkeypatch.setattr(
        "astrbot_plugin_office_assistant.tools.astrbot_adapter.WorkbookToolSet",
        lambda tools, *, workbook_store: SimpleNamespace(
            tools=list(tools), workbook_store=workbook_store
        ),
    )

    async def _after_export(_context, _path):
        return None

    toolset = build_workbook_toolset_from_registry(after_export=_after_export)

    assert [tool.name for tool in toolset.tools] == [spec.name for spec in specs]
    assert toolset.workbook_store is fake_store
    assert captured_after_export["export_workbook"] is _after_export


def test_agent_factory_build_workbook_toolset_delegates_to_adapter(
    monkeypatch,
):
    fake_toolset = SimpleNamespace(tools=[])
    monkeypatch.setattr(
        "astrbot_plugin_office_assistant.agent_tools.factory.build_workbook_toolset_from_registry",
        lambda workspace_dir=None, after_export=None: fake_toolset,
    )
    result = build_workbook_toolset()
    assert result is fake_toolset


def test_mcp_workbook_tool_registration_matches_registry_order(
    monkeypatch,
):
    registered_names: list[str] = []

    def _make_registrar(name: str):
        def _record(*_args, **_kwargs):
            registered_names.append(name)

        return _record

    specs = (
        WorkbookToolSpec(
            name="create_workbook",
            astrbot_factory=lambda *_args, **_kwargs: None,
            mcp_registrar=_make_registrar("create_workbook"),
        ),
        WorkbookToolSpec(
            name="write_rows",
            astrbot_factory=lambda *_args, **_kwargs: None,
            mcp_registrar=_make_registrar("write_rows"),
        ),
        WorkbookToolSpec(
            name="export_workbook",
            astrbot_factory=lambda *_args, **_kwargs: None,
            mcp_registrar=_make_registrar("export_workbook"),
        ),
    )
    monkeypatch.setattr(
        "astrbot_plugin_office_assistant.tools.mcp_adapter.get_workbook_tool_specs",
        lambda: specs,
    )

    register_workbook_tools_from_registry(
        server=SimpleNamespace(),
        store=SimpleNamespace(),
    )

    assert registered_names == [spec.name for spec in specs]
