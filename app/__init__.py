__all__ = [
    "AdminUsersResolver",
    "FileProcessingServices",
    "PluginRuntimeBundle",
    "PluginSettings",
    "RequestPipelineServices",
    "load_plugin_settings",
]


# Keep __all__ and __getattr__ in sync when exports change.
# This package uses lazy imports here to avoid importing runtime wiring
# during module import unless the caller actually requests an export.
def __getattr__(name: str):
    if name == "load_plugin_settings":
        from .settings import load_plugin_settings

        return load_plugin_settings
    if name == "PluginSettings":
        from .settings import PluginSettings

        return PluginSettings
    if name in {
        "AdminUsersResolver",
        "FileProcessingServices",
        "PluginRuntimeBundle",
        "RequestPipelineServices",
    }:
        from .runtime import __dict__ as runtime_namespace

        return runtime_namespace[name]
    raise AttributeError(name)
