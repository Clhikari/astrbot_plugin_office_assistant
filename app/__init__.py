__all__ = [
    "AdminUsersResolver",
    "FileProcessingServices",
    "PluginRuntimeBundle",
    "PluginSettings",
    "RequestPipelineServices",
    "load_plugin_settings",
]


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
