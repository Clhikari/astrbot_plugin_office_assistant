__all__ = ["create_server", "DocumentSessionStore"]


def __getattr__(name: str):
    if name == "create_server":
        from .server import create_server

        return create_server
    if name == "DocumentSessionStore":
        from ..domain.document.session_store import DocumentSessionStore

        return DocumentSessionStore
    raise AttributeError(name)
