from types import SimpleNamespace
from unittest.mock import MagicMock

import pytest

from astrbot_plugin_office_assistant import utils


@pytest.mark.parametrize(
    "app_name", ["Word.Application", "Excel.Application", "PowerPoint.Application"]
)
@pytest.mark.parametrize("operation_fails", [False, True])
def test_com_application_owns_its_instance_and_respects_powerpoint_visibility(
    monkeypatch, app_name, operation_fails
):
    class OfficeApplication:
        DisplayAlerts = True

        def __init__(self):
            self.visibility_assignments = []
            self.Quit = MagicMock()

        @property
        def Visible(self):
            return False

        @Visible.setter
        def Visible(self, value):
            if app_name == "PowerPoint.Application":
                raise RuntimeError("Hiding the application window is not allowed")
            self.visibility_assignments.append(value)

    existing_user_instance = OfficeApplication()
    private_instance = OfficeApplication()
    client = SimpleNamespace(
        Dispatch=MagicMock(return_value=existing_user_instance),
        DispatchEx=MagicMock(return_value=private_instance),
    )
    pythoncom = SimpleNamespace(CoInitialize=MagicMock(), CoUninitialize=MagicMock())
    monkeypatch.setattr(utils, "_WIN32COM_AVAILABLE", True)
    monkeypatch.setattr(
        utils, "win32com", SimpleNamespace(client=client), raising=False
    )
    monkeypatch.setattr(utils, "pythoncom", pythoncom, raising=False)

    def use_application():
        with utils.com_application(app_name) as application:
            assert application is private_instance
            if operation_fails:
                raise ValueError("conversion failed")

    if operation_fails:
        with pytest.raises(ValueError, match="conversion failed"):
            use_application()
    else:
        use_application()

    client.Dispatch.assert_not_called()
    client.DispatchEx.assert_called_once_with(app_name)
    private_instance.Quit.assert_called_once()
    existing_user_instance.Quit.assert_not_called()
    assert private_instance.visibility_assignments == (
        [] if app_name == "PowerPoint.Application" else [False]
    )
    pythoncom.CoInitialize.assert_called_once()
    pythoncom.CoUninitialize.assert_called_once()
