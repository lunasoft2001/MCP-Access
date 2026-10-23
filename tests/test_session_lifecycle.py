import os
import tempfile
import unittest
from types import SimpleNamespace
from unittest.mock import patch

import win32com.client

from mcp_access import core


class _FakeProject:
    def __init__(self, app):
        self._app = app

    @property
    def FullName(self):
        return self._app.current or ""


class _FakeApp:
    def __init__(self, current=None, visible=True, visible_error=None):
        self.current = current
        self._visible = visible
        self._visible_error = visible_error
        self.visible_sets = []
        self.open_calls = []
        self.close_calls = 0
        self.quit_calls = 0
        self.CurrentProject = _FakeProject(self)
        self.Forms = SimpleNamespace(Count=0)
        self.DoCmd = SimpleNamespace(Close=lambda *_: None)
        self.hWndAccessApp = 0
        self.AutomationSecurity = 1

    @property
    def Visible(self):
        return self._visible

    @Visible.setter
    def Visible(self, value):
        self.visible_sets.append(value)
        if self._visible_error:
            raise self._visible_error
        self._visible = value

    def OpenCurrentDatabase(self, path):
        self.open_calls.append(path)
        self.current = path

    def CloseCurrentDatabase(self):
        self.close_calls += 1
        self.current = None

    def Quit(self, *_):
        self.quit_calls += 1


class SessionLifecycleTests(unittest.TestCase):
    def setUp(self):
        core._Session._force_cleanup()
        self.temp_dir = tempfile.TemporaryDirectory()
        self.target = os.path.join(self.temp_dir.name, "target.accdb")
        self.other = os.path.join(self.temp_dir.name, "other.accdb")
        open(self.target, "wb").close()
        open(self.other, "wb").close()

    def tearDown(self):
        core._Session._force_cleanup()
        self.temp_dir.cleanup()

    def test_attached_instance_is_not_modified_or_closed(self):
        app = _FakeApp(current=self.target)
        with (
            patch.object(win32com.client, "GetActiveObject", return_value=app),
            patch.object(win32com.client, "DispatchEx") as dispatch_ex,
        ):
            self.assertIs(core._Session.connect(self.target), app)
            self.assertFalse(core._Session._owns_app)
            self.assertEqual(app.visible_sets, [])
            dispatch_ex.assert_not_called()

            core._Session._cm_cache["child"] = object()
            result = core._Session.quit()

        self.assertIn("remains open", result)
        self.assertEqual(app.close_calls, 0)
        self.assertEqual(app.quit_calls, 0)
        self.assertEqual(core._Session._cm_cache, {})

    def test_owned_instance_is_visible_and_quit_once(self):
        app = _FakeApp(current=None, visible=False)
        fake_windll = SimpleNamespace(
            user32=SimpleNamespace(keybd_event=lambda *_: None),
        )
        with (
            patch.object(win32com.client, "GetActiveObject", side_effect=RuntimeError("not running")),
            patch.object(win32com.client, "DispatchEx", return_value=app),
            patch.object(core.ctypes, "windll", fake_windll),
            patch.object(core._Session, "_get_access_process_id", return_value=4242),
            patch.object(core._Session, "_wait_for_process_exit", return_value=True) as wait_for_exit,
        ):
            self.assertIs(core._Session.connect(self.target), app)
            self.assertTrue(core._Session._owns_app)
            self.assertEqual(app.visible_sets, [True])
            self.assertEqual(app.open_calls, [os.path.abspath(self.target)])

            result = core._Session.quit()

        self.assertIn("exited successfully", result)
        self.assertEqual(app.close_calls, 0)
        self.assertEqual(app.quit_calls, 1)
        wait_for_exit.assert_called_once_with(
            4242,
            core._OWNED_SHUTDOWN_TIMEOUT_SECONDS,
        )

    def test_owned_shutdown_timeout_is_reported(self):
        app = _FakeApp(current=self.target)
        core._Session._app = app
        core._Session._db_open = self.target
        core._Session._owns_app = True

        with (
            patch.object(core._Session, "_get_access_process_id", return_value=4242),
            patch.object(core._Session, "_wait_for_process_exit", return_value=False),
        ):
            with self.assertRaisesRegex(RuntimeError, "did not exit within 10 seconds"):
                core._Session.quit()

        self.assertEqual(app.quit_calls, 1)
        self.assertIsNone(core._Session._app)
        self.assertFalse(core._Session._owns_app)

    def test_attached_instance_with_different_database_is_rejected(self):
        app = _FakeApp(current=self.other)
        with patch.object(win32com.client, "GetActiveObject", return_value=app):
            with self.assertRaisesRegex(RuntimeError, "will not be closed"):
                core._Session.connect(self.target)

        self.assertEqual(app.close_calls, 0)
        self.assertEqual(app.open_calls, [])
        self.assertEqual(app.current, self.other)

    def test_attached_instance_without_database_retains_opened_target(self):
        app = _FakeApp(current=None)
        fake_windll = SimpleNamespace(
            user32=SimpleNamespace(keybd_event=lambda *_: None),
        )
        with (
            patch.object(win32com.client, "GetActiveObject", return_value=app),
            patch.object(core.ctypes, "windll", fake_windll),
        ):
            self.assertIs(core._Session.connect(self.target), app)
            self.assertFalse(core._Session._owns_app)
            self.assertEqual(app.open_calls, [os.path.abspath(self.target)])

            result = core._Session.quit()

        self.assertIn("remains open", result)
        self.assertEqual(app.current, os.path.abspath(self.target))
        self.assertEqual(app.close_calls, 0)
        self.assertEqual(app.quit_calls, 0)

    def test_auto_decompile_does_not_close_uncached_user_instance(self):
        app = _FakeApp(current=self.target)
        with (
            patch.object(win32com.client, "GetActiveObject", return_value=app),
            patch.object(core.os.path, "exists", return_value=True),
        ):
            core._Session._decompile(self.target)

        self.assertFalse(core._Session._owns_app)
        self.assertEqual(app.close_calls, 0)
        self.assertEqual(app.quit_calls, 0)
        self.assertIn(self.target, core._Session._decompiled_dbs)

    def test_failed_owned_launch_does_not_leave_cached_session(self):
        app = _FakeApp(
            current=None,
            visible=False,
            visible_error=RuntimeError("invalid Visible reference"),
        )
        with patch.object(win32com.client, "DispatchEx", return_value=app):
            with self.assertRaisesRegex(RuntimeError, "Could not start"):
                core._Session._launch_owned()

        self.assertIsNone(core._Session._app)
        self.assertIsNone(core._Session._db_open)
        self.assertFalse(core._Session._owns_app)
        self.assertEqual(app.quit_calls, 1)


if __name__ == "__main__":
    unittest.main()
