"""
Core: COM session singleton, caches, logging, COM thread pool.
All other modules import shared state from here.
"""

import asyncio
import concurrent.futures
import ctypes
import logging
import os
import subprocess
import time
import sys
import threading
import winreg
from pathlib import Path
from typing import Any, Optional

# DPI awareness -- must be set before any window operations
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PROCESS_PER_MONITOR_DPI_AWARE
except Exception:
    pass

# ---------------------------------------------------------------------------
# Logging -- goes to stderr to avoid polluting the JSON-RPC stdout channel
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s — %(message)s",
    stream=sys.stderr,
)
log = logging.getLogger("access-mcp")

# ---------------------------------------------------------------------------
# COM thread pool -- single thread so all COM calls stay in the same STA.
# ---------------------------------------------------------------------------

def _com_thread_init():
    """Initializer for the COM worker thread -- calls CoInitialize once."""
    import pythoncom
    pythoncom.CoInitialize()
    log.info("COM thread initialized (thread=%s)", threading.current_thread().name)

_com_executor = concurrent.futures.ThreadPoolExecutor(
    max_workers=1,
    thread_name_prefix="com-worker",
    initializer=_com_thread_init,
)

# ---------------------------------------------------------------------------
# Access COM constants
# ---------------------------------------------------------------------------
AC_TYPE: dict[str, int] = {
    "query":  1,   # acQuery
    "form":   2,   # acForm
    "report": 3,   # acReport
    "macro":  4,   # acMacro
    "module": 5,   # acModule
}

# ---------------------------------------------------------------------------
# Caches to reduce COM calls in long sessions
# ---------------------------------------------------------------------------
_vbe_code_cache: dict = {}        # "type:name" -> full text of VBE module
_parsed_controls_cache: dict = {} # "form:name" / "report:name" -> _parse_controls() result
_OWNED_SHUTDOWN_TIMEOUT_SECONDS = 10.0

# ---------------------------------------------------------------------------
# COM Session -- singleton, keeps Access alive between calls
# ---------------------------------------------------------------------------
class _Session:
    """
    Maintains a single Access.Application instance across tool calls.
    An attached user instance is never closed or switched to another database.
    If that instance has no database open, MCP may open the requested database
    in it and leaves that database open when the session disconnects.
    An instance created by MCP can be switched and closed.
    """
    _app: Optional[Any] = None
    _db_open: Optional[str] = None
    _owns_app: bool = False
    _app_pid: Optional[int] = None
    _cm_cache: dict = {}   # "type:name" -> CodeModule COM object
    _decompiled_dbs: set = set()  # DBs already decompiled in this session

    @classmethod
    def connect(cls, db_path: str) -> Any:
        resolved = str(Path(db_path).resolve())
        if cls._app is not None:
            # Health check: verify COM session is still alive
            try:
                _ = cls._app.Visible  # cheap COM property access
            except Exception:
                log.warning("COM session stale — auto-reconnecting...")
                cls._force_cleanup()
        if cls._app is None:
            cls._launch()
        current = cls._current_database_path()
        cls._db_open = current
        if not cls._paths_equal(current, resolved):
            cls._switch(resolved)
        return cls._app

    @classmethod
    def _force_cleanup(cls):
        """Reset state without calling methods on a dead COM object."""
        cls._clear_object_caches()
        cls._app = None
        cls._db_open = None
        cls._owns_app = False
        cls._app_pid = None
        cls._decompiled_dbs.clear()

    @classmethod
    def _clear_object_caches(cls) -> None:
        cls._cm_cache.clear()
        _vbe_code_cache.clear()
        _parsed_controls_cache.clear()

    @staticmethod
    def _paths_equal(left: Optional[str], right: Optional[str]) -> bool:
        if not left or not right:
            return left == right
        return os.path.normcase(os.path.abspath(left)) == os.path.normcase(os.path.abspath(right))

    @classmethod
    def _current_database_path(cls) -> Optional[str]:
        if cls._app is None:
            return None
        try:
            current = str(cls._app.CurrentProject.FullName or "").strip()
            return str(Path(current).resolve()) if current else None
        except Exception:
            return None

    @classmethod
    def _launch(cls) -> None:
        try:
            import win32com.client
        except ImportError:
            raise RuntimeError(
                "pywin32 not installed. Run: pip install pywin32"
            )
        # Prefer attaching to an already-running Access instance so we don't
        # spawn a second process when the user already has Access open (e.g.
        # for interactive debugging). Fall back to DispatchEx only when none
        # is running — DispatchEx remains required after /decompile kills to
        # bypass stale ROT entries, but in that path no live instance exists.
        if not cls._try_attach_active():
            cls._launch_owned()

    @classmethod
    def _try_attach_active(cls) -> bool:
        try:
            import win32com.client

            cls._app = win32com.client.GetActiveObject("Access.Application")
        except Exception:
            cls._app = None
            cls._db_open = None
            cls._owns_app = False
            cls._app_pid = None
            return False
        cls._owns_app = False
        cls._app_pid = cls._get_access_process_id(cls._app)
        log.info("Attached to existing Access.Application instance")
        cls._db_open = cls._current_database_path()
        if cls._db_open:
            log.info("Existing Access has DB open: %s", cls._db_open)
        return True

    @classmethod
    def _launch_owned(cls) -> None:
        try:
            import win32com.client
        except ImportError:
            raise RuntimeError(
                "pywin32 not installed. Run: pip install pywin32"
            )
        try:
            log.info("Launching new Access.Application...")
            cls._app = win32com.client.DispatchEx("Access.Application")
            cls._owns_app = True
            cls._app.Visible = True   # required for VBE to be accessible via COM
            cls._app_pid = cls._get_access_process_id(cls._app)
            cls._db_open = cls._current_database_path()
            log.info("Access launched OK (PID=%s)", cls._app_pid or "unknown")
        except Exception as exc:
            app = cls._app
            if app is not None and cls._owns_app:
                try:
                    app.Quit()
                except Exception:
                    pass
            cls._force_cleanup()
            raise RuntimeError(f"Could not start Microsoft Access: {exc}") from exc

    @classmethod
    def reopen(cls, path: str) -> None:
        """Forces reopen with SHIFT (bypass AutoExec) via _switch().
        Use after CloseCurrentDatabase+CompactRepair in maintenance."""
        cls._db_open = None
        cls._switch(path)

    @classmethod
    def _decompile(cls, path: str) -> None:
        """Run MSACCESS /decompile + SHIFT on the DB before opening via COM.
        Strips orphaned p-code so compile errors are real, not phantom."""
        msaccess_candidates = [
            r"C:\Program Files\Microsoft Office\root\Office16\MSACCESS.EXE",
            r"C:\Program Files (x86)\Microsoft Office\root\Office16\MSACCESS.EXE",
        ]
        msaccess = next((p for p in msaccess_candidates if os.path.exists(p)), None)
        if not msaccess:
            log.warning("MSACCESS.EXE not found — skipping /decompile")
            cls._decompiled_dbs.add(path)  # don't retry
            return

        if cls._app is None:
            cls._try_attach_active()
        if cls._app is not None and not cls._owns_app:
            log.warning(
                "Skipping automatic /decompile because Access is user-owned: %s",
                path,
            )
            cls._decompiled_dbs.add(path)
            return

        # Close the owned COM session completely so the file is unlocked
        if cls._app is not None:
            log.info("Closing COM session for /decompile...")
            cls._clear_object_caches()
            try:
                cls._app.Quit(1)  # acQuitSaveNone
            except Exception:
                pass
            cls._app = None
            cls._db_open = None
            cls._owns_app = False
            cls._app_pid = None

        log.info("Decompiling %s ...", path)

        # Hold SHIFT while launching /decompile
        VK_SHIFT = 0x10
        KEYEVENTF_KEYUP = 0x0002
        _kbd = ctypes.windll.user32.keybd_event
        shift_held = False
        try:
            _kbd(VK_SHIFT, 0, 0, 0)
            time.sleep(0.3)
            shift_held = True
        except Exception:
            pass

        proc = subprocess.Popen(
            [msaccess, path, "/decompile"],
            creationflags=subprocess.CREATE_NEW_PROCESS_GROUP,
        )

        # Polling loop: 16 × 0.5s = 8s total.  Release SHIFT at ~3s mark
        # and poll for any blocking dialogs (wizards, recovery prompts)
        # via _dismiss_dialogs_by_pid.  Breaks early if the subprocess
        # exits on its own.
        for i in range(16):
            if i == 6 and shift_held:  # ~3s mark
                try:
                    _kbd(VK_SHIFT, 0, KEYEVENTF_KEYUP, 0)
                except Exception:
                    pass
                shift_held = False
            if proc.poll() is not None:
                break  # subprocess already exited
            try:
                from .vba_exec import _dismiss_dialogs_by_pid  # lazy — circular
                _dismiss_dialogs_by_pid(proc.pid)
            except Exception:
                pass
            time.sleep(0.5)

        # Ensure SHIFT released even on early exit
        if shift_held:
            try:
                _kbd(VK_SHIFT, 0, KEYEVENTF_KEYUP, 0)
            except Exception:
                pass
        try:
            subprocess.run(
                ["taskkill", "/F", "/T", "/PID", str(proc.pid)],
                capture_output=True, timeout=10,
            )
        except Exception:
            pass
        time.sleep(1)  # let Windows evict the dead process's ROT entry

        cls._decompiled_dbs.add(path)
        log.info("Decompile done for %s", path)

        # Re-launch a dedicated instance; do not attach to another user instance.
        cls._launch_owned()

    @staticmethod
    def _suppress_recovery_dialog() -> None:
        """Disable the 'last time you opened this file it caused a serious error' dialog."""
        key_path = r"Software\Microsoft\Office\16.0\Access\Resiliency"
        try:
            key = winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE)
            try:
                winreg.SetValueEx(key, "DisableAllCallersWarning", 0, winreg.REG_DWORD, 1)
                winreg.SetValueEx(key, "DoNotShowUI", 0, winreg.REG_DWORD, 1)
            finally:
                winreg.CloseKey(key)
            log.info("Recovery dialog suppressed via registry")
        except Exception as e:
            log.warning("Could not suppress recovery dialog: %s", e)

    @classmethod
    def _switch(cls, path: str) -> None:
        if not os.path.isfile(path):
            raise FileNotFoundError(f"File not found: {path}")

        current = cls._current_database_path()
        if cls._paths_equal(current, path):
            cls._db_open = current
            return

        if current is not None:
            if not cls._owns_app:
                raise RuntimeError(
                    "The attached Access instance already has a different database open "
                    f"({current}). It will not be closed to open {path}."
                )
            log.info("Closing previous DB: %s", current)
            cls._clear_object_caches()
            try:
                cls._app.CloseCurrentDatabase()
            except Exception as e:
                raise RuntimeError(f"Could not close the current database ({current}): {e}") from e
            still_open = cls._current_database_path()
            if still_open is not None:
                raise RuntimeError(f"Access did not close the current database: {still_open}")
        log.info("Opening DB: %s", path)

        cls._suppress_recovery_dialog()

        # Hold Shift during OpenCurrentDatabase to bypass AutoExec/startup forms
        VK_SHIFT = 0x10
        KEYEVENTF_KEYUP = 0x0002
        _kbd = ctypes.windll.user32.keybd_event
        shift_held = False
        try:
            _kbd(VK_SHIFT, 0, 0, 0)  # Press SHIFT
            time.sleep(0.3)  # Let the key state register before COM call
            shift_held = True
            log.info("SHIFT held for bypass")
        except Exception:
            log.warning("Could not simulate Shift — AutoExec may run")

        # Capture Access hwnd on THIS thread (COM worker — same apartment
        # that created _app).  COM STA proxies cannot be accessed from the
        # watchdog thread, so we must resolve hwnd here before spawning it.
        try:
            _h = cls._app.hWndAccessApp
            access_hwnd = int(_h() if callable(_h) else _h)
        except Exception as e_hwnd:
            log.warning("Could not capture Access hwnd for watchdog: %s", e_hwnd)
            access_hwnd = 0

        # Polling watchdog: after a 2s grace period, poll every 0.5s and
        # dismiss any dialog via proper button click (Cancel first — NEVER
        # VK_RETURN, which would advance wizards and create stray objects).
        _open_done = threading.Event()
        _dialog_screenshots: list = []

        def _watchdog():
            # Lazy import to avoid circular dependency with vba_exec
            from .vba_exec import _dismiss_access_dialogs
            # Grace period — normal opens complete in <2s
            if _open_done.wait(2):
                return
            log.warning("OpenCurrentDatabase blocked >2s — polling for dialogs")
            while not _open_done.is_set():
                if access_hwnd:
                    try:
                        if _dismiss_access_dialogs(
                            access_hwnd,
                            _dialog_screenshots if len(_dialog_screenshots) == 0 else None,
                        ):
                            log.warning("Dialog dismissed during OpenCurrentDatabase")
                    except Exception as e_wd:
                        log.warning("Watchdog dismiss error: %s", e_wd)
                _open_done.wait(0.5)

        watchdog_thread = threading.Thread(target=_watchdog, daemon=True)
        watchdog_thread.start()

        # Call OpenCurrentDatabase in the current thread (COM worker — same
        # apartment that created _app).  The watchdog will dismiss any dialog.
        # AutomationSecurity=3 as defence-in-depth.  Does NOT suppress Access
        # AutoExec macro objects (tested — Access ignores it for those), but
        # may prevent VBA auto-run code in edge cases where the Shift key
        # doesn't register (remote desktop, key event eaten by another app).
        try:
            cls._app.AutomationSecurity = 3   # msoAutomationSecurityForceDisable
        except Exception:
            pass
        try:
            cls._app.OpenCurrentDatabase(path)
        except Exception as e:
            actual = cls._current_database_path()
            if "already have the database open" not in str(e).lower() or not cls._paths_equal(actual, path):
                raise
        finally:
            _open_done.set()  # signal watchdog to stop
            try:
                cls._app.AutomationSecurity = 1  # msoAutomationSecurityLow — restore
            except Exception:
                pass
            if shift_held:
                try:
                    _kbd(VK_SHIFT, 0, KEYEVENTF_KEYUP, 0)  # Release SHIFT
                    log.info("SHIFT released")
                except Exception:
                    pass

        if _dialog_screenshots:
            log.warning("A blocking dialog was auto-dismissed. Screenshot: %s",
                        _dialog_screenshots[0])

        actual = cls._current_database_path()
        if not cls._paths_equal(actual, path):
            raise RuntimeError(
                f"Access did not open the requested database. Requested: {path}; "
                f"current: {actual or '(none)'}"
            )
        cls._db_open = actual

        # Close any auto-opened forms (safety net)
        try:
            for i in range(cls._app.Forms.Count - 1, -1, -1):
                try:
                    name = cls._app.Forms(i).Name
                    cls._app.DoCmd.Close(2, name)  # acForm
                    log.info("Closed auto-opened form: %s", name)
                except Exception:
                    pass
        except Exception:
            pass

        # Clear caches on DB switch
        cls._clear_object_caches()
        log.info("DB opened OK")

    @staticmethod
    def _get_access_process_id(app: Any) -> Optional[int]:
        try:
            import win32process

            hwnd_value = app.hWndAccessApp
            hwnd = int(hwnd_value() if callable(hwnd_value) else hwnd_value)
            if not hwnd:
                return None
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            return int(pid) if pid else None
        except Exception as exc:
            log.warning("Could not determine Access process ID: %s", exc)
            return None

    @staticmethod
    def _wait_for_process_exit(pid: int, timeout_seconds: float) -> Optional[bool]:
        """Wait for a process while continuing to pump the COM STA message queue."""
        from ctypes import wintypes

        import pythoncom

        synchronize = 0x00100000
        wait_object_0 = 0
        wait_timeout = 258
        wait_failed = 0xFFFFFFFF
        error_invalid_parameter = 87

        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
        kernel32.OpenProcess.restype = wintypes.HANDLE
        kernel32.WaitForSingleObject.argtypes = [wintypes.HANDLE, wintypes.DWORD]
        kernel32.WaitForSingleObject.restype = wintypes.DWORD
        kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
        kernel32.CloseHandle.restype = wintypes.BOOL

        process_handle = kernel32.OpenProcess(synchronize, False, pid)
        if not process_handle:
            error = ctypes.get_last_error()
            if error == error_invalid_parameter:
                return True
            log.warning("Could not open Access PID %s for shutdown verification (error=%s)", pid, error)
            return None

        deadline = time.monotonic() + timeout_seconds
        try:
            while True:
                remaining = deadline - time.monotonic()
                if remaining <= 0:
                    return False
                wait_ms = min(100, max(1, int(remaining * 1000)))
                result = kernel32.WaitForSingleObject(process_handle, wait_ms)
                if result == wait_object_0:
                    return True
                if result == wait_failed:
                    log.warning(
                        "Could not wait for Access PID %s (error=%s)",
                        pid,
                        ctypes.get_last_error(),
                    )
                    return None
                if result != wait_timeout:
                    log.warning("Unexpected wait result for Access PID %s: %s", pid, result)
                    return None
                pythoncom.PumpWaitingMessages()
        finally:
            kernel32.CloseHandle(process_handle)

    @classmethod
    def quit(cls, suppress_errors: bool = False) -> str:
        if cls._app is None:
            cls._force_cleanup()
            return "No active Access session."

        app = cls._app
        owns_app = cls._owns_app
        pid = (cls._app_pid or cls._get_access_process_id(app)) if owns_app else None
        cls._clear_object_caches()

        if not owns_app:
            log.info("Releasing attached COM session without closing Access...")
            cls._app = None
            cls._db_open = None
            cls._owns_app = False
            cls._app_pid = None
            cls._decompiled_dbs.clear()
            app = None
            return "COM session released; the attached Access instance remains open."

        quit_error = None
        try:
            log.info("Requesting shutdown of the MCP-owned Access instance...")
            app.Quit()
        except Exception as exc:
            quit_error = exc
        finally:
            cls._app = None
            cls._db_open = None
            cls._owns_app = False
            cls._app_pid = None
            cls._decompiled_dbs.clear()
            app = None

        if quit_error is not None:
            log.warning("Error closing Access: %s", quit_error)
            message = f"Could not close the MCP-owned Access instance: {quit_error}"
            if not suppress_errors:
                raise RuntimeError(message) from quit_error
            return message

        if pid is None:
            message = (
                "Shutdown requested for the Access instance created by MCP; "
                "process exit could not be verified because its PID was unavailable."
            )
            log.warning(message)
            return message

        exited = cls._wait_for_process_exit(pid, _OWNED_SHUTDOWN_TIMEOUT_SECONDS)
        if exited is True:
            return f"The Access instance created by MCP exited successfully (PID {pid})."
        if exited is None:
            message = (
                f"Shutdown requested for the Access instance created by MCP (PID {pid}), "
                "but process exit could not be verified."
            )
            log.warning(message)
            return message

        message = (
            f"Shutdown timed out: Access PID {pid} did not exit within "
            f"{_OWNED_SHUTDOWN_TIMEOUT_SECONDS:g} seconds."
        )
        log.warning(message)
        if not suppress_errors:
            raise RuntimeError(message)
        return message


def shutdown_com_thread() -> str:
    """Release the Access session and COM apartment on the COM worker thread."""
    try:
        return _Session.quit(suppress_errors=True)
    finally:
        try:
            import pythoncom
            pythoncom.CoUninitialize()
        except Exception as exc:
            log.warning("Could not uninitialize COM worker thread: %s", exc)


def _get_vb_project(app):
    """Return the VBProject that belongs to the current database.

    ``app.VBE.VBProjects(1)`` may return the wrong project (e.g. the
    ``acwzmain`` wizard library) after a decompile+compact cycle.  This
    helper enumerates all loaded VBProjects and picks the one whose
    ``.FileName`` matches ``_Session._db_open``.  Falls back to index 1
    if no match is found (single-project scenario).
    """
    db_path = _Session._db_open
    try:
        projects = app.VBE.VBProjects
        count = projects.Count
        if db_path:
            db_norm = os.path.normcase(os.path.abspath(db_path))
            for i in range(1, count + 1):
                try:
                    proj = projects(i)
                    fname = getattr(proj, "FileName", "") or ""
                    if fname and os.path.normcase(os.path.abspath(fname)) == db_norm:
                        return proj
                except Exception:
                    continue
        # Fallback: first project
        return projects(1)
    except Exception:
        return app.VBE.VBProjects(1)


def invalidate_all_caches():
    """Convenience: clear all 3 caches at once."""
    _vbe_code_cache.clear()
    _parsed_controls_cache.clear()
    _Session._cm_cache.clear()


def invalidate_object_caches(object_type: str, object_name: str):
    """Clear caches for a specific object."""
    cache_key = f"{object_type}:{object_name}"
    _vbe_code_cache.pop(cache_key, None)
    _parsed_controls_cache.pop(cache_key, None)
    _Session._cm_cache.pop(cache_key, None)
