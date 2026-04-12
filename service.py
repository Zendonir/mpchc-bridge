"""
MPC-HC Bridge — autostart manager (no admin rights required).

Usage:
  mpchc-bridge.exe install   — register autostart for current user (HKCU Run key)
  mpchc-bridge.exe start     — launch bridge now (without rebooting)
  mpchc-bridge.exe stop      — kill running bridge process
  mpchc-bridge.exe remove    — remove autostart entry
  mpchc-bridge.exe status    — show autostart registration status
  mpchc-bridge.exe debug     — run in foreground (console mode, Ctrl+C to stop)
"""

import asyncio
import os
import subprocess
import sys

if sys.platform == "win32":
    import winreg


# ── Service implementation ─────────────────────────────────────────────────────

def _run_service_loop(stop_event: asyncio.Event) -> None:
    """Run the aiohttp bridge until stop_event is set."""
    from aiohttp import web
    from bridge import BRIDGE_PORT, create_app

    async def _serve() -> None:
        app = create_app()
        runner = web.AppRunner(app)
        await runner.setup()
        site = web.TCPSite(runner, "0.0.0.0", BRIDGE_PORT)
        await site.start()
        await stop_event.wait()
        await runner.cleanup()

    asyncio.run(_serve())


try:
    import win32event
    import win32service
    import win32serviceutil
    import servicemanager

    class MpcHcBridgeService(win32serviceutil.ServiceFramework):
        """Windows SCM service class."""

        _svc_name_ = "MpcHcBridge"
        _svc_display_name_ = "MPC-HC Control Bridge"
        _svc_description_ = (
            "HTTP API bridge for UC Remote 3 — full MPC-HC control "
            "(playback, subtitle/audio track selection, seek, volume)."
        )

        def __init__(self, args: list) -> None:
            win32serviceutil.ServiceFramework.__init__(self, args)
            self._win32_stop = win32event.CreateEvent(None, 0, 0, None)
            self._loop: asyncio.AbstractEventLoop | None = None
            self._stop_async: asyncio.Event | None = None

        def SvcStop(self) -> None:
            self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
            win32event.SetEvent(self._win32_stop)
            if self._loop and self._stop_async:
                self._loop.call_soon_threadsafe(self._stop_async.set)

        def SvcDoRun(self) -> None:
            servicemanager.LogMsg(
                servicemanager.EVENTLOG_INFORMATION_TYPE,
                servicemanager.PYS_SERVICE_STARTED,
                (self._svc_name_, ""),
            )
            self._loop = asyncio.new_event_loop()
            asyncio.set_event_loop(self._loop)
            self._stop_async = asyncio.Event()

            from aiohttp import web
            from bridge import BRIDGE_PORT, create_app

            async def _serve() -> None:
                app = create_app()
                runner = web.AppRunner(app)
                await runner.setup()
                site = web.TCPSite(runner, "0.0.0.0", BRIDGE_PORT)
                await site.start()
                servicemanager.LogMsg(
                    servicemanager.EVENTLOG_INFORMATION_TYPE,
                    servicemanager.PYS_SERVICE_STARTED,
                    (self._svc_name_, f" listening on port {BRIDGE_PORT}"),
                )
                await self._stop_async.wait()  # type: ignore[union-attr]
                await runner.cleanup()

            try:
                self._loop.run_until_complete(_serve())
            finally:
                self._loop.close()

    _PYWIN32_AVAILABLE = True

except ImportError:
    _PYWIN32_AVAILABLE = False
    MpcHcBridgeService = None  # type: ignore[assignment,misc]


# ── CLI entry point ────────────────────────────────────────────────────────────

def _print_usage() -> None:
    print(__doc__)


_REG_RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
_REG_VALUE = "MpcHcBridge"


def _autostart_cmd() -> str:
    exe = os.path.abspath(sys.executable if getattr(sys, "frozen", False) else sys.argv[0])
    if not getattr(sys, "frozen", False):
        return f'"{exe}" "{os.path.abspath(__file__)}" debug'
    return f'"{exe}" debug'


def _install_interactive() -> None:
    """Register autostart via HKCU Run key — zero admin rights needed."""
    cmd = _autostart_cmd()
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _REG_RUN_KEY, 0, winreg.KEY_SET_VALUE) as k:
            winreg.SetValueEx(k, _REG_VALUE, 0, winreg.REG_SZ, cmd)
        print(f"Installed: '{_REG_VALUE}' autostart registered for current user.")
        print(f"  Command: {cmd}")
        print("Run 'start' to launch it now without rebooting.")
    except Exception as ex:  # pylint: disable=broad-exception-caught
        print(f"ERROR: {ex}")
        sys.exit(1)


def _remove_task() -> None:
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _REG_RUN_KEY, 0, winreg.KEY_SET_VALUE) as k:
            winreg.DeleteValue(k, _REG_VALUE)
        print(f"Removed autostart entry '{_REG_VALUE}'.")
    except FileNotFoundError:
        print("Already removed.")
    except Exception as ex:  # pylint: disable=broad-exception-caught
        print(f"ERROR: {ex}")


def _start_task() -> None:
    cmd = _autostart_cmd()
    parts = cmd.split('" "')
    # Split quoted command into executable and args
    exe = parts[0].lstrip('"')
    args = [a.rstrip('"') for a in parts[1:]] if len(parts) > 1 else []
    try:
        subprocess.Popen(
            [exe] + args,
            creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP,
            close_fds=True,
        )
        print("Bridge started.")
    except Exception as ex:  # pylint: disable=broad-exception-caught
        print(f"ERROR: {ex}")


def _status_task() -> None:
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _REG_RUN_KEY) as k:
            val, _ = winreg.QueryValueEx(k, _REG_VALUE)
        print(f"Autostart: registered\n  Command: {val}")
    except FileNotFoundError:
        print("Autostart: not installed")
    except Exception as ex:  # pylint: disable=broad-exception-caught
        print(f"ERROR: {ex}")


def main() -> None:
    """Dispatch based on arguments, or launch GUI if none given."""
    args = sys.argv[1:]

    # --service → called by Windows SCM, start the service dispatcher
    if args and args[0] == "--service":
        if not _PYWIN32_AVAILABLE:
            sys.exit(1)
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(MpcHcBridgeService)
        servicemanager.StartServiceCtrlDispatcher()
        return

    # No arguments → open GUI
    if not args:
        from gui import main as gui_main
        gui_main()
        return

    if args[0] in ("-h", "--help"):
        _print_usage()
        return

    if args[0] == "debug":
        from bridge import BRIDGE_PORT, main as bridge_main
        print(f"Running in debug mode on port {BRIDGE_PORT} — Ctrl+C to stop")
        bridge_main()
        return

    if args[0] == "install":
        _install_interactive()
        return

    if args[0] == "remove":
        _remove_task()
        return

    if args[0] == "start":
        _start_task()
        return

    if args[0] == "status":
        _status_task()
        return

    if args[0] == "stop":
        # Kill the running bridge process by port
        subprocess.run(
            ["taskkill", "/f", "/im", os.path.basename(sys.executable)],
            capture_output=True,
        )
        print("Stopped.")
        return

    if not _PYWIN32_AVAILABLE:
        print("ERROR: pywin32 is required for service management.")
        sys.exit(1)

    win32serviceutil.HandleCommandLine(MpcHcBridgeService)


if __name__ == "__main__":
    main()
