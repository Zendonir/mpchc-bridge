"""
MPC-HC Bridge — Windows Service wrapper.

Usage (run as Administrator):
  mpchc-bridge.exe install   — install as Windows service (auto-start)
  mpchc-bridge.exe start     — start the service
  mpchc-bridge.exe stop      — stop the service
  mpchc-bridge.exe restart   — restart the service
  mpchc-bridge.exe remove    — uninstall the service
  mpchc-bridge.exe status    — show current status
  mpchc-bridge.exe debug     — run in foreground (console mode, no service)
"""

import asyncio
import os
import subprocess
import sys


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


_TASK_NAME = "MpcHcBridge"


def _install_interactive() -> None:
    """Install MPC-HC Bridge as a Task Scheduler task for the current user.

    Runs at logon with full user privileges — no password required, and the
    task automatically has access to mapped network drives (e.g. Y:\\).
    """
    exe = os.path.abspath(sys.executable if getattr(sys, "frozen", False) else sys.argv[0])
    # When frozen (PyInstaller), sys.executable IS the exe.
    # In source mode, use this script's path but run with 'debug' arg.
    if not getattr(sys, "frozen", False):
        cmd = f'"{exe}" "{os.path.abspath(__file__)}" debug'
    else:
        cmd = f'"{exe}" debug'

    result = subprocess.run(
        [
            "schtasks", "/create",
            "/tn", _TASK_NAME,
            "/tr", cmd,
            "/sc", "ONLOGON",
            "/rl", "HIGHEST",
            "/f",
        ],
        capture_output=True, text=True,
    )
    if result.returncode == 0:
        print(f"Installed: task '{_TASK_NAME}' will start at logon for current user.")
        print("Run 'start' to launch it now without rebooting.")
    else:
        print(f"ERROR installing task:\n{result.stderr or result.stdout}")
        sys.exit(1)


def _remove_task() -> None:
    result = subprocess.run(
        ["schtasks", "/delete", "/tn", _TASK_NAME, "/f"],
        capture_output=True, text=True,
    )
    if result.returncode == 0:
        print(f"Removed task '{_TASK_NAME}'.")
    else:
        print(f"ERROR: {result.stderr or result.stdout}")


def _start_task() -> None:
    result = subprocess.run(
        ["schtasks", "/run", "/tn", _TASK_NAME],
        capture_output=True, text=True,
    )
    if result.returncode == 0:
        print(f"Task '{_TASK_NAME}' started.")
    else:
        print(f"ERROR: {result.stderr or result.stdout}")


def _status_task() -> None:
    result = subprocess.run(
        ["schtasks", "/query", "/tn", _TASK_NAME, "/fo", "LIST"],
        capture_output=True, text=True,
    )
    print(result.stdout or result.stderr)


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
