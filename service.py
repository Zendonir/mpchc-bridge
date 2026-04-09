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


def main() -> None:
    """Dispatch install/start/stop/remove/debug commands."""
    if not _PYWIN32_AVAILABLE:
        print("ERROR: pywin32 is required for service management.")
        print("       Install with: pip install pywin32")
        sys.exit(1)

    args = sys.argv[1:]

    if not args or args[0] in ("-h", "--help"):
        _print_usage()
        return

    if args[0] == "debug":
        # Run in foreground without SCM
        from bridge import BRIDGE_PORT, main as bridge_main
        print(f"Running in debug mode on port {BRIDGE_PORT} — Ctrl+C to stop")
        bridge_main()
        return

    if args[0] == "status":
        import subprocess
        result = subprocess.run(
            ["sc", "query", "MpcHcBridge"],
            capture_output=True, text=True
        )
        print(result.stdout or result.stderr)
        return

    # Delegate install/start/stop/restart/remove to pywin32
    # pywin32 reads sys.argv directly
    if len(sys.argv) == 1:
        # Called by SCM without arguments — start dispatcher
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(MpcHcBridgeService)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(MpcHcBridgeService)


if __name__ == "__main__":
    main()
