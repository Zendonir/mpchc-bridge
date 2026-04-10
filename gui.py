"""
MPC-HC Bridge — GUI Installer / Control Panel
Provides a simple tkinter window to install, uninstall, start and stop the service.
"""

import subprocess
import sys
import threading
import tkinter as tk
import webbrowser
from tkinter import font, messagebox, scrolledtext

BRIDGE_PORT = 13580
SVC_NAME = "MpcHcBridge"

# ── Service helpers (via sc.exe — no pywin32 needed in the GUI thread) ─────────


def _sc(*args: str, timeout: int = 10) -> tuple[int, str]:
    """Run an sc.exe command and return (returncode, output)."""
    try:
        r = subprocess.run(
            ["sc"] + list(args),
            capture_output=True,
            text=True,
            timeout=timeout,
        )
        return r.returncode, ((r.stdout or "") + (r.stderr or "")).strip()
    except subprocess.TimeoutExpired:
        return -1, f"Timeout after {timeout}s — service may still be starting"
    except Exception as ex:  # pylint: disable=broad-exception-caught
        return -1, str(ex)


def _self_exe() -> str:
    """Return path to the running executable."""
    return sys.executable if not getattr(sys, "frozen", False) else sys.argv[0]


def svc_status() -> str:
    """Return 'running', 'stopped', 'not_installed' or 'unknown'."""
    code, out = _sc("query", SVC_NAME)
    if code != 0:
        return "not_installed"
    if "RUNNING" in out:
        return "running"
    if "STOPPED" in out:
        return "stopped"
    return "unknown"


def svc_install() -> tuple[bool, str]:
    """Install and configure the service for auto-start."""
    exe = _self_exe()
    code, out = _sc(
        "create", SVC_NAME,
        f"binPath={exe} --service",
        "start=auto",
        "DisplayName=MPC-HC Control Bridge",
    )
    if code != 0:
        return False, out
    # Add description
    _sc("description", SVC_NAME,
        "HTTP API bridge for UC Remote 3 — full MPC-HC control")
    return True, out


def svc_uninstall() -> tuple[bool, str]:
    _sc("stop", SVC_NAME)
    code, out = _sc("delete", SVC_NAME)
    return code == 0, out


def svc_start() -> tuple[bool, str]:
    code, out = _sc("start", SVC_NAME, timeout=30)
    return code == 0, out


def svc_stop() -> tuple[bool, str]:
    code, out = _sc("stop", SVC_NAME)
    return code == 0, out


# ── GUI ────────────────────────────────────────────────────────────────────────

# Colours
CLR_BG = "#1e1e2e"
CLR_PANEL = "#2a2a3e"
CLR_ACCENT = "#89b4fa"
CLR_GREEN = "#a6e3a1"
CLR_RED = "#f38ba8"
CLR_YELLOW = "#f9e2af"
CLR_TEXT = "#cdd6f4"
CLR_MUTED = "#6c7086"
CLR_BTN = "#313244"
CLR_BTN_HOVER = "#45475a"


class App(tk.Tk):
    """Main application window."""

    def __init__(self) -> None:
        super().__init__()
        self.title("MPC-HC Bridge")
        self.resizable(False, False)
        self.configure(bg=CLR_BG)
        self._build_ui()
        self.after(100, self._refresh_status)
        # Refresh every 3 s
        self._schedule_refresh()

    # ── UI construction ────────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        pad = {"padx": 16, "pady": 8}

        # Header
        header = tk.Frame(self, bg=CLR_BG)
        header.pack(fill="x", **pad)
        tk.Label(
            header, text="MPC-HC Bridge", bg=CLR_BG, fg=CLR_ACCENT,
            font=("Segoe UI", 16, "bold"),
        ).pack(side="left")
        tk.Label(
            header, text=f"port {BRIDGE_PORT}", bg=CLR_BG, fg=CLR_MUTED,
            font=("Segoe UI", 10),
        ).pack(side="left", padx=(8, 0), pady=(4, 0))

        # Status row
        status_frame = tk.Frame(self, bg=CLR_PANEL, pady=10)
        status_frame.pack(fill="x", padx=16, pady=(0, 8))
        tk.Label(status_frame, text="Status:", bg=CLR_PANEL, fg=CLR_MUTED,
                 font=("Segoe UI", 10)).pack(side="left", padx=(12, 6))
        self._dot = tk.Label(status_frame, text="●", bg=CLR_PANEL,
                             font=("Segoe UI", 14))
        self._dot.pack(side="left")
        self._status_lbl = tk.Label(status_frame, bg=CLR_PANEL, fg=CLR_TEXT,
                                    font=("Segoe UI", 10, "bold"))
        self._status_lbl.pack(side="left", padx=(4, 0))

        # Buttons
        btn_frame = tk.Frame(self, bg=CLR_BG)
        btn_frame.pack(fill="x", padx=16, pady=(0, 8))

        self._btn_install = self._btn(btn_frame, "Install Service", self._on_install)
        self._btn_install.pack(side="left", padx=(0, 6))

        self._btn_uninstall = self._btn(btn_frame, "Uninstall", self._on_uninstall)
        self._btn_uninstall.pack(side="left", padx=(0, 6))

        self._btn_start = self._btn(btn_frame, "▶  Start", self._on_start)
        self._btn_start.pack(side="left", padx=(0, 6))

        self._btn_stop = self._btn(btn_frame, "■  Stop", self._on_stop)
        self._btn_stop.pack(side="left", padx=(0, 6))

        self._btn_browser = self._btn(
            btn_frame, "🌐  Test in Browser", self._on_browser, accent=True
        )
        self._btn_browser.pack(side="right")

        # Log output
        log_frame = tk.Frame(self, bg=CLR_BG)
        log_frame.pack(fill="both", padx=16, pady=(0, 12))
        tk.Label(log_frame, text="Output", bg=CLR_BG, fg=CLR_MUTED,
                 font=("Segoe UI", 9)).pack(anchor="w")
        self._log = scrolledtext.ScrolledText(
            log_frame, height=8, width=54,
            bg=CLR_PANEL, fg=CLR_TEXT, insertbackground=CLR_TEXT,
            font=("Consolas", 9), relief="flat", state="disabled",
            wrap="word",
        )
        self._log.pack(fill="both")

        # Geometry — centre on screen
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        x = (self.winfo_screenwidth() - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"+{x}+{y}")

    def _btn(self, parent: tk.Frame, text: str, cmd, accent: bool = False) -> tk.Button:
        bg = CLR_ACCENT if accent else CLR_BTN
        fg = CLR_BG if accent else CLR_TEXT
        b = tk.Button(
            parent, text=text, command=cmd,
            bg=bg, fg=fg, activebackground=CLR_BTN_HOVER, activeforeground=CLR_TEXT,
            relief="flat", padx=10, pady=6,
            font=("Segoe UI", 9),
            cursor="hand2",
        )
        return b

    # ── Status refresh ─────────────────────────────────────────────────────────

    def _refresh_status(self) -> None:
        status = svc_status()
        if status == "running":
            self._dot.config(fg=CLR_GREEN)
            self._status_lbl.config(text="Running", fg=CLR_GREEN)
        elif status == "stopped":
            self._dot.config(fg=CLR_YELLOW)
            self._status_lbl.config(text="Stopped", fg=CLR_YELLOW)
        elif status == "not_installed":
            self._dot.config(fg=CLR_MUTED)
            self._status_lbl.config(text="Not installed", fg=CLR_MUTED)
        else:
            self._dot.config(fg=CLR_RED)
            self._status_lbl.config(text="Unknown", fg=CLR_RED)

        installed = status != "not_installed"
        running = status == "running"
        self._btn_install.config(state="disabled" if installed else "normal")
        self._btn_uninstall.config(state="normal" if installed else "disabled")
        self._btn_start.config(state="disabled" if running else (
            "normal" if installed else "disabled"))
        self._btn_stop.config(state="normal" if running else "disabled")
        self._btn_browser.config(state="normal" if running else "disabled")

    def _schedule_refresh(self) -> None:
        self.after(3000, self._auto_refresh)

    def _auto_refresh(self) -> None:
        self._refresh_status()
        self._schedule_refresh()

    # ── Logging ────────────────────────────────────────────────────────────────

    def _log_write(self, msg: str) -> None:
        self._log.config(state="normal")
        self._log.insert("end", msg + "\n")
        self._log.see("end")
        self._log.config(state="disabled")

    # ── Button handlers ────────────────────────────────────────────────────────

    def _run_in_thread(self, fn) -> None:
        threading.Thread(target=fn, daemon=True).start()

    def _on_install(self) -> None:
        def _do():
            self._log_write("Installing service…")
            ok, out = svc_install()
            self._log_write(out)
            if ok:
                self._log_write("✔  Service installed. Starting…")
                ok2, out2 = svc_start()
                self._log_write(out2)
                if ok2:
                    self._log_write("✔  Service started.")
            else:
                self._log_write("✘  Install failed — run as Administrator?")
            self.after(0, self._refresh_status)
        self._run_in_thread(_do)

    def _on_uninstall(self) -> None:
        if not messagebox.askyesno(
            "Uninstall", "Stop and remove the MPC-HC Bridge service?"
        ):
            return
        def _do():
            self._log_write("Uninstalling service…")
            ok, out = svc_uninstall()
            self._log_write(out)
            self._log_write("✔  Done." if ok else "✘  Failed.")
            self.after(0, self._refresh_status)
        self._run_in_thread(_do)

    def _on_start(self) -> None:
        def _do():
            self._log_write("Starting service…")
            ok, out = svc_start()
            self._log_write(out)
            self._log_write("✔  Running." if ok else "✘  Failed.")
            self.after(0, self._refresh_status)
        self._run_in_thread(_do)

    def _on_stop(self) -> None:
        def _do():
            self._log_write("Stopping service…")
            ok, out = svc_stop()
            self._log_write(out)
            self._log_write("✔  Stopped." if ok else "✘  Failed.")
            self.after(0, self._refresh_status)
        self._run_in_thread(_do)

    def _on_browser(self) -> None:
        webbrowser.open(f"http://localhost:{BRIDGE_PORT}/status")


# ── Entry point ────────────────────────────────────────────────────────────────


def main() -> None:
    """Launch the GUI."""
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
