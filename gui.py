"""
MPC-HC Bridge — GUI Installer / Control Panel
Provides a simple tkinter window to install, uninstall, start and stop the service.
"""

import json
import subprocess
import sys
import threading
import tkinter as tk
import urllib.error
import urllib.parse
import urllib.request
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


# ── Windows Firewall helpers ───────────────────────────────────────────────────

FW_RULE_NAME = "MPC-HC Bridge (Port 13580)"


def _netsh(*args: str, timeout: int = 10) -> tuple[int, str]:
    """Run a netsh command and return (returncode, output)."""
    try:
        r = subprocess.run(
            ["netsh"] + list(args),
            capture_output=True,
            text=True,
            timeout=timeout,
        )
        return r.returncode, ((r.stdout or "") + (r.stderr or "")).strip()
    except subprocess.TimeoutExpired:
        return -1, f"Timeout after {timeout}s"
    except Exception as ex:  # pylint: disable=broad-exception-caught
        return -1, str(ex)


def fw_add_rule() -> tuple[bool, str]:
    """Add inbound TCP firewall rule for port 13580."""
    # Remove existing rule first to avoid duplicates
    _netsh("advfirewall", "firewall", "delete", "rule", f'name={FW_RULE_NAME}')
    code, out = _netsh(
        "advfirewall", "firewall", "add", "rule",
        f"name={FW_RULE_NAME}",
        "dir=in",
        "action=allow",
        "protocol=TCP",
        f"localport={BRIDGE_PORT}",
        "profile=any",
        "description=Allows UC Remote to connect to MPC-HC Bridge",
    )
    return code == 0, out


def fw_remove_rule() -> tuple[bool, str]:
    """Remove the inbound firewall rule for port 13580."""
    code, out = _netsh("advfirewall", "firewall", "delete", "rule", f"name={FW_RULE_NAME}")
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

        self._btn_firewall = self._btn(btn_frame, "🔓  Firewall", self._on_firewall)
        self._btn_firewall.pack(side="left", padx=(0, 6))

        self._btn_test = self._btn(btn_frame, "🧪  Test Controls", self._on_test)
        self._btn_test.pack(side="left", padx=(0, 6))

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
                self._log_write(f"Adding firewall rule for port {BRIDGE_PORT}…")
                ok3, out3 = fw_add_rule()
                self._log_write(out3)
                self._log_write(f"✔  Firewall rule added." if ok3 else "⚠  Firewall rule failed — add manually if needed.")
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
            self._log_write("Removing firewall rule…")
            _, out2 = fw_remove_rule()
            self._log_write(out2)
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

    def _on_firewall(self) -> None:
        def _do():
            self._log_write(f"Adding firewall rule for port {BRIDGE_PORT} (all profiles)…")
            ok, out = fw_add_rule()
            self._log_write(out)
            self._log_write("✔  Firewall rule set." if ok else "✘  Failed — run as Administrator?")
        self._run_in_thread(_do)

    def _on_browser(self) -> None:
        webbrowser.open(f"http://localhost:{BRIDGE_PORT}/status")

    def _on_test(self) -> None:
        TestWindow(self)


# ── Test / control window ──────────────────────────────────────────────────────


def _bridge_get(path: str) -> dict:
    """Synchronous GET to the bridge. Returns parsed JSON or error dict."""
    try:
        url = f"http://127.0.0.1:{BRIDGE_PORT}{path}"
        with urllib.request.urlopen(url, timeout=3) as r:
            return json.loads(r.read().decode())
    except Exception as ex:  # pylint: disable=broad-exception-caught
        return {"error": str(ex)}


def _bridge_post(path: str, params: dict | None = None) -> dict:
    """Synchronous POST to the bridge. Returns parsed JSON or error dict."""
    try:
        url = f"http://127.0.0.1:{BRIDGE_PORT}{path}"
        if params:
            url += "?" + urllib.parse.urlencode(params)
        req = urllib.request.Request(url, data=b"", method="POST")
        with urllib.request.urlopen(req, timeout=3) as r:
            return json.loads(r.read().decode())
    except Exception as ex:  # pylint: disable=broad-exception-caught
        return {"error": str(ex)}


class TestWindow(tk.Toplevel):
    """Full test & control panel for the MPC-HC Bridge."""

    def __init__(self, parent: tk.Tk) -> None:
        super().__init__(parent)
        self.title("MPC-HC Bridge — Test Controls")
        self.resizable(False, False)
        self.configure(bg=CLR_BG)
        self._build_ui()
        self._refresh_status()
        self._refresh_tracks()

    # ── helpers ────────────────────────────────────────────────────────────────

    def _run(self, fn) -> None:
        threading.Thread(target=fn, daemon=True).start()

    def _cmd(self, name: str) -> None:
        self._run(lambda: self._show(_bridge_post(f"/command/{name}")))

    def _show(self, result: dict) -> None:
        err = result.get("error")
        if err:
            self._set_status_text(f"Error: {err}", CLR_RED)
        self._refresh_log()

    def _refresh_log(self) -> None:
        data = _bridge_get("/debug/log")
        entries = data.get("log", [])
        if entries:
            self.after(0, lambda: self._write_log(entries))

    def _write_log(self, entries: list) -> None:
        self._log.config(state="normal")
        self._log.delete("1.0", "end")
        self._log.insert("end", "\n".join(entries))

    def _dump_raw(self, path: str) -> None:
        try:
            url = f"http://127.0.0.1:{BRIDGE_PORT}{path}"
            with urllib.request.urlopen(url, timeout=3) as r:
                text = r.read().decode(errors="replace")
            self.after(0, lambda: self._write_log([f"=== {path} ===", text]))
        except Exception as ex:  # pylint: disable=broad-exception-caught
            self.after(0, lambda: self._write_log([f"Error fetching {path}: {ex}"]))

    def _dump_controls(self) -> None:
        self._dump_raw("/debug/controls")

    def _dump_variables(self) -> None:
        self._dump_raw("/debug/variables")
        self._log.see("end")
        self._log.config(state="disabled")

    def _lbl(self, parent, text: str, fg=None, font_size: int = 9) -> tk.Label:
        return tk.Label(parent, text=text, bg=CLR_BG, fg=fg or CLR_MUTED,
                        font=("Segoe UI", font_size))

    def _section(self, parent, title: str) -> tk.LabelFrame:
        f = tk.LabelFrame(parent, text=title, bg=CLR_BG, fg=CLR_ACCENT,
                          font=("Segoe UI", 9, "bold"),
                          bd=1, relief="groove", padx=6, pady=4)
        return f

    def _btn(self, parent, text: str, cmd, width: int = 8) -> tk.Button:
        return tk.Button(parent, text=text, command=cmd,
                         bg=CLR_BTN, fg=CLR_TEXT, activebackground=CLR_BTN_HOVER,
                         activeforeground=CLR_TEXT, relief="flat",
                         padx=4, pady=4, width=width,
                         font=("Segoe UI", 9), cursor="hand2")

    # ── UI ─────────────────────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        pad = {"padx": 10, "pady": 4}

        # ── Status display ────────────────────────────────────────────────────
        sf = self._section(self, "Status")
        sf.pack(fill="x", **pad)

        self._status_text = tk.StringVar(value="—")
        tk.Label(sf, textvariable=self._status_text, bg=CLR_BG, fg=CLR_TEXT,
                 font=("Consolas", 9), justify="left", wraplength=460).pack(anchor="w")

        tk.Button(sf, text="↺  Refresh", command=self._refresh_status,
                  bg=CLR_BTN, fg=CLR_TEXT, activebackground=CLR_BTN_HOVER,
                  relief="flat", padx=6, pady=2,
                  font=("Segoe UI", 9), cursor="hand2").pack(anchor="e")

        # ── Playback ──────────────────────────────────────────────────────────
        pf = self._section(self, "Playback")
        pf.pack(fill="x", **pad)

        row1 = tk.Frame(pf, bg=CLR_BG)
        row1.pack()
        for text, name in [
            ("⏮", "prev"), ("⏪⏪", "seek_bwd_large"), ("⏪", "seek_bwd_small"),
            ("⏯", "play_pause"),
            ("⏩", "seek_fwd_small"), ("⏩⏩", "seek_fwd_large"), ("⏭", "next"),
        ]:
            self._btn(row1, text, lambda n=name: self._cmd(n), width=4).pack(side="left", padx=2)

        row2 = tk.Frame(pf, bg=CLR_BG)
        row2.pack(pady=(4, 0))
        for text, name in [
            ("■ Stop", "stop"), ("⏏ Close", "close"),
            ("◀ Frame", "frame_step_back"), ("Frame ▶", "frame_step"),
            ("Speed −", "speed_down"), ("Speed ✕", "speed_reset"), ("Speed +", "speed_up"),
        ]:
            self._btn(row2, text, lambda n=name: self._cmd(n), width=7).pack(side="left", padx=2)

        # ── Seek ──────────────────────────────────────────────────────────────
        skf = self._section(self, "Seek to position")
        skf.pack(fill="x", **pad)

        sk_row = tk.Frame(skf, bg=CLR_BG)
        sk_row.pack()
        self._lbl(sk_row, "Position (s):").pack(side="left")
        self._seek_var = tk.StringVar()
        tk.Entry(sk_row, textvariable=self._seek_var, width=8,
                 bg=CLR_PANEL, fg=CLR_TEXT, insertbackground=CLR_TEXT,
                 relief="flat", font=("Consolas", 10)).pack(side="left", padx=4)
        self._btn(sk_row, "Go", self._on_seek, width=4).pack(side="left")

        # ── Volume ────────────────────────────────────────────────────────────
        vf = self._section(self, "Volume")
        vf.pack(fill="x", **pad)

        v_row = tk.Frame(vf, bg=CLR_BG)
        v_row.pack()
        for text, name in [("🔉 −", "vol_down"), ("🔊 +", "vol_up"), ("🔇 Mute", "mute")]:
            self._btn(v_row, text, lambda n=name: self._cmd(n), width=7).pack(side="left", padx=2)

        tk.Frame(v_row, bg=CLR_MUTED, width=1, height=24).pack(side="left", padx=8)

        self._lbl(v_row, "Set level (0-100):").pack(side="left")
        self._vol_var = tk.StringVar()
        tk.Entry(v_row, textvariable=self._vol_var, width=5,
                 bg=CLR_PANEL, fg=CLR_TEXT, insertbackground=CLR_TEXT,
                 relief="flat", font=("Consolas", 10)).pack(side="left", padx=4)
        self._btn(v_row, "Set", self._on_set_volume, width=4).pack(side="left")

        # ── Subtitles ─────────────────────────────────────────────────────────
        subf = self._section(self, "Subtitles")
        subf.pack(fill="x", **pad)

        sub_row = tk.Frame(subf, bg=CLR_BG)
        sub_row.pack()
        for text, name in [
            ("◀ Prev", "sub_prev"), ("Next ▶", "sub_next"),
            ("Delay −", "sub_delay_minus"), ("Delay +", "sub_delay_plus"),
        ]:
            self._btn(sub_row, text, lambda n=name: self._cmd(n), width=7).pack(side="left", padx=2)

        # ── Audio ─────────────────────────────────────────────────────────────
        af = self._section(self, "Audio")
        af.pack(fill="x", **pad)

        a_row = tk.Frame(af, bg=CLR_BG)
        a_row.pack()
        for text, name in [
            ("Next Track", "audio_next"),
            ("Delay −", "audio_delay_minus"), ("Delay +", "audio_delay_plus"),
        ]:
            self._btn(a_row, text, lambda n=name: self._cmd(n), width=9).pack(side="left", padx=2)

        # ── View ──────────────────────────────────────────────────────────────
        vwf = self._section(self, "View")
        vwf.pack(fill="x", **pad)

        vw_row = tk.Frame(vwf, bg=CLR_BG)
        vw_row.pack()
        for text, name in [
            ("Fullscreen", "fullscreen"), ("Fit", "zoom_fit"),
            ("50%", "zoom_50"), ("100%", "zoom_100"), ("200%", "zoom_200"),
        ]:
            self._btn(vw_row, text, lambda n=name: self._cmd(n), width=9).pack(side="left", padx=2)

        # ── Open file ─────────────────────────────────────────────────────────
        of = self._section(self, "Open file in MPC-HC")
        of.pack(fill="x", **pad)

        o_row = tk.Frame(of, bg=CLR_BG)
        o_row.pack(fill="x")
        self._path_var = tk.StringVar()
        tk.Entry(o_row, textvariable=self._path_var, width=46,
                 bg=CLR_PANEL, fg=CLR_TEXT, insertbackground=CLR_TEXT,
                 relief="flat", font=("Consolas", 9)).pack(side="left", padx=(0, 4), fill="x", expand=True)
        self._btn(o_row, "Open", self._on_open, width=5).pack(side="left")

        # ── Tracks ────────────────────────────────────────────────────────────
        tf = self._section(self, "Tracks")
        tf.pack(fill="x", **pad)

        tr_hdr = tk.Frame(tf, bg=CLR_BG)
        tr_hdr.pack(fill="x")
        self._lbl(tr_hdr, "Audio", fg=CLR_ACCENT).pack(side="left")
        tk.Button(tr_hdr, text="↺ Refresh", command=lambda: self._run(self._refresh_tracks),
                  bg=CLR_BTN, fg=CLR_TEXT, activebackground=CLR_BTN_HOVER,
                  relief="flat", padx=4, pady=1, font=("Segoe UI", 9), cursor="hand2").pack(side="right")

        self._audio_frame = tk.Frame(tf, bg=CLR_BG)
        self._audio_frame.pack(fill="x", pady=(2, 4))
        tk.Label(self._audio_frame, text="—", bg=CLR_BG, fg=CLR_MUTED,
                 font=("Segoe UI", 9)).pack(anchor="w")

        self._lbl(tf, "Subtitles", fg=CLR_ACCENT).pack(anchor="w")
        self._sub_frame = tk.Frame(tf, bg=CLR_BG)
        self._sub_frame.pack(fill="x", pady=(2, 0))
        tk.Label(self._sub_frame, text="—", bg=CLR_BG, fg=CLR_MUTED,
                 font=("Segoe UI", 9)).pack(anchor="w")

        # ── Debug log ─────────────────────────────────────────────────────────
        lf = self._section(self, "Debug Log")
        lf.pack(fill="x", **pad)

        log_row = tk.Frame(lf, bg=CLR_BG)
        log_row.pack(fill="x")
        self._log = scrolledtext.ScrolledText(
            log_row, height=6, width=66,
            bg=CLR_PANEL, fg=CLR_YELLOW, insertbackground=CLR_TEXT,
            font=("Consolas", 8), relief="flat", state="disabled", wrap="none",
        )
        self._log.pack(fill="x")
        btn_row = tk.Frame(log_row, bg=CLR_BG)
        btn_row.pack(fill="x", pady=(2, 0))
        tk.Button(btn_row, text="↺  Refresh Log", command=lambda: self._run(self._refresh_log),
                  bg=CLR_BTN, fg=CLR_TEXT, activebackground=CLR_BTN_HOVER,
                  relief="flat", padx=6, pady=2,
                  font=("Segoe UI", 9), cursor="hand2").pack(side="left")
        tk.Button(btn_row, text="Dump controls.html",
                  command=lambda: self._run(self._dump_controls),
                  bg=CLR_BTN, fg=CLR_TEXT, activebackground=CLR_BTN_HOVER,
                  relief="flat", padx=6, pady=2,
                  font=("Segoe UI", 9), cursor="hand2").pack(side="left", padx=(4, 0))
        tk.Button(btn_row, text="Dump variables.html",
                  command=lambda: self._run(self._dump_variables),
                  bg=CLR_BTN, fg=CLR_TEXT, activebackground=CLR_BTN_HOVER,
                  relief="flat", padx=6, pady=2,
                  font=("Segoe UI", 9), cursor="hand2").pack(side="left", padx=(4, 0))

        # centre window
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        x = (self.winfo_screenwidth() - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"+{x}+{y}")

    # ── Status refresh ─────────────────────────────────────────────────────────

    def _set_status_text(self, text: str, color: str = CLR_TEXT) -> None:
        self._status_text.set(text)

    def _refresh_status(self) -> None:
        def _do():
            data = _bridge_get("/status")
            if "error" in data:
                self.after(0, lambda: self._set_status_text(f"⚠  {data['error']}", CLR_RED))
                return
            pos_ms = data.get("position_ms", 0)
            lines = [
                f"State:     {data.get('state', '?').upper()}   "
                f"Position:  {data.get('position_str', '?')} / {data.get('duration_str', '?')}  ({pos_ms} ms)",
                f"Volume:    {data.get('volume', '?')}%   "
                f"Muted:     {'Yes' if data.get('muted') else 'No'}   "
                f"Rate:      {data.get('playback_rate', 1.0)}x",
                f"File:      {data.get('file') or '—'}",
            ]
            self.after(0, lambda: self._set_status_text("\n".join(lines)))
        self._run(_do)

    def _refresh_tracks(self) -> None:
        def _do():
            data = _bridge_get("/tracks")
            if "error" in data:
                self.after(0, lambda: self._write_log([f"Tracks: {data['error']}"]))
                return
            self.after(0, lambda: self._build_track_buttons(data))
        self._run(_do)

    def _build_track_buttons(self, data: dict) -> None:
        for frame, kind, cmd_prev, cmd_next in [
            (self._audio_frame, "audio", "audio_prev", "audio_next"),
            (self._sub_frame,   "subtitle", "sub_prev", "sub_next"),
        ]:
            for w in frame.winfo_children():
                w.destroy()
            tracks = data.get(kind, [])
            if not tracks:
                tk.Label(frame, text="—", bg=CLR_BG, fg=CLR_MUTED,
                         font=("Segoe UI", 9)).pack(anchor="w")
                continue
            for t in tracks:
                pos = t["pos"]
                active = t.get("selected", False)
                lang = t.get("lang", "").upper() or "?"
                name = t.get("name", "") or ""
                codec = t.get("codec", "")
                label = f"{'▶  ' if active else '    '}{lang}  {codec}" + (f"  [{name}]" if name else "")
                bg = CLR_ACCENT if active else CLR_BTN
                fg = CLR_BG if active else CLR_TEXT
                tk.Button(frame, text=label, bg=bg, fg=fg,
                          activebackground=CLR_BTN_HOVER, activeforeground=CLR_TEXT,
                          relief="flat", padx=6, pady=3, anchor="w", width=50,
                          font=("Consolas", 8), cursor="hand2",
                          command=lambda k=kind, p=pos: self._run(
                              lambda: self._select_track_and_refresh(k, p)
                          )).pack(fill="x", pady=1)
            # cycling controls below the list
            ctrl = tk.Frame(frame, bg=CLR_BG)
            ctrl.pack(anchor="w", pady=(3, 0))
            self._btn(ctrl, "◀ Prev", lambda c=cmd_prev: self._cmd(c), width=6).pack(side="left", padx=(0, 2))
            self._btn(ctrl, "Next ▶", lambda c=cmd_next: self._cmd(c), width=6).pack(side="left")

    def _select_track_and_refresh(self, kind: str, pos: int) -> None:
        result = _bridge_post(f"/{kind}/select/{pos}")
        self._show(result)
        import time as _time
        _time.sleep(0.4)
        data = _bridge_get("/tracks")
        if "error" not in data:
            self.after(0, lambda: self._build_track_buttons(data))

    def _show_and_refresh(self, result: dict) -> None:
        self._show(result)
        import time as _time
        _time.sleep(0.3)
        self._refresh_tracks()

    # ── Action handlers ────────────────────────────────────────────────────────

    def _on_seek(self) -> None:
        try:
            pos_ms = int(float(self._seek_var.get()) * 1000)
        except ValueError:
            return
        self._run(lambda: self._show(_bridge_post("/seek", {"pos_ms": pos_ms})))

    def _on_set_volume(self) -> None:
        try:
            level = max(0, min(100, int(self._vol_var.get())))
        except ValueError:
            return
        self._run(lambda: self._show(_bridge_post("/volume", {"level": level})))

    def _on_open(self) -> None:
        path = self._path_var.get().strip()
        if path:
            self._run(lambda: self._show(_bridge_post("/open", {"path": path})))


# ── Entry point ────────────────────────────────────────────────────────────────


def main() -> None:
    """Launch the GUI."""
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
