"""
MPC-HC Control Bridge
Runs on the Windows PC alongside MPC-HC.
Exposes a full HTTP API so the UC Remote driver can control everything.

Endpoints:
  GET  /status                   — playback state, position, volume, file
  GET  /tracks                   — available audio + subtitle tracks
  POST /command/{cmd}            — named playback command (see CMD dict)
  POST /seek?pos_ms=<int>        — absolute seek in milliseconds
  POST /volume?level=<0-100>     — set exact volume
  POST /audio/{index}            — select audio track by index
  POST /subtitle/{index}         — select subtitle track (-1 = disable)
  POST /open?path=<filepath>     — open file in MPC-HC
"""

import re
import subprocess
import sys
from pathlib import Path

from aiohttp import ClientSession, ClientTimeout, web

# ── win32 via ctypes (no external dependency) ──────────────────────────────────
if sys.platform == "win32":
    import ctypes

    _user32 = ctypes.windll.user32
else:
    _user32 = None

MPCHC_CLASS = "MediaPlayerClassicW"
WM_COMMAND = 0x0111


def _find_hwnd() -> int:
    if _user32 is None:
        return 0
    return _user32.FindWindowW(MPCHC_CLASS, None)  # type: ignore[attr-defined]


def _post_wm_command(cmd_id: int) -> bool:
    hwnd = _find_hwnd()
    if not hwnd:
        return False
    _user32.PostMessageW(hwnd, WM_COMMAND, cmd_id, 0)  # type: ignore[attr-defined]
    return True


# ── MPC-HC command IDs ─────────────────────────────────────────────────────────
CMD: dict[str, int] = {
    # Playback
    "play_pause": 889,
    "stop": 890,
    "goto_start": 891,
    "frame_step": 892,
    "frame_step_back": 893,
    "speed_down": 894,
    "speed_up": 895,
    "speed_reset": 896,
    "seek_fwd_small": 899,  # ~5 s
    "seek_bwd_small": 900,
    "seek_fwd_large": 901,  # ~1 min
    "seek_bwd_large": 902,
    "prev": 921,
    "next": 922,
    # Volume
    "vol_up": 907,
    "vol_down": 908,
    "mute": 909,
    # View
    "fullscreen": 830,
    "zoom_25": 832,
    "zoom_50": 833,
    "zoom_100": 834,
    "zoom_200": 835,
    "zoom_fit": 836,
    # Subtitles
    "sub_toggle": 955,
    "sub_next": 954,
    "sub_prev": 953,
    "sub_delay_plus": 958,
    "sub_delay_minus": 957,
    # Audio
    "audio_next": 952,
    "audio_delay_plus": 946,
    "audio_delay_minus": 945,
    # Chapters
    "chapter_next": 918,
    "chapter_prev": 916,
    # File
    "close": 808,
    # DVD
    "dvd_menu_title": 923,
    "dvd_menu_root": 924,
    "dvd_menu_sub": 925,
    "dvd_menu_audio": 926,
    "dvd_menu_angle": 927,
    "dvd_menu_chapter": 928,
}

# Base WM_COMMAND offsets for track selection
_AUDIO_BASE = 50000
_SUB_BASE = 50935
_VIDEO_BASE = 50200

MPCHC_PORT = 13579
BRIDGE_PORT = 13580
_TIMEOUT = ClientTimeout(total=3)


# ── MPC-HC HTTP helpers ────────────────────────────────────────────────────────


async def _mpchc_get(
    session: ClientSession, path: str, params: dict | None = None
) -> str | None:
    try:
        url = f"http://localhost:{MPCHC_PORT}{path}"
        async with session.get(url, params=params, allow_redirects=False) as r:
            if r.status in (200, 302):
                return await r.text()
    except Exception:  # pylint: disable=broad-exception-caught
        pass
    return None


def _parse_variables(html: str) -> dict[str, str]:
    """Parse /variables.html <p id="key">value</p> tags."""
    return {
        m.group(1): m.group(2).strip()
        for m in re.finditer(r'<p\s+id="([^"]+)">([^<]*)</p>', html)
    }


def _parse_tracks(html: str) -> dict[str, list]:
    """Parse audio/subtitle track lists from /controls.html select elements."""
    result: dict[str, list] = {"audio": [], "subtitle": []}
    for sel_id, key in (("audiotrackid", "audio"), ("subtrackid", "subtitle")):
        m = re.search(
            rf'id="{sel_id}"[^>]*>(.*?)</select>', html, re.DOTALL | re.IGNORECASE
        )
        if m:
            for opt in re.finditer(
                r'<option\s+value="(-?\d+)"([^>]*)>([^<]+)</option>', m.group(1)
            ):
                result[key].append(
                    {
                        "index": int(opt.group(1)),
                        "name": opt.group(3).strip(),
                        "selected": "selected" in opt.group(2).lower(),
                    }
                )
    return result


def _find_mpchc_exe() -> str | None:
    """Locate mpc-hc64.exe via registry or common install paths."""
    if sys.platform != "win32":
        return None
    try:
        import winreg

        for key_path in (
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\mpc-hc64.exe",
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\mpc-hc.exe",
        ):
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path) as k:
                    return winreg.QueryValue(k, None)
            except OSError:
                pass
    except ImportError:
        pass
    for candidate in (
        r"C:\Program Files\MPC-HC\mpc-hc64.exe",
        r"C:\Program Files (x86)\MPC-HC\mpc-hc64.exe",
        r"C:\Program Files\MPC-HC\mpc-hc.exe",
        r"C:\Program Files (x86)\MPC-HC\mpc-hc.exe",
    ):
        if Path(candidate).exists():
            return candidate
    return None


# ── Request handlers ───────────────────────────────────────────────────────────


async def _status(req: web.Request) -> web.Response:
    """Full playback state."""
    html = await _mpchc_get(req.app["session"], "/variables.html")
    if html is None:
        return web.json_response({"error": "MPC-HC not reachable"}, status=503)
    v = _parse_variables(html)
    state_id = int(v.get("state", 0))
    return web.json_response(
        {
            "state": {0: "stopped", 1: "paused", 2: "playing"}.get(state_id, "unknown"),
            "state_id": state_id,
            "position_ms": int(v.get("position", 0)),
            "duration_ms": int(v.get("duration", 0)),
            "position_str": v.get("positionstring", ""),
            "duration_str": v.get("durationstring", ""),
            "volume": int(v.get("volumelevel", 0)),
            "muted": v.get("muted", "0") == "1",
            "playback_rate": float(v.get("playbackrate", "1.0")),
            "file": v.get("file", ""),
            "filepath": v.get("filepath", ""),
            "audio_track": v.get("audiotrack", ""),
            "subtitle_track": v.get("subtitletrack", ""),
        }
    )


async def _tracks(req: web.Request) -> web.Response:
    """List available audio and subtitle tracks with current selection."""
    html = await _mpchc_get(req.app["session"], "/controls.html")
    if html is None:
        return web.json_response({"error": "MPC-HC not reachable"}, status=503)
    return web.json_response(_parse_tracks(html))


async def _command(req: web.Request) -> web.Response:
    """Send a named command."""
    cmd = req.match_info["cmd"]
    if cmd not in CMD:
        valid = sorted(CMD.keys())
        return web.json_response({"error": f"Unknown command '{cmd}'", "valid": valid}, status=400)
    if not _post_wm_command(CMD[cmd]):
        return web.json_response({"error": "MPC-HC window not found"}, status=503)
    return web.json_response({"ok": True, "command": cmd, "wm_command": CMD[cmd]})


async def _seek(req: web.Request) -> web.Response:
    """Seek to absolute position. ?pos_ms=<milliseconds>"""
    try:
        pos_ms = int(req.rel_url.query["pos_ms"])
    except (KeyError, ValueError):
        return web.json_response({"error": "pos_ms required (milliseconds)"}, status=400)
    # MPC-HC /command.html accepts position in seconds
    result = await _mpchc_get(
        req.app["session"],
        "/command.html",
        {"wm_command": -1, "position": f"{pos_ms / 1000:.3f}"},
    )
    if result is None:
        return web.json_response({"error": "Seek failed — MPC-HC not reachable"}, status=503)
    return web.json_response({"ok": True, "position_ms": pos_ms})


async def _set_volume(req: web.Request) -> web.Response:
    """Set exact volume. ?level=<0-100>"""
    try:
        level = max(0, min(100, int(req.rel_url.query["level"])))
    except (KeyError, ValueError):
        return web.json_response({"error": "level required (0-100)"}, status=400)
    result = await _mpchc_get(
        req.app["session"],
        "/command.html",
        {"wm_command": -2, "volume": level},
    )
    if result is None:
        return web.json_response({"error": "Volume set failed"}, status=503)
    return web.json_response({"ok": True, "volume": level})


async def _audio_track(req: web.Request) -> web.Response:
    """Select audio track by index."""
    try:
        index = int(req.match_info["index"])
    except ValueError:
        return web.json_response({"error": "Invalid index"}, status=400)
    if not _post_wm_command(_AUDIO_BASE + index):
        return web.json_response({"error": "MPC-HC window not found"}, status=503)
    return web.json_response({"ok": True, "audio_track": index})


async def _subtitle_track(req: web.Request) -> web.Response:
    """Select subtitle track by index. Use index=-1 to toggle off."""
    try:
        index = int(req.match_info["index"])
    except ValueError:
        return web.json_response({"error": "Invalid index"}, status=400)
    cmd_id = CMD["sub_toggle"] if index < 0 else _SUB_BASE + index
    if not _post_wm_command(cmd_id):
        return web.json_response({"error": "MPC-HC window not found"}, status=503)
    return web.json_response({"ok": True, "subtitle_track": index})


async def _open_file(req: web.Request) -> web.Response:
    """Open a file in MPC-HC. ?path=<full filepath>"""
    path = req.rel_url.query.get("path", "").strip()
    if not path:
        return web.json_response({"error": "path required"}, status=400)
    exe = _find_mpchc_exe()
    if exe is None:
        return web.json_response({"error": "MPC-HC executable not found"}, status=503)
    try:
        subprocess.Popen([exe, path])
        return web.json_response({"ok": True, "path": path})
    except OSError as ex:
        return web.json_response({"error": str(ex)}, status=500)


async def _commands_list(req: web.Request) -> web.Response:  # noqa: ARG001
    """List all available named commands."""
    return web.json_response({"commands": sorted(CMD.keys())})


# ── App lifecycle ──────────────────────────────────────────────────────────────


async def _startup(app: web.Application) -> None:
    app["session"] = ClientSession(timeout=_TIMEOUT)


async def _cleanup(app: web.Application) -> None:
    await app["session"].close()


# ── App factory + main ─────────────────────────────────────────────────────────


def create_app() -> web.Application:
    """Create and return the aiohttp application (used by service wrapper too)."""
    app = web.Application()
    app.on_startup.append(_startup)
    app.on_cleanup.append(_cleanup)

    app.router.add_get("/status", _status)
    app.router.add_get("/tracks", _tracks)
    app.router.add_get("/commands", _commands_list)
    app.router.add_post("/command/{cmd}", _command)
    app.router.add_post("/seek", _seek)
    app.router.add_post("/volume", _set_volume)
    app.router.add_post("/audio/{index}", _audio_track)
    app.router.add_post("/subtitle/{index}", _subtitle_track)
    app.router.add_post("/open", _open_file)
    return app


def main() -> None:
    """Start the bridge server (standalone mode)."""
    print(f"MPC-HC Bridge  →  http://0.0.0.0:{BRIDGE_PORT}")
    print(f"MPC-HC target  →  http://localhost:{MPCHC_PORT}")
    web.run_app(create_app(), host="0.0.0.0", port=BRIDGE_PORT, print=lambda _: None)


if __name__ == "__main__":
    main()
