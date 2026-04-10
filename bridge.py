"""
MPC-HC Control Bridge
Runs on the Windows PC alongside MPC-HC.
Exposes a full HTTP API so the UC Remote driver can control everything.

Endpoints:
  GET  /status                   — playback state, position, volume, file
  GET  /tracks                   — available audio + subtitle tracks
  GET  /ws                       — WebSocket: pushed state changes (JSON diffs)
  POST /command/{cmd}            — named playback command (see CMD dict)
  POST /seek?pos_ms=<int>        — absolute seek in milliseconds
  POST /skip?offset_ms=<int>     — relative seek (±ms from current position)
  POST /volume?level=<0-100>     — set exact volume
  POST /audio/{index}            — select audio track by index
  POST /subtitle/{index}         — select subtitle track (-1 = disable)
  POST /open?path=<filepath>     — open file in MPC-HC
"""

import asyncio
import logging
import re
import subprocess
import sys
import time
from pathlib import Path

from aiohttp import ClientSession, ClientTimeout, web

_LOG = logging.getLogger(__name__)

# ── In-memory log ring buffer (for /debug/log endpoint) ───────────────────────
import collections

_LOG_BUFFER: collections.deque = collections.deque(maxlen=100)


class _BufHandler(logging.Handler):
    def emit(self, record: logging.LogRecord) -> None:
        _LOG_BUFFER.append(f"{self.formatter.formatTime(record, '%H:%M:%S')}  {record.getMessage()}")


_buf_handler = _BufHandler()
_buf_handler.setFormatter(logging.Formatter())
logging.getLogger().addHandler(_buf_handler)
logging.getLogger().setLevel(logging.WARNING)

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
    "frame_step": 891,
    "frame_step_back": 892,
    "speed_down": 894,
    "speed_up": 895,
    "speed_reset": 896,
    "seek_bwd_small": 899,  # ~5 s
    "seek_fwd_small": 900,
    "seek_bwd_large": 903,  # ~1 min
    "seek_fwd_large": 904,
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
    # Subtitles  (954 = Subtitle >, 955 = < Subtitle)
    "sub_next": 954,
    "sub_prev": 955,
    "sub_delay_plus": 958,
    "sub_delay_minus": 957,
    # Audio
    "audio_next": 952,
    "audio_prev": 953,
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
_CMD_TIMEOUT = ClientTimeout(total=1)  # commands don't need a long wait


# ── MPC-HC HTTP helpers ────────────────────────────────────────────────────────


async def _mpchc_get(
    session: ClientSession, path: str, params: dict | None = None
) -> str | None:
    try:
        url = f"http://127.0.0.1:{MPCHC_PORT}{path}"
        async with session.get(url, params=params, allow_redirects=False) as r:
            if r.status in (200, 302):
                return await r.text()
    except Exception:  # pylint: disable=broad-exception-caught
        pass
    return None


async def _mpchc_command(session: ClientSession, params: dict) -> bool:
    """Fire a command at MPC-HC and return immediately — don't read body."""
    t0 = time.monotonic()
    try:
        url = f"http://127.0.0.1:{MPCHC_PORT}/command.html"
        async with session.get(url, params=params, allow_redirects=False, timeout=_CMD_TIMEOUT) as r:
            ok = r.status in (200, 302)
            _LOG.warning("_mpchc_command %s → %s in %.3fs", params, r.status, time.monotonic() - t0)
            return ok
    except Exception as ex:  # pylint: disable=broad-exception-caught
        _LOG.warning("_mpchc_command %s failed in %.3fs: %s", params, time.monotonic() - t0, ex)
        return False


def _parse_variables(html: str) -> dict[str, str]:
    """Parse /variables.html <p id="key">value</p> tags."""
    return {
        m.group(1): m.group(2).strip()
        for m in re.finditer(r'<p\s+id="([^"]+)">([^<]*)</p>', html)
    }


def _parse_tracks(html: str) -> dict[str, list]:
    """Parse audio/subtitle track lists from /controls.html select elements (original MPC-HC only)."""
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


# ── MKV EBML track parser (pure Python, no external deps) ─────────────────────

def _ebml_id(buf: bytes, pos: int) -> tuple[int, int]:
    b = buf[pos]
    if b >= 0x80: return b, pos + 1
    if b >= 0x40: return (b << 8) | buf[pos + 1], pos + 2
    if b >= 0x20: return (b << 16) | (buf[pos + 1] << 8) | buf[pos + 2], pos + 3
    if b >= 0x10: return (b << 24) | (buf[pos + 1] << 16) | (buf[pos + 2] << 8) | buf[pos + 3], pos + 4
    raise ValueError(f"bad EBML ID {b:#x} at {pos}")


def _ebml_size(buf: bytes, pos: int) -> tuple[int, int]:
    b = buf[pos]
    if b >= 0x80: return b & 0x7F, pos + 1
    if b >= 0x40: return ((b & 0x3F) << 8) | buf[pos + 1], pos + 2
    if b >= 0x20: return ((b & 0x1F) << 16) | (buf[pos + 1] << 8) | buf[pos + 2], pos + 3
    if b >= 0x10: return ((b & 0x0F) << 24) | (buf[pos + 1] << 16) | (buf[pos + 2] << 8) | buf[pos + 3], pos + 4
    if b >= 0x08:
        v = (b & 0x07)
        for i in range(4): v = (v << 8) | buf[pos + 1 + i]
        return v, pos + 5
    if b == 0x01:
        v = 0
        for i in range(7): v = (v << 8) | buf[pos + 1 + i]
        return v, pos + 8
    return -1, pos + 1  # unknown / all-ones = infinite


def _parse_mkv_chapters(buf: bytes, pos: int, el_end: int, n: int) -> list[dict]:
    """Parse a Chapters element body from buf[pos:el_end]. Returns list of
    {"name": str, "time_ms": int} sorted by time, hidden chapters excluded."""
    _ID_EDITION            = 0x45B9
    _ID_CHAPTER_ATOM       = 0xB6
    _ID_CHAPTER_TIME_START = 0x91  # nanoseconds as big-endian uint
    _ID_CHAPTER_FLAG_HID   = 0x98
    _ID_CHAPTER_DISPLAY    = 0x80
    _ID_CHAP_STRING        = 0x85

    chapters: list[dict] = []
    while pos < min(el_end, n) - 2:
        try:
            eid, npos = _ebml_id(buf, pos)
            esz, npos = _ebml_size(buf, npos)
        except (IndexError, ValueError):
            break
        if esz < 0: esz = 0
        edition_end = npos + esz
        if eid == _ID_EDITION:
            apos = npos
            while apos < min(edition_end, n) - 2:
                try:
                    aid, apos2 = _ebml_id(buf, apos)
                    asz, apos2 = _ebml_size(buf, apos2)
                except (IndexError, ValueError):
                    break
                if asz < 0: asz = 0
                atom_end = apos2 + asz
                if aid == _ID_CHAPTER_ATOM:
                    time_ns, hidden, names = 0, False, []
                    fpos = apos2
                    while fpos < min(atom_end, n) - 2:
                        try:
                            fid, fpos2 = _ebml_id(buf, fpos)
                            fsz, fpos2 = _ebml_size(buf, fpos2)
                        except (IndexError, ValueError):
                            break
                        if fsz < 0: fsz = 0
                        fdata = buf[fpos2: fpos2 + fsz]
                        if fid == _ID_CHAPTER_TIME_START:
                            time_ns = int.from_bytes(fdata, "big")
                        elif fid == _ID_CHAPTER_FLAG_HID:
                            hidden = bool(int.from_bytes(fdata, "big"))
                        elif fid == _ID_CHAPTER_DISPLAY:
                            dpos = fpos2
                            while dpos < min(fpos2 + fsz, n) - 2:
                                try:
                                    did, dpos2 = _ebml_id(buf, dpos)
                                    dsz, dpos2 = _ebml_size(buf, dpos2)
                                except (IndexError, ValueError):
                                    break
                                if dsz < 0: dsz = 0
                                if did == _ID_CHAP_STRING:
                                    names.append(buf[dpos2: dpos2 + dsz].decode("utf-8", errors="replace"))
                                dpos = dpos2 + dsz
                        fpos = fpos2 + fsz
                    if not hidden and names:
                        chapters.append({"name": names[0], "time_ms": time_ns // 1_000_000})
                apos = atom_end
        pos = edition_end
    chapters.sort(key=lambda c: c["time_ms"])
    return chapters


def _read_mkv_tracks(filepath: str) -> dict[str, list] | None:
    """Parse MKV EBML for tracks, video info, and chapters.

    Returns:
      {"audio":    [{"pos": N, "name": "...", "lang": "ger", "codec": "A_AC3"}, …],
       "subtitle": […],
       "video":    [{"pos": 0, "codec": "V_MPEGH/ISO/HEVC", "width": 1920, "height": 1080,
                     "label": "1920x1080 H.265"}],
       "chapters": [{"name": "Chapter 1", "time_ms": 0}, …]}
    Returns None if the file cannot be read or is not MKV.
    """
    _ID_EBML     = 0x1A45DFA3
    _ID_SEGMENT  = 0x18538067
    _ID_SEEKHEAD = 0x114D9B74
    _ID_SEEK     = 0x4DBB
    _ID_SEEKID   = 0x53AB
    _ID_SEEKPOS  = 0x53AC
    _ID_TRACKS   = 0x1654AE6B
    _ID_CHAPTERS = 0x1043A770
    _ID_CLUSTER  = 0x1F43B675
    _ID_ENTRY    = 0xAE
    _ID_NUM      = 0xD7
    _ID_TYPE     = 0x83
    _ID_NAME     = 0x536E
    _ID_LANG     = 0x22B59C
    _ID_CODEC    = 0x86
    _ID_VIDEO    = 0xE0
    _ID_PX_W     = 0xB0
    _ID_PX_H     = 0xBA
    _TYPE_VIDEO  = 1
    _TYPE_AUDIO  = 2
    _TYPE_SUB    = 17

    _CODEC_SHORT = {
        "V_MPEGH/ISO/HEVC": "H.265", "V_MPEG4/ISO/AVC": "H.264",
        "V_AV1": "AV1", "V_VP9": "VP9", "V_VP8": "VP8",
        "V_MPEG2": "MPEG-2", "V_MPEG1": "MPEG-1",
    }

    try:
        with open(filepath, "rb") as fh:
            buf = fh.read(524288)  # 512 KB
        n = len(buf)
        if n < 8 or buf[:4] != b"\x1a\x45\xdf\xa3":
            return None  # not MKV

        pos = 0
        _, pos = _ebml_id(buf, pos)
        esz, pos = _ebml_size(buf, pos)
        pos += esz  # skip EBML header

        eid, pos = _ebml_id(buf, pos)
        if eid != _ID_SEGMENT:
            return None
        esz, pos = _ebml_size(buf, pos)
        seg_data_start = pos  # file offset where Segment body starts
        seg_end = (pos + esz) if esz >= 0 else n

        result: dict[str, list] = {"audio": [], "subtitle": [], "video": [], "chapters": []}
        chapters_offset = -1  # SeekPosition relative to seg_data_start

        # Scan all pre-Cluster Segment children
        while pos < min(seg_end, n) - 4:
            try:
                eid, npos = _ebml_id(buf, pos)
                esz, npos = _ebml_size(buf, npos)
            except (IndexError, ValueError):
                break
            if esz < 0 or esz > n:
                esz = n - npos
            el_end = npos + esz

            if eid == _ID_CLUSTER:
                break

            elif eid == _ID_SEEKHEAD:
                # Extract chapters file position from SeekHead
                spos = npos
                while spos < min(el_end, n) - 2:
                    try:
                        sid, spos2 = _ebml_id(buf, spos)
                        ssz, spos2 = _ebml_size(buf, spos2)
                    except (IndexError, ValueError):
                        break
                    if ssz < 0: ssz = 0
                    se_end = spos2 + ssz
                    if sid == _ID_SEEK:
                        sk_id, sk_off = 0, -1
                        epos = spos2
                        while epos < min(se_end, n) - 2:
                            try:
                                fid, epos2 = _ebml_id(buf, epos)
                                fsz, epos2 = _ebml_size(buf, epos2)
                            except (IndexError, ValueError):
                                break
                            if fsz < 0: fsz = 0
                            fdata = buf[epos2: epos2 + fsz]
                            if fid == _ID_SEEKID:
                                sk_id = int.from_bytes(fdata, "big")
                            elif fid == _ID_SEEKPOS:
                                sk_off = int.from_bytes(fdata, "big")
                            epos = epos2 + fsz
                        if sk_id == _ID_CHAPTERS and sk_off >= 0:
                            chapters_offset = sk_off
                    spos = se_end

            elif eid == _ID_CHAPTERS:
                result["chapters"] = _parse_mkv_chapters(buf, npos, el_end, n)

            elif eid == _ID_TRACKS:
                tpos = npos
                while tpos < min(el_end, n) - 2:
                    try:
                        tid, tpos2 = _ebml_id(buf, tpos)
                        tsz, tpos2 = _ebml_size(buf, tpos2)
                    except (IndexError, ValueError):
                        break
                    if tsz < 0: tsz = 0
                    te_end = tpos2 + tsz
                    if tid == _ID_ENTRY:
                        track: dict = {"number": 0, "type": 0, "name": "", "lang": "", "codec": "",
                                       "width": 0, "height": 0}
                        epos = tpos2
                        while epos < min(te_end, n) - 2:
                            try:
                                fid, epos2 = _ebml_id(buf, epos)
                                fsz, epos2 = _ebml_size(buf, epos2)
                            except (IndexError, ValueError):
                                break
                            if fsz < 0: fsz = 0
                            fdata = buf[epos2: epos2 + fsz]
                            if fid == _ID_NUM:    track["number"] = int.from_bytes(fdata, "big")
                            elif fid == _ID_TYPE: track["type"]   = int.from_bytes(fdata, "big")
                            elif fid == _ID_NAME: track["name"]   = fdata.decode("utf-8", errors="replace")
                            elif fid == _ID_LANG: track["lang"]   = fdata.decode("ascii", errors="replace").rstrip("\x00")
                            elif fid == _ID_CODEC: track["codec"] = fdata.decode("ascii", errors="replace").rstrip("\x00")
                            elif fid == _ID_VIDEO:
                                vpos = epos2
                                while vpos < min(epos2 + fsz, n) - 2:
                                    try:
                                        vid, vpos2 = _ebml_id(buf, vpos)
                                        vsz, vpos2 = _ebml_size(buf, vpos2)
                                    except (IndexError, ValueError):
                                        break
                                    if vsz < 0: vsz = 0
                                    vdata = buf[vpos2: vpos2 + vsz]
                                    if vid == _ID_PX_W: track["width"]  = int.from_bytes(vdata, "big")
                                    elif vid == _ID_PX_H: track["height"] = int.from_bytes(vdata, "big")
                                    vpos = vpos2 + vsz
                            epos = epos2 + fsz
                        if track["type"] == _TYPE_VIDEO:
                            short = _CODEC_SHORT.get(track["codec"], track["codec"].split("/")[-1])
                            track["pos"] = len(result["video"])
                            track["label"] = f"{track['width']}x{track['height']} {short}"
                            result["video"].append(track)
                        elif track["type"] == _TYPE_AUDIO:
                            track["pos"] = len(result["audio"])
                            result["audio"].append(track)
                        elif track["type"] == _TYPE_SUB:
                            track["pos"] = len(result["subtitle"])
                            result["subtitle"].append(track)
                    tpos = te_end

            pos = npos + esz

        # Chapters via SeekHead: may lie outside the initial 512 KB window
        if not result["chapters"] and chapters_offset >= 0:
            chap_abs = seg_data_start + chapters_offset
            try:
                with open(filepath, "rb") as fh:
                    fh.seek(chap_abs)
                    hdr = fh.read(12)
                cpos = 0
                c_eid, cpos = _ebml_id(hdr, cpos)
                c_esz, cpos = _ebml_size(hdr, cpos)
                if c_eid == _ID_CHAPTERS and c_esz > 0:
                    read_sz = min(c_esz, 262144)  # cap at 256 KB
                    with open(filepath, "rb") as fh:
                        fh.seek(chap_abs + cpos)
                        cbuf = fh.read(read_sz)
                    result["chapters"] = _parse_mkv_chapters(cbuf, 0, len(cbuf), len(cbuf))
            except Exception as ex:  # pylint: disable=broad-exception-caught
                _LOG.warning("chapters read at %d failed: %s", chap_abs, ex)

        return result

    except Exception as ex:  # pylint: disable=broad-exception-caught
        _LOG.warning("_read_mkv_tracks %s: %s", filepath, ex)
    return None


def _match_track(mkv_tracks: list[dict], current_name: str) -> int:
    """
    Match MPC-HC's current track string (e.g. 'A: German AC3…[ger]') against
    MKV track list. Returns 0-based pos, or -1 if not found.
    """
    if not current_name or not mkv_tracks:
        return 0  # assume first track
    cur = current_name.lower()
    # Extract 3-letter language code from brackets e.g. "[ger]"
    lang_m = re.search(r'\[([a-z]{3})\]', cur)
    cur_lang = lang_m.group(1) if lang_m else ""
    # Extract codec hint: ac3, dts, aac, flac, truehd, eac3, pcm, mp3, vorbis, opus, vobsub, ass, srt, subrip, pgs
    codec_hints = {
        "ac3": "A_AC3", "dts": "A_DTS", "aac": "A_AAC", "flac": "A_FLAC",
        "truehd": "A_TRUEHD", "eac3": "A_EAC3", "mp3": "A_MPEG", "opus": "A_OPUS",
        "vorbis": "A_VORBIS", "vobsub": "S_VOBSUB", "ass": "S_TEXT/ASS",
        "subrip": "S_TEXT/UTF8", "pgs": "S_HDMV/PGS",
    }
    cur_codec = ""
    for hint, codec in codec_hints.items():
        if hint in cur:
            cur_codec = codec
            break

    best_pos, best_score = 0, -1
    for t in mkv_tracks:
        score = 0
        if cur_lang and t.get("lang", "").lower() == cur_lang:
            score += 10
        if cur_codec and t.get("codec", "").upper().startswith(cur_codec.upper().split("/")[0]):
            score += 5
        if t.get("name") and t["name"].lower() in cur:
            score += 3
        if score > best_score:
            best_score, best_pos = score, t["pos"]
    return best_pos


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
    """List available audio and subtitle tracks, parsed from the open MKV file."""
    html = await _mpchc_get(req.app["session"], "/variables.html")
    if html is None:
        return web.json_response({"error": "MPC-HC not reachable"}, status=503)
    v = _parse_variables(html)
    filepath = v.get("filepath", "")
    cur_audio = v.get("audiotrack", "")
    cur_sub = v.get("subtitletrack", "")

    if not filepath:
        return web.json_response({"error": "No file loaded in MPC-HC"}, status=404)

    mkv = _read_mkv_tracks(filepath)
    if mkv is None:
        return web.json_response({"error": f"Cannot read track info from: {filepath}"}, status=422)

    cur_audio_pos = _match_track(mkv["audio"], cur_audio)
    cur_sub_pos = _match_track(mkv["subtitle"], cur_sub)

    for t in mkv["audio"]:
        t["selected"] = t["pos"] == cur_audio_pos
        t["label"] = f"{t['lang'].upper() or '?'}  {t['codec']}  {t['name']}".strip()
    for t in mkv["subtitle"]:
        t["selected"] = t["pos"] == cur_sub_pos
        t["label"] = f"{t['lang'].upper() or '?'}  {t['codec']}  {t['name']}".strip()

    return web.json_response({
        "audio": mkv["audio"],
        "subtitle": mkv["subtitle"],
        "video": mkv.get("video", []),
        "chapters": mkv.get("chapters", []),
        "current_audio_pos": cur_audio_pos,
        "current_sub_pos": cur_sub_pos,
        "filepath": filepath,
    })


async def _select_track(req: web.Request) -> web.Response:
    """Cycle to a specific track by position. POST /audio/select/{pos} or /subtitle/select/{pos}"""
    kind = req.match_info["kind"]   # "audio" or "subtitle"
    try:
        target_pos = int(req.match_info["pos"])
    except ValueError:
        return web.json_response({"error": "Invalid pos"}, status=400)

    # Get current state
    html = await _mpchc_get(req.app["session"], "/variables.html")
    if html is None:
        return web.json_response({"error": "MPC-HC not reachable"}, status=503)
    v = _parse_variables(html)
    filepath = v.get("filepath", "")
    if not filepath:
        return web.json_response({"error": "No file loaded"}, status=404)

    mkv = _read_mkv_tracks(filepath)
    if mkv is None:
        return web.json_response({"error": "Cannot read track info"}, status=422)

    tracks = mkv[kind]
    if not tracks:
        return web.json_response({"error": f"No {kind} tracks found"}, status=404)
    if target_pos < 0 or target_pos >= len(tracks):
        return web.json_response({"error": f"pos out of range 0..{len(tracks)-1}"}, status=400)

    cur_name = v.get("audiotrack" if kind == "audio" else "subtitletrack", "")
    cur_pos = _match_track(tracks, cur_name)

    n = len(tracks)
    steps_fwd = (target_pos - cur_pos) % n
    steps_bwd = (cur_pos - target_pos) % n

    if steps_fwd == 0:
        return web.json_response({"ok": True, "steps": 0, "direction": "none"})

    cmd_next = "audio_next" if kind == "audio" else "sub_next"
    cmd_prev = "audio_prev" if kind == "audio" else "sub_prev"

    if steps_fwd <= steps_bwd:
        cmd, steps = cmd_next, steps_fwd
    else:
        cmd, steps = cmd_prev, steps_bwd

    cmd_id = CMD[cmd]
    for _ in range(steps):
        if not _post_wm_command(cmd_id):
            await _mpchc_command(req.app["session"], {"wm_command": cmd_id})
        await asyncio.sleep(0.05)  # small gap between steps

    _LOG.warning("_select_track %s: cur=%d target=%d cmd=%s steps=%d", kind, cur_pos, target_pos, cmd, steps)
    return web.json_response({"ok": True, "steps": steps, "direction": cmd})


async def _command(req: web.Request) -> web.Response:
    """Send a named command."""
    t0 = time.monotonic()
    cmd = req.match_info["cmd"]
    if cmd not in CMD:
        valid = sorted(CMD.keys())
        return web.json_response({"error": f"Unknown command '{cmd}'", "valid": valid}, status=400)
    cmd_id = CMD[cmd]
    wm_ok = _post_wm_command(cmd_id)
    _LOG.warning("_command '%s': PostMessageW=%s in %.3fs", cmd, wm_ok, time.monotonic() - t0)
    if not wm_ok:
        # Service runs in Session 0 — fall back to MPC-HC HTTP (works across sessions)
        if not await _mpchc_command(req.app["session"], {"wm_command": cmd_id}):
            return web.json_response({"error": "MPC-HC not reachable"}, status=503)
    _LOG.warning("_command '%s' total: %.3fs", cmd, time.monotonic() - t0)
    return web.json_response({"ok": True, "command": cmd, "wm_command": cmd_id})


def _ms_to_hmsms(ms: int) -> str:
    """Convert milliseconds to MPC-HC position string HH:MM:SS:mmm.

    MPC-HC /command.html?wm_command=-1 parses the position field with
    _stscanf_s("%d%c%d%c%d%c%d") — four integers separated by any char.
    Sending decimal seconds (e.g. "46.245") is NOT parsed correctly;
    the colon-separated form "HH:MM:SS:mmm" always works.
    """
    ms = max(0, ms)
    hours, rem = divmod(ms, 3_600_000)
    minutes, rem = divmod(rem, 60_000)
    seconds, millis = divmod(rem, 1_000)
    return f"{hours}:{minutes:02d}:{seconds:02d}:{millis:03d}"


async def _seek(req: web.Request) -> web.Response:
    """Seek to absolute position. ?pos_ms=<milliseconds>"""
    try:
        pos_ms = int(req.rel_url.query["pos_ms"])
    except (KeyError, ValueError):
        return web.json_response({"error": "pos_ms required (milliseconds)"}, status=400)
    position = _ms_to_hmsms(pos_ms)
    _LOG.warning("_seek: pos_ms=%d  position=%s", pos_ms, position)
    if not await _mpchc_command(req.app["session"], {"wm_command": -1, "position": position}):
        return web.json_response({"error": "Seek failed — MPC-HC not reachable"}, status=503)
    return web.json_response({"ok": True, "position_ms": pos_ms})


async def _skip(req: web.Request) -> web.Response:
    """Relative seek by ±offset_ms milliseconds. ?offset_ms=<int>"""
    try:
        offset_ms = int(req.rel_url.query["offset_ms"])
    except (KeyError, ValueError):
        return web.json_response({"error": "offset_ms required (milliseconds, may be negative)"}, status=400)
    html = await _mpchc_get(req.app["session"], "/variables.html")
    if html is None:
        return web.json_response({"error": "MPC-HC not reachable"}, status=503)
    v = _parse_variables(html)
    pos_ms = int(v.get("position", 0))
    dur_ms = int(v.get("duration", 0))
    target_ms = max(0, min(dur_ms if dur_ms > 0 else pos_ms, pos_ms + offset_ms))
    position = _ms_to_hmsms(target_ms)
    _LOG.warning("_skip: offset_ms=%d  pos_ms=%d → target_ms=%d  position=%s", offset_ms, pos_ms, target_ms, position)
    if not await _mpchc_command(req.app["session"], {"wm_command": -1, "position": position}):
        return web.json_response({"error": "Seek failed — MPC-HC not reachable"}, status=503)
    return web.json_response({"ok": True, "position_ms": target_ms, "offset_ms": offset_ms})


async def _set_volume(req: web.Request) -> web.Response:
    """Set exact volume. ?level=<0-100>"""
    try:
        level = max(0, min(100, int(req.rel_url.query["level"])))
    except (KeyError, ValueError):
        return web.json_response({"error": "level required (0-100)"}, status=400)
    if not await _mpchc_command(req.app["session"], {"wm_command": -2, "volume": level}):
        return web.json_response({"error": "Volume set failed"}, status=503)
    return web.json_response({"ok": True, "volume": level})


async def _audio_track(req: web.Request) -> web.Response:
    """Select audio track by index."""
    try:
        index = int(req.match_info["index"])
    except ValueError:
        return web.json_response({"error": "Invalid index"}, status=400)
    cmd_id = _AUDIO_BASE + index
    if not _post_wm_command(cmd_id):
        if not await _mpchc_command(req.app["session"], {"wm_command": cmd_id}):
            return web.json_response({"error": "MPC-HC not reachable"}, status=503)
    return web.json_response({"ok": True, "audio_track": index})


async def _subtitle_track(req: web.Request) -> web.Response:
    """Select subtitle track by index. Use index=-1 to toggle off."""
    try:
        index = int(req.match_info["index"])
    except ValueError:
        return web.json_response({"error": "Invalid index"}, status=400)
    if index < 0:
        return web.json_response({"error": "Use subtitle cycling commands to toggle"}, status=400)
    cmd_id = _SUB_BASE + index
    if not _post_wm_command(cmd_id):
        if not await _mpchc_command(req.app["session"], {"wm_command": cmd_id}):
            return web.json_response({"error": "MPC-HC not reachable"}, status=503)
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


async def _debug_log(req: web.Request) -> web.Response:  # noqa: ARG001
    """Return last 100 log entries as JSON."""
    return web.json_response({"log": list(_LOG_BUFFER)})


async def _debug_controls(req: web.Request) -> web.Response:
    """Return raw /controls.html from MPC-HC for debugging."""
    html = await _mpchc_get(req.app["session"], "/controls.html")
    if html is None:
        return web.json_response({"error": "MPC-HC not reachable"}, status=503)
    return web.Response(text=html, content_type="text/html")


async def _debug_variables(req: web.Request) -> web.Response:
    """Return raw /variables.html from MPC-HC for debugging."""
    html = await _mpchc_get(req.app["session"], "/variables.html")
    if html is None:
        return web.json_response({"error": "MPC-HC not reachable"}, status=503)
    return web.Response(text=html, content_type="text/html")


async def _index(req: web.Request) -> web.Response:  # noqa: ARG001
    """Root endpoint — redirect to /status."""
    raise web.HTTPFound("/status")


# ── WebSocket push ─────────────────────────────────────────────────────────────

_ws_clients: set[web.WebSocketResponse] = set()
_PUSH_INTERVAL_PLAYING = 1.0   # seconds — while MPC-HC is playing
_PUSH_INTERVAL_IDLE = 3.0       # seconds — paused / stopped / unreachable


async def _broadcast(data: dict) -> None:
    """Send a JSON diff to all connected WebSocket clients."""
    for ws in list(_ws_clients):
        try:
            await ws.send_json(data)
        except Exception:  # pylint: disable=broad-exception-caught
            _ws_clients.discard(ws)


async def _ws_handler(req: web.Request) -> web.WebSocketResponse:
    """WebSocket endpoint — push-only stream of MPC-HC state changes."""
    ws = web.WebSocketResponse(heartbeat=30)
    await ws.prepare(req)
    _ws_clients.add(ws)
    try:
        async for _ in ws:
            pass  # client-to-server messages are ignored
    finally:
        _ws_clients.discard(ws)
    return ws


async def _push_task(app: web.Application) -> None:
    """Poll MPC-HC locally and broadcast changed fields to WebSocket clients."""
    prev: dict = {}
    while True:
        interval = _PUSH_INTERVAL_IDLE
        try:
            html = await _mpchc_get(app["session"], "/variables.html")
            if html is not None:
                v = _parse_variables(html)
                state_id = int(v.get("state", 0))
                current = {
                    "state_id": state_id,
                    "state": {0: "stopped", 1: "paused", 2: "playing"}.get(state_id, "unknown"),
                    "position_ms": int(v.get("position", 0)),
                    "duration_ms": int(v.get("duration", 0)),
                    "volume": int(v.get("volumelevel", 0)),
                    "muted": v.get("muted", "0") == "1",
                    "audio_track": v.get("audiotrack", ""),
                    "subtitle_track": v.get("subtitletrack", ""),
                }
                changed = {k: val for k, val in current.items() if prev.get(k) != val}
                if changed:
                    if _ws_clients:
                        await _broadcast(changed)
                    prev.update(changed)
                interval = _PUSH_INTERVAL_PLAYING if state_id == 2 else _PUSH_INTERVAL_IDLE
        except Exception:  # pylint: disable=broad-exception-caught
            pass
        await asyncio.sleep(interval)


# ── App lifecycle ──────────────────────────────────────────────────────────────


async def _startup(app: web.Application) -> None:
    app["session"] = ClientSession(timeout=_TIMEOUT)
    app["push_task"] = asyncio.create_task(_push_task(app))


async def _cleanup(app: web.Application) -> None:
    task = app.get("push_task")
    if task:
        task.cancel()
        try:
            await task
        except asyncio.CancelledError:
            pass
    await app["session"].close()


# ── App factory + main ─────────────────────────────────────────────────────────


def create_app() -> web.Application:
    """Create and return the aiohttp application (used by service wrapper too)."""
    app = web.Application()
    app.on_startup.append(_startup)
    app.on_cleanup.append(_cleanup)

    app.router.add_get("/", _index)
    app.router.add_get("/status", _status)
    app.router.add_get("/tracks", _tracks)
    app.router.add_post("/{kind:(audio|subtitle)}/select/{pos}", _select_track)
    app.router.add_get("/commands", _commands_list)
    app.router.add_get("/debug/log", _debug_log)
    app.router.add_get("/debug/controls", _debug_controls)
    app.router.add_get("/debug/variables", _debug_variables)
    app.router.add_get("/ws", _ws_handler)
    app.router.add_post("/command/{cmd}", _command)
    app.router.add_post("/seek", _seek)
    app.router.add_post("/skip", _skip)
    app.router.add_post("/volume", _set_volume)
    app.router.add_post("/audio/{index}", _audio_track)
    app.router.add_post("/subtitle/{index}", _subtitle_track)
    app.router.add_post("/open", _open_file)
    return app


def main() -> None:
    """Start the bridge server (standalone mode)."""
    logging.basicConfig(level=logging.WARNING, format="%(asctime)s %(levelname)s %(message)s")
    print(f"MPC-HC Bridge  →  http://0.0.0.0:{BRIDGE_PORT}")
    print(f"MPC-HC target  →  http://127.0.0.1:{MPCHC_PORT}")
    web.run_app(create_app(), host="0.0.0.0", port=BRIDGE_PORT, print=lambda _: None)


if __name__ == "__main__":
    main()
