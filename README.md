# MPC-HC Bridge

A lightweight Windows HTTP/WebSocket bridge that exposes full control of [MPC-HC](https://github.com/clsid2/mpc-hc) (Media Player Classic - Home Cinema) over the local network.

Designed as a companion to the [intg-mpchc-kodi](https://github.com/Zendonir/intg-mpchc-kodi) UC Remote 3 integration, but works with any HTTP client.

---

## Features

- **Full playback control** — play/pause, stop, seek, speed, chapters, next/prev
- **Volume control** — set exact level, mute toggle, up/down
- **Subtitle control** — cycle tracks, toggle on/off, adjust delay
- **Audio track control** — cycle tracks, adjust delay
- **Real-time state via WebSocket** — push-only stream of changed fields, no polling required
- **REST status endpoint** — position, duration, state, volume, current audio/subtitle track name
- **Open files** — launch MPC-HC with a specific file path
- **Windows Service** — auto-starts with Windows, runs in the background
- **GUI installer** — install, uninstall, start and stop the service without a terminal

---

## Download

Get the latest `mpchc-bridge-vX.X.X-windows-x64.exe` from the [Releases](../../releases) page.

The `.exe` is a single self-contained file — no Python installation required.

---

## Quick Start

1. **Enable MPC-HC web interface**
   In MPC-HC: *View → Options → Player → Web Interface*
   - Check **"Listen on port"** — default `13579`
   - Optionally check all sub-options for full state reporting

2. **Run the bridge**
   Double-click `mpchc-bridge.exe` to open the GUI.
   Click **Install Service** — the bridge installs itself as a Windows service that starts automatically with Windows.

3. **Verify**
   Click **Test in Browser** — opens `http://localhost:13580/status` in your browser.
   You should see a JSON response with the current playback state.

---

## Ports

| Port | Direction | Description |
|------|-----------|-------------|
| `13579` | Bridge → MPC-HC | MPC-HC built-in web interface (localhost only) |
| `13580` | Clients → Bridge | Bridge HTTP + WebSocket API (all interfaces) |

---

## HTTP API

All endpoints are on `http://<host>:13580`.

### GET /status

Returns the full current playback state.

```json
{
  "state": "playing",
  "state_id": 2,
  "position_ms": 123456,
  "duration_ms": 7984392,
  "position_str": "00:02:03",
  "duration_str": "02:13:04",
  "volume": 75,
  "muted": false,
  "playback_rate": 1.0,
  "file": "The Running Man (2025).mkv",
  "filepath": "Y:\\Filme\\The Running Man (2025).mkv",
  "audio_track": "Deutsch (AC3 5.1)",
  "subtitle_track": "Deutsch"
}
```

`state_id`: `0` = stopped, `1` = paused, `2` = playing

---

### GET /ws — WebSocket push stream

Connect to `ws://<host>:13580/ws` to receive real-time state updates.

The bridge polls MPC-HC locally and pushes **only changed fields** as JSON objects:

```json
{"position_ms": 124000}
{"state_id": 1, "state": "paused"}
{"audio_track": "English (DTS-HD MA 7.1)"}
{"volume": 80, "muted": false}
```

**Poll intervals:**
- Playing: every **1 second**
- Paused / stopped: every **3 seconds**
- A heartbeat ping is sent every 30 seconds to keep the connection alive

This is far more efficient than HTTP polling from a remote device — the client only maintains an idle TCP socket.

---

### GET /tracks

Returns available audio and subtitle tracks (if exposed by MPC-HC's web interface).

```json
{
  "audio": [
    {"index": 0, "name": "Deutsch (AC3 5.1)", "selected": true},
    {"index": 1, "name": "English (DTS-HD MA 7.1)", "selected": false}
  ],
  "subtitle": [
    {"index": 0, "name": "Deutsch", "selected": true},
    {"index": 1, "name": "English", "selected": false}
  ]
}
```

---

### GET /commands

Returns a list of all valid command names for `POST /command/{cmd}`.

---

### POST /command/{cmd}

Sends a named command to MPC-HC via Windows `WM_COMMAND`.

```
POST /command/play_pause
POST /command/sub_next
POST /command/audio_next
```

**Response:**
```json
{"ok": true, "command": "play_pause", "wm_command": 889}
```

**Available commands:**

| Category | Commands |
|----------|----------|
| Playback | `play_pause`, `stop`, `goto_start`, `frame_step`, `frame_step_back` |
| Speed | `speed_up`, `speed_down`, `speed_reset` |
| Seek | `seek_fwd_small` (~5s), `seek_bwd_small`, `seek_fwd_large` (~1min), `seek_bwd_large` |
| Navigation | `prev`, `next`, `chapter_next`, `chapter_prev` |
| Volume | `vol_up`, `vol_down`, `mute` |
| Subtitles | `sub_toggle`, `sub_next`, `sub_prev`, `sub_delay_plus`, `sub_delay_minus` |
| Audio | `audio_next`, `audio_delay_plus`, `audio_delay_minus` |
| View | `fullscreen`, `zoom_25`, `zoom_50`, `zoom_100`, `zoom_200`, `zoom_fit` |
| File | `close` |
| DVD | `dvd_menu_title`, `dvd_menu_root`, `dvd_menu_sub`, `dvd_menu_audio`, `dvd_menu_angle`, `dvd_menu_chapter` |

---

### POST /seek?pos_ms=\<int\>

Seeks to an absolute position in milliseconds.

```
POST /seek?pos_ms=120000
```

```json
{"ok": true, "position_ms": 120000}
```

---

### POST /volume?level=\<0-100\>

Sets the volume to an exact level.

```
POST /volume?level=75
```

```json
{"ok": true, "volume": 75}
```

---

### POST /audio/{index}

Selects an audio track by index (as returned by `/tracks`).

```
POST /audio/1
```

---

### POST /subtitle/{index}

Selects a subtitle track by index. Use `-1` to disable subtitles.

```
POST /subtitle/0
POST /subtitle/-1
```

---

### POST /open?path=\<filepath\>

Opens a file in MPC-HC. MPC-HC is launched if not already running.

```
POST /open?path=Y%3A%5CFilme%5CMovie.mkv
```

The path must be URL-encoded. MPC-HC is located automatically via the Windows registry or common install paths.

---

## Windows Service

The bridge installs as a Windows service named **MpcHcBridge** that:
- Starts automatically with Windows (before login)
- Runs silently in the background
- Is managed via the GUI or standard Windows service tools

### GUI

Double-click the `.exe` (no arguments) to open the control panel:

- **Install Service** — registers the service and starts it immediately (requires Administrator)
- **Uninstall** — stops and removes the service
- **Start / Stop** — manual service control
- **Test in Browser** — opens `/status` in the default browser

The status indicator updates every 3 seconds.

### Command line

Run as Administrator:

```
mpchc-bridge.exe install    # install + auto-start
mpchc-bridge.exe start      # start the service
mpchc-bridge.exe stop       # stop the service
mpchc-bridge.exe restart    # restart the service
mpchc-bridge.exe remove     # uninstall the service
mpchc-bridge.exe status     # show current status
mpchc-bridge.exe debug      # run in foreground (Ctrl+C to stop)
```

---

## Use with UC Remote 3 / intg-mpchc-kodi

The bridge was built for the [intg-mpchc-kodi](https://github.com/Zendonir/intg-mpchc-kodi) integration driver, which combines Kodi metadata with MPC-HC playback when using MPC-HC as an external player.

**Setup in the UC Remote integration:**

| Field | Value |
|-------|-------|
| MPC-HC host | IP address of the Windows PC |
| MPC-HC port | `13579` (direct MPC-HC, for HTTP fallback) |
| Bridge port | `13580` (enables WebSocket push) |

When the bridge port is set, the remote connects via WebSocket and receives pushed state updates — no periodic polling. This significantly reduces battery drain on the remote.

---

## Building from source

Requirements: Python 3.11+, `aiohttp`, `pywin32`

```bash
pip install -r requirements.txt
pip install pyinstaller

pyinstaller --onefile --name mpchc-bridge --noconsole \
  --hidden-import win32timezone \
  --hidden-import pywintypes \
  --hidden-import tkinter \
  service.py
```

The resulting `dist/mpchc-bridge.exe` is fully self-contained.

---

## Architecture

```
MPC-HC (port 13579)
      ↑  HTTP localhost
      |
mpchc-bridge (port 13580)
  ├── GET  /status        HTTP poll → JSON response
  ├── GET  /ws            WebSocket → push diffs to connected clients
  ├── POST /command/{cmd} WM_COMMAND via PostMessageW (no HTTP round-trip)
  ├── POST /seek          HTTP → MPC-HC /command.html
  ├── POST /volume        HTTP → MPC-HC /command.html
  ├── POST /audio/{n}     WM_COMMAND (offset 50000 + n)
  ├── POST /subtitle/{n}  WM_COMMAND (offset 50935 + n)
  └── POST /open          subprocess → mpc-hc64.exe <path>
```

Commands that require precise timing (play/pause, track switching) use `PostMessageW` directly to the MPC-HC window handle — bypassing the HTTP layer entirely for minimal latency.

---

## License

MIT
