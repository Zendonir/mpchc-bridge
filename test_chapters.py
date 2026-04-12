"""
Quick test script: read MKV chapters + tracks from a file.
Usage:  python test_chapters.py "Y:\Filme\Movie.mkv"
"""
import sys


# ── minimal EBML helpers ───────────────────────────────────────────────────────

def _ebml_id(buf, pos):
    b = buf[pos]
    if b >= 0x80: return b, pos + 1
    if b >= 0x40: return (b << 8) | buf[pos+1], pos + 2
    if b >= 0x20: return (b << 16) | (buf[pos+1] << 8) | buf[pos+2], pos + 3
    if b >= 0x10: return (b << 24) | (buf[pos+1] << 16) | (buf[pos+2] << 8) | buf[pos+3], pos + 4
    raise ValueError(f"bad EBML ID {b:#x}")

def _ebml_size(buf, pos):
    b = buf[pos]
    if b >= 0x80: return b & 0x7F, pos + 1
    if b >= 0x40: return ((b & 0x3F) << 8) | buf[pos+1], pos + 2
    if b >= 0x20: return ((b & 0x1F) << 16) | (buf[pos+1] << 8) | buf[pos+2], pos + 3
    if b >= 0x10: return ((b & 0x0F) << 24) | (buf[pos+1] << 16) | (buf[pos+2] << 8) | buf[pos+3], pos + 4
    if b >= 0x08:
        v = b & 0x07
        for i in range(4): v = (v << 8) | buf[pos+1+i]
        return v, pos + 5
    if b == 0x01:
        v = 0
        for i in range(7): v = (v << 8) | buf[pos+1+i]
        return v, pos + 8
    return -1, pos + 1


# ── chapter parser ─────────────────────────────────────────────────────────────

def _parse_mkv_chapters(buf, pos, el_end, n):
    _ID_EDITION            = 0x45B9
    _ID_CHAPTER_ATOM       = 0xB6
    _ID_CHAPTER_TIME_START = 0x91
    _ID_CHAPTER_FLAG_HID   = 0x98
    _ID_CHAPTER_DISPLAY    = 0x80
    _ID_CHAP_STRING        = 0x85

    chapters = []
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
                                    names.append(buf[dpos2: dpos2+dsz].decode("utf-8", errors="replace"))
                                dpos = dpos2 + dsz
                        fpos = fpos2 + fsz
                    if not hidden and names:
                        chapters.append({"name": names[0], "time_ms": time_ns // 1_000_000})
                apos = atom_end
        pos = edition_end
    chapters.sort(key=lambda c: c["time_ms"])
    return chapters


# ── full MKV parser ────────────────────────────────────────────────────────────

def read_mkv(filepath):
    _ID_CHAPTERS = 0x1043A770
    _ID_CLUSTER  = 0x1F43B675
    _ID_SEEKHEAD = 0x114D9B74
    _ID_SEEK     = 0x4DBB
    _ID_SEEKID   = 0x53AB
    _ID_SEEKPOS  = 0x53AC
    _ID_TRACKS   = 0x1654AE6B
    _ID_ENTRY    = 0xAE
    _ID_NUM      = 0xD7
    _ID_TYPE     = 0x83
    _ID_NAME     = 0x536E
    _ID_LANG     = 0x22B59C
    _ID_CODEC    = 0x86

    with open(filepath, "rb") as fh:
        buf = fh.read(2097152)  # 2 MB
    n = len(buf)
    assert buf[:4] == b"\x1a\x45\xdf\xa3", "Not an MKV file"

    pos = 0
    _, pos = _ebml_id(buf, pos)
    esz, pos = _ebml_size(buf, pos)
    pos += esz  # skip EBML header

    _, pos = _ebml_id(buf, pos)
    esz, pos = _ebml_size(buf, pos)
    seg_data_start = pos
    seg_end = (pos + esz) if esz >= 0 else n

    result = {"audio": [], "subtitle": [], "chapters": []}
    chapters_offset = -1

    while pos < min(seg_end, n) - 4:
        try:
            eid, npos = _ebml_id(buf, pos)
            esz, npos = _ebml_size(buf, npos)
        except (IndexError, ValueError):
            break
        if esz < 0 or esz > n: esz = n - npos
        el_end = npos + esz

        if eid == _ID_CLUSTER:
            break
        elif eid == _ID_CHAPTERS:
            result["chapters"] = _parse_mkv_chapters(buf, npos, el_end, n)
        elif eid == _ID_SEEKHEAD:
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
                        fdata = buf[epos2: epos2+fsz]
                        if fid == _ID_SEEKID:   sk_id = int.from_bytes(fdata, "big")
                        elif fid == _ID_SEEKPOS: sk_off = int.from_bytes(fdata, "big")
                        epos = epos2 + fsz
                    if sk_id == _ID_CHAPTERS and sk_off >= 0:
                        chapters_offset = sk_off
                spos = se_end
        elif eid == _ID_TRACKS:
            tpos = npos
            while tpos < min(el_end, n) - 2:
                try:
                    tid, tpos2 = _ebml_id(buf, tpos)
                    tsz, tpos2 = _ebml_size(buf, tpos2)
                except (IndexError, ValueError):
                    break
                if tsz < 0: tsz = 0
                te_end = min(tpos2 + tsz, n)
                if tid == _ID_ENTRY:
                    t = {"number": 0, "type": 0, "name": "", "lang": "", "codec": ""}
                    epos = tpos2
                    while epos < te_end - 2:
                        try:
                            fid, epos2 = _ebml_id(buf, epos)
                            fsz, epos2 = _ebml_size(buf, epos2)
                        except (IndexError, ValueError):
                            break
                        if fsz < 0: fsz = 0
                        fd = buf[epos2: epos2+fsz]
                        if fid == _ID_NUM:    t["number"] = int.from_bytes(fd, "big")
                        elif fid == _ID_TYPE: t["type"]   = int.from_bytes(fd, "big")
                        elif fid == _ID_NAME: t["name"]   = fd.decode("utf-8", errors="replace")
                        elif fid == _ID_LANG: t["lang"]   = fd.decode("ascii", errors="replace").rstrip("\x00")
                        elif fid == _ID_CODEC: t["codec"] = fd.decode("ascii", errors="replace").rstrip("\x00")
                        epos = epos2 + fsz
                    if t["type"] == 2:
                        t["pos"] = len(result["audio"])
                        result["audio"].append(t)
                    elif t["type"] == 17:
                        t["pos"] = len(result["subtitle"])
                        result["subtitle"].append(t)
                tpos = te_end
        pos = npos + esz

    # SeekHead fallback for chapters outside 2 MB window
    if not result["chapters"] and chapters_offset >= 0:
        chap_abs = seg_data_start + chapters_offset
        with open(filepath, "rb") as fh:
            fh.seek(chap_abs)
            hdr = fh.read(12)
        cpos = 0
        c_eid, cpos = _ebml_id(hdr, cpos)
        c_esz, cpos = _ebml_size(hdr, cpos)
        if c_eid == _ID_CHAPTERS and c_esz > 0:
            read_sz = min(c_esz, 262144)
            with open(filepath, "rb") as fh:
                fh.seek(chap_abs + cpos)
                cbuf = fh.read(read_sz)
            result["chapters"] = _parse_mkv_chapters(cbuf, 0, len(cbuf), len(cbuf))
        print(f"  (chapters found via SeekHead at offset {chap_abs})")

    return result


# ── main ───────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python test_chapters.py <path-to-file.mkv>")
        sys.exit(1)

    path = sys.argv[1]
    print(f"Reading: {path}\n")

    try:
        data = read_mkv(path)
    except Exception as ex:
        print(f"ERROR: {ex}")
        sys.exit(1)

    print(f"Audio tracks ({len(data['audio'])}):")
    for t in data["audio"]:
        print(f"  [{t['pos']}] #{t['number']}  lang={t['lang'] or '?'}  codec={t['codec']}  name={t['name']!r}")

    print(f"\nSubtitle tracks ({len(data['subtitle'])}):")
    for t in data["subtitle"]:
        print(f"  [{t['pos']}] #{t['number']}  lang={t['lang'] or '?'}  codec={t['codec']}  name={t['name']!r}")

    print(f"\nChapters ({len(data['chapters'])}):")
    if data["chapters"]:
        for c in data["chapters"]:
            ms = c["time_ms"]
            h, m, s = ms // 3600000, (ms % 3600000) // 60000, (ms % 60000) // 1000
            print(f"  {h:02d}:{m:02d}:{s:02d}  {c['name']}")
    else:
        print("  (none)")
