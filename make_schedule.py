"""Reads the weekly schedule workbook (.xlsb or .xlsx, sheet SCHEDULE) and saves schedule.json for the Hourly Tracker page.
Run by the GitHub workflow whenever the workbook in the repo changes. Needs only Python 3 (standard library)."""
import glob, json, os, re, struct, sys, zipfile, datetime
import xml.etree.ElementTree as ET
ALL = ["2A", "PXM6", "PL1", "MP4", "PL2", "2E"]


def serial_to_iso(v):
    return (datetime.date(1899, 12, 30) + datetime.timedelta(days=int(v))).isoformat()


def now_rel(date):
    """minutes since midnight of the plan date"""
    d = datetime.datetime.fromisoformat(date + "T00:00:00")
    return (datetime.datetime.now() - d).total_seconds() / 60.0


def pend(r):
    return r["s"] + r["hrs"] * 60 if r["hrs"] > 0 else r["e"]


# ======================================================================= reading the workbook (xlsb / xlsx)
def _records(data):
    i, n = 0, len(data)
    while i < n:
        b = data[i]; i += 1
        rt = b & 0x7F
        if b & 0x80:
            b2 = data[i]; i += 1
            rt |= (b2 & 0x7F) << 7
        ln, sh = 0, 0
        while True:
            b = data[i]; i += 1
            ln |= (b & 0x7F) << sh; sh += 7
            if not b & 0x80:
                break
        yield rt, data[i:i + ln]
        i += ln


def _wstr(b, o):
    cch = struct.unpack_from("<I", b, o)[0]
    return b[o + 4:o + 4 + cch * 2].decode("utf-16le", "replace"), o + 4 + cch * 2


def _rk(v):
    if v & 2:
        x = struct.unpack("<i", struct.pack("<I", v))[0] >> 2
    else:
        x = struct.unpack("<d", struct.pack("<II", 0, v & 0xFFFFFFFC))[0]
    return x / 100.0 if v & 1 else float(x)


def _read_xlsb(z, sheet):
    names = z.namelist()
    sst = []
    if "xl/sharedStrings.bin" in names:
        for rt, b in _records(z.read("xl/sharedStrings.bin")):
            if rt == 19:
                sst.append(_wstr(b, 1)[0])
    sheets = []
    for rt, b in _records(z.read("xl/workbook.bin")):
        if rt == 156:
            rid, o = _wstr(b, 8)
            nm, _ = _wstr(b, o)
            sheets.append((nm, rid))
    rels = {}
    for el in ET.fromstring(z.read("xl/_rels/workbook.bin.rels")):
        rels[el.get("Id")] = el.get("Target")
    path = None
    for nm, rid in sheets:
        if nm.strip().upper() == sheet.upper():
            t = rels.get(rid, "")
            path = t.lstrip("/") if t.startswith("/") else "xl/" + t
    if not path:
        raise ValueError('The workbook has no sheet named "%s".' % sheet)
    rows, r = {}, 0
    for rt, b in _records(z.read(path)):
        if rt == 0:
            r = struct.unpack_from("<I", b, 0)[0]
            continue
        if rt not in (2, 5, 6, 7, 8, 9):
            continue
        c = struct.unpack_from("<I", b, 0)[0]
        if rt == 2:
            v = _rk(struct.unpack_from("<I", b, 8)[0])
        elif rt in (5, 9):
            v = struct.unpack_from("<d", b, 8)[0]
        elif rt == 7:
            i = struct.unpack_from("<I", b, 8)[0]
            v = sst[i] if i < len(sst) else None
        else:
            v = _wstr(b, 8)[0]
        if v is not None and v != "":
            rows.setdefault(r + 1, {})[c] = v
    return rows


def _col(ref):
    n = 0
    for ch in re.match(r"[A-Z]+", ref).group(0):
        n = n * 26 + ord(ch) - 64
    return n - 1


def _read_xlsx(z, sheet):
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    sst = []
    if "xl/sharedStrings.xml" in z.namelist():
        for si in ET.fromstring(z.read("xl/sharedStrings.xml")).findall("m:si", ns):
            sst.append("".join(t.text or "" for t in si.iter("{%s}t" % ns["m"])))
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = {e.get("Id"): e.get("Target") for e in ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))}
    path = None
    for s in wb.find("m:sheets", ns):
        if s.get("name").strip().upper() == sheet.upper():
            t = rels[s.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")]
            path = t.lstrip("/") if t.startswith("/") else "xl/" + t
    if not path:
        raise ValueError('The workbook has no sheet named "%s".' % sheet)
    rows = {}
    for row in ET.fromstring(z.read(path)).iter("{%s}row" % ns["m"]):
        for c in row.findall("m:c", ns):
            v = c.find("m:v", ns)
            if v is None or v.text is None:
                continue
            t = c.get("t")
            if t in ("e", "b"):
                continue
            try:
                val = sst[int(v.text)] if t == "s" else (v.text if t in ("str", "inlineStr") else float(v.text))
            except (ValueError, IndexError):
                continue
            rows.setdefault(int(row.get("r")), {})[_col(c.get("r"))] = val
    return rows


def read_sheet(path, sheet="SCHEDULE"):
    z = zipfile.ZipFile(path)
    return _read_xlsb(z, sheet) if "xl/workbook.bin" in z.namelist() else _read_xlsx(z, sheet)


def _colof(d, label):
    for k, v in d.items():
        if isinstance(v, str) and v.strip().upper() == label:
            return k
    return -1


def parse_schedule(rows):
    """Finds columns by their header text. Returns {'LINE|date': [run, ...]}"""
    keys = sorted(rows)
    cols = dict(item=-1, prod=-1, start=-1, need=-1, std=-1, hrs=-1)
    for r in keys:
        d = rows[r]
        i, p, s, e = _colof(d, "ITEM CODE"), _colof(d, "PRODUCT"), _colof(d, "CASES NEEDED"), _colof(d, "TIME END")
        sd, th = _colof(d, "STANDARD/HR"), _colof(d, "TOTAL HR")
        if i >= 0: cols["item"] = i
        if p >= 0: cols["prod"] = p
        if s >= 0: cols["need"] = s
        if sd >= 0: cols["std"] = sd
        if th >= 0: cols["hrs"] = th
        if e > 0 and isinstance(d.get(e - 1), str) and d[e - 1].strip().upper() == "TIME START":
            cols["start"] = e - 1
    if cols["prod"] < 0 or cols["start"] < 0 or cols["need"] < 0:
        raise ValueError("Could not find the PRODUCT / CASES NEEDED / TIME START / TIME END headers.")
    heads = [r for r in keys if _colof(rows[r], "LINE:") >= 0]
    out = {}
    for bi, r in enumerate(heads):
        end = heads[bi + 1] - 1 if bi + 1 < len(heads) else keys[-1]
        lc, dc = _colof(rows[r], "LINE:"), _colof(rows[r], "DATE/DAY:")
        line = str(rows[r].get(lc + 1, "")).strip().upper()
        dv = rows[r].get(dc + 1) if dc >= 0 else None
        if not isinstance(dv, float) or line not in ALL:
            continue
        key = line + "|" + serial_to_iso(dv)
        out.setdefault(key, [])
        for rr in range(r + 1, end + 1):
            d = rows.get(rr)
            if not d:
                continue
            p = d.get(cols["prod"])
            if not isinstance(p, str) or not p.strip():
                continue
            s, e = d.get(cols["start"]), d.get(cols["start"] + 1)
            if not isinstance(s, float) or not isinstance(e, float) or e <= s or s >= 2:
                continue
            std = d.get(cols["std"]) if cols["std"] >= 0 and isinstance(d.get(cols["std"]), float) else 0
            hrs = d.get(cols["hrs"]) if cols["hrs"] >= 0 and isinstance(d.get(cols["hrs"]), float) else 0
            need = std * hrs if std * hrs > 0 else (d.get(cols["need"]) if isinstance(d.get(cols["need"]), float) else 0)
            it = d.get(cols["item"]) if cols["item"] >= 0 else ""
            it = str(int(it)) if isinstance(it, float) else str(it or "")
            out[key].append(dict(item=it, name=p.strip(), s=s * 1440, e=e * 1440, need=need, std=std, hrs=hrs))
    return out


def main():
    files = sorted(glob.glob("*.xlsb") + glob.glob("*.xlsx") + glob.glob("*.xlsm"), key=os.path.getmtime)
    if not files:
        sys.exit("No .xlsb/.xlsx workbook found in the repository root")
    path = files[-1]
    parsed = parse_schedule(read_sheet(path))
    plans = {}
    for key, runs in parsed.items():
        line, date = key.split("|")
        plans.setdefault(date, {})[line] = runs
    stamp = "%s %d" % (path, os.path.getmtime(path))
    old = {}
    try:
        old = json.load(open("schedule.json", encoding="utf-8"))
    except Exception:
        pass
    if old.get("plans") == plans:
        print("no change"); return
    json.dump({"updated": datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%d %H:%M:%S"), "source": path, "plans": plans}, open("schedule.json", "w"), separators=(",", ":"))
    print("saved", len(plans), "days from", path)


main()
