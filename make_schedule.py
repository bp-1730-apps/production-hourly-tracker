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


def add(lst, run):
    """Adds a run unless an identical one is already there (workbooks often repeat a block)."""
    same = [r for r in lst if r["item"] == run["item"] and r["s"] == run["s"] and r["e"] == run["e"] and abs(r["need"] - run["need"]) < 0.5]
    if not same:
        lst.append(run)


def parse_schedule(rows, week=None):
    """Finds columns by their header text. Returns {'LINE|date': [run, ...]}.
    Newer workbooks have planned start/end times; older ones do not (those runs get s=e=None = daily totals only).
    week = list of (first, last) ISO dates from the file name: a block dated outside that week is a stale copy and is skipped."""
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
    if cols["prod"] < 0 or cols["need"] < 0:
        raise ValueError("Could not find the PRODUCT / CASES NEEDED headers.")
    heads = [r for r in keys if _colof(rows[r], "LINE:") >= 0]
    out, skipped = {}, 0
    for bi, r in enumerate(heads):
        end = heads[bi + 1] - 1 if bi + 1 < len(heads) else keys[-1]
        lc, dc = _colof(rows[r], "LINE:"), _colof(rows[r], "DATE/DAY:")
        line = str(rows[r].get(lc + 1, "")).strip().upper()
        dv = rows[r].get(dc + 1) if dc >= 0 else None
        if not isinstance(dv, float) or line not in ALL:
            continue
        day = serial_to_iso(dv)
        if week and not any(a <= day <= b for a, b in week):
            skipped += 1
            continue
        key = line + "|" + day
        out.setdefault(key, [])
        for rr in range(r + 1, end + 1):
            d = rows.get(rr)
            if not d:
                continue
            p = d.get(cols["prod"])
            if not isinstance(p, str) or not p.strip():
                continue
            std = d.get(cols["std"]) if cols["std"] >= 0 and isinstance(d.get(cols["std"]), float) else 0
            hrs = d.get(cols["hrs"]) if cols["hrs"] >= 0 and isinstance(d.get(cols["hrs"]), float) else 0
            need = std * hrs if std * hrs > 0 else (d.get(cols["need"]) if isinstance(d.get(cols["need"]), float) else 0)
            it = d.get(cols["item"]) if cols["item"] >= 0 else ""
            it = str(int(it)) if isinstance(it, float) else str(it or "")
            if cols["start"] >= 0:
                s, e = d.get(cols["start"]), d.get(cols["start"] + 1)
                if isinstance(s, float) and isinstance(e, float) and e > s and s < 2:
                    add(out[key], dict(item=it, name=p.strip(), s=s * 1440, e=e * 1440, need=need, std=std, hrs=hrs))
                    continue
                if cols["start"] >= 0 and (isinstance(s, float) or isinstance(e, float)):
                    continue
            if need > 0:                                   # no times in this layout: keep the daily total
                add(out[key], dict(item=it, name=p.strip(), s=None, e=None, need=need, std=std, hrs=hrs))
    out["_skipped"] = skipped
    return out


def week_of(path):
    """('2026-02-02', '2026-02-06') from a name like 'SCHEDULE FOR WEEK 02.02.26-02.06.26'. Year typos are tolerated: month/day decide."""
    m = re.search(r"(\d\d)\.(\d\d)\.(\d\d)\s*-\s*(\d\d)\.(\d\d)\.(\d\d)", os.path.basename(path))
    if not m:
        return None
    m1, d1, y1, m2, d2, y2 = [int(x) for x in m.groups()]
    y2 += 2000
    y1 = y2 - 1 if (m1, d1) > (m2, d2) else y2
    out = []
    for k in (0, 1):                               # also try one year later, because some names have the year typed wrong
        try:
            a = datetime.date(y1 + k, m1, d1) - datetime.timedelta(days=1)
            b = datetime.date(y2 + k, m2, d2) + datetime.timedelta(days=1)
        except ValueError:
            continue
        out.append((a.isoformat(), b.isoformat()))
    return out or None


def rounded(runs):
    return [dict(item=r["item"], name=r["name"], s=None if r["s"] is None else round(r["s"], 3), e=None if r["e"] is None else round(r["e"], 3), need=round(r["need"], 2),
                 std=round(r["std"], 2), hrs=round(r["hrs"], 4)) for r in runs]


def pick_files():
    files = sorted(set(glob.glob("*.xls*") + glob.glob("**/*.xls*", recursive=True)))
    def order(p):                                  # copies first, so the original of the same week wins
        b = os.path.basename(p)
        return (0 if b.lower().startswith("copy of") else 1, p)
    return sorted(files, key=order)


def main():
    files = pick_files()
    if not files:
        sys.exit("No .xlsb/.xlsx workbook found in the repository")
    plans = {}                                     # everything saved so far: date -> line -> runs
    for f in glob.glob("data/plans-*.json"):
        for d, lines in json.load(open(f, encoding="utf-8")).get("plans", {}).items():
            plans.setdefault(d, {}).update(lines)
    before = json.dumps(plans, sort_keys=True)
    used = 0
    for path in files:                             # a day in a newer workbook replaces the saved one; other days stay
        try:
            parsed = parse_schedule(read_sheet(path), week_of(path))
        except Exception as ex:
            print("skipped", path, "-", str(ex)[:80]); continue
        skipped = parsed.pop("_skipped", 0)
        n = 0
        for key, runs in parsed.items():
            line, date = key.split("|")
            if runs:
                plans.setdefault(date, {})[line] = rounded(runs); n += 1
        used += 1
        print("read", os.path.basename(path), "-", n, "line-days" + (", %d stale block(s) ignored" % skipped if skipped else ""))
    if json.dumps(plans, sort_keys=True) == before and os.path.exists("schedule.json"):
        print("no change"); return
    os.makedirs("data", exist_ok=True)
    months = {}
    for d, lines in plans.items():
        months.setdefault(d[:7], {})[d] = lines
    for m, p in months.items():                    # one file per month, used by the Trends tab
        json.dump({"plans": p}, open("data/plans-%s.json" % m, "w"), separators=(",", ":"))
    today = datetime.datetime.now(datetime.timezone.utc) - datetime.timedelta(hours=8)
    cut = (today - datetime.timedelta(days=21)).strftime("%Y-%m-%d")
    recent = {d: l for d, l in plans.items() if d >= cut}      # the live page only needs the last 3 weeks and the future
    json.dump({"updated": today.strftime("%Y-%m-%d %H:%M:%S"), "plans": recent}, open("schedule.json", "w"), separators=(",", ":"))
    print("saved", len(plans), "days (", len(recent), "recent ) from", used, "workbook(s)")


main()
