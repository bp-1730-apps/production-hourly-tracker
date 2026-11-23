"""Fetches today's and yesterday's L2L pitch records for all six lines (MP4, PXM6, 2A, PL1, PL2, 2E) and saves them to l2l.json (read only).
Run by the GitHub workflow every ~5 minutes. Needs only Python 3 (standard library). Key comes from env L2L_KEY."""
import json, os, re, sys, urllib.parse, urllib.request, datetime
try:
    from zoneinfo import ZoneInfo
    NOW = datetime.datetime.now(ZoneInfo("America/Los_Angeles")).replace(tzinfo=None)
except Exception:
    NOW = datetime.datetime.utcnow() - datetime.timedelta(hours=8)
BASE = os.environ.get("L2L_BASE", "https://lakeviewfarms.leading2lean.com/api/1.0")
KEY = os.environ.get("L2L_KEY", "")
SITE = "2"
LCODE = {"2A": "BP-LINE1", "PXM6": "BP-LINE2", "MP4": "BP-LINE3", "PL1": "BP-LINE4", "PL2": "BP-LINE5", "2E": "BP-LINE6"}
LINES = {"MP4": 3049, "PXM6": 3193}          # known L2L numeric line ids; the others are looked up from the lines list below
OUT = os.environ.get("OUT_FILE", "l2l.json")
CODE_KEYS = ("code", "product_code", "productcode", "item_code", "itemcode", "item_number", "sku", "number", "partnumber", "part_number")
PROD = {}


def get(path, params):
    p = dict(params); p["site"] = SITE; p["auth"] = KEY
    url = BASE + "/" + path + "/?" + urllib.parse.urlencode(p)
    with urllib.request.urlopen(urllib.request.Request(url, headers={"User-Agent": "hourly-tracker-feed"}), timeout=40) as r:
        return json.loads(r.read().decode("utf-8", "replace"))


def as_list(j):
    if isinstance(j, list):
        return j
    if isinstance(j, dict):
        for k in ("objects", "results", "data"):
            if isinstance(j.get(k), list):
                return j[k]
    return []


def id_of(v):
    if v in (None, ""):
        return ""
    return str(v.get("id", "")) if isinstance(v, dict) else str(v)


def code_of(o):
    if not isinstance(o, dict):
        return ""
    for k in CODE_KEYS:
        if o.get(k) not in (None, ""):
            return str(o[k])
    for k in ("description", "name", "title", "label"):
        m = re.search(r"(?<!\d)(\d{6,7})(?!\d)", str(o.get(k) or ""))
        if m:
            return m.group(1)
    return ""


def product(pid):
    if not pid:
        return ""
    if pid not in PROD:
        code = ""
        for ep in ("productcomponents", "products"):
            try:
                j = get(ep, {"id": pid}); lst = as_list(j)
                if lst:
                    o = lst[0]
                elif isinstance(j, dict) and isinstance(j.get("data"), dict):   # a single record comes back as {"success":true,"data":{...}}
                    o = j["data"]
                else:
                    o = j if isinstance(j, dict) and "success" not in j else None
                code = code_of(o)
            except Exception:
                code = ""
            if code:
                break
        PROD[pid] = code
    return PROD[pid]


def find_lines():
    """Looks up the numeric id of every line from L2L's lines list (a line is found by its BP-LINEn code or its name)."""
    try:
        for rec in as_list(get("lines", {"limit": 500})):
            vals = [str(v).strip().upper() for v in rec.values() if isinstance(v, (str, int))]
            for name, code in LCODE.items():
                if rec.get("id") is not None and (code in vals or name.upper() in vals):
                    LINES[name] = rec["id"]
    except Exception as e:
        print("could not read the lines list:", e)
    for name in LCODE:
        print(name, "->", LINES.get(name, "NOT FOUND"))


def pull(line):
    d = NOW.date()
    a = (d - datetime.timedelta(days=2)).isoformat() + " 12:00:00"
    b = (d + datetime.timedelta(days=1)).isoformat() + " 12:00:00"
    rows, offset = [], 0
    while offset < 5000:
        j = get("pitches", {"line": LINES[line], "pitch_start__gte": a, "pitch_start__lte": b, "limit": 500, "offset": offset})
        if isinstance(j, dict) and j.get("success") is False:
            raise RuntimeError("L2L says: %s" % j.get("error"))
        got = as_list(j); rows += got
        size = j.get("limit") if isinstance(j, dict) and isinstance(j.get("limit"), int) and j.get("limit") > 0 else 500
        if not got or len(got) < min(size, 500):
            break
        offset += len(got)
    for r in rows:
        r["actual_product_code"] = product(id_of(r.get("actual_product") or r.get("actual_product_id")))
        r["planned_product_code"] = product(id_of(r.get("planned_product") or r.get("planned_product_id") or r.get("product") or r.get("product_id")))
    return rows


def main():
    if not KEY:
        sys.exit("L2L_KEY is not set")
    old = {}
    try:
        old = json.load(open(OUT, encoding="utf-8"))
    except Exception:
        pass
    out = {"updated": old.get("updated", ""), "lines": dict(old.get("lines") or {}), "errors": []}
    find_lines()
    for line in LCODE:
        if line not in LINES:
            out["errors"].append("%s: line id not found in L2L" % line); continue
        try:
            out["lines"][line] = pull(line)
            print(line, len(out["lines"][line]), "records")
        except Exception as e:
            out["errors"].append("%s: %s" % (line, e)); print("FAILED", line, e)
    if out["lines"] != old.get("lines") or out["errors"] != old.get("errors"):
        out["updated"] = NOW.strftime("%Y-%m-%d %H:%M")      # plant-local time of the last change
        json.dump(out, open(OUT, "w", encoding="utf-8"), separators=(",", ":"))
        print("saved", OUT)
    else:
        print("no change")
    if out["errors"] and not any(out["lines"].get(l) for l in LCODE):
        sys.exit(1)


main()
