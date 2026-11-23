"""Saves L2L actuals into data/l2l-YYYY-MM.json (one small file per month) so trends can look back over the year.
Daily run: last 3 days. Backfill: set ARCH_DAYS (e.g. 365). Read only; needs only Python 3. Key comes from env L2L_KEY."""
import json, os, sys, datetime
import l2l_feed as F

DAYS = int(os.environ.get("ARCH_DAYS", "3") or "3")
ACTUAL_KEYS = ("actual", "actual_quantity", "actual_qty", "quantity", "total_actual", "actual_count", "good_quantity")


def actual(r):
    for k in ACTUAL_KEYS:
        v = r.get(k)
        if v not in (None, ""):
            try:
                return round(float(v), 2)
            except ValueError:
                pass
    return None


def compact(r):
    a = actual(r)                      # blank actual (changeover) is not useful for trends
    s = str(r.get("pitch_start") or "").replace("T", " ")[:16]
    if a is None or not s:
        return None
    e = str(r.get("pitch_end") or "").replace("T", " ")[:16]
    return [s, e, a, r.get("actual_product_code") or r.get("planned_product_code") or ""]


def main():
    if not F.KEY:
        sys.exit("L2L_KEY is not set")
    F.find_lines()
    end = F.NOW.date() + datetime.timedelta(days=1)
    start = end - datetime.timedelta(days=DAYS)
    months = {}                       # month -> line -> {key: record}
    failed = 0
    for line in F.LCODE:
        if line not in F.LINES:
            print("skip", line, "(no id)"); failed += 1; continue
        d = start
        while d <= end:
            e = min(d + datetime.timedelta(days=6), end)
            try:
                rows = F.pull_range(line, d.isoformat() + " 00:00:00", e.isoformat() + " 23:59:59")
            except Exception as ex:
                print("FAILED", line, d, ex); failed += 1; d = e + datetime.timedelta(days=1); continue
            for r in rows:
                c = compact(r)
                if c:
                    months.setdefault(c[0][:7], {}).setdefault(line, {})["|".join([c[0], c[1], str(c[3])])] = c
            print(line, d, "to", e, len(rows), "records")
            d = e + datetime.timedelta(days=1)
    os.makedirs("data", exist_ok=True)
    for m, lines in sorted(months.items()):
        path = "data/l2l-%s.json" % m
        try:
            cur = json.load(open(path, encoding="utf-8")).get("lines", {})
        except Exception:
            cur = {}
        out = {}
        for line in set(cur) | set(lines):
            keyed = {"|".join([r[0], r[1], str(r[3])]): r for r in cur.get(line, [])}
            keyed.update(lines.get(line, {}))
            out[line] = sorted(keyed.values(), key=lambda r: r[0])
        if out != cur:
            json.dump({"lines": out}, open(path, "w", encoding="utf-8"), separators=(",", ":"))
            print("saved", path)
    if failed and not months:
        sys.exit(1)


if __name__ == "__main__":
    main()
