#!/usr/bin/env python3
"""Runs on the phone (Termux, Korean IP), started by run_intraday.sh every hour: reads Upbit's and Bithumb's notices, keeps the delisting ones
(거래지원 종료 / 상장폐지) of the last 60 days (Upbit announces about a month ahead), and sends them to the 'upbit-notices' workflow, which stores them in data/upbit_delist.json
for the daily perp report (GitHub's servers are refused by both exchanges). Sent when something changed, and at least every 6 hours.
Every run that cannot read a source writes one line to the trigger log; the file is fetched fresh from the public repo each hour, so it can be fixed without touching the phone."""
import datetime as dt
import html
import json
import os
import re
import time
import urllib.error
import urllib.request

LOG, STATE = os.environ["LOGF"], os.environ["STATE"]
KST = dt.timezone(dt.timedelta(hours=9))
UA = "Mozilla/5.0 (Linux; Android 14; SM-S918N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Mobile Safari/537.36"
DELIST = re.compile(r"거래지원\s*종료|상장\s*폐지|delist|trading support (?:end|termination)", re.I)
DAYS = 60


def log(m):
    with open(LOG, "a") as f:
        f.write(f"{time.strftime('%F %T')} notices: {m}\n")


def get(url, headers=None, timeout=30):
    h = {"User-Agent": UA, "Accept-Language": "ko-KR,ko;q=0.9"}
    h.update(headers or {})
    return urllib.request.urlopen(urllib.request.Request(url, headers=h), timeout=timeout).read()


def when(v):
    """ISO string with +09:00 from an ISO time, epoch (s or ms) or 'YYYY.MM.DD[ HH:MM]' text; None when not a time."""
    if v in (None, ""):
        return None
    try:
        x = float(v)
        x = x / 1000 if x > 1e12 else x
        return dt.datetime.fromtimestamp(x, KST).isoformat()
    except (TypeError, ValueError):
        pass
    s = str(v).strip()
    try:
        d = dt.datetime.fromisoformat(s.replace("Z", "+00:00"))
        return (d if d.tzinfo else d.replace(tzinfo=KST)).isoformat()
    except ValueError:
        pass
    m = re.search(r"(\d{4})[.\-/]\s*(\d{1,2})[.\-/]\s*(\d{1,2})(?:\D+(\d{1,2}):(\d{2}))?", s)
    if m:
        y, mo, d, h, mi = (int(x) if x else 0 for x in m.groups())
        try:
            return dt.datetime(y, mo, d, h, mi, tzinfo=KST).isoformat()
        except ValueError:
            return None
    return None


def walk(x, out):
    """Every dict that has a title and a time, anywhere in a JSON answer."""
    if isinstance(x, dict):
        t = next((x[k] for k in ("title", "annTitle", "subject", "name") if isinstance(x.get(k), str)), None)
        w = next((when(x[k]) for k in ("listed_at", "first_listed_at", "published_at", "publishedAt", "publish_date", "modified_at", "created_at", "createdAt", "date", "regDate", "reg_date", "pubDate") if x.get(k) not in (None, "") and when(x[k])), None)
        if t and w:
            out.append({"title": html.unescape(t).strip()[:200], "listed_at": w})
        for v in x.values():
            walk(v, out)
    elif isinstance(x, list):
        for v in x:
            walk(v, out)


def keep(items):
    cut = time.time() - DAYS * 86400
    out, seen = [], set()
    for n in items:
        try:
            ok = dt.datetime.fromisoformat(n["listed_at"]).timestamp() >= cut
        except Exception:
            ok = False
        k = (n["title"], n["listed_at"])
        if ok and k not in seen and DELIST.search(n["title"]):
            seen.add(k)
            out.append(n)
    return sorted(out, key=lambda n: (n["listed_at"], n["title"]))


def load_state():
    try:
        return json.load(open(STATE))
    except Exception:
        return {}


def upbit(st):
    base = "https://api-manager.upbit.com/api/v1/announcements"
    full = {"Accept": "application/json", "Origin": "https://upbit.com", "Referer": "https://upbit.com/"}
    plain = {"Accept": "application/json"}
    tries = [st.get("upbit_variant")] if st.get("upbit_variant") else []
    tries += [v for v in ([50, "trade", "full"], [20, "trade", "full"], [20, "trade", "plain"], [50, "trade", "plain"], [20, "", "plain"], [10, "trade", "plain"]) if v not in tries]
    last = ""
    for v in tries:
        pp, cat, hd = v
        items, ok = [], True
        for page in range(1, 13):
            q = f"?os=web&page={page}&per_page={pp}" + (f"&category={cat}" if cat else "")
            try:
                d = json.loads(get(base + q, full if hd == "full" else plain))
            except urllib.error.HTTPError as e:
                last = f"HTTP {e.code} (per_page={pp}, category={cat or '-'}, headers={hd})"
                ok = False
                break
            except Exception as e:
                last = f"{type(e).__name__} {str(e)[:50]}"
                ok = False
                break
            got = []
            walk(d, got)
            items += got
            if not got or min(g["listed_at"] for g in got) < (dt.datetime.now(KST) - dt.timedelta(days=DAYS)).isoformat():
                break
        if ok:
            st["upbit_variant"] = v
            return keep(items), len(items)
    log(f"Upbit read failed ({last})")
    return None, 0


def bithumb(st):
    """Bithumb's notice site; the address that works is remembered. Parses JSON answers and the HTML list alike."""
    cands = [st.get("bithumb_url")] if st.get("bithumb_url") else []
    cands += [u for u in ("https://feed.bithumb.com/notice", "https://feed.bithumb.com/notice?category=", "https://api.bithumb.com/v1/notices?count=100", "https://feed.bithumb.com/api/v1/notices") if u not in cands]
    last = ""
    for u in cands:
        items = []
        try:
            raw = get(u, {"Accept": "text/html,application/json"}).decode("utf-8", "replace")
        except urllib.error.HTTPError as e:
            last = f"{u.split('/')[2]}{u[u.index('/', 8):][:24]} HTTP {e.code}"
            continue
        except Exception as e:
            last = f"{u.split('/')[2]} {type(e).__name__}"
            continue
        try:
            walk(json.loads(raw), items)
        except ValueError:
            for m in re.finditer(r'<script id="__NEXT_DATA__"[^>]*>(.*?)</script>', raw, re.S):
                try:
                    walk(json.loads(m.group(1)), items)
                except ValueError:
                    pass
            for m in re.finditer(r'<a\b[^>]*href="[^"]*notice[^"]*?(\d+)[^"]*"[^>]*>(.*?)</a>', raw, re.S):
                txt = html.unescape(re.sub(r"<[^>]+>", " ", m.group(2)))
                txt = re.sub(r"\s+", " ", txt).strip()
                w = when(txt)
                if txt and w:
                    items.append({"title": re.sub(r"\d{4}[.\-/]\s*\d{1,2}[.\-/]\s*\d{1,2}(\s+\d{1,2}:\d{2})?", "", txt).strip()[:200], "listed_at": w})
        if items:
            st["bithumb_url"] = u
            return keep(items), len(items)
        last = f"{u.split('/')[2]}{u[u.index('/', 8):][:24]}: answered {len(raw)} bytes but no notice with a date was found" + (f" (starts: {re.sub(chr(10), ' ', raw[:80])!r})" if len(raw) < 20000 else "")
    log(f"Bithumb read failed ({last})")
    return None, 0


def main():
    st = load_state()
    up, n_up = upbit(st)
    bi, n_bi = bithumb(st)
    if up is not None or bi is not None:
        log(f"read Upbit {('%d notices, %d delisting' % (n_up, len(up))) if up is not None else 'FAILED'}; Bithumb {('%d notices, %d delisting' % (n_bi, len(bi))) if bi is not None else 'FAILED'}")
    sig = json.dumps([up, bi], ensure_ascii=False, sort_keys=True)
    if up is None and bi is None:
        json.dump(st, open(STATE, "w"))
        return
    if st.get("sig") == sig and time.time() - st.get("t", 0) < 6 * 3600:
        json.dump(st, open(STATE, "w"))
        return
    data = {"updated": dt.datetime.now(dt.timezone.utc).isoformat()}
    if up is not None:
        data["notices"] = up
    if bi is not None:
        data["bithumb"] = bi
    payload = json.dumps({"inputs": {"data": json.dumps(data, ensure_ascii=True)}, "ref": os.environ["BRANCH"]}).encode()
    try:
        r = urllib.request.Request(f"https://api.github.com/repos/{os.environ['REPO']}/actions/workflows/upbit-notices.yml/dispatches", data=payload, method="POST",
                                   headers={"Authorization": "Bearer " + os.environ["TOKEN"], "Accept": "application/vnd.github+json", "Content-Type": "application/json"})
        code = urllib.request.urlopen(r, timeout=30).status
    except urllib.error.HTTPError as e:
        code = e.code
    except Exception as e:
        code = type(e).__name__
    if code == 204:
        st.update(sig=sig, t=time.time())
        log("sent to GitHub")
    else:
        log(f"send FAILED (HTTP {code}; 403/404 = the token cannot start the upbit-notices workflow)")
    json.dump(st, open(STATE, "w"))


try:
    main()
except Exception as e:
    log(f"crashed ({type(e).__name__}: {str(e)[:80]})")
