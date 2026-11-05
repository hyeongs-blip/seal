#!/usr/bin/env python3
"""
Daily relative-strength ranking vs the majors basket (or SOL, selectable) -> Excel -> Telegram.

Run:  python ranker.py              (build Excel + send it to Telegram)
      python ranker.py --no-send    (build Excel only; --no-email works too)

Env vars for Telegram: TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID
Dependencies: pip install pandas numpy requests openpyxl
"""
import argparse
import csv
import io
import json
import os
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor

import numpy as np
import pandas as pd
import requests
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ----------------------------- settings -----------------------------
N_COINS = 200            # universe size (by market cap, after exclusions)
BENCH = "SOL"            # always loaded as the data anchor; it is the benchmark itself only when BENCH_MODE is "sol"
MAJORS = ["ETH", "SOL", "BNB", "XRP", "ADA", "DOGE", "AVAX", "LINK"]  # co-movement basket
# What the momentum maths (beta, weekly residuals, 30-day beta, hype tests) measures every coin against:
#   "majors" = equal-weight average of the other majors (the coin itself left out)      <- default
#   "sol"    = SOL alone (SOL then shows as BENCHMARK and is not a candidate)
# The "Run workflow" form on GitHub can override this for one run (bench_mode).
BENCH_MODE = os.environ.get("BENCH_MODE", "").strip().lower()
if BENCH_MODE not in ("majors", "sol"):
    BENCH_MODE = "majors"
BENCH_LBL = "majors" if BENCH_MODE == "majors" else "SOL"
BENCH_EX = [BENCH] if BENCH_MODE == "sol" else []     # coins that are the benchmark themselves (left out of ranking)
WINDOW_WEEKS = 12        # weeks used for ranking
SKIP_DAYS = 7            # the latest 7 days are NOT part of the 12 weeks
CONSIST_WEEKS = 3        # the most recent 3 weeks of the window are the consistency check
def _env_num(name, default):
    try:
        return float(os.environ.get(name, "").strip())
    except ValueError:
        return default


# 1-minute correlation floors; a coin below the floor of the chosen test is flagged IGNORE (untuned).
# The "Run workflow" form on GitHub can override all three for one run; empty or "default" there = the values below.
MIN_CORR_SOL = _env_num("MIN_CORR_SOL", 0.45)
MIN_CORR_MAJ = _env_num("MIN_CORR_MAJ", 0.45)
CORR_TEST = os.environ.get("CORR_TEST", "").strip().lower()      # which correlation decides IGNORE:
if CORR_TEST not in ("sol", "majors"):                           # which correlation decides IGNORE: "majors" (corr to the majors basket) or "sol"
    CORR_TEST = "majors"                                         # the benchmark for beta / residuals is set separately by BENCH_MODE
# a run started with explicit values in the Run workflow form is a what-if: it must not overwrite the saved history
WHATIF_RUN = (os.environ.get("CORR_TEST", "").strip().lower() in ("sol", "majors")
              or os.environ.get("BENCH_MODE", "").strip().lower() in ("majors", "sol")
              or bool(os.environ.get("MIN_CORR_SOL", "").strip()) or bool(os.environ.get("MIN_CORR_MAJ", "").strip()))
# --- pump / manipulation-prone coins (price-pattern test over the 12-week window; no volume or order-book data) ---
HYPE_ACTION = "exclude"  # "exclude" = strong-hype coins are kept out of the shortlists (and shaded); "shade" = only shade them
HYPE_RUNUP = 4.0         # biggest rise from an earlier low inside the window: 4.0 = +300%  (counts 2 points)
HYPE_SPIKE = 0.5         # biggest 24-hour rise inside the window: +50%  (1 point)
HYPE_VOL_MULT = 3.0      # 30-day daily volatility at least this many times the median coin  (1 point)
HYPE_DECOUPLED_PCT = 0.10  # R2 versus the benchmark in the lowest 10% of all coins  (1 point)
HYPE_FADE = 0.6          # after a run-up of 2x or more, price is at or below 60% of its window peak: the pump is being unwound  (1 point)
HYPE_SOUND_FLOOR = 0.40  # a run-up is judged SOUND (not hype) if on the way up the coin kept at least this hourly correlation to the benchmark (majors / SOL) ...
HYPE_SOUND_TOL = 0.15    # ... and no more than this far below its correlation before the boom
HYPE_SOUND_NOBASE = 0.50 # (if the window holds no pre-boom period to compare with, it must have kept at least this)
HYPE_SETTLE_DAYS = 30    # a run-up whose boom (price first up 300% from its low) is at least this old is also SOUND if the coin has not since acted like a hype coin
HYPE_STRONG, HYPE_WATCH = 3, 2   # points needed: strong (shaded purple, excluded) / watch (shaded peach, still eligible)
MINUTE_DAYS = 2          # days of 1-minute data for the co-movement test
SIZING_DAYS = 30         # days used for volatility and beta in sizing
N_SHORTLIST = 20        # candidates per side on the Portfolio sheet (all pre-selected; you trim the ones you skip)
# --- Timing sheet ---
EVENTS_FILE = "events.csv"   # utc_time (YYYY-MM-DD HH:MM, or a date for a whole day), label, type (scheduled / news)
TIMING_HOURS = 24 * 28       # 4 weeks of hourly data
WF_DAYS = 42                 # walk-forward charts: up to this many earlier daily reports (6 weeks)
WF5_DAYS = 14                # the 5 vs 5 OB/OS pair chart uses only the last 2 weeks
N_OBOS = 5                   # coins per side in the extra time-of-day chart (the most oversold longs and most overbought shorts)
DAY_KEYS = (("wd", "Weekdays (KST Mon-Fri)", "#1f4e79", "o"), ("sat", "Saturday (KST)", "#ed7d31", "s"), ("sun", "Sunday (KST)", "#548235", "^"))
EVENT_BEFORE_H, EVENT_AFTER_H = 1, 3   # a timed event removes this many hours before / after from the "calm" statistics
SPIKE_Z = 3.0                # an hour is a volatility spike if the market or the ratio moves more than this many std devs
KST = 9                      # Seoul = UTC+9
AVG_DAYS = 7            # coins are chosen by their average position over this many days (today plus the earlier ones)
VERSION = "v16 2026-11-04 (parallel price fetching; majors benchmark default, SOL selectable)"
STATE_FILE = "data/ranker/ranker_state.json"   # remembers each day's shortlists so the next report can show changes
OLD_STATE_FILES = ["data/ranker-state.json", "data/state.json", "state.json"]   # earlier names / places; read once as a fallback, then removed
SIZE_BY = "vol"         # "vol" = size by inverse 30-day volatility, "beta" = inverse 30-day beta (floored at 0.3)

NONCOIN_NAME_WORDS = ("tokenized", "tokenised", "bstock", "xstock")   # CoinGecko names of tokenized stocks / ETFs
EXCLUDE = {
    "MSTRB", "SNDKB", "MUB", "CRCLB", "NVDAB", "TSLAB", "SPCXB", "SOLXB",   # Binance bStocks (tokenized US stocks) seen or suspected
    "BTC", "USDT", "USDC", "DAI", "FDUSD", "USDE", "USDS", "TUSD", "PYUSD", "USD1", "USDD",
    "USDG", "RLUSD", "BUIDL", "USDY", "WBTC", "WETH", "STETH", "WSTETH", "WBETH", "WEETH",
    "CBBTC", "CBETH", "BSC-USD", "BNSOL", "JITOSOL", "MSOL", "XAUT", "PAXG",
}
BASE = "https://data-api.binance.vision"   # public market-data host, no API key needed


# ----------------------------- data fetching -----------------------------
POOL_WORKERS = 8
HOST_RPS = 12.0            # requests per second per host while the coins are fetched in parallel (Binance allows far more; this stays well inside its limits)
_TH = {"on": False, "lock": threading.Lock(), "next": {}}


def _throttle(url, backoff=0.0):
    """While fetching in parallel: spaces the requests to one host evenly (backoff > 0 = pause that host for that many seconds after a 429/418)."""
    if not _TH["on"]:
        return
    host = url.split("/")[2] if "//" in url else url
    with _TH["lock"]:
        now = time.time()
        if backoff:
            _TH["next"][host] = max(_TH["next"].get(host, 0.0), now + backoff)
            return
        t = max(now, _TH["next"].get(host, 0.0))
        _TH["next"][host] = t + 1.0 / HOST_RPS
    if t > now:
        time.sleep(t - now)


def get(url, params=None, tries=5):
    for i in range(tries):
        _throttle(url)
        r = requests.get(url, params=params, timeout=30)
        if r.status_code == 200:
            return r.json()
        if r.status_code in (418, 429):
            _throttle(url, backoff=10 * (i + 1))
            time.sleep(5 * (i + 1))
            continue
        if r.status_code == 451:
            sys.exit("Binance blocks this server's region (HTTP 451). Run the script on your own computer instead.")
        r.raise_for_status()
    raise RuntimeError(f"request failed: {url}")


def klines(sym, interval, start, end):
    """Closing prices of fully closed candles, indexed by candle open time (UTC)."""
    t, end_ms = int(start.timestamp() * 1000), int(end.timestamp() * 1000) - 1
    rows = []
    while t < end_ms:
        r = get(f"{BASE}/api/v3/klines",
                dict(symbol=sym + "USDT", interval=interval, startTime=t, endTime=end_ms, limit=1000))
        if not r:
            break
        rows += r
        t = r[-1][0] + 1
        if len(r) < 1000:
            break
        time.sleep(0.05)
    return pd.Series([float(x[4]) for x in rows],
                     index=pd.to_datetime([x[0] for x in rows], unit="ms"), dtype=float)


def fetch_all(syms, h_end, m_end, hours):
    """Hourly (about 12 weeks) and 1-minute (2 days) closes of every coin, fetched on a thread pool. Returns (hs, ms, failed coin notes); the order of syms is kept."""
    def one(s):
        try:
            return s, klines(s, "1h", h_end - pd.Timedelta(hours=hours + 15), h_end), klines(s, "1m", m_end - pd.Timedelta(minutes=MINUTE_DAYS * 1440), m_end), None
        except Exception as ex:
            return s, None, None, ex
    hs, ms, failed, n = {}, {}, [], 0
    _TH["on"] = True
    try:
        with ThreadPoolExecutor(max_workers=POOL_WORKERS) as pool:
            for s, h, m, ex in pool.map(one, syms):
                n += 1
                if ex is not None:
                    failed.append(f"{s} (fetch error)")
                    print(f"  {s}: {ex}")
                else:
                    hs[s], ms[s] = h, m
                print(f"[{n}/{len(syms)}] {s}", flush=True)
    finally:
        _TH["on"] = False
    return hs, ms, failed


def universe():
    """Top coins by market cap (CoinGecko) that have a Binance USDT spot pair."""
    cg = []
    for page in (1, 2):                      # 500 coins by market cap, so 200 are left after exclusions
        cg += get("https://api.coingecko.com/api/v3/coins/markets",
                  dict(vs_currency="usd", order="market_cap_desc", per_page=250, page=page))
        time.sleep(2)
    info = get(f"{BASE}/api/v3/exchangeInfo")
    tradable = {s["baseAsset"] for s in info["symbols"]
                if s["quoteAsset"] == "USDT" and s["status"] == "TRADING"}
    out, seen = [], set()
    for cap_rank, c in enumerate(cg, 1):
        sym = c["symbol"].upper()
        if sym in EXCLUDE or sym in seen or sym not in tradable:
            continue
        if any(w in (c.get("name") or "").lower() for w in NONCOIN_NAME_WORDS):
            continue
        seen.add(sym)
        out.append((sym, cap_rank))
        if len(out) == N_COINS:
            break
    return out


def perp_bases():
    try:
        r = requests.get("https://fapi.binance.com/fapi/v1/exchangeInfo", timeout=30)
        if r.status_code != 200:
            return None
        return {s["baseAsset"] for s in r.json()["symbols"]
                if s["contractType"] == "PERPETUAL" and s["quoteAsset"] == "USDT" and s["status"] == "TRADING"}
    except Exception:
        return None


def perp_bases_cg():
    """Fallback when Binance's futures host is blocked: CoinGecko's list of Binance futures tickers."""
    try:
        time.sleep(2)
        r = requests.get("https://api.coingecko.com/api/v3/derivatives/exchanges/binance_futures",
                         params={"include_tickers": "unexpired"}, timeout=60)
        if r.status_code != 200:
            return None
        bases = {t["base"].upper() for t in r.json().get("tickers", [])
                 if t.get("base") and str(t.get("contract_type", "perpetual")).lower().startswith("perp")}
        return bases if len(bases) >= 100 else None
    except Exception:
        return None


def has_perp(sym, bases):
    if bases is None:
        return "?"
    return "yes" if any((p + sym) in bases for p in ("", "1000", "1000000")) else "no"


# ----------------------------- analysis -----------------------------
def analyze(H, M, caprank, perps, history=None):
    """
    H: hourly closes (rows = hours, oldest first, last row = latest closed hour), columns = symbols
    M: 1-minute closes over the last MINUTE_DAYS days
    Returns (table, longs, shorts, short_scale)
    """
    lp = np.log(H)
    n = len(lp)
    e = n - 1 - SKIP_DAYS * 24                      # last hour of the 12-week window
    r = lp.diff()

    # Benchmark hourly returns, one column per coin. "sol": SOL's returns for everybody. "majors": equal-weight average of the
    # majors with the coin itself left out (so a major is never measured against itself).
    global BENCH_MODE, BENCH_LBL, BENCH_EX
    maj_h = [m for m in MAJORS if m in r.columns]
    if BENCH_MODE == "majors" and len(maj_h) < 4:
        print("Fewer than 4 majors loaded: falling back to SOL as the benchmark for this run")
        BENCH_MODE, BENCH_LBL, BENCH_EX = "sol", "SOL", [BENCH]
    if BENCH_MODE == "sol":
        B = pd.DataFrame({c: r[BENCH] for c in r.columns})
    else:
        rmj = r[maj_h]
        tot, cnt = rmj.sum(axis=1), rmj.notna().sum(axis=1)
        B = pd.DataFrame({c: ((tot - rmj[c].fillna(0.0)) / (cnt - rmj[c].notna().astype(int))) if c in maj_h else tot / cnt
                          for c in r.columns}).replace([np.inf, -np.inf], np.nan)

    def _dev(X):
        return X - X.mean()

    def _beta(X, Y):
        dx, dy = _dev(X), _dev(Y)
        return (dx * dy).sum() / (dy ** 2).sum()

    def _r2(X, Y):
        dx, dy = _dev(X), _dev(Y)
        return (dx * dy).sum() ** 2 / ((dx ** 2).sum() * (dy ** 2).sum())

    # Beta vs the benchmark and weekly residuals over the 12-week window (ends 7 days ago)
    rw, Bw = r.iloc[1:e + 1], B.iloc[1:e + 1]
    beta = _beta(rw, Bw)
    r2 = _r2(rw, Bw)
    ends = [e - 168 * k for k in range(WINDOW_WEEKS)]          # k=0 is the most recent week of the window
    wk = pd.DataFrame(lp.iloc[ends].values - lp.iloc[[x - 168 for x in ends]].values, columns=lp.columns)
    CB = B.fillna(0.0).cumsum()
    wkb = pd.DataFrame(CB.iloc[ends].values - CB.iloc[[x - 168 for x in ends]].values, columns=lp.columns)   # the benchmark's own weekly return
    resid = wk - wkb * beta.values                             # coin return minus beta * benchmark return
    weeks_up = (resid > 0).sum()
    recent = resid.iloc[:CONSIST_WEEKS]
    last_up, last_dn = (recent > 0).all(), (recent < 0).all()
    cum = resid.sum()

    # Volatility and beta for sizing: last 30 days, including the latest week
    rs = r.iloc[-SIZING_DAYS * 24:]
    vol = rs.std() * np.sqrt(24)
    beta30 = _beta(rs, B.iloc[-SIZING_DAYS * 24:])

    # Minute-by-minute co-movement with SOL and with the majors basket
    rm = np.log(M).diff().iloc[1:]
    corr_sol = rm.corrwith(rm[BENCH])      # 1-minute correlation to SOL (whatever the benchmark mode)
    maj = [m for m in MAJORS if m in rm.columns]
    corr_maj = pd.Series({c: rm[c].corr(rm[[m for m in maj if m != c]].mean(axis=1)) for c in rm.columns})

    df = pd.DataFrame({
        "cap_rank": pd.Series(caprank), "weeks_up": weeks_up, "last3_up": last_up, "last3_dn": last_dn,
        "cum_resid": cum, "beta": beta, "r2": r2, "corr_sol": corr_sol, "corr_maj": corr_maj,
        "vol": vol, "beta30": beta30, "chg24": H.iloc[-1] / H.iloc[-25] - 1, "chg6": H.iloc[-1] / H.iloc[-7] - 1,
    }).loc[H.columns]
    df["corr_b"] = df.corr_maj if BENCH_MODE == "majors" else df.corr_sol     # correlation to the active benchmark
    df["perp"] = [has_perp(s, perps) for s in df.index]
    df["inv"] = 1 / (df.vol if SIZE_BY == "vol" else df.beta30.clip(lower=0.3))
    # Overbought/oversold: 6h and 24h price changes, each divided by the coin's typical move over that horizon, 50:50
    sd = df.vol.replace(0, np.nan)
    df["ob"] = (0.5 * df.chg6 / (sd * np.sqrt(6 / 24)) + 0.5 * df.chg24 / sd).fillna(0.0)

    # Pump / manipulation-prone pattern score from the price history alone
    runup = (H / H.cummin()).max()                        # biggest rise from an earlier low in the window
    spike24 = (H / H.shift(24)).max() - 1                 # biggest 24-hour rise
    fade = H.iloc[-1] / H.max()                           # where price sits versus the window peak
    r2q = df.r2.drop(index=BENCH_EX, errors="ignore").quantile(HYPE_DECOUPLED_PCT)
    vmed = df.vol.drop(index=BENCH_EX, errors="ignore").median()
    lr = np.log(H).diff().fillna(0.0)
    Bh = B.fillna(0.0).values
    colix = {c: i for i, c in enumerate(H.columns)}
    spike30 = (H / H.shift(24)).iloc[-HYPE_SETTLE_DAYS * 24:].max() - 1

    def _cc(sv, a, b):
        x, y = lr[sv].values[a:b], Bh[a:b, colix[sv]]
        return float(np.corrcoef(x, y)[0, 1]) if b - a >= 48 and x.std() > 0 and y.std() > 0 else np.nan

    def sound_check(sv):
        """Why a big run-up looks organic: it kept its correlation to the benchmark, or a month passed without hype behaviour."""
        px = H[sv].values
        ratio = px / np.minimum.accumulate(px)
        kp = int(ratio.argmax())
        kl = int(px[:kp + 1].argmin())
        kb = int(np.argmax(ratio >= HYPE_RUNUP))           # the hour the boom was complete: price first stood at the run-up threshold
        out = []
        asc, pre = _cc(sv, kl + 1, kp + 1), _cc(sv, 1, kl + 1)
        if asc == asc:
            if pre == pre and asc >= max(HYPE_SOUND_FLOOR, pre - HYPE_SOUND_TOL):
                out.append(f"{BENCH_LBL} corr {asc:.2f} on the way up vs {pre:.2f} before")
            elif pre != pre and asc >= HYPE_SOUND_NOBASE:
                out.append(f"{BENCH_LBL} corr {asc:.2f} on the way up (no earlier data to compare)")
        age = (len(px) - 1 - kb) / 24
        if age >= HYPE_SETTLE_DAYS:
            c30 = _cc(sv, len(px) - HYPE_SETTLE_DAYS * 24, len(px))
            if c30 == c30 and c30 >= HYPE_SOUND_FLOOR and fade[sv] >= HYPE_FADE and spike30[sv] < HYPE_SPIKE:
                out.append(f"{age:.0f} days since the boom with no hype behaviour ({BENCH_LBL} corr {c30:.2f})")
        return "; ".join(out)
    pts, why, oks = [], [], []
    for sv in df.index:
        p, w, ok_txt = 0, [], ""
        if sv not in BENCH_EX:
            sound = sound_check(sv) if runup[sv] >= HYPE_RUNUP else ""
            if runup[sv] >= HYPE_RUNUP and sound:
                ok_txt = f"run-up {100 * (runup[sv] - 1):+.0f}% judged sound: {sound}"
            elif runup[sv] >= HYPE_RUNUP:
                p += 2
                w.append(f"run-up {100 * (runup[sv] - 1):+.0f}%")
            if spike24[sv] >= HYPE_SPIKE:
                p += 1
                w.append(f"24h spike {100 * spike24[sv]:+.0f}%")
            if df.vol[sv] >= HYPE_VOL_MULT * vmed:
                p += 1
                w.append(f"vol {df.vol[sv] / vmed:.1f}x median")
            if df.r2[sv] <= r2q:
                p += 1
                w.append(f"R2 vs {BENCH_LBL} {df.r2[sv]:.2f}")
            if runup[sv] >= 2 and fade[sv] <= HYPE_FADE and not ok_txt:
                p += 1
                w.append(f"{100 * (1 - fade[sv]):.0f}% below peak")
        pts.append(p)
        why.append(", ".join(w))
        oks.append(ok_txt)
    df["hype_pts"], df["hype_why"], df["hype_ok"] = pts, why, oks

    def flag(row):
        if row.name in BENCH_EX:
            return "BENCHMARK"
        low_s, low_m = not row.corr_sol >= MIN_CORR_SOL, not row.corr_maj >= MIN_CORR_MAJ
        bad = low_s if CORR_TEST == "sol" else low_m
        if not bad:
            return ""
        why = []
        if low_s and CORR_TEST == "sol":
            why.append(f"corr to SOL {row.corr_sol:.2f}")
        if low_m and CORR_TEST == "majors":
            why.append(f"corr to majors {row.corr_maj:.2f}")
        return "IGNORE: " + ", ".join(why)
    df["flag"] = df.apply(flag, axis=1)
    if HYPE_ACTION == "exclude":                          # strong-hype coins never become shortlist candidates
        strong = df.hype_pts >= HYPE_STRONG
        df.loc[strong & (df.flag == ""), "flag"] = "HYPE: " + df.hype_why
        df.loc[strong & df.flag.str.startswith("IGNORE"), "flag"] = df.flag + "; HYPE"

    # Score: high = long candidate, low = short candidate
    x = df.drop(index=BENCH_EX, errors="ignore")
    df["score"] = np.nan
    df.loc[x.index, "score"] = (x.weeks_up.rank(pct=True) + x.cum_resid.rank(pct=True)
                                + 0.5 * x.last3_up.astype(float) - 0.5 * x.last3_dn.astype(float))
    df = df.sort_values("score", ascending=False, na_position="last")
    df["rank"] = np.where(df.score.notna(), np.arange(1, len(df) + 1), np.nan)

    # Shortlists: top/bottom N among coins that are NOT flagged; all of them start selected
    ok = df[(df.flag == "") & df.score.notna()]
    df["elig_pos"] = np.nan
    df.loc[ok.index, "elig_pos"] = np.arange(1, len(ok) + 1)
    k = min(N_SHORTLIST, len(ok) // 2)
    # Choose by the AVERAGE position over the last AVG_DAYS days (today + earlier reports), not by today's rank alone.
    # Each day's position is turned into a fraction of that day's list so lists of different length are comparable;
    # a day on which the coin was not eligible counts as the middle (0.5).
    m = len(ok)
    today_pos = {sv: i + 1 for i, sv in enumerate(ok.index)}

    def avg_frac(sv, side):
        entries = [(today_pos[sv], m)] + [((lst.index(sv) + 1) if sv in lst else None, len(lst)) for lst in (history or [])]
        vals = [0.5 if i is None else (i if side == "L" else n + 1 - i) / n for i, n in entries]
        return sum(vals) / len(vals)
    fl = {sv: avg_frac(sv, "L") for sv in ok.index}
    fs = {sv: avg_frac(sv, "S") for sv in ok.index}
    long_ids = sorted(ok.index, key=lambda sv: (fl[sv], today_pos[sv]))[:k]
    taken = set(long_ids)
    short_ids = sorted((sv for sv in ok.index if sv not in taken), key=lambda sv: (fs[sv], -today_pos[sv]))[:k]
    cl, cs = ok.loc[long_ids].copy(), ok.loc[short_ids].copy()      # best first
    cl["avg_pos"], cs["avg_pos"] = [fl[sv] * m for sv in long_ids], [fs[sv] * m for sv in short_ids]
    cl["today_pos"], cs["today_pos"] = [today_pos[sv] for sv in long_ids], [m + 1 - today_pos[sv] for sv in short_ids]
    longs, shorts = cl.copy(), cs.copy()
    for side in (longs, shorts):
        side["w"] = side.inv / side.inv.sum()
    # Beta balance: scale the short side so sum(w*beta) matches the long side
    bl, bs = (longs.w * longs.beta30).sum(), (shorts.w * shorts.beta30).sum()
    scale = float(np.clip(bl / bs, 0.5, 2.0)) if bs > 0 and bl > 0 else 1.0
    longs["size"], shorts["size"] = longs.w, shorts.w * scale
    df["pick"] = ""
    df.loc[longs.index, "pick"] = "LONG"
    df.loc[shorts.index, "pick"] = "SHORT"
    return df, longs, shorts, scale, cl, cs


# ----------------------------- Excel -----------------------------
def load_history(now):
    """Earlier days' ordered lists of eligible coins (oldest first), at most AVG_DAYS - 1 of them."""
    today = now.strftime("%Y-%m-%d")
    try:
        with open(next((p for p in [STATE_FILE] + OLD_STATE_FILES if os.path.exists(p)), STATE_FILE)) as f:
            state = json.load(f)
    except Exception:
        return []
    days = [d for d in sorted(state) if d < today and state[d].get("ok")][-(AVG_DAYS - 1):]
    return [state[d]["ok"] for d in days]


def track(df, cl, cs, now, n_hist):
    """Saves today's lists for the coming days. Returns (info, new_state)."""
    today = now.strftime("%Y-%m-%d")
    try:
        with open(next((p for p in [STATE_FILE] + OLD_STATE_FILES if os.path.exists(p)), STATE_FILE)) as f:
            state = json.load(f)
    except Exception:
        state = {}
    state[today] = {"ok": list(df[df.elig_pos.notna()].sort_values("elig_pos").index),
                    "LONG": list(cl.index), "SHORT": list(cs.index)}
    for d in sorted(state)[:-(WF_DAYS + 4)]:
        del state[d]
    return {"days": n_hist}, state


# ----------------------------- Timing sheet -----------------------------
def recurring_events(start, end):
    """Events that repeat on a calendar (added automatically, no need to list them in events.csv)."""
    out = []
    d = start
    while d <= end:
        if d.dayofweek == 4:
            if (d + pd.Timedelta(days=7)).month != d.month:
                out.append(dict(t=d + pd.Timedelta(hours=8), label="Monthly crypto options expiry (last Friday 08:00 UTC)",
                                kind="recurring", allday=False))
            if 15 <= d.day <= 21:
                out.append(dict(t=d + pd.Timedelta(hours=20), kind="recurring", allday=False,
                                label="US quarterly options expiry (quad witching)" if d.month in (3, 6, 9, 12) else "US monthly options expiry"))
        if (d + pd.Timedelta(days=1)).month != d.month:
            out.append(dict(t=d + pd.Timedelta(hours=21), label="Month end (fund rebalancing)", kind="recurring", allday=False))
        d += pd.Timedelta(days=1)
    return out


def load_events(path=EVENTS_FILE):
    out = []
    try:
        with open(path, newline="", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                t = (row.get("utc_time") or "").strip()
                if not t:
                    continue
                allday = len(t) <= 10
                out.append(dict(t=pd.Timestamp(t if not allday else t + " 00:00"), label=(row.get("label") or "").strip(),
                                kind=(row.get("type") or "news").strip().lower(), allday=allday))
    except FileNotFoundError:
        pass
    out += recurring_events(pd.Timestamp.utcnow().tz_localize(None).normalize() - pd.Timedelta(days=110),
                            pd.Timestamp.utcnow().tz_localize(None).normalize() + pd.Timedelta(days=14))
    return out


def walk_forward(H, picks, max_days=14):
    """Point-in-time replay for the Timing sheet. Every earlier daily report saved its LONG and SHORT lists; those lists are held for the
    24 hours AFTER that report (so nothing is known in hindsight), with inverse-volatility weights from the data available then.
    Also the N_OBOS most oversold longs / most overbought shorts of that day's lists (OB/OS at the report time) are held the same 24 hours.
    The 24-hour windows follow each other, so the pieces chain into one hourly series. Returns None if no saved day can be replayed."""
    r = np.log(H).diff()
    n = len(H)
    marks = []
    for p in picks[-max_days:]:
        t0 = p["t"] if p.get("t") is not None else pd.Timestamp(p["date"]) - pd.Timedelta(hours=1)
        i0 = int(H.index.searchsorted(t0, side="right")) - 1
        if i0 < 200 or i0 >= n - 1 or H.index[i0] < t0 - pd.Timedelta(hours=3):
            continue
        marks.append((i0, p))
    parts, parts5, rows = [], [], []
    for k, (i0, p) in enumerate(marks):
        i1 = min(i0 + 24, n - 1)
        if k + 1 < len(marks):
            i1 = min(i1, marks[k + 1][0])
        if i1 <= i0:
            continue
        rv = r.iloc[max(1, i0 - SIZING_DAYS * 24 + 1):i0 + 1].std()
        lc = [x for x in p["long"] if x in H.columns and rv.get(x, np.nan) > 0]
        sc = [x for x in p["short"] if x in H.columns and rv.get(x, np.nan) > 0]
        if len(lc) < N_OBOS or len(sc) < N_OBOS:
            continue
        wt = lambda cs: (1 / rv[cs]) / (1 / rv[cs]).sum()
        sd = rv * np.sqrt(24)                                           # daily volatility, as in the ranking
        px = H.iloc[i0]
        ob = 0.5 * (px / H.iloc[i0 - 6] - 1) / (sd * np.sqrt(6 / 24)) + 0.5 * (px / H.iloc[i0 - 24] - 1) / sd
        lo, hi = list(ob[lc].nsmallest(N_OBOS).index), list(ob[sc].nlargest(N_OBOS).index)
        seg = r.iloc[i0 + 1:i1 + 1]
        s20 = (seg[lc] * wt(lc)).sum(axis=1) - (seg[sc] * wt(sc)).sum(axis=1)
        s5 = (seg[lo] * wt(lo)).sum(axis=1) - (seg[hi] * wt(hi)).sum(axis=1)
        parts.append(s20)
        parts5.append(s5)
        rows.append(dict(t0=H.index[i0], hours=len(seg), bps20=float(s20.sum() * 1e4), bps5=float(s5.sum() * 1e4), lo=lo, hi=hi))
    if not parts:
        return None
    return dict(sp=pd.concat(parts), sp5=pd.concat(parts5), rows=rows)


def timing_analysis(H, longs, shorts, events, picks=None):
    """Long basket / short basket ratio over the last 4 weeks and its time-of-day pattern."""
    r = np.log(H).diff().iloc[-TIMING_HOURS:]
    idx = r.index
    rl = (r[longs.index] * longs["w"]).sum(axis=1)
    rs = (r[shorts.index] * shorts["w"]).sum(axis=1)
    sp = rl - rs                                   # hourly log return of the long/short spread
    mk = r.mean(axis=1)                            # equal-weight market proxy
    ratio = 100 * np.exp(sp.cumsum())
    zs, zm = (sp - sp.mean()) / sp.std(), (mk - mk.mean()) / mk.std()
    spike = (zs.abs() > SPIKE_Z) | (zm.abs() > SPIKE_Z)

    first, last = idx[0], idx[-1]
    ev_mask = pd.Series(False, index=idx)
    ev_rows, upcoming = [], []
    for e in events:
        if e["allday"]:
            day = e["t"].normalize()
            span = (idx >= day) & (idx < day + pd.Timedelta(days=1))
            tf = day
        else:
            tf = e["t"].floor("h")
            span = (idx >= tf - pd.Timedelta(hours=EVENT_BEFORE_H)) & (idx <= tf + pd.Timedelta(hours=EVENT_AFTER_H))
        if e["t"] > last + pd.Timedelta(hours=1):
            if e["t"] <= last + pd.Timedelta(days=7):
                upcoming.append(e)
            continue
        if not span.any():
            continue
        ev_mask |= span
        msp, mmk = sp[span].sum(), mk[span].sum()
        w48 = (idx >= tf) & (idx < tf + pd.Timedelta(hours=EVENT_WINDOW_H))
        nh = int(w48.sum())
        s48 = float(sp[w48].sum())
        z48 = (s48 - sp.mean() * nh) / (sp.std() * np.sqrt(nh)) if nh >= 3 else np.nan
        ev_rows.append(dict(t=e["t"], label=e["label"], kind=e["kind"], allday=e["allday"], spread_bps=msp * 1e4,
                            market_pct=mmk * 100, z=abs(mmk) / (mk.std() * np.sqrt(span.sum())),
                            bps48=s48 * 1e4, vs_norm=(s48 - sp.mean() * nh) * 1e4, z48=z48, hours48=nh, neg=bool(z48 <= EVENT_NEG_Z)))
    calm = ~(ev_mask | spike)

    def table(mask, series):
        """Hour-of-day statistics on the Seoul clock (hour and day of week both in KST), drift removed, split weekday / Saturday / Sunday."""
        ki = series.index + pd.Timedelta(hours=KST)
        dk = np.where(ki.dayofweek < 5, "wd", np.where(ki.dayofweek == 5, "sat", "sun"))
        kh = np.asarray(ki.hour)
        m = mask.reindex(series.index).fillna(True).values.astype(bool)
        b = (series * 1e4)[m]
        d = pd.DataFrame({"bps": (b - b.mean()).values, "h": kh[m], "k": dk[m]})
        res = {}
        for key, *_ in DAY_KEYS:
            g = d[d.k == key].groupby("h")["bps"].agg(["mean", "std", "count"]).reindex(range(24))        # index = KST hour
            g["t"] = g["mean"] / (g["std"] / np.sqrt(g["count"]))
            c = g["mean"].fillna(0).cumsum()
            g["path"] = c - c.mean()
            res[key] = g
        return res

    # Extra basket: the N_OBOS most oversold longs and N_OBOS most overbought shorts of today's lists (OB/OS score), inverse-vol weights within each side
    lo5 = list(longs.nsmallest(N_OBOS, "ob").index)
    hi5 = list(shorts.nlargest(N_OBOS, "ob").index)
    w5l, w5s = longs.loc[lo5, "w"] / longs.loc[lo5, "w"].sum(), shorts.loc[hi5, "w"] / shorts.loc[hi5, "w"].sum()
    sp5 = (r[lo5] * w5l).sum(axis=1) - (r[hi5] * w5s).sum(axis=1)
    sp_rows = []
    for t in idx[spike]:
        sp_rows.append(dict(t=t, market_pct=mk[t] * 100, spread_bps=sp[t] * 1e4, z=max(abs(zs[t]), abs(zm[t])),
                            near_event=bool(ev_mask[t])))
    sp_rows.sort(key=lambda x: -x["z"])
    def counts(ser):
        kd = pd.Series((ser.index + pd.Timedelta(hours=KST)).normalize()).drop_duplicates()
        return dict(wd=int((kd.dt.dayofweek < 5).sum()), sat=int((kd.dt.dayofweek == 5).sum()), sun=int((kd.dt.dayofweek == 6).sum()))
    out = dict(idx=idx, ratio=ratio, ev_rows=ev_rows, upcoming=upcoming, spikes=sp_rows[:15], n_spikes=int(spike.sum()), spike_t=list(idx[spike]),
               calm=table(calm, sp), n_calm=int(calm.sum()), n_hours=len(idx), days=counts(sp), lo5=lo5, hi5=hi5,
               first=first, last=last, nl=len(longs), ns=len(shorts), wf=None)
    # Walk-forward set (no hindsight): each earlier day's own saved lists held for the 24 hours after that day's report
    wf = walk_forward(H, picks or [], max_days=WF_DAYS)
    if wf is not None:
        rows = wf["rows"]
        sp5w = wf["sp5"][wf["sp5"].index >= rows[-WF5_DAYS]["t0"]] if len(rows) > WF5_DAYS else wf["sp5"]
        out["wf"] = dict(ratio=100 * np.exp(wf["sp"].cumsum()), calm=table(calm, wf["sp"]), calm5=table(calm, sp5w), rows=rows,
                         days=counts(wf["sp"]), n5=min(len(rows), WF5_DAYS), days5=counts(sp5w))
    return out


def timing_charts(T):
    """Returns a dict of PNG buffers: ratio / tod (replay of today's lists), w_ratio / w_tod / w5 (walk-forward; None until earlier days are saved)."""
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
    except Exception:
        return None
    kst = lambda ts: ts + pd.Timedelta(hours=KST)

    def png(fig):
        buf = io.BytesIO()
        fig.savefig(buf, format="png")
        plt.close(fig)
        buf.seek(0)
        return buf

    def ratio_chart(rt, title, with_spikes=True):
        x = kst(rt.index)
        fig, ax = plt.subplots(figsize=(9, 4.2), dpi=100)
        for d in pd.Series(x.normalize()).drop_duplicates():                  # grey = Saturday and Sunday on the Seoul clock
            if d.dayofweek >= 5:
                ax.axvspan(d, d + pd.Timedelta(days=1), color="#eeeeee", zorder=0)
        ax.plot(x, rt.values, color="#1f4e79", lw=1.2, zorder=3)
        ymax = float(rt.max())
        for e in T["ev_rows"]:
            if not (rt.index[0] <= e["t"] <= rt.index[-1]):
                continue
            xt = kst(e["t"])
            ax.axvline(xt, color={"scheduled": "#c00000", "recurring": "#7030a0"}.get(e["kind"], "#ed7d31"), ls=":" if e["kind"] == "recurring" else "--", lw=0.9, zorder=2)
            ax.text(xt, ymax, " " + e["label"].split(";")[0][:26], rotation=90, va="top", ha="right", fontsize=6, color="#555555")
        st = pd.DatetimeIndex([t for t in T["spike_t"] if t in rt.index])
        if with_spikes and len(st):
            ax.scatter(kst(st), rt.reindex(st).values, marker="v", s=18, color="#c00000", zorder=4)
        ax.set_xlim(x[0], x[-1])
        ax.set_title(title, fontsize=7.5)
        ax.tick_params(labelsize=7)
        fig.autofmt_xdate()
        fig.tight_layout()
        return png(fig)

    def tod_chart(tab, title):
        fig, ax = plt.subplots(figsize=(9, 4.0), dpi=100)
        for key, name, col, mk in DAY_KEYS:
            ax.plot(range(24), tab[key]["path"].values, marker=mk, ms=3, lw=1.3, color=col, label=name)
        ax.axhline(0, color="#999999", lw=0.8)
        ax.set_xticks(range(24))
        ax.set_xlabel("Hour of day, KST", fontsize=8)
        ax.set_ylabel("typical ratio path, bps (0 = day average)", fontsize=8)
        ax.set_title(title, fontsize=7.5)
        ax.tick_params(labelsize=7)
        ax.legend(fontsize=7)
        fig.tight_layout()
        return png(fig)

    W = T["wf"]
    c = dict(
        ratio=ratio_chart(T["ratio"], "Long basket / short basket (start = 100), TODAY's lists replayed over the last 4 weeks, KST. Grey = Sat/Sun, dashed = listed event (purple dotted = recurring), red triangle = volatility spike"),
        tod=tod_chart(T["calm"], "Typical path of the ratio through the day (KST), TODAY's lists replayed over 4 weeks, calm hours only (low = better time to open, high = better time to close)"),
        w_ratio=None, w_tod=None, w5=None)
    if W is not None:
        c["w_ratio"] = ratio_chart(W["ratio"], f"WALK-FORWARD: each day's own saved lists held for the next 24h, chained (start = 100), last {len(W['rows'])} days, KST. Grey = Sat/Sun, dashed = listed event", with_spikes=False)
        c["w_tod"] = tod_chart(W["calm"], "WALK-FORWARD typical path of the ratio through the day (KST), each day's own saved lists held 24h after the report, calm hours only")
        c["w5"] = tod_chart(W["calm5"], f"WALK-FORWARD {N_OBOS} most oversold longs vs {N_OBOS} most overbought shorts of each day's lists (OB/OS at report time), held 24h, last {W['n5']} days, KST")
    return c


def write_timing_sheet(book, T, charts):
    from openpyxl.drawing.image import Image as XlImage
    ws = book.create_sheet("Timing", 2)
    bold = Font(bold=True)
    head = PatternFill("solid", fgColor="D9D9D9")
    kst_h = lambda h: (h + KST) % 24
    def lh(tab):
        o = {}
        for key, *_ in DAY_KEYS:
            p = tab[key]["path"]
            o[key] = (int(p.idxmin()), int(p.idxmax()), float(p.max() - p.min()))
        return (f"weekdays lowest around {o['wd'][0]:02d}:00 KST, highest {o['wd'][1]:02d}:00, swing {o['wd'][2]:.0f} bps; "
                f"Saturday lowest {o['sat'][0]:02d}:00, highest {o['sat'][1]:02d}:00, swing {o['sat'][2]:.0f} bps; "
                f"Sunday lowest {o['sun'][0]:02d}:00, highest {o['sun'][1]:02d}:00, swing {o['sun'][2]:.0f} bps")
    W = T["wf"]
    nt = int((T["calm"]["wd"]["t"].abs() > 2).sum())
    now_ratio = float(T["ratio"].iloc[-1])
    G = T.get("gauge")
    if G:
        zt = {"big": "BIG zone: conditions are comfortable, you could go up to 15x15",
              "normal": "NORMAL zone: stay with 10x10", "small": "SMALL zone: markets are weak, 5x5 or no position"}[G["cur_zone"]]
        gl = f"MARKET COMFORT GAUGE now {G['cur']:+.2f} (as of {G['cur_time'] + pd.Timedelta(hours=KST):%m-%d %H:%M} KST). {zt}."
    else:
        gl = "Market comfort gauge unavailable today: " + str(T.get("gauge_err", "no data")) + "."
    d_ = T["days"]
    lines = [
        gl,
        f"TOP charts = REPLAY: the {T['nl']} LONG and {T['ns']} SHORT coins of today's Portfolio sheet (chosen by 7-day average position, inverse-volatility weights within each side) replayed over the last 4 weeks, hourly, start = 100. Now {now_ratio:.1f}. "
        "Hindsight: the coins were picked partly on these weeks, so the ratio drifts up by construction.",
        "Replay typical path (calm hours, Seoul clock): " + lh(T["calm"]) + ".",
        ((f"WALK-FORWARD charts (below the replay): on each of the last {len(W['rows'])} days (up to {WF_DAYS}) the LONG and SHORT lists saved by that day's report were held for the 24 hours after it, "
          f"weights = inverse volatility from that day's data, pieces chained, start = 100. Now {float(W['ratio'].iloc[-1]):.1f}. No hindsight in the lists. "
          f"Typical path: " + lh(W["calm"]) + ".")
         if W else
         f"WALK-FORWARD charts: none yet. They need earlier daily reports saved in {STATE_FILE} (the file keeps the last {WF_DAYS + 4} days); they appear automatically as the days accumulate."),
    ] + ([
        f"{N_OBOS} vs {N_OBOS} OB/OS pair (last {W['n5']} days only): each day the {N_OBOS} most oversold longs and {N_OBOS} most overbought shorts of that day's lists (OB/OS at report time) held 24h. "
        "Typical path: " + lh(W["calm5"]) + f". Today's pair has no outcome yet: long {', '.join(T['lo5'])} / short {', '.join(T['hi5'])}.",
    ] if W else []) + [
        f"Calm hours (replay series) = all hours minus listed events (1h before to 3h after, or the whole day for news items) and detected volatility spikes: {T['n_calm']} of {T['n_hours']} hours kept, {T['n_spikes']} spike hours found.",
        f"Noise check (replay): {nt} of 24 weekday hours have |t| above 2 (about 1 would appear by chance). Replay has {d_['wd']} weekdays, {d_['sat']} Saturdays and {d_['sun']} Sundays"
        + (f"; walk-forward has {W['days']['wd']}, {W['days']['sat']} and {W['days']['sun']} (5 vs 5: {W['days5']['wd']}, {W['days5']['sat']} and {W['days5']['sun']})" if W else "")
        + ". Each Saturday or Sunday hour rests on only that many observations, so those lines are mostly noise.",
        "Caution 1 (replay): the baskets are TODAY's picks, chosen partly on these same weeks, so the ratio drifts up by construction. The time-of-day path removes the drift but is still tuned on the same 4 weeks.",
        "Caution 2: 'close when high, open when low' only works if the ratio mean-reverts. If it trends (your momentum idea), closing at a high can cut winners.",
        "Caution 3 (walk-forward): every 24h window starts at the report (about 09:00 KST), so the time-of-day path mixes the move since that hour with the clock effect, and a few big days dominate when only a few weeks are saved.",
        "Events come from events.csv (scheduled releases and news days); edit that file to add or fix items. Spikes are found automatically, so check news for the red-triangle hours.",
    ]
    for ln in lines:
        ws.append([ln])
    ws["A1"].font = Font(bold=True, color="C00000" if G and G["cur_zone"] == "small" else "1F4E79")
    ws.column_dimensions["A"].width = 12
    for col in "BCDEFGHIJKLMN":
        ws.column_dimensions[col].width = 14
    row = len(lines) + 2
    def put(buf, title=None):
        nonlocal row
        if title:
            ws.cell(row, 1, title).font = bold
            row += 1
        im = XlImage(buf)
        ws.add_image(im, f"A{row}")
        row += -(-im.height // 20) + 2                                  # default row height = 20 px

    if charts is not None:
        put(charts["ratio"])
        gimg = gauge_chart(G, T["ev_rows"]) if G else None
        if gimg is not None:
            put(gimg)
        put(charts["tod"])
        if charts["w_ratio"] is not None:
            put(charts["w_ratio"], "WALK-FORWARD (no hindsight): each day's own saved lists, held for the 24 hours after that day's report")
            put(charts["w_tod"])
            put(charts["w5"])
        else:
            ws.cell(row, 1, "Walk-forward charts appear here once earlier daily reports are saved.").font = bold
            row += 2
    else:
        ws.cell(row, 1, "(charts unavailable: install matplotlib)")
        row += 2

    def table(title, cols, rows, fmts=None):
        nonlocal row
        ws.cell(row, 1, title).font = bold
        row += 1
        for j, c in enumerate(cols, 1):
            cell = ws.cell(row, j, c)
            cell.font, cell.fill = bold, head
        row += 1
        for r in rows:
            for j, v in enumerate(r, 1):
                cell = ws.cell(row, j, v if not (isinstance(v, float) and np.isnan(v)) else "")
                if fmts and fmts[j - 1] and isinstance(v, (int, float)):
                    cell.number_format = fmts[j - 1]
            row += 1
        row += 1

    fl = "0.0"

    def hour_rows(tab):
        out = []
        for k in range(24):
            c = [tab[key].loc[k] for key, *_ in DAY_KEYS]
            out.append([k, (k - KST) % 24] + [v for g in c for v in (g["mean"], g["t"], g["path"])] + [int(g["count"]) for g in c])
        return out
    hcols = ["KST hour", "UTC hour", "Weekday avg bps", "Weekday t", "Weekday path", "Sat avg bps", "Sat t", "Sat path", "Sun avg bps", "Sun t", "Sun path",
             "Weekday hours", "Sat hours", "Sun hours"]
    hfmt = [None, None] + [fl] * 9 + [None] * 3
    table("REPLAY, hour of day (calm hours): average hourly move of the ratio in bps versus the average hour, trend removed (t = mean divided by its standard error; |t| above 2 is worth noticing). Days and hours on the Seoul clock.",
          hcols, hour_rows(T["calm"]), hfmt)
    if W:
        table("WALK-FORWARD, hour of day (same measures)", hcols, hour_rows(W["calm"]), hfmt)
        table(f"WALK-FORWARD {N_OBOS} vs {N_OBOS} OB/OS pair, hour of day (last {W['n5']} days)", hcols, hour_rows(W["calm5"]), hfmt)
    if W:
        table(f"Walk-forward days (up to {WF_DAYS}): the lists saved at each report held for the next 24h, in bps of the long/short spread (positive = longs beat shorts)",
              ["Report (KST)", "Hours covered", f"{T['nl']} vs {T['ns']} bps", f"{N_OBOS} vs {N_OBOS} bps", f"{N_OBOS} oversold longs (OB/OS)", f"{N_OBOS} overbought shorts (OB/OS)"],
              [[(w["t0"] + pd.Timedelta(hours=KST + 1)).strftime("%m-%d %H:%M"), w["hours"], w["bps20"], w["bps5"], ", ".join(w["lo"]), ", ".join(w["hi"])]
               for w in W["rows"]], [None, None, fl, fl, None, None])
    negs = [e for e in sorted(T["ev_rows"], key=lambda x: x["t"], reverse=True) if e["neg"]]
    table(f"Significantly NEGATIVE events only: the {T['nl']}x{T['ns']} spread over the next {EVENT_WINDOW_H}h was at least {abs(EVENT_NEG_Z):.0f} standard deviations below normal. "
          f"{len(T['ev_rows']) - len(negs)} other listed events in the window were neutral or positive and are hidden.",
          ["UTC time", "KST time", "Event", "Type", "Spread bps next 48h", "vs normal bps", "Size (std devs)", "Spread bps in event window", "Market %"],
          [[e["t"].strftime("%Y-%m-%d" + ("" if e["allday"] else " %H:%M")),
            (e["t"] + pd.Timedelta(hours=KST)).strftime("%m-%d" + ("" if e["allday"] else " %H:%M")), e["label"], e["kind"],
            e["bps48"], e["vs_norm"], e["z48"], e["spread_bps"], e["market_pct"]] for e in negs] or [["none in this window", "", "", "", "", "", "", "", ""]],
          [None, None, None, None, fl, fl, "0.0", fl, "0.00"])
    if G:
        kst_ = lambda t_: (t_ + pd.Timedelta(hours=KST)).strftime("%m-%d %H:%M")
        table("Shocks feeding the gauge right now (strongest first). A shock acts for 24h after the US close, fades by 48h, then rebounds until 120h",
              ["Market", "US close (KST)", "Drop / rise %", "Effect on gauge now", "Hours since close"],
              [[x["name"], kst_(x["close"]) + (" (session running)" if x["live"] else ""), x["drop"], x["now"], x["age"]] for x in G["active"]]
              or [["none: no move beyond its dead band in the last 5 days", "", "", "", ""]], [None, None, "0.0", "+0.00;-0.00", "0"])
        table(f"Gauge settings (edit the constants at the top of ranker.py): Nasdaq moves under {NASDAQ_DEAD:.0f}% are ignored and {NASDAQ_REF:.0f}% = 1.0 unit; single stocks under {STOCK_DEAD:.0f}% ignored, {STOCK_REF:.0f}% = 1.0; "
              f"VIX rises under {VIX_DEAD:.0f}% ignored, {VIX_REF:.0f}% = 1.0. Impact grows exponentially above the dead band. Weights: Nasdaq {W_NASDAQ}, each stock {W_STOCK}, VIX {W_VIX}; "
              f"stablecoin supply growth (5 days, versus its own normal pace) adds up to +-{STABLE_SCALE * 2:.1f} all the time.",
              ["Part", "Source", "Contribution now"], [[k, G["src"].get(k, ""), v] for k, v in G["parts"].items()], [None, None, "+0.00;-0.00"])
        if G["hist"]:
            hl_ = [[kst_(x["close"]), x["drop"], x["w0"], x["w1"], x["w2"], x["w3"]] for x in G["hist"]]
            ok_ = [[x[k] for x in G["hist"] if x[k] == x[k]] for k in ("w0", "w1", "w2", "w3")]
            hl_.append(["Average", "", *[float(np.mean(v)) if v else np.nan for v in ok_]])
            table("Check on the shape: spread move (bps vs a normal day) in four 24h blocks after each past Nasdaq drop beyond the dead band. The gauge assumes negative in blocks 1-2 and positive in blocks 3-4 - with so few cases this is only a sanity check",
                  ["US close (KST)", "Drop %", "0-24h", "24-48h", "48-72h", "72-96h"], hl_, [None, "0.0", fl, fl, fl, fl])
        zl = [[{"big": "BIG (gauge above %+.1f)" % GAUGE_BIG, "normal": "NORMAL", "small": "SMALL (gauge below %+.1f)" % GAUGE_SMALL}[t["zone"]],
               t["n"], t["f24"], t["f48"]] for t in G["table"]]
        zl.append(["All hours", G["all"]["n"], G["all"]["f24"], G["all"]["f48"]])
        table("Does the gauge help? Average move of the long/short spread AFTER the gauge was in each zone, in bps. If BIG is not clearly above SMALL, the gauge is not adding information",
              ["Zone at the time", "Hours", "Spread bps, next 24h", "Spread bps, next 48h"], zl, [None, None, fl, fl])
        if G["corr"]:
            c = G["corr"]
            ws.cell(row, 1, f"Gauge at the US close vs the spread over the next 48h: r = {c['r']:+.2f}, p = {c['p']:.3f}, {c['n']} days (p counts only half the days as independent). "
                            f"Positive and p below 0.05 would support the idea; otherwise treat the gauge as a hypothesis. Zone limits are my starting guesses.")
            row += 2
        if G["notes"]:
            ws.cell(row, 1, "Data notes: " + "; ".join(G["notes"][:4]))
            row += 2
    table(f"Largest detected volatility hours (market or ratio move beyond {SPIKE_Z:.0f} standard deviations). Check news for these.",
          ["UTC time", "KST time", "Market %", "Spread bps", "Size (std devs)", "Inside a listed event?"],
          [[x["t"].strftime("%Y-%m-%d %H:%M"), (x["t"] + pd.Timedelta(hours=KST)).strftime("%m-%d %H:%M"), x["market_pct"],
            x["spread_bps"], x["z"], "yes" if x["near_event"] else "NO - unexplained"] for x in T["spikes"]],
          [None, None, "0.00", fl, "0.0", None])
    if T["upcoming"]:
        table("Upcoming listed events (next 7 days)", ["UTC time", "KST time", "Event", "Type"],
              [[e["t"].strftime("%Y-%m-%d %H:%M"), (e["t"] + pd.Timedelta(hours=KST)).strftime("%m-%d %H:%M"), e["label"], e["kind"]]
               for e in T["upcoming"]])
    ws.column_dimensions["C"].width = 44


# ----------------------------- Timing sheet: market comfort gauge -----------------------------
# The gauge is 0 when nothing is wrong. Stock-market shocks push it down (exponentially with size, ignored below a dead band),
# the push lasts about 48 hours and then turns into a mild rebound. Stablecoin supply adds or removes a little all the time.
GAUGE_BIG, GAUGE_SMALL = 0.25, -0.5      # zones: at or above BIG = go bigger, at or below SMALL = go small (starting guesses)
NASDAQ_DEAD, NASDAQ_REF = 1.0, 2.0       # percent: Nasdaq drops below 1% are ignored; a 2% drop counts as 1.0 gauge unit
STOCK_DEAD, STOCK_REF = 4.0, 10.0        # each single big-tech stock (after-hours and pre-market included): ignored below 4%, a 10% drop = 1.0 unit
VIX_DEAD, VIX_REF = 15.0, 30.0           # percent RISE of the VIX: ignored below 15%, a 30% rise = 1.0 unit
EXP_SHAPE = 1.0                          # impact grows e^1.0 = 2.7x across the span from dead band to reference size, and keeps accelerating
IMPACT_CAP = 6.0
W_NASDAQ, W_STOCK, W_VIX = 1.0, 0.7, 0.25   # your judgement: VIX far below Nasdaq
STABLE_SCALE = 0.25                      # gauge units per standard deviation of stablecoin-supply growth versus its own normal pace (continuous effect, max +-0.5)
FULL_H, FADE_END_H = 24, 48              # a shock acts fully for 24h after the US close, fades to zero by 48h ...
REBOUND, REB_END_H = 0.4, 120            # ... then a rebound of 40% of the shock size, peaking around 84h and gone by 120h
GAUGE_SMOOTH_H = 120
EVENT_NEG_Z = -2.0        # a listed event counts as "significantly negative" if the spread's next-48h move is this many std devs below normal
EVENT_WINDOW_H = 48
ALERT_RECENT_H = 72       # only events from the last 72 hours go into the message title


def fetch_yahoo_hourly(sym, prepost=True):
    last = None
    for rng_ in ("3mo", "60d"):
        try:
            r = _http(f"https://query1.finance.yahoo.com/v8/finance/chart/{sym}",
                      {"interval": "1h", "range": rng_, "includePrePost": "true" if prepost else "false"}, {"User-Agent": "Mozilla/5.0"})
            res = r.json()["chart"]["result"][0]
            ts = pd.to_datetime(np.array(res["timestamp"], dtype=float), unit="s")
            s = pd.Series(res["indicators"]["quote"][0]["close"], index=ts, dtype=float).dropna()
            s = s[~s.index.duplicated(keep="last")].sort_index()
            if len(s) > 100:
                return s
        except Exception as e:
            last = e
    raise RuntimeError(last or "too little data")


def load_hourly_levels():
    out, notes = {}, []
    for key, sym, pp in [("NASDAQ", "QQQ", True), ("VIX", "^VIX", False)] + [(k, k, True) for k in BIGTECH]:
        try:
            out[key] = fetch_yahoo_hourly(sym, pp)
        except Exception as e:
            notes.append(f"{sym} hourly: {e}")
        time.sleep(0.4)
    mem = [k for k in BIGTECH if k in out]
    if len(mem) >= 4:
        lr = pd.concat([np.log(out[k]).diff() for k in mem], axis=1).mean(axis=1).dropna()
        out["BIGTECH"] = np.exp(lr.cumsum())
    return out, notes


def _on_grid(s, idx):
    """Value of series s (indexed by the time each value became known) at the END of every hourly label in idx."""
    ends = idx + pd.Timedelta(hours=1)
    u = s.reindex(s.index.union(ends)).ffill()
    return pd.Series(u.reindex(ends).values, index=idx)


def _closes(first, last):
    """US regular-session closes (16:00 New York time, weekdays) as naive UTC timestamps."""
    days = pd.bdate_range(first.normalize() - pd.Timedelta(days=3), last.normalize() + pd.Timedelta(days=3))
    return (days + pd.Timedelta(hours=16)).tz_localize("America/New_York").tz_convert("UTC").tz_localize(None)


def _impact(drop, dead, ref):
    drop = np.asarray(drop, float)
    k = EXP_SHAPE / (ref - dead)
    a = (np.exp(k * np.clip(drop - dead, 0, None)) - 1) / (np.exp(k * (ref - dead)) - 1)
    return np.where(drop > dead, np.minimum(a, IMPACT_CAP), 0.0)


def _kernel(age_h):
    """1 while the shock acts, falling to 0 at 48h, then negative (= a rebound) until 120h."""
    a = np.asarray(age_h, float)
    main = np.where(a <= FULL_H, 1.0, np.where(a <= FADE_END_H, 1 - (a - FULL_H) / (FADE_END_H - FULL_H), 0.0))
    reb = np.where((a > FADE_END_H) & (a <= REB_END_H), REBOUND * np.sin(np.pi * (a - FADE_END_H) / (REB_END_H - FADE_END_H)), 0.0)
    return main - reb


def shock_effect(Lg, closes, dead, ref, weight, up_is_bad=False):
    """Gauge contribution of one price series known at the END of each hour: zero normally, negative after a shock, positive during the rebound."""
    ends = Lg.index
    t = ends.values
    ln = np.log(Lg.values.astype(float))
    eff = np.zeros(len(Lg))
    lc = []
    for c in closes:
        v = Lg.asof(c) if c >= ends[0] else np.nan
        lc.append(np.log(v) if v == v and v > 0 else np.nan)
    pct = (lambda a, b: (np.exp(a - b) - 1) * 100) if up_is_bad else (lambda a, b: -(np.exp(a - b) - 1) * 100)    # stress move in percent
    sessions, last_t = [], ends[-1]
    for j in range(1, len(closes)):
        c0, c1, v0, v1 = closes[j - 1], closes[j], lc[j - 1], lc[j]
        if np.isnan(v0) or c0 > last_t:
            continue
        m = (t > np.datetime64(c0)) & (t <= np.datetime64(c1))
        if m.any():                                            # the session is running (the drop so far counts at once, pre-market included)
            drop = pct(ln[m], v0)
            eff[m] -= weight * _impact(drop, dead, ref)
            if c1 > last_t and drop[-1] > dead:
                sessions.append(dict(close=c1, drop=float(drop[-1]), a=float(_impact(drop[-1], dead, ref)) * weight, age=0.0, live=True))
        if np.isnan(v1) or c1 > last_t:
            continue
        d1 = float(pct(v1, v0))
        a1 = float(_impact(d1, dead, ref))
        if a1 <= 0:
            continue
        age = (t - np.datetime64(c1)) / np.timedelta64(1, "h")
        m2 = (age > 0) & (age <= REB_END_H)
        eff[m2] -= weight * a1 * _kernel(age[m2])
        sessions.append(dict(close=c1, drop=d1, a=a1 * weight, age=float((last_t - c1) / pd.Timedelta(hours=1)), live=False))
    return pd.Series(eff, index=ends), sessions


def gauge_analysis(H, longs, shorts, lv, D=None):
    from scipy import stats
    r = np.log(H).diff().iloc[1:]
    idx = r.index
    ends = idx + pd.Timedelta(hours=1)
    sp = (r[longs.index] * longs["w"]).sum(axis=1) - (r[shorts.index] * shorts["w"]).sum(axis=1)
    hl, hnotes = load_hourly_levels()
    closes = _closes(idx[0], idx[-1])
    src, parts, shocks = {}, {}, []

    def level(key):
        if key in hl:
            s = pd.Series(hl[key].values, index=hl[key].index + pd.Timedelta(hours=1))      # an hourly bar is known when it ends
            how = "hourly"
        elif key in lv:
            d = lv[key]
            cl_ = (d.index + pd.Timedelta(hours=16)).tz_localize("America/New_York", nonexistent="shift_forward").tz_convert("UTC").tz_localize(None)
            s = pd.Series(d.values, index=cl_)
            how = "daily only"
        else:
            return None, None
        s = s[~s.index.duplicated(keep="last")].sort_index()
        u = s.reindex(s.index.union(ends)).ffill()
        return pd.Series(u.reindex(ends).values, index=ends).dropna(), how

    spec = [("NASDAQ", "Nasdaq", NASDAQ_DEAD, NASDAQ_REF, W_NASDAQ, False), ("VIX", "VIX", VIX_DEAD, VIX_REF, W_VIX, True)] + \
           [(k, k, STOCK_DEAD, STOCK_REF, W_STOCK, False) for k in BIGTECH]
    for key, name, dead, ref, wt, upbad in spec:
        Lg, how = level(key)
        if Lg is None or len(Lg) < 50:
            continue
        eff, ses = shock_effect(Lg, closes, dead, ref, wt, upbad)
        parts[name] = eff.reindex(ends).fillna(0.0)
        src[name] = how
        for x in ses:
            x["name"], x["kind"] = name, key
        shocks += ses
    if "STABLE" in lv:
        s = pd.Series(lv["STABLE"].values, index=lv["STABLE"].index + pd.Timedelta(hours=24))
        u = s.reindex(s.index.union(ends)).ffill()
        lvl = pd.Series(u.reindex(ends).values, index=ends)
        ln = np.log(lvl.where(lvl > 0))
        ch5 = ln.diff(GAUGE_SMOOTH_H)                           # growth over the last 5 days
        sig = ch5.std()
        if sig and not np.isnan(sig):                           # compared with its own normal pace, so steady growth does not push the gauge up forever
            z = (ch5 - ch5.rolling(720, min_periods=240).mean()) / sig
            parts["Stablecoins"] = (STABLE_SCALE * z.clip(-2, 2)).fillna(0.0)
            src["Stablecoins"] = "daily"
    if "Nasdaq" not in parts and "Stablecoins" not in parts:
        raise RuntimeError("no market data could be loaded: " + "; ".join(hnotes[:3]))
    g = pd.Series(0.0, index=ends)
    for v in parts.values():
        g = g + v
    g.index = idx                                              # label = start of the hour, value known at its end
    f24 = sp.rolling(24).sum().shift(-24) * 1e4
    f48 = sp.rolling(48).sum().shift(-48) * 1e4
    zone = pd.Series(np.where(g >= GAUGE_BIG, "big", np.where(g <= GAUGE_SMALL, "small", "normal")), index=idx)
    ok = f48.notna()
    tab = []
    for z in ("big", "normal", "small"):
        m = ok & (zone == z)
        tab.append(dict(zone=z, n=int(m.sum()), f24=float(f24[m].mean()) if m.any() else np.nan, f48=float(f48[m].mean()) if m.any() else np.nan))
    allm = dict(f24=float(f24[ok].mean()), f48=float(f48[ok].mean()), n=int(ok.sum()))
    d = pd.DataFrame({"g": g, "f": f48})[idx.hour == 20].dropna()
    corr = None
    if len(d) >= 20 and d["g"].std() > 0:
        rr = stats.pearsonr(d["g"], d["f"])[0]
        ne = len(d) / 2
        t_ = rr * np.sqrt((ne - 2) / max(1e-12, 1 - rr ** 2))
        corr = dict(r=float(rr), p=float(2 * stats.t.sf(abs(t_), ne - 2)), n=len(d))
    hist = []                                                  # what the spread did after each past Nasdaq drop beyond the dead band
    mean_h = sp.mean() * 1e4
    for x in sorted([x for x in shocks if x["kind"] == "NASDAQ" and not x["live"]], key=lambda x: x["close"]):
        row = dict(close=x["close"], drop=x["drop"])
        for k, w0 in enumerate((0, 24, 48, 72)):
            msk = ((ends > x["close"] + pd.Timedelta(hours=w0)) & (ends <= x["close"] + pd.Timedelta(hours=w0 + 24)))
            row[f"w{k}"] = float(sp.values[msk].sum() * 1e4 - mean_h * msk.sum()) if msk.sum() == 24 else np.nan
        hist.append(row)
    active = []
    for x in shocks:
        if x["live"] or x["age"] <= REB_END_H:
            kk = float(_kernel(0.0 if x["live"] else x["age"]))
            active.append(dict(name=x["name"], close=x["close"], drop=x["drop"], now=-x["a"] * kk, age=x["age"], live=x["live"]))
    active.sort(key=lambda x: x["now"])
    mk_ = {}
    for x in shocks:
        if x["a"] >= 0.25:
            mk_.setdefault(x["close"], []).append(f"{x['name']} {x['drop']:+.0f}%" if x["kind"] == "VIX" else f"{x['name']} {-x['drop']:.1f}%")
    marks = [(t_, ", ".join(v)) for t_, v in sorted(mk_.items())]
    last_g = g.dropna()
    cur = float(last_g.iloc[-1])
    cur_zone = "big" if cur >= GAUGE_BIG else "small" if cur <= GAUGE_SMALL else "normal"
    hrs = idx[-TIMING_HOURS:]
    ratio = (np.exp(sp.reindex(hrs).cumsum()) - 1) * 100
    return dict(g=g, hours=hrs, g_chart=g.reindex(hrs), ratio=ratio, src=src, notes=hnotes, cur=cur, cur_zone=cur_zone,
                cur_time=idx[-1] + pd.Timedelta(hours=1), table=tab, all=allm, corr=corr, hist=hist, active=active, marks=marks,
                parts={k: float(v.iloc[-1]) for k, v in parts.items()})


def gauge_chart(G, ev_rows):
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
    except Exception:
        return None
    ink, mute = "#52514e", "#898781"
    kst = lambda ts: ts + pd.Timedelta(hours=KST)
    x = kst(G["hours"] + pd.Timedelta(hours=1))
    fig, (a1, a2) = plt.subplots(2, 1, figsize=(9, 5.6), dpi=100, sharex=True, gridspec_kw={"height_ratios": [3, 2]})
    gv = G["g_chart"].values
    lo, hi = max(-6.0, np.nanmin([np.nanmin(gv), -1.5])), np.nanmax([np.nanmax(gv), 1.0])
    a1.axhspan(GAUGE_BIG, hi + 0.3, color="#cde2fb", zorder=0)
    a1.axhspan(lo - 0.3, GAUGE_SMALL, color="#f6d2c3", zorder=0)
    a1.plot(x, np.clip(gv, lo, None), color="#2a78d6", lw=1.6, zorder=3)
    for tm, lab in G["marks"]:
        tk = kst(tm)
        if x[0] <= tk <= x[-1]:
            a1.axvline(tk, color="#b9b7ae", ls=":", lw=0.8, zorder=1)
            a1.text(tk, lo - 0.2, " " + lab, rotation=90, va="bottom", ha="right", fontsize=6.5, color="#52514e")
    a1.axhline(0, color="#c3c2b7", lw=0.8)
    a1.set_xlim(x[0], x[-1])
    a1.set_ylim(lo - 0.3, hi + 0.3)
    a1.text(x[0], hi + 0.2, " go bigger (up to 15x15)", fontsize=7.5, color=ink, va="top")
    a1.text(x[0], GAUGE_SMALL - 0.05, " go small: 5x5 or stand aside", fontsize=7.5, color=ink, va="top")
    a1.text(x[0], 0.05, " normal 10x10", fontsize=7.5, color=ink, va="bottom")
    for e in ev_rows:
        if e.get("neg"):
            for a in (a1, a2):
                a.axvline(kst(e["t"]), color="#d03b3b", ls="--", lw=0.9, zorder=2)
            a1.text(kst(e["t"]), hi + 0.25, " " + e["label"].split(";")[0][:28], rotation=90, va="top", ha="right", fontsize=6.5, color=ink)
    a1.set_title("Market comfort gauge, hourly, KST (higher = more comfortable). Dotted = shock close, dashed red = negative event", fontsize=8, color=ink)
    a2.plot(x, G["ratio"].values, color="#2a78d6", lw=1.3)
    a2.axhline(0, color="#c3c2b7", lw=0.8)
    a2.set_ylabel("long/short ratio, % since start", fontsize=7.5, color=ink)
    a1.set_ylabel("gauge units (2% Nasdaq drop = 1)", fontsize=7.5, color=ink)
    for a in (a1, a2):
        a.grid(axis="y", color="#e1e0d9", lw=0.6)
        a.tick_params(labelsize=7, colors=mute)
        for s_ in ("top", "right"):
            a.spines[s_].set_visible(False)
    fig.autofmt_xdate()
    fig.tight_layout()
    b = io.BytesIO()
    fig.savefig(b, format="png")
    plt.close(fig)
    b.seek(0)
    return b


def bad_news_lines(timing, drivers):
    """Significantly negative listed events of the last 72 hours and fresh stock-market shocks, as lines for the top of the Telegram message."""
    out = []
    if timing:
        for e in sorted(timing["ev_rows"], key=lambda x: x["t"], reverse=True):
            if e.get("neg") and e["t"] >= timing["last"] - pd.Timedelta(hours=ALERT_RECENT_H):
                out.append(f"- {e['label']} ({e['t']:%Y-%m-%d}): your 10x10 spread was {e['vs_norm']:+.0f} bps versus normal over the following 48h ({e['z48']:.1f} std devs).")
    if drivers and drivers[0]:
        out += ["- " + a for a in drivers[0]["alerts"]]
    return out


# ----------------------------- Drivers sheet: do outside markets explain the long/short ratio? -----------------------------
BIGTECH = ["GOOGL", "TSLA", "AAPL", "MSFT", "NVDA", "AMZN", "META"]
# key, label, bad: +1 means a RISE is stress for coins (VIX, yields, dollar), -1 means a FALL is stress (stocks, stablecoin supply)
DRIVERS = [("NASDAQ", "Nasdaq Composite", -1), ("BIGTECH", "Big-tech index (equal-weight 7)", -1), ("GOOGL", "Alphabet", -1),
           ("TSLA", "Tesla", -1), ("NVDA", "Nvidia", -1), ("VIX", "VIX (stock-market fear)", 1),
           ("US10Y", "US 10-year yield", 1), ("DXY", "US dollar index", 1), ("STABLE", "Stablecoin supply (DefiLlama)", -1)]
YAHOO = {"NASDAQ": "^IXIC", "GOOGL": "GOOGL", "TSLA": "TSLA", "AAPL": "AAPL", "MSFT": "MSFT", "NVDA": "NVDA", "AMZN": "AMZN",
         "META": "META", "VIX": "^VIX", "US10Y": "^TNX", "DXY": "DX-Y.NYB"}
STOOQ = {"NASDAQ": "^ndq", "GOOGL": "googl.us", "TSLA": "tsla.us", "AAPL": "aapl.us", "MSFT": "msft.us", "NVDA": "nvda.us",
         "AMZN": "amzn.us", "META": "meta.us"}       # backup source if Yahoo refuses
DRIVER_LAGS = [0, 1, 2, 3, "cum"]
LAG_NAME = {0: "same day", 1: "+1 day", 2: "+2 days", 3: "+3 days", "cum": "sum +1..+3"}
SHOCK_SIGMA = 2.0        # a "shock" day = stress move of at least this many standard deviations ...
SHOCK_MIN = 0.015        # ... and, for stocks, at least 1.5%
SHOCK_KEYS = ("NASDAQ", "BIGTECH", "GOOGL", "TSLA", "VIX")
MIN_DRIVER_DAYS = 25


def _http(url, params=None, headers=None, tries=3):
    last = None
    for k in range(tries):
        try:
            r = requests.get(url, params=params, headers=headers, timeout=30)
            if r.status_code == 429:
                time.sleep(5 * (k + 1))
                last = "HTTP 429 (rate limited)"
                continue
            r.raise_for_status()
            return r
        except Exception as e:
            last = e
            time.sleep(2)
    raise RuntimeError(f"{url.split('?')[0]}: {last}")


def fetch_yahoo(sym):
    r = _http(f"https://query1.finance.yahoo.com/v8/finance/chart/{sym}", {"interval": "1d", "range": "6mo"},
              {"User-Agent": "Mozilla/5.0"})
    res = r.json()["chart"]["result"][0]
    off = res.get("meta", {}).get("gmtoffset", 0) or 0
    ts = pd.to_datetime(np.array(res["timestamp"], dtype=float) + off, unit="s").normalize()
    s = pd.Series(res["indicators"]["quote"][0]["close"], index=ts, dtype=float).dropna()
    return s[~s.index.duplicated(keep="last")].sort_index()


def fetch_stooq(sym):
    r = _http("https://stooq.com/q/d/l/", {"s": sym, "i": "d"}, {"User-Agent": "Mozilla/5.0"})
    d = pd.read_csv(io.StringIO(r.text))
    if "Date" not in d.columns or "Close" not in d.columns:
        raise RuntimeError("no data")
    s = pd.Series(d["Close"].values, index=pd.to_datetime(d["Date"]), dtype=float).dropna()
    return s.sort_index().iloc[-140:]


def fetch_stable():
    r = _http("https://stablecoins.llama.fi/stablecoincharts/all")
    rows = r.json()
    s = pd.Series([float(x["totalCirculatingUSD"]["peggedUSD"]) for x in rows],
                  index=pd.to_datetime([int(x["date"]) for x in rows], unit="s").normalize(), dtype=float)
    return s[~s.index.duplicated(keep="last")].sort_index().iloc[-140:]


def load_driver_levels():
    lv, notes = {}, []
    for key in YAHOO:
        s = None
        try:
            s = fetch_yahoo(YAHOO[key])
        except Exception as e:
            if key in STOOQ:
                try:
                    s = fetch_stooq(STOOQ[key])
                    notes.append(f"{key}: Yahoo failed, used Stooq.")
                except Exception as e2:
                    notes.append(f"{key}: Yahoo failed ({e}); Stooq failed ({e2})")
            else:
                notes.append(f"{key}: {e}")
        if s is not None and len(s) > MIN_DRIVER_DAYS:
            lv[key] = s
        time.sleep(0.4)
    try:
        lv["STABLE"] = fetch_stable()
    except Exception as e:
        notes.append(f"STABLE: {e}")
    mem = [k for k in BIGTECH if k in lv]
    if len(mem) >= 4:
        lr = pd.concat([np.log(lv[k]).diff() for k in mem], axis=1).mean(axis=1).dropna()
        lv["BIGTECH"] = 100 * np.exp(lr.cumsum())
        notes.append("Big-tech index = equal-weight average of daily returns of: " + ", ".join(mem) + ".")
    else:
        notes.append("Big-tech index not built (fewer than 4 of the 7 stocks could be loaded).")
    return lv, notes


def _bh(ps):
    ps = np.asarray(ps, float)
    m = len(ps)
    order = np.argsort(ps)
    q = np.empty(m)
    prev = 1.0
    for k in range(m - 1, -1, -1):
        i = order[k]
        prev = min(prev, ps[i] * m / (k + 1))
        q[i] = prev
    return q


def drivers_analysis(H, longs, shorts, lv, events, now=None):
    from scipy import stats
    r = np.log(H).diff().iloc[1:]
    rl = (r[longs.index] * longs["w"]).sum(axis=1)
    rs = (r[shorts.index] * shorts["w"]).sum(axis=1)
    cum = pd.DataFrame({"spread": (rl - rs).cumsum(), "market": r.mean(axis=1).cumsum()})
    cal = lv["NASDAQ"] if "NASDAQ" in lv else lv.get("BIGTECH")
    if cal is None:
        raise RuntimeError("no stock-market calendar (Nasdaq and big-tech both missing)")
    t20 = pd.Timedelta(hours=20)                     # label 20:00 = the hour that ends at 21:00 UTC (US close, roughly)
    grid = pd.DatetimeIndex([d for d in cal.index if cum.index[0] <= d + t20 <= cum.index[-1]])
    if len(grid) < MIN_DRIVER_DAYS + 5:
        raise RuntimeError(f"only {len(grid)} trading days overlap")
    at = cum.asof(grid + t20)
    at.index = grid
    C = at.diff().iloc[1:]                           # coin move over each trading-day interval (close to close)
    S = {}
    for key, label, bad in DRIVERS:
        if key not in lv:
            continue
        v = lv[key].reindex(grid, method="ffill")
        mv = v.diff() if key == "US10Y" else np.log(v).diff()
        S[key] = (bad * mv).iloc[1:]                 # positive = stress
    rows = []
    for key in S:
        e = S[key]
        for tgt in ("spread", "market"):
            c = C[tgt]
            for lag in DRIVER_LAGS:
                if lag == "cum":
                    y, win = c.shift(-1) + c.shift(-2) + c.shift(-3), 3
                else:
                    y, win = c.shift(-lag), 1
                d = pd.concat([e, y], axis=1).dropna()
                if len(d) < 20 or d.iloc[:, 0].std() == 0:
                    continue
                rr, p = stats.pearsonr(d.iloc[:, 0], d.iloc[:, 1])
                rho = stats.spearmanr(d.iloc[:, 0], d.iloc[:, 1])[0]
                if win > 1:                          # overlapping 3-day sums: count only n/3 independent observations
                    ne = len(d) / win
                    tt = rr * np.sqrt((ne - 2) / max(1e-12, 1 - rr ** 2))
                    p = 2 * stats.t.sf(abs(tt), ne - 2)
                rows.append(dict(key=key, tgt=tgt, lag=lag, r=rr, rho=rho, p=p, n=len(d)))
    if rows:
        for x, q in zip(rows, _bh([x["p"] for x in rows])):
            x["q"] = q
    # ---- event study: what happened to the coins after a stress day in a driver ----
    names = {k: lab for k, lab, _ in DRIVERS}
    ev_rows, pooled, alerts, alert_short = [], [], [], []
    rng = np.random.default_rng(0)
    fwd = {t: C[t].rolling(4).sum().shift(-3) for t in ("spread", "market")}     # move over day 0..+3
    for key in SHOCK_KEYS:
        if key not in S:
            continue
        e = S[key]
        thr = SHOCK_SIGMA * e.std()
        if key != "VIX":
            thr = max(thr, SHOCK_MIN)
        pos_list = []
        for i in np.where(e.values >= thr)[0]:
            if pos_list and i - pos_list[-1] <= 1:
                continue
            pos_list.append(i)
        for i in pos_list:
            sp = C["spread"].values
            vals = [sp[i + k] * 1e4 if i + k < len(sp) else np.nan for k in range(4)]
            ev_rows.append(dict(key=key, name=names[key], date=e.index[i], move=e.values[i] * 100, d0=vals[0], d1=vals[1],
                                d2=vals[2], d3=vals[3], tot=(np.nansum(vals) if not np.isnan(vals).any() else np.nan),
                                mkt=np.nansum([C["market"].values[i + k] * 1e4 if i + k < len(sp) else np.nan for k in range(4)])))
        full = [i for i in pos_list if not np.isnan(fwd["spread"].values[i])]
        if full:
            res = dict(key=key, name=names[key], n=len(full), thr=thr * 100)
            for t in ("spread", "market"):
                pool = fwd[t].dropna().values
                obs = np.mean([fwd[t].values[i] for i in full])
                draws = np.array([rng.choice(pool, size=len(full), replace=False if len(pool) > len(full) else True).mean()
                                  for _ in range(5000)])
                res[t + "_bps"] = obs * 1e4
                res[t + "_p"] = float((draws <= obs).mean())          # chance a random set of days looks this bad
                res[t + "_all_bps"] = pool.mean() * 1e4
            pooled.append(res)
        recent = [i for i in range(max(0, len(e) - 2), len(e)) if e.values[i] >= thr]
        if recent:
            i = recent[-1]
            pr = next((p for p in pooled if p["key"] == key), None)
            bad_ = next(bd for k_, _, bd in DRIVERS if k_ == key)
            msg = (f"{names[key]} {'rose' if bad_ > 0 else 'fell'} {e.values[i] * 100:.1f}% on {e.index[i]:%Y-%m-%d} (shock threshold {thr * 100:.1f}%). ")
            if pr:
                msg += (f"Your own history ({pr['n']} earlier case(s)): the long/short spread averaged {pr['spread_bps']:+.0f} bps over day 0..+3 "
                        f"(all days: {pr['spread_all_bps']:+.0f}); chance this is luck {pr['spread_p']:.0%}.")
            else:
                msg += "No earlier case in this window with a full 4-day follow-up, so no evidence either way."
            alerts.append(msg)
            alert_short.append(f"{names[key]} {'+' if bad_ > 0 else '-'}{e.values[i] * 100:.1f}%")
    status = []
    for key, label, bad in DRIVERS:
        if key in S:
            e = S[key]
            status.append(dict(name=label, date=e.index[-1], d1=e.values[-1] * 100 * bad, d3=e.iloc[-3:].sum() * 100 * bad, key=key))
    # ---- chart data: last 4 weeks, everything as % change from the start of the window ----
    hrs = r.index[-TIMING_HOURS:]
    spr = (rl - rs).reindex(hrs).cumsum()
    chart = dict(hours=hrs, spread=(np.exp(spr) - 1) * 100, series={})
    g4 = grid[grid + t20 >= hrs[0]]
    for key in ("NASDAQ", "BIGTECH", "TSLA", "GOOGL"):
        if key in lv and len(g4) > 2:
            v = lv[key].reindex(g4, method="ffill")
            chart["series"][key] = ((v / v.iloc[0] - 1) * 100)
    ev = []
    lo, hi = hrs[0], hrs[-1]
    for x in events:
        if lo <= x["t"] <= hi:
            ev.append(x)
    return dict(rows=rows, ev_rows=ev_rows, pooled=pooled, alerts=alerts, alert_short=alert_short, status=status, chart=chart, events=ev,
                n_days=len(C), first=grid[0], last=grid[-1], names=names, keys=list(S))


def drivers_charts(D):
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        from matplotlib.colors import LinearSegmentedColormap
    except Exception:
        return None, None
    ink, mute, grid_c = "#52514e", "#898781", "#e1e0d9"
    kst = lambda ts: ts + pd.Timedelta(hours=KST)
    ch = D["chart"]
    fig, ax = plt.subplots(figsize=(9, 4.6), dpi=100)
    ax.set_facecolor("#fcfcfb")
    for e in D["events"]:
        ax.axvline(kst(e["t"]), color="#c3c2b7" if e["kind"] != "scheduled" else "#e3b4b4", ls=":", lw=0.9, zorder=1)
    ax.plot(kst(ch["hours"]), ch["spread"].values, color="#2a78d6", lw=1.8, label="Long/short ratio (hourly)", zorder=4)
    spec = [("NASDAQ", "Nasdaq", "#eb6834"), ("BIGTECH", "Big-tech index", "#1baf7a"), ("TSLA", "Tesla", "#6250d6"),
            ("GOOGL", "Alphabet", "#e87ba4")]
    for key, lab, col in spec:
        if key in ch["series"]:
            s = ch["series"][key]
            ax.plot(kst(s.index + pd.Timedelta(hours=21)), s.values, color=col, lw=1.1, marker="o", ms=3, label=lab, zorder=3)
    ax.axhline(0, color="#c3c2b7", lw=0.8)
    ax.set_ylabel("% change since start of window", fontsize=8, color=ink)
    ax.set_title("Long/short ratio against US stocks, last 4 weeks, KST. All lines are % change from the first day. Dotted = listed events", fontsize=8, color=ink)
    ax.grid(axis="y", color=grid_c, lw=0.6)
    for s_ in ("top", "right"):
        ax.spines[s_].set_visible(False)
    ax.tick_params(labelsize=7, colors=mute)
    ax.legend(fontsize=7, frameon=False, loc="upper left", ncol=5)
    fig.autofmt_xdate()
    fig.tight_layout()
    b1 = io.BytesIO()
    fig.savefig(b1, format="png")
    plt.close(fig)

    keys = [k for k in D["keys"] if any(x["key"] == k and x["tgt"] == "spread" for x in D["rows"])]
    if not keys:
        b1.seek(0)
        return b1, None
    M = np.full((len(keys), len(DRIVER_LAGS)), np.nan)
    stars = [[""] * len(DRIVER_LAGS) for _ in keys]
    for x in D["rows"]:
        if x["tgt"] == "spread" and x["key"] in keys:
            i, j = keys.index(x["key"]), DRIVER_LAGS.index(x["lag"])
            M[i, j] = x["r"]
            stars[i][j] = "**" if x["q"] < 0.10 else "*" if x["p"] < 0.05 else ""
    cmap = LinearSegmentedColormap.from_list("div", ["#eb6834", "#f0efec", "#2a78d6"])
    fig, ax = plt.subplots(figsize=(9, 0.55 * len(keys) + 1.9), dpi=100)
    ax.imshow(M, cmap=cmap, vmin=-0.5, vmax=0.5, aspect="auto")
    for i in range(len(keys)):
        for j in range(len(DRIVER_LAGS)):
            if not np.isnan(M[i, j]):
                ax.text(j, i, f"{M[i, j]:+.2f}{stars[i][j]}", ha="center", va="center", fontsize=8, color="#0b0b0b",
                        fontweight="bold" if stars[i][j] else "normal")
    ax.set_xticks(range(len(DRIVER_LAGS)))
    ax.set_xticklabels([LAG_NAME[l] for l in DRIVER_LAGS], fontsize=8)
    ax.set_yticks(range(len(keys)))
    ax.set_yticklabels([D["names"][k] for k in keys], fontsize=8)
    ax.set_title("Correlation of a stress day in each market with the long/short spread N trading days later.\n"
                 "Orange (negative) = stress is followed by the spread FALLING. * p<0.05 on its own, ** survives the many-tests correction (q<0.10)",
                 fontsize=7.5, color=ink)
    fig.tight_layout()
    b2 = io.BytesIO()
    fig.savefig(b2, format="png")
    plt.close(fig)
    b1.seek(0)
    b2.seek(0)
    return b1, b2


def write_drivers_sheet(book, D, charts, notes):
    from openpyxl.drawing.image import Image as XlImage
    ws = book.create_sheet("Drivers", 3)
    bold, head = Font(bold=True), PatternFill("solid", fgColor="D9D9D9")
    red = PatternFill("solid", fgColor="F8CBAD")
    lines = []
    if D is None:
        lines.append("Drivers study could not run today: " + (notes[0] if notes else "no data."))
        lines += ["Details:"] + [f"  {n}" for n in notes[1:12]]
        for ln in lines:
            ws.append([ln])
        ws.column_dimensions["A"].width = 120
        return
    for a in D["alerts"]:
        lines.append("SHOCK WINDOW - " + a)
    if not D["alerts"]:
        lines.append("No stress shock in the last 2 trading days in Nasdaq, big-tech, Tesla, Alphabet or VIX.")
    lines += [
        f"Study window: {D['n_days']} trading days, {D['first']:%Y-%m-%d} to {D['last']:%Y-%m-%d}. Each day = US close to US close (about 21:00 UTC). 'Spread' = long basket minus short basket (the same baskets as the Timing sheet), 'market' = equal-weight average of all coins loaded.",
        "Stress is signed so that positive always means bad for coins: stocks falling, stablecoin supply shrinking, VIX / yields / dollar rising.",
        "How to read the heat map: a correlation of -0.3 on '+2 days' means that when stress was higher than usual, the spread tended to be lower two trading days later. Roughly |r| above 0.25 can reach p<0.05 with this many days.",
        "p-values are two-sided. 'q' corrects for the dozens of combinations tested at once (Benjamini-Hochberg); with many tests several p<0.05 appear by chance, so trust q first. 'sum +1..+3' overlaps itself, so its p-value counts only n/3 independent days.",
        "Caution 1: the baskets are TODAY's picks, chosen on the same weeks, so the spread has an upward drift by construction (it cancels in correlations of daily changes, but selection can still colour the result). Only the latest 7 days are truly out of sample.",
        "Caution 2: this only says stress and the spread moved together in these weeks. It cannot say WHY, and ~60 trading days contain only a handful of real shocks. Treat 'wait 2 days after a shock' as a hypothesis until it survives more months.",
        "Event study below: a 'shock' = stress at least 2 standard deviations and, for stocks, at least 1.5%. The p-value is the share of 5000 random same-size sets of days whose average 4-day spread move was as bad or worse.",
    ] + [f"Data note: {n}" for n in notes[:6]]
    for ln in lines:
        ws.append([ln])
    ws.column_dimensions["A"].width = 30
    for col in "BCDEFGHIJ":
        ws.column_dimensions[col].width = 14
    ws.column_dimensions["C"].width = 18
    for r_ in range(1, 1 + len(D["alerts"])):
        ws.cell(r_, 1).font = Font(bold=True, color="C00000")
    row = len(lines) + 2
    if charts[0] is not None:
        ws.add_image(XlImage(charts[0]), f"A{row}")
        row += 25
        if charts[1] is not None:
            ws.add_image(XlImage(charts[1]), f"A{row}")
            row += int(0.55 * len([1 for k in D["keys"]]) * 5 + 14)
    else:
        ws.cell(row, 1, "(charts unavailable: install matplotlib)")
        row += 2

    def table(title, cols, rows_, fmts=None, hot=None):
        nonlocal row
        ws.cell(row, 1, title).font = bold
        row += 1
        for j, c in enumerate(cols, 1):
            cell = ws.cell(row, j, c)
            cell.font, cell.fill = bold, head
        row += 1
        for rr in rows_:
            for j, v in enumerate(rr, 1):
                cell = ws.cell(row, j, v if not (isinstance(v, float) and np.isnan(v)) else "")
                if fmts and fmts[j - 1] and isinstance(v, (int, float)):
                    cell.number_format = fmts[j - 1]
                if hot and hot(rr) and j == 1:
                    cell.fill = red
            row += 1
        row += 1

    table("Latest stress move in each market (positive = bad for coins)", ["Market", "Last date", "Last day %", "Last 3 days %"],
          [[s["name"], f"{s['date']:%Y-%m-%d}", s["d1"], s["d3"]] for s in D["status"]], [None, None, "0.00", "0.00"])
    rws = sorted([x for x in D["rows"] if x["tgt"] == "spread"], key=lambda x: x["p"])
    table("All tests on the long/short spread, smallest p first (r: negative = stress then falling spread)",
          ["Market", "Days later", "r", "Spearman rho", "p-value", "q (corrected)", "Days used"],
          [[D["names"][x["key"]], LAG_NAME[x["lag"]], x["r"], x["rho"], x["p"], x["q"], x["n"]] for x in rws],
          [None, None, "0.00", "0.00", "0.000", "0.000", None], hot=lambda rr: rr[5] < 0.10)
    rwm = sorted([x for x in D["rows"] if x["tgt"] == "market"], key=lambda x: x["p"])[:15]
    table("For comparison: the same tests on the whole coin market (equal-weight), 15 smallest p",
          ["Market", "Days later", "r", "Spearman rho", "p-value", "q (corrected)", "Days used"],
          [[D["names"][x["key"]], LAG_NAME[x["lag"]], x["r"], x["rho"], x["p"], x["q"], x["n"]] for x in rwm],
          [None, None, "0.00", "0.00", "0.000", "0.000", None])
    table("Event study, summary: average move over day 0..+3 after a stress day (bps = 0.01%)",
          ["Market", "Shocks", "Threshold %", "Spread avg bps", "Spread, all days bps", "Chance it is luck", "Whole market avg bps", "Market chance luck"],
          [[p["name"], p["n"], p["thr"], p["spread_bps"], p["spread_all_bps"], p["spread_p"], p["market_bps"], p["market_p"]] for p in D["pooled"]],
          [None, None, "0.0", "0", "0", "0%", "0", "0%"])
    table("Event study, every shock day (spread move in bps on the shock day and the next 3 trading days; blank = not happened yet)",
          ["Market", "Shock date", "Stress move %", "Day 0", "Day +1", "Day +2", "Day +3", "Total 0..+3", "Whole market total"],
          [[x["name"], f"{x['date']:%Y-%m-%d}", x["move"], x["d0"], x["d1"], x["d2"], x["d3"], x["tot"], x["mkt"]] for x in D["ev_rows"]],
          [None, None, "0.0", "0", "0", "0", "0", "0", "0"])
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 16


def write_excel(path, df, longs, shorts, scale, skipped, stamp, cl, cs, tr, timing=None, drivers=None):
    def trend(row):
        return "all up" if row.last3_up else "all down" if row.last3_dn else "mixed"

    df = df.sort_values("score", ascending=False, na_position="last", kind="stable")     # Ranking sheet order: highest score first
    med = df.inv.median()
    avg_by = {**{x: round(v, 1) for x, v in cl.avg_pos.items()}, **{x: round(v, 1) for x, v in cs.avg_pos.items()}}
    chg_lbl = [avg_by.get(sym, "") if p in ("LONG", "SHORT") else "" for sym, p in zip(df.index, df.pick)]
    wk = f"Weeks beating {BENCH_LBL} (of {WINDOW_WEEKS})"
    l3 = f"Last {CONSIST_WEEKS} wks vs {BENCH_LBL}"
    maj_mode = BENCH_MODE == "majors"
    rk_cols = [   # key, header, values: the order here is the column order on the Ranking sheet
        ("csol", "Maj r" if maj_mode else "SOL r", df.corr_b.values), ("sym", "Symbol", list(df.index)), ("rank", "Rank", df["rank"].values),
        ("pick", "Pick", df.pick.values),
        ("chg", "7d avg pos", chg_lbl), ("flag", "Flag", df.flag.values), ("score", "Score", df.score.values),
        ("wk", wk, df.weeks_up.values), ("l3", l3, df.apply(trend, axis=1).values),
        ("cum", f"Cum. residual vs {BENCH_LBL}", df.cum_resid.values), ("beta", "Beta", df.beta.values),
        ("r2", "R2 (hourly)", df.r2.values),
        ("cmaj", "SOL r (1m)" if maj_mode else "Maj r (1m)", (df.corr_sol if maj_mode else df.corr_maj).values),
        ("vol", "Daily vol (30d)", df.vol.values), ("cap", "Cap rank", df.cap_rank.values),
        ("c6", "6h change", df.chg6.values), ("c24", "24h change", df.chg24.values),
        ("ob", "OB/OS score", df.ob.values),
        ("rel", "Rel. size (1.0 = typical coin)", (df.inv / med).values),
        ("inp", "Portfolio", [""] * len(df)),
        ("hype", "Hype", [(f"{p}: {w}" if p >= HYPE_WATCH else "") + (" | " if p >= HYPE_WATCH and o else "") + (("ok: " + o) if o else "")
                          for p, w, o in zip(df.hype_pts, df.hype_why, df.hype_ok)]),
    ]
    out = pd.DataFrame({h: v for _, h, v in rk_cols})
    out[wk] = out[wk].astype(object)
    out.loc[out.Symbol.isin(BENCH_EX), [l3, wk]] = ""
    RK = {k: get_column_letter(i + 1) for i, (k, _, _) in enumerate(rk_cols)}

    pf_cols = [   # key, header: column order on the Portfolio sheet
        ("sym", "Symbol"), ("side", "Side"), ("pos", "Pos in side"), ("chg", "7d avg pos"), ("pos1", "1-day pos"),
        ("c6", "6h change"), ("c24", "24h change"), ("ob", "OB/OS score"), ("hype", "Hype"), ("inc", "Include"),
        ("beta", "Beta (30d)"), ("vol", "Daily vol (30d)"),
        ("size", "Size % of side"), ("rel", "Rel. size (1.0 = typical coin)"),
        ("amt", "Amount (USDT)"),
    ]
    PK = {k: get_column_letter(i + 1) for i, (k, _) in enumerate(pf_cols)}
    BL = get_column_letter(len(pf_cols) + 2)         # column of the budget / multiplier block

    def side_table(name, d, posmap):
        n_ = len(d)
        return pd.DataFrame({
            "Symbol": list(d.index), "Side": [name] * n_, "Pos in side": [posmap[x] for x in d.index],
            "7d avg pos": [round(v, 1) for v in d.avg_pos.values], "1-day pos": d.today_pos.values,
            "6h change": d.chg6.values, "24h change": d.chg24.values, "OB/OS score": d.ob.values, "Hype": [p if p >= HYPE_WATCH else ("ok" if o else "") for p, o in zip(d.hype_pts, d.hype_ok)],
            "Include": ["x"] * n_,
            "Beta (30d)": d.beta30.values, "Daily vol (30d)": d.vol.values,
            "Size % of side": [""] * n_, "Rel. size (1.0 = typical coin)": (d.inv / med).values,
            "Amount (USDT)": [""] * n_,
        })
    # rows: longs most oversold first (best to buy), shorts most overbought first (best to sell)
    pos_l, pos_s = {x: i + 1 for i, x in enumerate(cl.index)}, {x: i + 1 for i, x in enumerate(cs.index)}
    port = pd.concat([side_table("LONG", cl.sort_values("ob"), pos_l),
                      side_table("SHORT", cs.sort_values("ob", ascending=False), pos_s)])

    cmp_note = (f"Coins are CHOSEN by 7-day average position, not by today's rank: each day a coin's position in the list is turned into a "
                f"fraction of that day's list, averaged over today plus up to {AVG_DAYS - 1} earlier reports ({tr['days']} earlier report(s) available now; "
                "a day when the coin was not eligible counts as the middle). The 20 best averages go on the Portfolio sheet. "
                "7d avg pos = that average in today's list scale, 1-day pos = today's own position, Pos in side = rank by the average. "
                "Shorts are counted from the bottom of the list.")
    notes = [
        f"Generated {stamp} UTC. Script version: {VERSION}",
        f"Window: {WINDOW_WEEKS} rolling 7-day weeks ending 7 days before run time (latest 7 days excluded).",
        (f"Benchmark = {'MAJORS: the equal-weight average hourly return of ' + ', '.join(MAJORS) + ' with the coin itself left out' if BENCH_MODE == 'majors' else 'SOL alone (SOL is shown as BENCHMARK and is not a candidate)'}. "
         "Switch it for one run in GitHub > Actions > Run workflow (bench_mode), or permanently with BENCH_MODE at the top of ranker.py. Beta, residuals, the score, the 30-day beta behind the short multiplier and the Hype correlation tests all use it."),
        f"Residual = coin return minus beta x {BENCH_LBL} return, beta from hourly regression over the window.",
        f"Score = rank of weeks beating {BENCH_LBL} + rank of cumulative residual, +0.5 if last {CONSIST_WEEKS} weeks all up, -0.5 if all down.",
        "High score = long candidate, low score = short candidate.",
        ("WHAT-IF RUN: settings came from the Run workflow form, so today's lists were NOT saved to the 7-day history. " if WHATIF_RUN else "")
        + f"IGNORE test used for this file: '{CORR_TEST}' (SOL floor {MIN_CORR_SOL}, majors floor {MIN_CORR_MAJ}). Change it for one run in GitHub > Actions > Run workflow (form with the same names), or permanently in the settings at the top of ranker.py. The shortlists skip IGNORE coins. The IGNORE test and the benchmark for beta / score are separate settings.",
        "'IGNORE test' sheet = a what-if copy: type another mode or cutoffs in its yellow cells and it recounts instantly. It does not change the picks; to apply the numbers, run the workflow with them.",
        "Majors basket = equal-weight average of the 1-minute returns of " + ", ".join(MAJORS) + f" (the coin itself left out). Maj r = how closely a coin's {MINUTE_DAYS}-day, minute-by-minute moves follow that basket.",
        "IGNORE count if you changed the test (of " + str(int((~df.index.isin(BENCH_EX)).sum())) + " coins): " + "; ".join(
            f"{lab} {int(m.sum())}" for lab, m in [
                ("SOL r at 0.45", ~(df.corr_sol >= 0.45) & ~df.index.isin(BENCH_EX)), ("SOL r at 0.35", ~(df.corr_sol >= 0.35) & ~df.index.isin(BENCH_EX)),
                ("SOL r at 0.25", ~(df.corr_sol >= 0.25) & ~df.index.isin(BENCH_EX)),
                ("Maj r at 0.45", ~(df.corr_maj >= 0.45) & ~df.index.isin(BENCH_EX)), ("Maj r at 0.35", ~(df.corr_maj >= 0.35) & ~df.index.isin(BENCH_EX)),
                ("Maj r at 0.25", ~(df.corr_maj >= 0.25) & ~df.index.isin(BENCH_EX))]) + ".",
        "6h / 24h change = latest closed hourly price vs 6 / 24 hours earlier.",
        cmp_note,
        f"Portfolio sheet = shortlist of {N_SHORTLIST} per side, all marked x. Delete the x in the Include column to skip a coin; sizes and the short multiplier recalculate.",
        "Portfolio rows are sorted by OB/OS score = 50% x 6h change / (typical 6h move) + 50% x 24h change / (typical 24h move), typical move from the 30-day volatility. "
        "Positive = overbought, negative = oversold. Longs list the most oversold first, shorts the most overbought first. Pos in side is still the momentum rank.",
        f"Sizing: inverse of 30-day {'daily volatility' if SIZE_BY == 'vol' else 'beta (floored at 0.3)'} within each side (each side sums to 100%).",
        f"Short multiplier = long-side blended beta / short-side blended beta (both vs {BENCH_LBL}), capped 0.5 to 2. Script's value with all picks kept: {scale:.2f}x.",
        f"Amount (USDT), two ways (position size, not margin). Way 1: type a Long budget in Portfolio {BL}2; long side totals it, short side totals budget x multiplier. "
        f"Way 2: type in {BL}5 how many USDT a coin with Rel. size 1.0 should get; every long coin then gets Rel. size x that number, and the short side keeps the beta-balancing multiplier on top. {BL}5 filled = Way 2, empty = Way 1.",
        "The first (frozen) column of the Ranking sheet = 1-minute correlation to the active benchmark (Maj r = majors basket, SOL r = SOL); the other correlation is shown further right. The IGNORE test uses SOL r when CORR_TEST is sol, Maj r when it is majors. The Ranking sheet's Portfolio column just displays what you chose on the Portfolio sheet.",
        "Drivers sheet = how the long/short ratio moved against Nasdaq, a big-tech index, Tesla, Alphabet, Nvidia, VIX, US 10y yield, dollar index and stablecoin supply (daily data from Yahoo Finance and DefiLlama, about 60 trading days). It lists correlations with p-values at lags of 0 to 3 trading days and an event study of stock shocks. If those sites refuse the connection the sheet says so and everything else still works.",
        f"Hype = pump / manipulation-prone pattern points from price alone over the 12-week window: run-up from a low of {100 * (HYPE_RUNUP - 1):.0f}%+ (2 points), a 24h spike of {100 * HYPE_SPIKE:.0f}%+, volatility {HYPE_VOL_MULT:.0f}x the median coin, "
        f"R2 versus {BENCH_LBL} in the lowest {100 * HYPE_DECOUPLED_PCT:.0f}% of coins, and a fade to {100 * HYPE_FADE:.0f}% of the peak after a 2x run-up (1 point each). "
        f"{HYPE_WATCH}+ points = peach shading (still eligible), {HYPE_STRONG}+ = purple and, with HYPE_ACTION = '{HYPE_ACTION}', kept out of the shortlists. Rough screen only: it has no volume, exchange or market-maker data and will sometimes flag a real rally.",
        f"Sound run-ups (green 'ok' in the Hype column, not penalised): a big run-up is judged sound if the coin's hourly correlation to {BENCH_LBL} on the way up stayed at least {HYPE_SOUND_FLOOR} and within {HYPE_SOUND_TOL} of its correlation before the boom, "
        f"or if the boom (price first up {100 * (HYPE_RUNUP - 1):.0f}% from its low) is at least {HYPE_SETTLE_DAYS} days old and in the last {HYPE_SETTLE_DAYS} days the coin kept that correlation, stayed above {int(100 * HYPE_FADE)}% of its peak and had no {int(100 * HYPE_SPIKE)}% day. {int((df.hype_ok != '').sum())} coin(s) this run.",
        f"Hype this run: {int((df.hype_pts >= HYPE_STRONG).sum())} strong ({', '.join(df.index[df.hype_pts >= HYPE_STRONG][:25]) or 'none'}), "
        f"{int(((df.hype_pts >= HYPE_WATCH) & (df.hype_pts < HYPE_STRONG)).sum())} watch. If the strong list looks too long, raise HYPE_RUNUP or HYPE_STRONG at the top of ranker.py.",
        "Skipped: " + ("; ".join(skipped) if skipped else "none"),
    ]

    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        out.to_excel(xw, sheet_name="Ranking", index=False)
        port.to_excel(xw, sheet_name="Portfolio", index=False)
        pd.DataFrame({"Notes": notes}).to_excel(xw, sheet_name="Notes", index=False)

        fills = {k: PatternFill("solid", fgColor=c) for k, c in
                 dict(LONG="C6EFCE", SHORT="FFC7CE", IGNORE="D9D9D9", BENCH="DDEBF7", NEW="FFEB9C", HYPE="E4D5F2", WATCH="FCE4D6", OKAY="E2EFDA").items()}
        ws, wp = xw.sheets["Ranking"], xw.sheets["Portfolio"]

        def fmt_cells(sheet, keys_idx, pct, num2, money, new_col):
            for row in sheet.iter_rows(min_row=2, max_col=keys_idx):
                for c in row:
                    if c.column in pct:
                        c.number_format = "0.0%"
                    elif c.column in num2:
                        c.number_format = "0.00"
                    elif c.column in money:
                        c.number_format = "#,##0.0"
                row[new_col - 1].number_format = "0.0"

        # ---- Portfolio: editable picks (Include column), live sizes ----
        n = len(port) + 1
        R = lambda k: f"${PK[k]}$2:${PK[k]}${n}"
        for r in range(2, n + 1):
            wp[f"{PK['size']}{r}"].value = (f'=IF({PK["inc"]}{r}="","",{PK["rel"]}{r}/SUMIFS({R("rel")},{R("side")},'
                                            f'{PK["side"]}{r},{R("inc")},"<>"))')
            wp[f"{PK['amt']}{r}"].value = (f'=IF({PK["size"]}{r}="","",IF({PK["side"]}{r}="SHORT",'
                                           f'{PK["size"]}{r}*${BL}$9,{PK["size"]}{r})*${BL}$12)')
        sidef = lambda name: (f'SUMPRODUCT(({R("side")}="{name}")*({R("inc")}<>"")*{R("rel")}*{R("beta")})'
                              f'/SUMIFS({R("rel")},{R("side")},"{name}",{R("inc")},"<>")')
        lab = lambda r, t: (wp.__setitem__(f"{BL}{r}", t), setattr(wp[f"{BL}{r}"], "font", Font(bold=True)))
        yel = PatternFill("solid", fgColor="FFF2CC")
        lab(1, "Way 1: Long budget (USDT) - type yours")
        wp[f"{BL}2"].value, wp[f"{BL}2"].fill, wp[f"{BL}2"].font = 250, yel, Font(bold=True)
        lab(4, "Way 2: USDT for a coin with Rel. size 1.0")
        wp[f"{BL}5"].fill, wp[f"{BL}5"].font = yel, Font(bold=True)         # left empty on purpose
        wp[f"{BL}6"].value = "Fill this and it overrides Way 1. Empty = Way 1."
        wp[f"{BL}6"].font = Font(italic=True, color="7F7F7F")
        lab(8, "Short multiplier")
        wp[f"{BL}9"].value = f'=IFERROR(MAX(0.5,MIN(2,({sidef("LONG")})/({sidef("SHORT")}))),1)'
        wp[f"{BL}9"].number_format = "0.00"
        lab(11, "Long budget in use (USDT)")
        wp[f"{BL}12"].value = (f'=IF(ISNUMBER({BL}5),{BL}5*SUMIFS({R("rel")},{R("side")},"LONG",{R("inc")},"<>"),{BL}2)')
        lab(14, "Long total (USDT)")
        wp[f"{BL}15"].value = f'=SUMIFS({R("amt")},{R("side")},"LONG")'
        lab(17, "Short total (USDT)")
        wp[f"{BL}18"].value = f'=SUMIFS({R("amt")},{R("side")},"SHORT")'
        for a in ("12", "15", "18"):
            wp[f"{BL}{a}"].number_format = "#,##0.0"
        pidx = {k: i + 1 for i, (k, _) in enumerate(pf_cols)}
        for row in wp.iter_rows(min_row=2, max_col=len(pf_cols)):
            for c in row:
                c.fill = fills["LONG"] if row[pidx["side"] - 1].value == "LONG" else fills["SHORT"]
        fmt_cells(wp, len(pf_cols), {pidx[k] for k in ("c6", "c24", "vol", "size")},
                  {pidx["beta"], pidx["rel"], pidx["ob"]}, {pidx["amt"]}, pidx["chg"])
        for c in wp[1]:
            c.font = Font(bold=True)
        pw = dict(sym=10, side=8, pos=7, chg=10, pos1=9, c6=10, c24=10, ob=10, hype=6, inc=8, beta=10, vol=14, size=12, rel=14, amt=14)
        for k, w in pw.items():
            wp.column_dimensions[PK[k]].width = w
        wp.column_dimensions[get_column_letter(len(pf_cols) + 1)].width = 3
        wp.column_dimensions[BL].width = 30
        for r in range(2, n + 1):
            wp[f"{PK['inc']}{r}"].alignment = Alignment(horizontal="center")
        wp[f"{PK['inc']}1"].alignment = Alignment(horizontal="center")
        for r in range(2, n + 1):
            hv = wp[f"{PK['hype']}{r}"].value
            if hv == "ok":
                wp[f"{PK['hype']}{r}"].fill = fills["OKAY"]
            elif hv not in (None, ""):
                wp[f"{PK['hype']}{r}"].fill = fills["WATCH"] if hv < HYPE_STRONG else fills["HYPE"]
            wp[f"{PK['hype']}{r}"].alignment = Alignment(horizontal="center")
        wp[f"{PK['hype']}1"].alignment = Alignment(horizontal="center")
        wp.freeze_panes = "B1"                      # Symbol column stays in view when scrolling right

        # ---- Ranking: colours, formats, and what was chosen on Portfolio ----
        last = ws.max_row
        ridx = {k: i + 1 for i, (k, _, _) in enumerate(rk_cols)}
        for r in range(2, last + 1):
            for key, src in (("inp", "inc"),):
                ref = (f'INDEX(Portfolio!${PK[src]}$2:${PK[src]}${n},'
                       f'MATCH({RK["sym"]}{r},Portfolio!${PK["sym"]}$2:${PK["sym"]}${n},0))')
                ws[f"{RK[key]}{r}"].value = f'=IFERROR(IF({ref}="","",{ref}),"")'
        ws.freeze_panes = "C2"                      # the correlation column and Symbol stay in view; Rank scrolls
        for row in ws.iter_rows(min_row=2):
            pick, flag = row[ridx["pick"] - 1].value or "", row[ridx["flag"] - 1].value or ""
            hp = row[ridx["hype"] - 1].value or ""
            strong = bool(hp) and hp.split(":")[0].isdigit() and int(hp.split(":")[0]) >= HYPE_STRONG
            fill = fills["BENCH"] if flag == "BENCHMARK" else fills["HYPE"] if strong else fills["IGNORE"] if flag.startswith("IGNORE") \
                else fills.get(pick)
            if hp and not strong:
                row[ridx["hype"] - 1].fill = fills["OKAY"] if hp.startswith("ok:") else fills["WATCH"]
                if not hp.startswith("ok:"):
                    row[ridx["sym"] - 1].fill = fills["WATCH"]
            for c in row:
                if fill:
                    c.fill = fill
                if flag.startswith("IGNORE"):
                    c.font = Font(color="7F7F7F")
        fmt_cells(ws, len(rk_cols), {ridx[k] for k in ("cum", "vol", "c6", "c24")},
                  {ridx[k] for k in ("score", "beta", "r2", "csol", "cmaj", "rel", "ob")}, set(), ridx["chg"])
        rw = dict(rank=6, sym=9, pick=8, chg=14, flag=38, score=8, wk=14, l3=16, cum=14, beta=7, r2=10, csol=5,
                  cmaj=14, vol=12, cap=9, hype=45, c6=10, c24=10, ob=11, rel=14, inp=10)
        for k, w in rw.items():
            ws.column_dimensions[RK[k]].width = w
        for c in ws[1]:
            c.font = Font(bold=True)
        for k_ in ("csol", "inp", "wk"):
            for r in range(1, last + 1):
                ws[f"{RK[k_]}{r}"].alignment = Alignment(horizontal="center")
        if timing is not None:
            write_timing_sheet(xw.book, timing, timing_charts(timing))
        if drivers is not None:
            D_, notes_ = drivers
            write_drivers_sheet(xw.book, D_, drivers_charts(D_) if D_ else (None, None), notes_)
        write_ignore_sheet(xw.book, df)
        xw.sheets["Notes"].column_dimensions["A"].width = 120


def write_ignore_sheet(book, df):
    """What-if sheet: live IGNORE recount from SOL r and Corr majors with cutoffs typed in yellow cells (no effect on the picks)."""
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.worksheet.datavalidation import DataValidation
    ws = book.create_sheet("IGNORE test", 4)
    bold, yel = Font(bold=True), PatternFill("solid", fgColor="FFF2CC")
    n = len(df)
    first, last = 11, 10 + n
    rng = lambda c: f"${c}${first}:${c}${last}"
    ws["A1"] = "What-if only: change the yellow cells and the counts and the last column update. The picks and Portfolio sheet do not change. To use the numbers: GitHub > Actions > Run workflow."
    ws["A1"].font = Font(italic=True, color="7F7F7F")
    labels = [("Test mode (majors / sol)", CORR_TEST), ("SOL r cutoff", MIN_CORR_SOL), ("Majors cutoff", MIN_CORR_MAJ)]
    for i, (lab, val) in enumerate(labels, start=2):
        ws[f"A{i}"], ws[f"B{i}"] = lab, val
        ws[f"A{i}"].font, ws[f"B{i}"].fill, ws[f"B{i}"].font = bold, yel, bold
    dv = DataValidation(type="list", formula1='"majors,sol"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("B2")
    ign_file = df.flag.astype(str).str.startswith("IGNORE").values
    ws["A5"], ws["B5"] = "IGNORE at these settings", f'=COUNTIF({rng("G")},"IGNORE")'
    ws["A6"], ws["B6"] = f"IGNORE in this run's file ({CORR_TEST}, {MIN_CORR_SOL}/{MIN_CORR_MAJ})", int(ign_file.sum())
    ws["A7"], ws["B7"] = "Current picks that would become IGNORE", f'=COUNTIFS({rng("E")},"?*",{rng("G")},"IGNORE")'
    ws["A8"], ws["B8"] = "Coins this run ignored that would be allowed", f'=COUNTIFS({rng("F")},"IGNORE",{rng("G")},"")'
    for r in (5, 6, 7, 8):
        ws[f"A{r}"].font = bold
    hdr = ["Symbol", "SOL r (1m)", "Maj r (1m)", "Rank (score)", "Pick (this run)", "IGNORE in file", "IGNORE at your settings"]
    for j, h in enumerate(hdr, 1):
        c = ws.cell(10, j, h)
        c.font, c.fill = bold, PatternFill("solid", fgColor="D9D9D9")
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    shade = PatternFill("solid", fgColor="F2F2F2")
    for k, (sym, row) in enumerate(df.iterrows()):
        r = first + k
        fl = "BENCH" if sym in BENCH_EX else ("IGNORE" if ign_file[k] else "")
        vals = [sym, None if pd.isna(row.corr_sol) else float(row.corr_sol), None if pd.isna(row.corr_maj) else float(row.corr_maj),
                None if pd.isna(row["rank"]) else int(row["rank"]), row.pick or "", fl]
        for j, v in enumerate(vals, 1):
            ws.cell(r, j, v)
        lo_s, lo_m = f"IF(ISNUMBER(B{r}),B{r}<$B$3,TRUE)", f"IF(ISNUMBER(C{r}),C{r}<$B$4,TRUE)"
        ws[f"G{r}"] = f'=IF(F{r}="BENCH","",IF(IF($B$2="sol",{lo_s},{lo_m}),"IGNORE",""))'
        ws[f"B{r}"].number_format = ws[f"C{r}"].number_format = "0.00"
        for j in range(1, 8):
            ws.cell(r, j).alignment = Alignment(horizontal="left" if j == 1 else "center")
            if (k // 5) % 2 == 1:
                ws.cell(r, j).fill = shade
    ws.conditional_formatting.add(f"G{first}:G{last}", FormulaRule(formula=[f'G{first}="IGNORE"'], font=Font(bold=True, color="C00000")))
    ws.column_dimensions["A"].width = 44
    for col, w in zip("BCDEFG", (9, 11, 9, 10, 10, 12)):
        ws.column_dimensions[col].width = w
    ws.row_dimensions[10].height = 32


# ----------------------------- Telegram -----------------------------
def send_telegram(path, body, subject):
    """Sends the summary text (subject on the first line), then the Excel file, to the chat. The bot token is never printed."""
    token, chat = os.environ["TELEGRAM_BOT_TOKEN"].strip(), os.environ["TELEGRAM_CHAT_ID"].strip()
    api = f"https://api.telegram.org/bot{token}/"

    def call(method, **kw):
        try:
            r = requests.post(api + method, timeout=120, **kw)
        except requests.RequestException as e:
            raise RuntimeError(f"Telegram {method} failed: {type(e).__name__}") from None     # the exception text would contain the token
        if r.status_code != 200 or not r.json().get("ok"):
            raise RuntimeError(f"Telegram {method} failed: HTTP {r.status_code} {r.text[:200]}")

    text = subject + "\n\n" + body
    chunks, cur = [], ""
    for line in text.split("\n"):                       # messages are limited to 4096 characters: split on line breaks
        if len(cur) + len(line) + 1 > 3900:
            chunks.append(cur)
            cur = ""
        cur += line + "\n"
    chunks.append(cur)
    for ch in chunks:
        call("sendMessage", data={"chat_id": chat, "text": ch, "disable_web_page_preview": "true"})
    with open(path, "rb") as f:
        call("sendDocument", data={"chat_id": chat, "caption": subject[:1000]}, files={"document": (os.path.basename(path), f)})


# ----------------------------- main -----------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--no-send", "--no-email", dest="no_send", action="store_true")
    args = ap.parse_args()
    print(f"ranker.py {VERSION}")

    now = pd.Timestamp.now(tz="UTC")
    h_end, m_end = now.floor("h"), now.floor("min")
    hours = (WINDOW_WEEKS * 7 + SKIP_DAYS) * 24 + 1

    uni = universe()
    perps = perp_bases() or perp_bases_cg()
    skipped = []
    if perps is None:
        skipped.append("perp check unavailable: coins without a Binance perpetual may be included")
    else:
        keep = [(s_, r_) for s_, r_ in uni if s_ == BENCH or has_perp(s_, perps) != "no"]
        skipped.append(f"{len(uni) - len(keep)} coins dropped because Binance has no perpetual for them")
        uni = keep
    syms, caprank = [s for s, _ in uni], dict(uni)
    hs, ms, failed = fetch_all(syms, h_end, m_end, hours)
    skipped += failed

    if BENCH not in hs or len(hs[BENCH]) < hours:
        sys.exit("Could not load enough SOL history.")
    idx = hs[BENCH].index[-hours:]
    H = pd.DataFrame(hs).reindex(idx).ffill(limit=3)
    for c in H.columns[H.isna().any()]:
        skipped.append(f"{c} (under 12 weeks + 1 week of history)")
    H = H.dropna(axis=1)
    for c in H.columns[(H.max() / H.min()) < 1.03]:
        if c != BENCH:
            skipped.append(f"{c} (looks pegged)")
            H = H.drop(columns=c)
    M = pd.DataFrame(ms).reindex(ms[BENCH].index).ffill(limit=5)[H.columns]

    hist_lists = load_history(now)
    df, longs, shorts, scale, cl, cs = analyze(H, M, caprank, perps, hist_lists)
    stamp = now.strftime("%Y-%m-%d %H:%M")
    path = f"ranker_{now:%Y-%m-%d_%H%M}.xlsx"
    tr, state = track(df, cl, cs, now, len(hist_lists))
    today_s = now.strftime("%Y-%m-%d")
    state[today_s]["t"] = H.index[-1].isoformat()          # last closed hour behind today's lists (used to replay earlier days point-in-time)
    picks = [dict(date=d, t=(pd.Timestamp(v["t"]) if v.get("t") else None), long=v["LONG"], short=v["SHORT"])
             for d, v in sorted(state.items()) if d < today_s and v.get("LONG") and v.get("SHORT")]
    try:
        timing = timing_analysis(H, longs, shorts, load_events(), picks)
    except Exception as e:
        print(f"Timing sheet skipped: {e}")
        timing = None
    drivers, lv_main = None, {}
    try:
        lv, dnotes = load_driver_levels()
        lv_main = lv
        try:
            D = drivers_analysis(H, longs, shorts, lv, load_events(), now)
            drivers = (D, dnotes)
        except Exception as e:
            print(f"Drivers study skipped: {e}")
            drivers = (None, [str(e)] + dnotes)
    except Exception as e:
        print(f"Drivers data failed: {e}")
        drivers = (None, [str(e)])
    if timing is not None:
        try:
            timing["gauge"] = gauge_analysis(H, longs, shorts, lv_main, drivers[0] if drivers else None)
        except Exception as e:
            print(f"Gauge skipped: {e}")
            timing["gauge"], timing["gauge_err"] = None, str(e)
    write_excel(path, df, longs, shorts, scale, skipped, stamp, cl, cs, tr, timing, drivers)
    bad_news = bad_news_lines(timing, drivers)
    if WHATIF_RUN:
        print("What-if run (settings typed in the Run workflow form): saved history left untouched.")
    else:
        os.makedirs(os.path.dirname(STATE_FILE) or ".", exist_ok=True)
        with open(STATE_FILE, "w") as f:
            json.dump(state, f, indent=1)
        for old in OLD_STATE_FILES:
            if os.path.exists(old):
                os.remove(old)                # renamed / moved: the old copy is no longer needed

    fmt = lambda d: "\n".join(f"  {s:<8} {v * 100:5.1f}%" for s, v in d["size"].items())
    body = (("BAD NEWS (last 72 hours):\n" + "\n".join(bad_news) + "\n\n" if bad_news else "")
            + f"Ranking generated {stamp} UTC\n\nLONG, all {len(cl)} kept (size, % of long budget)\n{fmt(longs)}\n\n"
            f"SHORT, all {len(cs)} kept (size, % of long budget; short side = {scale:.2f}x long notional)\n{fmt(shorts)}\n\n"
            f"Shortlist LONG: {', '.join(cl.index)}\nShortlist SHORT: {', '.join(cs.index)}\n\n"
            + (f"Market comfort gauge {timing['gauge']['cur']:+.2f}, zone {timing['gauge']['cur_zone'].upper()} (big = up to 15x15, normal = 10x10, small = 5x5 or none).\n\n"
               if timing and timing.get("gauge") else "")
            + f"Average positions use {tr['days']} earlier report(s).\n\n"
            f"Flagged IGNORE: {int(df.flag.str.startswith('IGNORE').sum())} coins. Full table in the Excel file.")
    print(body)
    if args.no_send:
        print(f"Saved {path}")
    else:
        subj = f"Ranker {now:%Y-%m-%d}"
        send_telegram(path, body, ("[what-if settings] " if WHATIF_RUN else "") + subj)
        print("Sent to Telegram.")


if __name__ == "__main__":
    main()
