"""SCHD 배당 트래커 데이터 자동 갱신.

GitHub Actions가 매일 실행한다. 새 데이터가 있을 때만 data/*.json 을 고치고,
결과를 .run/ 폴더에 남겨 워크플로가 커밋·이슈 생성 여부를 판단하게 한다.

  .run/changed        데이터가 바뀌었으면 생성 (엑셀 재생성·커밋 트리거)
  .run/new_dividend.md 새 SCHD 배당이 추가됐으면 생성 (알림 이슈 본문)
  .run/errors.md      수집 실패·검증 실패가 있으면 생성 (경고 이슈 본문)

로컬 실행:  python scripts/update.py          (전체)
            python scripts/update.py --only dividends
"""
import argparse
import csv
import datetime as dt
import io
import json
import re
import sys
import time
from pathlib import Path

import requests

sys.path.insert(0, str(Path(__file__).resolve().parent))
import holdings_lib as HL  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
RUN = ROOT / ".run"
KST = dt.timezone(dt.timedelta(hours=9))
TODAY = dt.datetime.now(KST).date()

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36")
SCHWAB_PAGE = "https://www.schwabassetmanagement.com/products/schd"
SCHWAB_DIST = ("https://www.schwabassetmanagement.com/sites/g/files/eyrktu361/files/"
               "product_files/SCHD/SCHD_Fund_Distributions.CSV")
SCHWAB_HOLD = ("https://www.schwabassetmanagement.com/sites/g/files/eyrktu361/files/"
               "product_files/SCHD/SCHD_FundHoldings_{d}.CSV")

# 새 배당이 직전 배당 대비 이 범위를 벗어나면 반영하지 않고 경고만 남긴다
MAX_CHANGE = 0.40

_page_cache: dict = {}

errors: list[str] = []
notes: list[str] = []  # 매일 이슈를 만들 정도는 아닌 참고사항 (새 배당 알림에만 덧붙임)
changed: list[str] = []


def load(name):
    return json.loads((DATA / f"{name}.json").read_text(encoding="utf-8"))


def save(name, obj):
    (DATA / f"{name}.json").write_text(json.dumps(obj, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")


_browser = None


def _browser_session():
    """Schwab(Akamai)은 일반 파이썬 요청을 403으로 막고 실제 브라우저의 접속 방식(TLS 지문)은 통과시킨다.
    curl_cffi 로 크롬처럼 접속한다. 설치돼 있지 않으면 None."""
    global _browser
    if _browser is None:
        try:
            from curl_cffi import requests as cr
            _browser = cr.Session(impersonate="chrome")
        except Exception:  # noqa: BLE001
            _browser = False
    return _browser or None


def get(url, *, referer=None, timeout=30, tries=3):
    h = {"User-Agent": UA, "Accept": "text/html,application/json,text/csv,*/*",
         "Accept-Language": "en-US,en;q=0.9,ko;q=0.8"}
    if referer:
        h["Referer"] = referer
    browser = _browser_session() if "schwabassetmanagement.com" in url else None
    last = None
    for i in range(tries):
        try:
            if browser:
                r = browser.get(url, headers={"Referer": referer} if referer else None, timeout=timeout)
            else:
                r = requests.get(url, headers=h, timeout=timeout)
            if r.status_code == 200 and r.content:
                return r
            last = f"HTTP {r.status_code}"
        except Exception as e:  # noqa: BLE001  네트워크 오류는 재시도
            last = str(e)
        time.sleep(2 * (i + 1))
    raise RuntimeError(f"{url} → {last}")


def schwab_page():
    """SCHD 상품 페이지 HTML. 구성종목·공식 수익률이 같이 쓰므로 한 번만 받는다."""
    if "html" not in _page_cache:
        _page_cache["html"] = get(SCHWAB_PAGE).text
    return _page_cache["html"]


def mdy(s):
    return dt.datetime.strptime(s.strip(), "%m/%d/%Y").date().isoformat()


def split_factor(ex, splits):
    f = 1
    for s in splits:
        if ex < s["date"]:
            f *= s["ratio"]
    return f


# ---------------------------------------------------------------- SCHD 배당
def fetch_schwab_distributions():
    r = get(SCHWAB_DIST, referer=SCHWAB_PAGE)
    text = r.content.decode("utf-8-sig", errors="replace")
    rows = []
    for row in csv.DictReader(io.StringIO(text)):
        ex = (row.get("Ex-Date") or "").strip()
        if not re.match(r"\d{2}/\d{2}/\d{4}$", ex):
            continue
        amt = float(row["Total Distribution"])
        rows.append(dict(ex=mdy(ex), record=mdy(row["Record Date"]), pay=mdy(row["Payable Date"]),
                         amount=round(amt, 6), source="Schwab 공식"))
    if not rows:
        raise RuntimeError("Schwab 분배 CSV에서 행을 찾지 못함")
    return rows


def fetch_stockanalysis_distributions():
    """Schwab 차단 시 예비 출처. 분할조정 금액이라 2024-10-11 이후 배당에만 사용한다."""
    r = get("https://stockanalysis.com/etf/schd/dividend/")
    rows = []
    pat = re.compile(r"(\w{3} \d{1,2}, \d{4})\s*\|\s*\$([\d.]+)\s*\|\s*(\w{3} \d{1,2}, \d{4})\s*\|\s*(\w{3} \d{1,2}, \d{4})")
    plain = re.sub(r"<[^>]+>", "|", r.text)
    plain = re.sub(r"(\|\s*)+", "| ", plain)
    for m in pat.finditer(plain):
        f = lambda s: dt.datetime.strptime(s, "%b %d, %Y").date().isoformat()
        rows.append(dict(ex=f(m.group(1)), record=f(m.group(3)), pay=f(m.group(4)),
                         amount=float(m.group(2)), source="stockanalysis(예비)"))
    if not rows:
        raise RuntimeError("stockanalysis 배당 표를 읽지 못함")
    return rows


def update_dividends(meta):
    divs = load("dividends")
    known = {d["ex"] for d in divs}
    try:
        src = fetch_schwab_distributions()
    except Exception as e:  # noqa: BLE001
        notes.append(f"Schwab 공식 CSV 접근 실패({e}) → stockanalysis 예비 출처로 반영. Schwab 공식 금액과 한 번 대조해 주세요.")
        try:
            src = [r for r in fetch_stockanalysis_distributions() if r["ex"] >= "2024-10-11"]
        except Exception as e2:  # noqa: BLE001
            errors.append(f"예비 출처(stockanalysis)도 실패: {e2}")
            src = []

    added = []
    for r in sorted(src, key=lambda x: x["ex"]):
        if r["ex"] in known:
            continue
        prev = divs[-1]
        prev_adj = prev["amount"] / split_factor(prev["ex"], meta["splits"])
        new_adj = r["amount"] / split_factor(r["ex"], meta["splits"])
        chg = new_adj / prev_adj - 1
        if r["ex"] <= prev["ex"] or r["pay"] < r["ex"]:
            errors.append(f"날짜 이상으로 보류: {r}")
            continue
        if abs(chg) > MAX_CHANGE:
            errors.append(f"직전 대비 {chg:+.1%} 변화로 보류 (분할·특별배당 여부 확인 필요): {r}")
            continue
        r.update(close=None, fx=None, added=TODAY.isoformat())
        divs.append(r)
        known.add(r["ex"])
        added.append((r, chg))

    # 기존 행의 금액 정정(운용사 수정 공시) 감지 — 자동 반영하지 않고 알림만
    by_ex = {r["ex"]: r for r in src if r["source"] == "Schwab 공식"}
    for d in divs:
        s = by_ex.get(d["ex"])
        if s and abs(s["amount"] - d["amount"]) > 1e-4 and d["source"] == "Schwab 공식":
            errors.append(f"Schwab 공시 금액이 기존 기록과 다름: {d['ex']} 기존 {d['amount']} / 공시 {s['amount']}")

    fill_prices(divs)
    fill_fx(divs)
    save("dividends", divs)

    if added:
        changed.append("dividends")
        lines = ["| 배당락일 | 지급일 | 배당금 | 직전 대비 |", "|---|---|---|---|"]
        for r, chg in added:
            lines.append(f"| {r['ex']} | {r['pay']} | ${r['amount']:.4f} | {chg:+.2%} |")
        (RUN / "new_dividend.md").write_text(
            "SCHD 새 배당이 트래커에 반영됐습니다.\n\n" + "\n".join(lines) +
            "\n\n출처: " + ", ".join(sorted({r['source'] for r, _ in added})) +
            "\n\n지급일 환율은 지급일이 지난 뒤 자동으로 채워집니다.\n" +
            ("".join(f"\n> {n}" for n in notes)), encoding="utf-8")


def fill_prices(divs):
    need = [d for d in divs if d.get("close") is None and d["ex"] <= TODAY.isoformat()]
    if not need:
        return
    try:
        import yfinance as yf
        start = min(d["ex"] for d in need)
        end = (dt.date.fromisoformat(max(d["ex"] for d in need)) + dt.timedelta(days=5)).isoformat()
        px = yf.download("SCHD", start=start, end=end, progress=False, auto_adjust=False)
        close = px["Close"]
        if hasattr(close, "columns"):
            close = close.iloc[:, 0]
        series = {i.strftime("%Y-%m-%d"): float(v) for i, v in close.items()}
        for d in need:
            if d["ex"] in series:
                d["close"] = round(series[d["ex"]], 4)
                changed.append("close")
    except Exception as e:  # noqa: BLE001
        errors.append(f"배당락일 종가(Yahoo) 수집 실패: {e}")


def fill_fx(divs):
    need = [d for d in divs if d.get("fx") is None and d["pay"] <= TODAY.isoformat()]
    if not need:
        return
    try:
        rates = {}
        for page in range(1, 4):
            r = get("https://finance.daum.net/api/exchanges/FRX.KRWUSD/days?symbolCode=FRX.KRWUSD"
                    f"&page={page}&perPage=100&pagination=true",
                    referer="https://finance.daum.net/exchanges/FRX.KRWUSD")
            for x in r.json()["data"]:
                rates[x["date"][:10]] = x["basePrice"]
            if min(rates) <= min(d["pay"] for d in need):
                break
        for d in need:
            day = dt.date.fromisoformat(d["pay"])
            for back in range(7):  # 지급일이 국내 휴일이면 직전 영업일
                k = (day - dt.timedelta(days=back)).isoformat()
                if k in rates:
                    d["fx"] = rates[k]
                    changed.append("fx")
                    break
    except Exception as e:  # noqa: BLE001
        errors.append(f"지급일 환율(다음금융) 수집 실패: {e}")


# ---------------------------------------------------------------- 주가 흐름
def update_prices(force=False):
    """일별 종가·배당재투자 수정주가 전체 이력. 7일마다(또는 새 배당 반영 시) 통째로 다시 받는다.
    수정주가(adj)는 배당이 생길 때마다 과거 값 전체가 바뀌므로 덧붙이지 않고 교체한다."""
    p = load("prices")
    if not force and (TODAY - dt.date.fromisoformat(p["asOf"])).days < 7:
        return
    try:
        import yfinance as yf
        px = yf.download("SCHD", period="max", progress=False, auto_adjust=False)
        c, a = px["Close"], px["Adj Close"]
        if hasattr(c, "columns"):
            c, a = c.iloc[:, 0], a.iloc[:, 0]
        rows = [[i.strftime("%Y-%m-%d"), round(float(cv), 4), round(float(av), 4)]
                for i, cv, av in zip(px.index, c, a) if cv == cv and av == av]
        if len(rows) < len(p["rows"]) - 5 or rows[0][0] != p["rows"][0][0]:
            raise RuntimeError(f"주가 이력이 기존보다 짧거나 시작일이 다름 ({len(rows)}행, 시작 {rows[0][0] if rows else '-'})")
        if rows[-1][0] != p["asOf"] or force:
            save("prices", dict(asOf=rows[-1][0], source=p["source"], rows=rows))
            changed.append("prices")
    except Exception as e:  # noqa: BLE001
        notes.append(f"주가(Yahoo) 갱신 실패: {e}")


# ---------------------------------------------------------------- 구성종목
# 보유 주식 수는 Schwab 공식 파일 기준(매일 시도). 못 받으면 stockanalysis 상위 25종목으로 리밸런싱을 감지해 보정.
# 비중은 매일 '보유 주식 수 × 그날 종가'로 다시 계산하고, 매월 첫 거래일 비중은 holdings_monthly.json 에 남긴다.
# Schwab 파일과 같은 날 종가로 계산하면 공식 비중과 소수점 넷째 자리까지 같다(2026-09/10 대조).
OFFICIAL_STALE_DAYS = 30  # 공식 파일을 이만큼 못 받으면 '확인 필요' 이슈
SA_URL = "https://stockanalysis.com/etf/schd/holdings/"


def fetch_official_holdings(h):
    """새 공식 파일이 있으면 공식 기준(보유 주식 수·공식 비중)을 통째로 교체한다."""
    page = schwab_page()
    m = re.search(r"SCHD_FundHoldings_(\d{4}-\d{2}-\d{2})\.CSV", page)
    if not m:
        raise RuntimeError("상품 페이지에서 보유종목 CSV 링크를 찾지 못함")
    if m.group(1) <= h["official"]["asOf"]:
        return None
    text = get(SCHWAB_HOLD.format(d=m.group(1)), referer=SCHWAB_PAGE).content.decode("utf-8-sig", errors="replace")
    return HL.official_holdings(*HL.parse_schwab_csv(text))


def refresh_top_shares(h):
    """공식 파일을 못 받을 때의 보조: stockanalysis 무료 상위 25종목 보유 주식 수(Schwab 데이터와 같음)로
    리밸런싱을 감지해 상위 종목 주식 수를 바꾼다. 바꿨으면 True."""
    html = get(SA_URL).text
    m = re.search(r"As of</span>\s*([A-Z][a-z]{2} \d{1,2}, \d{4})", html)
    if not m:
        raise RuntimeError("stockanalysis 기준일을 찾지 못함")
    as_of = dt.datetime.strptime(m.group(1), "%b %d, %Y").date().isoformat()
    base_day = max(h["official"]["asOf"], (h.get("top") or {}).get("asOf", ""))
    if as_of <= base_day:
        return False
    sa = {x.group(1): float(x.group(2).replace(",", ""))
          for x in re.finditer(r's:"\$([A-Z0-9.\-]+)",as:"[\d.]+%",sh:"([\d,]+)"', html)}
    if len(sa) < 20:
        raise RuntimeError(f"stockanalysis 상위 종목이 {len(sa)}개뿐")
    base = {i["symbol"]: i["shares"] for i in h["items"] if HL.is_stock(i)}
    ratios = sorted(sa[k] / base[k] for k in sa if base.get(k))
    if len(ratios) < 15:
        raise RuntimeError("stockanalysis 종목과 기존 종목이 너무 다름")
    med = ratios[len(ratios) // 2]
    moved = [k for k in sa if not base.get(k) or abs(sa[k] / base[k] / med - 1) > 0.005]
    h["top"] = dict(asOf=as_of, source="stockanalysis 상위 25종목", checked=TODAY.isoformat())
    if not moved:
        return True  # 리밸런싱 없음 (주식 수 그대로) — 확인 날짜만 갱신
    names = HL.load("holdings_monthly")["names"]
    items = {i["symbol"]: i for i in h["items"]}
    for k, v in sa.items():
        if k in items:
            items[k] = {**items[k], "shares": round(v / med, 4)}
        else:  # 상위 25에 새로 들어온 종목
            nm, sec = names.get(k, [k, "Unclassified"])
            items[k] = dict(symbol=k, weight=0.0, name=nm, sector=sec, shares=round(v / med, 4), officialWeight=0.0)
    h["items"] = list(items.values())
    notes.append(f"stockanalysis {as_of} 상위 종목 주식 수가 바뀌어(리밸런싱 추정) 반영: {', '.join(moved)}")
    return True


def compute_weights(h):
    """공식 보유 주식 수 × 최근 종가로 비중 계산. 새 종가 날짜가 없으면 None."""
    import pandas as pd
    import yfinance as yf
    off = h["official"]["asOf"]
    stocks = [i for i in h["items"] if HL.is_stock(i)]
    tick = {i["symbol"]: HL.yahoo(i["symbol"]) for i in stocks}
    start = (dt.date.fromisoformat(off) - dt.timedelta(days=7)).isoformat()
    px = yf.download(sorted(set(tick.values())), start=start, progress=False, auto_adjust=False,
                     actions=True, threads=True)
    close = px["Close"]
    splits = px["Stock Splits"] if "Stock Splits" in px.columns.get_level_values(0) else None
    valid = close.notna().sum(axis=1)
    ok = valid[valid >= 0.9 * close.shape[1]]
    if ok.empty:
        raise RuntimeError("종가가 충분히 받아지지 않음")
    P = ok.index[-1]
    pday = P.strftime("%Y-%m-%d")
    if pday <= off or (h.get("basis") == "computed" and h["asOf"] == pday):
        return None
    mv, missing, stale = {}, [], []
    for i in stocks:
        t = tick[i["symbol"]]
        ser = close[t][close.index <= P].dropna() if t in close.columns else pd.Series(dtype=float)
        if ser.empty:
            missing.append(i)
            continue
        if ser.index[-1] != P:
            stale.append(i["symbol"])
        f = 1.0
        if splits is not None and t in splits.columns:
            sp = splits[t][(splits.index > pd.Timestamp(off)) & (splits.index <= P)]
            for v in sp:
                if v and v == v:
                    f *= float(v)
        mv[i["symbol"]] = i["shares"] * f * float(ser.iloc[-1])
    cash = sum(i["shares"] for i in h["items"] if HL.is_cash(i))
    gone = sum(i["officialWeight"] for i in missing) / 100
    if gone > 0.05:
        raise RuntimeError(f"종가를 못 받은 종목 비중이 너무 큼 ({gone * 100:.1f}%): " + ", ".join(i["symbol"] for i in missing))
    total = (sum(mv.values()) + cash) / (1 - gone)
    for i in missing:
        mv[i["symbol"]] = i["officialWeight"] / 100 * total
    items = []
    for i in h["items"]:
        w = mv[i["symbol"]] / total * 100 if HL.is_stock(i) else (i["shares"] / total * 100 if HL.is_cash(i) else 0.0)
        items.append({**i, "weight": round(w, 6)})
    items.sort(key=lambda x: -x["weight"])
    if missing or stale:
        notes.append("비중 계산: " + (f"종가 없음(공식 비중 유지) {', '.join(i['symbol'] for i in missing)} " if missing else "")
                     + (f"최근 종가가 {pday} 이전 {', '.join(stale)}" if stale else ""))
    return {**h, "asOf": pday, "basis": "computed", "items": items}


def update_holdings(force=False):
    h = HL.load("holdings")
    age = (TODAY - dt.date.fromisoformat(h["official"]["asOf"])).days
    try:
        new = fetch_official_holdings(h)  # 매일 시도 (크롬처럼 접속하면 GitHub에서도 받아짐)
        if new:
            h = new
            changed.append("holdings")
    except Exception as e:  # noqa: BLE001
        msg = f"구성종목 공식 파일 수집 실패: {e} (보유 주식 수 기준일 {h['official']['asOf']}, {age}일 경과)"
        (errors if age >= OFFICIAL_STALE_DAYS else notes).append(msg)
        try:
            if refresh_top_shares(h):
                changed.append("holdings")
        except Exception as e2:  # noqa: BLE001
            notes.append(f"stockanalysis 상위 종목 확인 실패: {e2}")
    try:
        new = compute_weights(h)
        if new:
            h = new
            changed.append("holdings")
    except Exception as e:  # noqa: BLE001
        notes.append(f"구성종목 비중 계산 실패: {e} (기존 비중 유지)")
    if "holdings" in changed:
        HL.save("holdings", h)
        monthly = HL.load("holdings_monthly")
        if HL.add_month(monthly, h["asOf"], h["items"], h["basis"]):
            HL.save_monthly(monthly)
            changed.append("holdings_monthly")


# ---------------------------------------------------------------- 공식 수익률·분배 일정
def update_official(meta, force=False):
    """Schwab 상품 페이지의 30일 SEC 수익률·분배수익률(TTM)을 기준일과 함께 저장한다.
    SEC 수익률은 매일 바뀌므로 기준일이 7일 이상 새로워졌을 때(또는 새 배당 반영 시)만 갱신해 매일 커밋을 막는다.
    GitHub 서버에서 Schwab이 막히면(403) 기존 값을 유지한다 — 화면에 기준일이 함께 표시된다."""
    o = meta.get("official")
    if not o:
        return
    try:
        html = schwab_page()
        got = {}
        for key, label in (("sec_yield", r"SEC Yield \(30 Day\)"), ("dist_yield_ttm", r"Distribution Yield \(TTM\)")):
            m = re.search(label + r".{0,400}?As of (\d{2}/\d{2}/\d{4}).{0,200}?<td>\s*([\d.]+)%\s*</td>", html, re.S)
            if not m:
                raise RuntimeError(f"상품 페이지에서 {key} 값을 찾지 못함")
            v = float(m.group(2)) / 100
            if not 0.005 < v < 0.10:
                raise RuntimeError(f"{key} 값이 이상함: {v:.4f}")
            got[key] = dict(value=round(v, 4), asOf=mdy(m.group(1)))
        old_sec = dt.date.fromisoformat(o["sec_yield"]["asOf"])
        new_sec = dt.date.fromisoformat(got["sec_yield"]["asOf"])
        if force or (new_sec - old_sec).days >= 7 or got["dist_yield_ttm"]["asOf"] != o["dist_yield_ttm"]["asOf"]:
            o.update(got, checked=TODAY.isoformat())
            changed.append("official")
    except Exception as e:  # noqa: BLE001
        age = (TODAY - dt.date.fromisoformat(o["sec_yield"]["asOf"])).days
        notes.append(f"Schwab 공식 수익률 갱신 실패: {e} (현재 SEC 기준일 {o['sec_yield']['asOf']}, {age}일 경과)")


def check_schedule(meta, divs):
    """분배 일정(meta.schedule)은 Schwab이 연 1회 PDF로 내므로 손으로 넣는다. 다음 예정일이 없으면 알림."""
    sch = meta.get("schedule") or {}
    last_ex = divs[-1]["ex"]
    if not any(r["ex"] > last_ex for r in sch.get("rows", [])):
        notes.append("다음 분배 예정일이 data/meta.json의 schedule에 없습니다. "
                     "Schwab 'Schwab Equity ETFs Distribution Schedule' 새해 일정을 확인해 넣어 주세요.")


# ---------------------------------------------------------------- 국내 SCHD형 ETF
def update_kr_etf():
    k = load("kr_etf")
    before = json.dumps(k, ensure_ascii=False)
    for code in k["order"]:
        e = k["etfs"][code]
        try:
            r = get(f"https://www.funetf.co.kr/api/public/product/view/etfdividend?itemId={e['isin']}",
                    referer="https://www.funetf.co.kr/")
            hist = {x[0]: x for x in e["history"]}
            for x in r.json():
                d = x.get("gijunYmd")
                if d and x.get("divAmt") is not None:
                    hist[d] = [d, float(x["divAmt"]), float(x.get("divRt") or 0)]
            e["history"] = [hist[d] for d in sorted(hist)]
        except Exception as ex:  # noqa: BLE001
            errors.append(f"{e['name']} 분배금(FunETF) 수집 실패: {ex}")
        time.sleep(1)
    if json.dumps(k, ensure_ascii=False) != before:
        save("kr_etf", k)
        changed.append("kr_etf")


# ---------------------------------------------------------------- 국내 SCHD형 실부담비용 (KOFIA)
KOFIA_URL = "https://dis.kofia.or.kr/proframeWeb/XMLSERVICES/"


def fetch_kofia_costs(basis):
    """금융투자협회 전자공시 '펀드별 보수비용비교' (상장지수 펀드 전체). 단위 %. 미공개 기준일은 빈 목록."""
    import xml.etree.ElementTree as ET
    payload = ('<?xml version="1.0" encoding="utf-8"?><message><proframeHeader>'
               '<pfmAppName>FS-DIS2</pfmAppName><pfmSvcName>DISFundFeeCmsSO</pfmSvcName>'
               '<pfmFnName>select</pfmFnName></proframeHeader><systemHeader></systemHeader>'
               f'<DISCondFuncDTO><tmpV30>{basis}</tmpV30><tmpV11></tmpV11><tmpV12>상장지수</tmpV12>'
               '<tmpV3></tmpV3><tmpV5></tmpV5><tmpV4></tmpV4></DISCondFuncDTO></message>')
    last = None
    for i in range(4):  # 응답이 1.6MB라 중간에 끊기는 일이 잦다 → 재시도
        try:
            r = requests.post(KOFIA_URL, data=payload.encode("utf-8"), timeout=120,
                              headers={"Content-Type": "application/xml; charset=UTF-8", "User-Agent": UA})
            root = ET.fromstring(r.content.decode("utf-8"))
            break
        except Exception as e:  # noqa: BLE001
            last = e
            time.sleep(3 * (i + 1))
    else:
        raise RuntimeError(f"KOFIA 조회 실패: {last}")
    num = lambda x: float(x) if x not in (None, "") else 0.0
    out = {}
    for m in root.findall(".//selectMeta"):
        g = lambda t: ((m.find(t).text or "").strip() if m.find(t) is not None else "")
        if not g("tmpV2"):
            continue
        ter, sales, trade = num(g("tmpV12")), num(g("tmpV13")) + num(g("tmpV14")), num(g("tmpV16"))
        out[g("tmpV15")] = dict(total_fee=num(g("tmpV9")), other=num(g("tmpV11")), ter=ter, trading=trade,
                                sales=sales, real=round(ter + sales + trade, 4))
    return out


def update_real_cost():
    """월말 기준·약 한 달 뒤 공개. 저장된 기준일보다 새 달(지난달 말)이 공개됐으면 4종목 실부담비용을 바꾼다."""
    k = load("kr_etf")
    have = min((k["etfs"][c].get("cost") or {}).get("basis", "0000-00-00") for c in k["order"])
    first = TODAY.replace(day=1)
    month_end = first - dt.timedelta(days=1)          # 지난달 말일
    if month_end.isoformat() <= have:
        return
    # 지난달 마지막 평일부터 거꾸로 (월말이 휴일이면 그 전 영업일이 기준일)
    cands, d = [], month_end
    while len(cands) < 4:
        if d.weekday() < 5:
            cands.append(d)
        d -= dt.timedelta(days=1)
    try:
        for b in cands:
            rows = fetch_kofia_costs(b.strftime("%Y%m%d"))
            if len(rows) < 500:  # 미공개(0건) 또는 휴일
                continue
            missing = []
            for c in k["order"]:
                e = k["etfs"][c]
                code = (e.get("cost") or {}).get("kofia_code")
                if code in rows:
                    e["cost"] = dict(basis=b.isoformat(), **rows[code], kofia_code=code)
                else:
                    missing.append(e["name"])
            if missing:
                errors.append(f"KOFIA {b} 실부담비용에서 못 찾은 종목: {', '.join(missing)} (kofia_code 확인 필요)")
            save("kr_etf", k)
            changed.append("kr_etf")
            notes.append(f"국내 SCHD형 실부담비용을 KOFIA {b} 기준으로 갱신")
            return
    except Exception as e:  # noqa: BLE001
        notes.append(f"KOFIA 실부담비용 조회 실패: {e}")


# ---------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--only", choices=["dividends", "prices", "holdings", "kr_etf"])
    ap.add_argument("--force-holdings", action="store_true")
    a = ap.parse_args()
    RUN.mkdir(exist_ok=True)
    for f in RUN.glob("*"):
        f.unlink()

    meta = load("meta")
    if a.only in (None, "dividends"):
        update_dividends(meta)
    if a.only in (None, "dividends"):
        update_official(meta, force="dividends" in changed)
        check_schedule(meta, load("dividends"))
    if a.only in (None, "prices"):
        update_prices(force="dividends" in changed)
    if a.only in (None, "holdings"):
        update_holdings(force=a.force_holdings)
    if a.only in (None, "kr_etf"):
        update_kr_etf()
        update_real_cost()

    if changed:
        meta["updated"] = TODAY.isoformat()
        save("meta", meta)
        (RUN / "changed").write_text(",".join(sorted(set(changed))), encoding="utf-8")
    if errors:
        (RUN / "errors.md").write_text("자동 갱신 중 확인이 필요한 항목이 있습니다.\n\n" +
                                       "\n".join(f"- {e}" for e in errors) + "\n", encoding="utf-8")
    print("changed:", sorted(set(changed)) or "없음")
    if notes:
        print("notes:", *notes, sep="\n  ")
    print("errors:", *errors, sep="\n  ") if errors else print("errors: 없음")
    return 0


if __name__ == "__main__":
    sys.exit(main())
