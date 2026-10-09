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


def get(url, *, referer=None, timeout=30, tries=3):
    h = {"User-Agent": UA, "Accept": "text/html,application/json,text/csv,*/*",
         "Accept-Language": "en-US,en;q=0.9,ko;q=0.8"}
    if referer:
        h["Referer"] = referer
    last = None
    for i in range(tries):
        try:
            r = requests.get(url, headers=h, timeout=timeout)
            if r.status_code == 200 and r.content:
                return r
            last = f"HTTP {r.status_code}"
        except requests.RequestException as e:  # 네트워크 오류는 재시도
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
def update_holdings(force=False):
    h = load("holdings")
    # Schwab은 GitHub 서버 접속을 막는 경우가 많다(403). 실패해도 이슈를 만들지 않고,
    # 30일이 지나도록 못 받았을 때만 '확인 필요'로 알린다.
    age = (TODAY - dt.date.fromisoformat(h["asOf"])).days
    if not force and age < 6:
        return
    try:
        page = schwab_page()
        m = re.search(r"SCHD_FundHoldings_(\d{4}-\d{2}-\d{2})\.CSV", page)
        if not m:
            raise RuntimeError("상품 페이지에서 보유종목 CSV 링크를 찾지 못함")
        as_of = m.group(1)
        if as_of <= h["asOf"]:
            return
        text = get(SCHWAB_HOLD.format(d=as_of), referer=SCHWAB_PAGE).content.decode("utf-8-sig", errors="replace")
        items = []
        for row in csv.DictReader(io.StringIO(text)):
            if (row.get("As-Of-Date") or "").strip() != as_of:
                continue
            items.append(dict(symbol=row["Symbol"].strip(), weight=round(float(row["Percent of Assets"]), 6),
                              name=row["Name"].strip(), sector=(row.get("Sector") or "").strip()))
        total = sum(i["weight"] for i in items)
        if len(items) < 50 or not 98 <= total <= 102:
            raise RuntimeError(f"보유종목 검증 실패 (종목 {len(items)}개, 비중 합 {total:.2f}%)")
        items.sort(key=lambda x: -x["weight"])
        save("holdings", dict(asOf=as_of, source=h["source"], items=items))
        changed.append("holdings")
    except Exception as e:  # noqa: BLE001
        msg = f"구성종목 수집 실패: {e} (현재 데이터 기준일 {h['asOf']}, {age}일 경과)"
        (errors if age >= 30 else notes).append(msg)


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
