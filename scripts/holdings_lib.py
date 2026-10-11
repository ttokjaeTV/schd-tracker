"""SCHD 구성종목 공통 처리 (표준 라이브러리만 사용 — PC에서도 바로 실행되게).

- Schwab 보유종목 CSV(SCHD_FundHoldings_YYYY-MM-DD.CSV) 해석·검증
- data/holdings.json 의 '공식 기준'(보유 주식 수·공식 비중) 교체
- data/holdings_monthly.json 월별 비중 기록 추가 (그 달 첫 기록만 남김)
"""
import csv
import datetime as dt
import io
import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
SOURCE = "Schwab Asset Management SCHD_FundHoldings CSV"

# Schwab 표기 → Yahoo 티커 (점·공백은 하이픈으로 바꾸고, 예외만 여기에)
YAHOO = {"CWENA": "CWEN-A"}


def yahoo(sym):
    return YAHOO.get(sym, sym.replace(".", "-").replace("/", "-").replace(" ", "-"))


def is_stock(item):
    return bool(item.get("sector"))


def is_cash(item):
    """섹터가 비어 있는 줄 중 선물이 아닌 것(달러 현금·MMF). 수량 = 금액(달러)."""
    return not item.get("sector") and not re.search(r"EMINI|MINI|FUT", item.get("name", ""), re.I)


def _f(s):
    s = (s or "").strip()
    return float(s) if s else None


def parse_schwab_csv(text):
    """CSV 본문 → (기준일, 행 목록). 행: symbol, shares, weight(없으면 0), name, sector"""
    text = text.lstrip("﻿")
    rows = []
    as_of = None
    for row in csv.DictReader(io.StringIO(text)):
        d = (row.get("As-Of-Date") or "").strip()
        if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", d):
            continue  # 하단 안내문 등
        as_of = as_of or d
        if d != as_of:
            continue
        rows.append(dict(symbol=row["Symbol"].strip(), shares=_f(row.get("Quantity")) or 0.0,
                         weight=_f(row.get("Percent of Assets")) or 0.0,
                         name=row["Name"].strip(), sector=(row.get("Sector") or "").strip()))
    if not as_of:
        raise RuntimeError("CSV에서 기준일(As-Of-Date) 행을 찾지 못함")
    stocks = [r for r in rows if r["sector"]]
    total = sum(r["weight"] for r in rows)
    if len(stocks) < 50 or not 98 <= total <= 102:
        raise RuntimeError(f"보유종목 검증 실패 (주식 {len(stocks)}종목, 비중 합 {total:.2f}%)")
    return as_of, rows


def load(name):
    return json.loads((DATA / f"{name}.json").read_text(encoding="utf-8"))


def save(name, obj):
    (DATA / f"{name}.json").write_text(json.dumps(obj, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")


def save_monthly(monthly):
    """한 달 = 한 줄로 저장 (파일 크기·커밋 차이를 작게)."""
    c = lambda o: json.dumps(o, ensure_ascii=False, separators=(",", ":"))
    head = {k: v for k, v in monthly.items() if k not in ("names", "months", "rebalances")}
    parts = ["{"] + [f"{c(k)}:{c(v)}," for k, v in head.items()]
    parts.append('"names":{')
    names = sorted(monthly["names"].items())
    parts += [f"{c(k)}:{c(v)}" + ("," if n < len(names) - 1 else "") for n, (k, v) in enumerate(names)]
    parts.append('},"months":[')
    ms = monthly["months"]
    parts += [c(m) + ("," if n < len(ms) - 1 else "") for n, m in enumerate(ms)]
    rb = monthly.get("rebalances") or []
    if rb:
        parts.append('],"rebalances":[')
        parts += [c(m) + ("," if n < len(rb) - 1 else "") for n, m in enumerate(rb)]
    parts.append("]}")
    (DATA / "holdings_monthly.json").write_text("\n".join(parts) + "\n", encoding="utf-8")


def official_holdings(as_of, rows):
    """공식 파일 그대로의 holdings.json (비중 = 공식 비중)."""
    items = [dict(symbol=r["symbol"], weight=round(r["weight"], 6), name=r["name"], sector=r["sector"],
                  shares=r["shares"], officialWeight=round(r["weight"], 6)) for r in rows]
    items.sort(key=lambda x: -x["weight"])
    return dict(asOf=as_of, basis="official", source=SOURCE,
                official=dict(asOf=as_of, source=SOURCE), items=items)


def add_month(monthly, as_of, items, basis):
    """그 달 기록이 없을 때만 추가. 추가했으면 True."""
    month = as_of[:7]
    if any(m["month"] == month for m in monthly["months"]):
        return False
    rec = dict(month=month, asOf=as_of, basis=basis,
               cash=round(sum(i["weight"] for i in items if is_cash(i)), 4),
               items=[[i["symbol"], round(i["weight"], 4), i["shares"]]
                      for i in sorted(items, key=lambda x: -x["weight"]) if is_stock(i)])
    for i in items:
        if is_stock(i):
            monthly["names"][i["symbol"]] = [i["name"], i["sector"]]
    monthly["months"].append(rec)
    monthly["months"].sort(key=lambda m: m["month"])
    return True


# ---------------------------------------------------------------- 리밸런싱일 기록
# Dow Jones U.S. Dividend 100 지수: 3·6·9·12월 셋째 금요일 장 마감 후 반영 → 다음 월요일 장 시작부터 적용.
# 3월은 연 1회 정기 종목 교체(+비중), 6·9·12월은 비중 재조정만. Schwab 일별 파일로 2021-12 ~ 2026-09 전부 확인.
# 3월 적용 월요일 파일은 새로 들어온 종목의 비중 칸이 비어 있어(주식 수만 있음) 검증에서 빠지고, 다음 거래일(화) 공식 비중으로 기록된다.
REBAL_MONTHS = (3, 6, 9, 12)
FIRST_REBAL = (2021, 12)  # Schwab 일별 보유종목 파일이 2021-09부터 있어 이때부터 기록 가능


def third_friday(y, m):
    d = dt.date(y, m, 1)
    d += dt.timedelta(days=(4 - d.weekday()) % 7)
    return d + dt.timedelta(days=14)


def rebal_schedule(until):
    """(연, 월, 셋째 금요일, 적용 월요일) — 적용 월요일이 until 이전인 것만."""
    y, m = FIRST_REBAL
    out = []
    while True:
        f = third_friday(y, m)
        mon = f + dt.timedelta(days=3)
        if mon > until:
            break
        out.append((y, m, f, mon))
        y, m = (y, m + 3) if m < 12 else (y + 1, 3)
    return out


def next_rebalance(today):
    y, m = today.year, today.month
    while True:
        if m in REBAL_MONTHS:
            mon = third_friday(y, m) + dt.timedelta(days=3)
            if mon >= today:
                return (y, m, mon)
        y, m = (y, m + 1) if m < 12 else (y + 1, 1)


def snapshot(rows):
    """공식 파일 행 → 기록용 (현금 비중, [[종목, 공식 비중, 보유 주식 수]...])"""
    stocks = sorted((r for r in rows if r["sector"]), key=lambda r: -r["weight"])
    cash = round(sum(r["weight"] for r in rows if is_cash(r)), 4)
    return cash, [[r["symbol"], round(r["weight"], 4), r["shares"]] for r in stocks]


def make_rebalance(kind, friday, monday, before_rows, after_rows, before_day, after_day, names):
    """직전 거래일 파일과 적용일 파일로 리밸런싱 기록 하나를 만든다. monday = 실제 적용 첫 거래일(월요일이 휴장이면 화요일)."""
    bc, bi = snapshot(before_rows)
    ac, ai = snapshot(after_rows)
    for r in before_rows + after_rows:
        if r["sector"]:
            names.setdefault(r["symbol"], [r["name"], r["sector"]])
    bs, as_ = {x[0]: x[2] for x in bi}, {x[0]: x[2] for x in ai}
    ratios = sorted(as_[k] / bs[k] for k in bs if k in as_ and bs[k])
    med = ratios[len(ratios) // 2] if ratios else 1
    moved = sum(1 for k in bs if k in as_ and bs[k] and abs(as_[k] / bs[k] / med - 1) > 0.01)
    return dict(date=after_day, before=before_day, kind=kind, thirdFriday=friday.isoformat(),
                effective=monday.isoformat(), added=len(set(as_) - set(bs)), removed=len(set(bs) - set(as_)),
                reweighted=moved, beforeCash=bc, cash=ac, beforeItems=bi, items=ai)

