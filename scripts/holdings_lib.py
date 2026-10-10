"""SCHD 구성종목 공통 처리 (표준 라이브러리만 사용 — PC에서도 바로 실행되게).

- Schwab 보유종목 CSV(SCHD_FundHoldings_YYYY-MM-DD.CSV) 해석·검증
- data/holdings.json 의 '공식 기준'(보유 주식 수·공식 비중) 교체
- data/holdings_monthly.json 월별 비중 기록 추가 (그 달 첫 기록만 남김)
"""
import csv
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
    head = {k: v for k, v in monthly.items() if k not in ("names", "months")}
    parts = ["{"] + [f"{c(k)}:{c(v)}," for k, v in head.items()]
    parts.append('"names":{')
    names = sorted(monthly["names"].items())
    parts += [f"{c(k)}:{c(v)}" + ("," if n < len(names) - 1 else "") for n, (k, v) in enumerate(names)]
    parts.append('},"months":[')
    ms = monthly["months"]
    parts += [c(m) + ("," if n < len(ms) - 1 else "") for n, m in enumerate(ms)]
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
