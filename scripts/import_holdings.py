"""Schwab 보유종목 파일을 받아 보유 주식 수(공식 기준)를 교체한다. (수동용)

평소에는 GitHub Actions(update.py)가 매일 자동으로 받는다. 자동 수집이 계속 막힐 때 직접 받은 파일을 넣는다.
표준 라이브러리만 쓰므로 패키지 설치 없이 실행된다.

  python3 scripts/import_holdings.py 받은파일.CSV
  python3 scripts/import_holdings.py 받은파일.CSV --rows 102 --qty 1306484445
      (--rows·--qty: 크롬에서 센 행 수·수량 합계. 옮겨 적다 틀리면 멈춘다)

- 새 파일 기준일이 지금 공식 기준일보다 늦을 때만 data/holdings.json 을 바꾼다.
- 그 달 월별 기록이 아직 없으면 data/holdings_monthly.json 에도 넣는다.
- 바뀐 게 있으면 마지막 줄에 CHANGED, 없으면 UNCHANGED 를 출력한다.
"""
import argparse
import base64
import gzip
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import holdings_lib as HL  # noqa: E402


def read_text(path):
    raw = Path(path).read_bytes()
    if raw.lstrip()[:10].startswith(b"As-Of-Date") or raw.lstrip()[:13].startswith("﻿As-Of-Date".encode()):
        return raw.decode("utf-8-sig")
    return gzip.decompress(base64.b64decode(b"".join(raw.split()))).decode("utf-8-sig")


def main():
    ap = argparse.ArgumentParser(description="Schwab 보유종목 파일로 보유 주식 수 교체")
    ap.add_argument("path")
    ap.add_argument("--rows", type=int, help="원본 행 수(검증용)")
    ap.add_argument("--qty", type=int, help="원본 수량 합계 반올림값(검증용)")
    a = ap.parse_args()
    as_of, rows = HL.parse_schwab_csv(read_text(a.path))
    qsum = round(sum(r["shares"] for r in rows))
    if (a.rows is not None and a.rows != len(rows)) or (a.qty is not None and abs(a.qty - qsum) > 1):
        print(f"검증 실패: 행 {len(rows)} (기대 {a.rows}), 수량 합계 {qsum} (기대 {a.qty}) — 옮겨 적은 내용을 확인하세요")
        print("FAILED")
        return 1
    h = HL.load("holdings")
    old = h["official"]["asOf"]
    stocks = [r for r in rows if r["sector"]]
    print(f"Schwab 파일 기준일 {as_of} · 주식 {len(stocks)}종목 · 비중 합 {sum(r['weight'] for r in rows):.2f}%")
    if as_of <= old:
        print(f"지금 기준일({old})보다 새롭지 않아 바꾸지 않음")
        print("UNCHANGED")
        return 0
    before = {i["symbol"] for i in h["items"] if HL.is_stock(i)}
    after = {r["symbol"] for r in stocks}
    new = HL.official_holdings(as_of, rows)
    HL.save("holdings", new)
    monthly = HL.load("holdings_monthly")
    if HL.add_month(monthly, as_of, new["items"], "official"):
        HL.save_monthly(monthly)
        print(f"월별 기록에 {as_of[:7]} 추가")
    print(f"보유 주식 수 기준일 {old} → {as_of}")
    if after - before:
        print("신규 편입:", ", ".join(sorted(after - before)))
    if before - after:
        print("편출:", ", ".join(sorted(before - after)))
    print("CHANGED")
    return 0


if __name__ == "__main__":
    sys.exit(main())
