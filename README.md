# SCHD 배당 트래커

SCHD(Schwab U.S. Dividend Equity ETF) 분기 배당 이력, 연간 배당·성장률, 원화 환산, 구성종목, 국내상장 미국배당다우존스 ETF 4종 비교를 한 페이지로 보여 줍니다.
GitHub Actions가 매일 새 데이터를 확인해 자동으로 갱신합니다.

## 폴더 구조

| 경로 | 내용 |
|---|---|
| `index.html` | 트래커 화면. `data/*.json`을 읽어서 그립니다 |
| `data/dividends.json` | SCHD 배당 (배당락일·기준일·지급일·공시 금액·배당락일 종가·지급일 환율·출처) |
| `data/holdings.json` | 구성종목 전체 목록 (Schwab 공시) |
| `data/kr_etf.json` | 국내 SCHD형 ETF 4종 분배 이력, 총보수 |
| `data/prices.json` | SCHD 일별 종가·배당 재투자 수정주가 (2011-10 상장 이후 전체) |
| `data/meta.json` | 갱신일, 주식분할 이력, 한글 종목명·섹터명, 원천징수율 |
| `downloads/schd_dividend_tracker.xlsx` | 엑셀 마스터 (데이터가 바뀔 때마다 자동 생성) |
| `scripts/update.py` | 데이터 수집·검증 |
| `scripts/build_xlsx.py` | 엑셀 생성 |
| `.github/workflows/update.yml` | 매일 한국시간 08:30 자동 실행 |

## 자동 갱신 규칙

- **SCHD 배당**: Schwab 공식 분배 CSV → 막히면 stockanalysis(예비, 2026-10 첫 실행 기준 GitHub 서버에서는 Schwab이 막혀 예비 출처로 동작). 새 배당이 직전 대비 ±40% 넘게 다르면 반영하지 않고 이슈로 알림
- **배당락일 종가**: Yahoo Finance (분할 반영, 배당 미조정 종가)
- **주가 흐름**: Yahoo Finance 전체 이력을 7일마다(새 배당 반영 시엔 즉시) 통째로 다시 받음
- **지급일 환율**: 하나은행 매매기준율(다음금융). 지급일이 지나야 채워짐
- **구성종목**: Schwab 보유종목 CSV, 6일 이상 지났을 때만 확인. Schwab이 GitHub 서버 접속을 막으면(403) 기존 데이터를 유지하고, 30일 넘게 못 받으면 `확인 필요` 이슈로 알림
- **국내 SCHD형 분배금**: FunETF 분배 이력
- 새 배당이 반영되면 `SCHD 새 배당 반영` 이슈가, 수집 실패가 있으면 `확인 필요` 이슈가 생깁니다 (GitHub 알림 메일로 받음)

## 손으로 고칠 때

- 값은 `data/*.json`에서 고칩니다. 엑셀은 다음 갱신 때 새로 만들어지므로 엑셀에서 고친 내용은 남지 않습니다.
- 주식분할이 생기면 `data/meta.json`의 `splits`에 `{"date": "분할일", "ratio": 배수}`를 추가합니다.
- 국내 ETF 총보수가 바뀌면 `data/kr_etf.json`의 `fee`와 `fee_source`를 고칩니다.
- Actions 탭 → `SCHD 데이터 자동 갱신` → `Run workflow`로 언제든 수동 실행할 수 있습니다.

## 로컬에서 보기

`index.html`을 더블클릭하면 데이터를 못 읽습니다(브라우저 보안). 폴더에서 `python -m http.server` 실행 후 http://localhost:8000 으로 엽니다.

출처: Schwab Asset Management, Yahoo Finance, 다음금융(하나은행 고시 환율), FunETF, 금융투자협회. 특정 종목의 매수·매도를 권하는 자료가 아닙니다.
