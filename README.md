# SCHD 배당 트래커

SCHD(Schwab U.S. Dividend Equity ETF) 분기 배당 이력, 연간 배당·성장률, 원화 환산, 구성종목, 국내상장 미국배당다우존스 ETF 4종 비교를 한 페이지로 보여 줍니다.
GitHub Actions가 매일 새 데이터를 확인해 자동으로 갱신합니다.

## 폴더 구조

| 경로 | 내용 |
|---|---|
| `index.html` | 트래커 화면. `data/*.json`을 읽어서 그립니다 |
| `data/dividends.json` | SCHD 배당 (배당락일·기준일·지급일·공시 금액·배당락일 종가·지급일 환율·출처) |
| `data/holdings.json` | 구성종목 전체 목록. `official` = Schwab 공시 보유 주식 수의 기준일, `asOf` = 비중을 계산한 종가 날짜 |
| `data/holdings_monthly.json` | 매월 첫 거래일 비중 기록 (2021-10부터, 한 달 = 한 줄) |
| `data/kr_etf.json` | 국내 SCHD형 ETF 4종 분배 이력, 총보수 |
| `data/prices.json` | SCHD 일별 종가·배당 재투자 수정주가 (2011-10 상장 이후 전체) |
| `data/meta.json` | 갱신일, 주식분할 이력, 한글 종목명·섹터명, 원천징수율 |
| `downloads/schd_dividend_tracker.xlsx` | 엑셀 마스터 (데이터가 바뀔 때마다 자동 생성) |
| `scripts/update.py` | 데이터 수집·검증 |
| `scripts/holdings_lib.py` | 구성종목 공통 처리 (Schwab CSV 해석·월별 기록, 표준 라이브러리만 사용) |
| `scripts/import_holdings.py` | (수동용) 직접 받은 Schwab 보유종목 파일로 보유 주식 수 교체 |
| `scripts/build_xlsx.py` | 엑셀 생성 |
| `.github/workflows/update.yml` | 매일 한국시간 08:30 자동 실행 |

## 자동 갱신 규칙

- **SCHD 배당**: Schwab 공식 분배 CSV → 막히면 stockanalysis(예비, 2026-10 첫 실행 기준 GitHub 서버에서는 Schwab이 막혀 예비 출처로 동작). 새 배당이 직전 대비 ±40% 넘게 다르면 반영하지 않고 이슈로 알림
- **배당락일 종가**: Yahoo Finance (분할 반영, 배당 미조정 종가)
- **주가 흐름**: Yahoo Finance 전체 이력을 7일마다(새 배당 반영 시엔 즉시) 통째로 다시 받음
- **지급일 환율**: 하나은행 매매기준율(다음금융). 지급일이 지나야 채워짐
- **Schwab 접속**: Schwab(Akamai)은 일반 파이썬 요청을 403으로 막고 브라우저 접속은 통과시킴 → `curl_cffi`로 크롬처럼 접속(배당 CSV·보유종목 CSV·공식 수익률 모두 적용)
- **구성종목**: 매일 Schwab 공식 보유종목 CSV(SCHD_FundHoldings_날짜.CSV)를 받아 그대로 반영. 못 받은 날은 '마지막 공식 보유 주식 수 × Yahoo 최근 종가'로 비중을 다시 계산하고(주식분할은 Yahoo 분할 정보로 보정), stockanalysis 무료 상위 25종목 주식 수(Schwab 데이터와 같음)로 리밸런싱을 감지해 상위 종목 주식 수를 바꿈. 2026-09-01 주식 수에서 시작해 이 방식으로 10/8을 계산하면 공시 비중과 최대 0.07%p 차이. 공식 파일을 30일 넘게 못 받으면 `확인 필요` 이슈
- **월별 비중 기록**: 그 달 첫 계산(보통 첫 거래일 종가)을 `holdings_monthly.json`에 추가. 2021-10 ~ 2026-10은 Schwab 공식 파일 그대로(`basis: official`), 이후는 계산값(`basis: computed`). Schwab 공식 파일은 2021-09-16 이후만 공개돼 있음
- **공식 수익률**: Schwab 상품 페이지의 30일 SEC 수익률·분배수익률(TTM)을 기준일과 함께 `data/meta.json`의 `official`에 저장. SEC 기준일이 7일 이상 새로워졌거나 분배수익률 기준일이 바뀌었을 때만 갱신(매일 커밋 방지). Schwab이 막히면 기존 값 유지
- **다음 배당 예정일**: Schwab이 연 1회 내는 'Schwab Equity ETFs Distribution Schedule'을 `data/meta.json`의 `schedule`에 손으로 입력. 금액이 발표되기 전까지 화면에는 날짜만 '예정'으로 표시
- **국내 SCHD형 분배금**: FunETF 분배 이력
- 새 배당이 반영되면 `SCHD 새 배당 반영` 이슈가, 수집 실패가 있으면 `확인 필요` 이슈가 생깁니다 (GitHub 알림 메일로 받음)

## 손으로 고칠 때

- 값은 `data/*.json`에서 고칩니다. 엑셀은 다음 갱신 때 새로 만들어지므로 엑셀에서 고친 내용은 남지 않습니다.
- 해가 바뀌면 `data/meta.json`의 `schedule.rows`에 새해 분기 일정(`q`·`ex`·`record`·`pay`)을 넣고 `source`·`checked`를 고칩니다. 예정일이 바닥나면 새 배당 알림 이슈에 안내가 붙습니다.
- 공식 수익률을 손으로 고칠 때는 `official`의 `value`(소수, 3.37% → 0.0337)와 `asOf`를 같이 고칩니다.
- 주식분할이 생기면 `data/meta.json`의 `splits`에 `{"date": "분할일", "ratio": 배수}`를 추가합니다.
- 보유 주식 수를 손으로 넣을 때(자동 수집이 계속 막힐 때만): Schwab 상품 페이지 Holdings → Export All Holdings로 받은 CSV를 `python scripts/import_holdings.py 파일.CSV`.
- 국내 ETF 총보수가 바뀌면 `data/kr_etf.json`의 `fee`와 `fee_source`를 고칩니다.
- 국내 SCHD형 ETF를 새로 추가하면 `data/kr_etf.json`에 상장일 `listed`(YYYY-MM-DD)도 넣습니다. 분기 합계 표의 '상장 전 / 상장 후 첫 분배 전' 구분에 씁니다.
- Actions 탭 → `SCHD 데이터 자동 갱신` → `Run workflow`로 언제든 수동 실행할 수 있습니다.

## 로컬에서 보기

`index.html`을 더블클릭하면 데이터를 못 읽습니다(브라우저 보안). 폴더에서 `python -m http.server` 실행 후 http://localhost:8000 으로 엽니다.

출처: Schwab Asset Management, Yahoo Finance, 다음금융(하나은행 고시 환율), FunETF, 금융투자협회. 특정 종목의 매수·매도를 권하는 자료가 아닙니다.
