# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 주요 명령어

```bat
# 엑셀 대시보드 생성 (매월 수동 실행)
scripts\run_manual.bat
# == python -m src.runner --no-telegram

# 캐시 무시하고 강제 재fetch
python -m src.runner --force-fetch

# 종목 추가 (CLI)
python -m src.add_stock --code 042700 --name 한미반도체 --hs 848620 --hs-desc "반도체 본더" --sido 28 --sgg 남동구 --note "HBM 본더" -y

# 종목 추가 (대화형)
scripts\add_stock.bat

# HS코드 후보 검증 스크립트 실행 예시
python verify_candidates3.py

# API/매핑 단독 스모크 테스트
python -m src.customs_api
python -m src.mapper
```

## 아키텍처

```
관세청 OpenAPI → customs_api.py → mapper.py → signals.py
                                     ↓               ↓
                              StockTimeseries    Signal list
                                     ↓               ↓
                               dashboard.py (openpyxl 엑셀)
                                                     ↓
                               telegram_bot.py (알림)
```

**`runner.py`** 가 오케스트레이터. `run()` 한 번이면 fetch → 시그널 → 엑셀 → 텔레그램 전 과정 실행.

### 모듈별 역할

| 모듈 | 핵심 함수 | 설명 |
|---|---|---|
| `customs_api.py` | `fetch_sigungu_range()` | API 호출 + 파일 캐시 (data/raw/). 1회 호출 최대 12개월 → `split_periods()`로 배치 |
| `mapper.py` | `build_stock_timeseries()` | 매핑 CSV 로드 → 시군구 키워드 필터 → 종목별 합산 + MoM/YoY 계산 |
| `signals.py` | `compute_all_signals()` | MOMENTUM/BREAKOUT/WARN 시그널 산출. 임계값은 settings.yaml |
| `dashboard.py` | `build_dashboard()` | 엑셀 생성. 시트: Summary, {종목}_상세, Universe, HS_Mapping, README |
| `add_stock.py` | `probe_data()` / `add_stock()` | 신규 매핑 전 API 검증 → stock_mapping.csv append |

### 데이터 흐름

1. **매핑 마스터**: `data/master/stock_mapping.csv` — 한 종목이 여러 행(HS×시도 조합)이 있을 수 있음. 같은 종목의 복수 행은 `weight`를 적용해 합산
2. **API 캐시**: `data/raw/sigungu_{hs}_sido{sido}_{start}-{end}.json` — 90일 유효. 동일 (HS, 시도) 조합은 build_stock_timeseries 내에서도 중복 fetch 방지
3. **금액 단위**: 관세청 API 응답 `expUsdAmt`는 **천 USD(kUSD)**. 대시보드 표시 시 `/1000`하면 백만 달러(M)
4. **데이터 지연**: 당월 데이터는 익익월 17~20일경 반영. `latest_available_yymm()`이 자동 탐지

## 핵심 제약사항 / 함정

- **전북특별자치도 시도 코드는 반드시 `52`** (구 전라북도 코드 `45`는 API 오류 99 반환)
- **API 파라미터 대소문자**: `HsSgn` (대문자 H), `sidoCd`, `strtYymm`, `endYymm` — 변경 시 API error 99
- **Windows 콘솔 한글**: 스크립트 진입점 최상단에 `sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')` 필수
- **엑셀 PermissionError**: Excel에서 파일 열려있으면 저장 실패 → `build_dashboard()`에서 타임스탬프 suffix로 fallback 처리됨
- **포항 시군구 합산 노이즈**: 에코프로비엠 HS 284190 + 포스코퓨처엠이 포항에서 같이 잡힘 → confidence='mid' 표시, 동종업계 지표로 활용

## 종목 추가 시 판단 기준

`probe_data(hs, sido, sgg)` 반환값 `total_exp_kusd` 기준:
- **> 50,000 kUSD ($50M)** → 강한 신호, high confidence
- **5,000–50,000 kUSD** → 중간, mid confidence
- **500–5,000 kUSD** → 약함, low confidence
- **0** → HS코드 또는 시군구 키워드 재검토 필요

## 설정 파일

- `config/.env` — `DATA_GO_KR_API_KEY` (data.go.kr에서 발급). 유출 시 재발급 필요
- `config/settings.yaml` — 시그널 임계값, API 배치 크기, 캐시 기간 등
- `config/.env.example` — 커밋용 템플릿 (`.env`는 .gitignore로 제외)

## 시그널 정의

| 타입 | 조건 |
|---|---|
| MOMENTUM | 최신월 YoY ≥ +30% **AND** 최근 3개월 모두 YoY > 0 |
| BREAKOUT | 최신월 수출액 > 직전 24개월 최고치 |
| WARN | 최신월 YoY ≤ -25% **AND** 최근 3개월 연속 수출액 감소 |
