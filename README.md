# 수출입통계 트래커

관세청 시군구별 품목별 수출입 데이터를 종목 단위로 추정·트래킹하는 개인용 시스템.

## 빠른 시작

```bat
# 1) 종목 추가
scripts\add_stock.bat

# 2) 트래커 실행 (엑셀 생성)
scripts\run_manual.bat

# 결과: output\수출입통계_트래커_YYYY-MM.xlsx
```

## 디렉터리

```
config/
  .env                       # API 키 (채팅 노출 시 재발급 필요)
  settings.yaml              # 시그널 임계값 등
  sido_codes.csv             # 시도 코드 (편집 불필요)
data/
  master/stock_mapping.csv   # 종목 매핑 마스터 (수동 관리)
  raw/                       # API 캐시 (자동)
src/                         # 코드
output/                      # 매월 엑셀
logs/                        # 실행 로그
scripts/                     # bat 파일
```

## 종목 추가 — 세 가지 방법

### 방법 1: 시군구 자동 디스커버리 ⭐ 추천
종목명 + HS코드만 알면 시군구는 데이터에서 자동으로 찾아줍니다.
```bat
scripts\discover_sgg.bat
# 또는
python -m src.discover_sgg --code 066970 --name 엘앤에프 --hs 284190 --hs-desc "양극재"
```
17개 시도 전부 fetch → 시군구별 수출액 ranking → Top-N 중 선택 → 자동 매핑. 동종업계 노이즈와 함께 보이므로 confidence를 한눈에 판단 가능.

### 방법 2: 클로드와 대화
> "한미반도체 추가해줘"

클로드가 사업보고서/공시 보고 HS코드 + 생산거점 추정, 데이터 검증 후 매핑 추가.

### 방법 3: 직접 명령어 (시군구를 이미 아는 경우)
```bat
python -m src.add_stock --code 042700 --name 한미반도체 ^
  --hs 848620 --hs-desc "반도체 본더 등 가공기계" ^
  --sido 28 --sgg "남동구" --note "HBM 본더 주력" -y
```

검증 결과 보여주고 OK 받으면 `data/master/stock_mapping.csv`에 append.

## 자동 갱신 (Windows 작업 스케줄러)

```
작업 스케줄러 → 작업 만들기 → 트리거: 매월 1일 09:00 → 동작: scripts\run_monthly_phase1.bat
                            트리거: 매월 16일 09:00 → 동작: scripts\run_monthly_phase2.bat
```

## 시그널

| 타입 | 정의 |
|---|---|
| MOMENTUM | YoY ≥ +30% AND 3개월 연속 양수 |
| BREAKOUT | 24개월 최고치 갱신 |
| WARN | YoY ≤ -25% AND 3개월 연속 감소 |

`config/settings.yaml`에서 임계값 조정 가능.

## 텔레그램 알림

`config/settings.yaml`의 `telegram.enabled: true` 설정 + `config/.env`에 토큰/chat_id 입력하면 작동.

토큰 발급:
1. 텔레그램 [@BotFather](https://t.me/BotFather)에서 `/newbot`
2. 본인 만든 봇과 대화 시작 (`/start`)
3. `https://api.telegram.org/bot<토큰>/getUpdates` 접속해 chat_id 확인

## 한계

- **포항·창원 등 산업 클러스터 시군구는 동종업계 합산 노이즈** (포스코퓨처엠과 에코프로비엠 같이 잡힘) → 동종업계 지표로 활용
- **다종목 사업** (SK이노, LG화학 등) 추정 정확도 낮음
- **내수 위주 종목** (금융·유통·통신) 분석 무의미 → 매핑 제외
- **데이터 지연**: 4월 데이터는 5월 17~20일경 OpenAPI 반영
