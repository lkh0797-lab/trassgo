"""에코프로비엠 v2 데이터 수집 — 충북 청주시 + 경북 포항시 HS 284190 36개월"""
import sys, io, urllib.request, re, json, time, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# .env 로드
env_path = 'config/.env'
KEY = None
with open(env_path, 'r', encoding='utf-8') as f:
    for line in f:
        if line.startswith('DATA_GO_KR_API_KEY='):
            KEY = line.split('=', 1)[1].strip()
            break

assert KEY, 'API key not found'

BASE = f'https://apis.data.go.kr/1220000/sigunguperprlstperacrs/getSigunguPerPrlstPerAcrs?serviceKey={KEY}&numOfRows=200&pageNo=1'

def fetch_period(start, end, hs, sido):
    url = f'{BASE}&strtYymm={start}&endYymm={end}&HsSgn={hs}&sidoCd={sido}'
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req, timeout=30) as resp:
        body = resp.read().decode('utf-8')
    items = re.findall(r'<item>(.*?)</item>', body, re.DOTALL)
    out = []
    for it in items:
        rec = {}
        for tag in ['priodTitle', 'sggNm', 'hsSgn', 'expCnt', 'expUsdAmt', 'impCnt', 'impUsdAmt', 'cmtrBlncAmt', 'korePrlstNm']:
            m = re.search(f'<{tag}>(.*?)</{tag}>', it)
            rec[tag] = m.group(1).strip() if m else None
        out.append(rec)
    return out

# 36개월: 2023-04 ~ 2026-03
periods = []
for y in [2023, 2024, 2025, 2026]:
    for m in range(1, 13):
        if y == 2023 and m < 4: continue
        if y == 2026 and m > 3: continue
        periods.append(f'{y}{m:02d}')

print(f'Periods: {periods[0]} ~ {periods[-1]}, count={len(periods)}')

# 12개월씩 batch
def fetch_batched(hs, sido, label):
    out = []
    for i in range(0, len(periods), 12):
        batch = periods[i:i+12]
        s, e = batch[0], batch[-1]
        print(f'  [{label}] {s}~{e}...', end='')
        recs = fetch_period(s, e, hs, sido)
        print(f' {len(recs)} records')
        out.extend(recs)
        time.sleep(0.5)
    return out

print('\n[1/2] 충청북도(43) HS 284190')
chungbuk = fetch_batched('284190', '43', '충북')

print('[2/2] 경상북도(47) HS 284190')
gyeongbuk = fetch_batched('284190', '47', '경북')

# 저장
out_data = {
    'meta': {
        'fetch_date': time.strftime('%Y-%m-%d %H:%M'),
        'hs_code': '284190',
        'hs_desc': '산의 금속산염류 (양극재 분류)',
        'period_range': f'{periods[0]} ~ {periods[-1]}',
        'sido_codes': {'43': '충청북도', '47': '경상북도'},
        'target_company': '에코프로비엠 (247540)',
        'extraction_logic': '충북 청주시 + 경북 포항시 합산 = 에코프로비엠 추정',
        'unit': 'expUsdAmt 단위는 천 USD',
    },
    'chungbuk': chungbuk,
    'gyeongbuk': gyeongbuk,
}

os.makedirs('data/raw', exist_ok=True)
with open('data/raw/customs_284190_v2.json', 'w', encoding='utf-8') as f:
    json.dump(out_data, f, ensure_ascii=False, indent=2)

print(f'\nSaved: data/raw/customs_284190_v2.json')

# 빠른 요약 — 청주시·포항시 월별
print('\n=== 청주시 + 포항시 양극재 수출 (단위: 천 USD) ===')
print(f'{"연월":<10s}{"청주(충북)":>15s}{"포항(경북)":>15s}{"합계":>15s}')
data_by_period = {}
for r in chungbuk:
    if '청주시' in r['sggNm']:
        p = r['priodTitle']
        v = int(r['expUsdAmt'].replace(',', ''))
        data_by_period.setdefault(p, {})['cheongju'] = v
for r in gyeongbuk:
    if '포항시' in r['sggNm']:
        p = r['priodTitle']
        v = int(r['expUsdAmt'].replace(',', ''))
        data_by_period.setdefault(p, {})['pohang'] = v

for p in sorted(data_by_period.keys()):
    cj = data_by_period[p].get('cheongju', 0)
    ph = data_by_period[p].get('pohang', 0)
    print(f'{p:<10s}{cj:>15,}{ph:>15,}{cj+ph:>15,}')
