"""관세청 OpenAPI 클라이언트 — 시군구별 품목별 + 품목별 국가별

캐싱 전략: 파일 기반 (data/raw/), 같은 (HS, 시도, 기간) 조합은 한 번만 호출.
"""
from __future__ import annotations

import json
import re
import time
import urllib.request
import urllib.error
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import yaml


PROJECT_ROOT = Path(__file__).resolve().parent.parent
CONFIG_DIR = PROJECT_ROOT / 'config'
RAW_CACHE_DIR = PROJECT_ROOT / 'data' / 'raw'


def load_settings() -> dict:
    with open(CONFIG_DIR / 'settings.yaml', 'r', encoding='utf-8') as f:
        return yaml.safe_load(f)


def load_api_key() -> str:
    env_path = CONFIG_DIR / '.env'
    with open(env_path, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if line.startswith('DATA_GO_KR_API_KEY='):
                return line.split('=', 1)[1].strip()
    raise RuntimeError('DATA_GO_KR_API_KEY not found in config/.env')


@dataclass(frozen=True)
class SigunguQuery:
    hs_code: str          # 6단위
    sido_code: str        # 11, 43, 47 등
    yymm_start: str       # YYYYMM
    yymm_end: str         # YYYYMM


@dataclass
class SigunguRecord:
    period: str           # YYYY.MM
    sido: str             # 충청북도 등 (응답엔 시군구만 있어 시도는 호출 시 알고 있음)
    sgg_name: str         # 충청북도 청주시
    hs_code: str
    exp_count: int
    exp_kusd: int         # 천 USD
    imp_count: int
    imp_kusd: int
    balance_kusd: int


_settings = load_settings()


def _build_url(base: str, params: dict) -> str:
    qs = '&'.join(f'{k}={v}' for k, v in params.items())
    return f'{base}?{qs}'


def _fetch_xml(url: str, retries: int | None = None, retry_delay: int | None = None,
               timeout: int | None = None) -> str:
    retries = retries or _settings['api']['retry_count']
    retry_delay = retry_delay or _settings['api']['retry_delay_sec']
    timeout = timeout or _settings['api']['request_timeout_sec']

    last_err = None
    for attempt in range(retries):
        try:
            req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                return resp.read().decode('utf-8')
        except (urllib.error.URLError, TimeoutError) as e:
            last_err = e
            if attempt < retries - 1:
                time.sleep(retry_delay)
    raise RuntimeError(f'API fetch failed after {retries} retries: {last_err}')


def _parse_sigungu_xml(xml: str, sido_name_map: dict[str, str]) -> list[SigunguRecord]:
    # 응답 헤더 체크
    rc = re.search(r'<resultCode>(.*?)</resultCode>', xml)
    if rc and rc.group(1).strip() != '00':
        msg = re.search(r'<resultMsg>(.*?)</resultMsg>', xml)
        raise RuntimeError(f'API error: {rc.group(1)} - {msg.group(1) if msg else "unknown"}')

    items = re.findall(r'<item>(.*?)</item>', xml, re.DOTALL)
    out: list[SigunguRecord] = []
    for it in items:
        def get(tag: str) -> str:
            m = re.search(f'<{tag}>(.*?)</{tag}>', it)
            return m.group(1).strip() if m else ''

        def to_int(s: str) -> int:
            s = s.strip().replace(',', '')
            if not s or s == '-':
                return 0
            try:
                return int(s)
            except ValueError:
                return 0

        sgg_name = get('sggNm')
        sido_name = sgg_name.split()[0] if sgg_name else ''
        out.append(SigunguRecord(
            period=get('priodTitle'),
            sido=sido_name,
            sgg_name=sgg_name,
            hs_code=get('hsSgn'),
            exp_count=to_int(get('expCnt')),
            exp_kusd=to_int(get('expUsdAmt')),
            imp_count=to_int(get('impCnt')),
            imp_kusd=to_int(get('impUsdAmt')),
            balance_kusd=to_int(get('cmtrBlncAmt')),
        ))
    return out


def _cache_path(query: SigunguQuery) -> Path:
    fname = f'sigungu_{query.hs_code}_sido{query.sido_code}_{query.yymm_start}-{query.yymm_end}.json'
    return RAW_CACHE_DIR / fname


def _is_cache_fresh(path: Path, max_age_days: int) -> bool:
    if not path.exists():
        return False
    age_sec = time.time() - path.stat().st_mtime
    return age_sec < max_age_days * 86400


def fetch_sigungu(query: SigunguQuery, *, use_cache: bool = True,
                  api_key: str | None = None) -> list[SigunguRecord]:
    """한 번 호출 = 한 (HS, 시도, 기간[≤12개월]) 조합."""
    cache_path = _cache_path(query)
    cache_days = _settings['data']['cache_days']

    if use_cache and _is_cache_fresh(cache_path, cache_days):
        with open(cache_path, 'r', encoding='utf-8') as f:
            return [SigunguRecord(**r) for r in json.load(f)]

    api_key = api_key or load_api_key()
    base = _settings['api']['customs_endpoint_sigungu']
    url = _build_url(base, {
        'serviceKey': api_key,
        'strtYymm': query.yymm_start,
        'endYymm': query.yymm_end,
        'HsSgn': query.hs_code,
        'sidoCd': query.sido_code,
        'numOfRows': 1000,
        'pageNo': 1,
    })

    xml = _fetch_xml(url)
    records = _parse_sigungu_xml(xml, {})

    RAW_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    with open(cache_path, 'w', encoding='utf-8') as f:
        json.dump([r.__dict__ for r in records], f, ensure_ascii=False, indent=2)

    time.sleep(_settings['api']['inter_call_delay_sec'])
    return records


def split_periods(yymm_start: str, yymm_end: str, batch_months: int = 12) -> list[tuple[str, str]]:
    """기간을 batch_months 단위로 분할."""
    def to_idx(yymm: str) -> int:
        return int(yymm[:4]) * 12 + int(yymm[4:]) - 1

    def to_yymm(idx: int) -> str:
        return f'{idx // 12:04d}{idx % 12 + 1:02d}'

    s, e = to_idx(yymm_start), to_idx(yymm_end)
    out = []
    cur = s
    while cur <= e:
        nxt = min(cur + batch_months - 1, e)
        out.append((to_yymm(cur), to_yymm(nxt)))
        cur = nxt + 1
    return out


def fetch_sigungu_range(hs_code: str, sido_code: str, yymm_start: str, yymm_end: str,
                        *, use_cache: bool = True) -> list[SigunguRecord]:
    """긴 기간을 12개월씩 분할 호출 후 병합."""
    api_key = load_api_key()
    batch_months = _settings['api']['batch_months']
    out: list[SigunguRecord] = []
    for s, e in split_periods(yymm_start, yymm_end, batch_months):
        recs = fetch_sigungu(
            SigunguQuery(hs_code=hs_code, sido_code=sido_code, yymm_start=s, yymm_end=e),
            use_cache=use_cache, api_key=api_key,
        )
        out.extend(recs)
    return out


def filter_by_sgg_keyword(records: Iterable[SigunguRecord], keyword: str) -> list[SigunguRecord]:
    """시군구 이름에 키워드 포함된 행만 필터 (예: '청주시')."""
    return [r for r in records if keyword in r.sgg_name]


def latest_available_yymm(probe_hs: str = '284190', probe_sido: str = '43') -> str:
    """가장 최신 가용 데이터 월 탐지 (현재 월부터 거꾸로 6개월 시도)."""
    from datetime import datetime
    now = datetime.now()
    api_key = load_api_key()
    base = _settings['api']['customs_endpoint_sigungu']
    for back in range(6):
        y, m = now.year, now.month - back
        while m <= 0:
            y -= 1
            m += 12
        yymm = f'{y:04d}{m:02d}'
        url = _build_url(base, {
            'serviceKey': api_key,
            'strtYymm': yymm, 'endYymm': yymm,
            'HsSgn': probe_hs, 'sidoCd': probe_sido,
            'numOfRows': 5, 'pageNo': 1,
        })
        try:
            xml = _fetch_xml(url, retries=1)
            items = re.findall(r'<item>', xml)
            if len(items) > 0:
                return yymm
        except Exception:
            continue
    raise RuntimeError('Cannot detect latest yymm')


if __name__ == '__main__':
    # Smoke test
    print('Latest available yymm:', latest_available_yymm())
    recs = fetch_sigungu_range('284190', '43', '202503', '202603')
    print(f'Fetched {len(recs)} records for 충북 HS 284190')
    cheongju = filter_by_sgg_keyword(recs, '청주시')
    for r in cheongju[:5]:
        print(f'  {r.period}  {r.sgg_name}  exp={r.exp_kusd:>12,} kUSD')
