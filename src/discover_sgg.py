"""종목명 + HS코드 → 유력 시군구 자동 디스커버리

워크플로우 (Q1=a + Q2=c + Q3=a):
  1. 17개 시도 전부 fetch (HS 코드 고정, 최근 12개월)
  2. 시군구별 수출액 누적 합계
  3. Top-N 시군구 ranking 표시
  4. 사용자가 매핑할 시군구 선택 (복수 가능)
  5. stock_mapping.csv에 append (confidence 자동 산출)

사용 예:
  python -m src.discover_sgg --code 042700 --name 한미반도체 --hs 848620 --hs-desc "본더"
  python -m src.discover_sgg --code 042700 --name 한미반도체 --hs 848620 --auto-top 1
"""
from __future__ import annotations

import argparse
import io
import sys
from collections import defaultdict
from datetime import datetime

from .customs_api import fetch_sigungu_range, latest_available_yymm
from .mapper import MappingRow, add_mapping_row, load_mappings, load_sido_codes


def _yymm_back(yymm: str, months: int) -> str:
    y, m = int(yymm[:4]), int(yymm[4:])
    idx = y * 12 + m - 1 - months
    return f'{idx // 12:04d}{idx % 12 + 1:02d}'


def _extract_sgg_keyword(sgg_name: str) -> str:
    """관세청 시군구명 → stock_mapping.csv용 키워드 추출.

    예시:
      '충청북도 청주시 흥덕구' → '청주시'
      '경상북도 포항시 남구'   → '포항시'
      '전라남도 해남군'         → '해남군'
      '서울특별시 강남구'       → '강남구'
    """
    tokens = sgg_name.split()
    if len(tokens) >= 2:
        return tokens[1]
    return sgg_name


def discover_sgg(hs_code: str, months: int = 12, *, use_cache: bool = True) -> list[dict]:
    """모든 시도에서 hs_code 데이터 fetch → 시군구별 수출액 ranking.

    Returns: list of dicts sorted by exp_kusd_total desc
    """
    yymm_end = latest_available_yymm()
    yymm_start = _yymm_back(yymm_end, months - 1)

    sido_map = load_sido_codes()  # {sido_code: sido_name}

    # 시군구 키 = (sido_code, sgg_name) — 동명 시군구 분리
    by_sgg: dict[tuple[str, str], dict] = {}

    print(f'\n시도 {len(sido_map)}개 fetch 중 (HS={hs_code}, {yymm_start}~{yymm_end})...')
    for sido_code, sido_name in sido_map.items():
        try:
            recs = fetch_sigungu_range(hs_code, sido_code, yymm_start, yymm_end, use_cache=use_cache)
        except Exception as e:
            print(f'  {sido_name} ({sido_code}): ERR {str(e)[:60]}')
            continue

        # 시군구별 누적
        local: dict[str, dict] = {}
        for r in recs:
            slot = local.setdefault(r.sgg_name, {
                'total': 0, 'periods': set(),
                'latest_p': '', 'latest_kusd': 0,
            })
            slot['total'] += r.exp_kusd
            slot['periods'].add(r.period)
            if r.period > slot['latest_p']:
                slot['latest_p'] = r.period
                slot['latest_kusd'] = r.exp_kusd

        for sgg_name, d in local.items():
            by_sgg[(sido_code, sgg_name)] = {
                'sido_code': sido_code,
                'sido_name': sido_name,
                'sgg_name': sgg_name,
                'sgg_keyword': _extract_sgg_keyword(sgg_name),
                'exp_kusd_total': d['total'],
                'months_with_data': len(d['periods']),
                'latest_exp_kusd': d['latest_kusd'],
                'latest_period': d['latest_p'],
            }

        n_sgg = len(local)
        total = sum(d['total'] for d in local.values())
        print(f'  {sido_name[:10]:<12} ({sido_code}): {n_sgg:>3}개 시군구, 합계 ${total/1000:>10,.1f}M')

    ranked = sorted(by_sgg.values(), key=lambda x: x['exp_kusd_total'], reverse=True)
    return ranked


def print_ranking(ranked: list[dict], top_n: int = 10) -> None:
    print()
    print('=' * 100)
    print(f'{"#":<4}{"시도":<14}{"시군구(원본)":<28}{"키워드":<10}'
          f'{"누적수출(M)":>14}{"최신월(M)":>13}{"월수":>8}{"점유":>8}')
    print('=' * 100)

    grand = sum(r['exp_kusd_total'] for r in ranked) or 1
    for i, r in enumerate(ranked[:top_n], 1):
        share = r['exp_kusd_total'] / grand * 100
        print(
            f'{i:<4}{r["sido_name"][:12]:<14}{r["sgg_name"][:26]:<28}'
            f'{r["sgg_keyword"][:8]:<10}'
            f'{r["exp_kusd_total"]/1000:>13,.1f}'
            f'{r["latest_exp_kusd"]/1000:>12,.1f}'
            f'{r["months_with_data"]:>5}/12'
            f'{share:>7.1f}%'
        )
    print('=' * 100)
    top_sum = sum(r['exp_kusd_total'] for r in ranked[:top_n])
    print(f'전체 시군구: {len(ranked)}개  /  Top-{top_n} 점유율: {top_sum/grand*100:.1f}%  '
          f'/  전체 합계: ${grand/1000:,.1f}M')


def _confidence_for(exp_kusd: int) -> str:
    if exp_kusd > 50_000:
        return 'high'
    if exp_kusd > 5_000:
        return 'mid'
    return 'low'


def prompt_select(ranked: list[dict], top_n: int) -> list[int]:
    """사용자에게 매핑할 시군구 선택 (복수 가능)."""
    print()
    print('매핑할 시군구를 선택하세요.')
    print('  - 단일:       "1"')
    print('  - 복수:       "1,3"')
    print('  - 범위:       "1-3"')
    print('  - 자동 Top-N: "auto:3"')
    print('  - 취소:       엔터')
    try:
        s = input('선택: ').strip().lower()
    except EOFError:
        return []
    if not s:
        return []

    if s.startswith('auto:'):
        try:
            n = int(s.split(':', 1)[1])
            return list(range(min(n, top_n, len(ranked))))
        except ValueError:
            return []

    out: set[int] = set()
    try:
        for token in s.split(','):
            token = token.strip()
            if '-' in token:
                a, b = token.split('-', 1)
                for i in range(int(a), int(b) + 1):
                    out.add(i - 1)
            else:
                out.add(int(token) - 1)
        return sorted(i for i in out if 0 <= i < min(top_n, len(ranked)))
    except ValueError:
        print('  잘못된 입력 — 취소.')
        return []


def add_selected(stock_code: str, stock_name: str, hs_code: str, hs_desc: str,
                 selected: list[dict]) -> int:
    """선택된 시군구를 stock_mapping.csv에 append (중복 방지)."""
    existing = load_mappings()
    today = datetime.now().strftime('%Y-%m-%d')
    added = 0

    for r in selected:
        dup = [
            m for m in existing
            if m.stock_code == stock_code
            and m.hs_code == hs_code
            and m.sido_code == r['sido_code']
            and m.sgg_keyword == r['sgg_keyword']
        ]
        if dup:
            print(f'  ⚠️ 중복 — 이미 존재: {stock_name} HS{hs_code} {r["sido_name"]}/{r["sgg_keyword"]}')
            continue

        confidence = _confidence_for(r['exp_kusd_total'])
        note = (
            f'discover_sgg 자동: 최근 12개월 ${r["exp_kusd_total"]/1000:,.1f}M '
            f'({r["months_with_data"]}/12mo)'
        )

        row = MappingRow(
            stock_code=stock_code, stock_name=stock_name,
            hs_code=hs_code, hs_desc=hs_desc,
            sido_code=r['sido_code'], sgg_keyword=r['sgg_keyword'],
            weight=1.0, confidence=confidence,
            note=note, added_date=today,
        )
        add_mapping_row(row)
        print(
            f'  ✅ 추가: {stock_name} HS{hs_code} '
            f'{r["sido_name"]}/{r["sgg_keyword"]} ({confidence}, '
            f'${r["exp_kusd_total"]/1000:,.1f}M)'
        )
        added += 1
    return added


def main() -> None:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    p = argparse.ArgumentParser(description='종목 → 유력 시군구 자동 디스커버리')
    p.add_argument('--code', required=True, help='종목코드 6자리')
    p.add_argument('--name', required=True, help='종목명')
    p.add_argument('--hs', required=True, help='HS코드 6자리')
    p.add_argument('--hs-desc', default='', help='HS 한글품명 (선택)')
    p.add_argument('--top-n', type=int, default=10, help='상위 N개 표시 (default 10)')
    p.add_argument('--months', type=int, default=12, help='최근 N개월 합산 (default 12)')
    p.add_argument('--auto-top', type=int, default=0,
                   help='prompt 없이 Top-N 자동 추가 (default 0=대화형)')
    p.add_argument('--no-cache', action='store_true', help='캐시 무시 강제 fetch')
    args = p.parse_args()

    print(f'\n{"=" * 60}')
    print(f'  시군구 디스커버리: {args.name} ({args.code})')
    print(f'  HS: {args.hs}{" (" + args.hs_desc + ")" if args.hs_desc else ""}')
    print(f'  최근 {args.months}개월 누적 기준')
    print('=' * 60)

    ranked = discover_sgg(args.hs, months=args.months, use_cache=not args.no_cache)
    if not ranked:
        print('  데이터 없음 — HS코드 재확인 필요.')
        return

    print_ranking(ranked, top_n=args.top_n)

    if args.auto_top > 0:
        idxs = list(range(min(args.auto_top, args.top_n, len(ranked))))
        print(f'\n자동 모드: Top-{len(idxs)} 추가 중...')
    else:
        idxs = prompt_select(ranked, args.top_n)
        if not idxs:
            print('  취소됨.')
            return

    selected = [ranked[i] for i in idxs]
    print()
    n = add_selected(args.code, args.name, args.hs, args.hs_desc, selected)
    print(f'\n총 {n}개 매핑 추가됨.')
    if n > 0:
        print('  → 다음 runner.py 실행 시 새 매핑이 트래킹에 포함됩니다.')
        print('  → scripts\\run_manual.bat')


if __name__ == '__main__':
    main()
