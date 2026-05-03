"""종목 추가 헬퍼 — 대화형 / 인자 기반 둘 다 지원

사용법:
  python -m src.add_stock --code 042700 --name 한미반도체 --hs 848620 --sido 28 --sgg 인천 --note 'HBM 본더'
  또는
  python -m src.add_stock  (인자 없으면 대화형 입력)

입력 후 자동:
  1. 시도 코드 검증 (sido_codes.csv 대조)
  2. 데이터 1개월 fetch — sgg_keyword가 실제로 매칭되는지, 수출액 규모가 합리적인지 표시
  3. 사용자 확인 후 stock_mapping.csv에 append
"""
from __future__ import annotations

import argparse
import io
import sys
from datetime import datetime

from .customs_api import (
    SigunguQuery, fetch_sigungu, latest_available_yymm, filter_by_sgg_keyword,
)
from .mapper import MappingRow, add_mapping_row, load_mappings, load_sido_codes


def prompt(label: str, default: str = '') -> str:
    s = input(f'{label}{(" [" + default + "]") if default else ""}: ').strip()
    return s or default


def verify_sido(sido_code: str) -> str:
    sido_map = load_sido_codes()
    if sido_code not in sido_map:
        raise ValueError(f'시도 코드 {sido_code}가 sido_codes.csv에 없음. 11/26/27/.../43/47 중 선택')
    return sido_map[sido_code]


def probe_data(hs_code: str, sido_code: str, sgg_keyword: str) -> dict:
    """1개월 fetch 후 매칭 검증 결과 반환"""
    yymm = latest_available_yymm()
    recs = fetch_sigungu(SigunguQuery(
        hs_code=hs_code, sido_code=sido_code,
        yymm_start=yymm, yymm_end=yymm,
    ))
    matched = filter_by_sgg_keyword(recs, sgg_keyword)
    return {
        'yymm': yymm,
        'total_sgg_in_sido': len(recs),
        'matched_count': len(matched),
        'matched_sgg_names': sorted(set(r.sgg_name for r in matched)),
        'total_exp_kusd': sum(r.exp_kusd for r in matched),
        'sample_records': matched[:5],
    }


def add_stock(stock_code: str, stock_name: str, hs_code: str, hs_desc: str,
              sido_code: str, sgg_keyword: str, weight: float = 1.0,
              confidence: str = 'mid', note: str = '', skip_confirm: bool = False) -> bool:
    print()
    print('=' * 60)
    print(f'  종목 추가 검증: {stock_name} ({stock_code})')
    print('=' * 60)

    sido_name = verify_sido(sido_code)
    print(f'  시도: {sido_name} ({sido_code})')
    print(f'  시군구 키워드: "{sgg_keyword}"')
    print(f'  HS: {hs_code} ({hs_desc})')
    print(f'  가중치: {weight}, 신뢰도: {confidence}')
    print()

    print('관세청 API 호출 — 최신월 데이터 검증 중...')
    probe = probe_data(hs_code, sido_code, sgg_keyword)
    print()
    print(f'  검증월: {probe["yymm"]}')
    print(f'  해당 시도의 시군구 데이터: {probe["total_sgg_in_sido"]}건')
    print(f'  키워드 매칭 시군구: {probe["matched_count"]}건')
    if probe['matched_sgg_names']:
        for n in probe['matched_sgg_names']:
            print(f'    - {n}')
    else:
        print('  ⚠️ 매칭 없음 — sgg_keyword 다시 확인 필요')
    print(f'  매칭된 행 수출액 합계: ${probe["total_exp_kusd"]/1000:,.1f}M (천 USD 단위)')
    print()

    existing = load_mappings()
    dup = [m for m in existing if m.stock_code == stock_code and m.hs_code == hs_code
           and m.sido_code == sido_code and m.sgg_keyword == sgg_keyword]
    if dup:
        print(f'  ⚠️ 중복 매핑 — 이미 동일 행 존재 ({dup[0].added_date} 추가).')
        return False

    if not skip_confirm:
        ok = prompt('이대로 매핑 추가할까요? (y/n)', 'n')
        if ok.lower() != 'y':
            print('  취소됨.')
            return False

    row = MappingRow(
        stock_code=stock_code,
        stock_name=stock_name,
        hs_code=hs_code,
        hs_desc=hs_desc,
        sido_code=sido_code,
        sgg_keyword=sgg_keyword,
        weight=weight,
        confidence=confidence,
        note=note,
        added_date=datetime.now().strftime('%Y-%m-%d'),
    )
    add_mapping_row(row)
    print(f'  ✅ 매핑 추가 완료. 총 매핑 {len(existing)+1}건')
    return True


def main():
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    p = argparse.ArgumentParser()
    p.add_argument('--code')
    p.add_argument('--name')
    p.add_argument('--hs')
    p.add_argument('--hs-desc', default='')
    p.add_argument('--sido')
    p.add_argument('--sgg')
    p.add_argument('--weight', type=float, default=1.0)
    p.add_argument('--confidence', default='mid')
    p.add_argument('--note', default='')
    p.add_argument('-y', '--yes', action='store_true')
    args = p.parse_args()

    if not all([args.code, args.name, args.hs, args.sido, args.sgg]):
        print('대화형 입력 모드')
        args.code = prompt('종목코드 (6자리)')
        args.name = prompt('종목명')
        args.hs = prompt('HS코드 (6자리)')
        args.hs_desc = prompt('HS 한글품명', '미입력')
        args.sido = prompt('시도 코드 (예: 43=충북, 47=경북, 28=인천)')
        args.sgg = prompt('시군구 키워드 (예: 청주시, 포항시, 남동구)')
        args.note = prompt('메모', '')

    add_stock(
        stock_code=args.code, stock_name=args.name,
        hs_code=args.hs, hs_desc=args.hs_desc,
        sido_code=args.sido, sgg_keyword=args.sgg,
        weight=args.weight, confidence=args.confidence,
        note=args.note, skip_confirm=args.yes,
    )


if __name__ == '__main__':
    main()
