"""종목 매핑 마스터 로드 + 다종목 시계열 합산 로직"""
from __future__ import annotations

import csv
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path

from .customs_api import (
    SigunguRecord, fetch_sigungu_range, filter_by_sgg_keyword,
)


PROJECT_ROOT = Path(__file__).resolve().parent.parent
MAPPING_CSV = PROJECT_ROOT / 'data' / 'master' / 'stock_mapping.csv'
SIDO_CSV = PROJECT_ROOT / 'config' / 'sido_codes.csv'


@dataclass
class MappingRow:
    stock_code: str
    stock_name: str
    hs_code: str
    hs_desc: str
    sido_code: str
    sgg_keyword: str
    weight: float
    confidence: str
    note: str
    added_date: str


@dataclass
class StockTimeseries:
    stock_code: str
    stock_name: str
    by_period: dict[str, dict] = field(default_factory=dict)  # {YYYY.MM: {exp_kusd, exp_count, ...}}
    mappings: list[MappingRow] = field(default_factory=list)


def load_mappings() -> list[MappingRow]:
    rows: list[MappingRow] = []
    with open(MAPPING_CSV, 'r', encoding='utf-8-sig', newline='') as f:
        reader = csv.DictReader(f)
        for r in reader:
            rows.append(MappingRow(
                stock_code=r['stock_code'].strip(),
                stock_name=r['stock_name'].strip(),
                hs_code=r['hs_code'].strip(),
                hs_desc=r['hs_desc'].strip(),
                sido_code=r['sido_code'].strip(),
                sgg_keyword=r['sgg_keyword'].strip(),
                weight=float(r['weight']),
                confidence=r['confidence'].strip(),
                note=r['note'].strip(),
                added_date=r['added_date'].strip(),
            ))
    return rows


def load_sido_codes() -> dict[str, str]:
    out = {}
    with open(SIDO_CSV, 'r', encoding='utf-8-sig', newline='') as f:
        reader = csv.DictReader(f)
        for r in reader:
            out[r['sido_code'].strip()] = r['sido_name'].strip()
    return out


def group_by_stock(mappings: list[MappingRow]) -> dict[str, list[MappingRow]]:
    out: dict[str, list[MappingRow]] = defaultdict(list)
    for m in mappings:
        out[m.stock_code].append(m)
    return out


def build_stock_timeseries(mappings: list[MappingRow], yymm_start: str, yymm_end: str,
                           use_cache: bool = True) -> dict[str, StockTimeseries]:
    """매핑된 종목별로 시계열 합산 — 가중치 적용 + 시군구 키워드 필터"""
    by_stock: dict[str, StockTimeseries] = {}

    # 동일 (HS, sido) 조합은 한 번만 fetch
    fetch_cache: dict[tuple[str, str], list[SigunguRecord]] = {}

    for m in mappings:
        key = (m.hs_code, m.sido_code)
        if key not in fetch_cache:
            fetch_cache[key] = fetch_sigungu_range(
                m.hs_code, m.sido_code, yymm_start, yymm_end, use_cache=use_cache,
            )
        recs = filter_by_sgg_keyword(fetch_cache[key], m.sgg_keyword)

        ts = by_stock.setdefault(
            m.stock_code,
            StockTimeseries(stock_code=m.stock_code, stock_name=m.stock_name),
        )
        ts.mappings.append(m)

        for r in recs:
            slot = ts.by_period.setdefault(r.period, {
                'exp_kusd': 0, 'exp_count': 0, 'imp_kusd': 0, 'imp_count': 0, 'balance_kusd': 0,
            })
            slot['exp_kusd'] += int(r.exp_kusd * m.weight)
            slot['exp_count'] += int(r.exp_count * m.weight)
            slot['imp_kusd'] += int(r.imp_kusd * m.weight)
            slot['imp_count'] += int(r.imp_count * m.weight)
            slot['balance_kusd'] += int(r.balance_kusd * m.weight)

    # 시그널 계산용 정렬 + MoM/YoY 부여
    for ts in by_stock.values():
        sorted_periods = sorted(ts.by_period.keys())
        for i, p in enumerate(sorted_periods):
            slot = ts.by_period[p]
            cur = slot['exp_kusd']
            slot['mom'] = (
                cur / ts.by_period[sorted_periods[i-1]]['exp_kusd'] - 1
                if i >= 1 and ts.by_period[sorted_periods[i-1]]['exp_kusd'] else None
            )
            slot['yoy'] = (
                cur / ts.by_period[sorted_periods[i-12]]['exp_kusd'] - 1
                if i >= 12 and ts.by_period[sorted_periods[i-12]]['exp_kusd'] else None
            )

    return by_stock


def add_mapping_row(row: MappingRow) -> None:
    """매핑 마스터에 한 줄 추가 (append)."""
    exists = MAPPING_CSV.exists()
    with open(MAPPING_CSV, 'a' if exists else 'w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f)
        if not exists:
            writer.writerow([
                'stock_code', 'stock_name', 'hs_code', 'hs_desc',
                'sido_code', 'sgg_keyword', 'weight', 'confidence',
                'note', 'added_date',
            ])
        writer.writerow([
            row.stock_code, row.stock_name, row.hs_code, row.hs_desc,
            row.sido_code, row.sgg_keyword, row.weight, row.confidence,
            row.note, row.added_date,
        ])


if __name__ == '__main__':
    maps = load_mappings()
    print(f'Loaded {len(maps)} mappings, {len(group_by_stock(maps))} unique stocks')
    for m in maps:
        print(f'  {m.stock_code} {m.stock_name} HS{m.hs_code} sido{m.sido_code}/{m.sgg_keyword} w={m.weight}')

    print('\nBuilding 36-month timeseries...')
    ts_by_stock = build_stock_timeseries(maps, '202304', '202603')
    for code, ts in ts_by_stock.items():
        latest = sorted(ts.by_period.keys())[-1]
        slot = ts.by_period[latest]
        print(f'\n{ts.stock_name} ({code}) latest={latest} exp={slot["exp_kusd"]:,} kUSD '
              f'YoY={slot["yoy"]*100:+.1f}%' if slot['yoy'] else f'  exp={slot["exp_kusd"]:,} kUSD')
