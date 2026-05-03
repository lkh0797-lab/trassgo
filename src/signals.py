"""시그널 계산 — MOMENTUM / BREAKOUT / WARN"""
from __future__ import annotations

from dataclasses import dataclass

from .customs_api import load_settings
from .mapper import StockTimeseries


@dataclass
class Signal:
    stock_code: str
    stock_name: str
    signal_type: str       # MOMENTUM, BREAKOUT, WARN
    period: str
    value_kusd: int
    yoy: float | None
    detail: str            # 사람이 읽을 수 있는 설명


def compute_signals(ts: StockTimeseries) -> list[Signal]:
    cfg = load_settings()['signals']
    sorted_p = sorted(ts.by_period.keys())
    if not sorted_p:
        return []

    latest = sorted_p[-1]
    slot = ts.by_period[latest]
    out: list[Signal] = []

    # MOMENTUM
    yoy_thr = cfg['momentum']['yoy_threshold']
    n_pos = cfg['momentum']['consecutive_positive_months']
    if slot['yoy'] is not None and slot['yoy'] >= yoy_thr:
        last_n = sorted_p[-n_pos:]
        if all(ts.by_period[p]['yoy'] is not None and ts.by_period[p]['yoy'] > 0 for p in last_n):
            out.append(Signal(
                stock_code=ts.stock_code, stock_name=ts.stock_name,
                signal_type='MOMENTUM', period=latest,
                value_kusd=slot['exp_kusd'], yoy=slot['yoy'],
                detail=f'YoY +{slot["yoy"]*100:.1f}% 및 {n_pos}개월 연속 YoY 양수',
            ))

    # BREAKOUT
    lookback = cfg['breakout']['lookback_months']
    history = sorted_p[-(lookback + 1):-1] if len(sorted_p) > 1 else []
    if history:
        max_hist = max(ts.by_period[p]['exp_kusd'] for p in history)
        if slot['exp_kusd'] > max_hist:
            out.append(Signal(
                stock_code=ts.stock_code, stock_name=ts.stock_name,
                signal_type='BREAKOUT', period=latest,
                value_kusd=slot['exp_kusd'], yoy=slot['yoy'],
                detail=f'직전 {lookback}개월 최고치 ${max_hist/1000:,.1f}M 갱신',
            ))

    # WARN
    yoy_thr_w = cfg['warn']['yoy_threshold']
    n_neg = cfg['warn']['consecutive_negative_months']
    if slot['yoy'] is not None and slot['yoy'] <= yoy_thr_w:
        if len(sorted_p) >= n_neg:
            last_n = sorted_p[-n_neg:]
            decreasing = all(
                ts.by_period[last_n[i]]['exp_kusd'] < ts.by_period[last_n[i-1]]['exp_kusd']
                for i in range(1, n_neg)
            )
            if decreasing:
                out.append(Signal(
                    stock_code=ts.stock_code, stock_name=ts.stock_name,
                    signal_type='WARN', period=latest,
                    value_kusd=slot['exp_kusd'], yoy=slot['yoy'],
                    detail=f'YoY {slot["yoy"]*100:.1f}% 및 {n_neg}개월 연속 감소',
                ))

    return out


def compute_all_signals(ts_by_stock: dict[str, StockTimeseries]) -> list[Signal]:
    out: list[Signal] = []
    for ts in ts_by_stock.values():
        out.extend(compute_signals(ts))
    return out


def format_signal_text(sig: Signal) -> str:
    icon = {'MOMENTUM': '🚀', 'BREAKOUT': '⭐', 'WARN': '⚠️'}.get(sig.signal_type, '📊')
    yoy_str = f'{sig.yoy*100:+.1f}%' if sig.yoy is not None else 'N/A'
    return (
        f'{icon} [{sig.signal_type}] {sig.stock_name} ({sig.stock_code})\n'
        f'   {sig.period}  ${sig.value_kusd/1000:,.1f}M (YoY {yoy_str})\n'
        f'   {sig.detail}'
    )
