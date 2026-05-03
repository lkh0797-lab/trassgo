"""전체 오케스트레이션 — 매핑된 종목 일괄 fetch + 시그널 + 엑셀 + 알림"""
from __future__ import annotations

import logging
import sys
from datetime import datetime
from pathlib import Path

from .customs_api import latest_available_yymm, load_settings
from .dashboard import build_dashboard
from .mapper import build_stock_timeseries, group_by_stock, load_mappings
from .signals import compute_all_signals, format_signal_text
from .telegram_bot import send_run_summary, send_signals


PROJECT_ROOT = Path(__file__).resolve().parent.parent
LOG_DIR = PROJECT_ROOT / 'logs'


def _setup_logger() -> logging.Logger:
    LOG_DIR.mkdir(exist_ok=True)
    log_path = LOG_DIR / f'run_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log'
    logger = logging.getLogger('tracker')
    logger.setLevel(logging.INFO)
    if not logger.handlers:
        fh = logging.FileHandler(log_path, encoding='utf-8')
        sh = logging.StreamHandler(sys.stdout)
        fmt = logging.Formatter('%(asctime)s [%(levelname)s] %(message)s', '%H:%M:%S')
        fh.setFormatter(fmt)
        sh.setFormatter(fmt)
        logger.addHandler(fh)
        logger.addHandler(sh)
    return logger


def _yymm_back(yymm: str, months: int) -> str:
    y, m = int(yymm[:4]), int(yymm[4:])
    idx = y * 12 + m - 1 - months
    return f'{idx // 12:04d}{idx % 12 + 1:02d}'


def run(*, force_fetch: bool = False, send_telegram: bool = True) -> Path:
    log = _setup_logger()
    log.info('=' * 60)
    log.info('수출입통계 트래커 실행 시작')

    settings = load_settings()
    fetch_months = settings['data']['fetch_months']

    log.info('관세청 API 최신 데이터 시점 탐지...')
    yymm_end = latest_available_yymm()
    yymm_start = _yymm_back(yymm_end, fetch_months - 1)
    log.info(f'  데이터 기간: {yymm_start} ~ {yymm_end} ({fetch_months}개월)')

    mappings = load_mappings()
    log.info(f'매핑 로드: {len(mappings)}건 ({len(group_by_stock(mappings))} 종목)')

    log.info('종목별 시계열 구축 중 (캐시 활용)...')
    ts_by_stock = build_stock_timeseries(
        mappings, yymm_start, yymm_end, use_cache=not force_fetch,
    )
    log.info(f'  완료: {len(ts_by_stock)} 종목 시계열')

    log.info('시그널 계산...')
    signals = compute_all_signals(ts_by_stock)
    log.info(f'  발생 시그널: {len(signals)}건')
    for s in signals:
        log.info(f'    {s.signal_type:10s} {s.stock_name} ({s.stock_code})')

    log.info('엑셀 대시보드 생성...')
    out_path = build_dashboard(ts_by_stock, mappings, yymm_start, yymm_end)
    log.info(f'  저장: {out_path}')

    if send_telegram and settings['telegram']['enabled']:
        log.info('텔레그램 알림 전송...')
        n_sent = send_signals(signals)
        send_run_summary(len(ts_by_stock), len(signals), str(out_path), yymm_end)
        log.info(f'  전송 완료: {n_sent}건')
    else:
        log.info('텔레그램 비활성 — 콘솔 출력만')
        for s in signals:
            log.info('\n' + format_signal_text(s))

    log.info('실행 완료')
    return out_path


if __name__ == '__main__':
    import argparse
    p = argparse.ArgumentParser()
    p.add_argument('--force-fetch', action='store_true', help='캐시 무시 강제 fetch')
    p.add_argument('--no-telegram', action='store_true', help='텔레그램 알림 끄기')
    args = p.parse_args()
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    run(force_fetch=args.force_fetch, send_telegram=not args.no_telegram)
