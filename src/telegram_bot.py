"""텔레그램 알림 - 토큰 미설정 시에도 print만 하고 정상 동작"""
from __future__ import annotations

import json
import urllib.parse
import urllib.request
from pathlib import Path

from .customs_api import load_settings


PROJECT_ROOT = Path(__file__).resolve().parent.parent


def load_telegram_credentials() -> tuple[str | None, str | None]:
    env_path = PROJECT_ROOT / 'config' / '.env'
    token = chat_id = None
    with open(env_path, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if line.startswith('TELEGRAM_BOT_TOKEN='):
                v = line.split('=', 1)[1].strip()
                token = v if v else None
            elif line.startswith('TELEGRAM_CHAT_ID='):
                v = line.split('=', 1)[1].strip()
                chat_id = v if v else None
    return token, chat_id


def send_message(text: str) -> bool:
    """텔레그램 메시지 전송. 토큰 미설정 시 콘솔에 출력만."""
    settings = load_settings()
    if not settings['telegram']['enabled']:
        print(f'[TELEGRAM disabled] {text[:200]}')
        return False

    token, chat_id = load_telegram_credentials()
    if not (token and chat_id):
        print(f'[TELEGRAM no credentials] {text[:200]}')
        return False

    url = f'https://api.telegram.org/bot{token}/sendMessage'
    data = urllib.parse.urlencode({
        'chat_id': chat_id,
        'text': text,
        'parse_mode': 'HTML',
    }).encode('utf-8')

    try:
        req = urllib.request.Request(url, data=data)
        with urllib.request.urlopen(req, timeout=15) as resp:
            result = json.loads(resp.read().decode('utf-8'))
            return result.get('ok', False)
    except Exception as e:
        print(f'[TELEGRAM error] {e}')
        return False


def send_signals(signals: list, immediate_types: list[str] | None = None) -> int:
    """시그널 발송 — 즉시 푸시 타입은 개별 메시지, 나머지는 묶어서 1건."""
    settings = load_settings()
    immediate_types = immediate_types or settings['telegram']['push_immediate']

    sent = 0
    immediate = [s for s in signals if s.signal_type in immediate_types]
    summary = [s for s in signals if s.signal_type not in immediate_types]

    from .signals import format_signal_text

    for s in immediate:
        if send_message(format_signal_text(s)):
            sent += 1

    if summary:
        from collections import defaultdict
        groups = defaultdict(list)
        for s in summary:
            groups[s.signal_type].append(s)
        text = '📊 <b>월간 시그널 요약</b>\n\n'
        for typ, sigs in groups.items():
            text += f'<b>[{typ}] ({len(sigs)}건)</b>\n'
            for s in sigs:
                text += f'  • {s.stock_name} ({s.stock_code}) - YoY {(s.yoy or 0)*100:+.1f}%\n'
            text += '\n'
        if send_message(text):
            sent += len(summary)

    return sent


def send_run_summary(n_stocks: int, n_signals: int, output_path: str, latest_yymm: str) -> bool:
    text = (
        f'✅ <b>수출입통계 트래커 실행 완료</b>\n'
        f'기준월: {latest_yymm}\n'
        f'분석 종목: {n_stocks}개\n'
        f'시그널 발생: {n_signals}건\n'
        f'엑셀: {output_path}'
    )
    return send_message(text)
