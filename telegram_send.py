import json
import sys
from datetime import datetime
from pathlib import Path

BASE_DIR = Path(__file__).parent


def load_config():
    with open(BASE_DIR / 'config.json', encoding='utf-8') as f:
        return json.load(f)


def send_file(filepath: str, caption: str = None) -> bool:
    try:
        import requests
    except ImportError:
        print("Ошибка: установите пакет requests → pip install requests")
        return False

    config = load_config()
    tg = config.get('telegram', {})
    token = tg.get('bot_token', '').strip()
    channel_id = tg.get('channel_id', '').strip()

    if not token:
        print("Ошибка: заполните telegram.bot_token в config.json")
        return False
    if not channel_id:
        print("Ошибка: заполните telegram.channel_id в config.json")
        return False

    if caption is None:
        date_str = datetime.today().strftime('%d.%m.%Y')
        caption = (
            f"Актуальный прайс от {date_str}\n"
            f"<blockquote>➡️ Генераторы,\n"
            f"➡️ Инструмент,\n"
            f"➡️ Запчасти</blockquote>"
        )

    url = f"https://api.telegram.org/bot{token}/sendDocument"
    filepath = Path(filepath)

    with open(filepath, 'rb') as f:
        resp = requests.post(
            url,
            files={'document': (filepath.name, f,
                   'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')},
            data={'chat_id': channel_id, 'caption': caption, 'parse_mode': 'HTML'},
            timeout=60,
        )

    if resp.status_code == 200:
        print(f"Файл отправлен в Telegram: {filepath.name}")
        return True
    else:
        err = resp.json().get('description', resp.text)
        print(f"Ошибка Telegram ({resp.status_code}): {err}")
        return False


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Использование: python telegram_send.py <файл.xlsx> [подпись]")
        sys.exit(1)
    caption = sys.argv[2] if len(sys.argv) > 2 else None
    ok = send_file(sys.argv[1], caption)
    sys.exit(0 if ok else 1)
