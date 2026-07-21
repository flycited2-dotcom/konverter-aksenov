import json
import sys
from datetime import datetime
from pathlib import Path

BASE_DIR = Path(__file__).parent


def load_secrets(base_dir=BASE_DIR):
    """Читает секреты Telegram (bot_token, channel_id) из secrets.json — файла
    ВНЕ git (он в .gitignore). Возвращает {} если файла нет.
    Секреты держим отдельно от config.json, иначе токен утекает в репозиторий
    и Telegram его автоматически отзывает."""
    path = Path(base_dir) / 'secrets.json'
    if not path.exists():
        return {}
    with open(path, encoding='utf-8') as f:
        return json.load(f)


def send_file(filepath: str, caption: str = None) -> bool:
    try:
        import requests
    except ImportError:
        print("Ошибка: установите пакет requests → pip install requests")
        return False

    secrets = load_secrets()
    token = str(secrets.get('bot_token', '')).strip()
    channel_id = str(secrets.get('channel_id', '')).strip()

    if not token or not channel_id:
        print("Ошибка: не найдены секреты Telegram.")
        print("Создайте файл secrets.json рядом с telegram_send.py (шаблон — secrets.example.json):")
        print('  {"bot_token": "ВАШ_ТОКЕН_ОТ_BOTFATHER", "channel_id": "-100XXXXXXXXXX"}')
        print("Файл secrets.json НЕ коммитится в git (он в .gitignore).")
        return False

    if caption is None:
        date_str = datetime.today().strftime('%d.%m.%Y')
        caption = (
            f"<b>Актуальный прайс от {date_str}</b>\n"
            f"<blockquote>➡️ Генераторы,\n"
            f"➡️ Инструмент,\n"
            f"➡️ Запчасти</blockquote>\n"
            f"\nhttps://splithub.ru/"
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
