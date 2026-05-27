import os
import sys
import time
import unicodedata
from datetime import datetime, timedelta
from pathlib import Path

BASE_DIR = Path(__file__).parent


def pick_input_file(arg=None):
    if arg:
        p = Path(arg)
        if not p.exists():
            print(f"Файл не найден: {arg}")
            sys.exit(1)
        return str(p)

    input_dir = BASE_DIR / 'input'
    input_dir.mkdir(exist_ok=True)
    all_files = [f for f in list(input_dir.glob('*.xlsx')) + list(input_dir.glob('*.xls'))
                 if not f.name.startswith('~$')]
    files = sorted(all_files, key=os.path.getmtime, reverse=True)
    if not files:
        print("=" * 55)
        print("  Нет файлов в папке input/")
        print("  Положите прайс поставщика (.xlsx или .xls) в папку:")
        print(f"  {input_dir}")
        print("  и запустите run.py снова")
        print("=" * 55)
        sys.exit(1)

    display_name = unicodedata.normalize('NFC', files[0].name)
    print(f"Найден файл: {display_name}")
    if len(files) > 1:
        print(f"  (найдено {len(files)} файлов, берём самый новый)")
    return str(files[0])


def ask_yes_no(question: str) -> bool:
    ans = input(f"{question} (y/n): ").strip().lower()
    return ans in ('y', 'д', 'да', 'yes', '1')


def wait_until(send_at: str):
    """Ожидает до указанного времени HH:MM. Если уже прошло — ждёт до следующего дня."""
    try:
        hh, mm = map(int, send_at.split(':'))
    except ValueError:
        print(f"Неверный формат времени: {send_at}. Используйте HH:MM (например 10:00)")
        sys.exit(1)

    now = datetime.now()
    target = now.replace(hour=hh, minute=mm, second=0, microsecond=0)
    if target <= now:
        target += timedelta(days=1)

    wait_sec = (target - now).total_seconds()
    h = int(wait_sec // 3600)
    m = int((wait_sec % 3600) // 60)

    print()
    print(f"  Запланировано: {target.strftime('%d.%m.%Y в %H:%M')}")
    print(f"  Ожидание: {h} ч {m} мин")
    print("  (оставьте окно открытым)")
    time.sleep(wait_sec)


def main():
    print()
    print("=" * 55)
    print("   Excel-автоматизация для оптовой торговли")
    print("=" * 55)

    # Разбор аргументов: [файл] [--send-at HH:MM] [--immediate]
    args = list(sys.argv[1:])
    send_at = None
    immediate = False

    if '--immediate' in args:
        immediate = True
        args.remove('--immediate')

    if '--schedule' in args:
        args.remove('--schedule')
        send_at = input("  Введите время отправки (например 10:00): ").strip()

    elif '--send-at' in args:
        idx = args.index('--send-at')
        if idx + 1 < len(args):
            send_at = args[idx + 1]
            args = args[:idx] + args[idx + 2:]
        else:
            print("Укажите время после --send-at, например: --send-at 10:00")
            sys.exit(1)

    input_path = pick_input_file(args[0] if args else None)

    print()
    print("Шаг 1 — Трансформация прайса...")
    from transform import transform
    output_path = transform(input_path)

    print()
    print("Шаг 2 — Отправка в Telegram")

    if send_at:
        wait_until(send_at)
        from telegram_send import send_file
        send_file(output_path)
    elif immediate:
        from telegram_send import send_file
        send_file(output_path)
    elif ask_yes_no("Отправить файл в Telegram?"):
        from telegram_send import send_file
        send_file(output_path)
    else:
        print("Пропускаем отправку в Telegram")

    print()
    print("=" * 55)
    print(f"  Готово! Файл сохранён:")
    print(f"  {output_path}")
    print("=" * 55)
    print()


if __name__ == '__main__':
    main()
