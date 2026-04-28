import os
import sys
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
    files = sorted(input_dir.glob('*.xlsx'), key=os.path.getmtime, reverse=True)
    if not files:
        print("=" * 55)
        print("  Нет файлов в папке input/")
        print("  Положите прайс поставщика (.xlsx) в папку:")
        print(f"  {input_dir}")
        print("  и запустите run.py снова")
        print("=" * 55)
        sys.exit(1)

    print(f"Найден файл: {files[0].name}")
    if len(files) > 1:
        print(f"  (найдено {len(files)} файлов, берём самый новый)")
    return str(files[0])


def ask_yes_no(question: str) -> bool:
    ans = input(f"{question} (y/n): ").strip().lower()
    return ans in ('y', 'д', 'да', 'yes', '1')


def main():
    print()
    print("=" * 55)
    print("   Excel-автоматизация для оптовой торговли")
    print("=" * 55)

    input_path = pick_input_file(sys.argv[1] if len(sys.argv) > 1 else None)

    print()
    print("Шаг 1 — Трансформация прайса...")
    from transform import transform
    output_path = transform(input_path)

    print()
    print("Шаг 2 — Отправка в Telegram")
    if ask_yes_no("Отправить файл в Telegram?"):
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
