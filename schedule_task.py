import subprocess
import sys
from datetime import datetime, timedelta
from pathlib import Path

BASE_DIR = Path(__file__).parent


def main():
    print()
    print("=" * 55)
    print("   Создание задачи в планировщике Windows")
    print("=" * 55)
    print()

    print("Шаг 1 — Трансформация прайса...")
    from run import pick_input_file
    from transform import transform
    input_path = pick_input_file()
    output_path = transform(input_path)

    print()
    time_str = input("  Введите время отправки (например 10:00): ").strip()

    try:
        hh, mm = map(int, time_str.split(':'))
    except ValueError:
        print(f"  Неверный формат: {time_str}. Используйте HH:MM")
        sys.exit(1)

    now = datetime.now()
    target = now.replace(hour=hh, minute=mm, second=0, microsecond=0)
    if target <= now:
        target += timedelta(days=1)

    task_name = f"БытТехОпт_{target.strftime('%Y%m%d_%H%M')}"
    python_exe = sys.executable
    send_script = str(BASE_DIR / 'telegram_send.py')
    end_boundary = target + timedelta(hours=1)

    # Создаём задачу через PowerShell (не зависит от локали Windows)
    # EndBoundary обязателен при использовании DeleteExpiredTaskAfter
    ps_script = f"""
$ErrorActionPreference = 'Stop'
$action  = New-ScheduledTaskAction -Execute '{python_exe}' -Argument '"{send_script}" "{output_path}"'
$trigger = New-ScheduledTaskTrigger -Once -At '{target.strftime("%Y-%m-%dT%H:%M:00")}'
$trigger.EndBoundary = '{end_boundary.strftime("%Y-%m-%dT%H:%M:00")}'
$settings = New-ScheduledTaskSettingsSet -DeleteExpiredTaskAfter (New-TimeSpan -Minutes 1)
Register-ScheduledTask -TaskName '{task_name}' -Action $action -Trigger $trigger -Settings $settings -Force | Out-Null
Write-Host 'OK'
"""

    result = subprocess.run(
        ['powershell', '-NoProfile', '-Command', ps_script],
        capture_output=True, text=True
    )

    if 'OK' in result.stdout and not result.stderr.strip():
        print()
        print(f"  Готово! Задача создана в планировщике Windows.")
        print(f"  Отправка: {target.strftime('%d.%m.%Y в %H:%M')}")
        print(f"  Это окно можно закрыть.")
        print()
    else:
        print(f"  Ошибка создания задачи:")
        print(result.stderr or result.stdout or "(нет вывода)")
        sys.exit(1)


if __name__ == '__main__':
    main()
