@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo   Прайс будет создан и отправлен в Telegram
echo ============================================
python run.py --immediate
pause
