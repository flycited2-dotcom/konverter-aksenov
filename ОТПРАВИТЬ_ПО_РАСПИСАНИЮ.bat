@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo   Отложенная отправка прайса в Telegram
echo ============================================
echo.
set /p SEND_TIME="Введите время отправки (например 10:00): "
echo.
echo Прайс создаётся сейчас, отправка в %SEND_TIME%
python run.py --send-at %SEND_TIME%
pause
