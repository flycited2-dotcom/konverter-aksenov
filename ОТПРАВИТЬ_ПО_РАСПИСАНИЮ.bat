@echo off
chcp 65001 >nul
cd /d "%~dp0"
python schedule_task.py
pause
