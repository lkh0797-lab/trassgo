@echo off
chcp 65001 >nul
cd /d "%~dp0\.."
echo === 종목 매핑 추가 ===
python -m src.add_stock
pause
