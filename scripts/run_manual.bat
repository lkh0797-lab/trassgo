@echo off
chcp 65001 >nul
cd /d "%~dp0\.."
echo === 수출입통계 트래커 수동 실행 ===
python -m src.runner
echo.
echo 완료. output 폴더에서 엑셀 확인.
pause
