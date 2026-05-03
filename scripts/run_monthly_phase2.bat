@echo off
chcp 65001 >nul
cd /d "%~dp0\.."
REM 매월 16일 실행 - 정밀 잠정치 (캐시 무효화)
python -m src.runner --force-fetch
