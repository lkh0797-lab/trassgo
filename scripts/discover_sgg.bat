@echo off
chcp 65001 >nul
cd /d "%~dp0\.."
echo === 시군구 디스커버리 (종목 + HS코드 → 유력 시군구 Top-N) ===
echo.
set /p CODE=종목코드(6자리):
set /p NAME=종목명:
set /p HS=HS코드(6자리):
set /p HSDESC=HS 한글품명(선택, 엔터 가능):
echo.
python -m src.discover_sgg --code %CODE% --name %NAME% --hs %HS% --hs-desc "%HSDESC%"
pause
