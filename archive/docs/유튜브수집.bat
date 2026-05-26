@echo off
cd /d "%~dp0"
echo [YouTube 수집 시작] 스니커즈 5개 아이템
echo.
py scripts/collect_youtube.py
echo.
echo [완료] 창을 닫아도 됩니다.
pause
