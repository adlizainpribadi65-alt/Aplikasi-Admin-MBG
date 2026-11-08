@echo off
echo ============================================
echo   Admin MBG - Local Server
echo ============================================
echo.
echo Server berjalan di: http://localhost:3000
echo Tekan Ctrl+C untuk menghentikan server
echo.
cd /d "%~dp0"
start "" "http://localhost:3000"
python -m http.server 3000
pause

