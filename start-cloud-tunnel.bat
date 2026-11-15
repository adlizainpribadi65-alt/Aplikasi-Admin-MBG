@echo off
title Cloudflare Tunnel - Admin MBG Public Access
echo ===================================================
echo   ADMIN MBG - CLOUDFLARE PUBLIC INTERNET TUNNEL
echo ===================================================
echo.
echo Sedang menyiapkan akses publik global...
echo Jika cloudflared belum terpasang, unduh dari:
echo https://github.com/cloudflare/cloudflared/releases
echo.

where cloudflared >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [PERHATIAN] Perintah 'cloudflared' belum terpasang di sistem.
    echo.
    echo Cara Pasang Cepat:
    echo 1. Unduh file 'cloudflared-windows-amd64.exe' dari GitHub Cloudflare.
    echo 2. Ubah nama file menjadi 'cloudflared.exe' dan letakkan di folder ini atau C:\Windows\
    echo 3. Jalankan file ini kembali.
    echo.
    pause
    exit /b
)

echo Menghubungkan port 80 (Apache XAMPP) ke Internet Publik...
cloudflared tunnel --url http://localhost:80
pause
