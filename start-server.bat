@echo off
echo ============================================
echo   Admin MBG - XAMPP Local Server
echo ============================================
echo.

:: Cek apakah XAMPP Apache sedang berjalan
tasklist /FI "IMAGENAME eq httpd.exe" 2>NUL | find /I "httpd.exe" >NUL
if errorlevel 1 (
    echo [!] Apache XAMPP belum berjalan!
    echo     Silakan buka XAMPP Control Panel dan Start Apache terlebih dahulu.
    echo.
    pause
    exit /b
)

:: Copy folder ke htdocs jika belum ada
set "SRC=%~dp0"
set "DEST=C:\xampp\htdocs\admin-mbg\"

if not exist "%DEST%" (
    echo [*] Menyalin project ke htdocs XAMPP...
    xcopy /E /I /Y "%SRC%" "%DEST%"
    echo [+] Berhasil disalin ke %DEST%
) else (
    echo [*] Memperbarui file di htdocs XAMPP...
    xcopy /E /I /Y "%SRC%" "%DEST%"
    echo [+] File diperbarui.
)

echo.
echo ============================================
echo   Membuka: http://localhost/admin-mbg/
echo ============================================
echo.
start "" "http://localhost/admin-mbg/"
pause

