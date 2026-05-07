@echo off
chcp 932 > nul
cd /d "%~dp0"

echo ============================================
echo  Pool App Deploy Tool
echo ============================================
echo.
echo Deploying from: %CD%
echo.

call firebase deploy --only hosting
if errorlevel 1 (
    echo.
    echo [ERROR] Deploy failed.
    echo firebase login or firebase use --add Ç™ïKóvÇ©Ç‡ÇµÇÍÇ‹ÇπÇÒÅB
    pause
    exit /b 1
)

echo.
echo ============================================
echo  Deploy Success!
echo ============================================
echo.
pause
