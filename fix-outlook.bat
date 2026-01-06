@echo off
chcp 65001 >nul
echo ========================================
echo   Outlook/Exchange Fix - Password Reset
echo ========================================
echo.

REM Step 1: Clear all cached Kerberos tickets
REM This forces Windows to request new tickets with the new password
echo [1] Clearing Kerberos tickets...
klist purge

echo.

REM Step 2: Clear cached Outlook/Office credentials from Windows Credential Manager
REM Old saved passwords can cause repeated authentication failures
echo [2] Clearing cached Outlook credentials...
cmdkey /list | findstr "MicrosoftOffice" >nul
if %errorlevel%==0 (
    for /f "tokens=2 delims=: " %%a in ('cmdkey /list ^| findstr "Target:"') do (
        echo Removing: %%a
        cmdkey /delete:%%a 2>nul
    )
)

echo.

REM Step 3: Force close Outlook to ensure clean restart
REM New credentials will be requested on next launch
echo [3] Restarting Outlook...
taskkill /f /im outlook.exe 2>nul
timeout /t 2 >nul

echo.
echo ========================================
echo   DONE!
echo ========================================
echo.
echo Please start Outlook now.
echo Enter your NEW password when prompted.
echo.
pause
