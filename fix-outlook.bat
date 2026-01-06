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

REM Step 2: Clear saved credentials from Windows Credential Manager
REM Old saved passwords can cause repeated authentication failures and account lockouts
echo [2] Clearing saved credentials...
for /f "tokens=1,2 delims=: " %%a in ('cmdkey /list ^| findstr /i "Target:"') do (
    echo %%b | findstr /i "outlook office365 microsoftoffice exchange mail" >nul
    if not errorlevel 1 (
        echo Removing: %%b
        cmdkey /delete:%%b 2>nul
    )
)

echo.

REM Step 3: Force close Outlook to ensure clean restart
REM New credentials will be requested on next launch
echo [3] Closing Outlook...
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
