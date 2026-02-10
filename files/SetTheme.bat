@ECHO OFF
SETLOCAL ENABLEDELAYEDEXPANSION

if not exist "C:\CompanyName\USMT\Logs\loadstate.txt" (
    start "" "C:\Windows\CompanyName\CompanyName.deskthemepack"
    mkdir "%LOCALAPPDATA%\Microsoft\Windows\Shell" >nul 2>&1
    robocopy "C:\Windows\CompanyName" "%LOCALAPPDATA%\Microsoft\Windows\Shell" "LayoutModification.xml" >nul 2>&1
    reg.exe delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband" /f >nul 2>&1
    timeout /t 5 >nul 2>&1
    taskkill /f /im SystemSettings.exe >nul 2>&1
) else (
    del "C:\CompanyName\USMT\Logs\loadstate.txt" /f >nul 2>&1
)

ENDLOCAL