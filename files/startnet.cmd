@echo off
setlocal enabledelayedexpansion
wpeinit

:: Find USB drive letter
for %%a in (C D E F G H I J K L M N O P Q R S T U V W X Y Z) do (
    if exist %%a:\sources\install.wim set USBDRIVE=%%a
)

:: Check if drive letter was found
if not defined USBDRIVE (
    echo No USB drive is plugged in.
    timeout 5 > nul
    exit
)

:: Run Check-NetworkConnectivity.ps1 script
PowerShell -NoProfile -ExecutionPolicy Bypass -NoExit -File "!USBDRIVE!:\CompanyNameImaging\Test-NetworkConnectivity.ps1"