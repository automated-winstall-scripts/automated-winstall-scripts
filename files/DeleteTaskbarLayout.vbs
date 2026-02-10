Set WshShell = CreateObject("WScript.Shell")

WshShell.Run "cmd.exe /c del ""%LOCALAPPDATA%\Microsoft\Windows\Shell\LayoutModification.xml"" /f >nul 2>&1", 0, False

WshShell.Run "cmd.exe /c del ""%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\DeleteTaskbarLayout.vbs"" /f >nul 2>&1", 0, False

Set WshShell = Nothing