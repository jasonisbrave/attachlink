' AttachLink watchdog hidden launcher (called by scheduled task; no console flash)
' NOTE: keep this file ASCII-only. WSH reads ANSI for .vbs without BOM; non-ASCII
' comments can mojibake and swallow the following newline (breaks line 2 silently).
Set sh = CreateObject("WScript.Shell")
sh.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File ""C:\Program Files\AttachLink\watchdog.ps1"" -XlamPath ""C:\Program Files\AttachLink\AttachLink.xlam""", 0, False
