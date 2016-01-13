@echo off
tasklist /FI "IMAGENAME eq outlook.exe" /FI "USERNAME eq %USERNAME%" 2>NUL | find /I /N "outlook.exe">NUL
if "%ERRORLEVEL%"=="1" "C:\Program Files\Microsoft Office\Office14\OUTLOOK.EXE"