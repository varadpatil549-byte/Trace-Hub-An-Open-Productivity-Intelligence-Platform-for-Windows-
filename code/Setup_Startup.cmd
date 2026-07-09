@echo off

set startup=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup

copy "%~dp0TraceHub.vbs" "%startup%\TraceHub.vbs" /Y

echo.
echo Trace HUB installed successfully!
echo Auto startup enabled.
pause