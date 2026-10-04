@echo off
rem Installs the Word Publisher Kit (Bulletin tab) for the current Windows user.
setlocal
set "SRC=%~dp0WordPublisherKit.dotm"
set "DEST=%APPDATA%\Microsoft\Word\STARTUP"

if not exist "%SRC%" (
  echo WordPublisherKit.dotm must be in the same folder as Install.cmd.
  pause & exit /b 1
)
tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>nul | find /I "WINWORD.EXE" >nul
if not errorlevel 1 (
  echo Please close Microsoft Word, then run Install.cmd again.
  pause & exit /b 1
)
if not exist "%DEST%" mkdir "%DEST%"
copy /Y "%SRC%" "%DEST%\WordPublisherKit.dotm" >nul
rem Files downloaded from email or the internet are blocked from running macros until unblocked.
powershell -NoProfile -ExecutionPolicy Bypass -Command "Unblock-File -LiteralPath '%DEST%\WordPublisherKit.dotm'" >nul 2>&1
echo.
echo Installed. Open Word and look for the "Bulletin" tab next to "Home".
echo.
pause
