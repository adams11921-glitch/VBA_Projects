@echo off
rem Removes the Word Publisher Kit.
set "FILE=%APPDATA%\Microsoft\Word\STARTUP\WordPublisherKit.dotm"
tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>nul | find /I "WINWORD.EXE" >nul
if not errorlevel 1 (
  echo Please close Microsoft Word, then run Uninstall.cmd again.
  pause & exit /b 1
)
if exist "%FILE%" del "%FILE%"
echo The Bulletin and Publisher tabs have been removed from Word.
pause
