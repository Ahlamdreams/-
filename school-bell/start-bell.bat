@echo off
REM ==== School Bell launcher (Windows) ====
REM Opens the app full screen with sound allowed automatically.
REM Data is stored in the "browser-data" folder next to this file (portable).
set "APPDIR=%~dp0"
set "URL=file:///%APPDIR:\=/%index.html"
set "PROFILE=%APPDIR%browser-data"
set "FLAGS=--user-data-dir="%PROFILE%" --autoplay-policy=no-user-gesture-required --start-fullscreen --no-first-run --disable-session-crashed-bubble --disable-features=Translate --app="%URL%""

set "BROWSER="
if exist "%ProgramFiles%\Google\Chrome\Application\chrome.exe" set "BROWSER=%ProgramFiles%\Google\Chrome\Application\chrome.exe"
if not defined BROWSER if exist "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe" set "BROWSER=%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"
if not defined BROWSER if exist "%LocalAppData%\Google\Chrome\Application\chrome.exe" set "BROWSER=%LocalAppData%\Google\Chrome\Application\chrome.exe"
if not defined BROWSER if exist "%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe" set "BROWSER=%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe"
if not defined BROWSER if exist "%ProgramFiles%\Microsoft\Edge\Application\msedge.exe" set "BROWSER=%ProgramFiles%\Microsoft\Edge\Application\msedge.exe"

if not defined BROWSER (
  echo Chrome or Edge was not found. Please install Google Chrome.
  pause
  exit /b 1
)
start "" "%BROWSER%" %FLAGS%
