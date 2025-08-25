@echo off
set CHROME_EXE="C:\Program Files\Google\Chrome\Application\chrome.exe"
if not exist %CHROME_EXE% set CHROME_EXE="C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"

set DEBUG_PROFILE=%USERPROFILE%\ChromeDebugProfile

start "" %CHROME_EXE% --remote-debugging-port=9222 --user-data-dir="%DEBUG_PROFILE%"
