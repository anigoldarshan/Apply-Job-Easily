@echo off
REM Opens a normal Chrome window the job bot can connect to.
REM Its logins are saved in the "chrome_profile" folder next to this file.

set "CHROME="
if exist "%ProgramFiles%\Google\Chrome\Application\chrome.exe" set "CHROME=%ProgramFiles%\Google\Chrome\Application\chrome.exe"
if exist "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe" set "CHROME=%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"
if exist "%LocalAppData%\Google\Chrome\Application\chrome.exe" set "CHROME=%LocalAppData%\Google\Chrome\Application\chrome.exe"
if not defined CHROME goto nochrome

start "" "%CHROME%" --remote-debugging-port=9222 --user-data-dir="%~dp0chrome_profile" --no-first-run https://accounts.google.com https://www.naukri.com/nlogin/login https://www.linkedin.com/login https://www.foundit.in/rio/login https://www.shine.com/myshine/login/ https://cutshort.io/login https://secure.indeed.com/auth
echo Chrome opened. Sign in to Google first, then to each job site with "Continue with Google".
echo Keep this Chrome window open, then run:  python job_bot.py
goto end

:nochrome
echo Google Chrome was not found. Install it from https://www.google.com/chrome
:end
pause
