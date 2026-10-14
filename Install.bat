@echo off
REM One-time setup: installs what the job bot and its dashboard need,
REM and puts a "Job Bot Dashboard" shortcut on your desktop.
cd /d "%~dp0"
echo.
echo  Installing Job Bot...
echo.
where python >nul 2>nul
if errorlevel 1 (
  echo  Python was not found.
  echo  Install Python 3.9 or newer from https://www.python.org/downloads/
  echo  and tick "Add python.exe to PATH" in the installer. Then run Install.bat again.
  echo.
  pause
  exit /b 1
)
python -m pip install --upgrade pip -q
python -m pip install -r requirements.txt
if errorlevel 1 (
  echo.
  echo  Installing packages failed. Check your internet connection and run Install.bat again.
  pause
  exit /b 1
)
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$s=(New-Object -ComObject WScript.Shell).CreateShortcut([Environment]::GetFolderPath('Desktop')+'\Job Bot Dashboard.lnk');" ^
  "$s.TargetPath='%~dp0Start Dashboard.bat'; $s.WorkingDirectory='%~dp0'; $s.IconLocation='shell32.dll,165'; $s.Save()"
echo.
echo  Done! A "Job Bot Dashboard" shortcut is on your desktop.
echo  Starting the dashboard now...
echo.
start "" "%~dp0Start Dashboard.bat"
timeout /t 3 >nul
