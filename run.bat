@echo off
cd /d "%~dp0"

if not exist "app.py" (
  echo app.py not found. Please EXTRACT the whole zip file first ^(right-click - Extract All^), then double-click run.bat inside the extracted folder.
  pause
  exit /b
)
where python >nul 2>nul
if errorlevel 1 (
  echo Python was not found. Install Python 3.10 - 3.13 ^(64-bit^) from python.org and tick "Add Python to PATH".
  pause
  exit /b
)

if not exist ".venv\Scripts\python.exe" (
  echo First-time setup, please wait...
  python -m venv .venv
  if errorlevel 1 goto :err
)

rem --- install packages: use the offline "wheels" folder first (works on networks that block PyPI), else download ---
if exist "wheels\" (
  ".venv\Scripts\python.exe" -m pip install -q --no-index --find-links=wheels -r requirements.txt
  if not errorlevel 1 goto :installed
  echo Offline install failed, trying to download from the internet...
)
".venv\Scripts\python.exe" -m pip install -q -r requirements.txt
if errorlevel 1 goto :err
:installed

rem --- find a free port (an old copy may still be running on 5000) ---
set PORT=5000
:findport
netstat -ano | findstr /R /C:":%PORT% .*LISTENING" >nul
if errorlevel 1 goto :portok
echo Port %PORT% is already in use ^(maybe an old copy is still running^), trying the next one...
set /a PORT+=1
if %PORT% GTR 5010 goto :noport
goto :findport

:portok
start "" cmd /c "timeout /t 3 >nul & start http://localhost:%PORT%"
echo.
echo Running at http://localhost:%PORT%   (close this window or press Ctrl+C to stop)
echo The page header should say "Version 2.0 - 6 steps". If not, an old copy is open in your browser.
echo IMPORTANT: this program has no login. Use it only on this computer, and close this window when finished.
echo.
".venv\Scripts\python.exe" app.py
pause
exit /b

:noport
echo No free port found between 5000 and 5010. Close old black windows and try again.
pause
exit /b

:err
echo.
echo Something went wrong. Check that Python 3.10-3.13 (64-bit) is installed. If the "wheels" folder is missing, you must be online for the first run.
pause
