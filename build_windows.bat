@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if errorlevel 1 goto no_launcher

py -3.14 -c "import sys; assert sys.version_info[:2] == (3, 14)" >nul 2>nul
if errorlevel 1 goto no_python314

if exist ".venv314\Scripts\python.exe" goto venv_ready
py -3.14 -m venv .venv314
if errorlevel 1 goto failed

:venv_ready
.venv314\Scripts\python.exe -c "import sys, tkinter; assert sys.version_info[:2] == (3, 14)"
if errorlevel 1 goto no_tkinter

.venv314\Scripts\python.exe -m pip install --upgrade pip
if errorlevel 1 goto failed
.venv314\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 goto failed
.venv314\Scripts\python.exe -m unittest -v
if errorlevel 1 goto failed

.venv314\Scripts\python.exe -m PyInstaller --noconfirm --clean --onefile --windowed --name MonthlyLevelStats app.py
if errorlevel 1 goto failed

echo.
echo Build completed: dist\MonthlyLevelStats.exe
pause
exit /b 0

:no_launcher
echo [ERROR] Python Launcher was not found. Install Python 3.14 first.
goto failed

:no_python314
echo [ERROR] Python 3.14 was not found. Install Python 3.14 first.
goto failed

:no_tkinter
echo [ERROR] Python 3.14 does not include tkinter. Reinstall Python with Tcl/Tk.
goto failed

:failed
echo.
echo [ERROR] Build failed. Review the messages above.
pause
exit /b 1
