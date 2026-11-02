@echo off
setlocal
cd /d "%~dp0"

if defined PYTHON_EXE goto custom
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" build.py
    goto result
)
py -3 --version >nul 2>nul
if not errorlevel 1 (
    py -3 build.py
    goto result
)
python build.py
goto result

:custom
"%PYTHON_EXE%" build.py

:result
if errorlevel 1 (
    echo Build failed. Check Python and install requirements.txt first.
    pause
    exit /b 1
)
echo Build complete. Open the release folder.
pause
