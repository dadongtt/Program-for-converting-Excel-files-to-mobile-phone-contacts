@echo off
setlocal
set "APP="
for %%F in ("%~dp0release\*.exe") do if not defined APP set "APP=%%~fF"

if not defined APP (
    echo Error: release EXE was not found.
    pause
    exit /b 1
)

start "" "%APP%"
