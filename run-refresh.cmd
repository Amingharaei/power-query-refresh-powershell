@echo off
REM ============================================================
REM  Runs the Excel refresh. Task Scheduler (or a double-click)
REM  points at THIS file. It self-locates via %~dp0, so the
REM  project folder can live anywhere.
REM ============================================================

REM %~dp0 = the folder this file lives in (the project folder).
cd /d "%~dp0"

title Excel Refresh Orchestrator (PowerShell)
echo.
echo  Excel Refresh Orchestrator (PowerShell)
echo  Starting up...
echo.

REM Prefer PowerShell 7 (pwsh) if it is installed, otherwise fall back to the
REM built-in Windows PowerShell 5.1 (powershell). -STA pins the single-threaded
REM apartment that Excel COM requires; -NoProfile keeps startup clean and fast.
where pwsh >nul 2>nul
if %ERRORLEVEL%==0 (
    pwsh -STA -NoProfile -ExecutionPolicy Bypass -File "%~dp0Refresh-Reports.ps1"
) else (
    powershell -STA -NoProfile -ExecutionPolicy Bypass -File "%~dp0Refresh-Reports.ps1"
)

REM To hide the console window during scheduled runs, add -WindowStyle Hidden, e.g.:
REM     pwsh -STA -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "%~dp0Refresh-Reports.ps1"

REM Propagate the script's exit code (0 = all ok, 1 = a failure) to Task Scheduler.
exit /b %ERRORLEVEL%
