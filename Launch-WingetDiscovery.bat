@echo off
REM Winget Package Discovery Tool Launcher
REM This batch file launches the Winget Discovery Tool with proper PowerShell settings

echo.
echo ================================================
echo   Winget Package Discovery Tool
echo ================================================
echo.
echo Starting the tool...
echo.

REM Check if PowerShell is available
where powershell >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: PowerShell not found!
    echo Please ensure PowerShell is installed on your system.
    pause
    exit /b 1
)

REM Launch PowerShell with the main script
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0WingetDiscovery.ps1"

REM Check exit code
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo An error occurred while running the tool.
    echo Please check the logs folder for details.
    pause
)

exit /b %ERRORLEVEL%
