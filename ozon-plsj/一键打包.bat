@echo off
setlocal

cd /d "%~dp0"
rem Remove trailing backslash so quoted path does not escape the closing quote
set "RELEASE_DIR=%~dp0"
if "%RELEASE_DIR:~-1%"=="\" set "RELEASE_DIR=%RELEASE_DIR:~0,-1%"
set "BUILD_SCRIPT=%~dp0..\tool\exe_build\build_release.ps1"

echo ========================================
echo   OzonTool Build
echo   Output: %RELEASE_DIR%
echo ========================================
echo.

if not exist "%BUILD_SCRIPT%" (
    echo [ERROR] Missing build script:
    echo   %BUILD_SCRIPT%
    pause
    exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%BUILD_SCRIPT%" -ReleaseDir "%RELEASE_DIR%"
if errorlevel 1 (
    echo.
    echo [FAILED] See errors above.
    pause
    exit /b 1
)

echo.
echo [OK] Find OzonTool_*.exe in this folder.
echo.
pause
exit /b 0
