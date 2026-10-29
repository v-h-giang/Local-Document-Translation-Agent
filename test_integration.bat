@echo off
setlocal EnableExtensions
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    echo ERROR: Missing .venv. Run setup.bat first.
    exit /b 1
)
rem Default: five key workbooks. For all twenty: test_integration.bat --full
rem Optional: --backend excel (requires Microsoft Excel desktop)
if "%~1"=="" (
    ".venv\Scripts\python.exe" tests\run_integration.py --smoke
) else (
    ".venv\Scripts\python.exe" tests\run_integration.py %*
)
set "TEST_EXIT_CODE=%ERRORLEVEL%"
echo.
echo Integration-test exit code: %TEST_EXIT_CODE%
if not defined NO_PAUSE pause
exit /b %TEST_EXIT_CODE%
