@echo off
setlocal
cd /d "%~dp0"

rem Edit these two values to change the double-click defaults.
set "DEFAULT_SOURCE_LANG=vi"
set "DEFAULT_TARGET_LANG=ja"

echo ============================================================
echo Local TranslateGemma Agent (default: vi to ja)
echo ============================================================
echo.

if not exist ".venv\Scripts\python.exe" (
    echo ERROR: Python virtual environment was not found.
    echo.
    echo Run setup.bat first.
    echo.
    pause
    exit /b 1
)

rem Defaults are overridable: run.bat --from ja --to vi
rem Arguments in %* come last, so argparse uses them over these defaults.
".venv\Scripts\python.exe" run_translate_all.py --from "%DEFAULT_SOURCE_LANG%" --to "%DEFAULT_TARGET_LANG%" %*
set "EXIT_CODE=%ERRORLEVEL%"

echo.
echo ============================================================
echo Finished. Exit code: %EXIT_CODE%
echo ============================================================
if not defined NO_PAUSE pause
exit /b %EXIT_CODE%
