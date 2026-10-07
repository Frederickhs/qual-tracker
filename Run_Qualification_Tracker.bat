@echo off
setlocal
title PM Qualification Activity Tracker
color 0A

cd /d "%~dp0"

echo ==============================================================
echo   PM Qualification Activity Tracker
echo ==============================================================
echo.

if not exist ".venv\Scripts\python.exe" (
    echo ERROR: The project Python environment was not found.
    echo Expected: .venv\Scripts\python.exe
    echo.
    echo Run the project setup instructions before using this runner.
    echo.
    pause
    exit /b 1
)

echo STEP 1 OF 2: Archiving previous output reports...
echo.
".venv\Scripts\python.exe" "Scripts\archive_outputs.py"
if errorlevel 1 (
    echo.
    echo RUN STOPPED: Previous reports could not be archived.
    echo Input reports were not changed.
    echo.
    pause
    exit /b 1
)

echo.
echo STEP 2 OF 2: Generating current venue reports...
echo.
".venv\Scripts\python.exe" "Scripts\report_generator.py"
if errorlevel 1 (
    echo.
    echo RUN FAILED: New reports were not completed.
    echo Previous reports remain safe in Archive.
    echo Input reports were not changed.
    echo.
    pause
    exit /b 1
)

echo.
echo ==============================================================
echo   SUCCESS: Current reports were generated.
echo ==============================================================
echo.
echo Opening the Output folder...
start "" "%~dp0Output"

echo.
echo Press any key to close this window.
pause >nul
endlocal
