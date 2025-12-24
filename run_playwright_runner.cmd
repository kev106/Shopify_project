@ECHO OFF
SETLOCAL EnableDelayedExpansion
SET "PROJECT_ROOT=C:\Users\kev\My Drive\Shopify_project"
SET "PYTHON=%PROJECT_ROOT%\.venv\Scripts\python.exe"
IF NOT EXIST "%PYTHON%" SET "PYTHON=python"
REM Change to project directory so .env is found and relative paths work
CD /D "%PROJECT_ROOT%"
REM Ensure downloads\debug exists
IF NOT EXIST "%PROJECT_ROOT%\downloads\debug" MKDIR "%PROJECT_ROOT%\downloads\debug"
REM Run with full logging
"%PYTHON%" "%PROJECT_ROOT%\playwright_runner.py" >> "%PROJECT_ROOT%\downloads\debug\scheduled_run.log" 2>&1
ECHO Return code: !ERRORLEVEL! >> "%PROJECT_ROOT%\downloads\debug\scheduled_run.log"
