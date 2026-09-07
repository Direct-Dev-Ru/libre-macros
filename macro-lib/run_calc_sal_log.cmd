@echo off
REM Run LibreOffice Calc with SAL debug log to console (Linux equivalent):
REM   SAL_LOG=+INFO.all libreoffice --calc
REM
REM On Windows soffice.exe is a launcher: it returns to the shell immediately.
REM This script runs soffice.bin instead — the process keeps the console until Calc exits.
REM
REM   run_calc_sal_log.cmd
REM   run_calc_sal_log.cmd C:\path\to\file.ods

setlocal

if not defined SAL_LOG set "SAL_LOG=+INFO.all"

if defined SOFFICE (
    set "SOFFICE_EXE=%SOFFICE%"
) else if exist "%ProgramFiles%\LibreOffice\program\soffice.exe" (
    set "SOFFICE_EXE=%ProgramFiles%\LibreOffice\program\soffice.exe"
) else if exist "%ProgramFiles(x86)%\LibreOffice\program\soffice.exe" (
    set "SOFFICE_EXE=%ProgramFiles(x86)%\LibreOffice\program\soffice.exe"
) else (
    where soffice >nul 2>&1
    if errorlevel 1 (
        echo soffice.exe not found. Set SOFFICE=full\path\to\soffice.exe
        exit /b 1
    )
    for /f "delims=" %%I in ('where soffice') do (
        set "SOFFICE_EXE=%%I"
        goto :found_soffice
    )
)
:found_soffice

for %%I in ("%SOFFICE_EXE%") do set "SOFFICE_BIN=%%~dpIsoffice.bin"
if not exist "%SOFFICE_BIN%" (
    echo soffice.bin not found next to %SOFFICE_EXE%
    exit /b 1
)

echo SAL_LOG=%SAL_LOG%
if defined SAL_LOG_FILE (
    echo SAL_LOG_FILE=%SAL_LOG_FILE%
) else (
    echo SAL_LOG_FILE=(not set, output to console^)
)
echo Starting: "%SOFFICE_BIN%" --calc %*
echo Waiting for Calc to exit...
echo.

"%SOFFICE_BIN%" --calc %*

echo.
echo Calc exited.

endlocal
