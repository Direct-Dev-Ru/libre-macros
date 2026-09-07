@echo off
REM Run AlterOffice Calc with SAL debug log to console (Linux equivalent):
REM   SAL_LOG=+INFO.all aoffice --calc
REM
REM On Windows aoffice.exe may return to the shell immediately.
REM Prefer aoffice.bin next to aoffice.exe when present — keeps the console until Calc exits.
REM
REM   run_calc_sal_log_aoffice.cmd
REM   run_calc_sal_log_aoffice.cmd C:\path\to\file.ods

setlocal

if not defined SAL_LOG set "SAL_LOG=+INFO.all"

if defined AOFFICE (
    set "AOFFICE_EXE=%AOFFICE%"
) else if exist "%ProgramFiles%\AlterOffice\binaries\aoffice.exe" (
    set "AOFFICE_EXE=%ProgramFiles%\AlterOffice\binaries\aoffice.exe"
) else if exist "%ProgramFiles(x86)%\AlterOffice\binaries\aoffice.exe" (
    set "AOFFICE_EXE=%ProgramFiles(x86)%\AlterOffice\binaries\aoffice.exe"
) else (
    where aoffice >nul 2>&1
    if errorlevel 1 (
        echo aoffice.exe not found. Set AOFFICE=full\path\to\aoffice.exe
        exit /b 1
    )
    for /f "delims=" %%I in ('where aoffice') do (
        set "AOFFICE_EXE=%%I"
        goto :found_aoffice
    )
)
:found_aoffice

for %%I in ("%AOFFICE_EXE%") do set "AOFFICE_BIN=%%~dpIaoffice.bin"
if exist "%AOFFICE_BIN%" (
    set "AOFFICE_RUN=%AOFFICE_BIN%"
) else (
    set "AOFFICE_RUN=%AOFFICE_EXE%"
)

echo SAL_LOG=%SAL_LOG%
if defined SAL_LOG_FILE (
    echo SAL_LOG_FILE=%SAL_LOG_FILE%
) else (
    echo SAL_LOG_FILE=(not set, output to console^)
)
echo Starting: "%AOFFICE_RUN%" --calc %*
echo Waiting for Calc to exit...
echo.

"%AOFFICE_RUN%" --calc %*

echo.
echo Calc exited.

endlocal
