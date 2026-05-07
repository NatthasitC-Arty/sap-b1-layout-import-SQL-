@echo off
chcp 65001 >nul
REM ============================================================
REM  Rollback Crystal Layouts imported from RPT_Import_Map.xlsx
REM  Deletes rows in RDOC matching (DocName + TypeCode), Author ignored
REM  Connection settings are in _settings.bat (shared, gitignored).
REM ============================================================
if not exist "%~dp0_settings.bat" (
    echo ERROR: _settings.bat not found.
    echo Copy _settings.bat.example to _settings.bat and edit it.
    pause
    exit /b 1
)
call "%~dp0_settings.bat"

REM AUTHOR is no longer used for matching; kept only for log display.
set AUTHOR=SDA

REM ============================================================
REM  MODE:
REM    -DryRun        = preview only (no delete)
REM    (leave empty)  = delete for real (will ask confirmation)
REM    -Force         = delete without asking
REM ============================================================
set MODE=

REM ============================================================
REM  CONFIGDIR: folder to scan for mapping .xlsx files
REM  Can be relative (Config) or absolute (C:\path\to\folder)
REM ============================================================
set CONFIGDIR=Config

setlocal enabledelayedexpansion
set "CFG=!CONFIGDIR!"
if not "!CFG:~1,1!"==":" set "CFG=%~dp0!CFG!"

echo ============================================
echo  Select mapping Excel file from:
echo  !CFG!
echo ============================================
set IDX=0
for %%F in ("!CFG!\*.xlsx") do (
    set /a IDX+=1
    set "FILE_!IDX!=%%~nxF"
    echo   !IDX!. %%~nxF
)
if %IDX%==0 (
    echo No .xlsx files found in !CFG!
    pause
    exit /b
)
echo.
set /p PICK=Enter number (1-%IDX%):
if "%PICK%"=="" (
    echo No selection. Exiting.
    pause
    exit /b
)
call set "MAPFILE=%%FILE_%PICK%%%"
if "%MAPFILE%"=="" (
    echo Invalid selection.
    pause
    exit /b
)
echo.
echo ============================================
echo  Rollback Layouts
echo  Server   : %SERVER%
echo  Database : %COMPANYDB%
echo  Author   : %AUTHOR%
echo  MapFile  : %MAPFILE%
echo  Mode     : %MODE%
echo ============================================
echo.
pause

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Scripts\Rollback-FromExcel.ps1" ^
    -Server "%SERVER%" ^
    -CompanyDB "%COMPANYDB%" ^
    -DBUser "%DBUSER%" ^
    -DBPassword "%DBPASSWORD%" ^
    -Author "%AUTHOR%" ^
    -MapFile "!CFG!\!MAPFILE!" ^
    -UseFileNameAsDocName ^
    %MODE%
endlocal

echo.
pause
