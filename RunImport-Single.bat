@echo off
chcp 65001 >nul
REM ============================================================
REM  Import a SINGLE layout by keyword (asks you what to import)
REM  Connection settings are in _settings.bat (shared, gitignored).
REM ============================================================
if not exist "%~dp0_settings.bat" (
    echo ERROR: _settings.bat not found.
    echo Copy _settings.bat.example to _settings.bat and edit it.
    pause
    exit /b 1
)
call "%~dp0_settings.bat"

REM AUTHOR: stored on NEW rows only (INSERT). UPDATE keeps original Author.
set AUTHOR=manager

REM ============================================================
REM  CONFIGDIR: folder to scan for mapping .xlsx files
REM  Can be relative (Config) or absolute (C:\path\...)
REM ============================================================
set CONFIGDIR=Config

setlocal enabledelayedexpansion
set "CFG=!CONFIGDIR!"
if not "!CFG:~1,1!"==":" set "CFG=%~dp0!CFG!"
set "RPT=!RPTROOT!"
if not "!RPT:~1,1!"==":" set "RPT=%~dp0!RPT!"

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
echo  Single Layout Import
echo  Server   : %SERVER%
echo  Database : %COMPANYDB%
echo  MapFile  : !MAPFILE!
echo  RptRoot  : !RPT!
echo ============================================

:KEYWORD_LOOP
echo.
set "FILTER="
set /p FILTER=Type keyword (e.g. Sale Order) -- empty Enter to quit:

if "%FILTER%"=="" goto DONE

echo.
echo Importing rows matching "%FILTER%" ...
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Scripts\Import_SQL_Direct.ps1" ^
    -Server "%SERVER%" ^
    -CompanyDB "%COMPANYDB%" ^
    -DBUser "%DBUSER%" ^
    -DBPassword "%DBPASSWORD%" ^
    -Author "%AUTHOR%" ^
    -MapFile "!CFG!\!MAPFILE!" ^
    -RptRoot "!RPT!" ^
    -FilterFileName "%FILTER%" ^
    -UseFileNameAsDocName ^
    -OnDuplicate Update

echo.
echo ============================================
echo  Done. Type next keyword, or empty Enter to quit.
echo ============================================
goto KEYWORD_LOOP

:DONE
endlocal
echo.
echo Bye.
