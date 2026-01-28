@echo off
setlocal EnableExtensions

rem Directory of this .cmd
set "SCRIPT_DIR=%~dp0"

rem Directory where CsvEditor.xlsm lives (relative to this script)
rem For the default case, it's the same as SCRIPT_DIR 
rem but it could be something like 
rem set "EDITOR_DIR=%SCRIPT_DIR%..\scripts\"
set "EDITOR_DIR=%SCRIPT_DIR%"

set "ARG=%~1"
set "BASE=%~n0"

rem ----------------------------
rem Determine CSV
rem ----------------------------

if not "%ARG%"=="" (
  set "CSVARG=%ARG%"
  goto :have_csv
)

set "CAND1=%SCRIPT_DIR%%BASE%"
if exist "%CAND1%" (
  set "CSVARG=%CAND1%"
) else (
  if /i "%BASE:~0,5%"=="edit-" if exist "%SCRIPT_DIR%%BASE:~5%" (
    set "CSVARG=%SCRIPT_DIR%%BASE:~5%"
  )
)

if not defined CSVARG (
  echo Could not infer CSV.
  exit /b 2
)

:have_csv

rem ----------------------------
rem Environment
rem ----------------------------

set "EXCEL_CSV_PATH=%CSVARG%"
set "EXCEL_CSV_CWD=%CD%"

rem ----------------------------
rem Launch editor
rem ----------------------------

set "WB=%EDITOR_DIR%CsvEditor.xlsm"

if not exist "%WB%" (
  echo Editor not found: "%WB%"
  exit /b 2
)

start "" "%WB%"
exit /b 0
