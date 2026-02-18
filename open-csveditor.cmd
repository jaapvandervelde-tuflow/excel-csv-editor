@echo off
setlocal EnableExtensions EnableDelayedExpansion

rem ==================================================================
rem Locations (customise these in copies of this script)
rem ==================================================================

rem Directory of this .cmd, no need to customise
set "SCRIPT_DIR=%~dp0"

rem Directory where CsvEditor.xlsm lives
rem Default: same as SCRIPT_DIR
rem Examples of custom value:
rem   set "EDITOR_DIR=%SCRIPT_DIR%..\scripts\"
rem   set "EDITOR_DIR=C:\Scripts\CsvEditor\"
set "EDITOR_DIR=%SCRIPT_DIR%"

rem Directory where CSV files live
rem Default: same as SCRIPT_DIR
rem Examples of custom value:
rem   set "DATA_DIR=%SCRIPT_DIR%..\csvs\"
rem   set "DATA_DIR=C:\Data\CsvFiles\"
set "DATA_DIR=%SCRIPT_DIR%"

rem Temp folder next to the editor workbook (created if missing)
rem Default: .temp\ folder inside EDITOR_DIR
set "EDITOR_TEMP_DIR=%EDITOR_DIR%.temp\\"

rem ==================================================================
rem Determine CSV path argument
rem ==================================================================

set "ARG=%~1"
set "BASE=%~n0"

if not "%ARG%"=="" (
  rem If argument is relative, make it relative to DATA_DIR
  set "CSVARG=%ARG%"
  goto :resolve_csv
)

rem No argument: infer from script name:
rem   <name>.csv.cmd  -> <name>.csv
rem   edit-<name>.csv.cmd -> <name>.csv
set "CAND1=%DATA_DIR%%BASE%"
if exist "%CAND1%" (
  set "CSVARG=%CAND1%"
) else (
  if /i "%BASE:~0,5%"=="edit-" if exist "%DATA_DIR%%BASE:~5%" (
    set "CSVARG=%DATA_DIR%%BASE:~5%"
  )
)

if not defined CSVARG (
  echo Could not infer CSV.
  echo Tried:
  echo   "%CAND1%"
  if /i "%BASE:~0,5%"=="edit-" echo   "%DATA_DIR%%BASE:~5%"
  exit /b 2
)

:resolve_csv
rem Iterating over single value for access to modifiers like ~nxF
for %%F in ("%CSVARG%") do set "CSVBASE=%%~nxF"
rem Make a filename-safe prefix (replace characters illegal in Windows filenames)
set "CSVSAFE=%CSVBASE%"
set "CSVSAFE=!CSVSAFE: =_!"


rem If CSVARG is not absolute/UNC, treat it as relative to DATA_DIR
call :is_abs "%CSVARG%"
if errorlevel 1 (
  set "CSVARG=%DATA_DIR%%CSVARG%"
)

if not exist "%CSVARG%" (
  echo CSV not found: "%CSVARG%"
  exit /b 2
)

rem ============================
rem Environment for the workbook
rem ============================

set "EXCEL_CSV_PATH=%CSVARG%"
set "EXCEL_CSV_CWD=%CD%"

rem Optional delimiter override:
rem set "EXCEL_CSV_DELIM=,"

rem ============================
rem Launch editor (copy-on-launch)
rem ============================

set "WB_SRC=%EDITOR_DIR%CsvEditor.xlsm"
if not exist "%WB_SRC%" (
  echo Editor not found: "%WB_SRC%"
  exit /b 2
)

rem Ensure temp directory exists
if not exist "%EDITOR_TEMP_DIR%" mkdir "%EDITOR_TEMP_DIR%" >nul 2>&1

rem ----------------------------
rem Ensure Excel trusted location for temp folder (idempotent)
rem ----------------------------

set "TRUST_PS1=%EDITOR_DIR%scripts\Ensure-ExcelTrustedLocation.ps1"
if exist "%TRUST_PS1%" (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%TRUST_PS1%" -TrustedPath "%EDITOR_TEMP_DIR%"
) else (
  echo Warning: Could not find trusted location script: "%TRUST_PS1%"
)

rem Cleanup old temp copies (silent on failures/locks)
call :cleanup_temp "%EDITOR_TEMP_DIR%"

rem Create a unique per-session copy name
set "STAMP=%RANDOM%_%RANDOM%_%TIME%"
set "STAMP=%STAMP::=%"
set "STAMP=%STAMP: =0%"
rem Marker token: __CsvEditor__ so cleanup can target these reliably
set "WB_COPY=%EDITOR_TEMP_DIR%!CSVSAFE!__CsvEditor__%STAMP%.xlsm"

copy /y "%WB_SRC%" "%WB_COPY%" >nul
if errorlevel 1 (
  echo Failed to copy editor workbook to: "%WB_COPY%"
  exit /b 2
)

call :find_excel_exe
if not defined EXCEL_EXE (
  echo Could not locate EXCEL.EXE
  exit /b 2
)

start "" "%EXCEL_EXE%" /x "%WB_COPY%"
exit /b 0

rem ============================
rem Helpers
rem ============================

:is_abs
rem returns errorlevel 0 if absolute or UNC; 1 otherwise
set "P=%~1"
if "%P:~1,1%"==":" exit /b 0
if "%P:~0,2%"=="\\" exit /b 0
exit /b 1

:cleanup_temp
set "TD=%~1"
rem Delete any prior copies; ignore errors (locked/open files)
del /q "%TD%*__CsvEditor__*.xlsm" >nul 2>&1
rem Also remove legacy naming pattern used previously
del /q "%TD%CsvEditor_*.xlsm" >nul 2>&1
exit /b 0


:find_excel_exe
set "EXCEL_EXE="
for %%P in (
  "%ProgramFiles%\Microsoft Office\root\Office16\EXCEL.EXE"
  "%ProgramFiles(x86)%\Microsoft Office\root\Office16\EXCEL.EXE"
  "%ProgramFiles%\Microsoft Office\Office16\EXCEL.EXE"
  "%ProgramFiles(x86)%\Microsoft Office\Office16\EXCEL.EXE"
) do (
  if not defined EXCEL_EXE if exist %%~P set "EXCEL_EXE=%%~P"
)
exit /b 0
