@echo off
rem =====================================================================
rem  Service_Manager - double-clickable wrapper around the build harness.
rem
rem  Usage (from anywhere):
rem    tools\run-tests.cmd            -> test mode, pauses at the end
rem    tools\run-tests.cmd build      -> production build, pauses
rem    set NOPAUSE=1 && tools\run-tests.cmd test
rem
rem  The real work is done by tools\build.vbs (see that file for details).
rem =====================================================================
setlocal

cd /d "%~dp0.."

set MODE=%1
if "%MODE%"=="" set MODE=test

cscript //nologo "tools\build.vbs" %MODE%
set RC=%ERRORLEVEL%

echo.
echo [run-tests] mode=%MODE% exit code=%RC%
if defined NOPAUSE goto :end
pause

:end
endlocal & exit /b %RC%
