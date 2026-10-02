@echo off
rem ---------------------------------------------------------------------------
rem  Service_Manager - puts the template back to the last committed version.
rem
rem  The file that is in place now is copied away first, so nothing is lost -
rem  not even the mistake:
rem
rem      backups\template\broken-<date>_<time>.xlsx
rem
rem  Everything else (sources, tools) is restored with:  git restore .
rem ---------------------------------------------------------------------------
setlocal
cd /d "%~dp0.."

git rev-parse --is-inside-work-tree >nul 2>&1
if errorlevel 1 (
  echo This folder is not a git repository.
  pause
  exit /b 1
)

for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyy-MM-dd_HH-mm-ss"') do set STAMP=%%i

if not exist backups\template mkdir backups\template
copy /y Service_Manager.xlsx "backups\template\broken-%STAMP%.xlsx" >nul
if errorlevel 1 (
  echo Cannot copy the template - is it open in Excel? Close it and try again.
  pause
  exit /b 1
)

git checkout HEAD -- Service_Manager.xlsx
if errorlevel 1 (
  echo git could not restore the template.
  pause
  exit /b 1
)

echo The template is back at the last commit:
git --no-pager log --oneline -1
echo The file you had is in backups\template\broken-%STAMP%.xlsx
echo.
echo Next: cscript //nologo tools\build.vbs inspect
pause
