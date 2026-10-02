@echo off
rem ---------------------------------------------------------------------------
rem  Service_Manager - a save point before a design session.
rem
rem  Commits everything in the repository, so "git restore" has a version to go
rem  back to. Run it before you start moving cells around in the template.
rem
rem      tools\save-point.cmd "moving the dashboard buttons"
rem      tools\save-point.cmd                       (message: "save point")
rem ---------------------------------------------------------------------------
setlocal
cd /d "%~dp0.."

set MESSAGE=%~1
if "%MESSAGE%"=="" set MESSAGE=save point
rem cmd treats these as redirection or piping even inside quotes when the caller
rem nests the quoting, so a message like "09:00-19:00 -> 09..18" once created a
rem file named "09..18)". They are dropped from the message.
set MESSAGE=%MESSAGE:>=%
set MESSAGE=%MESSAGE:<=%
set MESSAGE=%MESSAGE:|=%
set MESSAGE=%MESSAGE:&=%

git rev-parse --is-inside-work-tree >nul 2>&1
if errorlevel 1 (
  echo This folder is not a git repository. Run: git init -b main
  pause
  exit /b 1
)

git add -A
git commit -m "%MESSAGE%"
if errorlevel 1 echo Nothing to commit - the save point is already there.
echo.
git --no-pager log --oneline -3
pause
