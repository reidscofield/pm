@echo off
title Hydro-Wates Project Manager
cd /d "%~dp0"

where node >nul 2>nul
if errorlevel 1 (
  echo.
  echo  Node.js is required but was not found.
  echo  Install the LTS version from https://nodejs.org and try again.
  echo.
  pause
  exit /b 1
)

start "" /min cmd /c "timeout /t 2 >nul & start "" http://localhost:8743"
node server.js
pause
