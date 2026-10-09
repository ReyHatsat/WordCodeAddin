@echo off
setlocal
cd /d "%~dp0"
if not exist node_modules (
  echo Please run Install.cmd first. / Ejecute Install.cmd primero.
  pause
  exit /b 1
)
echo Leave this window open, then open Word.
echo Deje esta ventana abierta y abra Word.
call npm.cmd run dev-server
if errorlevel 1 pause
