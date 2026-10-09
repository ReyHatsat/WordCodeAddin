@echo off
setlocal
cd /d "%~dp0"
where npm.cmd >nul 2>&1
if errorlevel 1 (
  echo Please install Node.js 22.15 or later, then reopen this file.
  echo Instale Node.js 22.15 o posterior y vuelva a abrir este archivo.
  pause
  exit /b 1
)
echo Installing WordCodeAddin / Instalando WordCodeAddin...
call npm.cmd ci
if errorlevel 1 goto failed
echo Sign in with the same account as Word.
echo Inicie sesion con la misma cuenta que utiliza en Word.
call npm.cmd run signin:browser
if errorlevel 1 goto failed
call npm.cmd run start:desktop:word:browser
if errorlevel 1 goto failed
echo Keep the local server running. Use Start.cmd next time.
echo Mantenga el servidor local funcionando. Use Start.cmd la proxima vez.
pause
exit /b 0
:failed
echo Installation failed. Please review the error above and the README.
echo La instalacion fallo. Revise el error anterior y el README.
pause
exit /b 1
