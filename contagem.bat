@echo off
setlocal
cd /d "%~dp0"
title Contagem Auvo

echo.
echo  CONTAGEM AUTOMATICA - AUVO
echo  --------------------------
echo  Formato da data: AAAA-MM-DD  (ex: 2026-09-15)
echo.

for /f %%d in ('powershell -NoProfile -Command "(Get-Date).AddDays(-1).ToString('yyyy-MM-dd')"') do set "ONTEM=%%d"

set "INICIO="
set /p INICIO="  Data inicial [Enter = %ONTEM%]: "
if "%INICIO%"=="" set "INICIO=%ONTEM%"

set "FIM="
set /p FIM="  Data final   [Enter = %INICIO%]: "
if "%FIM%"=="" set "FIM=%INICIO%"

echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0auvo-contagem.ps1" -StartDate %INICIO% -EndDate %FIM%

echo.
pause
