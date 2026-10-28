@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ==============================================
echo   UNIALIADOS - CARGA MASIVA OPTIMIZADA
 echo ==============================================
echo.
echo Cierre Unialiados_Operacion_Unificada.xlsm y pause n8n antes de importar.
echo Este proceso crea un respaldo y permite correos repetidos.
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Importar_Unialiados_Optimizado.ps1" -WorkbookPath "%~dp0Unialiados_Operacion_Unificada.xlsm"
set "RESULTADO=%ERRORLEVEL%"
echo.
if not "%RESULTADO%"=="0" echo Hubo un error. Revise el mensaje anterior; el archivo de respaldo se conserva.
pause
exit /b %RESULTADO%
