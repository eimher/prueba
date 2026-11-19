@echo off
REM =====================================================================
REM ARCHIVO   : ejecutar_departamentos.cmd
REM PROPOSITO : Lanzador por doble clic de departamentos_colombia.ps1.
REM ALCANCE   : Ejecuta el .ps1 de la MISMA carpeta (ruta relativa %~dp0)
REM             con ExecutionPolicy Bypass solo para este proceso.
REM =====================================================================
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0departamentos_colombia.ps1"
REM %~dp0 = carpeta de este .cmd; evita depender del directorio actual
echo Codigo de salida: %ERRORLEVEL%
REM Muestra 0 (exito) o 1 (error) devuelto por el script
pause
REM Mantiene la ventana abierta para leer el resultado
