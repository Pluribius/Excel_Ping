@echo off
echo Instalando módulos de Python...

REM Comprobar si pip está disponible
where pip >nul 2>&1
if errorlevel 1 (
    echo Error: pip no se encuentra. Por favor, asegúrese de que Python esté instalado y agregado a la variable de entorno PATH.
    pause
    exit /b 1
)

REM Instalar openpyxl
echo Instalando openpyxl...
pip install openpyxl
if errorlevel 1 (
    echo Error: Falló la instalación de openpyxl. Por favor, revise su conexión a internet y la instalación de Python.
    pause
    exit /b 1
)
echo openpyxl instalado con éxito.

REM Instalar subprocess (Nota: subprocess es un módulo incorporado de Python, no es necesario instalarlo)


echo Todos los módulos de Python requeridos (excepto los incorporados) han sido instalados.
echo Instalación completada.
pause
exit /b 0