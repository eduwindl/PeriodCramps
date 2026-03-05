@echo off
chcp 65001 > nul
echo ========================================================
echo   Iniciando subida automatica a GitHub...
echo ========================================================
echo.

:: Agregar todos los cambios
git add .

:: Hacer commit con la fecha y hora actual, ignorando si falla por no haber cambios
git commit -m "Actualizacion automatica: %date% %time%"

:: Subir a la rama main
echo.
echo Subiendo a la nube...
git push origin main

echo.
echo ========================================================
echo   Procedimiento finalizado. ¡Cambios en GitHub!
echo ========================================================
pause
