@echo off
title Bot Mishiguene - Orquestador (RUN_BOT) - ENTORNO PRUEBA
cd /d "%~dp0"
set MISHI_ENTORNO=PRUEBA
echo ============================================================
echo  ORQUESTADOR MISHIGUENE  -  ENTORNO: PRUEBA (copias NO CARGAR)
echo  Escribe SOLO en las copias de prueba; rechaza IDs de produccion.
echo  Parar: Ctrl+C en esta ventana, o crear el archivo estado\STOP
echo ============================================================
python estado_bot.py
python orquestador_mishi.py --push-drive
echo.
echo Orquestador detenido (codigo %ERRORLEVEL%). Codigo 3 = ya hay otro consumidor activo (usar python estado_bot.py).
pause
