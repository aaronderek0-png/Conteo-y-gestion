@echo off
title Bot Mishiguene - Orquestador (RUN_BOT)
cd /d "%~dp0"
echo Orquestador Mishiguene: Telegram long polling + cola + escritura en bloque. Ctrl+C o crear estado\STOP para parar.
python orquestador_mishi.py --push-drive
echo.
echo Orquestador detenido (codigo %ERRORLEVEL%). Si fue 3: ya hay otro consumidor activo (ver estado_bot.py).
pause
