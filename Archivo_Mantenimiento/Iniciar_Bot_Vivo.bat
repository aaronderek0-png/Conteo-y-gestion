@echo off
title Bot Mishiguene en vivo
cd /d "%~dp0"
echo Bot Mishiguene escuchando el grupo "Facturas Mishiguene" cada 2 segundos (Ctrl+C para salir)...
python telegram_ingestor.py --watch --interval 2 --push-drive
pause
