@echo off
title Moskit Dashboard - Reiniciando
echo Parando aplicacao...
taskkill /F /IM python.exe /FI "WINDOWTITLE eq Moskit Dashboard" >nul 2>&1
timeout /t 2 /nobreak >nul
echo Iniciando novamente...
python app.py
pause
