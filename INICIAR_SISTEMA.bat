@echo off
title Sistema Excel - REUNION 123
cd /d "%~dp0"
echo Iniciando Sistema Excel REUNION 123 en http://localhost:3070 ...
start "" "http://localhost:3070"
node server.js
pause
