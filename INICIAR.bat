@echo off
chcp 65001 >nul
title Gerador de Documentos - Infobarra
cd /d "%~dp0"
echo.
echo   ============================================
echo    GERADOR DE DOCUMENTOS - INFOBARRA
echo   ============================================
echo.
echo    Iniciando servidor local na porta 8777...
echo    NAO FECHE esta janela enquanto usar o app.
echo.
start "" powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0server.ps1" -Port 8777
timeout /t 2 >nul
start "" "http://localhost:8777/"
echo    Navegador aberto em http://localhost:8777/
echo.
echo    Para encerrar: feche esta janela e a janela azul do servidor.
echo.
pause >nul
