@echo off
chcp 65001 >nul
title Instalar - Gerador de Documentos Infobarra
setlocal

REM ============================================================
REM  CONFIGURE AQUI a pasta de rede onde o app foi publicado:
set "ORIGEM=\\SERVIDOR\apps\GeradorDocs"
REM ============================================================

set "DESTINO=%LOCALAPPDATA%\GeradorDocsInfobarra"

echo.
echo   Instalando o Gerador de Documentos Infobarra
echo   De:   %ORIGEM%
echo   Para: %DESTINO%
echo.

if not exist "%ORIGEM%\index.html" (
  echo   [ERRO] Nao encontrei o app em "%ORIGEM%".
  echo   Edite a linha ORIGEM deste arquivo com a pasta de rede correta.
  echo.
  pause & exit /b 1
)

REM Copia o app (sem itens de desenvolvimento)
robocopy "%ORIGEM%" "%DESTINO%" /MIR /NFL /NDL /NJH /NJS /NP /R:1 /W:1 ^
  /XD ".git" "_source" "supabase" ".claude" ^
  /XF ".gitignore" "CLAUDE.md" "README.md" >nul

REM Cria o atalho na area de trabalho (aponta para o launcher escondido)
powershell -NoProfile -ExecutionPolicy Bypass -Command "$s=New-Object -ComObject WScript.Shell; $l=$s.CreateShortcut([Environment]::GetFolderPath('Desktop')+'\Gerador de Documentos.lnk'); $l.TargetPath=$env:WINDIR+'\System32\wscript.exe'; $l.Arguments='\"%DESTINO%\Abrir Gerador.vbs\"'; $l.WorkingDirectory='%DESTINO%'; $l.IconLocation='%DESTINO%\app.ico'; $l.Description='Gerador de Documentos Infobarra'; $l.Save()"

echo.
echo   Instalado! Foi criado o atalho "Gerador de Documentos" na area de trabalho.
echo   E so dar dois cliques nele para abrir o app.
echo.
pause
