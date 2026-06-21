# Launcher do app (uso pelo atalho da área de trabalho).
# Sobe o server.ps1 ESCONDIDO (se ainda não estiver no ar) e abre o app
# numa janela de aplicativo (Chrome/Edge --app), sem abas nem barra de endereço.
$ErrorActionPreference = "SilentlyContinue"
$base = Split-Path -Parent $MyInvocation.MyCommand.Path
$port = 8777
$url  = "http://localhost:$port/"

# Servidor já está ouvindo na porta?
$inUse = $false
try {
  $c = New-Object System.Net.Sockets.TcpClient
  $c.Connect('127.0.0.1', $port); $inUse = $true; $c.Close()
} catch {}

if (-not $inUse) {
  Start-Process powershell -WindowStyle Hidden -ArgumentList @(
    '-NoProfile','-ExecutionPolicy','Bypass','-File', (Join-Path $base 'server.ps1'), '-Port', "$port"
  )
  Start-Sleep -Milliseconds 1500
}

# Navegador em modo aplicativo (procura Chrome, depois Edge)
$chrome = @(
  "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
  "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
  "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
) | Where-Object { Test-Path $_ } | Select-Object -First 1

$edge = @(
  "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
  "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe"
) | Where-Object { Test-Path $_ } | Select-Object -First 1

if     ($chrome) { Start-Process $chrome "--app=$url" }
elseif ($edge)   { Start-Process $edge   "--app=$url" }
else             { Start-Process $url }
