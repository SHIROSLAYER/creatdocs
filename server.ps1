param([int]$Port = 8777)
# Servidor HTTP estático via HttpListener (localhost não requer admin).
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path

$mime = @{
  ".html"="text/html; charset=utf-8"; ".css"="text/css; charset=utf-8";
  ".js"="application/javascript; charset=utf-8"; ".json"="application/json; charset=utf-8";
  ".docx"="application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  ".png"="image/png"; ".jpg"="image/jpeg"; ".svg"="image/svg+xml"; ".ico"="image/x-icon";
  ".woff2"="font/woff2"; ".map"="application/json";
}

# Converte bytes de um .docx em PDF usando o Word (alta fidelidade / vetorial)
function Convert-DocxToPdf([byte[]]$docxBytes){
  $tmp = Join-Path $root "_source\_tmp"
  if (-not (Test-Path $tmp)){ New-Item -ItemType Directory -Force $tmp | Out-Null }
  $id = [System.Guid]::NewGuid().ToString("N")
  $docxPath = Join-Path $tmp ($id + ".docx")
  $pdfPath  = Join-Path $tmp ($id + ".pdf")
  [System.IO.File]::WriteAllBytes($docxPath, $docxBytes)
  $word = $null; $doc = $null
  try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Open($docxPath, $false, $true)   # ConfirmConversions=false, ReadOnly=true
    $doc.ExportAsFixedFormat($pdfPath, 17)                  # 17 = wdExportFormatPDF
    $doc.Close($false); $doc = $null
    $word.Quit(); $word = $null
    return [System.IO.File]::ReadAllBytes($pdfPath)
  } finally {
    if ($doc)  { try { $doc.Close($false) } catch {} }
    if ($word) { try { $word.Quit() } catch {} }
    try { Remove-Item $docxPath, $pdfPath -Force -ErrorAction SilentlyContinue } catch {}
  }
}

# Lê um .mdb (Access) e devolve os registros da tabela "Office Address List" como JSON.
# Usa o PowerShell 32-bit + Access COM (via _source\_export_access.ps1).
function Convert-MdbToJson([byte[]]$bytes){
  $tmp = Join-Path $root "_source\_tmp"
  if (-not (Test-Path $tmp)){ New-Item -ItemType Directory -Force $tmp | Out-Null }
  $id = [System.Guid]::NewGuid().ToString("N")
  $mdb = Join-Path $tmp ($id + ".mdb")
  $json = Join-Path $tmp ($id + ".json")
  [System.IO.File]::WriteAllBytes($mdb, $bytes)
  try {
    $ps32 = Join-Path $env:WINDIR "SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
    $script = Join-Path $root "_source\_export_access.ps1"
    & $ps32 -NoProfile -ExecutionPolicy Bypass -File $script -mdb $mdb -out $json | Out-Null
    if (-not (Test-Path $json)){ throw "Não foi possível ler o .mdb (Access)." }
    return [System.IO.File]::ReadAllText($json, [System.Text.Encoding]::UTF8)
  } finally {
    try { Remove-Item $mdb, $json -Force -ErrorAction SilentlyContinue } catch {}
  }
}

# Cria um .mdb (Access Jet 4) com a tabela "Office Address List" a partir de JSON {records:[...]}.
function Convert-JsonToMdb([string]$jsonText){
  $obj = $jsonText | ConvertFrom-Json
  $records = @($obj.records)
  $out = Join-Path $env:TEMP ([System.Guid]::NewGuid().ToString("N") + ".mdb")
  if (Test-Path $out){ Remove-Item $out -Force }
  $cols = New-Object System.Collections.Specialized.OrderedDictionary
  foreach ($r in $records){
    foreach ($p in $r.PSObject.Properties){
      if ($p.Name -notlike "_*" -and -not $cols.Contains($p.Name)){ $cols.Add($p.Name, $true) }
    }
  }
  $colNames = @($cols.Keys)
  if ($colNames.Count -eq 0){ $colNames = @("CONTRATO") }
  $acc = $null; $db = $null
  try {
    $acc = New-Object -ComObject Access.Application
    $db = $acc.DBEngine.CreateDatabase($out, ";LANGID=0x0416;CP=1252;COUNTRY=0", 64)
    $td = $db.CreateTableDef("Office Address List")
    foreach ($c in $colNames){ $fld = $td.CreateField([string]$c, 12); $td.Fields.Append($fld) }  # 12 = dbMemo
    $db.TableDefs.Append($td)
    $rs = $db.OpenRecordset("Office Address List")
    foreach ($r in $records){
      $rs.AddNew()
      foreach ($c in $colNames){
        $prop = $r.PSObject.Properties[$c]
        if ($prop -and $null -ne $prop.Value -and "$($prop.Value)" -ne ""){
          $rs.Fields.Item([string]$c).Value = [string]$prop.Value
        }
      }
      $rs.Update()
    }
    $rs.Close(); $db.Close(); $db = $null
    $acc.Quit(); $acc = $null
    return [System.IO.File]::ReadAllBytes($out)
  } finally {
    if ($db)  { try { $db.Close() } catch {} }
    if ($acc) { try { $acc.Quit() } catch {} }
    try { Remove-Item $out -Force -ErrorAction SilentlyContinue } catch {}
  }
}

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$Port/")
$listener.Start()
Write-Host "Static server em http://localhost:$Port/  (root: $root)"

while ($listener.IsListening){
  try {
    $ctx = $listener.GetContext()
    $req = $ctx.Request
    $res = $ctx.Response
    try {
      $path = [System.Uri]::UnescapeDataString($req.Url.AbsolutePath)
      if ($path -eq "/" -or $path -eq ""){ $path = "/index.html" }
      $rel  = $path.TrimStart('/').Replace('/', [System.IO.Path]::DirectorySeparatorChar)
      $full = [System.IO.Path]::GetFullPath((Join-Path $root $rel))
      $rootFull = [System.IO.Path]::GetFullPath($root)

      $res.Headers["Cache-Control"] = "no-store"
      $res.Headers["Access-Control-Allow-Origin"] = "*"

      # Armazenamento compartilhado (relatorios/operadores) — JSON em data/store/<key>.json
      if ($path -eq "/__store"){
        $key = [string]$req.QueryString["key"]
        if ($key -notmatch '^[a-z0-9_]+$'){
          $b = [System.Text.Encoding]::UTF8.GetBytes("bad key"); $res.StatusCode = 400
          $res.ContentLength64 = $b.Length; $res.OutputStream.Write($b,0,$b.Length); $res.OutputStream.Close(); continue
        }
        $storeDir = Join-Path $root "data\store"
        if (-not (Test-Path $storeDir)){ New-Item -ItemType Directory -Force $storeDir | Out-Null }
        $file = Join-Path $storeDir ($key + ".json")
        if ($req.HttpMethod -eq "POST"){
          $sr = New-Object System.IO.StreamReader($req.InputStream, [System.Text.Encoding]::UTF8)
          $body = $sr.ReadToEnd(); $sr.Close()
          [System.IO.File]::WriteAllText($file, $body, (New-Object System.Text.UTF8Encoding($false)))
          $b = [System.Text.Encoding]::UTF8.GetBytes("ok")
        } else {
          $txt = if (Test-Path $file){ [System.IO.File]::ReadAllText($file, [System.Text.Encoding]::UTF8) } else { "null" }
          $b = [System.Text.Encoding]::UTF8.GetBytes($txt)
        }
        $res.StatusCode = 200; $res.ContentType = "application/json; charset=utf-8"
        $res.ContentLength64 = $b.Length; $res.OutputStream.Write($b,0,$b.Length); $res.OutputStream.Close(); continue
      }

      # Upload de PDF assinado -> data/assinados/<name>
      if ($req.HttpMethod -eq "POST" -and $path -eq "/__upload"){
        $name = [string]$req.QueryString["name"]
        $name = ($name -replace '[^A-Za-z0-9._-]', '_')
        if (-not $name){ $name = "anexo.pdf" }
        $dir = Join-Path $root "data\assinados"
        if (-not (Test-Path $dir)){ New-Item -ItemType Directory -Force $dir | Out-Null }
        $ms = New-Object System.IO.MemoryStream; $req.InputStream.CopyTo($ms)
        [System.IO.File]::WriteAllBytes((Join-Path $dir $name), $ms.ToArray())
        $b = [System.Text.Encoding]::UTF8.GetBytes('{"path":"data/assinados/' + $name + '"}')
        $res.StatusCode = 200; $res.ContentType = "application/json; charset=utf-8"
        $res.ContentLength64 = $b.Length; $res.OutputStream.Write($b,0,$b.Length); $res.OutputStream.Close(); continue
      }

      if ($req.HttpMethod -eq "POST" -and $path -eq "/__mkmdb"){
        $sr = New-Object System.IO.StreamReader($req.InputStream, [System.Text.Encoding]::UTF8)
        $jsonText = $sr.ReadToEnd(); $sr.Close()
        try {
          $mdbBytes = Convert-JsonToMdb $jsonText
          $res.StatusCode = 200; $res.ContentType = "application/x-msaccess"
          $res.ContentLength64 = $mdbBytes.Length
          $res.OutputStream.Write($mdbBytes, 0, $mdbBytes.Length)
        } catch {
          $msg = [System.Text.Encoding]::UTF8.GetBytes("MKMDB_ERROR: " + $_.Exception.Message)
          $res.StatusCode = 500; $res.ContentType = "text/plain; charset=utf-8"
          $res.ContentLength64 = $msg.Length
          $res.OutputStream.Write($msg, 0, $msg.Length)
        }
        $res.OutputStream.Close()
        continue
      }

      if ($req.HttpMethod -eq "POST" -and $path -eq "/__mdb"){
        $ms = New-Object System.IO.MemoryStream
        $req.InputStream.CopyTo($ms)
        try {
          $jsonText = Convert-MdbToJson $ms.ToArray()
          $bytes = [System.Text.Encoding]::UTF8.GetBytes($jsonText)
          $res.StatusCode = 200; $res.ContentType = "application/json; charset=utf-8"
          $res.ContentLength64 = $bytes.Length
          $res.OutputStream.Write($bytes, 0, $bytes.Length)
        } catch {
          $msg = [System.Text.Encoding]::UTF8.GetBytes("MDB_ERROR: " + $_.Exception.Message)
          $res.StatusCode = 500; $res.ContentType = "text/plain; charset=utf-8"
          $res.ContentLength64 = $msg.Length
          $res.OutputStream.Write($msg, 0, $msg.Length)
        }
        $res.OutputStream.Close()
        continue
      }

      if ($req.HttpMethod -eq "POST" -and $path -eq "/__pdf"){
        $ms = New-Object System.IO.MemoryStream
        $req.InputStream.CopyTo($ms)
        try {
          $pdfBytes = Convert-DocxToPdf $ms.ToArray()
          $res.StatusCode = 200; $res.ContentType = "application/pdf"
          $res.ContentLength64 = $pdfBytes.Length
          $res.OutputStream.Write($pdfBytes, 0, $pdfBytes.Length)
        } catch {
          $msg = [System.Text.Encoding]::UTF8.GetBytes("PDF_ERROR: " + $_.Exception.Message)
          $res.StatusCode = 500; $res.ContentType = "text/plain; charset=utf-8"
          $res.ContentLength64 = $msg.Length
          $res.OutputStream.Write($msg, 0, $msg.Length)
        }
        $res.OutputStream.Close()
        continue
      }

      if (-not $full.StartsWith($rootFull, [System.StringComparison]::OrdinalIgnoreCase)){
        $res.StatusCode = 403; $bytes = [System.Text.Encoding]::UTF8.GetBytes("403")
        $res.ContentType = "text/plain; charset=utf-8"
      } elseif (Test-Path -LiteralPath $full -PathType Leaf){
        $ext = [System.IO.Path]::GetExtension($full).ToLower()
        $res.ContentType = if ($mime.ContainsKey($ext)){ $mime[$ext] } else { "application/octet-stream" }
        $res.StatusCode = 200
        $bytes = [System.IO.File]::ReadAllBytes($full)
      } else {
        $res.StatusCode = 404; $bytes = [System.Text.Encoding]::UTF8.GetBytes("404: $path")
        $res.ContentType = "text/plain; charset=utf-8"
      }
      $res.ContentLength64 = $bytes.Length
      $res.OutputStream.Write($bytes, 0, $bytes.Length)
    } catch {
      try { $res.StatusCode = 500 } catch {}
    } finally {
      try { $res.OutputStream.Close() } catch {}
    }
  } catch {
    # ignora erros de conexão e continua
  }
}
