$zipUrl = "https://github.com/r0tifer/PrinterExorcism/archive/refs/heads/main.zip"
$zipFile = "$env:TEMP\\PrinterExorcism.zip"
$destPath = "$env:TEMP\\PrinterExorcism-main"

Invoke-WebRequest -Uri $zipUrl -OutFile $zipFile -UseBasicParsing
Expand-Archive -Path $zipFile -DestinationPath $env:TEMP -Force
Import-Module (Join-Path $destPath 'PrinterExorcism.psm1') -Force
Write-Host "PrinterExorcism loaded from GitHub!" -ForegroundColor Green
