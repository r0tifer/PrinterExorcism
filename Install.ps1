# ───── CONFIG: Prepare for Summoning ─────
$zipUrl    = "https://github.com/r0tifer/PrinterExorcism/archive/refs/heads/main.zip"
$zipFile   = "$env:TEMP\PrinterExorcism.zip"
$destPath  = "$env:TEMP\PrinterExorcism-main"

# ───── PHASE 1: Summon the Exorcist ─────
Write-Host "`n🧭 Tracking down the Exorcist..." -ForegroundColor Cyan
Invoke-WebRequest -Uri $zipUrl -OutFile $zipFile -UseBasicParsing
Start-Sleep -Seconds 2

Write-Host "✅ The Exorcist has been found! Now begins the persuasion ritual..." -ForegroundColor Green
Start-Sleep -Seconds 3
Write-Host "🧳 He's packing up his holy relics and printer banishment scrolls..." -ForegroundColor Yellow

# ───── PHASE 2: Unseal the Relics ─────
Write-Host "`n📦 Unpacking the sacred arsenal..." -ForegroundColor Cyan
Expand-Archive -Path $zipFile -DestinationPath $env:TEMP -Force
Start-Sleep -Seconds 2

# ───── PHASE 3: Binding the Exorcist ─────
Write-Host "`n🏠 Relocating the Exorcist to his command chamber..." -ForegroundColor DarkCyan
Import-Module (Join-Path $destPath 'PrinterExorcism.psm1') -Force

# ───── Final Rites ─────
Write-Host "`n🔱 The Exorcist is in place and ready to purge the unholy printer spirits!" -ForegroundColor Green
Write-Host "🔥 Run 'Start-PrinterExorcismSession' to begin the reckoning." -ForegroundColor Magenta
