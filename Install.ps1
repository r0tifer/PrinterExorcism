<#
.SYNOPSIS
    Ensure the classic ConsoleHost speaks UTF-8 so supplementary emoji render.
#>

if ($Host.Name -eq 'ConsoleHost' -and [Console]::OutputEncoding.CodePage -ne 65001) {
    try {
        chcp 65001 > $null                       # switch code page
        [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)  # no BOM
    } catch {
        Write-Warning "⚠️  Could not switch console to UTF-8: $_"
    }
}

# ───── CONFIG: Prepare for Summoning ─────
$zipUrl    = "https://github.com/r0tifer/PrinterExorcism/archive/refs/heads/main.zip"
$zipFile   = "$env:TEMP\PrinterExorcism.zip"
$destPath  = "$env:TEMP\PrinterExorcism-main"

# ───── Emoji table: Prepare for Summoning ─────
$Emoji = @{
    OK          = '✅'      # 9989
    Hammer      = '⚒'      # 9874
    Anvil       = '⚙'      # 9881
    Coffee      = '☕'      # 9749
    Star        = '⭐'      # 11088
    Phone       = '☎'      # 9742
    Scissors    = '✂'      # 9986
    Check       = '✔'      # 10004
    Gear        = '⚙'
}

# ───── PHASE 1: Summon the Exorcist ─────
Write-Host "🧭 Tracking down the Exorcist..." -ForegroundColor Cyan
Invoke-WebRequest -Uri $zipUrl -OutFile $zipFile -UseBasicParsing
Start-Sleep -Seconds 2

Write-Host "✅ The Exorcist has been found! Now begins the persuasion ritual..." -ForegroundColor Green
Start-Sleep -Seconds 3
Write-Host "$($Emoji.Star) Persuasion successful — the Exorcist is on board!" -ForegroundColor Green

Write-Host ""
Write-Host "$($Emoji.Gear) He's gathering holy relics and printer-banishment scrolls..." -ForegroundColor Yellow

# ───── PHASE 2: Unseal the Relics ─────
Write-Host ""
Write-Host "📁 Unpacking the sacred arsenal..." -ForegroundColor Cyan
Expand-Archive -Path $zipFile -DestinationPath $env:TEMP -Force
Start-Sleep -Seconds 2

# ───── PHASE 3: Binding the Exorcist ─────
Write-Host "🏠 Relocating the Exorcist to his command chamber..." -ForegroundColor DarkCyan
Import-Module (Join-Path $destPath 'PrinterExorcism.psm1') -Force

# ───── Final Rites ─────
Write-Host ""
Write-Host "🔱 The Exorcist is in place and ready to purge the unholy printer spirits!" -ForegroundColor Green
Write-Host "🔥 Run 'Invoke-PrinterExorcism' to begin the reckoning." -ForegroundColor Magenta
