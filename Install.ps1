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
$moduleHome = "$env:ProgramFiles\WindowsPowerShell\Modules\PrinterExorcism"
$installedManifest = Join-Path $moduleHome 'PrinterExorcism.psd1'

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

# ───── Handle Progress Bars: Silence the workers ─────
function Invoke-NoProgress {
    param ( [scriptblock]$Script )

    $oldPref = $ProgressPreference
    try   { $Global:ProgressPreference = 'SilentlyContinue'; & $Script }
    finally { $Global:ProgressPreference = $oldPref }
}

# ───── PHASE 1: Summon the Exorcist ─────
Write-Host "🧭 Tracking down the Exorcist..." -ForegroundColor Cyan
Write-Host " "
Invoke-NoProgress {
    Invoke-WebRequest -Uri $zipUrl -OutFile $zipFile -UseBasicParsing
}

Write-Host "✅ The Exorcist has been found! Now begins the persuasion ritual..." -ForegroundColor Green
Start-Sleep -Seconds 3

Write-Host
Write-Host "$($Emoji.Star) Persuasion successful — the Exorcist is on board!" -ForegroundColor Green

Write-Host
Write-Host "$($Emoji.Gear) He's gathering holy relics and printer-banishment scrolls..." -ForegroundColor Yellow

# ───── PHASE 2: Unseal the Relics ─────
Write-Host
Write-Host "🏠 Relocating the Exorcist to his command chamber..." -ForegroundColor DarkCyan
Invoke-NoProgress {
    Expand-Archive -Path $zipFile -DestinationPath $env:TEMP -Force
}

New-Item -ItemType Directory -Path $moduleHome -Force | Out-Null
Copy-Item (Join-Path $destPath '*') $moduleHome -Recurse -Force
Start-Sleep -Seconds 2
Write-Host

# ───── PHASE 3: Binding the Exorcist ─────
Write-Host "📁 Unpacking the sacred arsenal..." -ForegroundColor Cyan
Import-Module $installedManifest -DisableNameChecking -Force

# ───── Final Rites ─────
Write-Host
Write-Host "🔱 The Exorcist is in place and ready to purge the unholy printer spirits!" -ForegroundColor Green
Write-Host "🔥 Run 'Invoke-PrinterExorcism' or 'Start-PrinterExorcismSession' to begin the reckoning." -ForegroundColor Magenta
