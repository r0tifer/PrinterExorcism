<#
.SYNOPSIS
    Ensure the classic ConsoleHost speaks UTF-8 so supplementary emoji render.
#>

[CmdletBinding()]
param(
    [ValidateSet('Auto', 'AllUsers', 'CurrentUser')]
    [string]$Scope = 'Auto'
)

$ErrorActionPreference = 'Stop'

if ($Host.Name -eq 'ConsoleHost' -and [Console]::OutputEncoding.CodePage -ne 65001) {
    try {
        chcp 65001 > $null                       # switch code page
        [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)  # no BOM
    } catch {
        Write-Warning "⚠️  Could not switch console to UTF-8: $_"
    }
}

# ───── CONFIG: Prepare for Summoning ─────
$zipUrl      = "https://github.com/r0tifer/PrinterExorcism/archive/refs/heads/main.zip"
$tempRoot    = if ([string]::IsNullOrWhiteSpace($env:TEMP)) { [System.IO.Path]::GetTempPath() } else { $env:TEMP }
$zipFile     = Join-Path $tempRoot 'PrinterExorcism.zip'
$extractRoot = Join-Path $tempRoot ("PrinterExorcism-{0}" -f ([guid]::NewGuid().ToString('N')))
$destPath    = Join-Path $extractRoot 'PrinterExorcism-main'

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

function Test-IsAdministrator {
    try {
        $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = [System.Security.Principal.WindowsPrincipal]::new($identity)
        return $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch {
        return $false
    }
}

function Get-ModuleRootFromPath {
    param(
        [ValidateSet('AllUsers', 'CurrentUser')]
        [string]$InstallScope
    )

    $modulePaths = @($env:PSModulePath -split [regex]::Escape([string][System.IO.Path]::PathSeparator)) |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    if ($InstallScope -eq 'AllUsers') {
        $programFiles = [Environment]::GetFolderPath([Environment+SpecialFolder]::ProgramFiles)
        if ([string]::IsNullOrWhiteSpace($programFiles)) {
            $programFiles = $env:ProgramFiles
        }

        $psHomePath = try { (Resolve-Path -LiteralPath $PSHOME).Path } catch { $PSHOME }
        $allUsersPath = $modulePaths |
            Where-Object {
                $_.StartsWith($programFiles, [System.StringComparison]::OrdinalIgnoreCase) -and
                -not $_.StartsWith($psHomePath, [System.StringComparison]::OrdinalIgnoreCase) -and
                ($_.TrimEnd('\', '/') -match '[\\/]Modules$')
            } |
            Select-Object -First 1

        if (-not [string]::IsNullOrWhiteSpace($allUsersPath)) {
            return $allUsersPath
        }

        return (Join-Path $programFiles 'WindowsPowerShell\Modules')
    }

    $userProfile = [Environment]::GetFolderPath([Environment+SpecialFolder]::UserProfile)
    if ([string]::IsNullOrWhiteSpace($userProfile)) {
        $userProfile = $env:USERPROFILE
    }

    $currentUserPath = $modulePaths |
        Where-Object {
            -not [string]::IsNullOrWhiteSpace($userProfile) -and
            $_.StartsWith($userProfile, [System.StringComparison]::OrdinalIgnoreCase) -and
            ($_.TrimEnd('\', '/') -match '[\\/]Modules$')
        } |
        Select-Object -First 1

    if (-not [string]::IsNullOrWhiteSpace($currentUserPath)) {
        return $currentUserPath
    }

    $documents = [Environment]::GetFolderPath([Environment+SpecialFolder]::MyDocuments)
    if ([string]::IsNullOrWhiteSpace($documents)) {
        $documents = Join-Path $userProfile 'Documents'
    }

    $profileModuleRoot = if ($PSVersionTable.PSEdition -eq 'Core') { 'PowerShell\Modules' } else { 'WindowsPowerShell\Modules' }
    return (Join-Path $documents $profileModuleRoot)
}

function Install-PrinterExorcismModule {
    param(
        [ValidateSet('AllUsers', 'CurrentUser')]
        [string]$InstallScope
    )

    if ($InstallScope -eq 'AllUsers' -and -not (Test-IsAdministrator)) {
        throw "The AllUsers install scope requires an elevated PowerShell process. Rerun as administrator or use -Scope CurrentUser."
    }

    $moduleRoot = Get-ModuleRootFromPath -InstallScope $InstallScope
    $moduleHome = Join-Path $moduleRoot 'PrinterExorcism'
    $installedManifest = Join-Path $moduleHome 'PrinterExorcism.psd1'

    Write-Host "Installing to $moduleHome ($InstallScope scope)..." -ForegroundColor DarkCyan
    New-Item -ItemType Directory -Path $moduleHome -Force | Out-Null
    Copy-Item -Path (Join-Path $destPath '*') -Destination $moduleHome -Recurse -Force

    if (-not (Test-Path -LiteralPath $installedManifest -PathType Leaf)) {
        throw "The module manifest was not copied to '$installedManifest'."
    }

    [pscustomobject]@{
        Scope    = $InstallScope
        Home     = $moduleHome
        Manifest = $installedManifest
    }
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
New-Item -ItemType Directory -Path $extractRoot -Force | Out-Null
Invoke-NoProgress {
    Expand-Archive -Path $zipFile -DestinationPath $extractRoot -Force
}

$requestedScope = $Scope
$installScope = if ($Scope -eq 'Auto') {
    if (Test-IsAdministrator) { 'AllUsers' } else { 'CurrentUser' }
} else {
    $Scope
}

try {
    $installResult = Install-PrinterExorcismModule -InstallScope $installScope
} catch {
    $installError = $_
    $installErrorText = $installError | Out-String
    $isAccessDenied = $installError.Exception.GetType().FullName -match 'UnauthorizedAccess' -or
        $installErrorText -match 'access.*denied|unauthorized'

    if ($requestedScope -eq 'Auto' -and $installScope -eq 'AllUsers' -and $isAccessDenied) {
        Write-Warning "Could not write to the all-users module path. Falling back to the current-user module path."
        $installResult = Install-PrinterExorcismModule -InstallScope 'CurrentUser'
    } else {
        throw "Failed to install PrinterExorcism using $installScope scope. $($installError.Exception.Message)"
    }
}

Start-Sleep -Seconds 2
Write-Host

# ───── PHASE 3: Binding the Exorcist ─────
Write-Host "📁 Unpacking the sacred arsenal..." -ForegroundColor Cyan
Import-Module $installResult.Manifest -DisableNameChecking -Force

# ───── Final Rites ─────
Write-Host
Write-Host "🔱 The Exorcist is in place at $($installResult.Home) and ready to purge the unholy printer spirits!" -ForegroundColor Green
Write-Host "🔥 Run 'Invoke-PrinterExorcism' or 'Start-PrinterExorcismSession' to begin the reckoning." -ForegroundColor Magenta
