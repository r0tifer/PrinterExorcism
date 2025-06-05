<#
.SYNOPSIS
    Safe, read-only discovery of printers and related registry keys.

.DESCRIPTION
    Loads current or target user's hive, enumerates all user-level printer keys,
    detects GPO-deployed printers, identifies default and ghost candidates.
    Outputs to terminal, and optionally JSON if -JSON is supplied.

.NOTES
    This script does don't perform any changes or destructive actions.The sciprt only enumerate the currentl mapped printers.
#>

param (
    [string]$TargetUser,
    [switch]$JSON
)

# ─── Import Shared Logger ──────────────────────────────
. "$PSScriptRoot\..\Common.ps1"

$LogPath = Join-Path $env:TEMP "PrinterDiscovery.log"
$ConsoleLevel = [LogVerbosity]::Info

function Log {
    param (
        [string]$Message,
        [LogVerbosity]$Level = [LogVerbosity]::Info
    )
    Log-PrinterEvent -msg $Message -Level $Level.ToString() -LogPath $LogPath -Verbosity $ConsoleLevel
}

# ─── Setup ─────────────────────────────────────────────
$Mounted = $false
if ($TargetUser) {
    $HivePath = "C:\Users\$TargetUser\NTUSER.DAT"
    $MountPoint = "HKU\TempHive_Discovery"

    try {
        reg load $MountPoint $HivePath | Out-Null
        $Root = "Registry::HKU\TempHive_Discovery"
        $Mounted = $true
        Log "🔍 Loaded hive for discovery: $TargetUser"
    } catch {
        Log "❌ Failed to load user hive: $_" Critical
        return
    }
} else {
    $Root = "HKCU:"
    $TargetUser = $env:USERNAME
}

$Discovery = [ordered]@{
    user = $TargetUser
    default_printer = $null
    registry_entries = @{}
    printers = @()
    timestamp = (Get-Date).ToString("s")
}

Write-Host "`n🖨  Printer Discovery for: $TargetUser" -ForegroundColor Cyan

# ─── Registry Printer Keys ─────────────────────────────
$paths = @(
    "Printers\Connections",
    "Printers\DevModePerUser",
    "Printers\DevModes2",
    "Software\Microsoft\Windows NT\CurrentVersion\Devices",
    "Software\Microsoft\Windows NT\CurrentVersion\PrinterPorts",
    "Software\Policies\Microsoft\Windows NT\Printers\Connections"
)

foreach ($sub in $paths) {
    $full = Join-Path $Root $sub
    if (Test-Path $full) {
        Write-Host "`n📁 $sub" -ForegroundColor Yellow
        $entries = Get-ChildItem -Path $full | ForEach-Object { $_.PSChildName }
        $entries | ForEach-Object { Write-Host "   - $_" }
        $Discovery.registry_entries[$sub] = $entries
    }
}

# ─── Default Printer ──────────────────────────────────
$defaultKey = Join-Path $Root "Software\Microsoft\Windows NT\CurrentVersion\Windows"
try {
    if (Test-Path $defaultKey) {
        $defaultPrinter = (Get-ItemProperty -Path $defaultKey -Name "Device").Device.Split(",")[0]
        $Discovery.default_printer = $defaultPrinter
        Write-Host "`n⭐ Default Printer: $defaultPrinter" -ForegroundColor Magenta
    }
} catch {
    Log "⚠️  Could not read default printer: $_" Warning
}

# ─── Get-WmiObject Results ─────────────────────────────
try {
    $Printers = Get-WmiObject -Class Win32_Printer
    Write-Host "`n🖨  WMI Printers:" -ForegroundColor Green
    foreach ($p in $Printers) {
        $printer = [ordered]@{
            name = $p.Name
            source = "WMI"
            is_default = ($p.Name -eq $Discovery.default_printer)
            ghost_candidate = ($p.Name -like "Copy*") -or ($p.Name -like "PPO*") -or ($p.Name -like "Front Desk*")
            gpo_linked = $false
        }

        if ($Discovery.registry_entries["Software\Policies\Microsoft\Windows NT\Printers\Connections"] -contains $p.Name) {
            $printer.gpo_linked = $true
        }

        $label = @()
        if ($printer.is_default) { $label += "Default" }
        if ($printer.gpo_linked) { $label += "GPO" }
        if ($printer.ghost_candidate) { $label += "GhostCandidate" }

        Write-Host "   - $($p.Name) [$($label -join ', ')]"
        $Discovery.printers += $printer
    }
} catch {
    Log "⚠️  Failed to get printers from WMI: $_" Warning
}

# ─── Get-Printer (if available) ───────────────────────
try {
    $allPrinters = Get-Printer
    Write-Host "`n🖨  Get-Printer Results:" -ForegroundColor Blue
    foreach ($gp in $allPrinters) {
        $exists = $Discovery.printers | Where-Object { $_.name -eq $gp.Name }
        if (-not $exists) {
            $printer = [ordered]@{
                name = $gp.Name
                source = "Get-Printer"
                is_default = ($gp.Name -eq $Discovery.default_printer)
                ghost_candidate = ($gp.Name -like "Copy*") -or ($gp.Name -like "PPO*") -or ($gp.Name -like "Front Desk*")
                gpo_linked = $false
            }

            if ($Discovery.registry_entries["Software\Policies\Microsoft\Windows NT\Printers\Connections"] -contains $gp.Name) {
                $printer.gpo_linked = $true
            }

            $label = @()
            if ($printer.is_default) { $label += "Default" }
            if ($printer.gpo_linked) { $label += "GPO" }
            if ($printer.ghost_candidate) { $label += "GhostCandidate" }

            Write-Host "   - $($gp.Name) [$($label -join ', ')]"
            $Discovery.printers += $printer
        }
    }
} catch {
    Log "⚠️  Get-Printer failed: $_" Warning
}

# ─── Output JSON if requested ─────────────────────────
if ($JSON) {
    $jsonOut = Join-Path $env:TEMP "PrinterDiscovery.$($TargetUser).json"
    $Discovery | ConvertTo-Json -Depth 5 | Set-Content -Path $jsonOut -Encoding UTF8
    Write-Host "`n📄 Discovery saved to: $jsonOut" -ForegroundColor Gray
}

# ─── Cleanup ──────────────────────────────────────────
if ($Mounted) {
    reg unload $MountPoint | Out-Null
    Write-Host "`n🧹 Unmounted user hive: $TargetUser" -ForegroundColor DarkGray
}
