# PrinterExorcist.ps1
# Two-phase printer cleanup tool with structured JSON logging
# Phase 1: Run as local user, perform cleanup, log failures
# Phase 2: Retry failed steps using elevation

param(
    [switch]$FullCleanup,
    [switch]$CompareGPO,
    [switch]$Automated,
    [string]$RetryPrinters,
    [string]$RetryGhosts,
    [string]$TargetUser,
    [switch]$RetryOnly,
    [ValidateSet("Info", "Warning", "Critical", "Debug")]
    [string]$Verbosity = "Critical"

)

enum LogVerbosity {
    Info = 0
    Warning = 1
    Critical = 2
    Debug = 3
}

# ───── Import shared log system ─────
. "$PSScriptRoot\Common.ps1"

# ───── Config ─────
$BuiltinPrinters = @(
    "Microsoft Print to PDF", "Microsoft XPS Document Writer",
    "Fax", "OneNote for Windows 10", "OneNote (Desktop)", "Adobe PDF"
)

if ($Automated) {
    # Local path avoids UNC prompts
    $DesktopPath = Join-Path $env:LOCALAPPDATA "PrinterExorcist"
    if (-not (Test-Path $DesktopPath)) {
        New-Item -ItemType Directory -Path $DesktopPath -Force | Out-Null
    }
} else {
    $DesktopPath = [Environment]::GetFolderPath("Desktop")
    if (-not (Test-Path $DesktopPath)) {
        $DesktopPath = Join-Path $env:USERPROFILE "Documents"
    }
}

$LogPath    = Join-Path $DesktopPath "PrinterCleanup.log"
$StatusPath = Join-Path $DesktopPath "PrinterCleanup.status.json"

# Define default console logging
$ConsoleLevel = [LogVerbosity]::Info

function Log {
    param (
        [string]$Message,
        [LogVerbosity]$Level = [LogVerbosity]::Info
    )
    Log-PrinterEvent -msg $Message -Level $Level.ToString() -LogPath $LogPath -Verbosity $ConsoleLevel
}

function Is-Admin {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

function Remove-HKLMPrinterConnections {
    Log "Clearing HKLM printer connection GUIDs..." "Info"

    $regBase = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Connections"
    if (Test-Path $regBase) {
        Get-ChildItem -Path $regBase | ForEach-Object {
            Log "🗑️  Removing: $($_.Name)" "Info"
            Remove-Item -Path $_.PsPath -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    # Optional: Also purge ghosted printer entries from full list
    Get-ChildItem "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers" |
        Where-Object { $_.PSChildName -match "Copy|PPO|Front" } |
        ForEach-Object {
            Log "🪓 Removing ghost printer: $($_.PSChildName)" "Info"
            Remove-Item -Path $_.PsPath -Recurse -Force -ErrorAction SilentlyContinue
        }

    Log "🔁 Restarting Print Spooler..." "Info"
    Restart-Service spooler -Force
}

function Write-StatusJson {
    param (
        [string[]]$Printers,
        [string[]]$Ghosts,
        [string[]]$CleanedPrinters,
        [string[]]$CleanedGhosts,
        [int]$Phase
    )

    $status = @{
        user             = $TargetUser
        elevated         = $IsAdmin
        phase            = $Phase
        failed_printers  = $Printers
        failed_ghosts    = $Ghosts
        cleaned_printers = $CleanedPrinters
        cleaned_ghosts   = $CleanedGhosts
        timestamp        = (Get-Date).ToString("s")
    }

    $json = $status | ConvertTo-Json -Depth 3
    $json | Set-Content -Path $StatusPath -Encoding UTF8
}

# Check if TargetUser is defined
if (-not $TargetUser) {
    $TargetUser = $env:USERNAME
    Log "No -TargetUser supplied. Defaulting to current user: $TargetUser" "Info"
}

Log "Starting printer cleanup for user: $TargetUser" "Info"

$IsAdmin = Is-Admin
$CleanedPrinters = @()
$CleanedGhosts   = @()
$FailedPrinters = @()
$FailedGhosts = @()

# ───── Registry Context ─────
$UserHivePath = "C:\Users\$TargetUser\NTUSER.DAT"
$MountKey = "HKU\TempUserHive"
$HKCU = "Registry::$MountKey"

# Only mount if not already mounted
if (-not (Test-Path $HKCU)) {
    try {
        reg load $MountKey $UserHivePath | Out-Null
        Log "Mounted user hive for: $TargetUser" "Info"
    } catch {
        Log "Failed to load user hive at ${UserHivePath}: $_" "Critical"
        exit 1
    }
}

# ───── User-space Registry Cleanup ─────
if (-not $RetryOnly) {
    Log "Removing user-space printer registry entries..." "Info"
    $UserRegPaths = @(
        "$HKCU\\Printers\\DevModePerUser",
        "$HKCU\\Printers\\DevModes2",
        "$HKCU\\Software\\Microsoft\\Windows NT\\CurrentVersion\\Devices",
        "$HKCU\\Software\\Microsoft\\Windows NT\\CurrentVersion\\PrinterPorts"
    )
    foreach ($path in $UserRegPaths) {
        if (Test-Path $path) {
            try {
                Remove-Item -Path $path -Recurse -Force -ErrorAction Stop
                Log "Removed: $path" "Info"
            } catch {
                Log "Failed to remove ${path}: $_" "Warning"
            }
        } else {
            Log "No entry at: $path" "Debug"
        }
    }

    $ConnectionsKey = "$HKCU\\Printers\\Connections"
    if ($CompareGPO) {
        $PolicyKey = "$HKCU\\Software\\Policies\\Microsoft\\Windows NT\\Printers\\Connections"
        $Active = @(Get-ChildItem -Path $ConnectionsKey -ErrorAction SilentlyContinue | ForEach-Object { $_.PSChildName })
        $Gpo = @(Get-ChildItem -Path $PolicyKey -ErrorAction SilentlyContinue | ForEach-Object { $_.PSChildName })
        $ToRemove = $Active | Where-Object { $Gpo -notcontains $_ }
        foreach ($conn in $ToRemove) {
            try {
                Remove-Item -Path (Join-Path $ConnectionsKey $conn) -Recurse -Force
                Log "Removed non-GPO printer connection: $conn" "Info"
            } catch {
                Log "Failed to remove non-GPO connection: $conn" "Warning"
            }
        }
    } else {
        if (Test-Path $ConnectionsKey) {
            try {
                Remove-Item -Path $ConnectionsKey -Recurse -Force
                Log "Removed all printer connections" "Info"
            } catch {
                Log "Failed to remove connections: $_" "Warning"
            }
        } else {
            Log "No printer connections found" "Debug"
        }
    }

    $DefaultKey = "$HKCU\\Software\\Microsoft\\Windows NT\\CurrentVersion\\Windows"
    if (Test-Path $DefaultKey) {
        try {
            $defaultDevice = (Get-ItemProperty -Path $DefaultKey -Name "Device").Device
            $deviceName = $defaultDevice.Split(",")[0]
            $Printers = Get-WmiObject -Query "Select * from Win32_Printer"
            if ($BuiltinPrinters -notcontains $deviceName -and ($Printers.Name -notcontains $deviceName)) {
                Remove-ItemProperty -Path $DefaultKey -Name "Device"
                Log "🧹 Removed invalid default printer: $deviceName" "Info"
            } else {
                Log "Default printer is valid: $deviceName" "Debug"
            }
        } catch {
            Log "Could not inspect default printer setting: $_" "Warning"
        }
    }
}

# ───── WMI Printer Cleanup ─────
$Printers = Get-WmiObject -Query "Select * from Win32_Printer"
foreach ($printer in $Printers) {
    if ($BuiltinPrinters -notcontains $printer.Name) {
        if ($RetryOnly) {
            if ($RetryPrinters -and ($RetryPrinters -split "\|") -contains $printer.Name) {
                try {
                    $printer.Delete() | Out-Null
                    Log "Retried and removed: $($printer.Name)" "Info"
                } catch {
                    Log "Still failed to remove: $($printer.Name)" "Critical"
                }
            }
        } else {
            try {
                $printer.Delete() | Out-Null
                Log "Removed printer: $($printer.Name)" "Info"
                $CleanedPrinters += $printer.Name
            } catch {
                Log "Failed to remove: $($printer.Name)" "Warning"
                $FailedPrinters += $printer.Name
            }
        }
    } elseif (-not $RetryOnly) {
        Log "Keeping built-in printer: $($printer.Name)" "Debug"
    }
}

# ───── Ghost printer detection deferral ─────
if (-not $RetryOnly -and -not $IsAdmin) {
    $FailedGhosts += "_DEFER_GHOST_CLEANUP_"
    Log "Skipped ghost detection: elevation required. Scheduling retry..." "Warning"
}

# ───── PnP Ghost Detection and Cleanup ─────
if ($IsAdmin -or $RetryOnly) {
    try {
        Write-Verbose "🔍 Enumerating printers via Get-Printer..."
        $AllPrinters = Get-Printer -ErrorAction Stop
    }
    catch {
        Log "Could not run Get-Printer: $($_.Exception.Message)" "Warning"
        $AllPrinters = @()
    }

    $GhostPatterns = @(
        'Copy*',
        'PPO*',
        'Front Desk Main*'
    )
    $FoundGhosts = @()
    foreach ($pattern in $GhostPatterns) {
        $FoundGhosts += $AllPrinters | Where-Object { $_.Name -like $pattern }
    }
    $FoundGhosts = $FoundGhosts | Sort-Object -Unique

    if ($FoundGhosts.Count -gt 0) {
        foreach ($ghost in $FoundGhosts) {
            try {
                Remove-Printer -Name $ghost.Name -ErrorAction Stop
                Log "Removed ghost printer via Remove-Printer: $($ghost.Name)" "Info"
                $CleanedGhosts += $ghost.Name
            }
            catch {
                Log "Failed to remove ghost printer via Remove-Printer: $($ghost.Name) - $_" "Critical"
                $FailedGhosts += $ghost.Name
            }
        }
    }
    else {
        Log "No ghost printers detected by Get-Printer." "Info"
    }

    $regPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers"
    try {
        $RegChildren = Get-ChildItem $regPath -ErrorAction Stop
    }
    catch {
        Log "Could not enumerate HKLM prn registry hive: $($_.Exception.Message)" "Warning"
        $RegChildren = @()
    }

    $RegGhosts = $RegChildren | Where-Object {
        $_.PSChildName -like 'Copy*' -or
        $_.PSChildName -like 'PPO*'  -or
        $_.PSChildName -like 'Front Desk Main*'
    }

    if ($RegGhosts.Count -gt 0) {
        foreach ($r in $RegGhosts) {
            $keyPath = "$regPath\$($r.PSChildName)"
            try {
                Remove-Item -Path $keyPath -Recurse -Force -ErrorAction Stop
                Log "Removed registry ghost key: $($r.PSChildName)" "Info"
                $CleanedGhosts += $r.PSChildName
            }
            catch {
                Log "Failed to remove registry ghost key: $($r.PSChildName) - $_" "Critical"
                $FailedGhosts += $r.PSChildName
            }
        }
    }
    elseif ($FoundGhosts.Count -eq 0) {
        Log "No registry-based ghost keys found." "Debug"
    }
}

# ───── HKLM Printer Connection Cleanup ─────
if ($IsAdmin) {
    Remove-HKLMPrinterConnections
}

# ───── Write status for this phase ─────
$phaseNum = if ($RetryOnly) { 2 } else { 1 }
Write-StatusJson -Printers $FailedPrinters -Ghosts $FailedGhosts -CleanedPrinters $CleanedPrinters -CleanedGhosts $CleanedGhosts -Phase $phaseNum

# ───── Elevation Retry Phase ─────
if (-not $RetryOnly -and ($FailedPrinters -or $FailedGhosts)) {
    $scriptPath = $MyInvocation.MyCommand.Definition
    $printerArg = ($FailedPrinters -join '|')
    $ghostArg   = ($FailedGhosts   -join '|')

    $argsList   = @(
        '-ExecutionPolicy', 'Bypass',
        '-File',           "`"$scriptPath`"",
        '-RetryOnly',
        '-RetryPrinters',  "`"$printerArg`"",
        '-RetryGhosts',    "`"$ghostArg`""
    )

    Log "Some cleanup failed. Relaunching with elevation to retry targeted steps…" "Warning"
    Log "Elevation Command: powershell.exe $($argsList -join ' ')" "Debug"

    $p = Start-Process -FilePath powershell.exe `
                       -ArgumentList $argsList `
                       -Verb RunAs -WindowStyle Hidden `
                       -PassThru -Wait

    if ($p.ExitCode -ne 0) {
        Log "Elevated run cancelled or failed (exit $($p.ExitCode))." "Critical"
    }
    exit
}


# ───── Cleanup Mounted Hive ─────
if ($HKCU -eq "Registry::$MountKey") {
    try {
        reg unload $MountKey | Out-Null
        Log "Unmounted hive for $TargetUser" "Debug"
    } catch {
        Log "Failed to unload hive for ${TargetUser}: $_" "Warning"
    }
}

# ───── Announce Finish ─────
Log "✅ Printer cleanup complete for user: $TargetUser" "Info"
