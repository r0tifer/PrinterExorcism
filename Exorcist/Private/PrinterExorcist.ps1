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
    [switch]$NoSelfElevation,
    [string]$StatusPath,
    [string]$LogPath,
    [ValidateSet("Info", "Warning", "Critical", "Debug")]
    [string]$Verbosity = "Critical"

)

enum LogVerbosity {
    Info = 0
    Warning = 1
    Critical = 2
    Debug = 3
}

# Import shared log system
    try {
        . "$PSScriptRoot\..\Common.ps1"
    } catch {
        Write-Host "Failed to import Common.ps1 logging module: $_" -ForegroundColor Red
        return 100
    }

# Config
$BuiltinPrinters = @(
    "Microsoft Print to PDF", "Microsoft XPS Document Writer",
    "Fax", "OneNote for Windows 10", "OneNote (Desktop)", "Adobe PDF"
)

if (-not $LogPath -or -not $StatusPath) {
    if ($Automated) {
        # Local path avoids UNC prompts
        $OutputDir = Join-Path $env:LOCALAPPDATA "PrinterExorcist"
        if (-not (Test-Path $OutputDir)) {
            New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
        }
    } else {
        $OutputDir = [Environment]::GetFolderPath("Desktop")
        if (-not (Test-Path $OutputDir)) {
            $OutputDir = Join-Path $env:USERPROFILE "Documents"
        }
    }

    if (-not $LogPath) {
        $LogPath = Join-Path $OutputDir "PrinterCleanup.log"
    }
    if (-not $StatusPath) {
        $StatusPath = Join-Path $OutputDir "PrinterCleanup.status.json"
    }
}

$OutputDirs = @($LogPath, $StatusPath) |
    Where-Object { $_ } |
    ForEach-Object { Split-Path $_ -Parent } |
    Sort-Object -Unique

foreach ($dir in $OutputDirs) {
    if ($dir -and -not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
}

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

function Get-LoadedUserHiveRoot {
    param(
        [Parameter(Mandatory)]
        [string]$UserName
    )

    $profileList = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"

    try {
        $profile = Get-ChildItem -Path $profileList -ErrorAction Stop |
            ForEach-Object {
                $props = Get-ItemProperty -Path $_.PSPath -Name ProfileImagePath -ErrorAction SilentlyContinue
                [pscustomobject]@{
                    Sid = $_.PSChildName
                    ProfileImagePath = $props.ProfileImagePath
                }
            } |
            Where-Object {
                $_.ProfileImagePath -and (Split-Path $_.ProfileImagePath -Leaf) -ieq $UserName
            } |
            Select-Object -First 1

        if ($profile) {
            $root = "Registry::HKEY_USERS\$($profile.Sid)"
            if (Test-Path $root) {
                return $root
            }
        }
    } catch {
        Log "Failed to resolve loaded hive for ${UserName}: $_" "Debug"
    }

    return $null
}

function Get-UserSid {
    param(
        [Parameter(Mandatory)]
        [string]$UserName
    )

    $profileList = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"

    try {
        $profile = Get-ChildItem -Path $profileList -ErrorAction Stop |
            ForEach-Object {
                $props = Get-ItemProperty -Path $_.PSPath -Name ProfileImagePath -ErrorAction SilentlyContinue
                [pscustomobject]@{
                    Sid = $_.PSChildName
                    ProfileImagePath = $props.ProfileImagePath
                }
            } |
            Where-Object {
                $_.ProfileImagePath -and (Split-Path $_.ProfileImagePath -Leaf) -ieq $UserName
            } |
            Select-Object -First 1

        if ($profile) {
            return $profile.Sid
        }
    } catch {
        Log "Failed to resolve SID for ${UserName}: $_" "Debug"
    }

    return $null
}

function Normalize-PrinterServerName {
    param(
        [string]$ServerName
    )

    if ([string]::IsNullOrWhiteSpace($ServerName)) {
        return $null
    }

    $normalized = $ServerName.Trim().TrimStart('\')
    if ($normalized -match '^(?<host>[^\.]+)\..+$') {
        $normalized = $Matches.host
    }

    return $normalized.ToLowerInvariant()
}

function Normalize-PrinterShareName {
    param(
        [string]$ShareName
    )

    if ([string]::IsNullOrWhiteSpace($ShareName)) {
        return $null
    }

    return $ShareName.Trim().ToLowerInvariant()
}

function Get-PrinterConnectionInfo {
    param(
        [string]$Name
    )

    if ([string]::IsNullOrWhiteSpace($Name)) {
        return $null
    }

    $value = $Name.Trim()
    $server = $null
    $share = $null

    if ($value -match '^\\\\(?<server>[^\\]+)\\(?<share>.+)$') {
        $server = $Matches.server
        $share = $Matches.share
    } elseif ($value -match '^,,(?<server>[^,]+),(?<share>.+)$') {
        $server = $Matches.server
        $share = $Matches.share
    } elseif ($value -match '^(?<share>.+?)\s+on\s+(?<server>.+)$') {
        $server = $Matches.server
        $share = $Matches.share
    }

    if (-not $server -or -not $share) {
        return $null
    }

    $serverShort = Normalize-PrinterServerName -ServerName $server
    $shareNormalized = Normalize-PrinterShareName -ShareName $share
    $canonicalKey = $null
    if ($serverShort -and $shareNormalized) {
        $canonicalKey = "$serverShort|$shareNormalized"
    }

    return [pscustomobject]@{
        RawName              = $Name
        Server               = $server.Trim()
        ServerShort          = $serverShort
        Share                = $share.Trim()
        ShareNormalized      = $shareNormalized
        CanonicalKey         = $canonicalKey
        PreferredDisplayName = "$($share.Trim()) on $($server.Trim())"
        PreferredPathName    = "\\$($server.Trim())\$($share.Trim())"
    }
}

function Get-GpoPrinterConnections {
    param(
        [string]$UserRegistryRoot = "HKCU:"
    )

    $paths = @(
        "$UserRegistryRoot\\Software\\Policies\\Microsoft\\Windows NT\\Printers\\Connections",
        "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\CurrentVersion\Print\Connections",
        "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers\Connections"
    ) | Select-Object -Unique

    $results = @()
    foreach ($path in $paths) {
        if (-not (Test-Path $path)) {
            continue
        }

        foreach ($item in Get-ChildItem -Path $path -ErrorAction SilentlyContinue) {
            $info = Get-PrinterConnectionInfo -Name $item.PSChildName
            if ($info) {
                $results += $info
            }
        }
    }

    return $results
}

function New-PrinterKeySet {
    return New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
}

function New-PrinterCountMap {
    return New-Object 'System.Collections.Generic.Dictionary[string,int]' ([System.StringComparer]::OrdinalIgnoreCase)
}

function Remove-ClientSideRenderingConnections {
    param(
        [string]$UserSid,
        [bool]$CompareGPO,
        [System.Collections.Generic.HashSet[string]]$AllowedGpoPrinterKeys
    )

    if ([string]::IsNullOrWhiteSpace($UserSid)) {
        return
    }

    $csrConnectionsPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Providers\Client Side Rendering Print Provider\$UserSid\Printers\Connections"
    if (-not (Test-Path $csrConnectionsPath)) {
        Log "No Client Side Rendering connection cache found for SID $UserSid" "Debug"
        return
    }

    foreach ($connection in Get-ChildItem -Path $csrConnectionsPath -ErrorAction SilentlyContinue) {
        $info = Get-PrinterConnectionInfo -Name $connection.PSChildName
        $keep = $false
        if ($CompareGPO -and $info -and $info.CanonicalKey) {
            $keep = $AllowedGpoPrinterKeys.Contains($info.CanonicalKey)
        }

        if ($keep) {
            Log "Keeping GPO-managed CSR connection: $($connection.PSChildName)" "Debug"
            continue
        }

        try {
            Remove-Item -Path $connection.PSPath -Recurse -Force -ErrorAction Stop
            Log "Removed CSR printer connection cache: $($connection.PSChildName)" "Info"
        } catch {
            Log "Failed to remove CSR printer connection cache: $($connection.PSChildName) - $_" "Warning"
        }
    }
}

function Remove-StalePrintEnumDevices {
    param(
        [string[]]$ActivePrinterNames,
        [string[]]$BuiltInPrinterNames
    )

    $printEnumPath = "HKLM:\SYSTEM\CurrentControlSet\Enum\SWD\PRINTENUM"
    if (-not (Test-Path $printEnumPath)) {
        Log "No SWD\\PRINTENUM device cache found." "Debug"
        return
    }

    $activeNameSet = New-PrinterKeySet
    $activeCanonicalKeySet = New-PrinterKeySet
    foreach ($printerName in $ActivePrinterNames) {
        if ([string]::IsNullOrWhiteSpace($printerName)) {
            continue
        }

        $null = $activeNameSet.Add($printerName)
        $info = Get-PrinterConnectionInfo -Name $printerName
        if ($info -and $info.CanonicalKey) {
            $null = $activeCanonicalKeySet.Add($info.CanonicalKey)
            if ($info.PreferredDisplayName) {
                $null = $activeNameSet.Add($info.PreferredDisplayName)
            }
            if ($info.PreferredPathName) {
                $null = $activeNameSet.Add($info.PreferredPathName)
            }
        }
    }

    $removedAny = $false

    foreach ($device in Get-ChildItem -Path $printEnumPath -ErrorAction SilentlyContinue) {
        $props = Get-ItemProperty -Path $device.PSPath -ErrorAction SilentlyContinue
        $friendlyName = $props.FriendlyName
        if ([string]::IsNullOrWhiteSpace($friendlyName)) {
            continue
        }

        if ($BuiltInPrinterNames -contains $friendlyName) {
            continue
        }

        $deviceInfo = Get-PrinterConnectionInfo -Name $friendlyName
        $isActive = $activeNameSet.Contains($friendlyName)
        if (-not $isActive -and $deviceInfo -and $deviceInfo.CanonicalKey) {
            $isActive = $activeCanonicalKeySet.Contains($deviceInfo.CanonicalKey)
        }

        if ($isActive) {
            Log "Keeping active PRINTENUM device: $friendlyName" "Debug"
            continue
        }

        $instanceId = "SWD\PRINTENUM\$($device.PSChildName)"
        $output = & pnputil.exe /remove-device $instanceId /force 2>&1
        $exitCode = $LASTEXITCODE

        if ($exitCode -eq 0 -or $exitCode -eq 3010) {
            Log "Removed stale PRINTENUM device: $friendlyName ($instanceId)" "Info"
            $script:CleanedGhosts += $friendlyName
            $removedAny = $true
            if ($exitCode -eq 3010) {
                Log "PRINTENUM device removal requested reboot: $friendlyName" "Warning"
            }
        } else {
            $message = if ($output) { ($output | Out-String).Trim() } else { "exit $exitCode" }
            Log "Failed to remove stale PRINTENUM device: $friendlyName ($instanceId) - $message" "Warning"
            $script:FailedGhosts += $friendlyName
        }
    }

    if ($removedAny) {
        $scanOutput = & pnputil.exe /scan-devices 2>&1
        $scanExitCode = $LASTEXITCODE
        if ($scanExitCode -eq 0) {
            Log "Triggered device rescan after PRINTENUM cleanup." "Info"
        } else {
            $message = if ($scanOutput) { ($scanOutput | Out-String).Trim() } else { "exit $scanExitCode" }
            Log "Device rescan after PRINTENUM cleanup reported: $message" "Warning"
        }
    }
}

function Remove-HKLMPrinterConnections {
    Log "Clearing HKLM printer connection GUIDs..." "Info"

    $regBase = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Connections"
    if (Test-Path $regBase) {
        Get-ChildItem -Path $regBase | ForEach-Object {
            Log "Removing: $($_.Name)" "Info"
            Remove-Item -Path $_.PsPath -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    # Optional: Also purge ghosted printer entries from full list
    Get-ChildItem "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers" |
        Where-Object { $_.PSChildName -match "Copy|PPO|Front" } |
        ForEach-Object {
            Log "Removing ghost printer: $($_.PSChildName)" "Info"
            Remove-Item -Path $_.PsPath -Recurse -Force -ErrorAction SilentlyContinue
        }

    Log "Restarting Print Spooler..." "Info"
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

# Registry Context
$CurrentUser = $env:USERNAME
$UserHivePath = "C:\Users\$TargetUser\NTUSER.DAT"
$MountKey = "HKU\TempUserHive"
$HKCU = "HKCU:"
$MountedHive = $false

if ($TargetUser -and $TargetUser -ne $CurrentUser) {
    $LoadedHiveRoot = Get-LoadedUserHiveRoot -UserName $TargetUser
    if ($LoadedHiveRoot) {
        $HKCU = $LoadedHiveRoot
        Log "Using already-loaded registry hive for: $TargetUser" "Info"
    } else {
        $HKCU = "Registry::$MountKey"

        # Only mount if not already mounted
        if (-not (Test-Path $HKCU)) {
            try {
                reg load $MountKey $UserHivePath | Out-Null
                $MountedHive = $true
                Log "Mounted user hive for: $TargetUser" "Info"
            } catch {
                Log "Failed to load user hive at ${UserHivePath}: $_" "Critical"
                exit 1
            }
        }
    }
} else {
    Log "Using live HKCU for current user: $TargetUser" "Info"
}

$GpoPrinterConnections = @()
$AllowedGpoPrinterKeys = New-PrinterKeySet
$PreferredGpoPrinterNames = New-PrinterKeySet
$TargetUserSid = Get-UserSid -UserName $TargetUser
if ($CompareGPO) {
    $GpoPrinterConnections = Get-GpoPrinterConnections -UserRegistryRoot $HKCU
    foreach ($connection in $GpoPrinterConnections) {
        if ($connection.CanonicalKey) {
            $null = $AllowedGpoPrinterKeys.Add($connection.CanonicalKey)
        }
        if ($connection.PreferredDisplayName) {
            $null = $PreferredGpoPrinterNames.Add($connection.PreferredDisplayName)
        }
        if ($connection.PreferredPathName) {
            $null = $PreferredGpoPrinterNames.Add($connection.PreferredPathName)
        }
    }

    Log "Discovered $($AllowedGpoPrinterKeys.Count) GPO-managed printer connection(s)." "Info"
}

# User-space Registry Cleanup
if (-not $RetryOnly) {
    Log "Removing user-space printer registry entries..." "Info"
    $UserRegPaths = @(
        "$HKCU\\Printers\\ConvertUserDevModesCount",
        "$HKCU\\Printers\\DevModePerUser",
        "$HKCU\\Printers\\DevModes2",
        "$HKCU\\Printers\\Defaults",
        "$HKCU\\Printers\\Settings",
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
        $Active = @(Get-ChildItem -Path $ConnectionsKey -ErrorAction SilentlyContinue)
        foreach ($conn in $Active) {
            $info = Get-PrinterConnectionInfo -Name $conn.PSChildName
            $keep = $false
            if ($info -and $info.CanonicalKey) {
                $keep = $AllowedGpoPrinterKeys.Contains($info.CanonicalKey)
            }
            if ($keep) {
                Log "Keeping GPO-managed printer connection: $($conn.PSChildName)" "Debug"
                continue
            }
            try {
                Remove-Item -Path $conn.PSPath -Recurse -Force
                Log "Removed non-GPO printer connection: $($conn.PSChildName)" "Info"
            } catch {
                Log "Failed to remove non-GPO connection: $($conn.PSChildName)" "Warning"
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
                Log "Removed invalid default printer: $deviceName" "Info"
            } else {
                Log "Default printer is valid: $deviceName" "Debug"
            }
        } catch {
            Log "Could not inspect default printer setting: $_" "Warning"
        }
    }
}

# WMI Printer Cleanup
$Printers = Get-WmiObject -Query "Select * from Win32_Printer"
foreach ($printer in $Printers) {
    if ($BuiltinPrinters -notcontains $printer.Name) {
        $printerInfo = Get-PrinterConnectionInfo -Name $printer.Name
        $isGpoManagedPrinter = $false
        if ($CompareGPO -and $printerInfo -and $printerInfo.CanonicalKey) {
            $isGpoManagedPrinter = $AllowedGpoPrinterKeys.Contains($printerInfo.CanonicalKey)
        }

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
            if ($CompareGPO -and $isGpoManagedPrinter) {
                Log "Keeping GPO-managed printer: $($printer.Name)" "Debug"
                continue
            }
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

# Ghost printer detection deferral
if (-not $RetryOnly -and -not $IsAdmin) {
    $FailedGhosts += "_DEFER_GHOST_CLEANUP_"
    Log "Skipped ghost detection: elevation required. Scheduling retry..." "Warning"
}

# PnP Ghost Detection and Cleanup
if ($IsAdmin -or $RetryOnly) {
    try {
        Write-Verbose "Enumerating printers via Get-Printer..."
        $AllPrinters = Get-Printer -ErrorAction Stop
    }
    catch {
        Log "Could not run Get-Printer: $($_.Exception.Message)" "Warning"
        $AllPrinters = @()
    }

    Remove-ClientSideRenderingConnections -UserSid $TargetUserSid -CompareGPO:$CompareGPO -AllowedGpoPrinterKeys $AllowedGpoPrinterKeys

    $GhostPatterns = @(
        'Copy*',
        'PPO*',
        'Front Desk Main*'
    )
    $CanonicalCounts = New-PrinterCountMap
    foreach ($printer in $AllPrinters) {
        $info = Get-PrinterConnectionInfo -Name $printer.Name
        if ($info -and $info.CanonicalKey) {
            if ($CanonicalCounts.ContainsKey($info.CanonicalKey)) {
                $CanonicalCounts[$info.CanonicalKey] += 1
            } else {
                $CanonicalCounts[$info.CanonicalKey] = 1
            }
        }
    }

    $FoundGhosts = @()
    foreach ($printer in $AllPrinters) {
        if ($BuiltinPrinters -contains $printer.Name) {
            continue
        }

        $printerInfo = Get-PrinterConnectionInfo -Name $printer.Name
        $isPatternGhost = $false
        foreach ($pattern in $GhostPatterns) {
            if ($printer.Name -like $pattern) {
                $isPatternGhost = $true
                break
            }
        }

        $isDuplicateConnection = $false
        if ($printerInfo -and $printerInfo.CanonicalKey -and $CanonicalCounts.ContainsKey($printerInfo.CanonicalKey)) {
            $isDuplicateConnection = $CanonicalCounts[$printerInfo.CanonicalKey] -gt 1
        }

        $isPreferredGpoName = $false
        if ($CompareGPO -and $PreferredGpoPrinterNames.Count -gt 0) {
            $isPreferredGpoName = $PreferredGpoPrinterNames.Contains($printer.Name)
        }

        $isGpoManagedPrinter = $false
        if ($CompareGPO -and $printerInfo -and $printerInfo.CanonicalKey) {
            $isGpoManagedPrinter = $AllowedGpoPrinterKeys.Contains($printerInfo.CanonicalKey)
        }

        $shouldRemove = $false
        if ($FullCleanup) {
            $shouldRemove = $true
        } elseif ($CompareGPO) {
            $shouldRemove = (-not $isGpoManagedPrinter) -or ($isDuplicateConnection -and -not $isPreferredGpoName)
        } else {
            $shouldRemove = $isPatternGhost -or $isDuplicateConnection
        }

        if ($shouldRemove) {
            $FoundGhosts += $printer
        }
    }

    $FoundGhosts = $FoundGhosts | Sort-Object Name -Unique

    if ($FoundGhosts.Count -gt 0) {
        foreach ($ghost in $FoundGhosts) {
            try {
                Remove-Printer -Name $ghost.Name -ErrorAction Stop
                Log "Removed printer via Remove-Printer: $($ghost.Name)" "Info"
                $CleanedGhosts += $ghost.Name
            }
            catch {
                Log "Failed to remove printer via Remove-Printer: $($ghost.Name) - $_" "Critical"
                $FailedGhosts += $ghost.Name
            }
        }
    }
    else {
        Log "No removable printer connections detected by Get-Printer." "Info"
    }

    $ActivePrinterNames = @($AllPrinters | Select-Object -ExpandProperty Name)
    Remove-StalePrintEnumDevices -ActivePrinterNames $ActivePrinterNames -BuiltInPrinterNames $BuiltinPrinters

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

# HKLM Printer Connection Cleanup
if ($IsAdmin) {
    Remove-HKLMPrinterConnections
}

# Write status for this phase
$phaseNum = if ($RetryOnly) { 2 } else { 1 }
Write-StatusJson -Printers $FailedPrinters -Ghosts $FailedGhosts -CleanedPrinters $CleanedPrinters -CleanedGhosts $CleanedGhosts -Phase $phaseNum

# Elevation Retry Phase
if (-not $NoSelfElevation -and -not $RetryOnly -and ($FailedPrinters -or $FailedGhosts)) {
    $scriptPath = $MyInvocation.MyCommand.Definition
    $printerArg = ($FailedPrinters -join '|')
    $ghostArg   = ($FailedGhosts   -join '|')

    $argsList   = @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File',           "`"$scriptPath`"",
        '-RetryOnly',
        '-RetryPrinters',  "`"$printerArg`"",
        '-RetryGhosts',    "`"$ghostArg`""
    )

    Log "Some cleanup failed. Relaunching with elevation to retry targeted steps..." "Warning"
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


# Cleanup Mounted Hive
if ($MountedHive) {
    try {
        reg unload $MountKey | Out-Null
        Log "Unmounted hive for $TargetUser" "Debug"
    } catch {
        Log "Failed to unload hive for ${TargetUser}: $_" "Warning"
    }
}

# Announce Finish
Log "Printer cleanup complete for user: $TargetUser" "Info"
