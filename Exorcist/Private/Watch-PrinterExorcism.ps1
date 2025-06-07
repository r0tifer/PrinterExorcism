function Start-PrinterExorcismSession {
<#
.SYNOPSIS
    Orchestrates the behavior and tracking of PrinterExorcist.ps1.
#>
    [CmdletBinding()]
    param(
        [switch]$FullCleanup,
        [switch]$CompareGPO,
        [switch]$Automated,
        [string]$RetryPrinters,
        [string]$RetryGhosts,
        [string]$TargetUser,
        [switch]$RetryOnly,
        [switch]$JSON,
        [int]$Verbosity = 2
    )

    enum LogVerbosity {
        Info = 0
        Warning = 1
        Critical = 2
        Debug = 3
    }

    # ───── Import shared log system ─────
    try {
        . "$PSScriptRoot\..\Common.ps1"
    } catch {
        Write-Host "Failed to import Common.ps1 logging module: $_" -ForegroundColor Red
        return 100
    }

    $ConsoleLevel = if ($Verbosity -eq 3) { [LogVerbosity]::Debug } else { [LogVerbosity]::Critical }

    function Log {
        param (
            [string]$Message,
            [LogVerbosity]$Level = [LogVerbosity]::Info
        )
        Log-PrinterEvent -msg $Message -Level $Level.ToString() -LogPath $LogPath -Verbosity $ConsoleLevel
    }

    function Output-PhaseSummary($phase, $status) {
        Write-Host ""
        Write-Host "📦 Phase $phase cleanup summary for: $($status.user)" -ForegroundColor Cyan
        Write-Host "   🖨  Printers cleaned:   $($status.cleaned_printers -join ', ')" -ForegroundColor Green
        Write-Host "   ❌ Printers failed:    $($status.failed_printers -join ', ')" -ForegroundColor Red
        Write-Host "   👻 Ghosts detected:    $($status.cleaned_ghosts -join ', ')" -ForegroundColor Green
        Write-Host "   ☠️  Ghosts failed:      $($status.failed_ghosts -join ', ')" -ForegroundColor Red
        Write-Host ""

        Log "📦 Phase $phase cleanup summary for user: $($status.user)" Info
        Log "🖨  Printers cleaned:   $($status.cleaned_printers -join ', ')" Info
        Log "❌ Printers failed:    $($status.failed_printers -join ', ')" Warning
        Log "👻 Ghosts detected:    $($status.cleaned_ghosts -join ', ')" Info
        Log "☠️  Ghosts failed:      $($status.failed_ghosts -join ', ')" Warning
    }

    function Invoke-PrinterDiscovery {
        param (
            [string]$TargetUser,
            [switch]$JSON
        )

        $DiscoveryScript = Join-Path $PSScriptRoot 'Discover-Printers.ps1'
        if (-not (Test-Path $DiscoveryScript)) {
            Write-Host "Cannot find discovery script at: $DiscoveryScript" -ForegroundColor Red
            return 101
        }

        $args = @()
        if ($TargetUser) { $args += @("-TargetUser", "`"$TargetUser`"") }
        if ($JSON)       { $args += "-JSON" }

        Write-Host "🔍 Running printer discovery..." -ForegroundColor Cyan
        & $DiscoveryScript @args
        return 0
    }

    # Begin path setup
    if ($Automated) {
        $StatusDir = Join-Path $env:LOCALAPPDATA 'PrinterExorcist'
        if (-not (Test-Path $StatusDir)) {
            New-Item -ItemType Directory -Path $StatusDir -Force | Out-Null
        }
    } else {
        $StatusDir = [Environment]::GetFolderPath('Desktop')
        if (-not (Test-Path $StatusDir)) {
            $StatusDir = Join-Path $env:USERPROFILE 'Documents'
        }
    }

    $StatusPath = Join-Path $StatusDir 'PrinterCleanup.status.json'
    $LogPath    = Join-Path $StatusDir 'PrinterCleanup.log'
    $PrinterScript = Join-Path $PSScriptRoot 'PrinterExorcist.ps1'

    if (-not (Test-Path $PrinterScript)) {
        Log "Cannot locate PrinterExorcist.ps1 at $PrinterScript" Critical
        return 99
    }

    function Build-ArgList {
        $args = @("-ExecutionPolicy", "Bypass", "-File", "`"$PrinterScript`"")
        if ($FullCleanup)    { $args += "-FullCleanup" }
        if ($CompareGPO)     { $args += "-CompareGPO" }
        if ($Automated)      { $args += "-Automated" }
        if ($RetryOnly)      { $args += "-RetryOnly" }
        if ($RetryPrinters)  { $args += @("-RetryPrinters", "`"$RetryPrinters`"") }
        if ($RetryGhosts)    { $args += @("-RetryGhosts", "`"$RetryGhosts`"") }
        if ($TargetUser)     { $args += @("-TargetUser", "`"$TargetUser`"") }
        return $args
    }

    function Wait-ForStatusFile {
        $maxWait = 90
        $elapsed = 0
        while (-not (Test-Path $StatusPath) -and $elapsed -lt $maxWait) {
            Start-Sleep -Seconds 2
            $elapsed += 2
        }
        return (Test-Path $StatusPath)
    }

    function Read-Status {
        try {
            return Get-Content -Raw -Path $StatusPath | ConvertFrom-Json
        } catch {
            Log "Failed to parse status file: $_" Critical
            return $null
        }
    }

    function Validate-StatusSchema {
        param ([psobject]$Status)

        $expectedFields = @{
            "user"             = "System.String"
            "elevated"         = "System.Boolean"
            "phase"            = "System.Int32"
            "failed_printers"  = "System.Object[]"
            "failed_ghosts"    = "System.Object[]"
            "cleaned_printers" = "System.Object[]"
            "cleaned_ghosts"   = "System.Object[]"
            "timestamp"        = "System.String"
        }

        foreach ($field in $expectedFields.Keys) {
            if (-not ($Status.PSObject.Properties.Name -contains $field)) {
                Log "Missing required field in status file: $field" Critical
                return $false
            }

            $actualType = $Status.$field.GetType().FullName
            $expectedType = $expectedFields[$field]

            if ($actualType -ne $expectedType) {
                Log "Field '$field' has invalid type. Expected: $expectedType, Got: $actualType" Critical
                return $false
            }
        }

        return $true
    }

    Log "Watcher initiated with user-mode run of PrinterExorcist.ps1..." Info

    # ─── Handle Discovery Mode ─────────────────────────────
    if ($PSBoundParameters.ContainsKey('JSON') -or $PSBoundParameters.ContainsKey('TargetUser')) {
        return Invoke-PrinterDiscovery -TargetUser:$TargetUser -JSON:$JSON
    }

    # Announce none discovery run
    Write-Host "Initiating printer exorcism sweep..." -ForegroundColor Yellow

    if (Test-Path $StatusPath) {
        Remove-Item $StatusPath -Force
        Log "Removed old status file." Debug
    }

    $argList = Build-ArgList
    Log "Launching Phase 1 with args: $($argList -join ' ')" Debug

    Write-Host "Starting Phase 1 (user mode)..." -ForegroundColor Cyan
    Start-Process powershell.exe -ArgumentList $argList -WindowStyle Hidden -Wait

    if (-not (Wait-ForStatusFile)) {
        Log "Phase 1 failed or timed out waiting for status file." Critical
        Write-Host "Phase 1 timeout. No status received." -ForegroundColor Red
        return 1
    }

    # Attempt to read in the json status file
    $status = Read-Status
    if (-not $status) { return 2 }

    # Ensure JSON matches expected schema
    if (-not (Validate-StatusSchema $status)) {
        Write-Host "Invalid structure in status JSON. Aborting." -ForegroundColor Red
        return 3
    }   

    Log "Phase $($status.phase) finished for $($status.user)" Info
    Output-PhaseSummary 1 $status

    $needsRetry = $status.failed_printers.Count -gt 0 -or $status.failed_ghosts.Count -gt 0

    if ($needsRetry -and $status.phase -eq 1) {
        if (Test-Path $StatusPath) {
            Remove-Item $StatusPath -Force
            Log "Cleared status file for Phase 2..." Debug
        }

        if (-not $TargetUser) {
            $TargetUser = $env:USERNAME
            Log "No -TargetUser provided; defaulting to current user: $TargetUser" Warning
        }

        $innerArgsList = @('-RetryOnly')
        if ($status.failed_printers.Count -gt 0) {
            $innerArgsList += '-RetryPrinters', "`"$($status.failed_printers -join '|')`""
        }
        if ($status.failed_ghosts.Count -gt 0) {
            $innerArgsList += '-RetryGhosts', "`"$($status.failed_ghosts -join '|')`""
        }
        if ($Automated)   { $innerArgsList += '-Automated'   }
        if ($FullCleanup) { $innerArgsList += '-FullCleanup' }
        if ($CompareGPO)  { $innerArgsList += '-CompareGPO'  }
        $innerArgsList += '-TargetUser', "`"$TargetUser`""

        $innerArgsString = $innerArgsList -join ' '
        $OrigLocalAppData = $env:LOCALAPPDATA

        Log "Phase 2 required - launching elevated (targeting user: $TargetUser)…" Info
        Log "Command: powershell.exe -NoProfile -ExecutionPolicy Bypass -Command & { `$env:LOCALAPPDATA = '$OrigLocalAppData'; & `"$PrinterScript`" $innerArgsString }" Debug

        Write-Host "Phase 2 required - launching elevated (targeting user: $TargetUser)…" -ForegroundColor Yellow
        $command = "& { `$env:LOCALAPPDATA = '$OrigLocalAppData'; & `"$PrinterScript`" $innerArgsString }"
        Start-Process -FilePath 'powershell.exe' `
            -ArgumentList '-NoProfile','-ExecutionPolicy','Bypass','-Command',$command `
            -Verb RunAs -WindowStyle Hidden `
            -PassThru -Wait | Out-Null

        if (-not (Wait-ForStatusFile)) {
            Log "Phase 2 timeout - status file never updated." Critical
            Write-Host "Phase 2 timeout - no updated status received." -ForegroundColor Red
            return 3
        }

        $status = Read-Status
        if (-not $status) { return 4 }
        Output-PhaseSummary 2 $status
    } elseif (-not $needsRetry) {
        Log "No elevation needed. Cleanup complete after Phase 1!" Info
        Write-Host "No elevation needed. Cleanup complete after Phase 1!" -ForegroundColor Green
    }

    Log "Printer exorcism complete. Final log available at: $LogPath" Info
    Write-Host "Printer exorcism complete. Final log available at: $LogPath" -ForegroundColor Cyan
    return 0
}
