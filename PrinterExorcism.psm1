enum LogVerbosity {
    Info = 0
    Warning = 1
    Critical = 2
    Debug = 3
}

function Invoke-PrinterExorcism {
    <#
    .SYNOPSIS
        Launches the structured printer cleanup process.

    .DESCRIPTION
        This function serves as the canonical entry point for launching the two-phase printer exorcism process.
        It delegates to Start-PrinterExorcismSession, passing through cleanup options such as GPO comparison,
        automation mode, and verbosity level.

        If TargetUser is specified, cleanup is performed against that user's registry hive.
        Otherwise, the currently logged-in user's profile is used.

    .PARAMETER FullCleanup
        Performs a full printer cleanup, including registry entries and WMI-printer mappings.

    .PARAMETER CompareGPO
        Only remove printers not defined in GPO-deployed connections. Skips all others.

    .PARAMETER Automated
        Suppresses prompts and assumes consent for all actions.

    .PARAMETER TargetUser
        The username whose hive should be loaded and cleaned. If omitted, defaults to current user.

    .PARAMETER Verbosity
        Sets the logging level for console and file output. Options are: Info, Warning, Critical, Debug.

    .EXAMPLE
        Invoke-PrinterExorcism -FullCleanup -Automated -Verbosity Debug

        Performs a full cleanup for the current user without prompts, logging at maximum verbosity.
    #>
    [CmdletBinding()]
    param(
        [switch]$FullCleanup,
        [switch]$CompareGPO,
        [switch]$Automated,
        [string]$TargetUser,
        [ValidateSet("Info", "Warning", "Critical", "Debug")]
        [string]$Verbosity = "Critical"
    )

    enum LogVerbosity {
        Info = 0
        Warning = 1
        Critical = 2
        Debug = 3
    }

    # If no actionable flags are set, show help and exit
    if (-not $FullCleanup -and -not $CompareGPO -and -not $Automated -and -not $TargetUser) {
        Write-Host "`nPrinterExorcism Module - CLI Usage" -ForegroundColor Cyan
        Write-Host "----------------------------------------"
        Write-Host "Available Flags:"
        Write-Host "  -FullCleanup    : Perform full cleanup on current or target user"
        Write-Host "  -CompareGPO     : Show printers not defined by GPO policy"
        Write-Host "  -Automated      : Perform actions without prompting"
        Write-Host "  -TargetUser     : Offline user hive to operate against"
        Write-Host "  -Verbosity      : Set log verbosity (Info, Warning, Critical, Debug)"
        Write-Host ""
        Write-Host "Examples:"
        Write-Host "  Invoke-PrinterExorcism -FullCleanup"
        Write-Host "  Invoke-PrinterExorcism -TargetUser 'jsmith' -Automated"
        Write-Host ""
        return
    }

    $verbosityLevel = [LogVerbosity]::$Verbosity
    $params = @{
        FullCleanup = $FullCleanup
        CompareGPO  = $CompareGPO
        Automated   = $Automated
        TargetUser  = $TargetUser
        Verbosity   = $verbosityLevel
    }

    return Start-PrinterExorcismSession @params
}

Export-ModuleMember -Function Invoke-PrinterExorcism

function Make-PrintersSuffer {
    <#
    .SYNOPSIS
        Executes an overly dramatic and theatrical printer purge.

    .DESCRIPTION
        This function provides a humorous wrapper around Invoke-PrinterExorcism with 
        dramatic CLI output. Designed for internal use, demos, or situations when you 
        truly want to make the printers suffer — and want everyone to know it.

    .PARAMETER FullCleanup
        Performs a full printer cleanup, including registry entries and WMI-printer mappings.

    .PARAMETER CompareGPO
        Only remove printers not defined in GPO-deployed connections. Skips all others.

    .PARAMETER Automated
        Suppresses prompts and assumes consent for all actions.

    .PARAMETER TargetUser
        The username whose hive should be loaded and cleaned. If omitted, defaults to current user.

    .PARAMETER Verbosity
        Sets the logging level for console and file output. Options are: Info, Warning, Critical, Debug.

    .EXAMPLE
        Make-PrintersSuffer -FullCleanup -Automated -Verbosity Debug

        Purges all printers for the current user while narrating the chaos.
    #>
    [CmdletBinding()]
    param(
        [switch]$FullCleanup,
        [switch]$CompareGPO,
        [switch]$Automated,
        [string]$TargetUser,
        [ValidateSet("Info", "Warning", "Critical", "Debug")]
        [string]$Verbosity = "Critical"
    )

    # Default to dramatic full auto cleanup if no switches are supplied
    if (-not $PSBoundParameters.ContainsKey('FullCleanup') -and
        -not $PSBoundParameters.ContainsKey('CompareGPO') -and
        -not $PSBoundParameters.ContainsKey('Automated') -and
        -not $PSBoundParameters.ContainsKey('TargetUser')) {

        $FullCleanup = $true
        $Automated   = $true
        Write-Host "🎯 No flags? Defaulting to FULL AUTO-EXORCISM MODE." -ForegroundColor Red
        Start-Sleep -Milliseconds 600
    }

    Write-Host "`n🧪 Brewing holy water..."
    Start-Sleep -Milliseconds 500
    Write-Host "📜 Reading the sacred Driver Uninstallation Scrolls..."
    Start-Sleep -Milliseconds 500
    Write-Host "🗡  Sharpening registry-cleaving axe..."
    Start-Sleep -Milliseconds 700
    Write-Host "💾 Mounting the blessed USB of banishment..."
    Start-Sleep -Milliseconds 700
    Write-Host "`n🔥 Initiating exorcism protocol...`n" -ForegroundColor Yellow

    Invoke-PrinterExorcism -FullCleanup:$FullCleanup `
                           -CompareGPO:$CompareGPO `
                           -Automated:$Automated `
                           -TargetUser:$TargetUser `
                           -Verbosity:$Verbosity
}
Export-ModuleMember -Function Make-PrintersSuffer

function Show-PrinterDiscovery {
    <#
    .SYNOPSIS
        Displays a list of printers and printer-related registry entries for a specific user.

    .DESCRIPTION
        This function wraps the standalone Discover-Printers.ps1 script for use within the module.
        If -JSON is set, it emits structured JSON output suitable for logging or pipelines.
        No changes are made to the system — this is a safe, read-only discovery routine.

    .PARAMETER TargetUser
        Specifies the user whose registry hive should be mounted and scanned for printer artifacts.

    .PARAMETER JSON
        If set, outputs results as JSON instead of formatted terminal output.

    .EXAMPLE
        Show-PrinterDiscovery -TargetUser "jsmith" -JSON
    #>
    [CmdletBinding()]
    param (
        [string]$TargetUser,
        [switch]$JSON
    )

    $scriptPath = Join-Path $PSScriptRoot "Private\Discover-Printers.ps1"
    $args = @()
    if ($TargetUser) { $args += @("-TargetUser", $TargetUser) }
    if ($JSON)       { $args += "-JSON" }

    & $scriptPath @args
}

function Start-SystemWidePrinterCleanup {
    <#
    .SYNOPSIS
        Performs a full printer cleanup across all user profiles on the system.

    .DESCRIPTION
        This function iterates through each user profile in C:\Users, loading each user's registry hive 
        (if available), and performs a comprehensive printer cleanup using PrinterExorcist.ps1. 
        Built-in printers are preserved. Cleanup includes registry entries, WMI printer objects, 
        and ghost printers.

        After per-user cleanup, system-level (HKLM and spooler) printers are also purged, 
        except for Microsoft built-in printers.

        Must be run as Administrator. This is equivalent to `-SystemWide` mode.
    #>
    [CmdletBinding()]
    param (
        [switch]$Automated,
        [string]$Verbosity = "Critical"
    )
    . "$PSScriptRoot\Common.ps1"

    $Users = Get-ChildItem "C:\Users" -Directory |
        Where-Object { Test-Path "$($_.FullName)\NTUSER.DAT" }

    foreach ($user in $Users) {
        $Target = $user.Name
        Write-Host "`n🧼 Running cleanup for user: $Target" -ForegroundColor Cyan
        $script = Join-Path $PSScriptRoot "PrinterExorcist.ps1"
        $args = @(
            '-ExecutionPolicy', 'Bypass',
            '-File', "`"$script`"",
            '-TargetUser', "`"$Target`"",
            '-Automated',
            '-FullCleanup',
            '-Verbosity', $Verbosity
        )

        $process = Start-Process -FilePath powershell.exe `
                                 -ArgumentList $args `
                                 -NoNewWindow -Wait -PassThru
        if ($process.ExitCode -ne 0) {
            Write-Host "⚠️ Cleanup failed for $Target (exit $($process.ExitCode))" -ForegroundColor Red
        }
    }

    # Global cleanup
    Write-Host "`n🗑  Performing HKLM and WMI printer cleanup..." -ForegroundColor Cyan
    $null = & "$PSScriptRoot\PrinterExorcist.ps1" -Automated -FullCleanup -RetryOnly -Verbosity $Verbosity
}

Export-ModuleMember -Function Show-PrinterDiscovery, Start-SystemWidePrinterCleanup
