function Log-PrinterEvent {
    param (
        [string]$msg,
        [string]$Level = "Info",  # Info, Warning, Critical, Debug
        [string]$LogPath,
        [int]$Verbosity = 2       # default = Critical
    )

    $levelMap = @{
        "Info"     = 0
        "Warning"  = 1
        "Critical" = 2
        "Debug"    = 3
    }

    $msgLevel = $levelMap[$Level]
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $entry = "$timestamp [$Level] $msg"

    # Write to file if log level is Critical or lower (or full debug mode)
    if ($msgLevel -le 2 -or $Verbosity -ge 3) {
        $entry | Add-Content -Path $LogPath
    }

    # Write to terminal if verbosity allows
    if ($Verbosity -ge $msgLevel) {
        switch ($Level) {
            "Info"     { Write-Host $msg -ForegroundColor Gray }
            "Warning"  { Write-Warning $msg }
            "Critical" { Write-Host $msg -ForegroundColor Red }
            "Debug"    { Write-Host "[DEBUG] $msg" -ForegroundColor DarkGray }
        }
    }
}
