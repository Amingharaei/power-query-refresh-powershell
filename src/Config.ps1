<#
Config.ps1 -- load and validate config.psd1.

Import-PowerShellDataFile reads the settings file in PowerShell's restricted
language (data sections only, no code execution), so there is no third-party
dependency and nothing in the file can run. There are no secrets in the config:
the Outlook email backend uses your signed-in desktop Outlook.

Import-RefreshConfig returns a plain object (PSCustomObject) with validated
settings. It throws on any missing or invalid setting; the caller treats that as
a configuration error and exits with code 2.
#>

function Import-RefreshConfig {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Config file not found: $Path"
    }

    try {
        $data = Import-PowerShellDataFile -LiteralPath $Path
    }
    catch {
        throw "$Path is not a valid PowerShell data file: $($_.Exception.Message)"
    }

    # --- Paths ---------------------------------------------------------------
    $paths = if ($data.ContainsKey('Paths')) { $data.Paths } else { @{} }
    if (-not $paths.ContainsKey('ReportsDir') -or [string]::IsNullOrWhiteSpace([string]$paths.ReportsDir)) {
        throw "Missing 'ReportsDir' under Paths in $Path."
    }
    $reportsDir = [string]$paths.ReportsDir

    $logDir =
        if ($paths.ContainsKey('LogDir') -and -not [string]::IsNullOrWhiteSpace([string]$paths.LogDir)) {
            [string]$paths.LogDir
        }
        else {
            Join-Path (Split-Path -Parent $reportsDir) 'excel-refresh-logs'
        }

    # --- Discovery -----------------------------------------------------------
    $discovery = if ($data.ContainsKey('Discovery')) { $data.Discovery } else { @{} }
    $include = if ($discovery.ContainsKey('Include') -and $discovery.Include) { [string[]]$discovery.Include } else { @('*.xlsx', '*.xlsm') }
    $exclude = if ($discovery.ContainsKey('Exclude') -and $discovery.Exclude) { [string[]]$discovery.Exclude } else { @() }
    $recurse = if ($discovery.ContainsKey('Recurse')) { [bool]$discovery.Recurse } else { $false }

    # --- Refresh -------------------------------------------------------------
    $refresh = if ($data.ContainsKey('Refresh')) { $data.Refresh } else { @{} }
    $timeout = if ($refresh.ContainsKey('TimeoutSeconds')) { [int]$refresh.TimeoutSeconds } else { 1800 }
    if ($timeout -le 0) {
        throw "Refresh.TimeoutSeconds must be a positive number of seconds."
    }
    $retryOnce = if ($refresh.ContainsKey('RetryOnce')) { [bool]$refresh.RetryOnce } else { $true }

    # --- Email ---------------------------------------------------------------
    $emailTable = if ($data.ContainsKey('Email')) { $data.Email } else { @{} }
    $emailEnabled = if ($emailTable.ContainsKey('Enabled')) { [bool]$emailTable.Enabled } else { $false }
    $recipients = if ($emailTable.ContainsKey('Recipients') -and $emailTable.Recipients) { [string[]]$emailTable.Recipients } else { @() }
    $sendOn = if ($emailTable.ContainsKey('SendOn') -and $emailTable.SendOn) { [string]$emailTable.SendOn } else { 'always' }

    if ($emailEnabled -and $recipients.Count -eq 0) {
        throw "Email.Recipients must be non-empty when Email.Enabled is true."
    }
    if ($sendOn -notin @('always', 'failure')) {
        throw "Email.SendOn must be 'always' or 'failure' (got '$sendOn')."
    }

    [pscustomobject]@{
        ReportsDir     = $reportsDir
        LogDir         = $logDir
        Include        = $include
        Exclude        = $exclude
        Recurse        = $recurse
        TimeoutSeconds = $timeout
        RetryOnce      = $retryOnce
        Email          = [pscustomobject]@{
            Enabled    = $emailEnabled
            Recipients = $recipients
            SendOn     = $sendOn
        }
    }
}
