#requires -Version 5.1
<#
.SYNOPSIS
    Excel Refresh Orchestrator (PowerShell) -- refresh a folder of Excel Power Query
    reports on a schedule, with per-query failure attribution and logs you can build
    reports on.

.DESCRIPTION
    Reads config.psd1, refreshes every workbook in the reports folder (each in its
    own Excel instance, under a watchdog), writes the event log, optionally emails an
    Outlook summary, and exits 0 if all workbooks succeeded or 1 if any failed. Task
    Scheduler reads that exit code as "Last Run Result".

.PARAMETER ConfigPath
    Path to the settings file. Defaults to config.psd1 next to this script.

.PARAMETER ReportsDir
    Override the reports folder for a one-off run.

.PARAMETER NoStaRelaunch
    Skip the automatic re-launch into a single-threaded apartment. For advanced use;
    see the STA note below.

.EXAMPLE
    pwsh -STA -NoProfile -ExecutionPolicy Bypass -File .\Refresh-Reports.ps1

.NOTES
    Windows only. Needs Excel installed and a logged-on, awake session at run time.
#>
[CmdletBinding()]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot 'config.psd1'),
    [string]$ReportsDir,
    [switch]$NoStaRelaunch
)

# --- Windows-only guard -----------------------------------------------------
# ($IsWindows exists only on PowerShell 6+, hence the version short-circuit; on
# Windows PowerShell 5.1 the platform is always Windows.)
if ($PSVersionTable.PSVersion.Major -ge 6 -and -not $IsWindows) {
    Write-Error 'This tool drives the Excel desktop app over COM and runs on Windows only.'
    exit 3
}

# --- STA guard --------------------------------------------------------------
# Excel COM must run in a single-threaded apartment (STA). A normal Windows
# PowerShell 5.1 session and a normal pwsh session are STA already, but jobs,
# runspaces, and the VS Code integrated console are MTA, where Excel COM is
# extremely slow or fails outright. If we find ourselves in MTA, relaunch the
# same host in STA and hand back its exit code. The relaunched child is passed
# -NoStaRelaunch so it never loops.
if (-not $NoStaRelaunch -and
    [System.Threading.Thread]::CurrentThread.GetApartmentState() -ne [System.Threading.ApartmentState]::STA) {

    $hostExe = (Get-Process -Id $PID).Path
    if ([string]::IsNullOrEmpty($hostExe)) {
        $hostExe = if ($PSVersionTable.PSEdition -eq 'Core') { 'pwsh' } else { 'powershell' }
    }

    $argList = @('-STA', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath, '-NoStaRelaunch')
    if ($PSBoundParameters.ContainsKey('ConfigPath')) { $argList += @('-ConfigPath', $ConfigPath) }
    if ($PSBoundParameters.ContainsKey('ReportsDir')) { $argList += @('-ReportsDir', $ReportsDir) }

    & $hostExe @argList
    exit $LASTEXITCODE
}

# --- Load the building blocks (dot-sourced so their functions/classes are ours) ---
. (Join-Path $PSScriptRoot 'src\Config.ps1')
. (Join-Path $PSScriptRoot 'src\EventLog.ps1')
. (Join-Path $PSScriptRoot 'src\ExcelEngine.ps1')
. (Join-Path $PSScriptRoot 'src\Notify.ps1')


function Write-Console {
    <#
      Best-effort console feedback for whoever is watching the window. Kept separate
      from the logger on purpose: the logger always writes to excel_refresh.log, but
      this window may have no usable console (e.g. launched with -WindowStyle Hidden
      under Task Scheduler), where Write-Host can throw. Swallowing that here means a
      console-write failure can never affect the refresh itself or the exit code.
    #>
    param([string]$Message = '')
    try { Write-Host $Message } catch { }
}

function Invoke-Run {
    [CmdletBinding()]
    param([Parameter(Mandatory)] $Config)

    Write-Console 'Excel Refresh Orchestrator (PowerShell)'
    Write-Console 'The reports are being refreshed one at a time. This can take a while'
    Write-Console 'depending on how many reports there are and how large they are.'
    Write-Console ''

    $logger = New-RefreshLogger -LogDir $Config.LogDir
    $events = New-RefreshEventLog -LogDir $Config.LogDir
    $logger.Info("Run $($events.RunId) start. Reports folder: $($Config.ReportsDir)")
    Write-RefreshEvent -EventLog $events -Scope 'run' -Target $Config.ReportsDir -EventName 'RUN' -Status 'STARTED'

    # List[object] (not ArrayList): its .Add() returns void, so it cannot leak an
    # index value into this function's output and corrupt the returned exit code.
    $results = New-Object 'System.Collections.Generic.List[object]'

    try {
        $workbooks = @(Find-RefreshWorkbook -ReportsDir $Config.ReportsDir -Include $Config.Include -Exclude $Config.Exclude -Recurse $Config.Recurse)
        $logger.Info("Found $($workbooks.Count) workbook(s).")
        Write-Console "Found $($workbooks.Count) report(s) in $($Config.ReportsDir)"
        Write-Console ''

        $i = 0
        foreach ($wbFile in $workbooks) {
            $i++
            $logger.Info("Refreshing $($wbFile.Name)")
            Write-Console "[$i/$($workbooks.Count)] Refreshing $($wbFile.Name) ..."
            Write-RefreshEvent -EventLog $events -Scope 'workbook' -Target $wbFile.Name -EventName 'WORKBOOK' -Status 'STARTED'

            $result = Invoke-WorkbookRefresh -Path $wbFile.FullName -TimeoutSeconds $Config.TimeoutSeconds -EventLog $events -Logger $logger -RetryOnce $Config.RetryOnce
            $results.Add($result)

            $failCount = @($result.Failed()).Count
            $detail = if ($failCount -gt 0) { "$failCount query failure(s)" } else { '' }
            Write-RefreshEvent -EventLog $events -Scope 'workbook' -Target $result.Workbook -EventName 'WORKBOOK' -Status $result.Status -DurationSeconds $result.Duration -ErrorType $result.ErrorType -ErrorMessage $result.ErrorMessage -Detail $detail

            $note = if ($detail) { ": $detail" } else { '' }
            $logger.Info(('  {0} {1} ({2:F1}s){3}' -f $result.Status, $result.Workbook, $result.Duration, $note))
            Write-Console ('    {0,-7} ({1:F1}s){2}' -f $result.Status, $result.Duration, $note)
        }
    }
    finally {
        $anyFailed = @($results | Where-Object { $_.Status -eq 'FAILED' }).Count -gt 0
        $runStatus = if ($anyFailed) { 'FAILED' } else { 'SUCCESS' }
        $failedCount = @($results | Where-Object { $_.Status -eq 'FAILED' }).Count
        $logger.Info("Run done. $($results.Count) workbook(s), failures=$anyFailed")
        Write-RefreshEvent -EventLog $events -Scope 'run' -Target $Config.ReportsDir -EventName 'RUN' -Status $runStatus -Detail "$($results.Count) workbook(s)"
        $events.Close()
        Invoke-RefreshEmail -Config $Config -Results $results.ToArray() -RunId $events.RunId -Logger $logger | Out-Null
        $logger.Close()
    }

    $okCount = $results.Count - $failedCount
    Write-Console ''
    Write-Console "Done. $okCount/$($results.Count) report(s) refreshed successfully."
    if ($failedCount -gt 0) {
        Write-Console 'Check excel_refresh.log and refresh-events.csv in the log folder for details.'
    }

    if ($failedCount -gt 0) { return 1 } else { return 0 }
}


# --- Main -------------------------------------------------------------------
try {
    $config = Import-RefreshConfig -Path $ConfigPath
}
catch {
    try { Write-Error "Configuration error: $($_.Exception.Message)" } catch { }
    exit 2
}

if ($PSBoundParameters.ContainsKey('ReportsDir') -and -not [string]::IsNullOrWhiteSpace($ReportsDir)) {
    $config.ReportsDir = $ReportsDir
}

exit (Invoke-Run -Config $config)
