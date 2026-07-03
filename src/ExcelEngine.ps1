<#
ExcelEngine.ps1 -- the Excel side of the job: find the workbooks, and refresh one
workbook in an isolated Excel instance guarded by a watchdog.

Refresh strategy -- each connection (query) is refreshed INDIVIDUALLY, inside
try/catch, so a failure is caught and attributed to the specific query. This is
the only reliable way to detect failures: Excel's "Refresh All", driven from code,
does NOT raise or report an error when a query fails, so a broken query would pass
silently and the workbook would be saved on stale data. See the README for the
trade-off this implies for queries that depend on each other.

If any query in a workbook fails, the workbook is NOT saved -- it keeps its last
good version rather than being left half-updated -- and the failure is recorded.

A fresh Excel instance is used per workbook (so one bad file can't sink the batch),
and a watchdog force-closes Excel if a refresh runs past the timeout.

Why the watchdog is written in C# (Add-Type) rather than in PowerShell: while the
pipeline thread is blocked inside a COM Refresh() call, PowerShell cannot run a
timer, event handler, or job callback -- nothing on that thread runs until the COM
call returns. Only a genuinely independent OS thread can act during the block. The
compiled ExcelWatchdog uses a System.Threading.Timer, which fires on a thread-pool
thread, so it can terminate Excel by PID even while PowerShell is stuck waiting.
#>

if (-not ('PqRefresh.ExcelWatchdog' -as [type])) {
    Add-Type -Language CSharp -TypeDefinition @'
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;

namespace PqRefresh
{
    public static class NativeMethods
    {
        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        // Excel.Application.Hwnd -> owning process id. Returns 0 if the handle is
        // not valid (the caller then falls back to a process snapshot diff).
        public static int GetProcessIdFromHwnd(long hwnd)
        {
            uint pid = 0;
            GetWindowThreadProcessId(new IntPtr(hwnd), out pid);
            return (int)pid;
        }
    }

    // Kills a target process if it is still running when the timeout elapses.
    // Runs entirely on a thread-pool thread, so it works even while the calling
    // PowerShell thread is blocked inside a COM call.
    public sealed class ExcelWatchdog : IDisposable
    {
        private readonly int _pid;
        private readonly Timer _timer;
        private int _fired;                 // guards against a double fire
        public bool TimedOut { get; private set; }

        public ExcelWatchdog(int pid, int timeoutMs)
        {
            _pid = pid;
            _timer = new Timer(OnElapsed, null, timeoutMs, Timeout.Infinite);
        }

        private void OnElapsed(object state)
        {
            if (Interlocked.Exchange(ref _fired, 1) != 0) { return; }
            try
            {
                Process p = Process.GetProcessById(_pid);
                if (!p.HasExited)
                {
                    TimedOut = true;        // set before Kill so the caller sees it
                    p.Kill();
                }
            }
            catch { /* process already gone -- nothing to do */ }
        }

        public void Cancel()
        {
            try { _timer.Change(Timeout.Infinite, Timeout.Infinite); } catch { }
        }

        public void Dispose()
        {
            try { _timer.Dispose(); } catch { }
        }
    }
}
'@
}


class RefreshConnectionResult {
    [string] $Name
    [string] $Status                    # SUCCESS | FAILED | SKIPPED
    [double] $Duration = 0.0
    [string] $ErrorType = ''
    [string] $ErrorMessage = ''

    RefreshConnectionResult([string]$name, [string]$status) {
        $this.Name = $name
        $this.Status = $status
    }

    RefreshConnectionResult([string]$name, [string]$status, [double]$duration, [string]$errorType, [string]$errorMessage) {
        $this.Name = $name
        $this.Status = $status
        $this.Duration = $duration
        $this.ErrorType = $errorType
        $this.ErrorMessage = $errorMessage
    }
}

class RefreshWorkbookResult {
    [string] $Workbook
    [string] $Status = 'SUCCESS'        # SUCCESS | FAILED
    [double] $Duration = 0.0
    [System.Collections.Generic.List[RefreshConnectionResult]] $Connections
    [string] $ErrorType = ''            # set only for a whole-workbook failure
    [string] $ErrorMessage = ''

    RefreshWorkbookResult([string]$workbook) {
        $this.Workbook = $workbook
        $this.Connections = [System.Collections.Generic.List[RefreshConnectionResult]]::new()
    }

    [RefreshConnectionResult[]] Failed() {
        return @($this.Connections | Where-Object { $_.Status -eq 'FAILED' })
    }
}


function Limit-String {
    param([string]$Value, [int]$Max)
    if ($null -eq $Value) { return '' }
    if ($Value.Length -gt $Max) { return $Value.Substring(0, $Max) }
    return $Value
}

function Find-RefreshWorkbook {
    <#
      Report files in ReportsDir, sorted, skipping Excel's ~$ lock files and any
      Exclude patterns. Matching is by file name against the Include/Exclude
      wildcards, which gives the same result as a glob for patterns like
      '*.xlsx' or '_archive*.xlsx'.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ReportsDir,
        [string[]]$Include = @('*.xlsx', '*.xlsm'),
        [string[]]$Exclude = @(),
        [bool]$Recurse = $false
    )

    if (-not (Test-Path -LiteralPath $ReportsDir -PathType Container)) {
        throw "Reports directory not found: $ReportsDir"
    }

    $all = Get-ChildItem -LiteralPath $ReportsDir -File -Recurse:$Recurse -ErrorAction Stop

    $matched = foreach ($file in $all) {
        $name = $file.Name
        if ($name -like '~$*') { continue }                                   # Excel's open-file lock marker
        if (-not ($Include | Where-Object { $name -like $_ })) { continue }   # must match an include pattern
        if ($Exclude | Where-Object { $name -like $_ }) { continue }          # skip excludes
        $file
    }

    return @($matched | Sort-Object -Property FullName)
}

function Disable-ConnectionBackgroundQuery {
    <#
      Turn off background refresh so the .Refresh() call blocks until the query is
      done. Returns $true if this is a refreshable data connection (OLEDB/ODBC).

      XlConnectionType: 1 = xlConnectionTypeOLEDB (Power Query lives here),
      2 = xlConnectionTypeODBC, 7 = xlConnectionTypeMODEL (Data Model -- not a data
      source, skipped). Others are not row-pulling connections and are skipped too.
    #>
    param([Parameter(Mandatory)] $Connection)

    $XL_OLEDB = 1
    $XL_ODBC = 2

    $ctype = $Connection.Type
    if ($ctype -eq $XL_OLEDB) {
        $Connection.OLEDBConnection.BackgroundQuery = $false
        return $true
    }
    if ($ctype -eq $XL_ODBC) {
        $Connection.ODBCConnection.BackgroundQuery = $false
        return $true
    }
    return $false
}

function Invoke-ConnectionRefreshWithRetry {
    # Refresh one connection, optionally retrying once. Throws if it fails on the
    # final attempt; the caller catches that and attributes it to this query.
    param(
        [Parameter(Mandatory)] $Connection,
        [Parameter(Mandatory)] $Logger,
        [bool]$RetryOnce = $true
    )

    try {
        $Connection.Refresh()
    }
    catch {
        if (-not $RetryOnce) { throw }
        # The first code-driven refresh right after opening a file sometimes throws
        # a spurious 'initialization of the data source failed'. A second attempt
        # usually succeeds; a genuinely broken query fails again and propagates.
        $Logger.Info("Retrying connection '$($Connection.Name)' after: $($_.Exception.Message)")
        $Connection.Refresh()
    }
}

function Invoke-WorkbookRefresh {
    <#
      Open, refresh each query (capturing per-query outcome), save only if all
      succeeded, then close -- all in this workbook's own Excel instance, under a
      watchdog that force-closes Excel if the timeout is exceeded.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][int]$TimeoutSeconds,
        [Parameter(Mandatory)] $EventLog,
        [Parameter(Mandatory)] $Logger,
        [bool]$RetryOnce = $true
    )

    $ErrorActionPreference = 'Stop'     # make cmdlet errors terminating so try/catch sees them

    $result = [RefreshWorkbookResult]::new([System.IO.Path]::GetFileName($Path))
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    $app = $null
    $wb = $null
    $conn = $null
    $watchdog = $null
    $excelPid = 0

    try {
        # Snapshot existing Excel PIDs so we can identify the fresh instance if the
        # Hwnd -> PID lookup fails.
        $before = @(Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)

        $app = New-Object -ComObject Excel.Application
        $app.Visible = $false
        $app.DisplayAlerts = $false
        $app.ScreenUpdating = $false
        try { $app.AskToUpdateLinks = $false } catch { }

        # Identify this instance's process so the watchdog can kill exactly it.
        try { $excelPid = [PqRefresh.NativeMethods]::GetProcessIdFromHwnd([long]$app.Hwnd) } catch { $excelPid = 0 }
        if ($excelPid -le 0) {
            $after = @(Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
            $new = @($after | Where-Object { $_ -notin $before })
            if ($new.Count -eq 1) { $excelPid = $new[0] }
        }

        if ($excelPid -gt 0) {
            $watchdog = [PqRefresh.ExcelWatchdog]::new($excelPid, $TimeoutSeconds * 1000)
        }
        else {
            $Logger.Error("Could not determine the Excel PID for '$($result.Workbook)'; running this workbook without a watchdog.")
        }

        # Open read-write, do not update external links (belt and suspenders with
        # AskToUpdateLinks = $false above). Args: (Filename, UpdateLinks=0, ReadOnly=$false).
        $wb = $app.Workbooks.Open($Path, 0, $false)

        try {
            foreach ($conn in $wb.Connections) {
                $name = $conn.Name

                $refreshable = $false
                try { $refreshable = Disable-ConnectionBackgroundQuery -Connection $conn } catch { $refreshable = $false }

                if (-not $refreshable) {
                    $result.Connections.Add([RefreshConnectionResult]::new($name, 'SKIPPED'))
                    Write-RefreshEvent -EventLog $EventLog -Scope 'connection' -Target $name -EventName 'REFRESH' -Status 'SKIPPED' -Detail 'not an OLEDB/ODBC connection'
                    continue
                }

                Write-RefreshEvent -EventLog $EventLog -Scope 'connection' -Target $name -EventName 'REFRESH' -Status 'STARTED'
                $connStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
                try {
                    Invoke-ConnectionRefreshWithRetry -Connection $conn -Logger $Logger -RetryOnce $RetryOnce
                    $connStopwatch.Stop()
                    $dur = $connStopwatch.Elapsed.TotalSeconds
                    $result.Connections.Add([RefreshConnectionResult]::new($name, 'SUCCESS', $dur, '', ''))
                    Write-RefreshEvent -EventLog $EventLog -Scope 'connection' -Target $name -EventName 'REFRESH' -Status 'SUCCESS' -DurationSeconds $dur
                }
                catch {
                    $connStopwatch.Stop()
                    $dur = $connStopwatch.Elapsed.TotalSeconds
                    $etype = $_.Exception.GetType().Name
                    $emsg = Limit-String $_.Exception.Message 500
                    $result.Connections.Add([RefreshConnectionResult]::new($name, 'FAILED', $dur, $etype, $emsg))
                    Write-RefreshEvent -EventLog $EventLog -Scope 'connection' -Target $name -EventName 'REFRESH' -Status 'FAILED' -DurationSeconds $dur -ErrorType $etype -ErrorMessage $emsg
                    $Logger.Error("Query '$name' failed: $($_.Exception.Message)")
                }
            }

            # Save only if every query succeeded; otherwise keep the last good file.
            if (@($result.Failed()).Count -gt 0) {
                $result.Status = 'FAILED'
                $Logger.Error("$($result.Workbook): $(@($result.Failed()).Count) query(ies) failed; not saving.")
            }
            else {
                $app.Calculate()
                $wb.Save()
            }
        }
        finally {
            try { if ($null -ne $wb) { $wb.Close($false) } } catch { }   # do not mask the real error
        }
    }
    catch {
        # Whole-workbook failure: open failed, watchdog kill, Excel crash, save error.
        $result.Status = 'FAILED'
        if ($null -ne $watchdog -and $watchdog.TimedOut) {
            $result.ErrorType = 'TimeoutError'
            $result.ErrorMessage = "Refresh exceeded $TimeoutSeconds s; Excel was force-closed."
        }
        else {
            $result.ErrorType = $_.Exception.GetType().Name
            $result.ErrorMessage = Limit-String $_.Exception.Message 500
        }
        $Logger.Error("Workbook '$($result.Workbook)' failed: $($_.Exception.Message)")
    }
    finally {
        if ($null -ne $watchdog) { $watchdog.Cancel(); $watchdog.Dispose() }

        if ($null -ne $app) {
            try { $app.Quit() } catch { }
            try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) } catch { }
        }
        $conn = $null
        $wb = $null
        $app = $null

        # Encourage the runtime-callable wrappers to be finalized so EXCEL.EXE exits.
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()

        # Backstop: if the instance somehow survived Quit (rare COM edge cases),
        # make sure it does not linger and accumulate over scheduled runs.
        if ($excelPid -gt 0) {
            $still = Get-Process -Id $excelPid -ErrorAction SilentlyContinue
            if ($null -ne $still) {
                try { Stop-Process -Id $excelPid -Force -ErrorAction SilentlyContinue } catch { }
            }
        }
    }

    $stopwatch.Stop()
    $result.Duration = $stopwatch.Elapsed.TotalSeconds
    return $result
}
