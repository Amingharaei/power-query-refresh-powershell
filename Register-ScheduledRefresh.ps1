#requires -Version 5.1
<#
.SYNOPSIS
    Register the Excel refresh as a scheduled task, correctly configured for Excel
    automation (interactive logon, no admin required).

.DESCRIPTION
    Creates (or replaces) a scheduled task that runs run-refresh.cmd on a Daily or
    Weekly trigger. The task uses LogonType Interactive on purpose: Excel COM needs a
    real, logged-on desktop session. A task set to "run whether logged on or not" runs
    in a non-interactive session where Excel automation fails, which is the single most
    common reason a scheduled Excel refresh "works when I run it but not on schedule".

    This is a convenience wrapper over the Task Scheduler GUI steps in the README; use
    whichever you prefer.

.PARAMETER TaskName
    Name for the scheduled task. Default: PowerQueryRefresh.

.PARAMETER At
    Start time as HH:mm (24-hour). Default: 02:00.

.PARAMETER Frequency
    Daily or Weekly. Default: Daily.

.PARAMETER DaysOfWeek
    Days for a Weekly trigger. Default: Monday.

.PARAMETER ExecutionTimeLimitHours
    Task Scheduler's own kill-switch for the whole run, independent of the per-workbook
    timeout in config.psd1. Default: 4.

.EXAMPLE
    .\Register-ScheduledRefresh.ps1 -At 03:30

.EXAMPLE
    .\Register-ScheduledRefresh.ps1 -Frequency Weekly -DaysOfWeek Monday,Thursday -At 22:00
#>
[CmdletBinding()]
param(
    [string]$TaskName = 'PowerQueryRefresh',

    [ValidatePattern('^\d{2}:\d{2}$')]
    [string]$At = '02:00',

    [ValidateSet('Daily', 'Weekly')]
    [string]$Frequency = 'Daily',

    [ValidateSet('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday')]
    [string[]]$DaysOfWeek = @('Monday'),

    [int]$ExecutionTimeLimitHours = 4
)

if ($PSVersionTable.PSVersion.Major -ge 6 -and -not $IsWindows) {
    Write-Error 'Scheduling this tool is Windows-only.'
    exit 3
}

$cmd = Join-Path $PSScriptRoot 'run-refresh.cmd'
if (-not (Test-Path -LiteralPath $cmd)) {
    Write-Error "run-refresh.cmd not found next to this script ($cmd)."
    exit 1
}

$action = New-ScheduledTaskAction -Execute 'cmd.exe' -Argument "/c `"$cmd`"" -WorkingDirectory $PSScriptRoot

$trigger =
    if ($Frequency -eq 'Daily') {
        New-ScheduledTaskTrigger -Daily -At $At
    }
    else {
        New-ScheduledTaskTrigger -Weekly -DaysOfWeek $DaysOfWeek -At $At
    }

# LogonType Interactive is the critical setting: Excel automation needs a real,
# logged-on desktop. RunLevel Limited means no elevation is requested, so a normal
# user can register and run this without admin rights.
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Limited

$settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -ExecutionTimeLimit (New-TimeSpan -Hours $ExecutionTimeLimitHours)

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $action `
    -Trigger $trigger `
    -Principal $principal `
    -Settings $settings `
    -Description 'Refresh a folder of Excel Power Query reports on a schedule (power-query-refresh-powershell).' `
    -Force | Out-Null

Write-Host "Registered scheduled task '$TaskName' ($Frequency at $At, LogonType Interactive)."
Write-Host 'The PC must be on and this user logged on at that time for Excel automation to run.'
Write-Host "Review or change it in Task Scheduler, or remove it with:  Unregister-ScheduledTask -TaskName '$TaskName'"
