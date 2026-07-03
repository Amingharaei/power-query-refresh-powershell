#requires -Version 5.1
<#
.SYNOPSIS
    Preflight check: verify this machine can actually run the refresh before you
    schedule it.

.DESCRIPTION
    Checks PowerShell version and edition, that we are on Windows, the apartment
    state, that config.psd1 parses and validates, that the reports and log folders
    exist, that Excel COM instantiates, and (if email is enabled) that Outlook COM
    instantiates. Prints an OK/FAIL line per check and exits 0 if everything passed
    or 1 if anything failed.

    This does not open or refresh any workbook; it only confirms the environment.
#>
[CmdletBinding()]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot 'config.psd1')
)

$allOk = $true

function Write-Check {
    param([string]$Label, [bool]$Good, [string]$Detail = '')
    $tag = if ($Good) { 'OK  ' } else { 'FAIL' }
    $suffix = if ($Detail) { "  ($Detail)" } else { '' }
    Write-Host ("[{0}] {1}{2}" -f $tag, $Label, $suffix)
}

Write-Host 'power-query-refresh-powershell -- environment check'
Write-Host ''

# PowerShell version and edition
$v = $PSVersionTable.PSVersion
$versionOk = ($v.Major -gt 5) -or ($v.Major -eq 5 -and $v.Minor -ge 1)
Write-Check 'PowerShell 5.1 or newer' $versionOk "$($PSVersionTable.PSEdition) $v"
if (-not $versionOk) { $allOk = $false }

# Windows
$isWin = ($PSVersionTable.PSVersion.Major -lt 6) -or $IsWindows
Write-Check 'Running on Windows' $isWin
if (-not $isWin) { $allOk = $false }

# Apartment state (informational: the launcher and entry script self-correct to STA)
$apt = [System.Threading.Thread]::CurrentThread.GetApartmentState()
Write-Check 'Single-threaded apartment (STA)' ($apt -eq [System.Threading.ApartmentState]::STA) `
    "$apt; run-refresh.cmd and Refresh-Reports.ps1 self-correct to STA if needed"

# Config
$config = $null
try {
    . (Join-Path $PSScriptRoot 'src\Config.ps1')
    $config = Import-RefreshConfig -Path $ConfigPath
    Write-Check 'config.psd1 parses and validates' $true $ConfigPath
}
catch {
    Write-Check 'config.psd1 parses and validates' $false $_.Exception.Message
    $allOk = $false
}

if ($null -ne $config) {
    $reportsOk = Test-Path -LiteralPath $config.ReportsDir -PathType Container
    Write-Check 'Reports folder exists' $reportsOk $config.ReportsDir
    if (-not $reportsOk) { $allOk = $false }

    if (-not (Test-Path -LiteralPath $config.LogDir)) {
        try { New-Item -ItemType Directory -Path $config.LogDir -Force | Out-Null } catch { }
    }
    $logOk = Test-Path -LiteralPath $config.LogDir -PathType Container
    Write-Check 'Log folder exists (created if missing)' $logOk $config.LogDir
    if (-not $logOk) { $allOk = $false }
}

# Excel COM: instantiate a throwaway invisible instance and quit it.
$excelOk = $false
$excelErr = ''
$app = $null
try {
    $app = New-Object -ComObject Excel.Application
    $app.Visible = $false
    $excelOk = $true
}
catch {
    $excelErr = $_.Exception.Message
}
finally {
    if ($null -ne $app) {
        try { $app.Quit() } catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) } catch { }
    }
    [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
}
Write-Check 'Excel COM (Excel.Application) available' $excelOk $excelErr
if (-not $excelOk) { $allOk = $false }

# Outlook COM: only relevant when email is enabled.
if ($null -ne $config -and $config.Email.Enabled) {
    $olOk = $false
    $olErr = ''
    $ol = $null
    try {
        $ol = New-Object -ComObject Outlook.Application
        $olOk = $true
    }
    catch {
        $olErr = $_.Exception.Message
    }
    finally {
        if ($null -ne $ol) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ol) } catch { } }
    }
    Write-Check 'Outlook COM available (email is enabled)' $olOk $olErr
    if (-not $olOk) { $allOk = $false }
}

Write-Host ''
if ($allOk) {
    Write-Host 'All checks passed. You can run run-refresh.cmd or schedule it.'
    exit 0
}
else {
    Write-Host 'One or more checks failed. Fix the FAIL items above before scheduling.'
    exit 1
}
