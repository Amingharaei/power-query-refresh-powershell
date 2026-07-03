<#
Notify.ps1 -- optional Outlook summary email.

Uses the Outlook desktop app you are already signed in to (Outlook.Application COM),
so no password or SMTP configuration is needed. Sending mail never changes the run's
result: if Outlook cannot send, that is logged and the run still reports its real
success/failure and exit code.
#>

function Build-RefreshSummary {
    [CmdletBinding()]
    param(
        [object[]]$Results = @(),
        [Parameter(Mandatory)][string]$RunId
    )

    $total = $Results.Count
    $failed = @($Results | Where-Object { $_.Status -eq 'FAILED' })
    $ok = $total - $failed.Count

    $subject = "[Excel refresh] $ok/$total ok"
    if ($failed.Count -gt 0) { $subject += ", $($failed.Count) FAILED" }
    $subject += " (run $RunId)"

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("Run $RunId")
    $lines.Add("Workbooks: $total   Succeeded: $ok   Failed: $($failed.Count)")
    $lines.Add('')

    foreach ($r in $Results) {
        $rFailed = @($r.Failed())
        $note = if ($rFailed.Count -gt 0) { '  (not saved)' } else { '' }
        $lines.Add(('{0,-8} {1}  ({2:F1}s){3}' -f $r.Status, $r.Workbook, $r.Duration, $note))
        foreach ($c in $rFailed) {
            $lines.Add("    - query '$($c.Name)': $($c.ErrorType): $($c.ErrorMessage)")
        }
        if ($r.Status -eq 'FAILED' -and $rFailed.Count -eq 0 -and $r.ErrorMessage) {
            $lines.Add("    - $($r.ErrorType): $($r.ErrorMessage)")
        }
    }

    return [pscustomobject]@{
        Subject = $subject
        Body    = ($lines -join "`n")
    }
}

function Send-RefreshOutlookMail {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]$Recipients,
        [Parameter(Mandatory)][string]$Subject,
        [Parameter(Mandatory)][string]$Body
    )

    $outlook = $null
    $mail = $null
    try {
        $outlook = New-Object -ComObject Outlook.Application
        $mail = $outlook.CreateItem(0)          # 0 = olMailItem
        $mail.To = ($Recipients -join '; ')
        $mail.Subject = $Subject
        $mail.Body = $Body
        $mail.Send()
    }
    finally {
        if ($null -ne $mail) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($mail) } catch { } }
        if ($null -ne $outlook) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) } catch { } }
    }
}

function Invoke-RefreshEmail {
    <#
      Decide whether to email (enabled, and SendOn matches the outcome), build the
      summary, and send it. Any failure here is logged and swallowed so it can
      never affect the run's exit code.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $Config,
        [object[]]$Results = @(),
        [Parameter(Mandatory)][string]$RunId,
        [Parameter(Mandatory)] $Logger
    )

    if (-not $Config.Email.Enabled) { return }

    $anyFailed = @($Results | Where-Object { $_.Status -eq 'FAILED' }).Count -gt 0
    if ($Config.Email.SendOn -eq 'failure' -and -not $anyFailed) { return }

    try {
        $summary = Build-RefreshSummary -Results $Results -RunId $RunId
        Send-RefreshOutlookMail -Recipients $Config.Email.Recipients -Subject $summary.Subject -Body $summary.Body
        $Logger.Info("Summary email sent to $($Config.Email.Recipients.Count) recipient(s).")
    }
    catch {
        $Logger.Error("Could not send summary email: $($_.Exception.Message)")
    }
}
