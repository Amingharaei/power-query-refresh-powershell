# Configuration for power-query-refresh-powershell.
#
# This is a PowerShell *data file* (.psd1): a restricted-language hashtable that
# Import-PowerShellDataFile reads WITHOUT executing any code, so it is safe and
# needs no third-party parser. It is the PowerShell-native counterpart to a JSON
# or TOML settings file.
#
# Edit the two paths under Paths before your first run. Use single-quoted strings
# for Windows paths so the backslashes are taken literally. There are no passwords
# or secrets in here: the optional email uses the Outlook desktop app you are
# already signed in to.

@{

    Paths = @{
        # The folder that holds the .xlsx / .xlsm reports you want refreshed.
        ReportsDir = 'D:\excel-reports'

        # The folder where logs are written (created automatically if missing).
        LogDir = 'D:\excel-refresh-logs'
    }

    Discovery = @{
        # Which files in the reports folder count as reports (wildcards).
        Include = @('*.xlsx', '*.xlsm')

        # Optional name patterns to skip, e.g. @('_archive*.xlsx', '*-template.xlsx').
        Exclude = @()

        # $false = only the reports folder itself. $true = include subfolders.
        Recurse = $false
    }

    Refresh = @{
        # Time limit PER WORKBOOK, in seconds. If a workbook runs longer than this,
        # its Excel instance is force-closed and the run moves on. 1800 = 30 minutes.
        TimeoutSeconds = 1800

        # Retry a query once if the first refresh throws. The first code-driven
        # refresh right after opening a file sometimes throws a spurious
        # "initialization of the data source failed"; a second attempt usually
        # succeeds, while a genuinely broken query fails again and is reported.
        RetryOnce = $true
    }

    Email = @{
        # Sends a run summary through your signed-in Outlook desktop app. No
        # password needed. Leave Enabled = $false to turn email off.
        Enabled    = $false
        Recipients = @('you@company.com')

        # 'always'  = email every run.
        # 'failure' = email only when at least one workbook fails.
        SendOn = 'always'
    }

}
