<#
EventLog.ps1 -- two logs, both written to the log folder.

  refresh-events.csv  -- one row per event (run start/end, and each workbook's
                         and query's result). Tabular and append-only, so you can
                         build reports on it over time in Power BI, Excel, or SQL.
                         Every row from one run shares a run_id.
  excel_refresh.log    -- a rotating, human-readable log for debugging a run.

Both writers keep AutoFlush on so a later crash cannot lose already-written rows,
and both write UTF-8 without a BOM and RFC-4180 CRLF line endings so the CSV loads
cleanly into any consumer.

The classes are consumed through the factory functions New-RefreshLogger and
New-RefreshEventLog. Callers use the returned objects dynamically; they never need
the class type names, which keeps this file's classes out of the caller's parse
scope.
#>

class RefreshLogger {
    hidden [System.IO.StreamWriter] $Writer
    hidden [string] $Path

    RefreshLogger([string]$logDir) {
        if (-not (Test-Path -LiteralPath $logDir)) {
            New-Item -ItemType Directory -Path $logDir -Force | Out-Null
        }
        $this.Path = Join-Path $logDir 'excel_refresh.log'
        $this.RotateIfLarge()
        $utf8 = New-Object System.Text.UTF8Encoding($false)          # $false => no BOM
        $this.Writer = [System.IO.StreamWriter]::new($this.Path, $true, $utf8)
        $this.Writer.NewLine = "`r`n"
        $this.Writer.AutoFlush = $true
    }

    # Size-based rotation, checked when the log is opened: keeps the file bounded
    # across runs (excel_refresh.log -> .log.1 -> ... -> .log.5, oldest dropped).
    hidden [void] RotateIfLarge() {
        [int]$maxBytes = 2000000
        [int]$backups = 5
        if ((Test-Path -LiteralPath $this.Path) -and ((Get-Item -LiteralPath $this.Path).Length -ge $maxBytes)) {
            $oldest = "$($this.Path).$backups"
            if (Test-Path -LiteralPath $oldest) { Remove-Item -LiteralPath $oldest -Force }
            for ($i = $backups - 1; $i -ge 1; $i--) {
                $src = "$($this.Path).$i"
                if (Test-Path -LiteralPath $src) {
                    Move-Item -LiteralPath $src -Destination "$($this.Path).$($i + 1)" -Force
                }
            }
            Move-Item -LiteralPath $this.Path -Destination "$($this.Path).1" -Force
        }
    }

    [void] Info([string]$message) { $this.WriteLine('INFO', $message) }
    [void] Error([string]$message) { $this.WriteLine('ERROR', $message) }

    hidden [void] WriteLine([string]$level, [string]$message) {
        $ts = [datetime]::Now.ToString('yyyy-MM-dd HH:mm:ss')
        $this.Writer.WriteLine(('{0} {1,-7} {2}' -f $ts, $level, $message))
    }

    [void] Close() {
        if ($null -ne $this.Writer) {
            $this.Writer.Flush()
            $this.Writer.Dispose()
            $this.Writer = $null
        }
    }
}

class RefreshEventLog {
    hidden [System.IO.StreamWriter] $Writer
    hidden [string] $Path
    [string] $RunId

    static [string[]] $Columns = @(
        'run_id', 'timestamp', 'scope', 'target', 'event', 'status',
        'duration_seconds', 'error_type', 'error_message', 'detail'
    )

    RefreshEventLog([string]$logDir) {
        if (-not (Test-Path -LiteralPath $logDir)) {
            New-Item -ItemType Directory -Path $logDir -Force | Out-Null
        }
        $this.RunId = [guid]::NewGuid().ToString('N').Substring(0, 12)
        $this.Path = Join-Path $logDir 'refresh-events.csv'
        $isNew = (-not (Test-Path -LiteralPath $this.Path)) -or ((Get-Item -LiteralPath $this.Path).Length -eq 0)

        $utf8 = New-Object System.Text.UTF8Encoding($false)          # $false => no BOM
        $this.Writer = [System.IO.StreamWriter]::new($this.Path, $true, $utf8)
        $this.Writer.NewLine = "`r`n"
        $this.Writer.AutoFlush = $true

        if ($isNew) {
            $header = ([RefreshEventLog]::Columns | ForEach-Object { [RefreshEventLog]::EscapeCsv($_) }) -join ','
            $this.Writer.WriteLine($header)
        }
    }

    # RFC 4180: quote a field only if it contains a quote, comma, or line break,
    # and double any embedded quotes. Matches Python csv.writer's QUOTE_MINIMAL.
    static [string] EscapeCsv([string]$field) {
        if ($null -eq $field) { return '' }
        if ($field -match '[",\r\n]') {
            return '"' + $field.Replace('"', '""') + '"'
        }
        return $field
    }

    [void] Emit(
        [string]$scope, [string]$target, [string]$eventName, [string]$status,
        [object]$durationSeconds, [string]$errorType, [string]$errorMessage, [string]$detail
    ) {
        $dur =
            if ($null -eq $durationSeconds -or "$durationSeconds" -eq '') { '' }
            else { ([double]$durationSeconds).ToString('F3', [System.Globalization.CultureInfo]::InvariantCulture) }

        $fields = @(
            $this.RunId,
            [System.DateTimeOffset]::Now.ToString("yyyy-MM-dd'T'HH:mm:ss.fffzzz"),
            $scope, $target, $eventName, $status, $dur, $errorType, $errorMessage, $detail
        )
        $line = ($fields | ForEach-Object { [RefreshEventLog]::EscapeCsv([string]$_) }) -join ','
        $this.Writer.WriteLine($line)
    }

    [void] Close() {
        if ($null -ne $this.Writer) {
            $this.Writer.Flush()
            $this.Writer.Dispose()
            $this.Writer = $null
        }
    }
}

function New-RefreshLogger {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$LogDir)
    return [RefreshLogger]::new($LogDir)
}

function New-RefreshEventLog {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$LogDir)
    return [RefreshEventLog]::new($LogDir)
}

# Thin wrapper so call sites can use named parameters and omit the optional ones
# (PowerShell class methods support neither). Delegates to RefreshEventLog.Emit.
function Write-RefreshEvent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $EventLog,
        [Parameter(Mandatory)][string]$Scope,
        [Parameter(Mandatory)][string]$Target,
        [Parameter(Mandatory)][string]$EventName,
        [Parameter(Mandatory)][string]$Status,
        [object]$DurationSeconds = $null,
        [string]$ErrorType = '',
        [string]$ErrorMessage = '',
        [string]$Detail = ''
    )
    $EventLog.Emit($Scope, $Target, $EventName, $Status, $DurationSeconds, $ErrorType, $ErrorMessage, $Detail)
}
