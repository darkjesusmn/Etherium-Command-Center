# =============================================================================
# Logging.psm1  -  Thread-safe logging for Etherium Command Center
# =============================================================================

$script:LogFile        = $null
$script:DebugEnabled   = $false

# ── File trim settings ────────────────────────────────────────────────────────
# Hard cap on how many lines the bot log file on disk is allowed to grow to.
# Every $script:LogTrimCheckEvery writes we check the file line count and trim
# it back to $script:LogMaxLines if it has exceeded the limit.
# This keeps bot_YYYYMMDD.log small forever regardless of session length.
$script:LogMaxLines       = 200   # maximum lines kept in the file
$script:LogTrimCheckEvery = 50    # check line count every N writes
$script:LogWriteCount     = 0
$script:LastFileWriteWarningAt = [datetime]::MinValue
$script:LastTrimWarningAt      = [datetime]::MinValue

function _WriteLoggingFallbackWarning {
    param(
        [string]$Message,
        [string]$Kind = 'General',
        [int]$ThrottleSeconds = 30
    )

    $now = Get-Date
    $lastAt = switch ($Kind) {
        'Write' { $script:LastFileWriteWarningAt }
        'Trim'  { $script:LastTrimWarningAt }
        default { [datetime]::MinValue }
    }

    if (($now - $lastAt).TotalSeconds -lt $ThrottleSeconds) { return }

    switch ($Kind) {
        'Write' { $script:LastFileWriteWarningAt = $now }
        'Trim'  { $script:LastTrimWarningAt = $now }
    }

    try {
        $entry = "[{0}][WARN][Logging] {1}" -f ($now.ToString('yyyy-MM-dd HH:mm:ss')), $Message
        Write-Host $entry -ForegroundColor Yellow
    } catch { }
}

# ── Initialise ────────────────────────────────────────────────────────────────
function Initialize-Logging {
    param([string]$LogDir = '.\Logs')
    if (-not (Test-Path $LogDir)) {
        New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
    }
    $stamp          = Get-Date -Format 'yyyyMMdd_HHmmss'
    $script:LogFile = Join-Path $LogDir "bot_$stamp.log"
    Write-Log "Logging initialised -> $($script:LogFile)" -Level INFO
}

# ── Return the current log file path (used by GUI to tail the file) ───────────
function Get-LogFilePath {
    return $script:LogFile
}

# ── Enable or disable DEBUG logging ──────────────────────────────────────────
function Set-DebugLoggingEnabled {
    param([bool]$Enabled)
    $script:DebugEnabled = [bool]$Enabled
}

# ── Write a log entry ─────────────────────────────────────────────────────────
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR','DEBUG')][string]$Level = 'INFO',
        [string]$Source = ''
    )

    if ($Level -eq 'DEBUG' -and -not $script:DebugEnabled) { return }
    $ts    = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $src   = if ($Source) { "[$Source] " } else { '' }
    $entry = "[$ts][$Level] $src$Message"

    # Always write to the log file - this is the reliable cross-runspace channel
    if ($script:LogFile) {
        try {
            Add-Content -Path $script:LogFile -Value $entry -ErrorAction Stop

            # Every $script:LogTrimCheckEvery writes, check if the file has grown
            # past $script:LogMaxLines and trim it back if so.
            # Doing this on every single write would be expensive; batching it
            # keeps the overhead near zero while still capping file size.
            $script:LogWriteCount++
            if ($script:LogWriteCount -ge $script:LogTrimCheckEvery) {
                $script:LogWriteCount = 0
                try {
                    $allLines = [System.IO.File]::ReadAllLines($script:LogFile)
                    if ($allLines.Count -gt $script:LogMaxLines) {
                        # Keep only the LAST $script:LogMaxLines lines
                        $keep = $allLines[($allLines.Count - $script:LogMaxLines)..($allLines.Count - 1)]
                        [System.IO.File]::WriteAllLines($script:LogFile, $keep)
                    }
                } catch {
                    _WriteLoggingFallbackWarning -Kind 'Trim' -Message "Failed to trim log file '$($script:LogFile)': $($_.Exception.Message)"
                }
            }
        }
        catch {
            _WriteLoggingFallbackWarning -Kind 'Write' -Message "Failed to write log entry to '$($script:LogFile)': $($_.Exception.Message)"
        }
    }

    # Echo to host console
    $colour = switch ($Level) {
        'ERROR' { 'Red'    }
        'WARN'  { 'Yellow' }
        'DEBUG' { 'Gray'   }
        default { 'White'  }
    }
    try { Write-Host $entry -ForegroundColor $colour } catch { }
}

Export-ModuleMember -Function Initialize-Logging, Get-LogFilePath, Set-DebugLoggingEnabled, Write-Log
