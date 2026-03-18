# =============================================================================
# Logging.psm1  -  Thread-safe logging for Etherium Command Center
# =============================================================================

$script:LogFile        = $null
$script:LogBuffer      = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
$script:DebugEnabled   = $false

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

    # Queue for any in-process consumers
    $script:LogBuffer.Enqueue($entry)

    # Always write to the log file - this is the reliable cross-runspace channel
    if ($script:LogFile) {
        try { Add-Content -Path $script:LogFile -Value $entry -ErrorAction Stop }
        catch { }
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

# ── Drain the in-process queue (kept for backwards compatibility) ─────────────
function Get-PendingLogEntries {
    param([int]$Max = 200)
    $results = [System.Collections.Generic.List[string]]::new()
    $item    = $null
    while ($results.Count -lt $Max -and $script:LogBuffer.TryDequeue([ref]$item)) {
        $results.Add($item)
    }
    return $results.ToArray()
}

Export-ModuleMember -Function Initialize-Logging, Get-LogFilePath, Set-DebugLoggingEnabled, Write-Log, Get-PendingLogEntries

