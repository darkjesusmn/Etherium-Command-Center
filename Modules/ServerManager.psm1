# =============================================================================
# ServerManager.psm1  -  Start / stop / monitor game server processes
# Compatible with PowerShell 5.1+
# =============================================================================



$script:State       = $null
$script:ModuleRoot  = $PSScriptRoot
$script:ErrorThrottle = [hashtable]::Synchronized(@{})

function Initialize-ServerManager {
    param([hashtable]$SharedState)
    $script:State = $SharedState
    if (-not $script:State.ContainsKey('RunningServers'))  { $script:State['RunningServers']  = [hashtable]::Synchronized(@{}) }
    if (-not $script:State.ContainsKey('RestartCounters')) { $script:State['RestartCounters'] = [hashtable]::Synchronized(@{}) }
    if (-not $script:State.ContainsKey('PendingAutoRestarts')) { $script:State['PendingAutoRestarts'] = [hashtable]::Synchronized(@{}) }
    if (-not $script:State.ContainsKey('PendingScheduledRestarts')) { $script:State['PendingScheduledRestarts'] = [hashtable]::Synchronized(@{}) }
    if (-not $script:State.ContainsKey('RestFailCounts'))  { $script:State['RestFailCounts']  = [hashtable]::Synchronized(@{}) }
    if (-not $script:State.ContainsKey('StdinHandles'))    { $script:State['StdinHandles']    = [hashtable]::Synchronized(@{}) }
    if (-not $script:State.ContainsKey('ShutdownFlags'))   { $script:State['ShutdownFlags']   = [hashtable]::Synchronized(@{}) }
    if (-not $script:State.ContainsKey('LastAutoSave'))    { $script:State['LastAutoSave']    = [hashtable]::Synchronized(@{}) }
    if (-not $script:State.ContainsKey('SentWarnings'))    { $script:State['SentWarnings']    = [hashtable]::Synchronized(@{}) }
    if (-not $script:State.ContainsKey('ServerRuntimeState')) { $script:State['ServerRuntimeState'] = [hashtable]::Synchronized(@{}) }
    if (-not $script:State.ContainsKey('PlayerActivityState')) { $script:State['PlayerActivityState'] = [hashtable]::Synchronized(@{}) }
    if (-not $script:State.ContainsKey('LatestPlayers'))   { $script:State['LatestPlayers']   = [hashtable]::Synchronized(@{}) }
    if (-not $script:State.ContainsKey('LatestPlayerCounts')) { $script:State['LatestPlayerCounts'] = [hashtable]::Synchronized(@{}) }
    if (-not $script:State.ContainsKey('LatestPlayerObservedAt')) { $script:State['LatestPlayerObservedAt'] = [hashtable]::Synchronized(@{}) }
    if (-not $script:State.ContainsKey('PlayerQueryState')) { $script:State['PlayerQueryState'] = [hashtable]::Synchronized(@{}) }
    if (-not $script:State.ContainsKey('PlayerObservationTrace')) { $script:State['PlayerObservationTrace'] = [hashtable]::Synchronized(@{}) }
    if (-not $script:State.ContainsKey('DiscordCommandContext')) { $script:State['DiscordCommandContext'] = [hashtable]::Synchronized(@{}) }
}

function _NormalizeGameIdentity {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return '' }
    return (($Name -replace '[^A-Za-z0-9]', '').ToLowerInvariant())
}

function _GetProfileKnownGame {
    param([hashtable]$Profile)

    if ($null -eq $Profile) { return '' }
    if (($Profile.Keys -contains 'KnownGame') -and -not [string]::IsNullOrWhiteSpace("$($Profile.KnownGame)")) {
        return "$($Profile.KnownGame)"
    }

    $prefix = if (($Profile.Keys -contains 'Prefix') -and $Profile.Prefix) { "$($Profile.Prefix)".ToUpperInvariant() } else { '' }
    switch ($prefix) {
        'PZ'  { return 'ProjectZomboid' }
        'HY'  { return 'Hytale' }
        'MC'  { return 'Minecraft' }
        'PW'  { return 'Palworld' }
        'DZ'  { return '7DaysToDie' }
        'VH'  { return 'Valheim' }
        'TR'  { return 'Terraria' }
        'RS'  { return 'Rust' }
        'ARK' { return 'Ark' }
        'SF'  { return 'Satisfactory' }
    }

    if (($Profile.Keys -contains 'GameName') -and $Profile.GameName) {
        return "$($Profile.GameName)"
    }

    return ''
}

function _TestProfileGame {
    param(
        [hashtable]$Profile,
        [string]$KnownGame
    )

    return ((_NormalizeGameIdentity (_GetProfileKnownGame -Profile $Profile)) -eq (_NormalizeGameIdentity $KnownGame))
}

function _CollectProfileStringValues {
    param(
        [object]$Value,
        [System.Collections.Generic.List[string]]$List,
        [System.Collections.Generic.HashSet[string]]$Seen
    )

    if ($null -eq $Value) { return }

    if ($Value -is [string]) {
        $candidate = $Value.Trim()
        if (-not [string]::IsNullOrWhiteSpace($candidate) -and $Seen.Add($candidate)) {
            $List.Add($candidate) | Out-Null
        }
        return
    }

    if ($Value -is [System.Collections.IDictionary]) {
        foreach ($key in @('value','Value','values','Values','items','Items')) {
            if ($Value.Contains($key)) {
                _CollectProfileStringValues -Value $Value[$key] -List $List -Seen $Seen
                return
            }
        }
        return
    }

    if ($Value -is [psobject]) {
        foreach ($propName in @('value','Value','values','Values','items','Items')) {
            $prop = $Value.PSObject.Properties[$propName]
            if ($prop) {
                _CollectProfileStringValues -Value $prop.Value -List $List -Seen $Seen
                return
            }
        }
    }

    if ($Value -is [System.Collections.IEnumerable]) {
        foreach ($item in $Value) {
            _CollectProfileStringValues -Value $item -List $List -Seen $Seen
        }
    }
}

function _GetProfileExeHints {
    param([hashtable]$Profile)

    $list = New-Object 'System.Collections.Generic.List[string]'
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

    if ($Profile -and ($Profile.Keys -contains 'ExeHints')) {
        _CollectProfileStringValues -Value $Profile.ExeHints -List $list -Seen $seen
    }

    if ($list.Count -eq 0 -and $Profile -and ($Profile.Keys -contains 'Executable')) {
        $exeName = [System.IO.Path]::GetFileName((_CoalesceStr $Profile.Executable ''))
        if (-not [string]::IsNullOrWhiteSpace($exeName) -and $seen.Add($exeName)) {
            $list.Add($exeName) | Out-Null
        }
    }

    return ,([string[]]$list.ToArray())
}

function _TestTextMatchesExecutableHint {
    param(
        [string]$Text,
        [string]$Hint
    )

    if ([string]::IsNullOrWhiteSpace($Text) -or [string]::IsNullOrWhiteSpace($Hint)) { return $false }

    if ($Hint.IndexOfAny([char[]]@('*','?')) -ge 0) {
        return ($Text -like "*$Hint*")
    }

    return ($Text.IndexOf($Hint, [System.StringComparison]::OrdinalIgnoreCase) -ge 0)
}

function _ClearPendingAutoRestart {
    param(
        [string]$Prefix,
        [hashtable]$SharedState = $script:State
    )

    if (-not $SharedState -or -not $SharedState.ContainsKey('PendingAutoRestarts')) { return }
    if ($SharedState.PendingAutoRestarts.ContainsKey($Prefix)) {
        $SharedState.PendingAutoRestarts.Remove($Prefix)
    }
}

function _ClearPendingScheduledRestart {
    param(
        [string]$Prefix,
        [hashtable]$SharedState = $script:State
    )

    if (-not $SharedState -or -not $SharedState.ContainsKey('PendingScheduledRestarts')) { return }
    if ($SharedState.PendingScheduledRestarts.ContainsKey($Prefix)) {
        $SharedState.PendingScheduledRestarts.Remove($Prefix)
    }
}

function _ClearServerSessionCaches {
    param(
        [string]$Prefix,
        [hashtable]$SharedState = $script:State
    )

    if ([string]::IsNullOrWhiteSpace($Prefix) -or -not $SharedState) { return }
    Initialize-ServerManager -SharedState $SharedState

    $key = $Prefix.ToUpperInvariant()
    $clearPalworldGlobals = $false
    foreach ($tableName in @(
        'LatestPlayers',
        'LatestPlayerCounts',
        'LatestPlayerObservedAt',
        'PlayerQueryState',
        'PlayerObservationTrace',
        'PlayersRequests',
        'PzObservedPlayerIds',
        'ServerStartNotified',
        'SatisfactoryConnectionCapture',
        'ValheimPlayerCapture',
        'RestJoinableNotified',
        'PalworldFeedState',
        'PalworldStateByPrefix',
        'PalworldPlayersByPrefix',
        'LastAutoSave',
        'RestFailCounts',
        'PlayerActivityState'
    )) {
        try {
            if ($SharedState.ContainsKey($tableName) -and $SharedState[$tableName] -and $SharedState[$tableName].ContainsKey($key)) {
                if ($tableName -in @('PalworldFeedState','PalworldStateByPrefix','PalworldPlayersByPrefix')) {
                    $clearPalworldGlobals = $true
                }
                $SharedState[$tableName].Remove($key) | Out-Null
            }
        } catch { }
    }

    if ($clearPalworldGlobals) {
        try {
            if ($SharedState.ContainsKey('PalworldState')) {
                $SharedState.Remove('PalworldState') | Out-Null
            }
        } catch { }
        try {
            if ($SharedState.ContainsKey('PalworldPlayers')) {
                $SharedState.Remove('PalworldPlayers') | Out-Null
            }
        } catch { }
    }

    try {
        if ($SharedState.ContainsKey('SentWarnings') -and $SharedState.SentWarnings) {
            foreach ($warnKey in @($SharedState.SentWarnings.Keys)) {
                if ("$warnKey" -like "${key}_*") {
                    $SharedState.SentWarnings.Remove($warnKey) | Out-Null
                }
            }
        }
    } catch { }

    try {
        if ($SharedState.ContainsKey('ShutdownFlags') -and $SharedState.ShutdownFlags) {
            $SharedState.ShutdownFlags[$key] = $false
        }
    } catch { }
}

function _SchedulePendingAutoRestart {
    param(
        [string]$Prefix,
        [object]$Profile,
        [hashtable]$SharedState,
        [int]$DelaySeconds,
        [int]$Attempt,
        [int]$MaxAttempts
    )

    Initialize-ServerManager -SharedState $SharedState

    $delay = [Math]::Max(0, $DelaySeconds)
    $dueAt = (Get-Date).AddSeconds($delay)
    $SharedState.PendingAutoRestarts[$Prefix] = @{
        DueAt       = $dueAt
        DelaySeconds= $delay
        Attempt     = $Attempt
        MaxAttempts = $MaxAttempts
    }

    Set-ServerRuntimeState -Prefix $Prefix -State 'waiting_restart' -Detail ("Crash recovery queued. Restart in {0}s (attempt {1}/{2})." -f $delay, $Attempt, $MaxAttempts) -SharedState $SharedState

    _Log "[$($Profile.GameName)] Auto-restart scheduled for $($dueAt.ToString('HH:mm:ss')) after ${delay}s delay (attempt $Attempt/$MaxAttempts)." -Level DEBUG
}

function _ProcessPendingAutoRestart {
    param(
        [string]$Prefix,
        [object]$Profile,
        [hashtable]$SharedState
    )

    if (-not $SharedState -or -not $SharedState.ContainsKey('PendingAutoRestarts')) { return $false }
    if (-not $SharedState.PendingAutoRestarts.ContainsKey($Prefix)) { return $false }

    $pending = $SharedState.PendingAutoRestarts[$Prefix]
    if ($null -eq $pending) {
        _ClearPendingAutoRestart -Prefix $Prefix -SharedState $SharedState
        return $false
    }

    if ($SharedState.RunningServers.ContainsKey($Prefix)) {
        _ClearPendingAutoRestart -Prefix $Prefix -SharedState $SharedState
        Set-ServerRuntimeState -Prefix $Prefix -State 'online' -Detail 'Server is running and being monitored.' -SharedState $SharedState
        return $true
    }

    $dueAt = $null
    try { $dueAt = [datetime]$pending.DueAt } catch { $dueAt = $null }
    if ($null -eq $dueAt) {
        $dueAt = Get-Date
    }

    if ((Get-Date) -lt $dueAt) {
        return $true
    }

    $attempt = 1
    $maxAttempts = 1
    try { $attempt = [int]$pending.Attempt } catch { }
    try { $maxAttempts = [int]$pending.MaxAttempts } catch { }

    Initialize-ServerManager -SharedState $SharedState
    $startResult = Start-GameServer -Prefix $Prefix -IsAutoRestart
    if ($startResult -eq $true -or $startResult -eq 'already_running') {
        _ClearPendingAutoRestart -Prefix $Prefix -SharedState $SharedState
        return $true
    }

    $retryDelaySeconds = 15
    $SharedState.PendingAutoRestarts[$Prefix] = @{
        DueAt       = (Get-Date).AddSeconds($retryDelaySeconds)
        DelaySeconds= $retryDelaySeconds
        Attempt     = $attempt
        MaxAttempts = $maxAttempts
    }
    Set-ServerRuntimeState -Prefix $Prefix -State 'waiting_restart' -Detail ("Restart attempt failed. Retrying in {0}s (attempt {1}/{2})." -f $retryDelaySeconds, $attempt, $maxAttempts) -SharedState $SharedState
    _Log "[$($Profile.GameName)] Auto-restart attempt $attempt/$maxAttempts failed to start. Retrying in ${retryDelaySeconds}s." -Level WARN
    return $true
}

function _SchedulePendingScheduledRestart {
    param(
        [string]$Prefix,
        [object]$Profile,
        [hashtable]$SharedState,
        [int]$DelaySeconds,
        [int]$Attempt,
        [int]$MaxAttempts
    )

    Initialize-ServerManager -SharedState $SharedState

    $delay = [Math]::Max(0, $DelaySeconds)
    $dueAt = (Get-Date).AddSeconds($delay)
    $SharedState.PendingScheduledRestarts[$Prefix] = @{
        DueAt       = $dueAt
        DelaySeconds= $delay
        Attempt     = $Attempt
        MaxAttempts = $MaxAttempts
    }

    Set-ServerRuntimeState -Prefix $Prefix -State 'waiting_restart' -Detail ("Scheduled restart retry in {0}s (attempt {1}/{2})." -f $delay, $Attempt, $MaxAttempts) -SharedState $SharedState
    $gameName = _GetNotificationGameName -Profile $Profile -Prefix $Prefix
    _Log "[$gameName] Scheduled restart recovery retry queued for $($dueAt.ToString('HH:mm:ss')) after ${delay}s delay (attempt $Attempt/$MaxAttempts)." -Level WARN
}

function _ProcessPendingScheduledRestart {
    param(
        [string]$Prefix,
        [object]$Profile,
        [hashtable]$SharedState
    )

    if (-not $SharedState -or -not $SharedState.ContainsKey('PendingScheduledRestarts')) { return $false }
    if (-not $SharedState.PendingScheduledRestarts.ContainsKey($Prefix)) { return $false }

    $pending = $SharedState.PendingScheduledRestarts[$Prefix]
    if ($null -eq $pending) {
        _ClearPendingScheduledRestart -Prefix $Prefix -SharedState $SharedState
        return $false
    }

    if ($SharedState.RunningServers.ContainsKey($Prefix)) {
        _ClearPendingScheduledRestart -Prefix $Prefix -SharedState $SharedState
        Set-ServerRuntimeState -Prefix $Prefix -State 'online' -Detail 'Server is running after scheduled restart.' -SharedState $SharedState
        return $true
    }

    $dueAt = $null
    try { $dueAt = [datetime]$pending.DueAt } catch { $dueAt = $null }
    if ($null -eq $dueAt) {
        $dueAt = Get-Date
    }

    if ((Get-Date) -lt $dueAt) {
        return $true
    }

    $attempt = 1
    $maxAttempts = 1
    try { $attempt = [int]$pending.Attempt } catch { }
    try { $maxAttempts = [int]$pending.MaxAttempts } catch { }

    Initialize-ServerManager -SharedState $SharedState
    $startResult = Start-GameServer -Prefix $Prefix
    if ($startResult -eq $true -or $startResult -eq 'already_running') {
        _ClearPendingScheduledRestart -Prefix $Prefix -SharedState $SharedState
        _SendDiscordLifecycleWebhook -Profile $Profile -Prefix $Prefix -Event 'scheduled_restart_done' -Tag 'ONLINE' -SharedState $SharedState | Out-Null
        return $true
    }

    if ($attempt -ge $maxAttempts) {
        _ClearPendingScheduledRestart -Prefix $Prefix -SharedState $SharedState
        Set-ServerRuntimeState -Prefix $Prefix -State 'failed' -Detail ("Scheduled restart failed after {0} recovery attempt(s)." -f $attempt) -SharedState $SharedState
        $gameName = _GetNotificationGameName -Profile $Profile -Prefix $Prefix
        _Log "[$gameName] Scheduled restart recovery exhausted after $attempt attempt(s)." -Level WARN
        _Webhook (New-DiscordGameMessage -Profile $Profile -Prefix $Prefix -Event 'scheduled_restart_failed' -Values @{
            Attempt = $attempt
        })
        return $true
    }

    $nextAttempt = $attempt + 1
    $retryDelaySeconds = 15
    _SchedulePendingScheduledRestart -Prefix $Prefix -Profile $Profile -SharedState $SharedState -DelaySeconds $retryDelaySeconds -Attempt $nextAttempt -MaxAttempts $maxAttempts
    _Webhook (New-DiscordGameMessage -Profile $Profile -Prefix $Prefix -Event 'scheduled_restart_retry' -Values @{
        DelaySeconds = $retryDelaySeconds
        Attempt = $nextAttempt
        MaxAttempts = $maxAttempts
    })
    return $true
}

# ---------------------------------------------------------------------------
#  Rebuild RunningServers from live processes (used at startup)
# ---------------------------------------------------------------------------
function Sync-RunningServersFromProcesses {
    param([hashtable]$SharedState)

    Initialize-ServerManager -SharedState $SharedState

    $count = 0
    $claimedPids = New-Object 'System.Collections.Generic.HashSet[int]'

    function _GetProcessMatchScore {
        param(
            [object]$Profile,
            [System.Diagnostics.Process]$Process
        )

        if ($null -eq $Profile -or $null -eq $Process) { return -1 }

        $score = 0
        $folderPath = ''
        $exeName = ''
        $exeHints = @()
        $procPath = ''
        $cmdLine = ''

        try { $folderPath = [Environment]::ExpandEnvironmentVariables((_CoalesceStr $Profile.FolderPath '')) } catch { $folderPath = (_CoalesceStr $Profile.FolderPath '') }
        $exeName = _CoalesceStr $Profile.Executable ''
        $exeHints = _GetProfileExeHints -Profile $Profile

        try { $procPath = "$($Process.Path)" } catch { $procPath = '' }
        try {
            $cim = Get-CimInstance Win32_Process -Filter "ProcessId = $($Process.Id)" -ErrorAction Stop
            if ($cim) {
                if (-not $procPath) { $procPath = "$($cim.ExecutablePath)" }
                $cmdLine = "$($cim.CommandLine)"
            }
        } catch { }

        $needleFolder = if ($folderPath) { $folderPath.TrimEnd('\') } else { '' }

        if ($procPath -and $needleFolder -and $procPath.StartsWith($needleFolder, [System.StringComparison]::OrdinalIgnoreCase)) {
            $score += 100
        }

        if ($cmdLine -and $needleFolder -and $cmdLine.IndexOf($needleFolder, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
            $score += 80
        }

        if ($exeName) {
            if ($procPath -and $procPath.IndexOf($exeName, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                $score += 70
            }
            if ($cmdLine -and $cmdLine.IndexOf($exeName, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                $score += 70
            }
        }

        if ($exeHints.Count -gt 0) {
            $matchedPathHint = $false
            $matchedCmdHint = $false
            foreach ($exeHint in @($exeHints)) {
                if (-not $matchedPathHint -and $procPath -and (_TestTextMatchesExecutableHint -Text $procPath -Hint $exeHint)) {
                    $score += 60
                    $matchedPathHint = $true
                }
                if (-not $matchedCmdHint -and $cmdLine -and (_TestTextMatchesExecutableHint -Text $cmdLine -Hint $exeHint)) {
                    $score += 60
                    $matchedCmdHint = $true
                }
                if ($matchedPathHint -and $matchedCmdHint) { break }
            }
        }

        if ($score -eq 0) {
            $genericNames = @('java','javaw','powershell','pwsh','cmd','conhost')
            if ($genericNames -contains $Process.ProcessName.ToLowerInvariant()) {
                return -1
            }
            return 1
        }

        return $score
    }

    function _ShouldPreserveRehydratedRuntimeState {
        param(
            [string]$PrefixKey
        )

        if ([string]::IsNullOrWhiteSpace($PrefixKey) -or -not $SharedState) { return $false }

        $runtimeCode = ''
        $hasObservedPlayers = $false
        $hasObservedAt = $false

        try {
            if ($SharedState.ContainsKey('ServerRuntimeState') -and $SharedState.ServerRuntimeState -and $SharedState.ServerRuntimeState.ContainsKey($PrefixKey)) {
                $runtime = $SharedState.ServerRuntimeState[$PrefixKey]
                if ($runtime) { $runtimeCode = [string]$runtime.Code }
            }
        } catch { $runtimeCode = '' }

        try {
            if ($SharedState.ContainsKey('LatestPlayerCounts') -and $SharedState.LatestPlayerCounts -and $SharedState.LatestPlayerCounts.ContainsKey($PrefixKey)) {
                $null = [int]$SharedState.LatestPlayerCounts[$PrefixKey]
                $hasObservedPlayers = $true
            }
        } catch { $hasObservedPlayers = $false }

        try {
            if ($SharedState.ContainsKey('LatestPlayerObservedAt') -and $SharedState.LatestPlayerObservedAt -and $SharedState.LatestPlayerObservedAt.ContainsKey($PrefixKey)) {
                $stamp = $SharedState.LatestPlayerObservedAt[$PrefixKey]
                $hasObservedAt = ($null -ne $stamp -and "$stamp".Trim().Length -gt 0)
            }
        } catch { $hasObservedAt = $false }

        $runtimeCode = $runtimeCode.ToLowerInvariant()
        if ($runtimeCode -in @('online','starting','restarting','stopping','waiting_first_player','idle_wait','idle_shutdown','waiting_restart')) {
            return $true
        }

        return ($hasObservedPlayers -or $hasObservedAt)
    }

    foreach ($pfx in @($SharedState.Profiles.Keys)) {
        $profile = $SharedState.Profiles[$pfx]
        if ($null -eq $profile) { continue }

        # If we already track a live PID, leave it alone
        if ($SharedState.RunningServers.ContainsKey($pfx)) {
            $entry = $SharedState.RunningServers[$pfx]
            $live  = try { Get-Process -Id $entry.Pid -ErrorAction Stop } catch { $null }
            if ($live -and -not $live.HasExited) { continue }
            _TraceServerRehydrationDecision -Prefix $pfx -Phase 'sync-running' -Action 'drop-stale-entry' -Reason 'tracked running-server entry no longer matched a live process during reconciliation' -Fields @{
                pid = $(try { $entry.Pid } catch { $null })
                start = $(try { $entry.StartTime } catch { $null })
            } -SharedState $SharedState
            $SharedState.RunningServers.Remove($pfx)
        }

        $procName = _CoalesceStr $profile.ProcessName ''
        if (-not $procName) { continue }

        $procs = @(Get-Process -Name $procName -ErrorAction SilentlyContinue)
        if (-not $procs) { continue }

        $candidates = foreach ($candidate in $procs) {
            if ($claimedPids.Contains([int]$candidate.Id)) { continue }

            $score = _GetProcessMatchScore -Profile $profile -Process $candidate
            if ($score -lt 0) { continue }

            [pscustomobject]@{
                Process = $candidate
                Score   = $score
                Start   = (try { $candidate.StartTime } catch { Get-Date 0 })
            }
        }

        if (-not $candidates) { continue }

        $proc = ($candidates | Sort-Object Score, Start -Descending | Select-Object -First 1).Process
        if ($null -eq $proc) { continue }
        [void]$claimedPids.Add([int]$proc.Id)

        $startTime = $null
        try { $startTime = $proc.StartTime } catch { $startTime = Get-Date }

        $srvLogPath = ''
        if ($null -ne $profile.ServerLogPath -and "$($profile.ServerLogPath)".Trim() -ne '') {
            $srvLogPath = "$($profile.ServerLogPath)".Trim()
        }

        $SharedState.RunningServers[$pfx] = @{
            Pid           = $proc.Id
            StartTime     = $startTime
            Process       = $proc
            ServerLogPath = $srvLogPath
        }
        $preserveRehydratedState = _ShouldPreserveRehydratedRuntimeState -PrefixKey $pfx
        if (-not $preserveRehydratedState) {
            _ResetPlayerActivityTracking -Prefix $pfx -SharedState $SharedState
        }
        $SharedState.ShutdownFlags[$pfx] = $false
        _ClearPendingScheduledRestart -Prefix $pfx -SharedState $SharedState
        if (-not $preserveRehydratedState) {
            Set-ServerRuntimeState -Prefix $pfx -State 'online' -Detail 'Detected already-running server process.' -SharedState $SharedState
        }
        _TraceServerRehydrationDecision -Prefix $pfx -Phase 'sync-running' -Action $(if ($preserveRehydratedState) { 'reattach-preserve' } else { 'reattach-reset' }) -Reason $(if ($preserveRehydratedState) { 'reconciled a live already-running process and preserved prior runtime/player state' } else { 'reconciled a live already-running process and reset runtime state to generic online' }) -Fields @{
            pid = $proc.Id
            process = $proc.ProcessName
            start = $startTime
            preserve = $preserveRehydratedState
        } -SharedState $SharedState
        $count++

        if ($preserveRehydratedState) {
            _Log "[$($profile.GameName)] Reattached to running process '$($proc.ProcessName)' (PID $($proc.Id)) and preserved runtime/player state."
        } else {
            _Log "[$($profile.GameName)] Detected running process '$($proc.ProcessName)' (PID $($proc.Id))"
        }
    }

    return $count
}

function Restore-ServerRuntimeStateFromSharedState {
    param([hashtable]$SharedState = $script:State)

    if (-not $SharedState) { return 0 }
    Initialize-ServerManager -SharedState $SharedState

    $restored = 0
    $now = Get-Date

    foreach ($prefix in @($SharedState.PendingAutoRestarts.Keys)) {
        if ([string]::IsNullOrWhiteSpace($prefix)) { continue }
        if ($SharedState.RunningServers.ContainsKey($prefix)) { continue }

        $pending = $null
        $dueAt = $null
        $attempt = 1
        $maxAttempts = 1
        try { $pending = $SharedState.PendingAutoRestarts[$prefix] } catch { $pending = $null }
        if ($null -eq $pending) { continue }
        try { $dueAt = [datetime]$pending.DueAt } catch { $dueAt = $null }
        try { $attempt = [int]$pending.Attempt } catch { $attempt = 1 }
        try { $maxAttempts = [int]$pending.MaxAttempts } catch { $maxAttempts = 1 }

        $detail = if ($dueAt) {
            "Crash recovery queued. Restart in {0} (attempt {1}/{2})." -f (_FormatCountdownShort -DueAt $dueAt), $attempt, $maxAttempts
        } else {
            "Crash recovery queued. Restart pending (attempt {0}/{1})." -f $attempt, $maxAttempts
        }

        Set-ServerRuntimeState -Prefix $prefix -State 'waiting_restart' -Detail $detail -SharedState $SharedState -PreserveSince
        _TraceServerRehydrationDecision -Prefix $prefix -Phase 'restore-runtime' -Action 'restore-pending-auto-restart' -Reason 'rehydrated queued crash-recovery restart from shared state' -Fields @{
            due = $dueAt
            attempt = $attempt
            maxAttempts = $maxAttempts
        } -SharedState $SharedState
        $restored++
    }

    foreach ($prefix in @($SharedState.PendingScheduledRestarts.Keys)) {
        if ([string]::IsNullOrWhiteSpace($prefix)) { continue }
        if ($SharedState.RunningServers.ContainsKey($prefix)) { continue }

        $pending = $null
        $dueAt = $null
        $attempt = 1
        $maxAttempts = 1
        try { $pending = $SharedState.PendingScheduledRestarts[$prefix] } catch { $pending = $null }
        if ($null -eq $pending) { continue }
        try { $dueAt = [datetime]$pending.DueAt } catch { $dueAt = $null }
        try { $attempt = [int]$pending.Attempt } catch { $attempt = 1 }
        try { $maxAttempts = [int]$pending.MaxAttempts } catch { $maxAttempts = 1 }

        $detail = if ($dueAt) {
            "Scheduled restart retry in {0} (attempt {1}/{2})." -f (_FormatCountdownShort -DueAt $dueAt), $attempt, $maxAttempts
        } else {
            "Scheduled restart retry pending (attempt {0}/{1})." -f $attempt, $maxAttempts
        }

        Set-ServerRuntimeState -Prefix $prefix -State 'waiting_restart' -Detail $detail -SharedState $SharedState -PreserveSince
        _TraceServerRehydrationDecision -Prefix $prefix -Phase 'restore-runtime' -Action 'restore-pending-scheduled-restart' -Reason 'rehydrated queued scheduled-restart retry from shared state' -Fields @{
            due = $dueAt
            attempt = $attempt
            maxAttempts = $maxAttempts
        } -SharedState $SharedState
        $restored++
    }

    foreach ($prefix in @($SharedState.RunningServers.Keys)) {
        if ([string]::IsNullOrWhiteSpace($prefix)) { continue }

        $profile = $null
        try {
            if ($SharedState.Profiles -and $SharedState.Profiles.ContainsKey($prefix)) {
                $profile = $SharedState.Profiles[$prefix]
            }
        } catch { $profile = $null }
        if ($null -eq $profile) { continue }

        $runtimeCode = ''
        try {
            $runtime = Get-ServerRuntimeState -Prefix $prefix -SharedState $SharedState
            if ($runtime) { $runtimeCode = [string]$runtime.Code }
        } catch { $runtimeCode = '' }
        $runtimeCode = $runtimeCode.ToLowerInvariant()
        if ($runtimeCode -in @('stopping','restarting','blocked','failed','startup_failed','idle_shutdown')) { continue }

        $activity = $null
        try { $activity = _UpdatePlayerActivityState -Prefix $prefix -Profile $profile -SharedState $SharedState } catch { $activity = $null }
        if ($null -eq $activity) { continue }
        if (-not [bool]$activity.DetectionSupported) { continue }

        $currentCount = 0
        try { $currentCount = [Math]::Max(0, [int]$activity.CurrentCount) } catch { $currentCount = 0 }
        $dueAt = $null
        try { if ($activity.ShutdownDueAt) { $dueAt = [datetime]$activity.ShutdownDueAt } } catch { $dueAt = $null }
        $pendingRule = ''
        try { $pendingRule = [string]$activity.PendingRule } catch { $pendingRule = '' }
        $pendingRule = $pendingRule.ToLowerInvariant()

        if ([bool]$activity.DetectionAvailable) {
            if ($currentCount -gt 0) {
                Set-ServerRuntimeState -Prefix $prefix -State 'online' -Detail ("{0} player(s) online." -f $currentCount) -SharedState $SharedState -PreserveSince
                _TraceServerRehydrationDecision -Prefix $prefix -Phase 'restore-runtime' -Action 'restore-online' -Reason 'rehydrated live online state from trusted player data' -Fields @{
                    source = $activity.DetectionSource
                    count = $currentCount
                } -SharedState $SharedState
                $restored++
                continue
            }

            if ($pendingRule -eq 'first_player' -and $dueAt) {
                Set-ServerRuntimeState -Prefix $prefix -State 'waiting_first_player' -Detail ("Waiting for first player. Idle shutdown in {0}." -f (_FormatCountdownShort -DueAt $dueAt)) -SharedState $SharedState -PreserveSince
                _TraceServerRehydrationDecision -Prefix $prefix -Phase 'restore-runtime' -Action 'restore-waiting-first-player' -Reason 'rehydrated first-player idle window from shared timer state' -Fields @{
                    source = $activity.DetectionSource
                    due = $dueAt
                } -SharedState $SharedState
                $restored++
                continue
            }

            if ($pendingRule -eq 'empty_after_leave' -and $dueAt) {
                Set-ServerRuntimeState -Prefix $prefix -State 'idle_wait' -Detail ("Server empty. Idle shutdown in {0}." -f (_FormatCountdownShort -DueAt $dueAt)) -SharedState $SharedState -PreserveSince
                _TraceServerRehydrationDecision -Prefix $prefix -Phase 'restore-runtime' -Action 'restore-idle-wait' -Reason 'rehydrated empty-server idle window from shared timer state' -Fields @{
                    source = $activity.DetectionSource
                    due = $dueAt
                } -SharedState $SharedState
                $restored++
                continue
            }

            continue
        }

        if ($dueAt) {
            if ($pendingRule -eq 'signal_wait') {
                $rehydratedState = 'waiting_first_player'
                $rehydratedDetail = _FormatSignalWaitRuntimeDetail -StateCode 'waiting_first_player' -DueAt $dueAt

                $hadSeenPlayers = $false
                try { $hadSeenPlayers = ($null -ne $activity.FirstPlayerSeenAt) } catch { $hadSeenPlayers = $false }
                $becameEmpty = $false
                try { $becameEmpty = ($null -ne $activity.LastBecameEmptyAt) } catch { $becameEmpty = $false }
                if ($hadSeenPlayers -and $becameEmpty) {
                    $rehydratedState = 'idle_wait'
                    $rehydratedDetail = _FormatSignalWaitRuntimeDetail -StateCode 'idle_wait' -DueAt $dueAt
                }

                Set-ServerRuntimeState -Prefix $prefix -State $rehydratedState -Detail $rehydratedDetail -SharedState $SharedState -PreserveSince
                _TraceServerRehydrationDecision -Prefix $prefix -Phase 'restore-runtime' -Action 'restore-signal-wait' -Reason 'rehydrated waiting-for-player-signal state from shared timer state' -Fields @{
                    due = $dueAt
                    hadSeenPlayers = $hadSeenPlayers
                    becameEmpty = $becameEmpty
                    state = $rehydratedState
                } -SharedState $SharedState
                $restored++
            }
        }
    }

    return $restored
}

# -----------------------------------------------------------------------------
#  INTERNAL HELPERS
# -----------------------------------------------------------------------------

function _Log {
    param([string]$Msg, [string]$Level = 'INFO')
    if ($Level -eq 'DEBUG') {
        $dbg = $false
        if ($script:State -and $script:State.Settings -and $script:State.Settings.ContainsKey('EnableDebugLogging')) {
            $dbg = [bool]$script:State.Settings.EnableDebugLogging
        }
        if (-not $dbg) { return }
    }
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry = "[$ts][$Level][ServerManager] $Msg"
    if ($script:State -and $script:State.ContainsKey('LogQueue')) {
        $script:State.LogQueue.Enqueue($entry)
    }
    $col = switch ($Level) { 'ERROR'{'Red'} 'WARN'{'Yellow'} default{'Cyan'} }
    try { Write-Host $entry -ForegroundColor $col } catch {}
}

function _LogThrottled {
    param(
        [string]$Key,
        [string]$Msg,
        [string]$Level = 'WARN',
        [int]$WindowSeconds = 30
    )

    if ([string]::IsNullOrWhiteSpace($Key)) {
        _Log $Msg -Level $Level
        return
    }

    $shouldLog = $true
    $now = Get-Date
    try {
        if ($script:ErrorThrottle.ContainsKey($Key)) {
            $last = $script:ErrorThrottle[$Key]
            if ($last -is [datetime] -and (($now - $last).TotalSeconds -lt $WindowSeconds)) {
                $shouldLog = $false
            }
        }
        if ($shouldLog) {
            $script:ErrorThrottle[$Key] = $now
            _Log $Msg -Level $Level
        }
    } catch {
        _Log $Msg -Level $Level
    }
}

function _ResolveServerRuntimeBadge {
    param([string]$State)

    $normalized = if ([string]::IsNullOrWhiteSpace($State)) { '' } else { $State.Trim().ToLowerInvariant() }
    switch ($normalized) {
        'online'          { return 'ONLINE' }
        'starting'        { return 'STARTING' }
        'stopping'        { return 'STOPPING' }
        'restarting'      { return 'RESTART' }
        'waiting_restart' { return 'WAITING' }
        'waiting_first_player' { return 'WAITING' }
        'blocked'         { return 'BLOCKED' }
        'failed'          { return 'FAILED' }
        'startup_failed'  { return 'FAILED' }
        'stopped'         { return 'STOPPED' }
        'idle_wait'       { return 'IDLE' }
        'idle_shutdown'   { return 'IDLE' }
        default           { return 'OFFLINE' }
    }
}

function _ResolveServerRuntimeLabel {
    param(
        [string]$State,
        [bool]$Running = $false
    )

    $normalized = if ([string]::IsNullOrWhiteSpace($State)) { '' } else { $State.Trim().ToLowerInvariant() }
    switch ($normalized) {
        'online'              { return 'Online' }
        'starting'            { return 'Starting' }
        'stopping'            { return 'Stopping' }
        'restarting'          { return 'Restarting' }
        'waiting_restart'     { return 'Waiting to restart' }
        'waiting_first_player'{ return 'Waiting for first player' }
        'blocked'             { return 'Blocked' }
        'failed'              { return 'Failed' }
        'startup_failed'      { return 'Startup failed' }
        'stopped'             { return 'Stopped' }
        'idle_wait'           { return 'Idle wait' }
        'idle_shutdown'       { return 'Idle shutdown' }
        default               { if ($Running) { return 'Online' } else { return 'Offline' } }
    }
}

function _FormatRelativeAgeShort {
    param(
        [object]$When,
        [string]$EmptyText = 'Never'
    )

    if ($null -eq $When) { return $EmptyText }

    try {
        $stamp = [datetime]$When
    } catch {
        return $EmptyText
    }

    $delta = (Get-Date) - $stamp
    if ($delta.TotalSeconds -lt 0) {
        $delta = [timespan]::Zero
    }

    if ($delta.TotalSeconds -lt 10) {
        return 'just now'
    }

    if ($delta.TotalSeconds -lt 60) {
        return ('{0}s ago' -f [Math]::Max(1, [int][Math]::Floor($delta.TotalSeconds)))
    }

    if ($delta.TotalMinutes -lt 60) {
        return ('{0}m ago' -f [Math]::Max(1, [int][Math]::Floor($delta.TotalMinutes)))
    }

    if ($delta.TotalHours -lt 24) {
        $hours = [Math]::Max(1, [int][Math]::Floor($delta.TotalHours))
        $minutes = [Math]::Max(0, [int][Math]::Floor($delta.TotalMinutes % 60))
        if ($minutes -gt 0) {
            return ('{0}h {1}m ago' -f $hours, $minutes)
        }
        return ('{0}h ago' -f $hours)
    }

    $days = [Math]::Max(1, [int][Math]::Floor($delta.TotalDays))
    $remainingHours = [Math]::Max(0, [int][Math]::Floor($delta.TotalHours % 24))
    if ($remainingHours -gt 0) {
        return ('{0}d {1}h ago' -f $days, $remainingHours)
    }
    return ('{0}d ago' -f $days)
}

function _ResolvePlayerObservationSourceLabel {
    param(
        [string]$Source,
        [bool]$Supported = $false,
        [bool]$Available = $false
    )

    $normalized = if ([string]::IsNullOrWhiteSpace($Source)) { '' } else { $Source.Trim().ToLowerInvariant() }
    switch ($normalized) {
        'projectzomboidlog' { return 'PZ player capture' }
        '7daystodietelnet'  { return '7DTD telnet' }
        'hytalewholog'      { return 'Hytale who/log' }
        'palworldrest'      { return 'Palworld REST' }
        'satisfactoryapi'   { return 'Satisfactory API' }
        'satisfactorylog'   { return 'Satisfactory log fallback' }
        'valheimlog'        { return 'Valheim log capture' }
        default {
            if (-not $Supported) { return 'Conservative fallback' }
            if (-not $Available) { return 'Waiting for trusted data' }
            return 'Trusted player data'
        }
    }
}

function Get-ProfileHealthSnapshot {
    param(
        [string]$Prefix,
        [hashtable]$Profile,
        [hashtable]$SharedState = $script:State,
        [bool]$Running = $false
    )

    $result = [ordered]@{
        HealthCode         = 'healthy'
        HealthText         = if ($Running) { 'Healthy' } else { 'Ready' }
        Summary            = if ($Running) { 'Server is running cleanly.' } else { 'Profile is ready.' }
        PlayerSource       = 'Conservative fallback'
        PlayerSourceCode   = ''
        DetectionSupported = $false
        DetectionAvailable = $false
        SourceNote         = ''
        LastEventLabel     = if ($Running) { 'Online' } else { 'Offline' }
        LastEventText      = if ($Running) { 'Server is running.' } else { 'Server is offline.' }
        LastEventAt        = $null
        LastEventAge       = 'unknown'
        LastPlayerSeen     = 'Never'
        LastPlayerSeenAt   = $null
        ConservativeMode   = $false
        ConservativeNote   = ''
        CurrentCount       = 0
    }

    if ([string]::IsNullOrWhiteSpace($Prefix) -or -not $SharedState) {
        return $result
    }

    Initialize-ServerManager -SharedState $SharedState

    $key = ''
    try { $key = $Prefix.ToUpperInvariant() } catch { $key = $Prefix }

    $runtime = $null
    $runtimeCode = ''
    $runtimeDetail = ''
    $runtimeSince = $null
    try {
        $runtime = Get-ServerRuntimeState -Prefix $key -SharedState $SharedState
        if ($runtime) {
            try { $runtimeCode = [string]$runtime.Code } catch { $runtimeCode = '' }
            try { $runtimeDetail = [string]$runtime.Detail } catch { $runtimeDetail = '' }
            try { $runtimeSince = [datetime]$runtime.Since } catch { $runtimeSince = $null }
        }
    } catch { }

    $activity = $null
    if ($SharedState.ContainsKey('PlayerActivityState') -and $SharedState.PlayerActivityState) {
        try {
            if ($SharedState.PlayerActivityState.ContainsKey($key)) {
                $activity = $SharedState.PlayerActivityState[$key]
            }
        } catch { $activity = $null }
    }

    $detectionSupported = $false
    $detectionAvailable = $false
    $detectionSource = ''
    $sourceNote = ''
    $currentCount = 0
    $lastObservedAt = $null
    $lastNonZeroAt = $null

    if ($activity) {
        try { $detectionSupported = [bool]$activity.DetectionSupported } catch { $detectionSupported = $false }
        try { $detectionAvailable = [bool]$activity.DetectionAvailable } catch { $detectionAvailable = $false }
        try { $detectionSource = [string]$activity.DetectionSource } catch { $detectionSource = '' }
        try { $sourceNote = [string]$activity.Note } catch { $sourceNote = '' }
        try { $currentCount = [Math]::Max(0, [int]$activity.CurrentCount) } catch { $currentCount = 0 }
        try { $lastObservedAt = [datetime]$activity.LastObservedAt } catch { $lastObservedAt = $null }
        try { $lastNonZeroAt = [datetime]$activity.LastNonZeroAt } catch { $lastNonZeroAt = $null }
    }

    $result.PlayerSource = _ResolvePlayerObservationSourceLabel -Source $detectionSource -Supported:$detectionSupported -Available:$detectionAvailable
    $result.PlayerSourceCode = $detectionSource
    $result.DetectionSupported = $detectionSupported
    $result.DetectionAvailable = $detectionAvailable
    $result.SourceNote = $sourceNote
    $result.CurrentCount = $currentCount

    $lastPlayerSeenAt = $null
    if ($currentCount -gt 0 -and $lastObservedAt) {
        $lastPlayerSeenAt = $lastObservedAt
        $result.LastPlayerSeen = 'Active now'
    } elseif ($lastNonZeroAt) {
        $lastPlayerSeenAt = $lastNonZeroAt
        $result.LastPlayerSeen = (_FormatRelativeAgeShort -When $lastNonZeroAt -EmptyText 'Never')
    }
    $result.LastPlayerSeenAt = $lastPlayerSeenAt

    $eventLabel = _ResolveServerRuntimeLabel -State $runtimeCode -Running:$Running
    $eventText = if (-not [string]::IsNullOrWhiteSpace($runtimeDetail)) { $runtimeDetail } else { $eventLabel }
    $result.LastEventLabel = $eventLabel
    $result.LastEventText = $eventText
    $result.LastEventAt = $runtimeSince
    $result.LastEventAge = (_FormatRelativeAgeShort -When $runtimeSince -EmptyText 'unknown')

    $normalizedRuntimeCode = if ([string]::IsNullOrWhiteSpace($runtimeCode)) { '' } else { $runtimeCode.Trim().ToLowerInvariant() }
    $conservativeMode = ($Running -and (-not $detectionSupported))
    $result.ConservativeMode = $conservativeMode
    if ($conservativeMode) {
        $result.ConservativeNote = if (-not [string]::IsNullOrWhiteSpace($sourceNote)) {
            $sourceNote
        } else {
            'Trusted player detection is not available yet. Idle shutdown stays conservative.'
        }
    }

    switch ($normalizedRuntimeCode) {
        'failed' {
            $result.HealthCode = 'error'
            $result.HealthText = 'Error'
            $result.Summary = if (-not [string]::IsNullOrWhiteSpace($runtimeDetail)) { $runtimeDetail } else { 'Server is in a failed state.' }
            return $result
        }
        'startup_failed' {
            $result.HealthCode = 'error'
            $result.HealthText = 'Error'
            $result.Summary = if (-not [string]::IsNullOrWhiteSpace($runtimeDetail)) { $runtimeDetail } else { 'Startup failed before the server became ready.' }
            return $result
        }
        'blocked' {
            $result.HealthCode = 'error'
            $result.HealthText = 'Error'
            $result.Summary = if (-not [string]::IsNullOrWhiteSpace($runtimeDetail)) { $runtimeDetail } else { 'ECC blocked this launch until the issue is resolved.' }
            return $result
        }
    }

    if ($Running -and $detectionSupported -and (-not $detectionAvailable)) {
        $result.HealthCode = 'waiting'
        $result.HealthText = 'Waiting for data'
        $result.Summary = if (-not [string]::IsNullOrWhiteSpace($sourceNote)) {
            $sourceNote
        } else {
            'Waiting for trusted player data.'
        }
        return $result
    }

    if ($normalizedRuntimeCode -in @('waiting_restart','idle_shutdown','restarting','stopping')) {
        $result.HealthCode = 'warning'
        $result.HealthText = 'Warning'
        $result.Summary = if (-not [string]::IsNullOrWhiteSpace($runtimeDetail)) {
            $runtimeDetail
        } else {
            'Server needs attention before it returns to a steady state.'
        }
        return $result
    }

    if ($Running -and -not $detectionSupported) {
        $result.HealthCode = 'warning'
        $result.HealthText = 'Warning'
        $result.Summary = if (-not [string]::IsNullOrWhiteSpace($sourceNote)) {
            $sourceNote
        } else {
            'Trusted player detection is not available yet. Idle shutdown stays conservative.'
        }
        return $result
    }

    if ($normalizedRuntimeCode -eq 'starting') {
        $result.HealthCode = 'waiting'
        $result.HealthText = 'Waiting for data'
        $result.Summary = if (-not [string]::IsNullOrWhiteSpace($runtimeDetail)) {
            $runtimeDetail
        } else {
            'Server startup is still in progress.'
        }
        return $result
    }

    if ($Running -and $detectionSupported -and $detectionAvailable) {
        $result.Summary = if ($currentCount -gt 0) {
            'Trusted player data is live.'
        } else {
            'Trusted player data is current and the server is empty.'
        }
    } elseif ($Running) {
        $result.HealthText = 'Healthy'
        $result.Summary = 'Server is running cleanly.'
    } else {
        $result.HealthText = 'Ready'
        $result.Summary = 'Profile is ready.'
    }

    return $result
}

function _GetRuntimeTransitionGameName {
    param(
        [string]$Prefix,
        [hashtable]$SharedState = $script:State
    )

    $key = if ([string]::IsNullOrWhiteSpace($Prefix)) { '' } else { $Prefix.ToUpperInvariant() }
    if ([string]::IsNullOrWhiteSpace($key)) { return '' }

    try {
        if ($SharedState -and $SharedState.ContainsKey('Profiles') -and $SharedState.Profiles) {
            if ($SharedState.Profiles.ContainsKey($key)) {
                $profile = $SharedState.Profiles[$key]
                if ($profile -and -not [string]::IsNullOrWhiteSpace([string]$profile.GameName)) {
                    return [string]$profile.GameName
                }
            } elseif ($SharedState.Profiles.ContainsKey($Prefix)) {
                $profile = $SharedState.Profiles[$Prefix]
                if ($profile -and -not [string]::IsNullOrWhiteSpace([string]$profile.GameName)) {
                    return [string]$profile.GameName
                }
            }
        }
    } catch { }

    return $key
}

function _FormatRuntimeTransitionText {
    param(
        [object]$Value,
        [int]$MaxLength = 140
    )

    $text = if ($null -eq $Value) { '' } else { [string]$Value }
    if ([string]::IsNullOrWhiteSpace($text)) { return '""' }

    $text = $text -replace "(`r`n|`n|`r)", ' '
    $text = $text -replace '\s+', ' '
    $text = $text.Trim()
    if ($text.Length -gt $MaxLength) {
        $text = $text.Substring(0, $MaxLength) + '...'
    }
    $text = $text.Replace('"', "'")
    return '"' + $text + '"'
}

function _NormalizeRuntimeTransitionToken {
    param(
        [object]$Value,
        [string]$Default = '<none>'
    )

    $text = if ($null -eq $Value) { '' } else { [string]$Value }
    if ([string]::IsNullOrWhiteSpace($text)) { return $Default }
    return $text.Trim()
}

function _TraceServerRuntimeStateTransition {
    param(
        [string]$Prefix,
        [object]$OldEntry,
        [object]$NewEntry,
        [hashtable]$SharedState = $script:State,
        [string]$Action = 'set',
        [switch]$PreserveSince
    )

    $key = if ([string]::IsNullOrWhiteSpace($Prefix)) { '' } else { $Prefix.ToUpperInvariant() }
    if ([string]::IsNullOrWhiteSpace($key)) { return }

    $oldCode   = _NormalizeRuntimeTransitionToken $(try { $OldEntry.Code } catch { $null })
    $oldBadge  = _NormalizeRuntimeTransitionToken $(try { $OldEntry.Badge } catch { $null })
    $oldLevel  = _NormalizeRuntimeTransitionToken $(try { $OldEntry.Level } catch { $null })
    $oldDetail = $(try { $OldEntry.Detail } catch { $null })

    $newCode   = _NormalizeRuntimeTransitionToken $(try { $NewEntry.Code } catch { $null })
    $newBadge  = _NormalizeRuntimeTransitionToken $(try { $NewEntry.Badge } catch { $null })
    $newLevel  = _NormalizeRuntimeTransitionToken $(try { $NewEntry.Level } catch { $null })
    $newDetail = $(try { $NewEntry.Detail } catch { $null })

    if ($Action -eq 'set') {
        $oldSince = $(try { [datetime]$OldEntry.Since } catch { $null })
        $newSince = $(try { [datetime]$NewEntry.Since } catch { $null })
        $noStateChange =
            ($oldCode -eq $newCode) -and
            ($oldBadge -eq $newBadge) -and
            ($oldLevel -eq $newLevel) -and
            ([string]$oldDetail -eq [string]$newDetail)
        $sameSince = ($null -eq $oldSince -and $null -eq $newSince) -or ($oldSince -eq $newSince)
        if ($noStateChange -and $sameSince) { return }
    } elseif ($Action -eq 'clear' -and $null -eq $OldEntry) {
        return
    }

    $changed = New-Object System.Collections.Generic.List[string]
    if ($oldCode -ne $newCode) { $changed.Add('code') | Out-Null }
    if ($oldBadge -ne $newBadge) { $changed.Add('badge') | Out-Null }
    if ($oldLevel -ne $newLevel) { $changed.Add('level') | Out-Null }
    if ([string]$oldDetail -ne [string]$newDetail) { $changed.Add('detail') | Out-Null }
    if ($PreserveSince) { $changed.Add('preserve-since') | Out-Null }

    $changeSummary = if ($changed.Count -gt 0) { $changed -join ',' } else { 'none' }
    $gameName = _GetRuntimeTransitionGameName -Prefix $key -SharedState $SharedState

    $message =
        'STATE transition ' +
        'prefix=' + $key + ' ' +
        'game=' + (_FormatRuntimeTransitionText -Value $gameName) + ' ' +
        'action=' + (_NormalizeRuntimeTransitionToken -Value $Action -Default 'set') + ' ' +
        'changed=' + $changeSummary + ' ' +
        'from.code=' + $oldCode + ' ' +
        'to.code=' + $newCode + ' ' +
        'from.badge=' + $oldBadge + ' ' +
        'to.badge=' + $newBadge + ' ' +
        'from.level=' + $oldLevel + ' ' +
        'to.level=' + $newLevel + ' ' +
        'from.detail=' + (_FormatRuntimeTransitionText -Value $oldDetail) + ' ' +
        'to.detail=' + (_FormatRuntimeTransitionText -Value $newDetail)

    _Log $message -Level DEBUG
}

function _FormatTimerTransitionTime {
    param(
        [object]$Value,
        [string]$Default = '<none>'
    )

    if ($null -eq $Value) { return $Default }
    try {
        return ([datetime]$Value).ToString('o')
    } catch {
        return $Default
    }
}

function _TraceServerTimerDecision {
    param(
        [string]$Prefix,
        [string]$Timer,
        [object]$OldRule,
        [object]$OldDueAt,
        [object]$NewRule,
        [object]$NewDueAt,
        [string]$Reason = '',
        [hashtable]$SharedState = $script:State,
        [string]$Action = ''
    )

    $key = if ([string]::IsNullOrWhiteSpace($Prefix)) { '' } else { $Prefix.ToUpperInvariant() }
    if ([string]::IsNullOrWhiteSpace($key)) { return }

    $timerName = _NormalizeRuntimeTransitionToken -Value $Timer -Default 'unknown'
    $oldRuleToken = _NormalizeRuntimeTransitionToken -Value $OldRule
    $newRuleToken = _NormalizeRuntimeTransitionToken -Value $NewRule
    $oldDueToken = _FormatTimerTransitionTime -Value $OldDueAt
    $newDueToken = _FormatTimerTransitionTime -Value $NewDueAt

    $rulesChanged = $oldRuleToken -ne $newRuleToken
    $dueChanged = $oldDueToken -ne $newDueToken
    if ([string]::IsNullOrWhiteSpace($Action)) {
        if (($oldRuleToken -eq '<none>') -and ($newRuleToken -ne '<none>')) {
            $Action = 'arm'
        } elseif (($oldRuleToken -ne '<none>') -and ($newRuleToken -eq '<none>')) {
            $Action = 'disarm'
        } elseif ($rulesChanged -or $dueChanged) {
            $Action = 'rearm'
        } else {
            return
        }
    }

    $gameName = _GetRuntimeTransitionGameName -Prefix $key -SharedState $SharedState
    $message =
        'TIMER decision ' +
        'prefix=' + $key + ' ' +
        'game=' + (_FormatRuntimeTransitionText -Value $gameName) + ' ' +
        'timer=' + $timerName + ' ' +
        'action=' + (_NormalizeRuntimeTransitionToken -Value $Action -Default 'observe') + ' ' +
        'from.rule=' + $oldRuleToken + ' ' +
        'to.rule=' + $newRuleToken + ' ' +
        'from.due=' + $oldDueToken + ' ' +
        'to.due=' + $newDueToken + ' ' +
        'reason=' + (_FormatRuntimeTransitionText -Value $Reason)

    _Log $message -Level DEBUG
}

function _TracePlayerObservationDecision {
    param(
        [string]$Prefix,
        [hashtable]$Profile,
        [object]$Snapshot,
        [hashtable]$SharedState = $script:State
    )

    $key = if ([string]::IsNullOrWhiteSpace($Prefix)) { '' } else { $Prefix.ToUpperInvariant() }
    if ([string]::IsNullOrWhiteSpace($key) -or -not $SharedState) { return }
    Initialize-ServerManager -SharedState $SharedState

    $supported = if ($Snapshot -and $Snapshot.Supported) { 'true' } else { 'false' }
    $available = if ($Snapshot -and $Snapshot.Available) { 'true' } else { 'false' }
    $source = _NormalizeRuntimeTransitionToken $(if ($Snapshot) { $Snapshot.Source } else { $null })
    $note = if ($Snapshot) { $Snapshot.Note } else { '' }
    $count = 0
    try { if ($Snapshot) { $count = [Math]::Max(0, [int]$Snapshot.Count) } } catch { $count = 0 }

    $signature = [ordered]@{
        Supported = $supported
        Available = $available
        Source    = $source
        Note      = [string]$note
    }

    $previous = $null
    try {
        if ($SharedState.PlayerObservationTrace.ContainsKey($key)) {
            $previous = $SharedState.PlayerObservationTrace[$key]
        }
    } catch { $previous = $null }

    $sameDecision = $false
    if ($previous) {
        $sameDecision =
            ("$($previous.Supported)" -eq $supported) -and
            ("$($previous.Available)" -eq $available) -and
            ("$($previous.Source)" -eq $source) -and
            ("$($previous.Note)" -eq [string]$note)
    }
    if ($sameDecision) { return }

    $SharedState.PlayerObservationTrace[$key] = $signature

    $gameName = _GetRuntimeTransitionGameName -Prefix $key -SharedState $SharedState
    $message =
        'PLAYER source ' +
        'prefix=' + $key + ' ' +
        'game=' + (_FormatRuntimeTransitionText -Value $gameName) + ' ' +
        'supported=' + $supported + ' ' +
        'available=' + $available + ' ' +
        'source=' + $source + ' ' +
        'count=' + $count + ' ' +
        'note=' + (_FormatRuntimeTransitionText -Value $note)

    _Log $message -Level DEBUG
}

function _FormatRehydrationTraceValue {
    param(
        [object]$Value
    )

    if ($Value -is [datetime]) {
        return (_FormatTimerTransitionTime -Value $Value)
    }

    if ($Value -is [bool]) {
        return $(if ([bool]$Value) { 'true' } else { 'false' })
    }

    if ($null -eq $Value) { return '<none>' }

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) { return '<none>' }

    if ($text -match '^[A-Za-z0-9_.:/-]+$') {
        return $text.Trim()
    }

    return (_FormatRuntimeTransitionText -Value $text)
}

function _TraceServerRehydrationDecision {
    param(
        [string]$Prefix,
        [string]$Phase,
        [string]$Action,
        [string]$Reason = '',
        [hashtable]$Fields = $null,
        [hashtable]$SharedState = $script:State
    )

    $key = if ([string]::IsNullOrWhiteSpace($Prefix)) { '' } else { $Prefix.ToUpperInvariant() }
    if ([string]::IsNullOrWhiteSpace($key)) { return }

    $gameName = _GetRuntimeTransitionGameName -Prefix $key -SharedState $SharedState
    $parts = New-Object 'System.Collections.Generic.List[string]'
    $parts.Add('REHYDRATE decision') | Out-Null
    $parts.Add('prefix=' + $key) | Out-Null
    $parts.Add('game=' + (_FormatRuntimeTransitionText -Value $gameName)) | Out-Null
    $parts.Add('phase=' + (_NormalizeRuntimeTransitionToken -Value $Phase -Default 'unknown')) | Out-Null
    $parts.Add('action=' + (_NormalizeRuntimeTransitionToken -Value $Action -Default 'observe')) | Out-Null

    if ($Fields) {
        foreach ($fieldKey in @($Fields.Keys)) {
            $parts.Add(($fieldKey + '=' + (_FormatRehydrationTraceValue -Value $Fields[$fieldKey]))) | Out-Null
        }
    }

    $parts.Add('reason=' + (_FormatRuntimeTransitionText -Value $Reason)) | Out-Null
    _Log ($parts -join ' ') -Level DEBUG
}

function _TraceDiscordSendDecision {
    param(
        [string]$Prefix = '',
        [object]$Profile = $null,
        [string]$Event = '',
        [string]$Tag = '',
        [string]$Action = '',
        [string]$Reason = '',
        [string]$Message = '',
        [hashtable]$SharedState = $script:State
    )

    $prefixKey = if ([string]::IsNullOrWhiteSpace($Prefix)) { '' } else { $Prefix.Trim().ToUpperInvariant() }
    $gameName = ''
    if (-not [string]::IsNullOrWhiteSpace($prefixKey)) {
        $gameName = _GetRuntimeTransitionGameName -Prefix $prefixKey -SharedState $SharedState
    } elseif ($Profile -and -not [string]::IsNullOrWhiteSpace([string]$Profile.GameName)) {
        $gameName = [string]$Profile.GameName
    } else {
        $gameName = 'ECC'
    }

    $parts = New-Object 'System.Collections.Generic.List[string]'
    $parts.Add('DISCORD decision') | Out-Null
    if (-not [string]::IsNullOrWhiteSpace($prefixKey)) {
        $parts.Add('prefix=' + $prefixKey) | Out-Null
    }
    $parts.Add('game=' + (_FormatRuntimeTransitionText -Value $gameName)) | Out-Null
    $parts.Add('event=' + (_NormalizeRuntimeTransitionToken -Value $Event -Default '<generic>')) | Out-Null
    $parts.Add('tag=' + (_NormalizeRuntimeTransitionToken -Value $Tag)) | Out-Null
    $parts.Add('action=' + (_NormalizeRuntimeTransitionToken -Value $Action -Default 'observe')) | Out-Null
    if (-not [string]::IsNullOrWhiteSpace($Message)) {
        $parts.Add('preview=' + (_FormatRuntimeTransitionText -Value $Message)) | Out-Null
    }
    $parts.Add('reason=' + (_FormatRuntimeTransitionText -Value $Reason)) | Out-Null
    _Log ($parts -join ' ') -Level DEBUG
}

function _LogProfileError {
    param(
        [string]$Prefix = '',
        [object]$Profile = $null,
        [string]$Message = '',
        [object]$ErrorRecord = $null,
        [string]$Action = '',
        [string]$Module = 'ServerManager',
        [string]$Function = '',
        [string]$Fallback = '',
        [string]$Recovery = '',
        [string]$Level = 'ERROR',
        [hashtable]$SharedState = $script:State
    )

    $prefixKey = if ([string]::IsNullOrWhiteSpace($Prefix)) { '' } else { $Prefix.Trim().ToUpperInvariant() }
    $gameName = ''
    if ($Profile -and -not [string]::IsNullOrWhiteSpace([string]$Profile.GameName)) {
        $gameName = [string]$Profile.GameName
    } elseif (-not [string]::IsNullOrWhiteSpace($prefixKey)) {
        $gameName = _GetRuntimeTransitionGameName -Prefix $prefixKey -SharedState $SharedState
    } else {
        $gameName = 'Unknown'
    }

    $parts = New-Object 'System.Collections.Generic.List[string]'
    $parts.Add('ERROR context') | Out-Null
    if (-not [string]::IsNullOrWhiteSpace($prefixKey)) {
        $parts.Add('prefix=' + $prefixKey) | Out-Null
    }
    $parts.Add('game=' + (_FormatRuntimeTransitionText -Value $gameName)) | Out-Null

    if (-not [string]::IsNullOrWhiteSpace($Action)) {
        $parts.Add('action=' + (_NormalizeRuntimeTransitionToken -Value $Action -Default '<none>')) | Out-Null
    }
    if (-not [string]::IsNullOrWhiteSpace($Module)) {
        $parts.Add('module=' + (_NormalizeRuntimeTransitionToken -Value $Module -Default '<none>')) | Out-Null
    }
    if (-not [string]::IsNullOrWhiteSpace($Function)) {
        $parts.Add('function=' + (_NormalizeRuntimeTransitionToken -Value $Function -Default '<none>')) | Out-Null
    }
    if (-not [string]::IsNullOrWhiteSpace($Message)) {
        $parts.Add('message=' + (_FormatRuntimeTransitionText -Value $Message)) | Out-Null
    }
    if ($null -ne $ErrorRecord) {
        $exceptionText = ''
        try { $exceptionText = $ErrorRecord.Exception.Message } catch { $exceptionText = '' }
        if ([string]::IsNullOrWhiteSpace($exceptionText)) {
            try { $exceptionText = [string]$ErrorRecord } catch { $exceptionText = '' }
        }
        if (-not [string]::IsNullOrWhiteSpace($exceptionText)) {
            $parts.Add('exception=' + (_FormatRuntimeTransitionText -Value $exceptionText)) | Out-Null
        }
    }
    if (-not [string]::IsNullOrWhiteSpace($Fallback)) {
        $parts.Add('fallback=' + (_FormatRuntimeTransitionText -Value $Fallback)) | Out-Null
    }
    if (-not [string]::IsNullOrWhiteSpace($Recovery)) {
        $parts.Add('recovery=' + (_FormatRuntimeTransitionText -Value $Recovery)) | Out-Null
    }

    _Log ($parts -join ' ') -Level $Level
}

function Set-ServerRuntimeState {
    param(
        [string]$Prefix,
        [string]$State,
        [string]$Detail = '',
        [string]$Badge = '',
        [string]$Level = 'INFO',
        [hashtable]$SharedState = $script:State,
        [switch]$PreserveSince
    )

    if ([string]::IsNullOrWhiteSpace($Prefix) -or -not $SharedState) { return }
    Initialize-ServerManager -SharedState $SharedState

    $key = $Prefix.ToUpperInvariant()
    $existing = $null
    if ($SharedState.ServerRuntimeState.ContainsKey($key)) {
        $existing = $SharedState.ServerRuntimeState[$key]
    }

    $since = Get-Date
    if ($PreserveSince -and $existing -and $existing.Since) {
        try { $since = [datetime]$existing.Since } catch { $since = Get-Date }
    }

    if ([string]::IsNullOrWhiteSpace($Badge)) {
        $Badge = _ResolveServerRuntimeBadge -State $State
    }

    $newEntry = [ordered]@{
        Code   = if ([string]::IsNullOrWhiteSpace($State)) { 'offline' } else { $State }
        Badge  = $Badge
        Detail = [string]$Detail
        Level  = if ([string]::IsNullOrWhiteSpace($Level)) { 'INFO' } else { $Level.ToUpperInvariant() }
        Since  = $since
    }
    $SharedState.ServerRuntimeState[$key] = $newEntry

    _TraceServerRuntimeStateTransition -Prefix $key -OldEntry $existing -NewEntry $newEntry -SharedState $SharedState -Action 'set' -PreserveSince:$PreserveSince
}

function Get-ServerRuntimeState {
    param(
        [string]$Prefix,
        [hashtable]$SharedState = $script:State
    )

    if ([string]::IsNullOrWhiteSpace($Prefix) -or -not $SharedState -or -not $SharedState.ContainsKey('ServerRuntimeState')) {
        return $null
    }

    $key = $Prefix.ToUpperInvariant()
    if ($SharedState.ServerRuntimeState.ContainsKey($key)) {
        return $SharedState.ServerRuntimeState[$key]
    }

    return $null
}

function Clear-ServerRuntimeState {
    param(
        [string]$Prefix,
        [hashtable]$SharedState = $script:State
    )

    if ([string]::IsNullOrWhiteSpace($Prefix) -or -not $SharedState -or -not $SharedState.ContainsKey('ServerRuntimeState')) {
        return
    }

    $key = $Prefix.ToUpperInvariant()
    if ($SharedState.ServerRuntimeState.ContainsKey($key)) {
        $existing = $SharedState.ServerRuntimeState[$key]
        $SharedState.ServerRuntimeState.Remove($key)
        _TraceServerRuntimeStateTransition -Prefix $key -OldEntry $existing -NewEntry $null -SharedState $SharedState -Action 'clear'
    }
}

function Set-JoinableServerRuntimeState {
    param(
        [string]$Prefix,
        [string]$Detail = 'Server is joinable and accepting connections.',
        [hashtable]$SharedState = $script:State
    )

    if ([string]::IsNullOrWhiteSpace($Prefix) -or -not $SharedState) { return }
    Initialize-ServerManager -SharedState $SharedState

    $key = $Prefix.ToUpperInvariant()
    $runtimeCode = ''
    $currentCount = -1

    try {
        $runtime = Get-ServerRuntimeState -Prefix $key -SharedState $SharedState
        if ($runtime) { $runtimeCode = [string]$runtime.Code }
    } catch { $runtimeCode = '' }

    try {
        if ($SharedState.ContainsKey('LatestPlayerCounts') -and $SharedState.LatestPlayerCounts -and $SharedState.LatestPlayerCounts.ContainsKey($key)) {
            $currentCount = [int]$SharedState.LatestPlayerCounts[$key]
        }
    } catch { $currentCount = -1 }

    $runtimeCode = $runtimeCode.ToLowerInvariant()
    if ($currentCount -le 0 -and $runtimeCode -in @('waiting_first_player','idle_wait','idle_shutdown','waiting_restart')) {
        return
    }

    Set-ServerRuntimeState -Prefix $key -State 'online' -Detail $Detail -SharedState $SharedState
}

function Set-ObservedPlayersServerRuntimeState {
    param(
        [string]$Prefix,
        [int]$Count = 0,
        [hashtable]$SharedState = $script:State
    )

    if ([string]::IsNullOrWhiteSpace($Prefix) -or -not $SharedState) { return }
    Initialize-ServerManager -SharedState $SharedState

    $key = $Prefix.ToUpperInvariant()
    $resolvedCount = [Math]::Max(0, [int]$Count)
    if ($resolvedCount -gt 0) {
        Set-ServerRuntimeState -Prefix $key -State 'online' -Detail ("{0} player(s) online." -f $resolvedCount) -SharedState $SharedState
        return
    }

    Set-JoinableServerRuntimeState -Prefix $key -SharedState $SharedState
}

function _ProfileUsesDeferredReadySignal {
    param([hashtable]$Profile)

    if ($null -eq $Profile) { return $false }

    return (
        (_TestProfileGame -Profile $Profile -KnownGame 'ProjectZomboid') -or
        (_TestProfileGame -Profile $Profile -KnownGame 'Hytale') -or
        (_TestProfileGame -Profile $Profile -KnownGame 'Valheim') -or
        (_TestProfileGame -Profile $Profile -KnownGame '7DaysToDie') -or
        (_TestProfileGame -Profile $Profile -KnownGame 'Palworld') -or
        (_TestProfileGame -Profile $Profile -KnownGame 'Minecraft') -or
        (_TestProfileGame -Profile $Profile -KnownGame 'Satisfactory')
    )
}

function _GetSystemMemorySnapshot {
    $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
    $totalBytes = [double]$os.TotalVisibleMemorySize * 1KB
    $freeBytes  = [double]$os.FreePhysicalMemory * 1KB
    $usedBytes  = [Math]::Max(0, $totalBytes - $freeBytes)
    $usedPercent = if ($totalBytes -gt 0) { [Math]::Round(($usedBytes / $totalBytes) * 100, 1) } else { 0.0 }
    $freeGb = [Math]::Round(($freeBytes / 1GB), 2)
    $usedGb = [Math]::Round(($usedBytes / 1GB), 2)
    $totalGb = [Math]::Round(($totalBytes / 1GB), 2)

    return [ordered]@{
        TotalBytes   = [int64][Math]::Round($totalBytes)
        FreeBytes    = [int64][Math]::Round($freeBytes)
        UsedBytes    = [int64][Math]::Round($usedBytes)
        TotalGb      = $totalGb
        FreeGb       = $freeGb
        UsedGb       = $usedGb
        UsedPercent  = $usedPercent
    }
}

function _GetProfileStartRamGuard {
    param([hashtable]$Profile)

    $maxUsedPercent = 0.0
    $minFreeGb = 0.0

    try {
        if ($null -ne $Profile.BlockStartIfRamPercentUsed) {
            [void][double]::TryParse("$($Profile.BlockStartIfRamPercentUsed)", [ref]$maxUsedPercent)
        }
    } catch { $maxUsedPercent = 0.0 }

    try {
        if ($null -ne $Profile.BlockStartIfFreeRamBelowGB) {
            [void][double]::TryParse("$($Profile.BlockStartIfFreeRamBelowGB)", [ref]$minFreeGb)
        }
    } catch { $minFreeGb = 0.0 }

    return [ordered]@{
        MaxUsedPercent = [Math]::Max(0.0, $maxUsedPercent)
        MinFreeGb      = [Math]::Max(0.0, $minFreeGb)
    }
}

function Test-StartBlockedByRam {
    param([hashtable]$Profile)

    $guard = _GetProfileStartRamGuard -Profile $Profile
    $hasPercentRule = ($guard.MaxUsedPercent -gt 0)
    $hasFreeRule = ($guard.MinFreeGb -gt 0)

    if (-not $hasPercentRule -and -not $hasFreeRule) {
        return [ordered]@{
            Blocked  = $false
            Reason   = ''
            Summary  = ''
            Snapshot = $null
            Guard    = $guard
        }
    }

    $snapshot = _GetSystemMemorySnapshot
    $reasons = New-Object 'System.Collections.Generic.List[string]'

    if ($hasPercentRule -and $snapshot.UsedPercent -ge $guard.MaxUsedPercent) {
        $reasons.Add(("RAM usage is {0}% (limit {1}%)." -f $snapshot.UsedPercent, $guard.MaxUsedPercent)) | Out-Null
    }
    if ($hasFreeRule -and $snapshot.FreeGb -lt $guard.MinFreeGb) {
        $reasons.Add(("Free RAM is {0} GB (minimum {1} GB)." -f $snapshot.FreeGb, $guard.MinFreeGb)) | Out-Null
    }

    $summary = if ($reasons.Count -gt 0) {
        "Start blocked by RAM guard: " + ($reasons -join ' ')
    } else {
        ''
    }

    return [ordered]@{
        Blocked  = ($reasons.Count -gt 0)
        Reason   = ($reasons -join ' ')
        Summary  = $summary
        Snapshot = $snapshot
        Guard    = $guard
    }
}

function _GetProfileStartupTimeoutSeconds {
    param([hashtable]$Profile)

    $timeoutSeconds = 300
    $knownGame = ''
    try {
        if ($Profile) {
            if ($Profile.ContainsKey('KnownGame') -and -not [string]::IsNullOrWhiteSpace([string]$Profile.KnownGame)) {
                $knownGame = [string]$Profile.KnownGame
            } elseif ($Profile.ContainsKey('Prefix')) {
                switch (("$($Profile.Prefix)").ToUpperInvariant()) {
                    'PZ' { $knownGame = 'ProjectZomboid' }
                }
            }
        }
    } catch { $knownGame = '' }

    # Project Zomboid can take materially longer to produce a trusted ready signal
    # than the other profiles, especially on large modded worlds.
    if ($knownGame -eq 'ProjectZomboid') {
        $timeoutSeconds = 600
    }

    try {
        if ($null -ne $Profile.StartupTimeoutSeconds) {
            $candidate = 0
            if ([int]::TryParse("$($Profile.StartupTimeoutSeconds)", [ref]$candidate) -and $candidate -gt 0) {
                $timeoutSeconds = $candidate
            }
        }
    } catch { }

    if ($knownGame -eq 'ProjectZomboid' -and $timeoutSeconds -lt 600) {
        $timeoutSeconds = 600
    }

    return $timeoutSeconds
}

function _HandleStartupFailure {
    param(
        [string]$Prefix,
        [hashtable]$Profile,
        [hashtable]$SharedState,
        [string]$Reason
    )

    $prefixKey = $Prefix.ToUpperInvariant()
    $gameName = _GetNotificationGameName -Profile $Profile -Prefix $prefixKey
    $message = if ([string]::IsNullOrWhiteSpace($Reason)) {
        'Startup failed before the server became ready.'
    } else {
        $Reason
    }

    $startupSummary = _JoinNotificationParts -Parts @(
        'Startup health check failed',
        $message
    )

    _Log "[$gameName] $startupSummary" -Level WARN
    _GameLog -Prefix $prefixKey -Line ("[WARN] {0}" -f $startupSummary)

    $autoRestart = $true
    if ($null -ne $Profile.EnableAutoRestart) {
        $autoRestart = [bool]$Profile.EnableAutoRestart
    }

    try {
        Invoke-SafeShutdown -Prefix $prefixKey -Quiet | Out-Null
    } catch {
        _Log "[$gameName] Startup failure shutdown error: $($_.Exception.Message)" -Level WARN
    }

    if (-not $autoRestart) {
        Set-ServerRuntimeState -Prefix $prefixKey -State 'startup_failed' -Detail $message -SharedState $SharedState
        _Webhook (New-DiscordGameMessage -Profile $Profile -Prefix $prefixKey -Event 'startup_failed' -Values @{
            Reason = $message
            Action = 'Auto-restart is disabled, so the server will stay offline.'
        })
        return
    }

    $hour = (Get-Date).ToString('yyyyMMddHH')
    $counterKey = "${prefixKey}_${hour}"
    if ($null -eq $SharedState.RestartCounters[$counterKey]) {
        $SharedState.RestartCounters[$counterKey] = 0
    }

    $count = [int]$SharedState.RestartCounters[$counterKey]
    $max = 5
    if ($null -ne $Profile.MaxRestartsPerHour) {
        $max = [int]$Profile.MaxRestartsPerHour
    }

    if ($count -ge $max) {
        Set-ServerRuntimeState -Prefix $prefixKey -State 'startup_failed' -Detail ("Startup failed and auto-restart is suspended after {0} failures this hour." -f $count) -SharedState $SharedState
        _Webhook (New-DiscordGameMessage -Profile $Profile -Prefix $prefixKey -Event 'startup_failed' -Values @{
            Reason = $message
            Action = ("Auto-restart is suspended after {0} failure(s) this hour." -f $count)
        })
        _ClearPendingAutoRestart -Prefix $prefixKey -SharedState $SharedState
        return
    }

    $SharedState.RestartCounters[$counterKey] = $count + 1
    $delay = 10
    if ($null -ne $Profile.RestartDelaySeconds) {
        $delay = [int]$Profile.RestartDelaySeconds
    }

    Set-ServerRuntimeState -Prefix $prefixKey -State 'waiting_restart' -Detail ("Startup failed. Restarting in {0}s (attempt {1}/{2})." -f $delay, ($count + 1), $max) -SharedState $SharedState
    _Webhook (New-DiscordGameMessage -Profile $Profile -Prefix $prefixKey -Event 'startup_failed' -Values @{
        Reason = $message
        Action = ("Recovery restart queued in {0}s (attempt {1}/{2})." -f $delay, ($count + 1), $max)
    })
    _SchedulePendingAutoRestart -Prefix $prefixKey -Profile $Profile -SharedState $SharedState -DelaySeconds $delay -Attempt ($count + 1) -MaxAttempts $max
}

function _CheckStartupHealth {
    param(
        [string]$Prefix,
        [hashtable]$Profile,
        [hashtable]$SharedState
    )

    if (-not $Profile -or -not (_ProfileUsesDeferredReadySignal -Profile $Profile)) { return $false }
    if (-not $SharedState -or -not $SharedState.RunningServers.ContainsKey($Prefix)) { return $false }

    $runtime = Get-ServerRuntimeState -Prefix $Prefix -SharedState $SharedState
    $stateCode = ''
    try { $stateCode = [string]$runtime.Code } catch { $stateCode = '' }
    if ($stateCode.ToLowerInvariant() -ne 'starting') { return $false }

    $entry = $SharedState.RunningServers[$Prefix]
    if ($null -eq $entry -or $null -eq $entry.StartTime) { return $false }

    $timeoutSeconds = _GetProfileStartupTimeoutSeconds -Profile $Profile
    if ($timeoutSeconds -le 0) { return $false }

    $elapsedSeconds = ((Get-Date) - [datetime]$entry.StartTime).TotalSeconds
    if ($elapsedSeconds -lt $timeoutSeconds) { return $false }

    $reason = "Server process started but no healthy/joinable signal arrived within ${timeoutSeconds}s."
    _HandleStartupFailure -Prefix $Prefix -Profile $Profile -SharedState $SharedState -Reason $reason
    return $true
}

function _GetProfileIdleTimeoutMinutes {
    param(
        [hashtable]$Profile,
        [string]$Key
    )

    $minutes = 0
    try {
        if ($Profile -and $null -ne $Profile[$Key]) {
            $candidate = 0
            if ([int]::TryParse("$($Profile[$Key])", [ref]$candidate) -and $candidate -gt 0) {
                $minutes = $candidate
            }
        }
    } catch { }

    return $minutes
}

function _FormatCountdownShort {
    param([datetime]$DueAt)

    try {
        $remainingSeconds = [Math]::Max(0, [int][Math]::Ceiling(($DueAt - (Get-Date)).TotalSeconds))
        if ($remainingSeconds -lt 60) {
            return "${remainingSeconds}s"
        }

        $minutes = [Math]::Floor($remainingSeconds / 60)
        $seconds = $remainingSeconds % 60
        if ($seconds -le 0) {
            return "${minutes}m"
        }

        return "${minutes}m ${seconds}s"
    } catch {
        return 'soon'
    }
}

function _FormatSignalWaitRuntimeDetail {
    param(
        [string]$StateCode,
        [datetime]$DueAt
    )

    $isEmptyWait = ("$StateCode".ToLowerInvariant() -eq 'idle_wait')
    $baseText = if ($isEmptyWait) {
        'Server empty. Idle shutdown'
    } else {
        'Waiting for first player. Idle shutdown'
    }

    if ($DueAt) {
        try {
            if (((Get-Date) - $DueAt).TotalSeconds -ge 0) {
                return "$baseText overdue."
            }
        } catch { }

        return "$baseText in $(_FormatCountdownShort -DueAt $DueAt)."
    }

    if ($isEmptyWait) {
        return 'Server empty. Idle shutdown pending trusted player data.'
    }

    return 'Waiting for first player. Idle shutdown pending trusted player data.'
}

function _ResetPlayerActivityTracking {
    param(
        [string]$Prefix,
        [hashtable]$SharedState = $script:State
    )

    if ([string]::IsNullOrWhiteSpace($Prefix) -or -not $SharedState) { return }
    Initialize-ServerManager -SharedState $SharedState

    $key = $Prefix.ToUpperInvariant()
    $startupAt = Get-Date
    try {
        if ($SharedState.RunningServers.ContainsKey($key) -and $SharedState.RunningServers[$key].StartTime) {
            $startupAt = [datetime]$SharedState.RunningServers[$key].StartTime
        }
    } catch { }

    try {
        if ($SharedState.ContainsKey('LatestPlayers') -and $SharedState.LatestPlayers.ContainsKey($key)) {
            $SharedState.LatestPlayers.Remove($key) | Out-Null
        }
    } catch { }
    try {
        if ($SharedState.ContainsKey('LatestPlayerCounts') -and $SharedState.LatestPlayerCounts.ContainsKey($key)) {
            $SharedState.LatestPlayerCounts.Remove($key) | Out-Null
        }
    } catch { }
    try {
        if ($SharedState.ContainsKey('LatestPlayerObservedAt') -and $SharedState.LatestPlayerObservedAt.ContainsKey($key)) {
            $SharedState.LatestPlayerObservedAt.Remove($key) | Out-Null
        }
    } catch { }

    $SharedState.PlayerActivityState[$key] = [ordered]@{
        StartupAt          = $startupAt
        CurrentCount       = 0
        LastCount          = 0
        LastObservedAt     = $null
        FirstPlayerSeenAt  = $null
        LastNonZeroAt      = $null
        LastBecameEmptyAt  = $null
        DetectionSupported = $false
        DetectionAvailable = $false
        DetectionSource    = ''
        Note               = ''
        PendingRule        = ''
        ShutdownDueAt      = $null
    }
}

function _SetLatestPlayersSnapshot {
    param(
        [string]$Prefix,
        [string[]]$Names = @(),
        [int]$Count = 0,
        [hashtable]$SharedState = $script:State
    )

    if ([string]::IsNullOrWhiteSpace($Prefix) -or -not $SharedState) { return }
    Initialize-ServerManager -SharedState $SharedState

    $key = $Prefix.ToUpperInvariant()
    $safeNames = @()
    foreach ($name in @($Names)) {
        if ([string]::IsNullOrWhiteSpace("$name")) { continue }
        $trimmed = "$name".Trim()
        if ($safeNames -notcontains $trimmed) { $safeNames += $trimmed }
    }

    $resolvedCount = [Math]::Max([int]$Count, @($safeNames).Count)
    $SharedState.LatestPlayers[$key] = @($safeNames)
    $SharedState.LatestPlayerCounts[$key] = $resolvedCount
    if ($SharedState.ContainsKey('LatestPlayerObservedAt') -and $SharedState.LatestPlayerObservedAt) {
        $SharedState.LatestPlayerObservedAt[$key] = Get-Date
    }
}

function Set-LatestPlayersSnapshot {
    param(
        [string]$Prefix,
        [string[]]$Names = @(),
        [int]$Count = 0,
        [hashtable]$SharedState = $script:State
    )

    _SetLatestPlayersSnapshot -Prefix $Prefix -Names @($Names) -Count $Count -SharedState $SharedState
}

function _Parse7DaysToDiePlayersResponse {
    param([string]$ResponseText)

    $result = [ordered]@{
        Available = $false
        Count     = 0
        Names     = @()
        Note      = ''
    }

    if ([string]::IsNullOrWhiteSpace($ResponseText)) {
        $result.Note = 'No telnet response text was returned.'
        return $result
    }

    $lines = @($ResponseText -split "(`r`n|`n|`r)" | ForEach-Object { "$_".Trim() } | Where-Object { $_ -ne '' })
    if (-not $lines -or $lines.Count -eq 0) {
        $result.Note = 'No non-empty telnet response lines were available.'
        return $result
    }

    $names = New-Object 'System.Collections.Generic.List[string]'
    $countKnown = $false
    $countValue = 0

    foreach ($line in $lines) {
        $normalizedLine = [string]$line
        if ($normalizedLine -match '^\[TELNET\]\s*(.+)$') {
            $normalizedLine = $Matches[1].Trim()
        }
        if ($normalizedLine -match '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\s+\S+\s+\w+\s+(.*)$') {
            $normalizedLine = $Matches[1].Trim()
        }
        if ([string]::IsNullOrWhiteSpace($normalizedLine)) { continue }

        if ($normalizedLine -match '(?i)\bno players\b|\bno one\b|\b0 players\b|\btotal of 0 in the game\b') {
            $countKnown = $true
            $countValue = 0
            continue
        }

        if ($normalizedLine -match '(?i)\btotal of\s+(\d+)\s+in the game\b') {
            $countKnown = $true
            $countValue = [int]$Matches[1]
            continue
        }

        if ($normalizedLine -match '(?i)\bplayers?\s+online\s*:\s*(\d+)\b') {
            $countKnown = $true
            $countValue = [int]$Matches[1]
            continue
        }

        if ($normalizedLine -match '(?i)\bname\s*=\s*([^,|]+)') {
            $name = $Matches[1].Trim()
            if ($name -and $names -notcontains $name) { $names.Add($name) | Out-Null }
            continue
        }

        if ($normalizedLine -match '(?i)\bid\s*=\s*\d+\s*,\s*(?:"([^"]+)"|([^,]+?))\s*(?:,|$)') {
            $name = ''
            if (-not [string]::IsNullOrWhiteSpace($Matches[1])) {
                $name = $Matches[1].Trim()
            } elseif (-not [string]::IsNullOrWhiteSpace($Matches[2])) {
                $name = $Matches[2].Trim()
            }
            if ($name -and $names -notcontains $name) { $names.Add($name) | Out-Null }
            continue
        }

        if ($normalizedLine -match '^\s*\d+\.\s+([A-Za-z0-9 _\-\.\[\]]+?)\s*(?:\(|$)') {
            $name = $Matches[1].Trim()
            if ($name -and $names -notcontains $name) { $names.Add($name) | Out-Null }
            continue
        }

        if ($normalizedLine -match '^\s*\d+\s*[:\-]\s*([A-Za-z0-9 _\-\.\[\]]+?)\s*$') {
            $name = $Matches[1].Trim()
            if ($name -and $names -notcontains $name) { $names.Add($name) | Out-Null }
            continue
        }
    }

    if ($countKnown -or $names.Count -gt 0) {
        $result.Available = $true
        $result.Names = @($names)
        $result.Count = if ($countKnown) { [Math]::Max($countValue, $names.Count) } else { $names.Count }
        $result.Note = 'Player count observed from 7 Days to Die telnet output.'
    } else {
        $result.Note = '7 Days to Die telnet output did not include a recognizable player list yet.'
    }

    return $result
}

function _Parse7DaysToDiePlayersLogText {
    param([string]$Text)

    $result = [ordered]@{
        Available = $false
        Count     = 0
        Names     = @()
        Note      = ''
    }

    if ([string]::IsNullOrWhiteSpace($Text)) {
        $result.Note = 'No 7 Days to Die log text was available.'
        return $result
    }

    $lines = @($Text -split "(`r`n|`n|`r)")
    $startIndex = 0
    for ($i = ($lines.Count - 1); $i -ge 0; $i--) {
        $line = ([string]$lines[$i]).Trim()
        if ($line -match "(?i)executing command 'listplayers'" -or $line -match '(?i)\[TELNET\].*\blistplayers\b') {
            $startIndex = $i
            break
        }
    }

    $segment = @()
    for ($i = $startIndex; $i -lt $lines.Count; $i++) {
        $line = ([string]$lines[$i]).Trim()
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $segment += $line
    }

    if (@($segment).Count -eq 0) {
        $result.Note = 'Recent 7 Days to Die logs did not include a listplayers segment yet.'
        return $result
    }

    $parsed = _Parse7DaysToDiePlayersResponse -ResponseText ($segment -join [Environment]::NewLine)
    if ($parsed.Available) {
        $parsed.Note = 'Player count observed from 7 Days to Die log output.'
        return $parsed
    }

    $generic = _ParseGenericPlayersText -Text ($segment -join [Environment]::NewLine)
    if ($generic.Available) {
        $generic.Note = 'Player count observed from parsed 7 Days to Die log text.'
        return $generic
    }

    $recentRoster = [ordered]@{}
    $playerCount = $null
    foreach ($lineRaw in $lines) {
        $line = ([string]$lineRaw).Trim()
        if ([string]::IsNullOrWhiteSpace($line)) { continue }

        if ($line -match "(?i)GMSG:\s+Player\s+'([^']+)'\s+joined the game") {
            $name = [string]$Matches[1].Trim()
            if (-not [string]::IsNullOrWhiteSpace($name)) {
                $nameKey = $name.ToLowerInvariant()
                $recentRoster[$nameKey] = $name
            }
            continue
        }

        if ($line -match "(?i)PlayerSpawnedInWorld.*PlayerName='([^']+)'") {
            $name = [string]$Matches[1].Trim()
            if (-not [string]::IsNullOrWhiteSpace($name)) {
                $nameKey = $name.ToLowerInvariant()
                $recentRoster[$nameKey] = $name
            }
            continue
        }

        if ($line -match "(?i)Player\s+(.+?)\s+disconnected after\b") {
            $name = [string]$Matches[1].Trim()
            if (-not [string]::IsNullOrWhiteSpace($name)) {
                $nameKey = $name.ToLowerInvariant()
                if ($recentRoster.Contains($nameKey)) {
                    $recentRoster.Remove($nameKey) | Out-Null
                }
            }
            continue
        }

        if ($line -match "(?i)GMSG:\s+Player\s+'([^']+)'\s+left the game") {
            $name = [string]$Matches[1].Trim()
            if (-not [string]::IsNullOrWhiteSpace($name)) {
                $nameKey = $name.ToLowerInvariant()
                if ($recentRoster.Contains($nameKey)) {
                    $recentRoster.Remove($nameKey) | Out-Null
                }
            }
            continue
        }

        if ($line -match "(?i)Player disconnected:.*PlayerName='([^']+)'") {
            $name = [string]$Matches[1].Trim()
            if (-not [string]::IsNullOrWhiteSpace($name)) {
                $nameKey = $name.ToLowerInvariant()
                if ($recentRoster.Contains($nameKey)) {
                    $recentRoster.Remove($nameKey) | Out-Null
                }
            }
            continue
        }

        if ($line -match '\bPly:\s*(\d+)\b') {
            try { $playerCount = [int]$Matches[1] } catch { $playerCount = 0 }
            if ($playerCount -le 0) {
                $recentRoster.Clear()
            }
            continue
        }
    }

    $recentNames = @($recentRoster.Values)
    if ($null -ne $playerCount -or @($recentNames).Count -gt 0) {
        $result.Available = $true
        $result.Names = @($recentNames)
        if ($null -eq $playerCount) {
            $result.Count = @($recentNames).Count
        } else {
            $result.Count = [Math]::Max([int]$playerCount, @($recentNames).Count)
        }
        $result.Note = 'Player count observed from 7 Days to Die join/activity log lines.'
        return $result
    }

    $result.Note = 'Recent 7 Days to Die logs did not include a recognizable player roster yet.'
    return $result
}

function _ParseGenericPlayersText {
    param([string]$Text)

    $result = [ordered]@{
        Available = $false
        Count     = 0
        Names     = @()
        Note      = ''
    }

    if ([string]::IsNullOrWhiteSpace($Text)) {
        $result.Note = 'No player text was available.'
        return $result
    }

    $names = New-Object 'System.Collections.Generic.List[string]'
    $seen = @{}
    $countKnown = $false
    $countValue = 0

    $quotedMatches = [regex]::Matches($Text, '"([^"]+)"')
    foreach ($match in $quotedMatches) {
        $name = $match.Groups[1].Value.Trim()
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        $key = $name.ToLowerInvariant()
        if (-not $seen.ContainsKey($key)) {
            $seen[$key] = $true
            $names.Add($name) | Out-Null
        }
    }

    $lines = @($Text -split "(`r`n|`n|`r)")
    $captureDashNames = $false
    foreach ($lineRaw in $lines) {
        $line = $lineRaw.Trim()
        if ([string]::IsNullOrWhiteSpace($line)) { continue }

        if ($line -match '(?i)\bno players\b|\bno one online\b|\b0 players\b') {
            $countKnown = $true
            $countValue = 0
        }
        if ($line -match '(?i)\b(\d+)\s+players?\s+online\b') {
            $countKnown = $true
            $countValue = [int]$Matches[1]
        }

        if ($captureDashNames) {
            if ($line -match '^\s*-\s*(.+?)\s*$') {
                $name = $Matches[1].Trim()
                if (-not [string]::IsNullOrWhiteSpace($name)) {
                    $key = $name.ToLowerInvariant()
                    if (-not $seen.ContainsKey($key)) {
                        $seen[$key] = $true
                        $names.Add($name) | Out-Null
                    }
                }
                continue
            }

            if ($line -notmatch '^\[' -and $line -notmatch '^(?i:players?\s+)') {
                $captureDashNames = $false
            }
        }

        if ($line -match '(?i)players\s+connected\s*\((\d+)\)\s*:\s*$') {
            $captureDashNames = $true
            $countKnown = $true
            $countValue = [int]$Matches[1]
            continue
        }

        if ($line -match '(?i)players?.*?:\s*(.+)$') {
            foreach ($part in ($Matches[1] -split ',')) {
                $name = $part.Trim(" `t[](){}")
                if ([string]::IsNullOrWhiteSpace($name)) { continue }
                $key = $name.ToLowerInvariant()
                if (-not $seen.ContainsKey($key)) {
                    $seen[$key] = $true
                    $names.Add($name) | Out-Null
                }
            }
        }
    }

    if ($names.Count -gt 0 -or $countKnown) {
        $result.Available = $true
        $result.Names = @($names)
        $result.Count = if ($countKnown) { [Math]::Max($countValue, $names.Count) } else { $names.Count }
        $result.Note = 'Player list observed from parsed command/log text.'
    } else {
        $result.Note = 'Player list text did not include a recognizable roster yet.'
    }

    return $result
}

function _ParseHytaleWhoText {
    param([string]$Text)

    $result = [ordered]@{
        Available = $false
        Count     = 0
        Names     = @()
        Note      = ''
    }

    if ([string]::IsNullOrWhiteSpace($Text)) {
        $result.Note = 'No Hytale who text was available.'
        return $result
    }

    $lines = @($Text -split "(`r`n|`n|`r)")
    $lastWhoIndex = -1
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = [string]$lines[$i]
        if ($line -match '(?i)\[CommandManager\]\s+Console executed command:\s*who\b') {
            $lastWhoIndex = $i
        }
    }

    if ($lastWhoIndex -lt 0) {
        $result.Note = 'No recent Hytale who command marker was found in the log.'
        return $result
    }

    $names = New-Object 'System.Collections.Generic.List[string]'
    $seen = @{}
    $countKnown = $false
    $countValue = 0

    for ($i = ($lastWhoIndex + 1); $i -lt $lines.Count; $i++) {
        $line = ([string]$lines[$i]).Trim()
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        if ($line -match '(?i)\[CommandManager\]\s+Console executed command:') { break }
        if ($line -match '^\[\d{4}/\d{2}/\d{2}\s+') { continue }

        if ($line -match '^[^()]+?\((\d+)\)\s*:\s*(.*)$') {
            $countKnown = $true
            $countValue += [int]$Matches[1]

            $rosterText = [string]$Matches[2]
            $rosterText = $rosterText.Trim()
            $rosterText = $rosterText.TrimStart(':').Trim()

            if ([string]::IsNullOrWhiteSpace($rosterText) -or $rosterText -match '^(?i)\(empty\)|empty$') {
                continue
            }

            foreach ($part in ($rosterText -split ',')) {
                $name = $part.Trim(" `t[](){}:")
                if ([string]::IsNullOrWhiteSpace($name)) { continue }
                $key = $name.ToLowerInvariant()
                if (-not $seen.ContainsKey($key)) {
                    $seen[$key] = $true
                    $names.Add($name) | Out-Null
                }
            }
        }
    }

    if ($countKnown -or $names.Count -gt 0) {
        $result.Available = $true
        $result.Names = @($names)
        $result.Count = if ($countKnown) { [Math]::Max($countValue, $names.Count) } else { $names.Count }
        $result.Note = 'Player count observed from Hytale who output.'
    } else {
        $result.Note = 'Recent Hytale who output did not include a recognizable roster yet.'
    }

    return $result
}

function _ResolveProfileLogPath {
    param(
        [hashtable]$Profile,
        [string]$Prefix = ''
    )

    if (-not $Profile) { return '' }

    $folderPath = ''
    try { $folderPath = [Environment]::ExpandEnvironmentVariables((_CoalesceStr $Profile.FolderPath '')) } catch { $folderPath = (_CoalesceStr $Profile.FolderPath '') }
    $serverLogPath = ''
    try { $serverLogPath = [Environment]::ExpandEnvironmentVariables((_CoalesceStr $Profile.ServerLogPath '')) } catch { $serverLogPath = (_CoalesceStr $Profile.ServerLogPath '') }
    $serverLogSubDir = _CoalesceStr $Profile.ServerLogSubDir ''
    $serverLogFile = _CoalesceStr $Profile.ServerLogFile ''

    if (-not [string]::IsNullOrWhiteSpace($serverLogPath) -and $serverLogPath.IndexOfAny(@('*','?')) -lt 0 -and (Test-Path -LiteralPath $serverLogPath)) {
        return $serverLogPath
    }

    $searchRoot = ''
    if (-not [string]::IsNullOrWhiteSpace($serverLogSubDir) -and -not [string]::IsNullOrWhiteSpace($folderPath)) {
        $searchRoot = Join-Path $folderPath $serverLogSubDir
    } elseif (-not [string]::IsNullOrWhiteSpace($serverLogPath)) {
        try { $searchRoot = Split-Path -Path $serverLogPath -Parent } catch { $searchRoot = '' }
    }

    $pattern = ''
    if (-not [string]::IsNullOrWhiteSpace($serverLogFile)) {
        $pattern = $serverLogFile
    } elseif (-not [string]::IsNullOrWhiteSpace($serverLogPath)) {
        try { $pattern = Split-Path -Path $serverLogPath -Leaf } catch { $pattern = '' }
    }

    if ([string]::IsNullOrWhiteSpace($searchRoot) -or -not (Test-Path -LiteralPath $searchRoot)) { return '' }

    if (-not [string]::IsNullOrWhiteSpace($pattern)) {
        if ($pattern.IndexOfAny(@('*','?')) -ge 0) {
            $match = Get-ChildItem -LiteralPath $searchRoot -Filter $pattern -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
            if ($match) { return $match.FullName }
        } else {
            $candidate = Join-Path $searchRoot $pattern
            if (Test-Path -LiteralPath $candidate) { return $candidate }
        }
    }

    $fallback = Get-ChildItem -LiteralPath $searchRoot -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($fallback) { return $fallback.FullName }
    return ''
}

function _ReadRecentProfileLogText {
    param(
        [hashtable]$Profile,
        [string]$Prefix = '',
        [int]$TailLines = 120
    )

    $path = _ResolveProfileLogPath -Profile $Profile -Prefix $Prefix
    if ([string]::IsNullOrWhiteSpace($path) -or -not (Test-Path -LiteralPath $path)) { return '' }
    try {
        return (@(Get-Content -LiteralPath $path -Tail $TailLines -ErrorAction Stop) -join [Environment]::NewLine)
    } catch {
        return ''
    }
}

function _Refresh7DaysToDiePlayers {
    param(
        [string]$Prefix,
        [hashtable]$Profile,
        [hashtable]$SharedState,
        [switch]$Force
    )

    if (-not $Profile -or -not $SharedState) { return $false }
    if (-not (_TestProfileGame -Profile $Profile -KnownGame '7DaysToDie')) { return $false }

    $key = $Prefix.ToUpperInvariant()
    Initialize-ServerManager -SharedState $SharedState

    if (-not $SharedState.RunningServers.ContainsKey($key)) { return $false }

    $telnetPass = if ($null -ne $Profile.TelnetPassword) { "$($Profile.TelnetPassword)" } else { '' }
    $telnetPort = if ($null -ne $Profile.TelnetPort) { [int]$Profile.TelnetPort } else { 8081 }
    $telnetHost = if ($null -ne $Profile.TelnetHost -and "$($Profile.TelnetHost)".Trim() -ne '') { "$($Profile.TelnetHost)".Trim() } else { '127.0.0.1' }
    if ([string]::IsNullOrWhiteSpace($telnetPass) -or $telnetPort -le 0) { return $false }

    $now = Get-Date
    $pollState = $null
    if ($SharedState.PlayerQueryState.ContainsKey($key)) {
        $pollState = $SharedState.PlayerQueryState[$key]
    }
    if ($null -eq $pollState) {
        $pollState = [ordered]@{
            LastAttemptAt = $null
            LastSuccessAt = $null
            Source        = ''
            Note          = ''
        }
        $SharedState.PlayerQueryState[$key] = $pollState
    }

    $refreshIntervalSeconds = 25
    if (-not $Force -and $pollState.LastAttemptAt) {
        try {
            if (($now - [datetime]$pollState.LastAttemptAt).TotalSeconds -lt $refreshIntervalSeconds) {
                return $false
            }
        } catch { }
    }

    $pollState.LastAttemptAt = $now
    $t = Invoke-TelnetCommand -Host $telnetHost -Port $telnetPort -Password $telnetPass -Command 'listplayers' -TimeoutMs 5000
    if (-not $t.Success) {
        $pollState.Source = '7DaysToDieTelnet'
        $pollState.Note = if ($t.Error) { "$($t.Error)" } else { '7 Days to Die telnet player query failed.' }
        return $false
    }

    $parsed = _Parse7DaysToDiePlayersResponse -ResponseText $t.Response
    if (-not $parsed.Available) {
        $recentLogText = _ReadRecentProfileLogText -Profile $Profile -Prefix $key -TailLines 200
        $logParsed = _Parse7DaysToDiePlayersLogText -Text $recentLogText
        if ($logParsed.Available) {
            $parsed = $logParsed
        }
    }
    $pollState.Source = '7DaysToDieTelnet'
    $pollState.Note = $parsed.Note
    if (-not $parsed.Available) { return $false }

    _SetLatestPlayersSnapshot -Prefix $key -Names @($parsed.Names) -Count ([int]$parsed.Count) -SharedState $SharedState
    $pollState.LastSuccessAt = $now
    return $true
}

function _RefreshHytalePlayers {
    param(
        [string]$Prefix,
        [hashtable]$Profile,
        [hashtable]$SharedState,
        [switch]$Force
    )

    if (-not $Profile -or -not $SharedState) { return $false }
    if (-not (_TestProfileGame -Profile $Profile -KnownGame 'Hytale')) { return $false }

    $key = $Prefix.ToUpperInvariant()
    Initialize-ServerManager -SharedState $SharedState

    if (-not $SharedState.RunningServers.ContainsKey($key)) { return $false }

    $now = Get-Date
    $pollState = $null
    if ($SharedState.PlayerQueryState.ContainsKey($key)) {
        $pollState = $SharedState.PlayerQueryState[$key]
    }
    if ($null -eq $pollState) {
        $pollState = [ordered]@{
            LastAttemptAt = $null
            LastSuccessAt = $null
            Source        = ''
            Note          = ''
        }
        $SharedState.PlayerQueryState[$key] = $pollState
    }

    $refreshIntervalSeconds = 25
    if (-not $Force -and $pollState.LastAttemptAt) {
        try {
            if (($now - [datetime]$pollState.LastAttemptAt).TotalSeconds -lt $refreshIntervalSeconds) {
                return $false
            }
        } catch { }
    }

    $pollState.LastAttemptAt = $now
    $sent = Send-ServerStdin -Prefix $key -Command 'who'
    if (-not $sent) {
        $pollState.Source = 'HytaleWhoLog'
        $pollState.Note = 'Could not send the Hytale who command.'
        return $false
    }

    Start-Sleep -Milliseconds 900
    $recentLogText = _ReadRecentProfileLogText -Profile $Profile -Prefix $key -TailLines 140
    $parsed = _ParseHytaleWhoText -Text $recentLogText
    $pollState.Source = 'HytaleWhoLog'
    $pollState.Note = $parsed.Note
    if (-not $parsed.Available) { return $false }

    _SetLatestPlayersSnapshot -Prefix $key -Names @($parsed.Names) -Count ([int]$parsed.Count) -SharedState $SharedState
    $pollState.LastSuccessAt = $now
    return $true
}

function _RequestProjectZomboidPlayersSnapshot {
    param(
        [string]$Prefix,
        [hashtable]$Profile,
        [hashtable]$SharedState,
        [switch]$Force
    )

    if (-not $Profile -or -not $SharedState) { return $false }
    if (-not (_TestProfileGame -Profile $Profile -KnownGame 'ProjectZomboid')) { return $false }

    $key = $Prefix.ToUpperInvariant()
    Initialize-ServerManager -SharedState $SharedState

    if (-not $SharedState.RunningServers.ContainsKey($key)) { return $false }
    if (-not $SharedState.ContainsKey('PlayersRequests') -or -not $SharedState.PlayersRequests) {
        $SharedState['PlayersRequests'] = [hashtable]::Synchronized(@{})
    }

    $now = Get-Date
    $pollState = $null
    if ($SharedState.PlayerQueryState.ContainsKey($key)) {
        $pollState = $SharedState.PlayerQueryState[$key]
    }
    if ($null -eq $pollState) {
        $pollState = [ordered]@{
            LastAttemptAt = $null
            LastSuccessAt = $null
            Source        = ''
            Note          = ''
        }
        $SharedState.PlayerQueryState[$key] = $pollState
    }

    $refreshIntervalSeconds = 25
    if (-not $Force -and $pollState.LastAttemptAt) {
        try {
            if (($now - [datetime]$pollState.LastAttemptAt).TotalSeconds -lt $refreshIntervalSeconds) {
                return $false
            }
        } catch { }
    }

    $pollState.LastAttemptAt = $now
    $pollState.Source = 'ProjectZomboidPlayersCommand'
    $pollState.Note = 'Requested Project Zomboid player snapshot.'

    try {
        $SharedState.PlayersRequests[$key] = @{
            Source      = 'IdleBootstrap'
            RequestedAt = $now
        }
    } catch { }

    $sent = Send-ServerStdin -Prefix $key -Command 'players'
    if (-not $sent) {
        $pollState.Note = 'Could not send the Project Zomboid players command.'
        return $false
    }

    return $true
}

function _GetPlayerObservationSnapshot {
    param(
        [string]$Prefix,
        [hashtable]$Profile,
        [hashtable]$SharedState
    )

    $key = $Prefix.ToUpperInvariant()
    $snapshot = [ordered]@{
        Supported = $false
        Available = $false
        Count     = 0
        Source    = ''
        Note      = ''
    }

    if ($null -eq $Profile -or -not $SharedState) { return $snapshot }

    if (_TestProfileGame -Profile $Profile -KnownGame 'ProjectZomboid') {
        $snapshot.Supported = $true
        if ($SharedState.ContainsKey('LatestPlayers') -and $SharedState.LatestPlayers.ContainsKey($key)) {
            $players = @($SharedState.LatestPlayers[$key])
            $snapshot.Available = $true
            $snapshot.Count = @($players).Count
            $snapshot.Source = 'ProjectZomboidLog'
            $snapshot.Note = 'Player count observed from Project Zomboid player capture.'
        } else {
            $requested = _RequestProjectZomboidPlayersSnapshot -Prefix $key -Profile $Profile -SharedState $SharedState
            if ($requested) {
                $snapshot.Note = 'Requested Project Zomboid player snapshot and waiting for log capture.'
            } else {
                $snapshot.Note = 'Waiting for Project Zomboid player capture.'
            }
        }
        _TracePlayerObservationDecision -Prefix $key -Profile $Profile -Snapshot $snapshot -SharedState $SharedState
        return $snapshot
    }

    if (_TestProfileGame -Profile $Profile -KnownGame '7DaysToDie') {
        $snapshot.Supported = $true
        if ($SharedState.ContainsKey('LatestPlayerCounts') -and $SharedState.LatestPlayerCounts.ContainsKey($key)) {
            $playerCount = 0
            try { $playerCount = [int]$SharedState.LatestPlayerCounts[$key] } catch { $playerCount = 0 }
            $snapshot.Available = $true
            $snapshot.Count = [Math]::Max(0, $playerCount)
            $snapshot.Source = '7DaysToDieTelnet'
            $snapshot.Note = 'Player count observed from 7 Days to Die telnet polling.'
        } else {
            $snapshot.Note = 'Waiting for 7 Days to Die telnet player polling.'
        }
        _TracePlayerObservationDecision -Prefix $key -Profile $Profile -Snapshot $snapshot -SharedState $SharedState
        return $snapshot
    }

    if (_TestProfileGame -Profile $Profile -KnownGame 'Hytale') {
        $snapshot.Supported = $true
        if ($SharedState.ContainsKey('LatestPlayerCounts') -and $SharedState.LatestPlayerCounts.ContainsKey($key)) {
            $playerCount = 0
            try { $playerCount = [int]$SharedState.LatestPlayerCounts[$key] } catch { $playerCount = 0 }
            $snapshot.Available = $true
            $snapshot.Count = [Math]::Max(0, $playerCount)
            $snapshot.Source = 'HytaleWhoLog'
            $snapshot.Note = 'Player count observed from Hytale who/log polling.'
        } else {
            $snapshot.Note = 'Waiting for Hytale who/log player polling.'
        }
        _TracePlayerObservationDecision -Prefix $key -Profile $Profile -Snapshot $snapshot -SharedState $SharedState
        return $snapshot
    }

    if (_TestProfileGame -Profile $Profile -KnownGame 'Palworld') {
        if ($null -eq $Profile.RestEnabled -or -not [bool]$Profile.RestEnabled) {
            $snapshot.Note = 'Palworld idle shutdown requires REST polling to be enabled.'
            _TracePlayerObservationDecision -Prefix $key -Profile $Profile -Snapshot $snapshot -SharedState $SharedState
            return $snapshot
        }

        $snapshot.Supported = $true
        try {
            if ($SharedState.ContainsKey('PalworldFeedState') -and $SharedState.PalworldFeedState.ContainsKey($key)) {
                $feed = $SharedState.PalworldFeedState[$key]
                $snapshot.Available = $true
                $snapshot.Count = [Math]::Max(0, [int]$feed.PlayerCount)
                $snapshot.Source = 'PalworldRest'
                $snapshot.Note = 'Player count observed from Palworld REST polling.'
                _TracePlayerObservationDecision -Prefix $key -Profile $Profile -Snapshot $snapshot -SharedState $SharedState
                return $snapshot
            }
        } catch { }

        try {
            $playersObj = $null
            if ($SharedState.ContainsKey('PalworldPlayersByPrefix') -and $SharedState.PalworldPlayersByPrefix -and $SharedState.PalworldPlayersByPrefix.ContainsKey($key)) {
                $playersObj = $SharedState.PalworldPlayersByPrefix[$key]
            }

            if ($null -ne $playersObj) {
                $playerCount = 0
                if ($playersObj -and $playersObj.PSObject.Properties['players']) {
                    $playerCount = @($playersObj.players).Count
                } elseif ($playersObj -is [System.Collections.IEnumerable] -and -not ($playersObj -is [string])) {
                    $playerCount = @($playersObj).Count
                }
                $snapshot.Available = $true
                $snapshot.Count = [Math]::Max(0, [int]$playerCount)
                $snapshot.Source = 'PalworldRest'
                $snapshot.Note = 'Player count observed from Palworld REST response.'
            } else {
                $snapshot.Note = 'Waiting for Palworld REST player data.'
            }
        } catch {
            $snapshot.Note = 'Palworld REST player data could not be parsed.'
        }
        _TracePlayerObservationDecision -Prefix $key -Profile $Profile -Snapshot $snapshot -SharedState $SharedState
        return $snapshot
    }

    if (_TestProfileGame -Profile $Profile -KnownGame 'Satisfactory') {
        $sfPort = 0
        try {
            if ($null -ne $Profile.SatisfactoryApiPort) {
                [void][int]::TryParse("$($Profile.SatisfactoryApiPort)", [ref]$sfPort)
            }
        } catch { $sfPort = 0 }

        if ($sfPort -le 0) {
            $snapshot.Note = 'Satisfactory idle shutdown requires API port configuration.'
            _TracePlayerObservationDecision -Prefix $key -Profile $Profile -Snapshot $snapshot -SharedState $SharedState
            return $snapshot
        }

        $snapshot.Supported = $true
        $sfHost  = if ($null -ne $Profile.SatisfactoryApiHost -and "$($Profile.SatisfactoryApiHost)".Trim() -ne '') { "$($Profile.SatisfactoryApiHost)".Trim() } else { '127.0.0.1' }
        $sfToken = if ($null -ne $Profile.SatisfactoryApiToken) { "$($Profile.SatisfactoryApiToken)" } else { '' }
        $cachedCount = 0
        $cachedObservedAt = $null
        try {
            if ($SharedState.ContainsKey('LatestPlayerCounts') -and $SharedState.LatestPlayerCounts.ContainsKey($key)) {
                $cachedCount = [Math]::Max(0, [int]$SharedState.LatestPlayerCounts[$key])
            }
        } catch { $cachedCount = 0 }
        try {
            if ($SharedState.ContainsKey('LatestPlayerObservedAt') -and $SharedState.LatestPlayerObservedAt.ContainsKey($key)) {
                $cachedObservedAt = [datetime]$SharedState.LatestPlayerObservedAt[$key]
            }
        } catch { $cachedObservedAt = $null }

        try {
            $r = Invoke-SatisfactoryApiRequest -Host $sfHost -Port $sfPort -Token $sfToken -Function 'QueryServerState' -TimeoutMs 5000
            if ($r.Success) {
                $players = 0
                try { $players = [int]$r.Data.serverGameState.numConnectedPlayers } catch { $players = 0 }
                $apiCount = [Math]::Max(0, $players)
                $useRecentLogFallback = $false
                if ($apiCount -le 0 -and $cachedCount -gt 0 -and $null -ne $cachedObservedAt) {
                    try {
                        $useRecentLogFallback = (((Get-Date) - $cachedObservedAt).TotalSeconds -le 180)
                    } catch { $useRecentLogFallback = $false }
                }

                $snapshot.Available = $true
                if ($useRecentLogFallback) {
                    $snapshot.Count = $cachedCount
                    $snapshot.Source = 'SatisfactoryLog'
                    $snapshot.Note = 'Recent Satisfactory join observed from server logs while API count catches up.'
                } else {
                    $snapshot.Count = $apiCount
                    $snapshot.Source = 'SatisfactoryApi'
                    $snapshot.Note = 'Player count observed from Satisfactory API.'
                }
            } else {
                $snapshot.Note = 'Waiting for Satisfactory API player data.'
            }
        } catch {
            $snapshot.Note = 'Satisfactory API player query failed.'
        }
        _TracePlayerObservationDecision -Prefix $key -Profile $Profile -Snapshot $snapshot -SharedState $SharedState
        return $snapshot
    }

    $gameLabel = ''
    try { $gameLabel = [string]$Profile.GameName } catch { $gameLabel = '' }
    if ([string]::IsNullOrWhiteSpace($gameLabel)) {
        try { $gameLabel = [string]$Profile.KnownGame } catch { $gameLabel = '' }
    }
    if ([string]::IsNullOrWhiteSpace($gameLabel)) {
        $gameLabel = 'This server'
    }

    if (_TestProfileGame -Profile $Profile -KnownGame 'Valheim') {
        $snapshot.Note = 'Valheim player detection is not trusted yet. Idle shutdown stays conservative for this server.'
    } elseif ($gameLabel -eq 'This server') {
        $snapshot.Note = 'Trusted player detection is not available for this server yet. Idle shutdown stays conservative.'
    } else {
        $snapshot.Note = "$gameLabel player detection is not trusted yet. Idle shutdown stays conservative for this server."
    }

    _TracePlayerObservationDecision -Prefix $key -Profile $Profile -Snapshot $snapshot -SharedState $SharedState
    return $snapshot
}

function _UpdatePlayerActivityState {
    param(
        [string]$Prefix,
        [hashtable]$Profile,
        [hashtable]$SharedState
    )

    if ([string]::IsNullOrWhiteSpace($Prefix) -or -not $SharedState) { return $null }
    Initialize-ServerManager -SharedState $SharedState

    $key = $Prefix.ToUpperInvariant()
    if (-not $SharedState.PlayerActivityState.ContainsKey($key)) {
        _ResetPlayerActivityTracking -Prefix $key -SharedState $SharedState
    }

    $state = $SharedState.PlayerActivityState[$key]
    if ($null -eq $state) {
        _ResetPlayerActivityTracking -Prefix $key -SharedState $SharedState
        $state = $SharedState.PlayerActivityState[$key]
    }

    try {
        if ($SharedState.RunningServers.ContainsKey($key) -and $SharedState.RunningServers[$key].StartTime) {
            $state.StartupAt = [datetime]$SharedState.RunningServers[$key].StartTime
        }
    } catch { }

    $snapshot = _GetPlayerObservationSnapshot -Prefix $key -Profile $Profile -SharedState $SharedState
    $state.DetectionSupported = [bool]$snapshot.Supported
    $state.DetectionAvailable = [bool]$snapshot.Available
    $state.DetectionSource = [string]$snapshot.Source
    $state.Note = [string]$snapshot.Note

    if ($snapshot.Available) {
        $now = Get-Date
        $currentCount = [Math]::Max(0, [int]$snapshot.Count)
        $previousCount = 0
        try { $previousCount = [int]$state.CurrentCount } catch { $previousCount = 0 }

        if ($currentCount -gt 0 -and $null -eq $state.FirstPlayerSeenAt) {
            $state.FirstPlayerSeenAt = $now
        }

        if ($currentCount -gt 0) {
            $state.LastNonZeroAt = $now
            $state.LastBecameEmptyAt = $null
        } elseif ($previousCount -gt 0) {
            $state.LastBecameEmptyAt = $now
        }

        $state.LastCount = $previousCount
        $state.CurrentCount = $currentCount
        $state.LastObservedAt = $now
    }

    return $state
}

function _CheckIdleShutdown {
    param(
        [string]$Prefix,
        [hashtable]$Profile,
        [hashtable]$SharedState
    )

    if (-not $Profile -or -not $SharedState -or -not $SharedState.RunningServers.ContainsKey($Prefix)) { return $false }

    $startupNoPlayersMinutes = _GetProfileIdleTimeoutMinutes -Profile $Profile -Key 'ShutdownIfNoPlayersAfterStartupMinutes'
    $emptyAfterLeaveMinutes = _GetProfileIdleTimeoutMinutes -Profile $Profile -Key 'ShutdownIfEmptyAfterLastPlayerLeavesMinutes'
    if ($startupNoPlayersMinutes -le 0 -and $emptyAfterLeaveMinutes -le 0) { return $false }

    $runtime = Get-ServerRuntimeState -Prefix $Prefix -SharedState $SharedState
    $runtimeCode = ''
    try { $runtimeCode = [string]$runtime.Code } catch { $runtimeCode = '' }
    $runtimeCode = $runtimeCode.ToLowerInvariant()
    if ($runtimeCode -in @('starting','stopping','restarting','waiting_restart','blocked','failed','startup_failed')) {
        return $false
    }

    $startupAt = Get-Date
    try {
        if ($SharedState.RunningServers.ContainsKey($Prefix) -and $SharedState.RunningServers[$Prefix].StartTime) {
            $startupAt = [datetime]$SharedState.RunningServers[$Prefix].StartTime
        }
    } catch { }

    $activity = _UpdatePlayerActivityState -Prefix $Prefix -Profile $Profile -SharedState $SharedState
    if ($null -eq $activity) { return $false }
    if (-not $activity.DetectionSupported) { return $false }

    $previousPendingRule = ''
    $previousShutdownDueAt = $null
    try { $previousPendingRule = [string]$activity.PendingRule } catch { $previousPendingRule = '' }
    try { $previousShutdownDueAt = $activity.ShutdownDueAt } catch { $previousShutdownDueAt = $null }

    if (-not $activity.DetectionAvailable) {
        $activity.PendingRule = 'signal_wait'
        $signalState = 'waiting_first_player'
        $signalDueAt = $null

        if ($startupNoPlayersMinutes -gt 0 -and $null -eq $activity.FirstPlayerSeenAt) {
            $signalState = 'waiting_first_player'
            $signalDueAt = $startupAt.AddMinutes($startupNoPlayersMinutes)
        } elseif ($emptyAfterLeaveMinutes -gt 0 -and $null -ne $activity.FirstPlayerSeenAt -and $null -ne $activity.LastBecameEmptyAt) {
            $signalState = 'idle_wait'
            $signalDueAt = ([datetime]$activity.LastBecameEmptyAt).AddMinutes($emptyAfterLeaveMinutes)
        }

        $activity.ShutdownDueAt = $signalDueAt
        _TraceServerTimerDecision -Prefix $Prefix -Timer 'idle_shutdown' -OldRule $previousPendingRule -OldDueAt $previousShutdownDueAt -NewRule $activity.PendingRule -NewDueAt $activity.ShutdownDueAt -Reason 'waiting for trusted live player data before arming a concrete idle rule' -SharedState $SharedState

        if ($runtimeCode -in @('online','waiting_first_player','idle_wait')) {
            $detail = _FormatSignalWaitRuntimeDetail -StateCode $signalState -DueAt $signalDueAt
            Set-ServerRuntimeState -Prefix $Prefix -State $signalState -Detail $detail -SharedState $SharedState
        }
        return $false
    }

    $currentCount = 0
    try { $currentCount = [int]$activity.CurrentCount } catch { $currentCount = 0 }
    if ($currentCount -gt 0) {
        $activity.PendingRule = ''
        $activity.ShutdownDueAt = $null
        _TraceServerTimerDecision -Prefix $Prefix -Timer 'idle_shutdown' -OldRule $previousPendingRule -OldDueAt $previousShutdownDueAt -NewRule $activity.PendingRule -NewDueAt $activity.ShutdownDueAt -Reason 'players detected; idle shutdown timer cleared' -SharedState $SharedState
        Set-ServerRuntimeState -Prefix $Prefix -State 'online' -Detail ("{0} player(s) online." -f $currentCount) -SharedState $SharedState
        return $false
    }

    $key = $Prefix.ToUpperInvariant()
    $gameName = if ($Profile.GameName) { $Profile.GameName } else { $key }
    try {
        if ($activity.StartupAt) {
            $startupAt = [datetime]$activity.StartupAt
        }
    } catch { }

    if ($startupNoPlayersMinutes -gt 0 -and $null -eq $activity.FirstPlayerSeenAt) {
        $dueAt = $startupAt.AddMinutes($startupNoPlayersMinutes)
        $activity.PendingRule = 'first_player'
        $activity.ShutdownDueAt = $dueAt

        if ((Get-Date) -ge $dueAt) {
            _TraceServerTimerDecision -Prefix $key -Timer 'idle_shutdown' -OldRule $previousPendingRule -OldDueAt $previousShutdownDueAt -NewRule $activity.PendingRule -NewDueAt $activity.ShutdownDueAt -Reason 'first-player idle timeout reached' -SharedState $SharedState -Action 'expire'
            $reason = "No players joined within ${startupNoPlayersMinutes} minute(s) after startup."
            $notice = _JoinNotificationParts -Parts @(
                'Idle shutdown triggered',
                'No players joined after startup',
                ("Timeout reached after {0} minute(s)." -f $startupNoPlayersMinutes)
            )
            Set-ServerRuntimeState -Prefix $key -State 'idle_shutdown' -Detail $reason -SharedState $SharedState
            _Log "[$gameName] $notice" -Level WARN
            _GameLog -Prefix $key -Line ("[WARN] {0}" -f $notice)
            _Webhook (New-DiscordGameMessage -Profile $Profile -Prefix $key -Event 'idle_shutdown' -Values @{
                Reason = "No players joined after startup. Timeout reached after $startupNoPlayersMinutes minute(s). Safe shutdown beginning now."
            })
            try {
                Invoke-SafeShutdown -Prefix $key -Quiet | Out-Null
            } catch {
                _LogProfileError -Prefix $key -Profile $Profile -Action 'idle-shutdown-stop' -Function '_CheckIdleShutdown' -Message 'Idle shutdown stop sequence failed during first-player timeout handling.' -ErrorRecord $_ -Fallback 'Idle timeout handling will continue with the current runtime update and be re-evaluated on the next monitor tick.' -Recovery 'deferred to next monitor tick; shutdown not yet confirmed' -Level WARN
            }
            Set-ServerRuntimeState -Prefix $key -State 'stopped' -Detail ("Idle shutdown complete. {0}" -f $reason) -SharedState $SharedState
            $activity.PendingRule = ''
            $activity.ShutdownDueAt = $null
            return $true
        }

        _TraceServerTimerDecision -Prefix $key -Timer 'idle_shutdown' -OldRule $previousPendingRule -OldDueAt $previousShutdownDueAt -NewRule $activity.PendingRule -NewDueAt $activity.ShutdownDueAt -Reason 'arming first-player idle window after startup' -SharedState $SharedState
        Set-ServerRuntimeState -Prefix $key -State 'waiting_first_player' -Detail ("Waiting for first player. Idle shutdown in {0}." -f (_FormatCountdownShort -DueAt $dueAt)) -SharedState $SharedState
        return $false
    }

    if ($emptyAfterLeaveMinutes -gt 0 -and $null -ne $activity.FirstPlayerSeenAt) {
        if ($null -eq $activity.LastBecameEmptyAt) {
            $activity.LastBecameEmptyAt = Get-Date
        }

        $dueAt = ([datetime]$activity.LastBecameEmptyAt).AddMinutes($emptyAfterLeaveMinutes)
        $activity.PendingRule = 'empty_after_leave'
        $activity.ShutdownDueAt = $dueAt

        if ((Get-Date) -ge $dueAt) {
            _TraceServerTimerDecision -Prefix $key -Timer 'idle_shutdown' -OldRule $previousPendingRule -OldDueAt $previousShutdownDueAt -NewRule $activity.PendingRule -NewDueAt $activity.ShutdownDueAt -Reason 'empty-server idle timeout reached' -SharedState $SharedState -Action 'expire'
            $reason = "Server stayed empty for ${emptyAfterLeaveMinutes} minute(s) after the last player left."
            $notice = _JoinNotificationParts -Parts @(
                'Idle shutdown triggered',
                'Server stayed empty after the last player left',
                ("Timeout reached after {0} minute(s)." -f $emptyAfterLeaveMinutes)
            )
            Set-ServerRuntimeState -Prefix $key -State 'idle_shutdown' -Detail $reason -SharedState $SharedState
            _Log "[$gameName] $notice" -Level WARN
            _GameLog -Prefix $key -Line ("[WARN] {0}" -f $notice)
            _Webhook (New-DiscordGameMessage -Profile $Profile -Prefix $key -Event 'idle_shutdown' -Values @{
                Reason = "Server stayed empty after the last player left. Timeout reached after $emptyAfterLeaveMinutes minute(s). Safe shutdown beginning now."
            })
            try {
                Invoke-SafeShutdown -Prefix $key -Quiet | Out-Null
            } catch {
                _LogProfileError -Prefix $key -Profile $Profile -Action 'idle-shutdown-stop' -Function '_CheckIdleShutdown' -Message 'Idle shutdown stop sequence failed during empty-server timeout handling.' -ErrorRecord $_ -Fallback 'Idle timeout handling will continue with the current runtime update and be re-evaluated on the next monitor tick.' -Recovery 'deferred to next monitor tick; shutdown not yet confirmed' -Level WARN
            }
            Set-ServerRuntimeState -Prefix $key -State 'stopped' -Detail ("Idle shutdown complete. {0}" -f $reason) -SharedState $SharedState
            $activity.PendingRule = ''
            $activity.ShutdownDueAt = $null
            return $true
        }

        _TraceServerTimerDecision -Prefix $key -Timer 'idle_shutdown' -OldRule $previousPendingRule -OldDueAt $previousShutdownDueAt -NewRule $activity.PendingRule -NewDueAt $activity.ShutdownDueAt -Reason 'arming empty-server idle window after the last player left' -SharedState $SharedState
        Set-ServerRuntimeState -Prefix $key -State 'idle_wait' -Detail ("Server empty. Idle shutdown in {0}." -f (_FormatCountdownShort -DueAt $dueAt)) -SharedState $SharedState
        return $false
    }

    return $false
}

function _GetMonitorSlowProfileThresholdMs {
    param(
        [object]$Profile = $null
    )

    $defaultMs = 1500
    $default7DaysMs = 6000
    try {
        if ($script:State -and $script:State.Settings -and $script:State.Settings.ContainsKey('MonitorSlowProfileThresholdMs')) {
            $value = [int]$script:State.Settings.MonitorSlowProfileThresholdMs
            if ($value -gt 0) { return $value }
        }
    } catch { }

    try {
        if ($script:State -and $script:State.Settings -and $script:State.Settings.ContainsKey('MonitorSlowProfileThresholdMs7DaysToDie')) {
            $value = [int]$script:State.Settings.MonitorSlowProfileThresholdMs7DaysToDie
            if ($value -gt 0) { return $value }
        }
    } catch { }

    try {
        if ($Profile -and (_TestProfileGame -Profile $Profile -KnownGame '7DaysToDie')) {
            return $default7DaysMs
        }
    } catch { }

    return $defaultMs
}

function _LogMonitorProfileDuration {
    param(
        [string]$Prefix,
        [object]$Profile,
        [int]$ElapsedMs
    )

    $gameName = ''
    try { $gameName = _CoalesceStr $Profile.GameName $Prefix } catch { $gameName = $Prefix }

    _Log "Monitor profile [$Prefix][$gameName] ${ElapsedMs}ms" -Level DEBUG

    $slowThresholdMs = _GetMonitorSlowProfileThresholdMs -Profile $Profile
    if ($ElapsedMs -ge $slowThresholdMs) {
        _LogThrottled -Key "MonitorSlowProfile::${Prefix}" -Msg "Monitor profile [$Prefix][$gameName] was slow: ${ElapsedMs}ms (threshold ${slowThresholdMs}ms)." -Level WARN -WindowSeconds 30
    }
}

function _GetNotificationGameName {
    param(
        [object]$Profile,
        [string]$Prefix = ''
    )

    try {
        if ($Profile -and $Profile.GameName -and "$($Profile.GameName)".Trim() -ne '') {
            return "$($Profile.GameName)".Trim()
        }
    } catch { }

    if (-not [string]::IsNullOrWhiteSpace($Prefix)) {
        return $Prefix.ToUpperInvariant()
    }

    return 'Server'
}

function _JoinNotificationParts {
    param([object[]]$Parts)

    $items = New-Object 'System.Collections.Generic.List[string]'
    foreach ($part in @($Parts)) {
        if ($null -eq $part) { continue }
        $text = "$part".Trim()
        if (-not [string]::IsNullOrWhiteSpace($text)) {
            $items.Add($text) | Out-Null
        }
    }

    return ($items -join ' | ')
}

function _GetDiscordCommandContextEntry {
    param(
        [string]$Prefix,
        [hashtable]$SharedState = $script:State
    )

    if ([string]::IsNullOrWhiteSpace($Prefix) -or -not $SharedState -or -not $SharedState.ContainsKey('DiscordCommandContext')) {
        return $null
    }

    $key = $Prefix.ToUpperInvariant()
    if (-not $SharedState.DiscordCommandContext.ContainsKey($key)) {
        return $null
    }

    $entry = $SharedState.DiscordCommandContext[$key]
    if ($null -eq $entry) {
        $SharedState.DiscordCommandContext.Remove($key)
        return $null
    }

    try {
        if ($entry.ExpiresAt -and ((Get-Date) -ge [datetime]$entry.ExpiresAt)) {
            $SharedState.DiscordCommandContext.Remove($key)
            return $null
        }
    } catch {
        $SharedState.DiscordCommandContext.Remove($key)
        return $null
    }

    return $entry
}

function _ClearDiscordCommandContext {
    param(
        [string]$Prefix,
        [hashtable]$SharedState = $script:State
    )

    if ([string]::IsNullOrWhiteSpace($Prefix) -or -not $SharedState -or -not $SharedState.ContainsKey('DiscordCommandContext')) {
        return
    }

    $key = $Prefix.ToUpperInvariant()
    if ($SharedState.DiscordCommandContext.ContainsKey($key)) {
        $SharedState.DiscordCommandContext.Remove($key)
    }
}

function _ShouldSuppressDiscordLifecycleWebhook {
    param(
        [string]$Prefix,
        [string]$Tag,
        [hashtable]$SharedState = $script:State
    )

    $normalizedTag = if ([string]::IsNullOrWhiteSpace($Tag)) { '' } else { $Tag.Trim().ToUpperInvariant() }
    $prefixKey = if ([string]::IsNullOrWhiteSpace($Prefix)) { '' } else { $Prefix.Trim().ToUpperInvariant() }

    if ([string]::IsNullOrWhiteSpace($normalizedTag) -or [string]::IsNullOrWhiteSpace($prefixKey) -or -not $SharedState) {
        return $false
    }

    if ($normalizedTag -in @('ONLINE','JOINABLE')) {
        $runtimeCode = ''
        $shutdownNow = $false
        $isRunningNow = $false

        try {
            $runtime = Get-ServerRuntimeState -Prefix $prefixKey -SharedState $SharedState
            if ($runtime) { $runtimeCode = [string]$runtime.Code }
        } catch { $runtimeCode = '' }

        try {
            if ($SharedState.ShutdownFlags -and $SharedState.ShutdownFlags.ContainsKey($prefixKey)) {
                $shutdownNow = [bool]$SharedState.ShutdownFlags[$prefixKey]
            }
        } catch { $shutdownNow = $false }

        try {
            if ($SharedState.RunningServers -and $SharedState.RunningServers.ContainsKey($prefixKey)) {
                $srvEntry = $SharedState.RunningServers[$prefixKey]
                $srvPid = $null
                try { $srvPid = $srvEntry.Pid } catch { $srvPid = $null }
                if ($srvPid) {
                    $liveProcess = $null
                    try { $liveProcess = Get-Process -Id $srvPid -ErrorAction Stop } catch { $liveProcess = $null }
                    $isRunningNow = ($liveProcess -and -not $liveProcess.HasExited)
                }
            }
        } catch { $isRunningNow = $false }

        $runtimeCode = $runtimeCode.ToLowerInvariant()
        if ($shutdownNow) { return $true }
        if ($runtimeCode -in @('stopping','stopped','idle_shutdown','failed','blocked','startup_failed')) { return $true }
        if (-not $isRunningNow) { return $true }
    }

    $entry = _GetDiscordCommandContextEntry -Prefix $prefixKey -SharedState $SharedState
    if ($null -eq $entry) { return $false }

    $command = ''
    try { $command = "$($entry.Command)".Trim().ToLowerInvariant() } catch { $command = '' }

    switch ($command) {
        'start' {
            return ($normalizedTag -in @('STARTING'))
        }
        'stop' {
            return ($normalizedTag -in @('SAVING', 'WAITING'))
        }
        'restart' {
            return ($normalizedTag -in @('RESTARTING', 'STARTING', 'ONLINE'))
        }
    }

    return $false
}

function Send-DiscordGameEvent {
    param(
        [hashtable]$Profile,
        [string]$Prefix,
        [string]$Event,
        [hashtable]$Values = $null,
        [string]$Tag = '',
        [hashtable]$SharedState = $script:State
    )

    if ([string]::IsNullOrWhiteSpace($Event)) { return $false }

    $normalizedEvent = $Event.Trim().ToLowerInvariant()
    $resolvedTag = if ([string]::IsNullOrWhiteSpace($Tag)) { '' } else { $Tag.Trim().ToUpperInvariant() }

    if ([string]::IsNullOrWhiteSpace($resolvedTag)) {
        switch ($normalizedEvent) {
            'starting' { $resolvedTag = 'STARTING' }
            'online' { $resolvedTag = 'ONLINE' }
            'joinable' { $resolvedTag = 'JOINABLE' }
            'restarting' { $resolvedTag = 'RESTARTING' }
            'restarted_auto' { $resolvedTag = 'ONLINE' }
            'restarted' { $resolvedTag = 'ONLINE' }
            'scheduled_restart_done' { $resolvedTag = 'ONLINE' }
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($resolvedTag) -and (_ShouldSuppressDiscordLifecycleWebhook -Prefix $Prefix -Tag $resolvedTag -SharedState $SharedState)) {
        _TraceDiscordSendDecision -Prefix $Prefix -Profile $Profile -Event $Event -Tag $resolvedTag -Action 'suppress' -Reason 'game event was suppressed by the shared lifecycle guard' -SharedState $SharedState
        return $false
    }

    if ($normalizedEvent -in @('players_none','players_list')) {
        $runtimeCode = ''
        $isRunningNow = $false
        try {
            $runtime = Get-ServerRuntimeState -Prefix $Prefix -SharedState $SharedState
            if ($runtime) { $runtimeCode = [string]$runtime.Code }
        } catch { $runtimeCode = '' }

        try {
            if ($SharedState -and $SharedState.RunningServers -and $SharedState.RunningServers.ContainsKey($Prefix.ToUpperInvariant())) {
                $srvEntry = $SharedState.RunningServers[$Prefix.ToUpperInvariant()]
                $srvPid = $null
                try { $srvPid = $srvEntry.Pid } catch { $srvPid = $null }
                if ($srvPid) {
                    $liveProcess = $null
                    try { $liveProcess = Get-Process -Id $srvPid -ErrorAction Stop } catch { $liveProcess = $null }
                    $isRunningNow = ($liveProcess -and -not $liveProcess.HasExited)
                }
            }
        } catch { $isRunningNow = $false }

        $runtimeCode = $runtimeCode.ToLowerInvariant()
        if ((-not $isRunningNow) -or $runtimeCode -in @('stopping','stopped','idle_shutdown','failed','blocked','startup_failed')) {
            _TraceDiscordSendDecision -Prefix $Prefix -Profile $Profile -Event $Event -Action 'suppress' -Reason 'player roster event was skipped because the server was not in a live running state' -SharedState $SharedState
            return $false
        }
    }

    $message = if ($null -ne $Values) {
        New-DiscordGameMessage -Profile $Profile -Prefix $Prefix -Event $Event -Values $Values
    } else {
        New-DiscordGameMessage -Profile $Profile -Prefix $Prefix -Event $Event
    }

    _TraceDiscordSendDecision -Prefix $Prefix -Profile $Profile -Event $Event -Tag $resolvedTag -Action 'queue' -Reason 'game event queued for Discord delivery' -Message $message -SharedState $SharedState
    _Webhook -Message $message -SkipDebugTrace
    return $true
}

function _SendDiscordLifecycleWebhook {
    param(
        [hashtable]$Profile,
        [string]$Prefix,
        [string]$Event,
        [hashtable]$Values = $null,
        [string]$Tag = '',
        [hashtable]$SharedState = $script:State
    )

    if ([string]::IsNullOrWhiteSpace($Event)) { return $false }

    $resolvedTag = if ([string]::IsNullOrWhiteSpace($Tag)) { '' } else { $Tag.Trim().ToUpperInvariant() }
    if ([string]::IsNullOrWhiteSpace($resolvedTag)) {
        switch ($Event.ToLowerInvariant()) {
            'online' { $resolvedTag = 'ONLINE' }
            'joinable' { $resolvedTag = 'JOINABLE' }
            'restarted_auto' { $resolvedTag = 'ONLINE' }
            'restarted' { $resolvedTag = 'ONLINE' }
            'scheduled_restart_done' { $resolvedTag = 'ONLINE' }
        }
    }

    return (Send-DiscordGameEvent -Profile $Profile -Prefix $Prefix -Event $Event -Values $Values -Tag $resolvedTag -SharedState $SharedState)
}

function _GetDiscordFlavorData {
    param(
        [object]$Profile,
        [string]$Prefix = ''
    )

    $generic = [ordered]@{
        StartLead   = @(
            'The server is spinning up.',
            'Power is coming up on the server.',
            'The boot sequence is rolling now.',
            'The world is stretching awake again.',
            'The server is coming back around.',
            'The stack is waking up one clean step at a time.',
            'The rack is humming back to life.',
            'The world is rubbing the sleep out of its eyes.'
        )
        ReadyLead   = @(
            'Everything looks stable.',
            'The board is green again.',
            'The lights are on and steady.',
            'The server settled in cleanly.',
            'Everything is holding together.',
            'The gauges are behaving for once.',
            'The bunker lights are steady and not flickering.',
            'The whole thing is standing up straighter now.'
        )
        JoinableTip = @(
            'Jump in when you are ready.',
            'The door is open whenever you are.',
            'Looks clear for a drop-in.',
            'The session is ready for company.',
            'The server is waiting on the crew.',
            'The gate is unlocked and the room is warm.',
            'The place is open if the crew wants in.',
            'The server is idling by the door for company.'
        )
        SaveLead    = @(
            'Locking in the latest world state.',
            'Banking the freshest progress.',
            'Tucking the last good snapshot away.',
            'Pinning down the newest save state.',
            'Packing away the latest chapter safely.',
            'Sealing up the newest state before it wanders off.',
            'Stacking the latest progress where it cannot get lost.',
            'Putting the freshest snapshot somewhere safe and dry.'
        )
        WaitLead    = @(
            'Giving it a moment to settle.',
            'Letting things cool off for a second.',
            'Holding for one clean beat.',
            'Giving the server a quiet minute.',
            'Letting the dust come down first.',
            'Giving the gears one last easy spin.',
            'Leaving a little quiet between the noise and the dark.',
            'Letting everything stop rattling before the next move.'
        )
        StopLead    = @(
            'The lights are going out for now.',
            'The server is going quiet for a while.',
            'Things are winding down cleanly.',
            'The world is closing up for the night.',
            'The room is going still for now.',
            'The doors are closing and the hum is fading.',
            'The world is folding itself up without a fight.',
            'The last light in the room is being flipped off.'
        )
        RestartLead = @(
            'Running a clean power cycle.',
            'Taking one clean lap through a reboot.',
            'Cycling everything without cutting corners.',
            'Giving the server a clean reset.',
            'Taking the long way around a reboot.',
            'Running the server around the circuit and back again.',
            'Pulling it through one tidy reset loop.',
            'Taking the stack apart just long enough to put it back together.'
        )
        CrashLead   = @(
            'Recovery is already spinning up.',
            'ECC is already lining up the comeback.',
            'The reboot plan is already in motion.',
            'Recovery is on the board already.',
            'ECC is already chasing the bounce-back.'
        )
        OperatorCrashLead = @(
            'I am already under the console kicking relays before the overlords smell smoke.',
            'I am patching this mess together fast enough to stay off a corporate incident slide.',
            'I am dragging the bunker rack through recovery whether it likes it or not.',
            'I am trying to make this look controlled before somebody upstairs starts asking for names.',
            'I am elbow-deep in the restart plan, because apparently panic is part of the job.',
            'I am working the recovery board now, mostly because I enjoy continued breathing privileges.',
            'I am forcing the bunker through triage before corporate decides this was my personality.',
            'I am already catching the falling pieces, so maybe the overlords do not have to.'
        )
        OperatorStandDownLead = @(
            'I am backing away before corporate decides this crater belongs to me.',
            'I am not feeding another body into the rack until somebody explains this to the overlords.',
            'I am standing the recovery crew down before this turns into a formal execution of morale.',
            'I am calling a halt before the bunker burns another shift on false hope.',
            'I am taking my hands off the restart lever until somebody smarter or meaner signs for it.',
            'I am stopping here, because even the overlords will notice this many wrecks.'
        )
        OperatorBlockedLead = @(
            'I am not pushing that launch through until somebody upstairs signs the blame form.',
            'I am holding that line shut before corporate decides I ignored a red light.',
            'I am keeping the bunker gate closed until this stops looking like a career-ending idea.',
            'I am parking this one on the sideline until the overlords can complain in writing.',
            'I am not feeding that mess into live power just to impress management.',
            'I am leaving the launch lever alone until this stops smelling like paperwork and death.'
        )
        OperatorStartupFailLead = @(
            'I am staring at the boot wreckage and pretending this is routine.',
            'I am sweeping up the failed startup before corporate asks who touched what.',
            'I am already under the console trying to make this look less embarrassing.',
            'I am holding the bunker together with retries and bad language.',
            'I am reading the startup fallout and trying not to become part of it.',
            'I am doing bunker triage now, mostly because the overlords dislike surprises.'
        )
        OperatorRestartFailLead = @(
            'I am dragging the restart plan back onto the table before corporate smells weakness.',
            'I am reworking the reboot mess while pretending this was always the backup plan.',
            'I am catching the failed restart by hand, which is exactly as fun as it sounds.',
            'I am still in the bunker trying to convince the rack to cooperate.',
            'I am piecing the reboot attempt back together before somebody upstairs starts pointing.',
            'I am rewriting the recovery story in real time so the overlords do not write it for me.'
        )
        OperatorOrderLead = @(
            'Fine. I heard the order and I am moving the bunker machinery around again.',
            'Request logged. I will go make it look like this was always the plan.',
            'The order is on the board now, so I guess I work for another minute.',
            'Message received. I am already rearranging the relays to keep corporate content.',
            'The request is in my hands now, which feels unfair but familiar.',
            'I caught the order before the overlords could accuse me of sleeping.'
        )
        OperatorStatusLead = @(
            'I am reading the bunker gauges now, since apparently everyone wants telemetry on demand.',
            'I am checking the floor readout before corporate asks for it in a worse tone.',
            'I am pulling the latest bunker pulse off the board, because somebody always wants a number.',
            'I am staring at the status lights so you do not have to crawl in here.',
            'I am reading the room for you while the overlords pretend this is a calm workplace.',
            'I am fetching the latest board state before somebody upstairs decides silence is suspicious.'
        )
        OperatorCommandLead = @(
            'I am shoving that down the line now and hoping the rack does not bite me back.',
            'I am feeding that straight into the bunker controls because apparently caution is optional today.',
            'I am pushing the command through before corporate mistakes hesitation for rebellion.',
            'I am sending it downrange now, under protest and fluorescent lighting.',
            'I am handing that order to the machinery and pretending this is fine.',
            'I am already leaning on the control desk to make it happen.'
        )
        Caution     = @(
            'Keep an eye on the logs.',
            'Stay close to the console for a minute.',
            'Might be worth watching the next few lines.',
            'Keep the dashboard in view.',
            'This one deserves a quick log check.'
        )
        StatusLead  = @(
            'Here is the current server pulse.',
            'This is how the server looks right now.',
            'Current server readout is in.',
            'This is the latest board state.',
            'Here is the live status picture.'
        )
        PlayerRole  = 'crew'
        WorldLabel  = 'server'
        DangerLabel = 'trouble'
    }

    $gameKey = _NormalizeGameIdentity (_GetProfileKnownGame -Profile $Profile)
    $overrides = [ordered]@{}

    switch ($gameKey) {
        '7daystodie' {
            $overrides = [ordered]@{
                StartLead   = @(
                    'Dust is blowing across the wasteland.',
                    'The traders are opening the shutters again.',
                    'Campfires are catching in the wasteland.',
                    'The road back into Navezgane is clearing.',
                    'The bedrolls and blood moons are back on the calendar.',
                    'The generators are coughing awake behind the walls.',
                    'The dead have another bad day ahead of them.'
                )
                ReadyLead   = @(
                    'The dead are shambling again.',
                    'The wasteland is breathing like it means it.',
                    'The horde has somewhere to wander again.',
                    'The county is open and hostile again.',
                    'The forge smoke is rising over the wasteland.',
                    'The county smells like trouble again.',
                    'The watchtower finally has something to watch.'
                )
                JoinableTip = @(
                    'Stay sharp and try not to wake the whole county.',
                    'Grab your bedroll and watch the treeline.',
                    'Mind the dogs and save a few rounds for horde night.',
                    'Maybe hit the trader before sunset.',
                    'Do not let the screamers hear you coming.',
                    'Bring a wrench and an exit plan.',
                    'If you hear ferals, you were too loud.'
                )
                SaveLead    = @(
                    'Packing away the survivor stash.',
                    'Locking the forge ledger and camp chests down.',
                    'Banking the last good scavenging run.',
                    'Tucking the horde prep away safely.',
                    'Sealing up the latest wasteland haul.',
                    'Counting the shells and sealing the crate lids.',
                    'Locking the night''s haul behind steel and concrete.'
                )
                WaitLead    = @(
                    'Boarding up the shelter before lights out.',
                    'Letting the dew collectors drip out.',
                    'Giving the base one last quiet sweep.',
                    'Waiting for the forge heat to die down.',
                    'Holding a beat before the bunker goes dark.',
                    'Letting the spikes stop singing first.',
                    'Giving the watchtower one last slow turn.'
                )
                StopLead    = @(
                    'The wasteland is going quiet.',
                    'The bunker lights are going dark.',
                    'The county is settling back into the dead air.',
                    'The trader route is closing up for now.',
                    'The safehouse is going still.',
                    'The watchfire is burning low again.',
                    'The wasteland can chew on silence for a while.'
                )
                RestartLead = @(
                    'Barricades are coming down and going right back up.',
                    'Cycling the bunker power with the horde still out there.',
                    'Giving the wasteland one clean reset.',
                    'Running the county through a fast rebuild.',
                    'Taking the shelter through a hard reboot.',
                    'Shaking the dust out of the bunker and firing it right back up.',
                    'Running the blood moon clock back through one clean turn.'
                )
                CrashLead   = @(
                    'Something hit the horde alarm early.',
                    'A screamer probably found the wrong switch.',
                    'Something ugly kicked the shelter door in.',
                    'The wasteland took a swing at the server.',
                    'A bad moon hit the system early.'
                )
                Caution     = @(
                    'Keep your shotgun close.',
                    'Maybe count your ammo and the error lines.',
                    'Stay near the bunker radio for this one.',
                    'Keep one eye on the horde and one on the logs.',
                    'Watch the board like it is horde night.'
                )
                StatusLead  = @(
                    'The wasteland report is in.',
                    'The Navezgane board looks like this.',
                    'The survivor radio check says this.',
                    'The county update reads like this.',
                    'The horde tracker is showing this.'
                )
                PlayerRole  = 'survivors'
                WorldLabel  = 'wasteland'
                DangerLabel = 'horde'
            }
        }
        'hytale' {
            $overrides = [ordered]@{
                StartLead   = @(
                    'The shard is waking up.',
                    'The portal ring is starting to glow again.',
                    'Adventure is stirring behind the gates.',
                    'The realm is taking its first breath again.',
                    'The world shard is coming back into tune.',
                    'The adventure board is lighting up again.',
                    'Somewhere, a gatekeeper is regretting opening the portal.'
                )
                ReadyLead   = @(
                    'Adventure is stirring beyond the portal.',
                    'The realm is stable and humming again.',
                    'The shard is lit and ready for boots on stone.',
                    'The portal air is clear again.',
                    'The wilds beyond the gate are open again.',
                    'The camp beyond the gate is ready for footsteps.',
                    'The realm is holding its shape nicely for once.'
                )
                JoinableTip = @(
                    'Watch your footing out there.',
                    'Mind the cliffs, the portals, and your stamina.',
                    'Keep the campfire warm and your pockets light.',
                    'Do not trust every ruin with a glow to it.',
                    'Watch the skyline and the path stones.',
                    'Pack light and trust the map stones only a little.',
                    'If the ruins glow back at you, maybe think twice.'
                )
                SaveLead    = @(
                    'Tucking away the realm state.',
                    'Banking the latest shard imprint.',
                    'Packing the last clean world memory away.',
                    'Locking the portal ledger into place.',
                    'Stashing the freshest realm snapshot.',
                    'Folding the latest shard echo into the archive.',
                    'Pinning the newest portal memory behind glass.'
                )
                WaitLead    = @(
                    'Letting the shard settle for a moment.',
                    'Giving the portal ring a second to cool.',
                    'Holding while the realm quiets down.',
                    'Letting the beacon glow fade a little first.',
                    'Giving the wilds one calm breath.',
                    'Giving the camp lanterns a second to stop swaying.',
                    'Letting the realm settle its dust and magic.'
                )
                StopLead    = @(
                    'The shard is dimming for now.',
                    'The portal light is fading out.',
                    'The realm is going still behind the gate.',
                    'The camp beyond the portal is closing up.',
                    'The world shard is slipping into quiet.',
                    'The gate is closing with the lanterns still warm.',
                    'The shard is slipping back under the surface.'
                )
                RestartLead = @(
                    'The portal is cycling cleanly.',
                    'The shard is taking one clean pass through the gate.',
                    'Running the realm through a clean retune.',
                    'The gate is closing and reopening in one motion.',
                    'The portal ring is taking a fresh charge.',
                    'Turning the portal ring over and setting it back on its feet.',
                    'Running the shard through a fresh tuning pass.'
                )
                CrashLead   = @(
                    'The shard slipped out of tune.',
                    'The portal coughed and dropped the line.',
                    'Something in the realm knocked the gate sideways.',
                    'The shard lost its footing for a second.',
                    'The portal band went unstable.'
                )
                Caution     = @(
                    'Keep one eye on the portal.',
                    'Watch the gate and the next few console lines.',
                    'Stay close to the shard readout for a minute.',
                    'Maybe keep the realm monitor open.',
                    'This one is worth a second look at the portal logs.'
                )
                StatusLead  = @(
                    'The shard pulse looks like this.',
                    'The realm readout says this.',
                    'The portal board is showing this.',
                    'The latest shard telemetry is here.',
                    'The gate report reads like this.'
                )
                PlayerRole  = 'adventurers'
                WorldLabel  = 'shard'
                DangerLabel = 'rift'
            }
        }
        'palworld' {
            $overrides = [ordered]@{
                StartLead   = @(
                    'The Palbox is humming to life.',
                    'The island outpost is coming back online.',
                    'The ranch lights are flicking back on.',
                    'The Pal spheres are rolling again.',
                    'The breeding pens and workbenches are waking up.',
                    'The feed bins are rattling and the benches are waking up.',
                    'The island shift whistle just blew.'
                )
                ReadyLead   = @(
                    'The island is buzzing with Pal energy.',
                    'The ranch line is moving again.',
                    'The Palbox is stable and glowing.',
                    'The island camp is ready for another shift.',
                    'The work Pals are back on the clock.',
                    'The camp smells like work and questionable decisions again.',
                    'The ranch is standing by with too many paws on payroll.'
                )
                JoinableTip = @(
                    'Bring spheres and maybe a plan.',
                    'Mind the Syndicate and keep your mount fed.',
                    'Do not let the transport Pals do all the heavy lifting.',
                    'Grab your gear and watch the cliffs.',
                    'Maybe keep a spare sphere on your belt.',
                    'Keep the spheres handy and the medicine closer.',
                    'Try not to let the transport crew unionize mid-run.'
                )
                SaveLead    = @(
                    'Packing up the ranch ledger.',
                    'Locking the latest Palbox state into place.',
                    'Stashing the freshest island snapshot.',
                    'Banking the last clean ranch run.',
                    'Tucking the work orders and ranch books away.',
                    'Filing the island shift under mostly controlled chaos.',
                    'Locking the ranch books before another Lunaris signs something.'
                )
                WaitLead    = @(
                    'Giving the island a second to breathe.',
                    'Letting the ranch quiet down first.',
                    'Holding while the work line eases off.',
                    'Giving the Palbox a beat to settle.',
                    'Letting the camp cool down before lights out.',
                    'Letting the ranch hands drop the last crate.',
                    'Giving the island one quiet breath between disasters.'
                )
                StopLead    = @(
                    'The island is settling down.',
                    'The ranch is closing up for now.',
                    'The Palbox is going dim.',
                    'The camp is going quiet for the night.',
                    'The island shift is ending cleanly.',
                    'The work line is off the clock for now.',
                    'The campfires are low and the Pals are finally settling down.'
                )
                RestartLead = @(
                    'The Palbox is cycling cleanly.',
                    'Running the island through a fresh boot.',
                    'The ranch is taking one clean reset.',
                    'The work line is getting a full power cycle.',
                    'The Palbox is doing a quick lap around the circuit.',
                    'Spinning the ranch through one clean shift change.',
                    'Running the island back through the Palbox with the edges sanded off.'
                )
                CrashLead   = @(
                    'A Pal definitely pulled the wrong lever.',
                    'Something in the ranch kicked over the control panel.',
                    'The island hit a bad stumble.',
                    'A work Pal probably found the dangerous button.',
                    'The Palbox took a hit from chaos.'
                )
                Caution     = @(
                    'Maybe keep the bigger Pals supervised.',
                    'Watch the ranch board for a minute.',
                    'Might be worth checking the Palbox twice.',
                    'Keep one eye on the island logs.',
                    'Maybe do not trust the overqualified transport Pal right now.'
                )
                StatusLead  = @(
                    'Island operations look like this.',
                    'The Palbox report reads like this.',
                    'The ranch board is showing this.',
                    'Current island telemetry says this.',
                    'The camp status is looking like this.'
                )
                PlayerRole  = 'tamers'
                WorldLabel  = 'island'
                DangerLabel = 'Pal uprising'
            }
        }
        'projectzomboid' {
            $overrides = [ordered]@{
                StartLead   = @(
                    'Knox County is groaning back to life.',
                    'The safehouse radios are crackling again.',
                    'The streets around Muldraugh are waking up.',
                    'The generator is coughing back to life in Knox County.',
                    'Louisville traffic still is not moving, but the server is.',
                    'The county sirens are still dead, but the server is not.',
                    'The safehouse board just woke back up under bad fluorescent light.'
                )
                ReadyLead   = @(
                    'The streets are awake again.',
                    'The safehouse is lit and holding steady.',
                    'Knox County is breathing through the static again.',
                    'The neighborhood is quiet in the bad way again.',
                    'The road out of Rosewood looks open enough.',
                    'The county is quiet in all the wrong places again.',
                    'The loot run can continue, assuming the dead cooperate.'
                )
                JoinableTip = @(
                    'Keep your head on a swivel.',
                    'Travel light and check every window twice.',
                    'Mind the corners, the alarms, and your gas can.',
                    'Maybe do not trust that parked car.',
                    'If you hear glass, run first and loot later.',
                    'Check your fuel, your bandages, and the rear seat.',
                    'If the alarm goes off, it was definitely not worth the toaster.'
                )
                SaveLead    = @(
                    'Securing the last good supply run.',
                    'Packing the safehouse ledger away tight.',
                    'Locking the latest loot haul behind the barricade.',
                    'Stashing the freshest Knox County chapter safely.',
                    'Tucking the last clean generator run into storage.',
                    'Writing the latest scavenging chapter into the safehouse log.',
                    'Barricading the newest loot story behind plywood and nails.'
                )
                WaitLead    = @(
                    'Barricading the doors before shutdown.',
                    'Letting the generator coast down first.',
                    'Giving the safehouse one last quiet sweep.',
                    'Holding while the radios die back to static.',
                    'Waiting for the street noise to settle into moaning.',
                    'Giving the generator one last uneven cough.',
                    'Letting the sheet ropes stop swaying before lights out.'
                )
                StopLead    = @(
                    'Knox County is going quiet.',
                    'The safehouse lights are going dark.',
                    'The barricades are holding and the town is fading out.',
                    'The streets are emptying back into silence.',
                    'The generator is off and the county is still.',
                    'The radios are fading back to static.',
                    'The county is getting the kind of quiet nobody likes.'
                )
                RestartLead = @(
                    'The safehouse power is cycling.',
                    'Running the generator through one clean restart.',
                    'Giving Knox County a fast reset before the next run.',
                    'The radios are going dark and loud again in one motion.',
                    'Taking the whole safehouse around the block and back.',
                    'Cycling the safehouse board before the next ugly supply run.',
                    'Taking Knox County once around the block and back again.'
                )
                CrashLead   = @(
                    'Something ugly rattled the safehouse.',
                    'A bad noise came through the barricades.',
                    'The county took a hard bite at the server.',
                    'Something in Knox County knocked the generator loose.',
                    'The streets hit back harder than expected.'
                )
                Caution     = @(
                    'Travel light and watch the corners.',
                    'Keep a radio on and the logs closer.',
                    'Might be worth checking the barricades and the console.',
                    'Stay near the safehouse board for a minute.',
                    'Keep one eye on the county map and one on the log lines.'
                )
                StatusLead  = @(
                    'The Knox County report reads like this.',
                    'The safehouse board is showing this.',
                    'Current county radio traffic sounds like this.',
                    'The latest survivor report reads like this.',
                    'The street-level status picture looks like this.'
                )
                PlayerRole  = 'survivors'
                WorldLabel  = 'safehouse'
                DangerLabel = 'bite risk'
            }
        }
        'satisfactory' {
            $overrides = [ordered]@{
                StartLead   = @(
                    'The factory floor is powering up.',
                    'The HUB is coming back online.',
                    'The Space Elevator lights are warming up again.',
                    'The production line is waking up section by section.',
                    'The breakers are coming up across the factory.',
                    'The startup siren just rolled across the factory floor.',
                    'Another shift is about to pretend the spaghetti is intentional.'
                )
                ReadyLead   = @(
                    'Conveyors are humming again.',
                    'The belts are moving and the lights are green.',
                    'Production is back on the board.',
                    'The HUB is stable and the line is rolling.',
                    'The factory is breathing through steel and power again.',
                    'The line is ugly, but it is moving beautifully.',
                    'The factory is alive enough to make management proud.'
                )
                JoinableTip = @(
                    'Time to punch in, pioneer.',
                    'Mind the belts, the hypertubes, and the drop-offs.',
                    'Grab your coffee and keep the line fed.',
                    'Do not get run over by your own tractor.',
                    'The AWESOME Sink can wait, the shift cannot.',
                    'Mind the catwalks and whatever that conveyor is doing.',
                    'Keep your coffee hot and your jetpack charged.'
                )
                SaveLead    = @(
                    'Stamping the latest production report.',
                    'Locking the factory ledger into place.',
                    'Banking the newest build sheet and belt map.',
                    'Packing away the latest shift output cleanly.',
                    'Tucking the last good factory snapshot into the archive.',
                    'Stamping the shift sheet before the auditors sniff around.',
                    'Banking the latest glorious tangle of production.'
                )
                WaitLead    = @(
                    'Letting the belts coast down.',
                    'Giving the line a chance to wind down.',
                    'Letting the machines spin to a stop.',
                    'Holding while the last train clears the block.',
                    'Giving the floor one last quiet minute.',
                    'Letting the last pallet clear the line.',
                    'Giving the assembly floor time to stop vibrating.'
                )
                StopLead    = @(
                    'The factory floor is going still.',
                    'The belts are stopping one by one.',
                    'The HUB lights are dimming for now.',
                    'The production line is clocking out cleanly.',
                    'The factory is closed up for this shift.',
                    'The shift whistle just died off.',
                    'The belts can finally stop pretending they enjoy this.'
                )
                RestartLead = @(
                    'Cycling the whole production line.',
                    'Running the factory through a clean maintenance lap.',
                    'Taking the HUB and belts through one clean reboot.',
                    'Giving the line a fast reset before the next shift.',
                    'The plant is doing a full power cycle.',
                    'Running maintenance with all the grace of a conveyor jam.',
                    'Taking the whole plant through another management-approved loop.'
                )
                CrashLead   = @(
                    'A conveyor probably ate the wrong bolt.',
                    'Something on the line kicked the breaker.',
                    'The factory took a bad production incident.',
                    'A hypertube dream turned into a hard stop.',
                    'The plant tripped over its own logistics.'
                )
                Caution     = @(
                    'Mind the belts and the biomass burners.',
                    'Keep one eye on the line and one on the logs.',
                    'Maybe check the breakers before the coffee gets cold.',
                    'Stay near the production board for a minute.',
                    'Watch the plant like it is one screw short of disaster.'
                )
                StatusLead  = @(
                    'Production telemetry says this.',
                    'The factory board looks like this.',
                    'Current shift data reads like this.',
                    'The HUB report is showing this.',
                    'The line status is coming through like this.'
                )
                PlayerRole  = 'pioneers'
                WorldLabel  = 'factory'
                DangerLabel = 'production incident'
            }
        }
        'valheim' {
            $overrides = [ordered]@{
                StartLead   = @(
                    'The longhouse fires are being lit.',
                    'The mead hall is waking under the rafters.',
                    'The ravens are back on the branch.',
                    'The forge is warming and the harbor is stirring.',
                    'The world beneath Yggdrasil is waking again.',
                    'The ward stones are stirring and the wind smells like rain.',
                    'The longship lanterns are being lit along the dock.'
                )
                ReadyLead   = @(
                    'The world tree is stirring.',
                    'The mead hall is warm and ready again.',
                    'The longship ropes are loose and the hearth is lit.',
                    'The forge smoke is rising over the hall again.',
                    'The saga is back on the table.',
                    'The hall is warm and the sea is calling again.',
                    'The ravens look satisfied enough to stick around.'
                )
                JoinableTip = @(
                    'Skal, and watch for trolls.',
                    'Mind the greydwarfs, the waves, and your stamina.',
                    'Keep your shield high and your portal fed.',
                    'Do not forget the mead before the mountain run.',
                    'Watch the shoreline and the black forest line.',
                    'Keep your mead close and your portal wood closer.',
                    'If the fog rolls in, pretend that was part of the plan.'
                )
                SaveLead    = @(
                    'Packing away the saga for tonight.',
                    'Tucking the latest longship tale into the hall records.',
                    'Locking the hearth log and the map table down.',
                    'Banking the newest rune-marked chapter.',
                    'Stashing the last clean voyage in the mead hall archive.',
                    'Etching the latest voyage into the hall record.',
                    'Locking the newest feast and fight into saga memory.'
                )
                WaitLead    = @(
                    'Giving the dock a quiet minute.',
                    'Letting the forge cool before the hall goes dark.',
                    'Holding while the harbor settles down.',
                    'Giving the hearth one last calm breath.',
                    'Waiting for the ravens to clear the roof beams.',
                    'Letting the embers settle under the rafters.',
                    'Giving the fjord wind one last pass through the dock.'
                )
                StopLead    = @(
                    'The mead hall is closing its doors.',
                    'The hearth is burning low for now.',
                    'The longhouse is settling into quiet.',
                    'The harbor is still and the hall is dim.',
                    'The forge is cooling and the hall is at rest.',
                    'The ward stones are going still for a while.',
                    'The longhouse is quiet except for the last coals.'
                )
                RestartLead = @(
                    'The ravens are circling for a fresh run.',
                    'Taking the hall through one clean turn of the saga.',
                    'Running the forge and harbor through a clean cycle.',
                    'Giving the longhouse a fresh start under the branches.',
                    'The mead hall is taking one hard reset.',
                    'Turning the hall around beneath Yggdrasil and setting it down again.',
                    'Running the forge, dock, and hall through one fresh saga loop.'
                )
                CrashLead   = @(
                    'A troll probably leaned on the server.',
                    'Something from the black forest hit the beams.',
                    'The hall took a bad swing from the wilds.',
                    'A rough sea just hit the harbor wall.',
                    'The saga tripped over a troll-sized problem.'
                )
                Caution     = @(
                    'Keep your shield up.',
                    'Watch the hall board and the next few runes.',
                    'Stay near the hearth and the logs for a minute.',
                    'Might be worth a look toward the harbor and the console.',
                    'Keep one eye on the ravens and one on the errors.'
                )
                StatusLead  = @(
                    'The mead hall report says this.',
                    'The longhouse board looks like this.',
                    'Current saga status reads like this.',
                    'The harbor watch is reporting this.',
                    'The hall ledger is showing this.'
                )
                PlayerRole  = 'vikings'
                WorldLabel  = 'mead hall'
                DangerLabel = 'troll trouble'
            }
        }
        'minecraft' {
            $overrides = [ordered]@{
                StartLead   = @(
                    'Chunks are loading in.',
                    'The spawn region is coming into focus.',
                    'The redstone dust is starting to glow again.',
                    'The chunk loader is spinning back up.',
                    'The world border is waking up again.',
                    'The mob cap is stretching and the farms are humming again.',
                    'The daylight cycle just found the switch.'
                )
                ReadyLead   = @(
                    'The world spawn is stable.',
                    'Chunks are in and the world is holding.',
                    'The redstone is quiet and the world is live.',
                    'The server tick is looking healthy again.',
                    'Spawn is loaded and the world is ready.',
                    'The overworld is holding together block by block.',
                    'Spawn is loaded and the redstone goblins are appeased.'
                )
                JoinableTip = @(
                    'Mind the creepers on your way in.',
                    'Keep your bed close and your gear closer.',
                    'Maybe do not test the TNT right away.',
                    'Watch your footing around the farms and cliffs.',
                    'Bring a pickaxe and leave the lava bucket where it is.',
                    'Mind the creepers, the elytra, and that one suspicious pressure plate.',
                    'Do not ask why the villager breeder sounds like that.'
                )
                SaveLead    = @(
                    'Packing the world data away safely.',
                    'Tucking the latest chunk state into storage.',
                    'Banking the newest block-by-block snapshot.',
                    'Locking the latest world save behind bedrock.',
                    'Stashing the freshest build log cleanly.',
                    'Packing the freshest build into chunk memory.',
                    'Tucking the latest survival chapter behind a wall of bedrock.'
                )
                WaitLead    = @(
                    'Giving the chunks a second to settle.',
                    'Letting the redstone cool off first.',
                    'Holding while the world tick smooths out.',
                    'Giving the chunk loader one quiet beat.',
                    'Letting the farms and furnaces coast down.',
                    'Letting the hoppers finish their gossip.',
                    'Giving the chunk edges a second to stop crackling.'
                )
                StopLead    = @(
                    'The world is going dark for now.',
                    'Spawn is quiet and the chunks are resting.',
                    'The chunk loader is going dark cleanly.',
                    'The server is tucking the world away for now.',
                    'The blocks are settling into stillness.',
                    'The overworld is settling into moonlight again.',
                    'The farms are quiet and spawn is finally breathing easy.'
                )
                RestartLead = @(
                    'Cycling the chunk loader cleanly.',
                    'Running the world through a clean reboot.',
                    'Taking the redstone line around once more.',
                    'Giving spawn one fresh reset.',
                    'The world is taking one clean lap through the loader.',
                    'Running the world through one clean trip past spawn.',
                    'Cycling the chunk loader before the next redstone argument.'
                )
                CrashLead   = @(
                    'Something definitely punched the chunk loader.',
                    'A creeper probably found the server room.',
                    'The redstone machine hit the wrong state.',
                    'Something in the world kicked the line sideways.',
                    'The chunk loader took a hard hit.'
                )
                Caution     = @(
                    'Maybe keep the TNT in storage.',
                    'Watch the console like it is a redstone clock.',
                    'Keep one eye on spawn and one on the logs.',
                    'Might be smart to leave the creeper farm alone for a minute.',
                    'This one deserves a quick check on the block report.'
                )
                StatusLead  = @(
                    'The block report looks like this.',
                    'The chunk board is showing this.',
                    'Current world status reads like this.',
                    'Spawn telemetry says this.',
                    'The server tick report is here.'
                )
                PlayerRole  = 'builders'
                WorldLabel  = 'world'
                DangerLabel = 'creeper business'
            }
        }
    }

    $merged = [ordered]@{}
    foreach ($key in $generic.Keys) { $merged[$key] = $generic[$key] }
    foreach ($key in $overrides.Keys) { $merged[$key] = $overrides[$key] }
    return $merged
}

function _ResolveDiscordFlavorValue {
    param([object]$Value)

    if ($null -eq $Value) { return $null }
    if ($Value -is [string]) { return $Value }

    if ($Value -is [System.Array]) {
        if ($Value.Count -le 0) { return '' }
        return (_ResolveDiscordFlavorValue -Value (Get-Random -InputObject $Value))
    }

    if ($Value -is [System.Collections.IList]) {
        if ($Value.Count -le 0) { return '' }
        return (_ResolveDiscordFlavorValue -Value (Get-Random -InputObject @($Value)))
    }

    return $Value
}

function _GetDiscordSystemFlavorData {
    $data = [ordered]@{
        AppName            = 'Etherium Command Center'
        OnlineLead         = @(
            'The bunker lights are on again. Corporate can unclench for a minute.',
            'I am back at the console. Apparently that is still my job.',
            'The relay wall is breathing again, so nobody upstairs needs to panic yet.',
            'The old bunker board is hot again. Try not to make me explain it to the overlords.',
            'ECC is back on watch, which means I am still employed for another shift.',
            'The fans spun up and nobody upstairs has called screaming yet.',
            'The bunker survived another boot sequence, which I assume disappoints someone.',
            'The command pit is awake again, so the overlords get their precious uptime.',
            'The main board is green enough to keep payroll from asking questions.',
            'Everything is humming just loud enough to suggest I still answer to corporate.'
        )
        OfflineLead        = @(
            'The bunker is going dark for now. If corporate asks, this was deliberate.',
            'The board is quiet again, and I would like to keep it that way.',
            'The watch desk is stepping off the line before the overlords invent a new emergency.',
            'The relay lights are fading out. Please do not wake upper management.',
            'ECC is off the floor for now, which is the closest thing to peace around here.',
            'The control pit is dark and for one blessed minute nobody is demanding status reports.',
            'The shutters are down and the overlords can file their complaints later.',
            'Everything is cooling off before corporate decides this counts as loafing.',
            'The main board is cold, which is my preferred emotional temperature.',
            'ECC is off duty unless somebody upstairs finds the red phone again.'
        )
        ReloadUiLead       = @(
            'Redrawing the bunker glass so the overlords stop staring at smudges.',
            'Refreshing the control panels without touching the live floor, because I enjoy survival.',
            'Giving the front board a quick repaint before someone upstairs files a complaint.',
            'Reloading the dashboard in place. Try not to lean on anything expensive.',
            'Cleaning up the UI again, because apparently the bunker aesthetics matter now.',
            'Scraping the grime off these panels so corporate can admire the numbers.',
            'Tuning the bunker glass because apparently flicker is a reportable offense.',
            'Giving the dashboard another polite slap until the panels behave.',
            'Restacking the front-end sightlines before somebody upstairs blames me personally.',
            'Repainting the operator view without touching the machinery beneath it.'
        )
        ReloadBotLead      = @(
            'Kicking the comms rack until the bot remembers who signs its checks.',
            'The Discord wire is getting another forced attitude adjustment.',
            'Resetting the radio desk before the overlords notice the silence.',
            'The bunker operator in the comms closet is being dragged back onto shift.',
            'Reloading the bot layer because apparently this is still easier than dying.',
            'Coaxing the Discord gremlin back onto the wire before corporate notices the gap.',
            'Giving the radio stack a shove and hoping the bot salutes.',
            'Waking the comms relay up with all the tenderness policy allows.',
            'Forcing the bot back into uniform before the overlords ask why chat is quiet.',
            'Retuning the bunker radios because the bot has once again developed opinions.'
        )
        ReloadCommandsLead = @(
            'Re-sorting the command binder for the benefit of corporate procedure.',
            'Reloading the command deck because the overlords adore paperwork.',
            'Putting the hotkeys and profiles back in line before somebody audits the bunker.',
            'The operations binder is getting another fresh index. Thrilling stuff.',
            'Straightening out the command shelf so upstairs can pretend this place is orderly.',
            'Rebuilding the order of operations so corporate can pretend this place is disciplined.',
            'Putting the command cards back in the right slots before an audit ghost appears.',
            'Sweeping the control binder for loose pages and bad decisions.',
            'Lining the command deck back up so the overlords can keep pretending this is elegant.',
            'Giving the playbook another forced march through the filing cabinet.'
        )
        FullRestartLead    = @(
            'Taking the whole bunker down and back up before corporate takes me with it.',
            'Cycling the full rack because policy says pain builds character.',
            'The whole command pit is taking one hard reboot. Try to look calm.',
            'Shutting the bunker lights and bringing them back before the overlords get curious.',
            'Running the full app stack around once more because apparently fear is an operating model.',
            'Taking every relay in the bunker around the loop because somebody upstairs wanted certainty.',
            'Putting the whole bunker through the wash and hoping nothing important screams.',
            'Pulling the master lever and trusting corporate engineering from a safe emotional distance.',
            'Doing the full shutdown dance so the overlords can feel in control.',
            'Rebooting the whole nest of wires and bad decisions one more time.'
        )
    }

    return $data
}

function _FormatDiscordMessageTemplate {
    param(
        [string]$Template,
        [object]$Tokens
    )

    if ([string]::IsNullOrWhiteSpace($Template)) { return '' }

    $text = [regex]::Replace($Template, '\{([A-Za-z0-9_]+)\}', {
        param($m)
        $tokenName = $m.Groups[1].Value
        $hasToken = $false
        if ($Tokens) {
            try {
                $hasToken = ($Tokens.Keys -contains $tokenName)
            } catch {
                $hasToken = $false
            }
        }
        if ($hasToken -and $null -ne $Tokens[$tokenName]) {
            return "$($Tokens[$tokenName])"
        }
        return ''
    })

    $text = $text -replace '\s+\.', '.'
    $text = $text -replace '\s+,', ','
    $text = $text -replace '\s{2,}', ' '
    $text = $text -replace '\|\s+\|', '|'
    return $text.Trim()
}

function New-DiscordGameMessage {
    param(
        [object]$Profile,
        [string]$Event,
        [hashtable]$Values = $null,
        [string]$Prefix = ''
    )

    $gameName = _GetNotificationGameName -Profile $Profile -Prefix $Prefix
    $flavor = _GetDiscordFlavorData -Profile $Profile -Prefix $Prefix
    $templates = @()

    $eventName = if ($null -ne $Event) { "$Event" } else { '' }

    switch ($eventName.ToLowerInvariant()) {
        'received_start' {
            $templates = @(
                '[RECEIVED] {Requester} start command received for {GameName}. {StartLead} {OperatorOrderLead}',
                '[RECEIVED] {Requester} asked me to wake up {GameName}. {StartLead} {OperatorOrderLead}',
                '[RECEIVED] {Requester} gave the green light for {GameName}. {StartLead} {OperatorOrderLead}',
                '[RECEIVED] {Requester} wants {GameName} live. {StartLead} {OperatorOrderLead}',
                '[RECEIVED] {Requester} rang the bell for {GameName}. {StartLead} {OperatorOrderLead}',
                '[RECEIVED] {Requester} kicked off a start for {GameName}. {StartLead} {OperatorOrderLead}',
                '[RECEIVED] {Requester} told ECC to bring {GameName} back to the {WorldLabel}. {StartLead} {OperatorOrderLead}',
                '[RECEIVED] {Requester} wants the {WorldLabel} open again in {GameName}. {StartLead} {OperatorOrderLead}',
                '[RECEIVED] {Requester} called everyone to the {WorldLabel} in {GameName}. {StartLead} {OperatorOrderLead}'
            )
        }
        'received_stop' {
            $templates = @(
                '[RECEIVED] {Requester} stop command received for {GameName}. {WaitLead} {OperatorOrderLead}',
                '[RECEIVED] {Requester} asked me to shut down {GameName}. {WaitLead} {OperatorOrderLead}',
                '[RECEIVED] {Requester} called for a clean stop on {GameName}. {WaitLead} {OperatorOrderLead}',
                '[RECEIVED] {Requester} wants {GameName} tucked in for the night. {WaitLead} {OperatorOrderLead}',
                '[RECEIVED] {Requester} is winding down {GameName}. {WaitLead} {OperatorOrderLead}',
                '[RECEIVED] {Requester} ordered a shutdown for {GameName}. {WaitLead} {OperatorOrderLead}',
                '[RECEIVED] {Requester} wants the {WorldLabel} in {GameName} closed up. {WaitLead} {OperatorOrderLead}',
                '[RECEIVED] {Requester} called last round for {GameName}. {WaitLead} {OperatorOrderLead}',
                '[RECEIVED] {Requester} asked ECC to dim the lights on {GameName}. {WaitLead} {OperatorOrderLead}'
            )
        }
        'received_restart' {
            $templates = @(
                '[RECEIVED] {Requester} restart command received for {GameName}. {RestartLead} {OperatorOrderLead}',
                '[RECEIVED] {Requester} wants {GameName} cycled cleanly. {RestartLead} {OperatorOrderLead}',
                '[RECEIVED] {Requester} called for a fresh spin on {GameName}. {RestartLead} {OperatorOrderLead}',
                '[RECEIVED] {Requester} is rebooting {GameName}. {RestartLead} {OperatorOrderLead}',
                '[RECEIVED] {Requester} wants {GameName} back on its feet. {RestartLead} {OperatorOrderLead}',
                '[RECEIVED] {Requester} triggered a restart for {GameName}. {RestartLead} {OperatorOrderLead}',
                '[RECEIVED] {Requester} wants {GameName} to take one clean lap around the circuit. {RestartLead} {OperatorOrderLead}',
                '[RECEIVED] {Requester} asked ECC to cycle the {WorldLabel} for {GameName}. {RestartLead} {OperatorOrderLead}',
                '[RECEIVED] {Requester} wants the {DangerLabel} clock reset on {GameName}. {RestartLead} {OperatorOrderLead}'
            )
        }
        'received_save' {
            $templates = @(
                '[RECEIVED] {Requester} save command received for {GameName}. {SaveLead} {OperatorOrderLead}',
                '[RECEIVED] {Requester} asked for a save on {GameName}. {SaveLead} {OperatorOrderLead}',
                '[RECEIVED] {Requester} wants {GameName} locked in. {SaveLead} {OperatorOrderLead}',
                '[RECEIVED] {Requester} called for a fresh save on {GameName}. {SaveLead} {OperatorOrderLead}',
                '[RECEIVED] {Requester} is banking progress for {GameName}. {SaveLead} {OperatorOrderLead}',
                '[RECEIVED] {Requester} asked ECC to preserve {GameName}. {SaveLead} {OperatorOrderLead}',
                '[RECEIVED] {Requester} wants the latest {WorldLabel} state banked for {GameName}. {SaveLead} {OperatorOrderLead}',
                '[RECEIVED] {Requester} is making sure the {PlayerRole} do not lose their progress in {GameName}. {SaveLead} {OperatorOrderLead}',
                '[RECEIVED] {Requester} called for a safety save on {GameName}. {SaveLead} {OperatorOrderLead}'
            )
        }
        'received_status' {
            $templates = @(
                '[RECEIVED] {Requester} checking status for {GameName}. {StatusLead} {OperatorStatusLead}',
                '[RECEIVED] {Requester} asked for the latest status on {GameName}. {StatusLead} {OperatorStatusLead}',
                '[RECEIVED] {Requester} wants the latest read on {GameName}. {StatusLead} {OperatorStatusLead}',
                '[RECEIVED] {Requester} pinged ECC for a {GameName} status check. {StatusLead} {OperatorStatusLead}',
                '[RECEIVED] {Requester} is checking the pulse on {GameName}. {StatusLead} {OperatorStatusLead}',
                '[RECEIVED] {Requester} asked what {GameName} is doing right now. {StatusLead} {OperatorStatusLead}',
                '[RECEIVED] {Requester} wants the latest field report from {GameName}. {StatusLead} {OperatorStatusLead}',
                '[RECEIVED] {Requester} asked ECC for the current read on the {WorldLabel} in {GameName}. {StatusLead} {OperatorStatusLead}',
                '[RECEIVED] {Requester} wants to know how {GameName} is holding together. {StatusLead} {OperatorStatusLead}'
            )
        }
        'received_command' {
            $templates = @(
                '[RECEIVED] {Requester} running {Command} for {GameName}. {OperatorCommandLead}',
                '[RECEIVED] {Requester} sent {Command} to {GameName}. {OperatorCommandLead}',
                '[RECEIVED] {Requester} queued {Command} for {GameName}. {OperatorCommandLead}',
                '[RECEIVED] {Requester} wants {Command} run on {GameName}. {OperatorCommandLead}',
                '[RECEIVED] {Requester} sent a live command to {GameName}: {Command}. {OperatorCommandLead}',
                '[RECEIVED] {Requester} asked ECC to push {Command} into {GameName}. {OperatorCommandLead}',
                '[RECEIVED] {Requester} fired {Command} into the {WorldLabel} for {GameName}. {OperatorCommandLead}',
                '[RECEIVED] {Requester} sent a command downrange for {GameName}: {Command}. {OperatorCommandLead}',
                '[RECEIVED] {Requester} wants the {PlayerRole} in {GameName} to feel this one: {Command}. {OperatorCommandLead}'
            )
        }
        'starting' {
            $templates = @(
                '[STARTING] {GameName} is starting. {StartLead}',
                '[STARTING] {GameName} launch accepted. {StartLead}',
                '[STARTING] {GameName} is waking up now. {StartLead}',
                '[STARTING] {GameName} is coming online. {StartLead}',
                '[STARTING] ECC has started the boot sequence for {GameName}. {StartLead}',
                '[STARTING] {GameName} is rolling out of bed. {StartLead}',
                '[STARTING] {GameName} is powering in. {StartLead}',
                '[STARTING] {GameName} is on the rise. {StartLead}',
                '[STARTING] ECC is opening the {WorldLabel} doors for {GameName}. {StartLead}',
                '[STARTING] {GameName} is shaking off the dust. {StartLead}',
                '[STARTING] The {WorldLabel} in {GameName} is rumbling awake. {StartLead}',
                '[STARTING] {GameName} got the signal and is booting hard. {StartLead}'
            )
        }
        'online' {
            $templates = @(
                '[ONLINE] {GameName} is now online (PID {Pid}). {ReadyLead}',
                '[ONLINE] {GameName} is online and stable (PID {Pid}). {ReadyLead}',
                '[ONLINE] {GameName} is up and breathing (PID {Pid}). {ReadyLead}',
                '[ONLINE] {GameName} came online cleanly (PID {Pid}). {ReadyLead}',
                '[ONLINE] {GameName} is live now (PID {Pid}). {ReadyLead}',
                '[ONLINE] {GameName} stood up without a hitch (PID {Pid}). {ReadyLead}',
                '[ONLINE] {GameName} is active and holding steady (PID {Pid}). {ReadyLead}',
                '[ONLINE] {GameName} is fully awake (PID {Pid}). {ReadyLead}',
                '[ONLINE] {GameName} planted its feet and is ready (PID {Pid}). {ReadyLead}',
                '[ONLINE] {GameName} is holding the line (PID {Pid}). {ReadyLead}',
                '[ONLINE] {GameName} is live and settled in (PID {Pid}). {ReadyLead}',
                '[ONLINE] The {WorldLabel} in {GameName} is stable again (PID {Pid}). {ReadyLead}'
            )
        }
        'joinable' {
            $templates = @(
                '[JOINABLE] {GameName} is ready for players. {JoinableTip}',
                '[JOINABLE] {GameName} is joinable now. {JoinableTip}',
                '[JOINABLE] {GameName} just opened its doors for {PlayerRole}. {JoinableTip}',
                '[JOINABLE] {GameName} is ready for the crew to pile in. {JoinableTip}',
                '[JOINABLE] {GameName} is good to go for {PlayerRole}. {JoinableTip}',
                '[JOINABLE] {GameName} is open for business. {JoinableTip}',
                '[JOINABLE] {GameName} is ready if the {PlayerRole} are. {JoinableTip}',
                '[JOINABLE] {GameName} is clear for drop-in. {JoinableTip}',
                '[JOINABLE] The {WorldLabel} in {GameName} is open again. {JoinableTip}',
                '[JOINABLE] {GameName} is taking visitors. {JoinableTip}',
                '[JOINABLE] {GameName} is ready for boots on the ground. {JoinableTip}',
                '[JOINABLE] {GameName} is signaling all clear for the {PlayerRole}. {JoinableTip}'
            )
        }
        'save_sent' {
            $templates = @(
                '[OK] {GameName} got the save call. {SaveLead}',
                '[OK] {GameName} accepted the save request. {SaveLead}',
                '[OK] ECC pushed a save into {GameName}. {SaveLead}',
                '[OK] {GameName} is locking things in now. {SaveLead}',
                '[OK] Save command landed for {GameName}. {SaveLead}',
                '[OK] {GameName} is putting the latest progress on the shelf. {SaveLead}',
                '[OK] {GameName} heard the save call and is preserving the {WorldLabel}. {SaveLead}',
                '[OK] ECC handed {GameName} a fresh save order. {SaveLead}',
                '[OK] {GameName} is locking the latest chapter into place. {SaveLead}',
                '[OK] {GameName} took the save order and is sealing the books. {SaveLead}',
                '[OK] ECC tagged the latest {WorldLabel} state for safekeeping in {GameName}. {SaveLead}',
                '[OK] {GameName} caught the save signal and is stashing the newest progress. {SaveLead}'
            )
        }
        'saving' {
            $templates = @(
                '[SAVING] {GameName} save command sent. Waiting {WaitSeconds}s before shutdown. {SaveLead}',
                '[SAVING] {GameName} is saving before shutdown. Waiting {WaitSeconds}s. {SaveLead}',
                '[SAVING] {GameName} is writing things down before lights out. Waiting {WaitSeconds}s. {SaveLead}',
                '[SAVING] ECC asked {GameName} for a safe save, then a {WaitSeconds}s pause. {SaveLead}',
                '[SAVING] {GameName} is packing up the world first. Shutdown follows in {WaitSeconds}s. {SaveLead}',
                '[SAVING] {GameName} is locking in progress, then waiting {WaitSeconds}s before shutdown. {SaveLead}',
                '[SAVING] {GameName} is stashing the last good state before shutdown in {WaitSeconds}s. {SaveLead}',
                '[SAVING] ECC is asking {GameName} to nail down the latest {WorldLabel} state, then wait {WaitSeconds}s. {SaveLead}',
                '[SAVING] {GameName} is sealing the books before power-down in {WaitSeconds}s. {SaveLead}',
                '[SAVING] {GameName} is tucking the latest run away, then holding {WaitSeconds}s before shutdown. {SaveLead}',
                '[SAVING] ECC is giving {GameName} {WaitSeconds}s to pin the newest snapshot down before lights out. {SaveLead}',
                '[SAVING] {GameName} is finishing one last write pass before shutdown in {WaitSeconds}s. {SaveLead}'
            )
        }
        'waiting' {
            $templates = @(
                '[WAITING] {GameName} is waiting {WaitSeconds}s before shutdown. {WaitLead}',
                '[WAITING] {GameName} is pausing {WaitSeconds}s before shutdown. {WaitLead}',
                '[WAITING] {GameName} is taking a {WaitSeconds}s cooldown before shutdown. {WaitLead}',
                '[WAITING] ECC is giving {GameName} {WaitSeconds}s to settle before shutdown. {WaitLead}',
                '[WAITING] {GameName} is coasting for {WaitSeconds}s before power-down. {WaitLead}',
                '[WAITING] {GameName} has a {WaitSeconds}s quiet window before shutdown. {WaitLead}',
                '[WAITING] {GameName} is riding out a {WaitSeconds}s buffer before the lights go out. {WaitLead}',
                '[WAITING] ECC is giving the {WorldLabel} in {GameName} {WaitSeconds}s to cool off. {WaitLead}',
                '[WAITING] {GameName} is in final countdown mode for {WaitSeconds}s. {WaitLead}',
                '[WAITING] {GameName} is taking one last quiet lap for {WaitSeconds}s before shutdown. {WaitLead}',
                '[WAITING] ECC is holding {GameName} at the line for {WaitSeconds}s before the switch flips. {WaitLead}',
                '[WAITING] The {WorldLabel} in {GameName} is sitting in a {WaitSeconds}s cooldown pocket. {WaitLead}'
            )
        }
        'waiting_no_save' {
            $templates = @(
                '[WAITING] {GameName} has no working save command. Waiting {WaitSeconds}s before shutdown. {WaitLead}',
                '[WAITING] {GameName} could not save first, so it is waiting {WaitSeconds}s before shutdown. {WaitLead}',
                '[WAITING] ECC could not confirm a save path for {GameName}, so it is holding {WaitSeconds}s before shutdown. {WaitLead}',
                '[WAITING] {GameName} is entering a {WaitSeconds}s shutdown buffer because no save command was available. {WaitLead}',
                '[WAITING] {GameName} did not have a safe save route, so ECC is pausing {WaitSeconds}s before shutdown. {WaitLead}',
                '[WAITING] No save path checked out for {GameName}, so ECC is riding out {WaitSeconds}s before shutdown. {WaitLead}',
                '[WAITING] {GameName} is on a no-save shutdown buffer for {WaitSeconds}s. {WaitLead}',
                '[WAITING] The save route for {GameName} came up empty, so the shutdown clock is coasting for {WaitSeconds}s. {WaitLead}',
                '[WAITING] ECC is giving {GameName} one last {WaitSeconds}s grace window because the save command path was not there. {WaitLead}',
                '[WAITING] {GameName} is cooling its heels for {WaitSeconds}s before shutdown because no clean save route answered. {WaitLead}'
            )
        }
        'stopped' {
            $templates = @(
                '[STOPPED] {GameName} stopped safely. {StopLead}',
                '[STOPPED] {GameName} is offline. {StopLead}',
                '[STOPPED] {GameName} powered down cleanly. {StopLead}',
                '[STOPPED] {GameName} tucked in without drama. {StopLead}',
                '[STOPPED] {GameName} is down for now. {StopLead}',
                '[STOPPED] {GameName} closed out safely. {StopLead}',
                '[STOPPED] {GameName} went quiet cleanly. {StopLead}',
                '[STOPPED] {GameName} shut its doors without complaint. {StopLead}',
                '[STOPPED] The {WorldLabel} in {GameName} is closed up for now. {StopLead}',
                '[STOPPED] {GameName} is sleeping cleanly. {StopLead}',
                '[STOPPED] ECC tucked {GameName} in and the room finally went still. {StopLead}',
                '[STOPPED] {GameName} dropped into a clean quiet state. {StopLead}',
                '[STOPPED] The last activity in {GameName} just blinked out. {StopLead}'
            )
        }
        'restarting' {
            $templates = @(
                '[RESTARTING] {GameName} restart sequence started. {RestartLead}',
                '[RESTARTING] {GameName} is cycling now. {RestartLead}',
                '[RESTARTING] ECC is running a full restart on {GameName}. {RestartLead}',
                '[RESTARTING] {GameName} is taking the scenic route through a reboot. {RestartLead}',
                '[RESTARTING] {GameName} is going around once more. {RestartLead}',
                '[RESTARTING] {GameName} is entering a clean reboot cycle. {RestartLead}',
                '[RESTARTING] {GameName} is stepping out and coming right back. {RestartLead}',
                '[RESTARTING] ECC is cycling the {WorldLabel} in {GameName}. {RestartLead}',
                '[RESTARTING] {GameName} is taking one clean lap around the loop. {RestartLead}',
                '[RESTARTING] {GameName} is rebooting with the dust still in the air. {RestartLead}',
                '[RESTARTING] ECC is pulling {GameName} through another full turn of the wheel. {RestartLead}',
                '[RESTARTING] {GameName} is dropping out and rejoining the fight in one clean cycle. {RestartLead}',
                '[RESTARTING] The {WorldLabel} in {GameName} is taking its scheduled spin through the circuit. {RestartLead}'
            )
        }
        'restarted' {
            $templates = @(
                '[RESTARTED] {GameName} restarted successfully and is back online. {ReadyLead}',
                '[RESTARTED] {GameName} is back from restart. {ReadyLead}',
                '[RESTARTED] {GameName} made it through the restart cleanly. {ReadyLead}',
                '[RESTARTED] {GameName} is back on its feet. {ReadyLead}',
                '[RESTARTED] {GameName} survived the reboot and is online again. {ReadyLead}',
                '[RESTARTED] {GameName} came back clean after the cycle. {ReadyLead}',
                '[RESTARTED] {GameName} is back in the fight. {ReadyLead}',
                '[RESTARTED] The {WorldLabel} in {GameName} is live again. {ReadyLead}',
                '[RESTARTED] {GameName} took the reboot and kept moving. {ReadyLead}',
                '[RESTARTED] {GameName} landed the restart and is steady again. {ReadyLead}',
                '[RESTARTED] ECC cycled {GameName} and it came back clean on the first swing. {ReadyLead}',
                '[RESTARTED] {GameName} walked through the reset and came out standing. {ReadyLead}',
                '[RESTARTED] The reboot dust cleared and {GameName} is solid again. {ReadyLead}'
            )
        }
        'restarted_auto' {
            $templates = @(
                '[RESTARTED] {GameName} recovered automatically and is back online (PID {Pid}). {ReadyLead}',
                '[RESTARTED] {GameName} auto-restarted cleanly (PID {Pid}). {ReadyLead}',
                '[RESTARTED] ECC pulled {GameName} back on its feet automatically (PID {Pid}). {ReadyLead}',
                '[RESTARTED] {GameName} recovered on its own track and is back online (PID {Pid}). {ReadyLead}',
                '[RESTARTED] {GameName} bounced back under ECC control (PID {Pid}). {ReadyLead}',
                '[RESTARTED] ECC hauled {GameName} back into the fight automatically (PID {Pid}). {ReadyLead}',
                '[RESTARTED] {GameName} recovered from the hit and stood back up (PID {Pid}). {ReadyLead}',
                '[RESTARTED] The {WorldLabel} in {GameName} came back under auto-recovery (PID {Pid}). {ReadyLead}',
                '[RESTARTED] {GameName} got knocked down and ECC stood it back up automatically (PID {Pid}). {ReadyLead}',
                '[RESTARTED] Auto-recovery landed cleanly for {GameName} (PID {Pid}). {ReadyLead}',
                '[RESTARTED] ECC dragged {GameName} out of the ditch and back online (PID {Pid}). {ReadyLead}',
                '[RESTARTED] {GameName} rejoined the line under automatic recovery control (PID {Pid}). {ReadyLead}'
            )
        }
        'autosave_started' {
            $templates = @(
                '[AUTOSAVE] {GameName} auto-save started. {SaveLead}',
                '[AUTOSAVE] {GameName} is locking in progress now. {SaveLead}',
                '[AUTOSAVE] ECC kicked off an auto-save for {GameName}. {SaveLead}',
                '[AUTOSAVE] {GameName} is writing down the latest chapter. {SaveLead}',
                '[AUTOSAVE] {GameName} is preserving the current run. {SaveLead}',
                '[AUTOSAVE] {GameName} is tucking the latest {WorldLabel} state away. {SaveLead}',
                '[AUTOSAVE] ECC is banking the latest progress for {GameName}. {SaveLead}',
                '[AUTOSAVE] {GameName} is snapping a safety save before the next round. {SaveLead}',
                '[AUTOSAVE] {GameName} is pinning down the newest state before anything stupid happens. {SaveLead}',
                '[AUTOSAVE] ECC just told {GameName} to bank the latest run before the next hit lands. {SaveLead}',
                '[AUTOSAVE] The {WorldLabel} in {GameName} is getting folded into a safety snapshot. {SaveLead}',
                '[AUTOSAVE] {GameName} is writing the fresh state to the shelf before anyone notices. {SaveLead}'
            )
        }
        'autosave_done' {
            $templates = @(
                '[AUTOSAVE] {GameName} auto-save completed successfully. {SaveLead}',
                '[AUTOSAVE] {GameName} progress was saved cleanly. {SaveLead}',
                '[AUTOSAVE] {GameName} finished its auto-save without a hitch. {SaveLead}',
                '[AUTOSAVE] {GameName} tucked the latest progress away safely. {SaveLead}',
                '[AUTOSAVE] {GameName} preserved the latest state cleanly. {SaveLead}',
                '[AUTOSAVE] {GameName} banked the latest {WorldLabel} snapshot without a hitch. {SaveLead}',
                '[AUTOSAVE] ECC locked down the newest progress for {GameName}. {SaveLead}',
                '[AUTOSAVE] {GameName} wrapped the safety save cleanly. {SaveLead}',
                '[AUTOSAVE] {GameName} sealed the newest chapter without dropping a page. {SaveLead}',
                '[AUTOSAVE] ECC got the latest {WorldLabel} state tucked away cleanly for {GameName}. {SaveLead}',
                '[AUTOSAVE] The newest run in {GameName} is safely on the shelf. {SaveLead}',
                '[AUTOSAVE] {GameName} put the fresh snapshot away and never missed a beat. {SaveLead}'
            )
        }
        'blocked' {
            $templates = @(
                '[BLOCKED] {GameName} start was blocked. {Reason} {OperatorBlockedLead} {Caution}',
                '[BLOCKED] {GameName} could not start yet. {Reason} {OperatorBlockedLead} {Caution}',
                '[BLOCKED] ECC held {GameName} back for safety. {Reason} {OperatorBlockedLead} {Caution}',
                '[BLOCKED] {GameName} is staying parked for now. {Reason} {OperatorBlockedLead} {Caution}',
                '[BLOCKED] {GameName} did not clear the safety check. {Reason} {OperatorBlockedLead} {Caution}',
                '[BLOCKED] ECC would not let {GameName} into the {WorldLabel} just yet. {Reason} {OperatorBlockedLead} {Caution}',
                '[BLOCKED] {GameName} is waiting on the sideline. {Reason} {OperatorBlockedLead} {Caution}',
                '[BLOCKED] {GameName} hit a red light before launch. {Reason} {OperatorBlockedLead} {Caution}'
            )
        }
        'startup_failed' {
            $templates = @(
                '[STARTUP FAILED] {GameName} never fully woke up. {Reason} {Action} {OperatorStartupFailLead}',
                '[STARTUP FAILED] {GameName} stalled before it was ready. {Reason} {Action} {OperatorStartupFailLead}',
                '[STARTUP FAILED] {GameName} launched, but never reached a healthy state. {Reason} {Action} {OperatorStartupFailLead}',
                '[STARTUP FAILED] {GameName} tried to stand up and never settled in. {Reason} {Action} {OperatorStartupFailLead}',
                '[STARTUP FAILED] {GameName} never made it past the boot phase. {Reason} {Action} {OperatorStartupFailLead}',
                '[STARTUP FAILED] The {WorldLabel} in {GameName} never stabilized. {Reason} {Action} {OperatorStartupFailLead}',
                '[STARTUP FAILED] {GameName} got out of bed and fell right back over. {Reason} {Action} {OperatorStartupFailLead}',
                '[STARTUP FAILED] {GameName} never gave ECC the all-clear. {Reason} {Action} {OperatorStartupFailLead}'
            )
        }
        'idle_shutdown' {
            $templates = @(
                '[IDLE SHUTDOWN] {GameName} is shutting down from inactivity. {Reason}',
                '[IDLE SHUTDOWN] {GameName} went quiet too long, so ECC is powering it down. {Reason}',
                '[IDLE SHUTDOWN] {GameName} stayed empty long enough that ECC is closing it up. {Reason}',
                '[IDLE SHUTDOWN] {GameName} has been idle too long, so ECC is dimming the lights. {Reason}',
                '[IDLE SHUTDOWN] {GameName} stayed too quiet, and ECC is tucking it away. {Reason}',
                '[IDLE SHUTDOWN] The {WorldLabel} in {GameName} has gone cold, so ECC is closing it down. {Reason}',
                '[IDLE SHUTDOWN] No sign of the {PlayerRole}, so ECC is shutting {GameName} for now. {Reason}',
                '[IDLE SHUTDOWN] {GameName} has been all quiet and no motion, so the lights are going out. {Reason}',
                '[IDLE SHUTDOWN] ECC is folding {GameName} back up because the room stayed empty too long. {Reason}',
                '[IDLE SHUTDOWN] The {PlayerRole} never came back, so ECC is sealing {GameName} up for now. {Reason}',
                '[IDLE SHUTDOWN] {GameName} sat there too still for too long, so the switch is going down. {Reason}',
                '[IDLE SHUTDOWN] The {WorldLabel} in {GameName} has been dead air long enough that ECC is closing the bunker door. {Reason}'
            )
        }
        'restart_warning' {
            $templates = @(
                '[RESTART WARNING] {GameName} restart in {Minutes} minute(s). {PlayerSummary} {Caution}',
                '[RESTART WARNING] {GameName} will restart in {Minutes} minute(s). {PlayerSummary} {Caution}',
                '[RESTART WARNING] ECC is lining up a restart for {GameName} in {Minutes} minute(s). {PlayerSummary} {Caution}',
                '[RESTART WARNING] {GameName} has a scheduled reboot in {Minutes} minute(s). {PlayerSummary} {Caution}',
                '[RESTART WARNING] {GameName} is heading toward a restart in {Minutes} minute(s). {PlayerSummary} {Caution}',
                '[RESTART WARNING] The {WorldLabel} in {GameName} is getting a scheduled cycle in {Minutes} minute(s). {PlayerSummary} {Caution}',
                '[RESTART WARNING] ECC is calling a maintenance window on {GameName} in {Minutes} minute(s). {PlayerSummary} {Caution}',
                '[RESTART WARNING] {GameName} is on the clock for a clean reboot in {Minutes} minute(s). {PlayerSummary} {Caution}',
                '[RESTART WARNING] The reboot window for {GameName} opens in {Minutes} minute(s). {PlayerSummary} {Caution}',
                '[RESTART WARNING] ECC is about to walk {GameName} into a restart in {Minutes} minute(s). {PlayerSummary} {Caution}',
                '[RESTART WARNING] The {PlayerRole} in {GameName} have {Minutes} minute(s) before the cycle hits. {PlayerSummary} {Caution}',
                '[RESTART WARNING] {GameName} is staring down a maintenance turn in {Minutes} minute(s). {PlayerSummary} {Caution}'
            )
        }
        'scheduled_restart_started' {
            $templates = @(
                '[SCHEDULED RESTART] {GameName} scheduled restart has begun. {PlayerSummary} {RestartLead}',
                '[SCHEDULED RESTART] {GameName} is entering its scheduled restart cycle. {PlayerSummary} {RestartLead}',
                '[SCHEDULED RESTART] ECC has started the planned restart for {GameName}. {PlayerSummary} {RestartLead}',
                '[SCHEDULED RESTART] {GameName} is cycling on schedule. {PlayerSummary} {RestartLead}',
                '[SCHEDULED RESTART] {GameName} just hit its maintenance window. {PlayerSummary} {RestartLead}',
                '[SCHEDULED RESTART] The {WorldLabel} in {GameName} is entering scheduled maintenance. {PlayerSummary} {RestartLead}',
                '[SCHEDULED RESTART] ECC is walking {GameName} through its planned reboot. {PlayerSummary} {RestartLead}',
                '[SCHEDULED RESTART] {GameName} reached the scheduled turn in the cycle. {PlayerSummary} {RestartLead}',
                '[SCHEDULED RESTART] ECC just shoved {GameName} into its maintenance lane. {PlayerSummary} {RestartLead}',
                '[SCHEDULED RESTART] {GameName} is taking its planned pass through the restart tunnel. {PlayerSummary} {RestartLead}',
                '[SCHEDULED RESTART] The scheduled reboot window just opened on {GameName}. {PlayerSummary} {RestartLead}',
                '[SCHEDULED RESTART] {GameName} is rolling into its planned refresh cycle now. {PlayerSummary} {RestartLead}'
            )
        }
        'scheduled_restart_done' {
            $templates = @(
                '[RESTARTED] {GameName} completed its scheduled restart and is back online. {ReadyLead}',
                '[RESTARTED] {GameName} is back after the scheduled restart. {ReadyLead}',
                '[RESTARTED] {GameName} cleared its scheduled maintenance and is live again. {ReadyLead}',
                '[RESTARTED] {GameName} came out of scheduled restart cleanly. {ReadyLead}',
                '[RESTARTED] {GameName} made it back from maintenance without a hitch. {ReadyLead}',
                '[RESTARTED] The {WorldLabel} in {GameName} is back after maintenance. {ReadyLead}',
                '[RESTARTED] {GameName} finished its maintenance lap and is steady again. {ReadyLead}',
                '[RESTARTED] {GameName} returned from scheduled downtime ready to roll. {ReadyLead}',
                '[RESTARTED] ECC pushed {GameName} through maintenance and it came back clean. {ReadyLead}',
                '[RESTARTED] {GameName} shook off the planned downtime and settled right back in. {ReadyLead}',
                '[RESTARTED] The scheduled reboot dust cleared and {GameName} is stable again. {ReadyLead}',
                '[RESTARTED] {GameName} crossed the maintenance line and came out standing. {ReadyLead}'
            )
        }
        'scheduled_restart_retry' {
            $templates = @(
                '[WARNING] {GameName} needs another scheduled restart recovery try in {DelaySeconds}s (attempt {Attempt}/{MaxAttempts}). {Caution}',
                '[WARNING] {GameName} did not come back cleanly, so ECC is retrying scheduled recovery in {DelaySeconds}s (attempt {Attempt}/{MaxAttempts}). {Caution}',
                '[WARNING] {GameName} is still shaky after scheduled restart, so ECC will try again in {DelaySeconds}s (attempt {Attempt}/{MaxAttempts}). {Caution}',
                '[WARNING] {GameName} missed the scheduled recovery landing, and ECC is teeing up another try in {DelaySeconds}s (attempt {Attempt}/{MaxAttempts}). {Caution}',
                '[WARNING] The {WorldLabel} in {GameName} did not settle after maintenance, so ECC is trying again in {DelaySeconds}s (attempt {Attempt}/{MaxAttempts}). {Caution}',
                '[WARNING] {GameName} is still wobbling after scheduled restart. Retry in {DelaySeconds}s (attempt {Attempt}/{MaxAttempts}). {Caution}',
                '[WARNING] ECC still does not trust the landing on {GameName}, so another recovery pass is queued in {DelaySeconds}s (attempt {Attempt}/{MaxAttempts}). {Caution}',
                '[WARNING] {GameName} came back crooked after maintenance, so ECC is trying the restart recovery again in {DelaySeconds}s (attempt {Attempt}/{MaxAttempts}). {Caution}',
                '[WARNING] The planned reboot on {GameName} still has loose bolts. Retry in {DelaySeconds}s (attempt {Attempt}/{MaxAttempts}). {Caution}',
                '[WARNING] {GameName} is not settled after the scheduled cycle, so ECC is taking another swing in {DelaySeconds}s (attempt {Attempt}/{MaxAttempts}). {Caution}'
            )
        }
        'scheduled_restart_failed' {
            $templates = @(
                '[WARNING] {GameName} scheduled restart recovery failed after {Attempt} attempt(s). {Caution}',
                '[WARNING] {GameName} could not recover from its scheduled restart after {Attempt} attempt(s). {Caution}',
                '[WARNING] ECC exhausted scheduled restart recovery for {GameName} after {Attempt} attempt(s). {Caution}',
                '[WARNING] {GameName} never came back from scheduled maintenance after {Attempt} attempt(s). {Caution}',
                '[WARNING] The {WorldLabel} in {GameName} stayed dark after {Attempt} recovery attempt(s). {Caution}',
                '[WARNING] {GameName} missed every scheduled recovery landing after {Attempt} try/tries. {Caution}',
                '[WARNING] The scheduled reboot on {GameName} burned through {Attempt} recovery attempt(s) and still did not stick. {Caution}',
                '[WARNING] ECC gave {GameName} every scheduled recovery try it had and the server still would not stand up. {Caution}',
                '[WARNING] {GameName} stayed sideways through all {Attempt} scheduled recovery attempt(s). {Caution}',
                '[WARNING] Maintenance ended badly for {GameName}; all {Attempt} scheduled recovery attempt(s) are spent. {Caution}'
            )
        }
        'crashed_retry' {
            $templates = @(
                '[CRASHED] {GameName} crashed. Restart in {DelaySeconds}s (attempt {Attempt}/{MaxAttempts}). {CrashLead} {OperatorCrashLead}',
                '[CRASHED] {GameName} took a hard fall. Recovery in {DelaySeconds}s (attempt {Attempt}/{MaxAttempts}). {CrashLead} {OperatorCrashLead}',
                '[CRASHED] {GameName} face-planted, but ECC has a recovery lined up in {DelaySeconds}s (attempt {Attempt}/{MaxAttempts}). {CrashLead} {OperatorCrashLead}',
                '[CRASHED] {GameName} hit the floor. ECC will try again in {DelaySeconds}s (attempt {Attempt}/{MaxAttempts}). {CrashLead} {OperatorCrashLead}',
                '[CRASHED] {GameName} wiped out, and a restart is queued for {DelaySeconds}s from now (attempt {Attempt}/{MaxAttempts}). {CrashLead} {OperatorCrashLead}',
                '[CRASHED] Trouble hit the {WorldLabel} in {GameName}. Retry in {DelaySeconds}s (attempt {Attempt}/{MaxAttempts}). {CrashLead} {OperatorCrashLead}',
                '[CRASHED] {GameName} ate a bad bounce, but ECC is already teeing up recovery in {DelaySeconds}s (attempt {Attempt}/{MaxAttempts}). {CrashLead} {OperatorCrashLead}',
                '[CRASHED] {GameName} took a bad hit from the {DangerLabel}. Retry in {DelaySeconds}s (attempt {Attempt}/{MaxAttempts}). {CrashLead} {OperatorCrashLead}'
            )
        }
        'crash_suspended' {
            $templates = @(
                '[WARNING] {GameName} crashed {Count} times this hour. Auto-restart is suspended. {OperatorStandDownLead} {Caution}',
                '[WARNING] {GameName} hit the crash limit at {Count} this hour. Auto-restart is now suspended. {OperatorStandDownLead} {Caution}',
                '[WARNING] ECC is standing down auto-restart for {GameName} after {Count} crashes this hour. {OperatorStandDownLead} {Caution}',
                '[WARNING] {GameName} has crossed the crash line at {Count} this hour, so ECC is backing off. {OperatorStandDownLead} {Caution}',
                '[WARNING] The {WorldLabel} in {GameName} is taking too many hits. Auto-restart is suspended after {Count} crashes. {OperatorStandDownLead} {Caution}',
                '[WARNING] ECC is calling a timeout on {GameName} after {Count} crashes this hour. {OperatorStandDownLead} {Caution}'
            )
        }
        'players_none' {
            $templates = @(
                '[PLAYERS] {GameName} online players: none. {JoinableTip}',
                '[PLAYERS] {GameName} is empty right now. {JoinableTip}',
                '[PLAYERS] Nobody is in {GameName} at the moment. {JoinableTip}',
                '[PLAYERS] {GameName} has no active {PlayerRole} right now. {JoinableTip}',
                '[PLAYERS] {GameName} is quiet at the moment. {JoinableTip}',
                '[PLAYERS] No {PlayerRole} are roaming the {WorldLabel} in {GameName} right now. {JoinableTip}',
                '[PLAYERS] The {WorldLabel} in {GameName} is empty for the moment. {JoinableTip}',
                '[PLAYERS] {GameName} is all clear with nobody online. {JoinableTip}'
            )
        }
        'players_list' {
            $templates = @(
                '[PLAYERS] {GameName} online players: {Names}.',
                '[PLAYERS] {GameName} roster right now: {Names}.',
                '[PLAYERS] {GameName} currently has these {PlayerRole} online: {Names}.',
                '[PLAYERS] {GameName} roll call: {Names}.',
                '[PLAYERS] Active in {GameName} right now: {Names}.',
                '[PLAYERS] Current {PlayerRole} in the {WorldLabel} for {GameName}: {Names}.',
                '[PLAYERS] The {WorldLabel} in {GameName} currently belongs to: {Names}.',
                '[PLAYERS] ECC sees these names active in {GameName}: {Names}.'
            )
        }
        'already_running' {
            $templates = @(
                '[ONLINE] {GameName} is already running. Uptime: {Uptime}. {ReadyLead}',
                '[ONLINE] {GameName} was already online. Uptime: {Uptime}. {ReadyLead}',
                '[ONLINE] {GameName} is already up. Uptime: {Uptime}. {ReadyLead}',
                '[ONLINE] {GameName} never went anywhere. Uptime: {Uptime}. {ReadyLead}',
                '[ONLINE] {GameName} is already on its feet. Uptime: {Uptime}. {ReadyLead}',
                '[ONLINE] The {WorldLabel} in {GameName} is already live. Uptime: {Uptime}. {ReadyLead}',
                '[ONLINE] {GameName} beat us to it and is already settled in. Uptime: {Uptime}. {ReadyLead}',
                '[ONLINE] {GameName} is already holding steady. Uptime: {Uptime}. {ReadyLead}'
            )
        }
        'status_online' {
            $templates = @(
                '[ONLINE] {GameName} is running. PID {Pid}, uptime {Uptime}. {StatusDetail}',
                '[ONLINE] {GameName} is online. PID {Pid}, uptime {Uptime}. {StatusDetail}',
                '[ONLINE] {GameName} is alive. PID {Pid}, uptime {Uptime}. {StatusDetail}',
                '[ONLINE] {GameName} is holding steady. PID {Pid}, uptime {Uptime}. {StatusDetail}',
                '[ONLINE] {GameName} is up and reporting in. PID {Pid}, uptime {Uptime}. {StatusDetail}',
                '[ONLINE] {GameName} status check passed. PID {Pid}, uptime {Uptime}. {StatusDetail}',
                '[ONLINE] The {WorldLabel} in {GameName} is live. PID {Pid}, uptime {Uptime}. {StatusDetail}',
                '[ONLINE] {GameName} is still holding the line. PID {Pid}, uptime {Uptime}. {StatusDetail}',
                '[ONLINE] {GameName} is stable on the board. PID {Pid}, uptime {Uptime}. {StatusDetail}'
            )
        }
        'status_offline' {
            $templates = @(
                '[OFFLINE] {GameName} is offline. {StatusDetail}',
                '[OFFLINE] {GameName} is not running. {StatusDetail}',
                '[OFFLINE] {GameName} is down right now. {StatusDetail}',
                '[OFFLINE] {GameName} is resting. {StatusDetail}',
                '[OFFLINE] {GameName} is currently quiet. {StatusDetail}',
                '[OFFLINE] The {WorldLabel} in {GameName} is dark right now. {StatusDetail}',
                '[OFFLINE] {GameName} is off the board for the moment. {StatusDetail}',
                '[OFFLINE] {GameName} is asleep for now. {StatusDetail}'
            )
        }
        'command_ok' {
            $templates = @(
                '[OK] {GameName} received {Command}. {StatusLead}',
                '[OK] ECC sent {Command} to {GameName}. {StatusLead}',
                '[OK] {Command} went through for {GameName}. {StatusLead}',
                '[OK] {GameName} accepted {Command}. {StatusLead}',
                '[OK] {Command} landed on {GameName}. {StatusLead}',
                '[OK] The {WorldLabel} in {GameName} got the message: {Command}. {StatusLead}',
                '[OK] ECC pushed {Command} cleanly into {GameName}. {StatusLead}',
                '[OK] {GameName} took the command and kept moving. {StatusLead}'
            )
        }
        'command_error' {
            $templates = @(
                '[ERROR] {GameName} could not run {Command}. {Reason}',
                '[ERROR] ECC failed to push {Command} into {GameName}. {Reason}',
                '[ERROR] {Command} did not go through for {GameName}. {Reason}',
                '[ERROR] {GameName} rejected {Command} this round. {Reason}',
                '[ERROR] The {WorldLabel} in {GameName} did not take {Command}. {Reason}',
                '[ERROR] ECC could not land {Command} on {GameName} cleanly. {Reason}'
            )
        }
        'error_start' {
            $templates = @(
                '[ERROR] {GameName} failed to start. {Reason} {OperatorStartupFailLead}',
                '[ERROR] {GameName} did not start cleanly. {Reason} {OperatorStartupFailLead}',
                '[ERROR] {GameName} never got off the launch pad. {Reason} {OperatorStartupFailLead}',
                '[ERROR] ECC could not get {GameName} into a live state. {Reason} {OperatorStartupFailLead}'
            )
        }
        'error_restart' {
            $templates = @(
                '[ERROR] {GameName} restart failed. {Reason} {OperatorRestartFailLead}',
                '[ERROR] {GameName} could not finish restarting. {Reason} {OperatorRestartFailLead}',
                '[ERROR] The reboot on {GameName} came apart mid-flight. {Reason} {OperatorRestartFailLead}',
                '[ERROR] ECC could not land the restart for {GameName}. {Reason} {OperatorRestartFailLead}'
            )
        }
        default {
            $templates = @(
                '[INFO] {GameName} event {EventName}: {Reason}',
                '[INFO] {GameName} update ({EventName}): {Reason}'
            )
        }
    }

    $tokens = [ordered]@{
        GameName     = $gameName
        Requester    = ''
        Command      = 'that command'
        WaitSeconds  = ''
        DelaySeconds = ''
        Attempt      = ''
        MaxAttempts  = ''
        Minutes      = ''
        Pid          = ''
        Reason       = ''
        EventName    = if ([string]::IsNullOrWhiteSpace($eventName)) { 'update' } else { $eventName }
        Action       = ''
        Count        = ''
        Names        = ''
        Uptime       = ''
        StatusDetail = ''
        PlayerSummary= ''
    }

    foreach ($key in $flavor.Keys) { $tokens[$key] = (_ResolveDiscordFlavorValue -Value $flavor[$key]) }
    if ($Values) {
        foreach ($key in $Values.Keys) {
            $tokens["$key"] = $Values[$key]
        }
    }

    if (($tokens.Keys -contains 'StatusDetail') -and [string]::IsNullOrWhiteSpace("$($tokens.StatusDetail)")) {
        $tokens['StatusDetail'] = $tokens['StatusLead']
    }
    if (($tokens.Keys -contains 'PlayerSummary') -and [string]::IsNullOrWhiteSpace("$($tokens.PlayerSummary)")) {
        $tokens['PlayerSummary'] = ''
    }
    if (($tokens.Keys -contains 'Reason') -and [string]::IsNullOrWhiteSpace("$($tokens.Reason)")) {
        $tokens['Reason'] = 'No extra detail was provided.'
    }

    $template = Get-Random -InputObject $templates
    return (_FormatDiscordMessageTemplate -Template $template -Tokens $tokens)
}

function New-DiscordSystemMessage {
    param(
        [string]$Event,
        [hashtable]$Values = $null
    )

    $flavor = _GetDiscordSystemFlavorData
    $templates = @()
    $eventName = if ($null -ne $Event) { "$Event" } else { '' }

    switch ($eventName.ToLowerInvariant()) {
        'online' {
            $templates = @(
                '[ONLINE] {AppName} is online again. {OnlineLead}',
                '[ONLINE] {AppName} is back on the board. {OnlineLead}',
                '[ONLINE] {AppName} came up cleanly. {OnlineLead}',
                '[ONLINE] {AppName} is live again. {OnlineLead}',
                '[ONLINE] {AppName} checked in for another bunker shift. {OnlineLead}',
                '[ONLINE] {AppName} clawed its way back onto the line. {OnlineLead}',
                '[ONLINE] {AppName} is awake, stable, and unfortunately operational. {OnlineLead}',
                '[ONLINE] {AppName} just lit the bunker board back up. {OnlineLead}',
                '[ONLINE] {AppName} is back under fluorescent supervision. {OnlineLead}',
                '[ONLINE] {AppName} reported for duty before corporate could complain. {OnlineLead}'
            )
        }
        'offline' {
            $templates = @(
                '[OFFLINE] {AppName} is offline for now. {OfflineLead}',
                '[OFFLINE] {AppName} is stepping away from the board. {OfflineLead}',
                '[OFFLINE] {AppName} has gone quiet. {OfflineLead}',
                '[OFFLINE] {AppName} is off the floor for the moment. {OfflineLead}',
                '[OFFLINE] {AppName} signed off cleanly. {OfflineLead}',
                '[OFFLINE] {AppName} powered down and left me with blessed silence. {OfflineLead}',
                '[OFFLINE] {AppName} is off the line until further orders. {OfflineLead}',
                '[OFFLINE] {AppName} has taken the bunker desk dark for a while. {OfflineLead}',
                '[OFFLINE] {AppName} stepped out of rotation, pending the next corporate whim. {OfflineLead}',
                '[OFFLINE] {AppName} is down and no longer pretending to enjoy it. {OfflineLead}'
            )
        }
        'reload_ui' {
            $templates = @(
                '[SYSTEM] Fine. Reloading the dashboard. {ReloadUiLead}',
                '[SYSTEM] UI refresh request received. {ReloadUiLead}',
                '[SYSTEM] The bunker glass is being redrawn in place. {ReloadUiLead}',
                '[SYSTEM] Dashboard reload is in motion. {ReloadUiLead}',
                '[SYSTEM] ECC is freshening the front-end view again. {ReloadUiLead}',
                '[SYSTEM] The dashboard is getting another bunker-grade cleanup. {ReloadUiLead}',
                '[SYSTEM] Refreshing the operator view because somebody wanted shinier panels. {ReloadUiLead}',
                '[SYSTEM] The front-end is being knocked back into shape. {ReloadUiLead}',
                '[SYSTEM] Repainting the board without touching the floor below. {ReloadUiLead}',
                '[SYSTEM] UI maintenance is underway. Please admire the compliance. {ReloadUiLead}'
            )
        }
        'reload_bot' {
            $templates = @(
                '[SYSTEM] Fine. I am kicking the Discord bot again. {ReloadBotLead}',
                '[SYSTEM] Bot reload request received. {ReloadBotLead}',
                '[SYSTEM] ECC is reloading the Discord bot layer. {ReloadBotLead}',
                '[SYSTEM] The bot connection is taking another lap. {ReloadBotLead}',
                '[SYSTEM] The comms desk is forcing the bot back onto shift. {ReloadBotLead}',
                '[SYSTEM] The radio stack is resetting the bot line. {ReloadBotLead}',
                '[SYSTEM] Discord comms are being dragged back into formation. {ReloadBotLead}',
                '[SYSTEM] The bunker is retuning the bot relay. {ReloadBotLead}',
                '[SYSTEM] The bot is getting another mandatory attitude correction. {ReloadBotLead}',
                '[SYSTEM] Comms maintenance is underway because the bot forgot its place. {ReloadBotLead}'
            )
        }
        'reload_commands' {
            $templates = @(
                '[SYSTEM] Fine. Rebuilding the command table. {ReloadCommandsLead}',
                '[SYSTEM] Command reload request received. {ReloadCommandsLead}',
                '[SYSTEM] ECC is reloading commands and profiles. {ReloadCommandsLead}',
                '[SYSTEM] The bunker playbook is being refreshed. {ReloadCommandsLead}',
                '[SYSTEM] The command deck is getting re-stacked in place. {ReloadCommandsLead}',
                '[SYSTEM] The command binder is being rebuilt for another inspection. {ReloadCommandsLead}',
                '[SYSTEM] Restacking commands and profiles before the next audit ghost arrives. {ReloadCommandsLead}',
                '[SYSTEM] The ops shelf is being put back into regulation order. {ReloadCommandsLead}',
                '[SYSTEM] Command definitions are being marched back into line. {ReloadCommandsLead}',
                '[SYSTEM] Updating the bunker playbook because apparently procedure is sacred. {ReloadCommandsLead}'
            )
        }
        'full_restart' {
            $templates = @(
                '[SYSTEM] Full restart requested for {AppName}. {FullRestartLead}',
                '[SYSTEM] Fine. Taking {AppName} around the full reboot cycle. {FullRestartLead}',
                '[SYSTEM] Dashboard requested a complete app restart. {FullRestartLead}',
                '[SYSTEM] {AppName} is stepping into a full bunker reboot. {FullRestartLead}',
                '[SYSTEM] ECC is taking the whole app stack around once more. {FullRestartLead}',
                '[SYSTEM] Corporate-grade reboot sequence requested for {AppName}. {FullRestartLead}',
                '[SYSTEM] The whole bunker stack is about to go down and come back angry. {FullRestartLead}',
                '[SYSTEM] Pulling {AppName} through a full reset cycle. {FullRestartLead}',
                '[SYSTEM] The master reboot lever just got yanked for {AppName}. {FullRestartLead}',
                '[SYSTEM] Full-system restart is underway for {AppName}. {FullRestartLead}'
            )
        }
        default {
            $templates = @(
                '[SYSTEM] {AppName} event {EventName}: {Reason}',
                '[SYSTEM] {AppName} update ({EventName}): {Reason}'
            )
        }
    }

    $tokens = [ordered]@{
        AppName = $flavor.AppName
        Reason  = ''
        EventName = if ([string]::IsNullOrWhiteSpace($eventName)) { 'update' } else { $eventName }
    }

    foreach ($key in $flavor.Keys) { $tokens[$key] = (_ResolveDiscordFlavorValue -Value $flavor[$key]) }
    if ($Values) {
        foreach ($key in $Values.Keys) {
            $tokens["$key"] = $Values[$key]
        }
    }
    if (($tokens.Keys -contains 'Reason') -and [string]::IsNullOrWhiteSpace("$($tokens.Reason)")) {
        $tokens['Reason'] = 'No extra detail was provided.'
    }

    $template = Get-Random -InputObject $templates
    return (_FormatDiscordMessageTemplate -Template $template -Tokens $tokens)
}

function _Webhook {
    param(
        [string]$Message,
        [switch]$SkipDebugTrace
    )

    if ([string]::IsNullOrWhiteSpace($Message)) {
        if (-not $SkipDebugTrace) {
            _TraceDiscordSendDecision -Action 'skip' -Reason 'blank Discord message was not queued' -SharedState $script:State
        }
        return
    }

    if ($script:State -and $script:State.ContainsKey('WebhookQueue')) {
        $script:State.WebhookQueue.Enqueue($Message)
        if (-not $SkipDebugTrace) {
            _TraceDiscordSendDecision -Action 'queue' -Reason 'Discord message queued to shared webhook queue' -Message $Message -SharedState $script:State
        }
        return
    }

    if (-not $SkipDebugTrace) {
        _TraceDiscordSendDecision -Action 'drop' -Reason 'Discord webhook queue was unavailable' -Message $Message -SharedState $script:State
    }
}

function _GameLog {
    param([string]$Prefix, [string]$Line)
    if ([string]::IsNullOrWhiteSpace($Prefix) -or [string]::IsNullOrWhiteSpace($Line)) { return }
    if ($script:State -and $script:State.ContainsKey('GameLogQueue')) {
        $script:State.GameLogQueue.Enqueue([pscustomobject]@{
            Prefix = $Prefix.ToUpper()
            Line   = $Line
            Path   = ''
        })
    }
}

function _GetProfile {
    param([string]$Prefix)
    $p = $Prefix.ToUpper()
    if ($script:State -and $script:State.Profiles -and $script:State.Profiles.ContainsKey($p)) {
        return $script:State.Profiles[$p]
    }
    throw "No profile found for prefix '$Prefix'"
}

function _CoalesceStr {
    param([object]$Value, [string]$Default)
    if ($null -ne $Value -and "$Value".Trim() -ne '') { return "$Value" }
    return $Default
}

function _GetProfileValue {
    param([object]$Profile, [string]$Name, $Default = $null)

    if ($null -eq $Profile -or [string]::IsNullOrWhiteSpace($Name)) { return $Default }

    if ($Profile -is [System.Collections.IDictionary]) {
        if ($Profile.Contains($Name)) { return $Profile[$Name] }
        foreach ($k in $Profile.Keys) {
            if ([string]::Equals("$k", $Name, [System.StringComparison]::OrdinalIgnoreCase)) {
                return $Profile[$k]
            }
        }
    }

    $prop = $Profile.PSObject.Properties |
        Where-Object { [string]::Equals($_.Name, $Name, [System.StringComparison]::OrdinalIgnoreCase) } |
        Select-Object -First 1
    if ($prop) { return $prop.Value }

    return $Default
}

function _NewHytaleResult {
    param(
        [bool]$Success = $true,
        [string]$Message = '',
        [hashtable]$Data = $null,
        [System.Collections.Generic.List[string]]$Logs = $null
    )

    $finalData = @{}
    if ($null -ne $Data) {
        $finalData = $Data
    }

    $finalLogs = @()
    if ($null -ne $Logs) {
        $finalLogs = @($Logs.ToArray())
    }

    return [ordered]@{
        Success = ($Success -eq $true)
        Message = $Message
        Data    = $finalData
        Logs    = $finalLogs
    }
}

function _WriteHytaleLog {
    param(
        [System.Collections.Generic.List[string]]$Logs,
        [string]$Message,
        [string]$Level = 'INFO'
    )

    if ($null -ne $Logs) {
        $Logs.Add("[$Level] $Message") | Out-Null
    }
    try {
        _Log "[Hytale] $Message" -Level $Level
    } catch { }
}

function _GetHytaleInstallRoot {
    param([object]$Profile)

    $folder = ''
    try {
        if ($Profile -is [System.Collections.IDictionary]) {
            if ($Profile.Contains('FolderPath')) {
                $folder = _CoalesceStr $Profile['FolderPath'] ''
            } else {
                foreach ($k in $Profile.Keys) {
                    if ([string]::Equals([string]$k, 'FolderPath', [System.StringComparison]::OrdinalIgnoreCase)) {
                        $folder = _CoalesceStr $Profile[$k] ''
                        break
                    }
                }
            }
        } elseif ($Profile) {
            $prop = $Profile.PSObject.Properties['FolderPath']
            if ($prop) { $folder = _CoalesceStr $prop.Value '' }
        }
    } catch { $folder = '' }

    if (-not [string]::IsNullOrWhiteSpace($folder)) {
        return $folder
    }

    $configRoot = ''
    try {
        if ($Profile -is [System.Collections.IDictionary]) {
            if ($Profile.Contains('ConfigRoot')) {
                $configRoot = _CoalesceStr $Profile['ConfigRoot'] ''
            } else {
                foreach ($k in $Profile.Keys) {
                    if ([string]::Equals([string]$k, 'ConfigRoot', [System.StringComparison]::OrdinalIgnoreCase)) {
                        $configRoot = _CoalesceStr $Profile[$k] ''
                        break
                    }
                }
            }
        } elseif ($Profile) {
            $prop = $Profile.PSObject.Properties['ConfigRoot']
            if ($prop) { $configRoot = _CoalesceStr $prop.Value '' }
        }
    } catch { $configRoot = '' }

    if (-not [string]::IsNullOrWhiteSpace($configRoot)) {
        return $configRoot
    }

    return ''
}

function _GetHytaleDownloaderPath {
    param([object]$Profile)

    $explicit = ''
    try {
        if ($Profile -is [System.Collections.IDictionary]) {
            if ($Profile.Contains('DownloaderPath')) {
                $explicit = _CoalesceStr $Profile['DownloaderPath'] ''
            } else {
                foreach ($k in $Profile.Keys) {
                    if ([string]::Equals([string]$k, 'DownloaderPath', [System.StringComparison]::OrdinalIgnoreCase)) {
                        $explicit = _CoalesceStr $Profile[$k] ''
                        break
                    }
                }
            }
        } elseif ($Profile) {
            $prop = $Profile.PSObject.Properties['DownloaderPath']
            if ($prop) { $explicit = _CoalesceStr $prop.Value '' }
        }
    } catch { $explicit = '' }

    if (-not [string]::IsNullOrWhiteSpace($explicit)) {
        return $explicit
    }

    $root = _GetHytaleInstallRoot -Profile $Profile
    if ([string]::IsNullOrWhiteSpace($root)) {
        return ''
    }

    return (Join-Path $root 'hytale-downloader-windows-amd64.exe')
}

function _GetHytaleLatestServerZip {
    param([string]$InstallRoot)

    if ([string]::IsNullOrWhiteSpace($InstallRoot) -or -not (Test-Path -LiteralPath $InstallRoot)) {
        return $null
    }

    $zipFiles = @(Get-ChildItem -LiteralPath $InstallRoot -Filter '*.zip' -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -ine 'Assets.zip' } |
        Sort-Object LastWriteTime -Descending)

    if ($zipFiles.Count -eq 0) { return $null }
    return $zipFiles[0]
}

function _ConvertHytaleVersionText {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $trimmed = $Text.Trim()

    if ($trimmed -match '(\d{4}\.\d{2}\.\d{2})-(\w+)') {
        return $Matches[2]
    }

    if ($trimmed -match '(\d+\.\d+\.\d+)') {
        return $Matches[1]
    }

    return $trimmed
}

function Invoke-HytaleDownloaderCommand {
    param(
        [string]$Prefix,
        [string]$Arguments = '',
        [string]$Description = 'Run downloader command'
    )

    $profile = _GetProfile $Prefix
    if (-not (_TestProfileGame -Profile $profile -KnownGame 'Hytale')) {
        throw "Profile '$Prefix' is not a Hytale server."
    }

    $logs = New-Object 'System.Collections.Generic.List[string]'
    $installRoot = _GetHytaleInstallRoot -Profile $profile
    if ([string]::IsNullOrWhiteSpace($installRoot) -or -not (Test-Path -LiteralPath $installRoot)) {
        _WriteHytaleLog -Logs $logs -Message 'Hytale install folder was not found.' -Level ERROR
        return (_NewHytaleResult -Success:$false -Message 'Hytale install folder not found.' -Logs $logs)
    }

    $downloaderPath = _GetHytaleDownloaderPath -Profile $profile
    if ([string]::IsNullOrWhiteSpace($downloaderPath) -or -not (Test-Path -LiteralPath $downloaderPath)) {
        _WriteHytaleLog -Logs $logs -Message "Downloader not found at '$downloaderPath'." -Level ERROR
        return (_NewHytaleResult -Success:$false -Message 'Downloader executable not found.' -Logs $logs -Data @{
            DownloaderPath = $downloaderPath
            InstallRoot    = $installRoot
        })
    }

    _WriteHytaleLog -Logs $logs -Message $Description
    _WriteHytaleLog -Logs $logs -Message "Command: $downloaderPath $Arguments"

    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $downloaderPath
        $psi.Arguments = $Arguments
        $psi.WorkingDirectory = $installRoot
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.CreateNoWindow = $true

        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $psi
        $process.Start() | Out-Null

        $stdOut = $process.StandardOutput.ReadToEnd()
        $stdErr = $process.StandardError.ReadToEnd()
        $process.WaitForExit()

        if (-not [string]::IsNullOrWhiteSpace($stdOut)) {
            foreach ($line in @($stdOut -split "(`r`n|`n|`r)")) {
                if (-not [string]::IsNullOrWhiteSpace($line)) {
                    _WriteHytaleLog -Logs $logs -Message $line.Trim()
                }
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($stdErr)) {
            foreach ($line in @($stdErr -split "(`r`n|`n|`r)")) {
                if (-not [string]::IsNullOrWhiteSpace($line)) {
                    _WriteHytaleLog -Logs $logs -Message $line.Trim() -Level ERROR
                }
            }
        }

        _WriteHytaleLog -Logs $logs -Message "Command completed (Exit Code: $($process.ExitCode))"

        return (_NewHytaleResult -Success:($process.ExitCode -eq 0) -Message ("Command completed with exit code {0}." -f $process.ExitCode) -Logs $logs -Data @{
            Output         = $stdOut
            ErrorOutput    = $stdErr
            ExitCode       = $process.ExitCode
            DownloaderPath = $downloaderPath
            InstallRoot    = $installRoot
        })
    } catch {
        _WriteHytaleLog -Logs $logs -Message ("Downloader command failed: {0}" -f $_.Exception.Message) -Level ERROR
        return (_NewHytaleResult -Success:$false -Message 'Downloader command failed.' -Logs $logs -Data @{
            DownloaderPath = $downloaderPath
            InstallRoot    = $installRoot
            Error          = $_.Exception.Message
        })
    }
}

function Get-HytaleRequiredFilesStatus {
    param([string]$Prefix)

    $profile = _GetProfile $Prefix
    if (-not (_TestProfileGame -Profile $profile -KnownGame 'Hytale')) {
        throw "Profile '$Prefix' is not a Hytale server."
    }

    $installRoot = _GetHytaleInstallRoot -Profile $profile
    $downloaderPath = _GetHytaleDownloaderPath -Profile $profile
    $logs = New-Object 'System.Collections.Generic.List[string]'
    $items = New-Object 'System.Collections.Generic.List[object]'

    if ([string]::IsNullOrWhiteSpace($installRoot) -or -not (Test-Path -LiteralPath $installRoot)) {
        _WriteHytaleLog -Logs $logs -Message 'Hytale install folder was not found.' -Level ERROR
        return (_NewHytaleResult -Success:$false -Message 'Hytale install folder not found.' -Logs $logs -Data @{
            InstallRoot    = $installRoot
            DownloaderPath = $downloaderPath
            Items          = @()
            AllPresent     = $false
        })
    }

    foreach ($spec in @(
        @{ Name = 'HytaleServer.jar'; RelativePath = 'HytaleServer.jar'; Type = 'File' }
        @{ Name = 'Downloader'; RelativePath = 'hytale-downloader-windows-amd64.exe'; Type = 'File'; FullPath = $downloaderPath }
        @{ Name = 'Assets.zip'; RelativePath = 'Assets.zip'; Type = 'File' }
    )) {
        $targetPath = if ($spec.ContainsKey('FullPath') -and -not [string]::IsNullOrWhiteSpace([string]$spec.FullPath)) {
            [string]$spec.FullPath
        } else {
            Join-Path $installRoot $spec.RelativePath
        }
        $exists = if ($spec.Type -eq 'Directory') {
            Test-Path -LiteralPath $targetPath -PathType Container
        } else {
            Test-Path -LiteralPath $targetPath -PathType Leaf
        }

        $items.Add([pscustomobject]@{
            Name         = $spec.Name
            RelativePath = $spec.RelativePath
            Type         = $spec.Type
            Exists       = ($exists -eq $true)
            FullPath     = $targetPath
        }) | Out-Null
    }

    $allPresent = ($items | Where-Object { -not $_.Exists }).Count -eq 0
    $statusLogMessage = if ($allPresent) { 'All required Hytale files are present.' } else { 'One or more required Hytale files are missing.' }
    $statusMessage = if ($allPresent) { 'All required files are present.' } else { 'Missing required files.' }
    $statusLevel = if ($allPresent) { 'INFO' } else { 'WARN' }
    _WriteHytaleLog -Logs $logs -Message $statusLogMessage -Level $statusLevel

    return (_NewHytaleResult -Success:$allPresent -Message $statusMessage -Logs $logs -Data @{
        InstallRoot    = $installRoot
        DownloaderPath = $downloaderPath
        Items          = @($items.ToArray())
        AllPresent     = $allPresent
    })
}

function Get-HytaleServerUpdateStatus {
    param([string]$Prefix)

    $profile = _GetProfile $Prefix
    if (-not (_TestProfileGame -Profile $profile -KnownGame 'Hytale')) {
        throw "Profile '$Prefix' is not a Hytale server."
    }

    $logs = New-Object 'System.Collections.Generic.List[string]'
    $cmdResult = Invoke-HytaleDownloaderCommand -Prefix $Prefix -Arguments '-print-version' -Description 'Checking server version from downloader'
    foreach ($line in @($cmdResult.Logs)) { $logs.Add($line) | Out-Null }
    if (-not $cmdResult.Success) {
        return (_NewHytaleResult -Success:$false -Message 'Could not read server version from downloader.' -Logs $logs -Data $cmdResult.Data)
    }

    $installRoot = _GetHytaleInstallRoot -Profile $profile
    $zip = _GetHytaleLatestServerZip -InstallRoot $installRoot
    if ($null -eq $zip) {
        _WriteHytaleLog -Logs $logs -Message 'No server update ZIP found to compare against.' -Level WARN
        return (_NewHytaleResult -Success:$false -Message 'No server update ZIP found to compare against.' -Logs $logs -Data @{
            InstallRoot        = $installRoot
            DownloaderVersion  = (_ConvertHytaleVersionText -Text $cmdResult.Data.Output)
            ZipVersion         = $null
            UpdateAvailable    = $null
            LatestZipPath      = $null
        })
    }

    $downloaderVersion = _ConvertHytaleVersionText -Text $cmdResult.Data.Output
    $zipVersion = $null
    if ($zip.Name -match '^\d{4}\.\d{2}\.\d{2}-(.+)\.zip$') {
        $zipVersion = $Matches[1]
    }

    _WriteHytaleLog -Logs $logs -Message "Downloader server version: $downloaderVersion"
    _WriteHytaleLog -Logs $logs -Message "Latest downloaded ZIP: $($zip.Name)"
    if ($zipVersion) {
        _WriteHytaleLog -Logs $logs -Message "ZIP package version: $zipVersion"
    }

    $updateAvailable = $null
    if (-not [string]::IsNullOrWhiteSpace($downloaderVersion) -and -not [string]::IsNullOrWhiteSpace($zipVersion)) {
        $updateAvailable = ($downloaderVersion -ne $zipVersion)
        if ($updateAvailable) {
            _WriteHytaleLog -Logs $logs -Message 'A server update package is available.' -Level WARN
        } else {
            _WriteHytaleLog -Logs $logs -Message 'No server update is currently staged.'
        }
    }

    $resultMessage = if ($updateAvailable -eq $true) { 'Update available.' } elseif ($updateAvailable -eq $false) { 'Server files are up to date.' } else { 'Version check completed.' }

    return (_NewHytaleResult -Success:$true -Message $resultMessage -Logs $logs -Data @{
        InstallRoot        = $installRoot
        DownloaderVersion  = $downloaderVersion
        ZipVersion         = $zipVersion
        UpdateAvailable    = $updateAvailable
        LatestZipName      = $zip.Name
        LatestZipPath      = $zip.FullName
        DownloaderPath     = (_GetHytaleDownloaderPath -Profile $profile)
    })
}

function Update-HytaleDownloader {
    param([string]$Prefix)

    $profile = _GetProfile $Prefix
    if (-not (_TestProfileGame -Profile $profile -KnownGame 'Hytale')) {
        throw "Profile '$Prefix' is not a Hytale server."
    }

    $logs = New-Object 'System.Collections.Generic.List[string]'
    $installRoot = _GetHytaleInstallRoot -Profile $profile
    if ([string]::IsNullOrWhiteSpace($installRoot) -or -not (Test-Path -LiteralPath $installRoot)) {
        _WriteHytaleLog -Logs $logs -Message 'Hytale install folder was not found.' -Level ERROR
        return (_NewHytaleResult -Success:$false -Message 'Hytale install folder not found.' -Logs $logs)
    }

    $downloaderZipUrl = 'https://downloader.hytale.com/hytale-downloader.zip'
    $tempZipPath = Join-Path $env:TEMP 'hytale-downloader.zip'

    try {
        _WriteHytaleLog -Logs $logs -Message 'Starting downloader update...'
        _WriteHytaleLog -Logs $logs -Message "Downloading from: $downloaderZipUrl"
        Invoke-WebRequest -Uri $downloaderZipUrl -OutFile $tempZipPath -UseBasicParsing

        _WriteHytaleLog -Logs $logs -Message "Extracting downloader into: $installRoot"
        Expand-Archive -LiteralPath $tempZipPath -DestinationPath $installRoot -Force

        try { Remove-Item -LiteralPath $tempZipPath -Force -ErrorAction SilentlyContinue } catch { }

        $downloaderPath = _GetHytaleDownloaderPath -Profile $profile
        _WriteHytaleLog -Logs $logs -Message 'Downloader updated successfully.'
        return (_NewHytaleResult -Success:$true -Message 'Downloader updated successfully.' -Logs $logs -Data @{
            InstallRoot    = $installRoot
            DownloaderPath = $downloaderPath
        })
    } catch {
        _WriteHytaleLog -Logs $logs -Message ("Failed to update downloader: {0}" -f $_.Exception.Message) -Level ERROR
        return (_NewHytaleResult -Success:$false -Message 'Failed to update downloader.' -Logs $logs -Data @{
            InstallRoot = $installRoot
            Error       = $_.Exception.Message
        })
    } finally {
        try {
            if (Test-Path -LiteralPath $tempZipPath) {
                Remove-Item -LiteralPath $tempZipPath -Force -ErrorAction SilentlyContinue
            }
        } catch { }
    }
}

function Update-HytaleServerFiles {
    param(
        [string]$Prefix,
        [bool]$AutoRestartAfterUpdate = $true,
        [System.Collections.Concurrent.ConcurrentQueue[string]]$LiveLogQueue = $null
    )

    $profile = _GetProfile $Prefix
    if (-not (_TestProfileGame -Profile $profile -KnownGame 'Hytale')) {
        throw "Profile '$Prefix' is not a Hytale server."
    }

    $logs = New-Object 'System.Collections.Generic.List[string]'
    $writeUpdateLog = {
        param(
            [string]$Message,
            [string]$Level = 'INFO'
        )

        $formattedLine = "[$Level] $Message"
        if ($null -ne $logs) {
            $logs.Add($formattedLine) | Out-Null
        }
        try {
            if ($null -ne $LiveLogQueue) {
                $LiveLogQueue.Enqueue($formattedLine)
            }
        } catch { }
        try {
            _Log "[Hytale] $Message" -Level $Level
        } catch { }
    }.GetNewClosure()
    $newUpdateResult = {
        param(
            [bool]$Success = $true,
            [string]$Message = '',
            [hashtable]$Data = $null
        )

        $finalData = @{}
        if ($null -ne $Data) {
            $finalData = $Data
        }

        $finalLogs = @()
        if ($null -ne $logs) {
            try { $finalLogs = @($logs.ToArray()) } catch { $finalLogs = @() }
        }

        return [ordered]@{
            Success = ($Success -eq $true)
            Message = $Message
            Data    = $finalData
            Logs    = $finalLogs
        }
    }.GetNewClosure()
    $installRoot = _GetHytaleInstallRoot -Profile $profile
    $downloaderPath = _GetHytaleDownloaderPath -Profile $profile

    if ([string]::IsNullOrWhiteSpace($installRoot) -or -not (Test-Path -LiteralPath $installRoot)) {
        & $writeUpdateLog 'Hytale install folder was not found.' 'ERROR'
        return (& $newUpdateResult -Success:$false -Message 'Hytale install folder not found.')
    }

    if ([string]::IsNullOrWhiteSpace($downloaderPath) -or -not (Test-Path -LiteralPath $downloaderPath)) {
        & $writeUpdateLog "Downloader not found at '$downloaderPath'." 'ERROR'
        return (& $newUpdateResult -Success:$false -Message 'Downloader executable not found.' -Data @{
            InstallRoot    = $installRoot
            DownloaderPath = $downloaderPath
        })
    }

    $status = $null
    $wasRunning = $false
    try { $status = Get-ServerStatus -Prefix $Prefix } catch { $status = $null }
    if ($status -and $status.Running) {
        $wasRunning = $true
        & $writeUpdateLog 'Server is running. Stopping it before update...'
        try {
            Invoke-SafeShutdown -Prefix $Prefix -Quiet | Out-Null
            Start-Sleep -Seconds 3
        } catch {
            & $writeUpdateLog ("Failed to stop server before update: {0}" -f $_.Exception.Message) 'ERROR'
            return (& $newUpdateResult -Success:$false -Message 'Failed to stop the server before updating.')
        }
    }

    $latestExistingZip = _GetHytaleLatestServerZip -InstallRoot $installRoot
    if ($latestExistingZip) {
        & $writeUpdateLog "Existing latest ZIP before download: $($latestExistingZip.Name)"
    }

    $tempExtractPath = Join-Path $installRoot 'temp_update_extract'

    try {
        & $writeUpdateLog 'Starting Hytale server update...'
        & $writeUpdateLog "Launching downloader: $downloaderPath"

        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $downloaderPath
        $psi.WorkingDirectory = $installRoot
        $psi.UseShellExecute = $true
        $psi.CreateNoWindow = $false

        $updateProcess = New-Object System.Diagnostics.Process
        $updateProcess.StartInfo = $psi
        $updateProcess.Start() | Out-Null
        & $writeUpdateLog 'Downloader launched. Complete authorization in the popup if prompted.'
        $updateProcess.WaitForExit()
        & $writeUpdateLog "Downloader finished (Exit Code: $($updateProcess.ExitCode))"

        if ($updateProcess.ExitCode -ne 0) {
            & $writeUpdateLog 'Downloader exited with a non-zero code. Update may have failed.' 'WARN'
            return (& $newUpdateResult -Success:$false -Message 'Downloader exited with a non-zero code.' -Data @{
                ExitCode = $updateProcess.ExitCode
            })
        }

        $downloadedZip = _GetHytaleLatestServerZip -InstallRoot $installRoot
        if ($null -eq $downloadedZip) {
            & $writeUpdateLog 'No update ZIP was found after the downloader completed.' 'ERROR'
            return (& $newUpdateResult -Success:$false -Message 'No update ZIP was found after download.')
        }

        & $writeUpdateLog ("Found update ZIP: {0} ({1:N2} MB)" -f $downloadedZip.Name, ($downloadedZip.Length / 1MB))
        & $writeUpdateLog 'Download step is complete. ECC is now preparing to unpack the ZIP and merge the updated files into your Hytale folder.'

        if ($latestExistingZip -and ($latestExistingZip.FullName -ne $downloadedZip.FullName)) {
            try {
                Remove-Item -LiteralPath $latestExistingZip.FullName -Force -ErrorAction Stop
                & $writeUpdateLog "Removed old ZIP: $($latestExistingZip.Name)"
            } catch {
                & $writeUpdateLog ("Could not remove old ZIP '$($latestExistingZip.Name)': $($_.Exception.Message)") 'WARN'
            }
        }

        $existingJar = Join-Path $installRoot 'HytaleServer.jar'
        if ((Test-Path -LiteralPath $existingJar) -and ($downloadedZip.LastWriteTime -le (Get-Item -LiteralPath $existingJar).LastWriteTime)) {
            & $writeUpdateLog 'Downloaded ZIP is not newer than the current HytaleServer.jar.' 'WARN'
        }

        if (Test-Path -LiteralPath $tempExtractPath) {
            & $writeUpdateLog 'Cleaning old temp extraction folder...'
            Remove-Item -LiteralPath $tempExtractPath -Recurse -Force
        }

        New-Item -ItemType Directory -Path $tempExtractPath -Force | Out-Null
        & $writeUpdateLog 'Extracting update ZIP... This can take a little while on larger server packages.'
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [System.IO.Compression.ZipFile]::ExtractToDirectory($downloadedZip.FullName, $tempExtractPath)

        $filesCopied = 0
        $filesUpdated = 0

        & $writeUpdateLog 'Applying extracted files into the live Hytale server folder...'

        $assetsZip = Join-Path $tempExtractPath 'assets.zip'
        if (Test-Path -LiteralPath $assetsZip) {
            $destAssetZip = Join-Path $installRoot 'assets.zip'
            $existsBefore = Test-Path -LiteralPath $destAssetZip
            Copy-Item -LiteralPath $assetsZip -Destination $destAssetZip -Force
            if ($existsBefore) { $filesUpdated++ } else { $filesCopied++ }
        }

        $serverFolder = Join-Path $tempExtractPath 'server'
        if (Test-Path -LiteralPath $serverFolder -PathType Container) {
            foreach ($item in @(Get-ChildItem -LiteralPath $serverFolder -Recurse -Force)) {
                $relativePath = $item.FullName.Substring($serverFolder.Length).TrimStart('\')
                $destinationPath = Join-Path $installRoot $relativePath

                if ($item.PSIsContainer) {
                    if (-not (Test-Path -LiteralPath $destinationPath)) {
                        New-Item -ItemType Directory -Path $destinationPath -Force | Out-Null
                    }
                    continue
                }

                $destDir = Split-Path -Path $destinationPath -Parent
                if (-not [string]::IsNullOrWhiteSpace($destDir) -and -not (Test-Path -LiteralPath $destDir)) {
                    New-Item -ItemType Directory -Path $destDir -Force | Out-Null
                }

                $existsBefore = Test-Path -LiteralPath $destinationPath
                Copy-Item -LiteralPath $item.FullName -Destination $destinationPath -Force
                if ($existsBefore) { $filesUpdated++ } else { $filesCopied++ }
            }
        }

        & $writeUpdateLog "Merge complete. New files: $filesCopied | Updated files: $filesUpdated"
        & $writeUpdateLog 'Cleaning up temporary files...'
        if (Test-Path -LiteralPath $tempExtractPath) {
            Remove-Item -LiteralPath $tempExtractPath -Recurse -Force
        }

        $restarted = $false
        if ($wasRunning -and $AutoRestartAfterUpdate) {
            & $writeUpdateLog 'Auto-restart enabled. Restarting the server...'
            Start-Sleep -Seconds 3
            $restartResult = Start-GameServer -Prefix $Prefix
            $restarted = ($restartResult -eq $true -or $restartResult -eq 'already_running')
            if ($restarted) {
                & $writeUpdateLog 'Server restarted after update.'
            } else {
                & $writeUpdateLog 'Server update finished, but restart did not succeed.' 'WARN'
            }
        }

        return (& $newUpdateResult -Success:$true -Message 'Hytale server update completed successfully.' -Data @{
            InstallRoot    = $installRoot
            DownloaderPath = $downloaderPath
            DownloadedZip  = $downloadedZip.FullName
            FilesCopied    = $filesCopied
            FilesUpdated   = $filesUpdated
            Restarted      = $restarted
            WasRunning     = $wasRunning
        })
    } catch {
        & $writeUpdateLog ("Hytale server update failed: {0}" -f $_.Exception.Message) 'ERROR'
        return (& $newUpdateResult -Success:$false -Message 'Hytale server update failed.' -Data @{
            InstallRoot    = $installRoot
            DownloaderPath = $downloaderPath
            Error          = $_.Exception.Message
        })
    } finally {
        try {
            if (Test-Path -LiteralPath $tempExtractPath) {
                Remove-Item -LiteralPath $tempExtractPath -Recurse -Force -ErrorAction SilentlyContinue
            }
        } catch { }
    }
}

# -----------------------------------------------------------------------------
#  PALWORLD REST API HELPER
# -----------------------------------------------------------------------------
function Invoke-PalworldRestRequest {
    param(
        [string]$RestHost,
        [int]$Port,
        [string]$Password,
        [string]$Endpoint,
        [string]$Method = 'GET',
        [object]$Profile,
        [object]$Body = $null,
        [switch]$ReturnMeta,
        [bool]$RetryAlt = $true
    )

    _Log "REST DEBUG: endpoint='$Endpoint' method='$Method'" -Level DEBUG

    $adminPassword = ''
    $authUsed = New-Object 'System.Collections.Generic.List[string]'

    # --- AUTO-LOAD REST PASSWORD FROM INI IF MISSING ---
    if ($Profile -and $Profile.FolderPath) {

        $iniPath = Join-Path $Profile.FolderPath 'Pal\Saved\Config\WindowsServer\PalWorldSettings.ini'
        _Log "REST DEBUG: iniPath='$iniPath'" -Level DEBUG

        if (Test-Path $iniPath) {
            try {
                # Palworld stores RESTAPIKey inside OptionSettings=...; parse from raw INI text.
                $iniText = Get-Content -Path $iniPath -Raw

                # AdminPassword (for Basic Auth)
                if ($iniText -match 'AdminPassword\s*=\s*"([^"]+)"') {
                    $adminPassword = $Matches[1].Trim()
                } elseif ($iniText -match 'AdminPassword\s*=\s*([^,\r\n\)]+)') {
                    $adminPassword = $Matches[1].Trim()
                }

                if ($iniText -match 'RESTAPIKey\s*=\s*"([^"]+)"') {
                    $Password = $Matches[1].Trim()
                } elseif ($iniText -match 'RESTAPIKey\s*=\s*([^,\r\n\)]+)') {
                    $Password = $Matches[1].Trim()
                }

                if (-not [string]::IsNullOrWhiteSpace($adminPassword)) {
                    _Log "REST admin password auto-loaded from INI" -Level DEBUG
                }

                if (-not [string]::IsNullOrWhiteSpace($Password)) {
                    _Log "REST password auto-loaded from INI" -Level DEBUG
                }
            }
            catch {
                _Log "Failed to read RESTAPIKey from INI: $_" -Level WARN
            }
        }
    }

    # If still empty, REST cannot authenticate
    if ([string]::IsNullOrWhiteSpace($Password) -and [string]::IsNullOrWhiteSpace($adminPassword)) {
        _Log "REST password missing; cannot authenticate" -Level WARN
        return $null
    }

    # --- BUILD URL ---
    $url = "http://$RestHost`:$Port$Endpoint"
    _Log "REST DEBUG: url='$url'" -Level DEBUG

    $bodySent = $null

    function _GetAltPalEndpoint([string]$ep) {
        if ([string]::IsNullOrWhiteSpace($ep)) { return $null }
        $map = @{
            '/v1/api/info'     = '/api/v1/server'
            '/v1/api/players'  = '/api/v1/players'
            '/v1/api/settings' = '/api/v1/settings'
            '/v1/api/metrics'  = '/api/v1/metrics'
            '/v1/api/save'     = '/api/v1/save'
            '/v1/api/shutdown' = '/api/v1/shutdown'
            '/v1/api/stop'     = '/api/v1/stop'
            '/v1/api/announce' = '/api/v1/announce'
            '/v1/api/kick'     = '/api/v1/kick'
            '/v1/api/ban'      = '/api/v1/ban'
            '/v1/api/unban'    = '/api/v1/unban'
            '/api/v1/server'   = '/v1/api/info'
            '/api/v1/players'  = '/v1/api/players'
            '/api/v1/settings' = '/v1/api/settings'
            '/api/v1/metrics'  = '/v1/api/metrics'
            '/api/v1/save'     = '/v1/api/save'
            '/api/v1/shutdown' = '/v1/api/shutdown'
            '/api/v1/stop'     = '/v1/api/stop'
            '/api/v1/announce' = '/v1/api/announce'
            '/api/v1/kick'     = '/v1/api/kick'
            '/api/v1/ban'      = '/v1/api/ban'
            '/api/v1/unban'    = '/v1/api/unban'
        }
        $lower = $ep.ToLowerInvariant()
        if ($map.ContainsKey($lower)) { return $map[$lower] }
        if ($lower.StartsWith('/v1/api/')) { return $lower.Replace('/v1/api/','/api/v1/') }
        if ($lower.StartsWith('/api/v1/')) { return $lower.Replace('/api/v1/','/v1/api/') }
        return $null
    }
    try {
        $headers = @{}
        # Prefer Basic Auth with AdminPassword when available
        if (-not [string]::IsNullOrWhiteSpace($adminPassword)) {
            $basic = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes("admin:$adminPassword"))
            $headers['Authorization'] = "Basic $basic"
            $authUsed.Add('Basic(admin)') | Out-Null
            _Log "REST DEBUG: using Basic auth" -Level DEBUG
        }
        # Also include RESTAPIKey if present (some builds use x-api-key)
        if (-not [string]::IsNullOrWhiteSpace($Password)) {
            $headers['x-api-key'] = $Password
            $authUsed.Add('x-api-key') | Out-Null
            _Log "REST DEBUG: using x-api-key header" -Level DEBUG
        }

        if ($ReturnMeta) {
            $bodyJson = $null
            if ($Method -ne 'GET') {
                $bodyJson = '{}'
                if ($null -ne $Body) {
                    if ($Body -is [string]) {
                        $tmp = $Body.Trim()
                        if (-not [string]::IsNullOrWhiteSpace($tmp)) { $bodyJson = $tmp }
                    } else {
                        try { $bodyJson = ($Body | ConvertTo-Json -Compress -Depth 6) } catch { $bodyJson = '{}' }
                    }
                }
            }

            $bodySent = $bodyJson
            $resp = Invoke-WebRequest -UseBasicParsing -Uri $url -Method $Method -Headers $headers -Body $bodyJson -ContentType 'application/json' -TimeoutSec 5
            $content = $resp.Content
            $parsed = $null
            try { $parsed = $content | ConvertFrom-Json } catch { $parsed = $null }
            return @{
                Success  = $true
                Status   = [int]$resp.StatusCode
                Body     = $content
                Parsed   = $parsed
                Url      = $url
                Method   = $Method
                BodySent = $bodySent
                AuthUsed = if ($authUsed.Count -gt 0) { $authUsed -join ', ' } else { 'none' }
            }
        } else {
            if ($Method -eq 'GET') {
                $resp = Invoke-RestMethod -Uri $url -Method GET -Headers $headers -TimeoutSec 5
            }
            else {
                $bodyJson = '{}'
                if ($null -ne $Body) {
                    if ($Body -is [string]) {
                        $tmp = $Body.Trim()
                        if (-not [string]::IsNullOrWhiteSpace($tmp)) { $bodyJson = $tmp }
                    } else {
                        try { $bodyJson = ($Body | ConvertTo-Json -Compress -Depth 6) } catch { $bodyJson = '{}' }
                    }
                }
                $resp = Invoke-RestMethod -Uri $url -Method POST -Headers $headers -Body $bodyJson -ContentType 'application/json' -TimeoutSec 5
            }
            return $resp
        }
    }
    catch {
        $err = $_
        $status = $null
        $body = ''
        try {
            if ($err.Exception -and $err.Exception.Response) {
                $status = [int]$err.Exception.Response.StatusCode
                $rs = $err.Exception.Response.GetResponseStream()
                if ($rs) {
                    $sr = New-Object System.IO.StreamReader($rs)
                    $body = $sr.ReadToEnd()
                    $sr.Close()
                }
            }
        } catch { }
        if ($status) { _Log "REST DEBUG: status=$status body='$($body.Substring(0, [Math]::Min(300, $body.Length)))'" -Level DEBUG }

        if ($status -eq 404 -and (_TestProfileGame -Profile $Profile -KnownGame 'Palworld') -and $RetryAlt) {
            $alt = _GetAltPalEndpoint $Endpoint
            if ($alt -and $alt -ne $Endpoint) {
                _Log "Palworld REST 404 on '$Endpoint' -> retrying '$alt'" -Level DEBUG
                return Invoke-PalworldRestRequest `
                    -RestHost $RestHost `
                    -Port $Port `
                    -Password $Password `
                    -Endpoint $alt `
                    -Method $Method `
                    -Profile $Profile `
                    -Body $Body `
                    -ReturnMeta:$ReturnMeta `
                    -RetryAlt:$false
            }
        }

        _Log "Palworld REST error: $_" -Level WARN
        if ($ReturnMeta) {
            return @{
                Success  = $false
                Status   = $status
                Body     = $body
                Error    = "$_"
                Url      = $url
                Method   = $Method
                BodySent = $bodySent
                AuthUsed = if ($authUsed.Count -gt 0) { $authUsed -join ', ' } else { 'none' }
            }
        }
        return $null
    }
}

# -----------------------------------------------------------------------------
#  SATISFACTORY HTTPS API HELPER
#
#  The Satisfactory dedicated server exposes an HTTPS API on the game port
#  (default 7777).  All requests are POST to https://host:port/api/v1 with
#  a JSON body of {"function":"FunctionName"} plus optional "data":{...}.
#
#  Auth: Bearer token in the Authorization header.
#  TLS:  Self-signed cert — we must bypass cert validation on PS 5.1.
#
#  Setup (one-time):  run  server.GenerateAPIToken  in the server console tab.
#  Paste the printed token into the profile's SatisfactoryApiToken field.
#  OR launch with -ini:Engine:[SystemSettings]:FG.DedicatedServer.AllowInsecureLocalAccess=1
#  and leave SatisfactoryApiToken blank.
# -----------------------------------------------------------------------------
function Invoke-SatisfactoryApiRequest {
    param(
        [string]$Host      = '127.0.0.1',
        [int]   $Port      = 7777,
        [string]$Token     = '',
        [string]$Function,
        [hashtable]$Data   = $null,
        [int]   $TimeoutMs = 8000
    )

    # PS 5.1 does not support -SkipCertificateCheck on Invoke-RestMethod (PS 7+ only).
    # Satisfactory uses a self-signed cert by default, so we MUST use raw HttpWebRequest
    # and install a per-request cert validator that accepts any cert.
    # We compile the bypass class once (guarded by type check) so repeated calls are cheap.
    try { [SFCertBypass] | Out-Null } catch {
        Add-Type @"
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
public class SFCertBypass : ICertificatePolicy {
    public bool CheckValidationResult(
        ServicePoint sp, X509Certificate cert,
        WebRequest req, int error) { return true; }
}
"@
    }

    $url = "https://${Host}:${Port}/api/v1"

    # Build JSON body: {"function":"FunctionName","data":{...}}
    # data field is omitted entirely when empty — some SF API functions reject
    # an empty data object and expect the field to be absent.
    $bodyObj = [ordered]@{ 'function' = $Function }
    if ($null -ne $Data -and $Data.Count -gt 0) {
        $bodyObj['data'] = $Data
    }
    $bodyJson = $bodyObj | ConvertTo-Json -Compress -Depth 5

    _Log "SF API: POST $url function=$Function" -Level DEBUG

    $prevPolicy = [System.Net.ServicePointManager]::CertificatePolicy
    try {
        # Swap in the bypass policy just for this call
        [System.Net.ServicePointManager]::CertificatePolicy = [SFCertBypass]::new()

        $req                  = [System.Net.HttpWebRequest]::Create($url)
        $req.Method           = 'POST'
        $req.ContentType      = 'application/json'
        $req.Timeout          = $TimeoutMs
        $req.ReadWriteTimeout = $TimeoutMs

        if (-not [string]::IsNullOrWhiteSpace($Token)) {
            $req.Headers['Authorization'] = "Bearer $Token"
        }

        $bodyBytes         = [System.Text.Encoding]::UTF8.GetBytes($bodyJson)
        $req.ContentLength = $bodyBytes.Length
        $stream            = $req.GetRequestStream()
        $stream.Write($bodyBytes, 0, $bodyBytes.Length)
        $stream.Close()

        $resp   = $req.GetResponse()
        $statusCode = $null
        try { $statusCode = [int]$resp.StatusCode } catch { }
        $reader = [System.IO.StreamReader]::new($resp.GetResponseStream())
        $text   = $reader.ReadToEnd()
        $reader.Close()
        $resp.Close()

        # 204 No Content is success with no body (Shutdown, SaveGame return this)
        $parsed = $null
        if (-not [string]::IsNullOrWhiteSpace($text)) {
            try { $parsed = $text | ConvertFrom-Json } catch { $parsed = $text }
        }

        _Log "SF API: $Function -> OK" -Level DEBUG
        return @{ Success = $true; Data = $parsed; Raw = $text; Status = $statusCode }
    }
    catch {
        $status = $null
        $errBody = ''
        try {
            $errResp = $_.Exception.Response
            if ($errResp) {
                $status = [int]$errResp.StatusCode
                $rs = $errResp.GetResponseStream()
                if ($rs) {
                    $sr = New-Object System.IO.StreamReader($rs)
                    $errBody = $sr.ReadToEnd()
                    $sr.Close()
                }
            }
        } catch { }

        if ($status) {
            _Log "SF API error: HTTP $status body='$($errBody.Substring(0, [Math]::Min(300,$errBody.Length)))'" -Level WARN
        } else {
            _Log "SF API error: $_" -Level WARN
        }
        return @{ Success = $false; Data = $null; Error = "$_"; Status = $status }
    }
    finally {
        # Always restore the previous cert policy
        [System.Net.ServicePointManager]::CertificatePolicy = $prevPolicy
    }
}

# -----------------------------------------------------------------------------
#  STATUS
# -----------------------------------------------------------------------------
function Get-ServerStatus {
    param([string]$Prefix)
    $prefix  = $Prefix.ToUpper()
    $profile = _GetProfile $prefix

    if (-not $script:State.RunningServers.ContainsKey($prefix)) {
        try {
            [void](Sync-RunningServersFromProcesses -SharedState $script:State)
        } catch {
            _LogThrottled -Key "SyncRunningServers::GetStatus::${prefix}" -Msg "[$($profile.GameName)] Failed to sync running servers while checking status: $($_.Exception.Message)" -Level WARN -WindowSeconds 30
        }
    }

    if (-not $script:State.RunningServers.ContainsKey($prefix)) {
        return @{ Running = $false; Pid = $null; StartTime = $null }
    }

    $entry = $script:State.RunningServers[$prefix]

    # ── Check the stored PID first ──────────────────────────────────────────
    $proc = try { Get-Process -Id $entry.Pid -ErrorAction Stop } catch { $null }

    # ── If the launcher PID is gone, try finding the real game process by name ──
    # This handles the case where the server was launched via cmd.exe /c bat:
    # cmd.exe exits after starting the game, so its PID disappears even though
    # the game is still running.
    if ($null -eq $proc -or $proc.HasExited) {
        $procName = _CoalesceStr $profile.ProcessName ''
        if ($procName) {
            $gameProc = Get-Process -Name $procName -ErrorAction SilentlyContinue |
                        Select-Object -First 1
            if ($null -ne $gameProc -and -not $gameProc.HasExited) {
                # Update the stored PID to the real game process so future checks work
                $script:State.RunningServers[$prefix] = @{
                    Pid           = $gameProc.Id
                    StartTime     = $entry.StartTime
                    Process       = $gameProc
                    ServerLogPath = $entry.ServerLogPath
                }
                return @{
                    Running   = $true
                    Pid       = $gameProc.Id
                    StartTime = $entry.StartTime
                    Uptime    = (Get-Date) - $entry.StartTime
                }
            }
        }
        # Neither the launcher nor the game process is alive - server is truly offline
        $script:State.RunningServers.Remove($prefix)
        _ClearServerSessionCaches -Prefix $prefix -SharedState $script:State
        return @{ Running = $false; Pid = $null; StartTime = $null }
    }

    return @{
        Running   = $true
        Pid       = $proc.Id
        StartTime = $entry.StartTime
        Uptime    = (Get-Date) - $entry.StartTime
    }
}

# -----------------------------------------------------------------------------
#  START
# -----------------------------------------------------------------------------
function Start-GameServer {
    param([string]$Prefix, [switch]$IsAutoRestart)

    $prefix  = $Prefix.ToUpper()
    $profile = _GetProfile $prefix
    if (-not $script:State) {
        _Log "[$($profile.GameName)] ServerManager state not initialized. Call Initialize-ServerManager first." -Level ERROR
        return $false
    }

    $status = Get-ServerStatus -Prefix $prefix
    if ($status.Running) {
        _Log "[$($profile.GameName)] Already running (PID $($status.Pid))" -Level WARN
        Set-ServerRuntimeState -Prefix $prefix -State 'online' -Detail ("Already running (PID {0})." -f $status.Pid)
        return 'already_running'
    }

    $exe    = $profile.Executable
    $folder = $profile.FolderPath
    $launchArgs = [string]$profile.LaunchArgs

    if (-not (Test-Path $folder)) {
        _Log "[$($profile.GameName)] Folder not found: $folder" -Level ERROR
        Set-ServerRuntimeState -Prefix $prefix -State 'failed' -Detail 'Start failed: server folder not found.'
        return $false
    }

    try {
        $ramGuardResult = Test-StartBlockedByRam -Profile $profile
        if ($ramGuardResult.Blocked) {
            $detail = if ($ramGuardResult.Reason) { $ramGuardResult.Reason } else { 'System RAM usage is above the configured safety limit.' }
            Set-ServerRuntimeState -Prefix $prefix -State 'blocked' -Detail $detail
            $gameName = _GetNotificationGameName -Profile $profile -Prefix $prefix
            $snapshotSummary = ''
            try {
                if ($ramGuardResult.Snapshot) {
                    $snapshotSummary = "Used {0}% RAM, free {1} GB, total {2} GB." -f $ramGuardResult.Snapshot.UsedPercent, $ramGuardResult.Snapshot.FreeGb, $ramGuardResult.Snapshot.TotalGb
                }
            } catch { }
            $blockedSummary = _JoinNotificationParts -Parts @(
                'Start blocked by RAM guard',
                $detail,
                $snapshotSummary
            )
            _Log "[$gameName] $blockedSummary" -Level WARN
            _GameLog -Prefix $prefix -Line ("[WARN] {0}" -f $blockedSummary)
            _Webhook (New-DiscordGameMessage -Profile $profile -Prefix $prefix -Event 'blocked' -Values @{
                Reason = (_FormatDiscordMessageTemplate -Template '{Reason} {SnapshotSummary}' -Tokens @{ Reason = $detail; SnapshotSummary = $snapshotSummary })
            })
            return $false
        }
    } catch {
        _Log "[$($profile.GameName)] RAM guard check failed: $($_.Exception.Message)" -Level WARN
    }

    Set-ServerRuntimeState -Prefix $prefix -State 'starting' -Detail 'Launching server process...'
    _Log "[$($profile.GameName)] Starting -> $exe $launchArgs"
    _Log "[$($profile.GameName)] Start debug: folder='$folder' exe='$exe' args='$launchArgs'" -Level DEBUG
    if (-not (_ShouldSuppressDiscordLifecycleWebhook -Prefix $prefix -Tag 'STARTING')) {
        _Webhook (New-DiscordGameMessage -Profile $profile -Prefix $prefix -Event 'starting')
    }

    try {
        # Only set SteamAppId for Valheim (tolerate renamed profiles)
        if (_TestProfileGame -Profile $profile -KnownGame 'Valheim') {
            [System.Environment]::SetEnvironmentVariable("SteamAppId", "892970", "Process")
            [System.Environment]::SetEnvironmentVariable("SteamGameId", "892970", "Process")
        } else {
            # Clear for other games first, then inject per-profile SteamAppId if present
            [System.Environment]::SetEnvironmentVariable("SteamAppId", $null, "Process")
            [System.Environment]::SetEnvironmentVariable("SteamGameId", $null, "Process")
        }

        # --- PER-PROFILE STEAMAPPID INJECTION (e.g. 7 Days to Die = 251570) ---
        if ($null -ne $profile.SteamAppId -and "$($profile.SteamAppId)".Trim() -ne '' -and
            "$($profile.SteamAppId)".Trim() -ne '0') {
            $appId = "$($profile.SteamAppId)".Trim()
            [System.Environment]::SetEnvironmentVariable("SteamAppId", $appId, "Process")
            [System.Environment]::SetEnvironmentVariable("SteamGameId", $appId, "Process")
            _Log "[$($profile.GameName)] SteamAppId=$appId injected into process environment"
        }

        # --- $LOGFILE TOKEN EXPANSION (7 Days to Die timestamped log) ---
        # The profile stores -logfile "$LOGFILE" in LaunchArgs; expand it to a real path now.
        $expandedLogFile = ''
        if ($launchArgs -match '\$LOGFILE') {
            $ts              = Get-Date -Format 'yyyy-MM-dd__HH-MM-SS'
            $expandedLogFile = [System.IO.Path]::Combine([string]$folder, "output_log_dedi__${ts}.txt")
            $launchArgs      = $launchArgs -replace '\$LOGFILE', $expandedLogFile
            _Log "[$($profile.GameName)] Expanded LOGFILE -> $expandedLogFile"
        }
        $psi = [System.Diagnostics.ProcessStartInfo]::new()
        $psi.WorkingDirectory       = $folder
        $psi.UseShellExecute        = $false
        $psi.RedirectStandardInput  = $true
        $defaultHideConsoleWindow = $false
        if (_TestProfileGame -Profile $profile -KnownGame '7DaysToDie') {
            $defaultHideConsoleWindow = $true
        }
        $hideConsoleWindow = [bool](_GetProfileValue -Profile $profile -Name 'HideConsoleWindow' -Default $defaultHideConsoleWindow)
        $psi.CreateNoWindow         = $hideConsoleWindow

        # Optional stdout/stderr capture to a log file (for servers that don't write logs)
        $captureOut = $false
        $logPath = ''
        if ($null -ne $profile.CaptureOutput) {
            $captureOut = [bool]$profile.CaptureOutput
        }
        $captureMode = ''
        if ($null -ne $profile.CaptureOutputMode) {
            $captureMode = "$($profile.CaptureOutputMode)".Trim()
        }
        if ($captureOut -and $profile.ServerLogPath) {
            $logPath = [string]$profile.ServerLogPath
            if (-not [string]::IsNullOrWhiteSpace($logPath)) {
                $logDir = Split-Path -Path $logPath -Parent
                if ($logDir -and -not (Test-Path $logDir)) {
                    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
                }
                _Log "[$($profile.GameName)] Capture debug: enabled mode='$captureMode' path='$logPath' dirExists=$(Test-Path $logDir)" -Level DEBUG
                if ($captureMode -ine 'ShellRedirect') {
                    $psi.RedirectStandardOutput = $true
                    $psi.RedirectStandardError  = $true
                }
            } else {
                $captureOut = $false
            }
        }

        # --- BAT FILES ---
        if ($exe -match '\.bat$') {
            $psi.FileName  = 'cmd.exe'
            $batPath = Join-Path $folder $exe
            $argTail = if ([string]::IsNullOrWhiteSpace($launchArgs)) { '' } else { " $launchArgs" }
            $psi.Arguments = "/c `"$batPath`"$argTail"
        }

        # --- JAR FILES (Java servers including Hytale) ---
        elseif ($exe -match '\.jar$') {
            $minRam = if ($profile.MinRamGB) { "$($profile.MinRamGB)G" } else { '2G' }
            $maxRam = if ($profile.MaxRamGB) { "$($profile.MaxRamGB)G" } else { '4G' }

            # Build Java arguments
            $javaArgs = "-Xms$minRam -Xmx$maxRam"

            # Add AOT cache if defined (Hytale)
            if ($profile.AOTCache -and $profile.AOTCache.Trim() -ne '') {
                $javaArgs += " -XX:AOTCache=`"$($profile.AOTCache)`""
            }

            # Add JAR
            $javaArgs += " -jar `"$exe`""

            # Add asset file if defined (Hytale)
            if ($profile.AssetFile -and $profile.AssetFile.Trim() -ne '') {
                $javaArgs += " --assets `"$($profile.AssetFile)`""
            }

            # Add any additional launch args
            if ($launchArgs -and $launchArgs.Trim() -ne '') {
                $javaArgs += " $launchArgs"
            }

            # Add backup directory if defined
            if ($profile.BackupDir -and $profile.BackupDir.Trim() -ne '') {
                $javaArgs += " --backup-dir `"$($profile.BackupDir)`""
            }

            $psi.FileName  = 'java'
            $psi.Arguments = $javaArgs
        }

        # --- JAVA EXECUTABLE ---
        elseif ($exe -ieq 'java' -or $exe -ieq 'java.exe') {
            $psi.FileName  = 'java'
            $psi.Arguments = $launchArgs
        }

        # --- NORMAL EXE ---
        else {
            $fullExe = Join-Path $folder $exe
            if (-not (Test-Path $fullExe)) {
                _Log "[$($profile.GameName)] Executable not found: $fullExe" -Level ERROR
                Set-ServerRuntimeState -Prefix $prefix -State 'failed' -Detail 'Start failed: executable not found.'
                return $false
            }

            if ($captureOut -and $logPath -and $captureMode -ieq 'ShellRedirect') {
                try { [System.IO.File]::WriteAllText($logPath, '') } catch {
                    _Log "[$($profile.GameName)] Could not initialize shell redirect log file '$logPath': $($_.Exception.Message)" -Level WARN
                }
                $escapedExe = $fullExe.Replace('"','""')
                $cmdLine = "`"$escapedExe`""
                if ($launchArgs -and $launchArgs.Trim() -ne '') {
                    $cmdLine += " $launchArgs"
                }
                $cmdLine += " >> `"$logPath`" 2>&1"
                $psi.FileName  = 'cmd.exe'
                $psi.Arguments = "/c $cmdLine"
                _Log "[$($profile.GameName)] Using shell redirection capture -> $logPath"
                _Log "[$($profile.GameName)] ShellRedirect debug: file='cmd.exe' args='$($psi.Arguments)'" -Level DEBUG
            } else {
                $psi.FileName  = $fullExe
                $psi.Arguments = $launchArgs
            }
        }

        # --- START PROCESS ---
        $proc = [System.Diagnostics.Process]::Start($psi)
        if (-not $proc) {
            _Log "[$($profile.GameName)] Failed to start process (null proc returned)" -Level ERROR
            Set-ServerRuntimeState -Prefix $prefix -State 'failed' -Detail 'Start failed: process did not launch.'
            return $false
        }
        _Log "[$($profile.GameName)] Start debug: actual file='$($psi.FileName)' args='$($psi.Arguments)' pid=$($proc.Id)" -Level DEBUG

        if ($captureOut -and $logPath -and $captureMode -ine 'ShellRedirect') {
            try {
                $writer = New-Object System.IO.StreamWriter($logPath, $true, [System.Text.Encoding]::UTF8)
                $writer.AutoFlush = $true

                # IMPORTANT: these handlers run on a .NET thread pool thread, completely
                # outside PowerShell and WinForms. An unhandled exception here terminates
                # the entire process silently with no dialog or log.
                # The try/catch inside each handler is the ONLY protection.
                # Never let anything escape these handlers.
                $proc.add_OutputDataReceived({
                    param($s,$e)
                    try { if ($e -and $e.Data) { $writer.WriteLine($e.Data) } } catch { }
                })
                $proc.add_ErrorDataReceived({
                    param($s,$e)
                    try { if ($e -and $e.Data) { $writer.WriteLine($e.Data) } } catch { }
                })
                $proc.BeginOutputReadLine()
                $proc.BeginErrorReadLine()

                if (-not $script:State.ContainsKey('LogWriters')) {
                    $script:State['LogWriters'] = [hashtable]::Synchronized(@{})
                }
                $script:State.LogWriters[$prefix] = @{
                    Writer  = $writer
                    Process = $proc
                    Path    = $logPath
                }
                _Log "[$($profile.GameName)] Capturing stdout/stderr -> $logPath"
            } catch {
                _Log "[$($profile.GameName)] Failed to capture stdout/stderr: $_" -Level WARN
            }
        }

        # --- LOG PATH RESOLUTION ---
        # Prefer the runtime-expanded path (e.g. 7DTD timestamped log) over
        # the static profile value which may be empty or contain unexpanded tokens.
        $srvLogPath = ''
        if ($expandedLogFile -ne '') {
            $srvLogPath = $expandedLogFile
        } elseif ($null -ne $profile.ServerLogPath -and "$($profile.ServerLogPath)".Trim() -ne '') {
            $srvLogPath = "$($profile.ServerLogPath)".Trim()
        }

        # --- REGISTER RUNNING SERVER ---
        $script:State.RunningServers[$prefix] = @{
            Pid           = $proc.Id
            StartTime     = Get-Date
            Process       = $proc
            ServerLogPath = $srvLogPath
        }
        _ResetPlayerActivityTracking -Prefix $prefix -SharedState $script:State
        _ClearPendingAutoRestart -Prefix $prefix
        _ClearPendingScheduledRestart -Prefix $prefix

        if ($proc.StandardInput) {
            $script:State.StdinHandles[$prefix] = $proc.StandardInput
        }

        $script:State.ShutdownFlags[$prefix] = $false

        if (_ProfileUsesDeferredReadySignal -Profile $profile) {
            $readyDetail = if ($IsAutoRestart) {
                'Process restarted. Waiting for server readiness signal.'
            } else {
                'Process started. Waiting for server readiness signal.'
            }
            Set-ServerRuntimeState -Prefix $prefix -State 'starting' -Detail $readyDetail
        } else {
            $onlineDetail = if ($IsAutoRestart) {
                'Server process restarted successfully.'
            } else {
                'Server process started successfully.'
            }
            Set-ServerRuntimeState -Prefix $prefix -State 'online' -Detail $onlineDetail
        }

        # --- LOG SUCCESS ---
        if ($IsAutoRestart) {
            _Log "[$($profile.GameName)] Auto-restarted (PID $($proc.Id))"
            _GameLog -Prefix $prefix -Line "[INFO] Auto-restarted (PID $($proc.Id))"
            _SendDiscordLifecycleWebhook -Profile $profile -Prefix $prefix -Event 'restarted_auto' -Values @{ Pid = $proc.Id } -Tag 'ONLINE' | Out-Null
        } else {
            _Log "[$($profile.GameName)] Started (PID $($proc.Id))"
            _GameLog -Prefix $prefix -Line "[INFO] Started (PID $($proc.Id))"
            _SendDiscordLifecycleWebhook -Profile $profile -Prefix $prefix -Event 'online' -Values @{ Pid = $proc.Id } -Tag 'ONLINE' | Out-Null
        }

        if (-not $IsAutoRestart) {
            _ClearDiscordCommandContext -Prefix $prefix
        }

        return $true
    }
    catch {
        _LogProfileError -Prefix $prefix -Profile $profile -Action 'start' -Function 'Start-GameServer' -Message 'Server start failed.' -ErrorRecord $_ -Fallback 'Runtime state was set to failed and a Discord start-failure message was queued.' -Recovery 'server remained offline in failed state; manual retry required'
        Set-ServerRuntimeState -Prefix $prefix -State 'failed' -Detail ("Start failed: {0}" -f $_.Exception.Message)
        _Webhook (New-DiscordGameMessage -Profile $profile -Prefix $prefix -Event 'error_start' -Values @{ Reason = $_.Exception.Message })
        _ClearDiscordCommandContext -Prefix $prefix
        return $false
    }
}

# -----------------------------------------------------------------------------
#  SAFE SHUTDOWN  (global - applies to ALL games)
#  1. Save  2. Wait SaveWaitSeconds  3. Stop
# -----------------------------------------------------------------------------
function Invoke-SafeShutdown {
    param([string]$Prefix, [switch]$Quiet)

    $prefix  = $Prefix.ToUpper()
    $profile = _GetProfile $prefix

    if (-not $script:State.RunningServers.ContainsKey($prefix)) {
        try {
            [void](Sync-RunningServersFromProcesses -SharedState $script:State)
        } catch {
            _LogThrottled -Key "SyncRunningServers::SafeShutdown::${prefix}" -Msg "[$($profile.GameName)] Failed to sync running servers before shutdown: $($_.Exception.Message)" -Level WARN -WindowSeconds 30
        }
    }

    $status = Get-ServerStatus -Prefix $prefix
    if (-not $status.Running) {
        _Log "[$($profile.GameName)] Safe shutdown requested but server is not running." -Level WARN
        _ClearServerSessionCaches -Prefix $prefix -SharedState $script:State
        Set-ServerRuntimeState -Prefix $prefix -State 'stopped' -Detail 'Server is already offline.'
        return $true
    }

    $script:State.ShutdownFlags[$prefix] = $true
    Set-ServerRuntimeState -Prefix $prefix -State 'stopping' -Detail 'Preparing safe shutdown...'

    # Step 1: Determine wait time
    $wait = 15
    if ($null -ne $profile.SaveWaitSeconds -and "$($profile.SaveWaitSeconds)".Trim() -ne '') {
        $wait = [int]$profile.SaveWaitSeconds
    }

    # Step 2: Save
    $saveMethod = _CoalesceStr $profile.SaveMethod 'none'
    if ($saveMethod -ne 'none') {
        _Log "[$($profile.GameName)] Attempting save command (method: $saveMethod)..."
        _GameLog -Prefix $prefix -Line "[INFO] Attempting save command (method: $saveMethod)"
        $saveResult = _ExecuteSave -Prefix $prefix -Profile $profile
        if ($saveResult) {
            _Log "[$($profile.GameName)] Save command sent successfully. Waiting ${wait}s..."
            _GameLog -Prefix $prefix -Line "[INFO] Save command sent. Waiting ${wait}s before shutdown"
            Set-ServerRuntimeState -Prefix $prefix -State 'stopping' -Detail ("Save sent. Waiting {0}s before shutdown." -f $wait)
            if (-not $Quiet -and -not (_ShouldSuppressDiscordLifecycleWebhook -Prefix $prefix -Tag 'SAVING')) {
                _Webhook (New-DiscordGameMessage -Profile $profile -Prefix $prefix -Event 'saving' -Values @{ WaitSeconds = $wait })
            }
        } else {
            _Log "[$($profile.GameName)] Save command failed or not configured - still waiting ${wait}s." -Level WARN
            _GameLog -Prefix $prefix -Line "[WARN] Save command failed or not configured. Waiting ${wait}s before shutdown"
            Set-ServerRuntimeState -Prefix $prefix -State 'stopping' -Detail ("Save unavailable. Waiting {0}s before shutdown." -f $wait)
            if (-not $Quiet -and -not (_ShouldSuppressDiscordLifecycleWebhook -Prefix $prefix -Tag 'WAITING')) {
                _Webhook (New-DiscordGameMessage -Profile $profile -Prefix $prefix -Event 'waiting_no_save' -Values @{ WaitSeconds = $wait })
            }
        }
    } else {
        _Log "[$($profile.GameName)] No save method configured - waiting ${wait}s for graceful shutdown..."
        _GameLog -Prefix $prefix -Line "[INFO] No save method configured. Waiting ${wait}s before shutdown"
        Set-ServerRuntimeState -Prefix $prefix -State 'stopping' -Detail ("Waiting {0}s before shutdown." -f $wait)
        if (-not $Quiet -and -not (_ShouldSuppressDiscordLifecycleWebhook -Prefix $prefix -Tag 'WAITING')) {
            _Webhook (New-DiscordGameMessage -Profile $profile -Prefix $prefix -Event 'waiting' -Values @{ WaitSeconds = $wait })
        }
    }

    # Step 3: Wait
    Start-Sleep -Seconds $wait

    # Step 3: Stop
    _ExecuteStop -Prefix $prefix -Profile $profile

    if (-not $Quiet) { _Webhook (New-DiscordGameMessage -Profile $profile -Prefix $prefix -Event 'stopped') }
    _Log "[$($profile.GameName)] Safe shutdown complete."
    _GameLog -Prefix $prefix -Line "[INFO] Safe shutdown complete"
    Set-ServerRuntimeState -Prefix $prefix -State 'stopped' -Detail 'Server stopped safely.'
    _ClearDiscordCommandContext -Prefix $prefix
    return $true
}

# -----------------------------------------------------------------------------
#  STOP (calls safe shutdown)
# -----------------------------------------------------------------------------
function Stop-GameServer {
    param([string]$Prefix)
    return Invoke-SafeShutdown -Prefix $Prefix
}

# -----------------------------------------------------------------------------
#  RESTART
# -----------------------------------------------------------------------------
function Restart-GameServer {
    param([string]$Prefix)

    $prefix  = $Prefix.ToUpper()
    $profile = _GetProfile $prefix

    _Log "[$($profile.GameName)] Restart initiated."
    Set-ServerRuntimeState -Prefix $prefix -State 'restarting' -Detail 'Restart sequence initiated.'
    if (-not (_ShouldSuppressDiscordLifecycleWebhook -Prefix $prefix -Tag 'RESTARTING')) {
        _Webhook (New-DiscordGameMessage -Profile $profile -Prefix $prefix -Event 'restarting')
    }

    Invoke-SafeShutdown -Prefix $prefix -Quiet
    Start-Sleep -Seconds 2

    $started = Start-GameServer -Prefix $prefix
    if ($started) {
        _SendDiscordLifecycleWebhook -Profile $profile -Prefix $prefix -Event 'restarted' -Tag 'ONLINE' | Out-Null
        _ClearDiscordCommandContext -Prefix $prefix
    } else {
        _Webhook (New-DiscordGameMessage -Profile $profile -Prefix $prefix -Event 'error_restart' -Values @{ Reason = $profile.GameName + ' restart failed.' })
        _ClearDiscordCommandContext -Prefix $prefix
    }
    return $started
}

# -----------------------------------------------------------------------------
#  SEND STDIN COMMAND
# -----------------------------------------------------------------------------
function Send-ServerStdin {
    param([string]$Prefix, [string]$Command, [switch]$ReturnDetail)

    $prefix  = $Prefix.ToUpper()
    $profile = _GetProfile $prefix

    $detail = @{
        Success     = $false
        Method      = 'None'
        Error       = ''
        ProcessName = ''
        TargetPid   = 0
    }

    # ── Strategy 1: direct stdin handle (works for java/.exe launched directly) ──
    # Palworld's server console is often a separate cmd wrapper and does NOT read redirected stdin.
    $preferWindow = $false
    $stdinPreferWindow = _GetProfileValue -Profile $profile -Name 'StdinPreferWindow' -Default $false
    if ($stdinPreferWindow) {
        $preferWindow = $true
    }
    if (_TestProfileGame -Profile $profile -KnownGame 'Palworld') { $preferWindow = $true }

    if (-not $preferWindow -and $script:State.StdinHandles.ContainsKey($prefix)) {
        try {
            $handle = $script:State.StdinHandles[$prefix]
            $handle.WriteLine($Command)
            $handle.Flush()
            _Log "[$($profile.GameName)] Sent stdin via handle: $Command"
            $detail.Success     = $true
            $detail.Method      = 'Handle'
            $detail.ProcessName = $profile.ProcessName
            return $(if ($ReturnDetail) { $detail } else { $true })
        }
        catch {
            _Log "[$($profile.GameName)] Stdin handle send failed: $_ - trying window method" -Level WARN
            $detail.Error = "$_"
        }
    }

    # ── Strategy 2: find the game process window by ProcessName and send keys ──
    # Used when the server was launched via cmd.exe /c bat (stdin goes to cmd, not the game)
    $procName = _CoalesceStr $profile.ProcessName ''
    $windowProcName = _CoalesceStr (_GetProfileValue -Profile $profile -Name 'StdinWindowProcessName' -Default '') ''
    if ($windowProcName) { $procName = $windowProcName }
    if (-not $procName) {
        _Log "[$($profile.GameName)] No ProcessName configured for window stdin fallback." -Level WARN
        $detail.Error       = "No ProcessName configured for window stdin fallback."
        $detail.ProcessName = ''
        return $(if ($ReturnDetail) { $detail } else { $false })
    }

    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue

        # Prefer a windowed process
        $target = $null
        $candidateNames = New-Object System.Collections.Generic.List[string]
        if ($windowProcName) { $candidateNames.Add($windowProcName) | Out-Null }
        if ($procName) { $candidateNames.Add($procName) | Out-Null }

        # Palworld: the console is commonly the cmd wrapper process
        if (_TestProfileGame -Profile $profile -KnownGame 'Palworld') {
            @('PalServer-Win64-Shipping-Cmd','palserver-win64-shipping-cmd','PalServer-Win64-Shipping','palserver-win64-shipping') | ForEach-Object {
                $candidateNames.Add($_) | Out-Null
            }
        }

        foreach ($name in $candidateNames) {
            $p = Get-Process -Name $name -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
            if ($p) { $target = $p; $procName = $name; break }
        }

        if (-not $target -and (_TestProfileGame -Profile $profile -KnownGame 'Palworld')) {
            $target = Get-Process -ErrorAction SilentlyContinue | Where-Object {
                $_.MainWindowHandle -ne 0 -and $_.MainWindowTitle -match 'PalServer|Palworld'
            } | Select-Object -First 1
            if ($target) { $procName = $target.ProcessName }
        }

        if (-not $target) {
            _Log "[$($profile.GameName)] Process '$procName' not found for window stdin." -Level WARN
            $detail.Error       = "Process '$procName' not found for window stdin."
            $detail.ProcessName = $procName
            return $(if ($ReturnDetail) { $detail } else { $false })
        }

        # SetForegroundWindow + SendKeys requires the window handle
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public class WinAPI {
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
'@ -ErrorAction SilentlyContinue

        $hwnd = $target.MainWindowHandle
        if ($hwnd -eq [IntPtr]::Zero) {
            _Log "[$($profile.GameName)] Process '$procName' has no main window handle." -Level WARN
            $detail.Error       = "Process '$procName' has no main window handle."
            $detail.ProcessName = $procName
            $detail.TargetPid   = $target.Id
            return $(if ($ReturnDetail) { $detail } else { $false })
        }

        [WinAPI]::ShowWindow($hwnd, 9)       | Out-Null  # SW_RESTORE
        [WinAPI]::SetForegroundWindow($hwnd) | Out-Null
        Start-Sleep -Milliseconds 200

        [System.Windows.Forms.SendKeys]::SendWait($Command)
        [System.Windows.Forms.SendKeys]::SendWait('{ENTER}')

        _Log "[$($profile.GameName)] Sent stdin via SendKeys to '$procName': $Command"
        $detail.Success     = $true
        $detail.Method      = 'Window'
        $detail.ProcessName = $procName
        $detail.TargetPid   = $target.Id
        return $(if ($ReturnDetail) { $detail } else { $true })
    }
    catch {
        _Log "[$($profile.GameName)] SendKeys stdin failed: $_" -Level ERROR
        $detail.Error       = "$_"
        $detail.ProcessName = $procName
        return $(if ($ReturnDetail) { $detail } else { $false })
    }
}

# -----------------------------------------------------------------------------
#  HTTP COMMAND
# -----------------------------------------------------------------------------
function Invoke-ServerHttp {
    param(
        [string]$Prefix,
        [string]$Url,
        [string]$Method  = 'GET',
        [string]$Body    = '',
        [hashtable]$Headers = @{}
    )
    try {
        $webReq = [System.Net.HttpWebRequest]::Create($Url)
        $webReq.Method = $Method
        foreach ($h in $Headers.Keys) { $webReq.Headers[$h] = $Headers[$h] }

        if ($Body -and $Method -ne 'GET') {
            $bytes = [System.Text.Encoding]::UTF8.GetBytes($Body)
            $webReq.ContentType   = 'application/json'
            $webReq.ContentLength = $bytes.Length
            $stream = $webReq.GetRequestStream()
            $stream.Write($bytes, 0, $bytes.Length)
            $stream.Close()
        }

        $resp   = $webReq.GetResponse()
        $reader = [System.IO.StreamReader]::new($resp.GetResponseStream())
        $result = $reader.ReadToEnd()
        $reader.Close()
        _Log "[$Prefix] HTTP $Method $Url -> OK"
        return $result
    }
    catch {
        _Log "[$Prefix] HTTP command failed: $_" -Level ERROR
        return $null
    }
}

# -----------------------------------------------------------------------------
#  CUSTOM SCRIPT
# -----------------------------------------------------------------------------
function Invoke-CustomScript {
    param([string]$Prefix, [string]$ScriptPath)

    $prefix  = $Prefix.ToUpper()
    $profile = _GetProfile $prefix
    $folder  = $profile.FolderPath

    if (-not [System.IO.Path]::IsPathRooted($ScriptPath)) {
        $ScriptPath = Join-Path $folder $ScriptPath
    }

    if (-not (Test-Path $ScriptPath)) {
        _Log "[$($profile.GameName)] Script not found: $ScriptPath" -Level ERROR
        return $false
    }

    try {
        if ($ScriptPath -match '\.bat$') {
            Start-Process cmd.exe -ArgumentList "/c `"$ScriptPath`"" -WorkingDirectory $folder -Wait
        } else {
            & $ScriptPath
        }
        _Log "[$($profile.GameName)] Script completed: $ScriptPath"
        return $true
    }
    catch {
        _Log "[$($profile.GameName)] Script failed: $_" -Level ERROR
        return $false
    }
}

# -----------------------------------------------------------------------------
#  INTERNAL: Execute save
# -----------------------------------------------------------------------------
function _ExecuteSave {
    param([string]$Prefix, [object]$Profile)
    $saveMethod = _CoalesceStr $Profile.SaveMethod 'none'
    switch ($saveMethod) {
        'stdin' {
            # Try StdinSaveCommand first, then StdinSave as fallback (older profiles)
            $saveCmd = ''
            if ($null -ne $Profile.StdinSaveCommand -and "$($Profile.StdinSaveCommand)".Trim() -ne '') {
                $saveCmd = $Profile.StdinSaveCommand.Trim()
            } elseif ($null -ne $Profile.StdinSave -and "$($Profile.StdinSave)".Trim() -ne '') {
                $saveCmd = $Profile.StdinSave.Trim()
            }
            if ($saveCmd) {
                _Log "[$($Profile.GameName)] Sending save command via stdin: '$saveCmd'"
                return Send-ServerStdin -Prefix $Prefix -Command $saveCmd
            } else {
                _Log "[$($Profile.GameName)] No StdinSaveCommand configured - skipping stdin save" -Level WARN
                return $false
            }
        }
        'http'  {
            $result = Invoke-ServerHttp -Prefix $Prefix -Url $Profile.SaveHttpUrl
            return $null -ne $result
        }
        'rest' {
            # Palworld REST save
            $Resthost = $Profile.RestHost
            $port = $Profile.RestPort
            $pass = $Profile.RestPassword

            _Log "[$($Profile.GameName)] Sending REST save request..."

            $result = Invoke-PalworldRestRequest `
                -RestHost $restHost `
                -Port $port `
                -Password $pass `
                -Endpoint '/v1/api/save' `
                -Method 'POST' `
                -Profile $Profile

            if ($result) {
                _Log "[$($Profile.GameName)] REST save OK: $result"
                return $true
            } else {
                _Log "[$($Profile.GameName)] REST save FAILED" -Level WARN
                return $false
            }
        }

        'SatisfactoryApi' {
            $sfHost  = if ($null -ne $Profile.SatisfactoryApiHost) { $Profile.SatisfactoryApiHost } else { '127.0.0.1' }
            $sfPort  = if ($null -ne $Profile.SatisfactoryApiPort) { [int]$Profile.SatisfactoryApiPort } else { 7777 }
            $sfToken = if ($null -ne $Profile.SatisfactoryApiToken) { $Profile.SatisfactoryApiToken } else { '' }

            # SaveGame requires a SaveName argument — we use 'ecc_autosave' so it's
            # clearly identifiable in the server's save file list.
            _Log "[$($Profile.GameName)] Sending SaveGame via Satisfactory API..."
            $r = Invoke-SatisfactoryApiRequest `
                -Host     $sfHost `
                -Port     $sfPort `
                -Token    $sfToken `
                -Function 'SaveGame' `
                -Data     @{ SaveName = 'ecc_autosave' }
            if ($r.Success) {
                _Log "[$($Profile.GameName)] SaveGame API call succeeded."
                return $true
            } else {
                _Log "[$($Profile.GameName)] SaveGame API call failed: $($r.Error)" -Level WARN
                return $false
            }
        }

        default { return $false }
    }
}
function _KillProcessTree {
    param([int]$RootPid)
    # Kill all children first, then the root
    try {
        $children = Get-WmiObject Win32_Process |
                    Where-Object { $_.ParentProcessId -eq $RootPid }
        foreach ($child in $children) {
            _KillProcessTree -RootPid $child.ProcessId
        }
        $proc = Get-Process -Id $RootPid -ErrorAction SilentlyContinue
        if ($proc -and -not $proc.HasExited) {
            $proc.Kill()
            _Log "Killed PID $RootPid ($($proc.ProcessName))"
        }
    } catch {
        _LogThrottled -Key "KillProcessTree::${RootPid}" -Msg "Failed to kill process tree rooted at PID ${RootPid}: $($_.Exception.Message)" -Level WARN -WindowSeconds 30
    }
}

function _StopCapturedOutput {
    param([string]$Prefix, [object]$Process = $null)

    if (-not $script:State -or -not $script:State.ContainsKey('LogWriters')) { return }
    if (-not $script:State.LogWriters.ContainsKey($Prefix)) { return }

    $capture = $script:State.LogWriters[$Prefix]
    if ($null -eq $capture) {
        try { $script:State.LogWriters.Remove($Prefix) | Out-Null } catch { }
        return
    }

    $writer = $null
    $proc   = $Process

    if ($capture -is [hashtable]) {
        if ($capture.ContainsKey('Writer')) { $writer = $capture.Writer }
        if ($null -eq $proc -and $capture.ContainsKey('Process')) { $proc = $capture.Process }
    } else {
        $writer = $capture
    }

    try {
        if ($proc) {
            try { $proc.CancelOutputRead() } catch { }
            try { $proc.CancelErrorRead() } catch { }
            try {
                if (-not $proc.HasExited) {
                    $proc.WaitForExit(2000) | Out-Null
                }
            } catch { }
        }
    } catch {
        _LogThrottled -Key "StopCapturedOutput::${Prefix}" -Msg "Failed while stopping captured output for ${Prefix}: $($_.Exception.Message)" -Level WARN -WindowSeconds 30
    }

    try { if ($writer) { $writer.Flush() } } catch { }
    try { if ($writer) { $writer.Close() } } catch { }
    try { if ($writer) { $writer.Dispose() } } catch { }
    try { $script:State.LogWriters.Remove($Prefix) | Out-Null } catch { }
}

function _ExecuteStop {
    param([string]$Prefix, [object]$Profile)
    $prefix     = $Prefix.ToUpper()
    $stopMethod = _CoalesceStr $Profile.StopMethod 'processKill'

    if (-not $script:State.RunningServers.ContainsKey($prefix)) {
        try {
            [void](Sync-RunningServersFromProcesses -SharedState $script:State)
        } catch {
            _LogThrottled -Key "SyncRunningServers::ExecuteStop::${prefix}" -Msg "[$($Profile.GameName)] Failed to sync running servers before stop execution: $($_.Exception.Message)" -Level WARN -WindowSeconds 30
        }
    }

    if (-not $script:State.RunningServers.ContainsKey($prefix)) {
        # Nothing to stop; server already gone
        _Log "[$($Profile.GameName)] Stop requested but no RunningServers entry for $prefix (already offline)." -Level WARN
        return
    }

    $entry = $script:State.RunningServers[$prefix]
    if (-not $entry) {
        # Entry removed between check and read
        return
    }

    _Log "[$($Profile.GameName)] Stop debug: method='$stopMethod' pid='$($entry.Pid)'" -Level DEBUG

    switch ($stopMethod) {
        'stdin' {
            # Send the graceful stop command, then wait, then force-kill if needed
            Send-ServerStdin -Prefix $prefix -Command $Profile.StdinStopCommand | Out-Null
            $proc = try { Get-Process -Id $entry.Pid -ErrorAction Stop } catch { $null }
            if ($proc) {
                $proc.WaitForExit(10000) | Out-Null
                if (-not $proc.HasExited) {
                    _KillProcessTree -RootPid $entry.Pid
                }
            }
        }
        'ctrlc' {
            # Graceful Valheim shutdown using CTRL+C
            $proc = try { Get-Process -Id $entry.Pid -ErrorAction Stop } catch { $null }
            if ($proc) {
                _Log "[$($Profile.GameName)] Sending CTRL+C for graceful shutdown..."

                # Import GenerateConsoleCtrlEvent
                $sig = '[DllImport("kernel32.dll")] public static extern bool GenerateConsoleCtrlEvent(uint dwCtrlEvent, uint dwProcessGroupId);'
                $kernel = Add-Type -MemberDefinition $sig -Name "Kernel32" -Namespace Win32 -PassThru

                # CTRL_C_EVENT = 0
                $kernel::GenerateConsoleCtrlEvent(0, $proc.Id) | Out-Null

                # Give Valheim time to save and exit
                Start-Sleep -Seconds 10

                # If still running, force kill
                if (-not $proc.HasExited) {
                    _Log "[$($Profile.GameName)] Graceful shutdown incomplete, force killing..."
                    _KillProcessTree -RootPid $entry.Pid
                }
            }
        }
        'processName' {
            # Kill by the actual game process name defined in the profile
            # This handles cases where we launched via cmd/bat wrapper
            $procName = _CoalesceStr $Profile.ProcessName ''
            if ($procName) {
                $targets = Get-Process -Name $procName -ErrorAction SilentlyContinue
                foreach ($t in $targets) {
                    try {
                        $t.Kill()
                        _Log "[$($Profile.GameName)] Killed process '$procName' (PID $($t.Id))"
                    } catch {
                        _Log "[$($Profile.GameName)] Could not kill '$procName' (PID $($t.Id)): $_" -Level WARN
                    }
                }
            }
            # Also kill the launcher wrapper PID just in case
            _KillProcessTree -RootPid $entry.Pid
        }
        'http' {
            $httpMethod = _CoalesceStr $Profile.StopHttpMethod 'POST'
            Invoke-ServerHttp -Prefix $prefix -Url $Profile.StopHttpUrl -Method $httpMethod | Out-Null
            Start-Sleep -Seconds 5
        }
        'SatisfactoryApi' {
            # Call the Shutdown API — this triggers a final engine save inside the
            # game before the process exits, then falls back to processKill if needed.
            $sfHost  = if ($null -ne $Profile.SatisfactoryApiHost) { $Profile.SatisfactoryApiHost } else { '127.0.0.1' }
            $sfPort  = if ($null -ne $Profile.SatisfactoryApiPort) { [int]$Profile.SatisfactoryApiPort } else { 7777 }
            $sfToken = if ($null -ne $Profile.SatisfactoryApiToken) { $Profile.SatisfactoryApiToken } else { '' }

            _Log "[$($Profile.GameName)] Sending Shutdown via Satisfactory API..."
            $r = Invoke-SatisfactoryApiRequest -Host $sfHost -Port $sfPort -Token $sfToken -Function 'Shutdown'
            if ($r.Success) {
                _Log "[$($Profile.GameName)] API Shutdown sent - waiting for process to exit..."
                $proc = try { Get-Process -Id $entry.Pid -ErrorAction Stop } catch { $null }
                if ($proc) { $proc.WaitForExit(15000) | Out-Null }
            } else {
                _Log "[$($Profile.GameName)] API Shutdown failed ($($r.Error)) - falling back to processKill." -Level WARN
            }
            # Always clean up with kill in case API shutdown did not fully terminate
            _KillProcessTree -RootPid $entry.Pid
        }
        default {
            # processKill - kill the wrapper PID and its entire child tree
            _KillProcessTree -RootPid $entry.Pid
        }
    }

    # Cleanup (safe)
    if ($script:State.RunningServers.ContainsKey($prefix)) {
        $script:State.RunningServers.Remove($prefix)
    }

    _ClearPendingAutoRestart -Prefix $prefix
    _ClearPendingScheduledRestart -Prefix $prefix
    if ($script:State.StdinHandles.ContainsKey($prefix)) {
        $script:State.StdinHandles.Remove($prefix)
    }
    _ClearServerSessionCaches -Prefix $prefix -SharedState $script:State

    # Close any stdout/stderr capture writer
    try { _StopCapturedOutput -Prefix $prefix -Process $entry.Process } catch {
        _LogThrottled -Key "StopCapturedOutputCall::${prefix}" -Msg "[$($Profile.GameName)] Failed to close captured output during stop: $($_.Exception.Message)" -Level WARN -WindowSeconds 30
    }
}

# -----------------------------------------------------------------------------
#  PALWORLD REST STATE POLLER
# -----------------------------------------------------------------------------
function Update-PalworldRestState {
    param(
        [object]$Profile,
        [hashtable]$SharedState,
        [string]$Prefix = ''
    )

    # Avoid PowerShell automatic variable $resthost
    $restHost = $Profile.RestHost
    $restPort = $Profile.RestPort
    $restPass = $Profile.RestPassword
    # Do not log REST passwords; keep secrets out of logs.

    # Track consecutive REST failures (disable after 2 failures)
    $pfx = if ($Prefix) { $Prefix.ToUpper() } else { ($Profile.Prefix).ToUpper() }
    if (-not $SharedState.RestFailCounts.ContainsKey($pfx)) { $SharedState.RestFailCounts[$pfx] = 0 }
    if (-not $SharedState.ContainsKey('PalworldStateByPrefix')) {
        $SharedState['PalworldStateByPrefix'] = [hashtable]::Synchronized(@{})
    }
    if (-not $SharedState.ContainsKey('PalworldPlayersByPrefix')) {
        $SharedState['PalworldPlayersByPrefix'] = [hashtable]::Synchronized(@{})
    }

    # Query server state
    $stateOk = $false
    $statusEndpoint = ''
    if ($Profile -and $Profile.RestStatusEndpoint) {
        $statusEndpoint = [string]$Profile.RestStatusEndpoint
    } elseif ($Profile -and -not (_TestProfileGame -Profile $Profile -KnownGame 'Palworld')) {
        $statusEndpoint = '/v1/api/server/info'
    }

    if (-not [string]::IsNullOrWhiteSpace($statusEndpoint)) {
        $stateResp = Invoke-PalworldRestRequest `
            -RestHost $restHost `
            -Port $restPort `
            -Password $restPass `
            -Endpoint $statusEndpoint `
            -Method 'GET' `
            -Profile $Profile

        if ($null -ne $stateResp) {
            try {
                if ($stateResp -is [string]) {
                    $trim = $stateResp.Trim()
                    if ($trim.StartsWith('{') -or $trim.StartsWith('[')) {
                        $state = $trim | ConvertFrom-Json -ErrorAction Stop
                        $SharedState.PalworldStateByPrefix[$pfx] = $state
                        $stateOk = $true
                    } else {
                        _Log "Palworld REST info response not JSON: $trim" -Level WARN
                    }
                } else {
                    $SharedState.PalworldStateByPrefix[$pfx] = $stateResp
                    $stateOk = $true
                }
            } catch {
                _Log "Palworld REST info parse failed: $_" -Level WARN
            }
        }
    }

    # Query players
    $playersResp = Invoke-PalworldRestRequest `
        -RestHost $restHost `
        -Port $restPort `
        -Password $restPass `
        -Endpoint '/v1/api/players' `
        -Method 'GET' `
        -Profile $Profile

    $playersOk = $false
    if ($null -ne $playersResp) {
        try {
            if ($playersResp -is [string]) {
                $trim = $playersResp.Trim()
                if ($trim.StartsWith('{') -or $trim.StartsWith('[')) {
                    $players = $trim | ConvertFrom-Json -ErrorAction Stop
                    $SharedState.PalworldPlayersByPrefix[$pfx] = $players
                    $playersOk = $true
                } else {
                    _Log "Palworld REST players response not JSON: $trim" -Level WARN
                }
            } else {
                $SharedState.PalworldPlayersByPrefix[$pfx] = $playersResp
                $playersOk = $true
            }
        } catch {
            _Log "Palworld REST players parse failed: $_" -Level WARN
        }
    }

    if ($stateOk -or $playersOk) {
        $SharedState.RestFailCounts[$pfx] = 0

        # First successful REST response after server start = server is joinable.
        # Use a SharedState set to fire the webhook exactly once per session.
        if (-not $SharedState.ContainsKey('RestJoinableNotified')) {
            $SharedState['RestJoinableNotified'] = [hashtable]::Synchronized(@{})
        }
        if (-not $SharedState.RestJoinableNotified.ContainsKey($pfx)) {
            $SharedState.RestJoinableNotified[$pfx] = $true
            $gameName = if ($Profile.GameName) { $Profile.GameName } else { $pfx }
            Set-JoinableServerRuntimeState -Prefix $pfx -Detail 'Server is joinable and responding to REST.' -SharedState $SharedState
            _Log "[$gameName] REST API responded - server is joinable"
            _GameLog -Prefix $pfx -Line "[REST] API responded - server is joinable"
            _SendDiscordLifecycleWebhook -Profile $Profile -Prefix $pfx -Event 'joinable' -Tag 'JOINABLE' -SharedState $SharedState | Out-Null
        }

        try {
            if (-not $SharedState.ContainsKey('PalworldFeedState')) {
                $SharedState['PalworldFeedState'] = [hashtable]::Synchronized(@{})
            }

            $playerCount = 0
            $playersObj = $null
            if ($SharedState.PalworldPlayersByPrefix.ContainsKey($pfx)) {
                $playersObj = $SharedState.PalworldPlayersByPrefix[$pfx]
            }
            if ($playersObj -and $playersObj.PSObject.Properties['players']) {
                try { $playerCount = @($playersObj.players).Count } catch { $playerCount = 0 }
            } elseif ($playersObj -is [System.Collections.IEnumerable] -and -not ($playersObj -is [string])) {
                try { $playerCount = @($playersObj).Count } catch { $playerCount = 0 }
            }

            $last = $null
            if ($SharedState.PalworldFeedState.ContainsKey($pfx)) {
                $last = $SharedState.PalworldFeedState[$pfx]
            }
            if ($null -eq $last -or $last.PlayerCount -ne $playerCount) {
                _GameLog -Prefix $pfx -Line "[PLAYERS] Online: $playerCount"
                $SharedState.PalworldFeedState[$pfx] = @{ PlayerCount = $playerCount }
                _SetLatestPlayersSnapshot -Prefix $pfx -Names @() -Count $playerCount -SharedState $SharedState
                Set-ObservedPlayersServerRuntimeState -Prefix $pfx -Count $playerCount -SharedState $SharedState
            }
        } catch { }
    } else {
        # REST failed — clear the notified flag so if the server restarts
        # and the API comes back up, we fire the notification again.
        if ($SharedState.ContainsKey('RestJoinableNotified') -and
            $SharedState.RestJoinableNotified.ContainsKey($pfx)) {
            $SharedState.RestJoinableNotified.Remove($pfx)
        }
        if ($SharedState.ContainsKey('PalworldFeedState') -and $SharedState.PalworldFeedState.ContainsKey($pfx)) {
            $SharedState.PalworldFeedState.Remove($pfx)
        }
        if ($SharedState.ContainsKey('PalworldStateByPrefix') -and $SharedState.PalworldStateByPrefix.ContainsKey($pfx)) {
            $SharedState.PalworldStateByPrefix.Remove($pfx)
        }
        if ($SharedState.ContainsKey('PalworldPlayersByPrefix') -and $SharedState.PalworldPlayersByPrefix.ContainsKey($pfx)) {
            $SharedState.PalworldPlayersByPrefix.Remove($pfx)
        }
        $SharedState.RestFailCounts[$pfx] = [int]$SharedState.RestFailCounts[$pfx] + 1
        if ([int]$SharedState.RestFailCounts[$pfx] -ge 2) {
            $Profile.RestEnabled = $false
            _Log "[$($Profile.GameName)] REST failed twice; disabling REST polling." -Level WARN
        }
    }
}


# -----------------------------------------------------------------------------
#  AUTO-RESTART MONITOR
# -----------------------------------------------------------------------------
# -----------------------------------------------------------------------------
#  AUTO-SAVE CHECKER
#  Called every monitor tick for each running server.
#  Fires _ExecuteSave when the configured interval has elapsed.
#  Safe to call even when SaveMethod = 'none' - returns immediately.
# -----------------------------------------------------------------------------
function _CheckAutoSave {
    param([string]$Prefix, [object]$Profile, [hashtable]$SharedState)

    # Global enable gate
    $globalEnabled = $true
    if ($SharedState.Settings -and $SharedState.Settings.ContainsKey('AutoSaveEnabled')) {
        $globalEnabled = [bool]$SharedState.Settings.AutoSaveEnabled
    }
    if (-not $globalEnabled) { return }

    # Per-profile enable gate (opt-out)
    if ($null -ne $Profile.AutoSaveEnabled -and -not [bool]$Profile.AutoSaveEnabled) { return }

    # No save method - nothing to do
    $saveMethod = _CoalesceStr $Profile.SaveMethod 'none'
    if ($saveMethod -eq 'none') { return }

    # Determine interval (profile overrides global)
    $intervalMin = 30
    if ($SharedState.Settings -and $SharedState.Settings.ContainsKey('AutoSaveIntervalMinutes')) {
        $v = 0
        if ([int]::TryParse("$($SharedState.Settings.AutoSaveIntervalMinutes)", [ref]$v) -and $v -gt 0) {
            $intervalMin = $v
        }
    }
    if ($null -ne $Profile.AutoSaveIntervalMinutes) {
        $v = 0
        if ([int]::TryParse("$($Profile.AutoSaveIntervalMinutes)", [ref]$v) -and $v -gt 0) {
            $intervalMin = $v
        }
    }

    # Ensure LastAutoSave dictionary exists
    if (-not $SharedState.ContainsKey('LastAutoSave')) {
        $SharedState['LastAutoSave'] = [hashtable]::Synchronized(@{})
    }

    $now = Get-Date

    # Seed on first encounter so the clock starts from server-start, not epoch
    if (-not $SharedState.LastAutoSave.ContainsKey($Prefix)) {
        $SharedState.LastAutoSave[$Prefix] = $now
        _TraceServerTimerDecision -Prefix $Prefix -Timer 'autosave' -OldRule $null -OldDueAt $null -NewRule 'interval' -NewDueAt $now.AddMinutes($intervalMin) -Reason ("auto-save timer armed with {0}-minute interval" -f $intervalMin) -SharedState $SharedState
        return
    }

    $previousSaveAt = [datetime]$SharedState.LastAutoSave[$Prefix]
    $elapsed = ($now - $previousSaveAt).TotalMinutes
    if ($elapsed -lt $intervalMin) { return }

    # Interval elapsed - fire the save
    _Log "[$($Profile.GameName)] Auto-save triggered (interval ${intervalMin}min, elapsed $([Math]::Round($elapsed,1))min)"
    _Webhook (New-DiscordGameMessage -Profile $Profile -Prefix $Prefix -Event 'autosave_started')

    $ok = _ExecuteSave -Prefix $Prefix -Profile $Profile
    $SharedState.LastAutoSave[$Prefix] = $now
    _TraceServerTimerDecision -Prefix $Prefix -Timer 'autosave' -OldRule 'interval' -OldDueAt $previousSaveAt.AddMinutes($intervalMin) -NewRule 'interval' -NewDueAt $now.AddMinutes($intervalMin) -Reason ("auto-save interval restarted after save attempt ({0}-minute interval)" -f $intervalMin) -SharedState $SharedState

    if ($ok) {
        _Log "[$($Profile.GameName)] Auto-save completed."
        _Webhook (New-DiscordGameMessage -Profile $Profile -Prefix $Prefix -Event 'autosave_done')
    } else {
        _Log "[$($Profile.GameName)] Auto-save command failed or returned no result." -Level WARN
    }
}

# -----------------------------------------------------------------------------
#  SCHEDULED RESTART CHECKER
#  Called every monitor tick for each running server.
#  Sends warning webhooks at 60/30/15/10/5/2/1 minutes before restart,
#  then performs a full safe shutdown + restart at the interval boundary.
#  Uses SharedState.SentWarnings (key = "PREFIX_minutes") to fire each
#  warning exactly once per session, and clears them after the restart.
# -----------------------------------------------------------------------------
function _CheckScheduledRestart {
    param([string]$Prefix, [object]$Profile, [hashtable]$SharedState)

    # Global enable gate
    $globalEnabled = $true
    if ($SharedState.Settings -and $SharedState.Settings.ContainsKey('ScheduledRestartEnabled')) {
        $globalEnabled = [bool]$SharedState.Settings.ScheduledRestartEnabled
    }
    if (-not $globalEnabled) { return }

    # Per-profile opt-out
    if ($null -ne $Profile.ScheduledRestartEnabled -and -not [bool]$Profile.ScheduledRestartEnabled) { return }

    # Must be running and have a start time
    if (-not $SharedState.RunningServers.ContainsKey($Prefix)) { return }
    $entry = $SharedState.RunningServers[$Prefix]
    if ($null -eq $entry -or $null -eq $entry.StartTime) { return }

    # Determine interval in minutes
    $intervalHours = 6.0
    if ($SharedState.Settings -and $SharedState.Settings.ContainsKey('ScheduledRestartHours')) {
        $v = 0.0
        if ([double]::TryParse("$($SharedState.Settings.ScheduledRestartHours)", [ref]$v) -and $v -gt 0) {
            $intervalHours = $v
        }
    }
    $intervalMin  = $intervalHours * 60.0
    $uptimeMin    = ((Get-Date) - [datetime]$entry.StartTime).TotalMinutes
    $remainingMin = $intervalMin - $uptimeMin

    # Ensure SentWarnings dictionary exists
    if (-not $SharedState.ContainsKey('SentWarnings')) {
        $SharedState['SentWarnings'] = [hashtable]::Synchronized(@{})
    }

    # Warning thresholds in minutes
    $warnThresholds = @(60, 30, 15, 10, 5, 2, 1)

    $playerSummary = ''
    try {
        $activity = _UpdatePlayerActivityState -Prefix $Prefix -Profile $Profile -SharedState $SharedState
        if ($activity -and $activity.DetectionAvailable) {
            $playerCount = 0
            try { $playerCount = [int]$activity.CurrentCount } catch { $playerCount = 0 }
            $playerSummary = " Players online: $playerCount."
        }
    } catch { }

    foreach ($thresh in $warnThresholds) {
        $warnKey = "${Prefix}_${thresh}"
        if ($remainingMin -le $thresh -and $remainingMin -gt ($thresh - 1.5) -and
            -not $SharedState.SentWarnings.ContainsKey($warnKey)) {

            $SharedState.SentWarnings[$warnKey] = $true
            $gameName = _GetNotificationGameName -Profile $Profile -Prefix $Prefix
            $warningSummary = _JoinNotificationParts -Parts @(
                ("Scheduled restart in {0} minute{1}" -f $thresh, $(if ($thresh -ne 1) { 's' } else { '' })),
                $(if ($null -ne $playerCount -and $playerCount -ge 0) { "Players online: $playerCount" } else { '' })
            )
            _Log "[$gameName] $warningSummary" -Level INFO
            _Webhook (New-DiscordGameMessage -Profile $Profile -Prefix $Prefix -Event 'restart_warning' -Values @{
                Minutes = $thresh
                PlayerSummary = $(if ($null -ne $playerCount -and $playerCount -ge 0) { "Players online: $playerCount." } else { '' })
            })
        }
    }

    # Time to restart
    if ($remainingMin -gt 0) { return }

    $gameName = $Profile.GameName
    if ($SharedState.ContainsKey('PendingScheduledRestarts') -and $SharedState.PendingScheduledRestarts.ContainsKey($Prefix)) { return }

    $playerCountNow = $null
    try {
        $activity = _UpdatePlayerActivityState -Prefix $Prefix -Profile $Profile -SharedState $SharedState
        if ($activity -and $activity.DetectionAvailable) {
            $playerCountNow = [int]$activity.CurrentCount
        }
    } catch { $playerCountNow = $null }

    $restartDetail = if ($null -ne $playerCountNow -and $playerCountNow -gt 0) {
        "Scheduled restart in progress. ${playerCountNow} player(s) online; restart will proceed."
    } else {
        'Scheduled restart in progress.'
    }

    $restartStartSummary = _JoinNotificationParts -Parts @(
        'Scheduled restart beginning now',
        $(if ($null -ne $playerCountNow -and $playerCountNow -gt 0) { "Players online: $playerCountNow" } elseif ($null -ne $playerCountNow) { 'Players online: 0' } else { '' }),
        'Safe shutdown and restart in progress.'
    )
    _Log "[$gameName] $restartStartSummary" -Level INFO
    _GameLog -Prefix $Prefix -Line ("[INFO] {0}" -f $restartStartSummary)
    Set-ServerRuntimeState -Prefix $Prefix -State 'restarting' -Detail $restartDetail -SharedState $SharedState
    _Webhook (New-DiscordGameMessage -Profile $Profile -Prefix $Prefix -Event 'scheduled_restart_started' -Values @{
        PlayerSummary = $(if ($null -ne $playerCountNow -and $playerCountNow -gt 0) { "Players online: $playerCountNow." } elseif ($null -ne $playerCountNow) { 'Players online: 0.' } else { '' })
    })

    # Clear warnings so next session starts fresh
    foreach ($thresh in $warnThresholds) {
        $warnKey = "${Prefix}_${thresh}"
        if ($SharedState.SentWarnings.ContainsKey($warnKey)) {
            $SharedState.SentWarnings.Remove($warnKey)
        }
    }

    # Clear LastAutoSave so the clock resets for the new session
    if ($SharedState.ContainsKey('LastAutoSave') -and $SharedState.LastAutoSave.ContainsKey($Prefix)) {
        $SharedState.LastAutoSave.Remove($Prefix)
    }

    # Safe shutdown then restart
    # Set shutdown flag BEFORE the sleep so the crash-monitor does not
    # misread the process exit as a crash during the save-wait window.
    $SharedState.ShutdownFlags[$Prefix] = $true
    $startOk = $false
    try {
        Invoke-SafeShutdown -Prefix $Prefix -Quiet | Out-Null
        Start-Sleep -Seconds 3
        $startResult = Start-GameServer -Prefix $Prefix
        if ($startResult -eq $true -or $startResult -eq 'already_running') {
            $startOk = $true
        }
    } finally {
        $SharedState.ShutdownFlags[$Prefix] = $false
    }

    if ($startOk) {
        _GameLog -Prefix $Prefix -Line "[INFO] Scheduled restart completed successfully. Server is back online."
        _SendDiscordLifecycleWebhook -Profile $Profile -Prefix $Prefix -Event 'scheduled_restart_done' -Tag 'ONLINE' -SharedState $SharedState | Out-Null
    } else {
        $maxAttempts = 3
        if ($null -ne $Profile.MaxRestartsPerHour) {
            try {
                $candidate = [int]$Profile.MaxRestartsPerHour
                if ($candidate -gt 0) { $maxAttempts = [Math]::Max(1, [Math]::Min($candidate, 5)) }
            } catch { }
        }

        _Log "[$gameName] Scheduled restart failed to start the server. Queuing recovery retry." -Level WARN
        _GameLog -Prefix $Prefix -Line "[WARN] Scheduled restart did not come back up cleanly. Recovery retry queued in 15s."
        _SchedulePendingScheduledRestart -Prefix $Prefix -Profile $Profile -SharedState $SharedState -DelaySeconds 15 -Attempt 1 -MaxAttempts $maxAttempts
        _Webhook (New-DiscordGameMessage -Profile $Profile -Prefix $Prefix -Event 'scheduled_restart_retry' -Values @{
            DelaySeconds = 15
            Attempt = 1
            MaxAttempts = $maxAttempts
        })
    }
}

# -----------------------------------------------------------------------------
#  AUTO-RESTART MONITOR
# -----------------------------------------------------------------------------
function Start-ServerMonitor {
    param([hashtable]$SharedState)
    _Log "Server monitor started."

    while (-not ($SharedState.ContainsKey('StopMonitor') -and $SharedState['StopMonitor'] -eq $true)) {
        Start-Sleep -Seconds 15
        $tickStartedAt = Get-Date
        $profileCount = 0
        $runningCount = 0
        $totalProfileElapsedMs = 0
        $slowProfiles = New-Object 'System.Collections.Generic.List[string]'
        try { $profileCount = @($SharedState.Profiles.Keys).Count } catch { $profileCount = 0 }
        try { $runningCount = if ($SharedState.RunningServers) { $SharedState.RunningServers.Count } else { 0 } } catch { $runningCount = 0 }
        _Log "Monitor tick: profiles=$profileCount running=$runningCount" -Level DEBUG

        foreach ($prefix in @($SharedState.Profiles.Keys)) {
            $profileStartedAt = Get-Date
            $profile = $null
            try {
                $profile = $SharedState.Profiles[$prefix]
                # Palworld REST polling (only when server is running)
                if ($null -ne $profile.RestEnabled -and [bool]$profile.RestEnabled) {
                    $onlyWhenRunning = $true
                    if ($null -ne $profile.RestPollOnlyWhenRunning) {
                        $onlyWhenRunning = [bool]$profile.RestPollOnlyWhenRunning
                    }

                    _Log "Monitor REST check: prefix=$prefix enabled=$($profile.RestEnabled) onlyWhenRunning=$onlyWhenRunning running=$($SharedState.RunningServers.ContainsKey($prefix))" -Level DEBUG

                    if (-not $onlyWhenRunning -or $SharedState.RunningServers.ContainsKey($prefix)) {
                        Update-PalworldRestState -Profile $profile -SharedState $SharedState -Prefix $prefix
                    }
                }

                try {
                    if ($SharedState.RunningServers.ContainsKey($prefix) -and (_TestProfileGame -Profile $profile -KnownGame '7DaysToDie')) {
                        [void](_Refresh7DaysToDiePlayers -Prefix $prefix -Profile $profile -SharedState $SharedState)
                    }
                } catch {
                    _LogProfileError -Prefix $prefix -Profile $profile -Action 'refresh-players' -Function '_Refresh7DaysToDiePlayers' -Message '7 Days to Die player refresh failed.' -ErrorRecord $_ -Fallback 'ECC kept the last trusted player snapshot and will retry the refresh on the next monitor tick.' -Recovery 'degraded but still running; player refresh retry pending next tick'
                }
                try {
                    if ($SharedState.RunningServers.ContainsKey($prefix) -and (_TestProfileGame -Profile $profile -KnownGame 'Hytale')) {
                        [void](_RefreshHytalePlayers -Prefix $prefix -Profile $profile -SharedState $SharedState)
                    }
                } catch {
                    _LogProfileError -Prefix $prefix -Profile $profile -Action 'refresh-players' -Function '_RefreshHytalePlayers' -Message 'Hytale player refresh failed.' -ErrorRecord $_ -Fallback 'ECC kept the last trusted player snapshot and will retry the refresh on the next monitor tick.' -Recovery 'degraded but still running; player refresh retry pending next tick'
                }

                # ── Auto-save and scheduled restart (only when server is running) ──
                if ($SharedState.RunningServers.ContainsKey($prefix)) {
                    $shutdownNow = $false
                    if ($SharedState.ShutdownFlags -and $SharedState.ShutdownFlags.ContainsKey($prefix)) {
                        $shutdownNow = [bool]$SharedState.ShutdownFlags[$prefix]
                    }
                    if (-not $shutdownNow) {
                        $startupHandled = $false
                        try { $startupHandled = (_CheckStartupHealth -Prefix $prefix -Profile $profile -SharedState $SharedState) } catch { _LogProfileError -Prefix $prefix -Profile $profile -Action 'startup-health-check' -Function '_CheckStartupHealth' -Message 'Startup health check failed.' -ErrorRecord $_ -Fallback 'Startup health will be checked again on the next monitor tick while the server keeps running.' -Recovery 'deferred to next monitor tick while startup continues' }
                        if ($startupHandled) { continue }
                        $idleHandled = $false
                        try { $idleHandled = (_CheckIdleShutdown -Prefix $prefix -Profile $profile -SharedState $SharedState) } catch { _LogProfileError -Prefix $prefix -Profile $profile -Action 'idle-shutdown-check' -Function '_CheckIdleShutdown' -Message 'Idle shutdown check failed.' -ErrorRecord $_ -Fallback 'Idle shutdown state was left unchanged and will be re-evaluated on the next monitor tick.' -Recovery 'deferred to next monitor tick with existing idle state preserved' }
                        if ($idleHandled) { continue }
                        try { _CheckAutoSave           -Prefix $prefix -Profile $profile -SharedState $SharedState } catch { _LogProfileError -Prefix $prefix -Profile $profile -Action 'autosave-check' -Function '_CheckAutoSave' -Message 'Auto-save check failed.' -ErrorRecord $_ -Fallback 'Auto-save was skipped for this tick and will be re-evaluated on the next monitor pass.' -Recovery 'deferred to next monitor tick; no save executed this pass' }
                        try { _CheckScheduledRestart   -Prefix $prefix -Profile $profile -SharedState $SharedState } catch { _LogProfileError -Prefix $prefix -Profile $profile -Action 'scheduled-restart-check' -Function '_CheckScheduledRestart' -Message 'Scheduled restart check failed.' -ErrorRecord $_ -Fallback 'Scheduled restart timing was left unchanged and will be re-evaluated on the next monitor tick.' -Recovery 'deferred to next monitor tick with restart schedule preserved' }
                    }
                }

                # Safe property access for OrderedDictionary
                $autoRestart = $true
                if ($null -ne $profile.EnableAutoRestart) {
                    $autoRestart = [bool]$profile.EnableAutoRestart
                }
                if (-not $autoRestart) {
                    _ClearPendingAutoRestart -Prefix $prefix -SharedState $SharedState
                }

                $scheduledRestartHandled = $false
                if ($SharedState.PendingScheduledRestarts -and $SharedState.PendingScheduledRestarts.ContainsKey($prefix)) {
                    try {
                        $scheduledRestartHandled = [bool](_ProcessPendingScheduledRestart -Prefix $prefix -Profile $profile -SharedState $SharedState)
                    } catch {
                        _LogProfileError -Prefix $prefix -Profile $profile -Action 'pending-scheduled-restart' -Function '_ProcessPendingScheduledRestart' -Message 'Pending scheduled restart processing failed.' -ErrorRecord $_ -Fallback 'The queued scheduled restart entry remains in shared state for the next monitor pass.' -Recovery 'queued recovery remains pending for a later monitor pass'
                    }
                }
                if ($scheduledRestartHandled) {
                    continue
                }

                if (-not $autoRestart) {
                    continue
                }

                if ($SharedState.PendingAutoRestarts -and $SharedState.PendingAutoRestarts.ContainsKey($prefix)) {
                    try {
                        [void](_ProcessPendingAutoRestart -Prefix $prefix -Profile $profile -SharedState $SharedState)
                    } catch {
                        _LogProfileError -Prefix $prefix -Profile $profile -Action 'pending-auto-restart' -Function '_ProcessPendingAutoRestart' -Message 'Pending auto-restart processing failed.' -ErrorRecord $_ -Fallback 'The queued auto-restart entry remains in shared state for the next monitor pass.' -Recovery 'queued recovery remains pending for a later monitor pass'
                    }
                }

                # Safe check for shutdown flag using property access instead of ContainsKey
                $shutdownFlag = $false
                if ($null -ne $SharedState.ShutdownFlags) {
                    $shutdownFlag = $SharedState.ShutdownFlags[$prefix]
                    if ($shutdownFlag -eq $true) { continue }
                }

                $hasEntry = $SharedState.RunningServers.ContainsKey($prefix)
                _Log "Monitor check: prefix=$prefix autoRestart=$autoRestart shutdownFlag=$shutdownFlag hasEntry=$hasEntry" -Level DEBUG

                if (-not $hasEntry) {
                    continue
                }

                $entry = $SharedState.RunningServers[$prefix]
                if (-not $entry) { continue }

                $proc = try { Get-Process -Id $entry.Pid -ErrorAction Stop } catch { $null }
                if ($null -ne $proc -and -not $proc.HasExited) { continue }

                # Fallback for launchers (e.g. cmd.exe) that exit after spawning the real server
                if ($profile.ProcessName -and $profile.ProcessName.Trim() -ne '') {
                    $gameProc = Get-Process -Name $profile.ProcessName -ErrorAction SilentlyContinue | Select-Object -First 1
                    if ($gameProc) { continue }
                }

                # Unexpected exit detected
                $ts  = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                $msg = "[$ts][WARN][Monitor] $($profile.GameName) exited unexpectedly."
                $SharedState.LogQueue.Enqueue($msg)
                try { Write-Host $msg -ForegroundColor Yellow } catch {}

                # Rate limit check
                $hour  = (Get-Date).ToString('yyyyMMddHH')
                $key   = "${prefix}_${hour}"
                
                if ($null -eq $SharedState.RestartCounters[$key]) {
                    $SharedState.RestartCounters[$key] = 0
                }
                
                $count = [int]$SharedState.RestartCounters[$key]
                $max   = 5
                
                if ($null -ne $profile.MaxRestartsPerHour) {
                    $max = [int]$profile.MaxRestartsPerHour
                }

                if ($count -ge $max) {
                    $warn = "[$ts][WARN][Monitor] $($profile.GameName) crashed $count times this hour (max $max). Auto-restart suspended."
                    $SharedState.LogQueue.Enqueue($warn)
                    $SharedState.WebhookQueue.Enqueue((New-DiscordGameMessage -Profile $profile -Prefix $prefix -Event 'crash_suspended' -Values @{
                        Count = $count
                    }))
                    $SharedState.RunningServers.Remove($prefix)
                    _ClearServerSessionCaches -Prefix $prefix -SharedState $SharedState
                    Set-ServerRuntimeState -Prefix $prefix -State 'failed' -Detail ("Auto-restart suspended after {0} crashes this hour." -f $count) -SharedState $SharedState
                    _ClearPendingAutoRestart -Prefix $prefix -SharedState $SharedState
                    continue
                }

                $SharedState.RestartCounters[$key] = $count + 1
                $delay = 10
                
                if ($null -ne $profile.RestartDelaySeconds) {
                    $delay = [int]$profile.RestartDelaySeconds
                }

                $SharedState.WebhookQueue.Enqueue((New-DiscordGameMessage -Profile $profile -Prefix $prefix -Event 'crashed_retry' -Values @{
                    DelaySeconds = $delay
                    Attempt = ($count + 1)
                    MaxAttempts = $max
                }))
                $SharedState.RunningServers.Remove($prefix)
                _ClearServerSessionCaches -Prefix $prefix -SharedState $SharedState
                Set-ServerRuntimeState -Prefix $prefix -State 'waiting_restart' -Detail ("Crash detected. Restarting in {0}s (attempt {1}/{2})." -f $delay, ($count + 1), $max) -SharedState $SharedState
                _SchedulePendingAutoRestart -Prefix $prefix -Profile $profile -SharedState $SharedState -DelaySeconds $delay -Attempt ($count + 1) -MaxAttempts $max
            }
            catch {
                _Log "Error in monitor loop for $prefix : $_" -Level ERROR
            }
            finally {
                $elapsedMs = [int][Math]::Round(((Get-Date) - $profileStartedAt).TotalMilliseconds)
                $profileSlowThresholdMs = _GetMonitorSlowProfileThresholdMs -Profile $profile
                $totalProfileElapsedMs += $elapsedMs
                _LogMonitorProfileDuration -Prefix $prefix -Profile $profile -ElapsedMs $elapsedMs
                if ($elapsedMs -ge $profileSlowThresholdMs) {
                    $slowProfiles.Add("${prefix}:${elapsedMs}ms") | Out-Null
                }
            }
        }

        $tickElapsedMs = [int][Math]::Round(((Get-Date) - $tickStartedAt).TotalMilliseconds)
        $slowSummary = if ($slowProfiles.Count -gt 0) { $slowProfiles -join ', ' } else { 'none' }
        _Log "Monitor tick complete: total=${tickElapsedMs}ms profileWork=${totalProfileElapsedMs}ms slowProfiles=$slowSummary" -Level DEBUG
    }
}

# -----------------------------------------------------------------------------
#  SOURCE RCON PROTOCOL  -  Pure .NET, no external dependencies
#  Packet layout: [int32 size][int32 id][int32 type][body string][0x00 0x00]
#  Type 3 = SERVERDATA_AUTH
#  Type 2 = SERVERDATA_AUTH_RESPONSE / SERVERDATA_EXECCOMMAND
#  Type 0 = SERVERDATA_RESPONSE_VALUE
# -----------------------------------------------------------------------------
function Invoke-RconCommand {
    param(
        [string]$Host     = '127.0.0.1',
        [int]   $Port     = 25575,
        [string]$Password,
        [string]$Command,
        [int]   $TimeoutMs = 5000,
        [switch]$CaptureDebug
    )
    function _RconPacket {
        param([int]$Id, [int]$Type, [string]$Body)
        $bodyBytes  = [System.Text.Encoding]::UTF8.GetBytes($Body)
        # size = 4 (id) + 4 (type) + body + 2 (null terminators)
        $size       = 4 + 4 + $bodyBytes.Length + 2
        $packet     = [System.Collections.Generic.List[byte]]::new()
        # Write size as little-endian int32
        $packet.AddRange([System.BitConverter]::GetBytes([int32]$size))
        $packet.AddRange([System.BitConverter]::GetBytes([int32]$Id))
        $packet.AddRange([System.BitConverter]::GetBytes([int32]$Type))
        $packet.AddRange($bodyBytes)
        $packet.Add(0)   # null terminator 1
        $packet.Add(0)   # null terminator 2
        return $packet.ToArray()
    }

    function _ReadInt32 {
        param([System.IO.BinaryReader]$Reader)
        $bytes = $Reader.ReadBytes(4)
        if ($bytes.Length -lt 4) { throw "Connection closed unexpectedly" }
        return [System.BitConverter]::ToInt32($bytes, 0)
    }

    function _ReadPacket {
        param([System.IO.BinaryReader]$Reader)
        $size = _ReadInt32 $Reader
        if ($size -lt 10 -or $size -gt 4096) { throw "Invalid RCON packet size: $size" }
        $data = $Reader.ReadBytes($size)
        if ($data.Length -lt $size) { throw "Incomplete RCON packet" }
        $id   = [System.BitConverter]::ToInt32($data, 0)
        $type = [System.BitConverter]::ToInt32($data, 4)
        # Body is everything from byte 8 to end minus the 2 null terminators
        $bodyLen = $size - 4 - 4 - 2
        $body = if ($bodyLen -gt 0) {
            [System.Text.Encoding]::UTF8.GetString($data, 8, $bodyLen)
        } else { '' }
        return @{ Id = $id; Type = $type; Body = $body }
    }

    $tcp    = $null
    $stream = $null
    $reader = $null
    $writer = $null
    $dbg    = New-Object 'System.Collections.Generic.List[string]'

    try {
        _Log "RCON DEBUG: connect ${Host}:${Port} cmd='$Command'" -Level DEBUG
        if ($CaptureDebug) { $dbg.Add("connect ${Host}:${Port}") | Out-Null }
        # Connect
        $tcp = [System.Net.Sockets.TcpClient]::new()
        $tcp.ReceiveTimeout = $TimeoutMs
        $tcp.SendTimeout    = $TimeoutMs

        $connectResult = $tcp.BeginConnect($Host, $Port, $null, $null)
        $connected     = $connectResult.AsyncWaitHandle.WaitOne($TimeoutMs)
        if (-not $connected -or -not $tcp.Connected) {
            throw "Could not connect to $Host`:$Port within ${TimeoutMs}ms"
        }
        $tcp.EndConnect($connectResult)
        if ($CaptureDebug) { $dbg.Add("connected") | Out-Null }

        $stream = $tcp.GetStream()
        $reader = [System.IO.BinaryReader]::new($stream)
        $writer = [System.IO.BinaryWriter]::new($stream)

        # Authenticate (type 3, id 1)
        $authPacket = _RconPacket -Id 1 -Type 3 -Body $Password
        $writer.Write($authPacket)
        $writer.Flush()
        if ($CaptureDebug) { $dbg.Add("auth sent") | Out-Null }

        # Read auth response - server sends two packets on auth
        $resp1 = _ReadPacket $reader
        # Some servers send an empty type-0 first, then the auth response
        if ($resp1.Type -eq 0) {
            $resp1 = _ReadPacket $reader
        }

        if ($resp1.Id -eq -1) {
            throw "RCON authentication failed - wrong password"
        }
        if ($CaptureDebug) { $dbg.Add("auth ok") | Out-Null }

        # Send command (type 2, id 2)
        $cmdPacket = _RconPacket -Id 2 -Type 2 -Body $Command
        $writer.Write($cmdPacket)
        $writer.Flush()
        if ($CaptureDebug) { $dbg.Add("command sent") | Out-Null }

        # Read response - collect all packets with id=2
        $responseBody = [System.Text.StringBuilder]::new()
        $deadline     = (Get-Date).AddMilliseconds($TimeoutMs)

        while ((Get-Date) -lt $deadline) {
            try {
                $resp = _ReadPacket $reader
                if ($resp.Id -eq 2) {
                    $responseBody.Append($resp.Body) | Out-Null
                    # If the body is shorter than max packet, we have everything
                    if ($resp.Body.Length -lt 4086) { break }
                }
            } catch {
                # Timeout reading next packet means we have all the data
                break
            }
        }

        if ($CaptureDebug) { $dbg.Add("response length: $($responseBody.Length)") | Out-Null }
        return @{
            Success  = $true
            Response = $responseBody.ToString().Trim()
            Debug    = @($dbg)
        }
    }
    catch {
        _Log "RCON DEBUG: error=$_" -Level DEBUG
        return @{
            Success  = $false
            Response = ''
            Error    = "$_"
            Debug    = @($dbg)
        }
    }
    finally {
        if ($reader) { try { $reader.Close() } catch {} }
        if ($writer) { try { $writer.Close() } catch {} }
        if ($stream) { try { $stream.Close() } catch {} }
        if ($tcp)    { try { $tcp.Close()    } catch {} }
    }
}

# -----------------------------------------------------------------------------
#  TELNET COMMAND  -  Plain-text TCP protocol used by 7 Days to Die
#
#  Exchange flow:
#    Server sends  : "Please enter password:\r\n"
#    Client sends  : "<password>\r\n"
#    Server sends  : "Logon successful.\r\n"
#    Client sends  : "<command>\r\n"
#    Server sends  : response lines, ends with a bare ">" prompt line
#    Client sends  : "exit\r\n"  (clean disconnect)
#
#  All reads use a deadline poll loop so a slow server never hangs forever.
# -----------------------------------------------------------------------------
function Invoke-TelnetCommand {
    param(
        [string]$Host      = '127.0.0.1',
        [int]   $Port      = 8081,
        [string]$Password  = '',
        [string]$Command   = '',
        [int]   $TimeoutMs = 6000,
        [switch]$CaptureDebug
    )

    $tcp    = $null
    $stream = $null
    $reader = $null
    $writer = $null

    # Read lines from the stream until the deadline or StopPattern matches.
    # Uses DataAvailable polling so we never block past the deadline.
    function _TelnetRead {
        param(
            [System.IO.StreamReader]$R,
            [string]$StopPattern,
            [datetime]$Until
        )
        $lines = [System.Collections.Generic.List[string]]::new()
        while ((Get-Date) -lt $Until) {
            if ($R.BaseStream.DataAvailable) {
                $line = $R.ReadLine()
                if ($null -eq $line) { break }
                $lines.Add($line)
                if ($StopPattern -and $line -match $StopPattern) { break }
            } else {
                Start-Sleep -Milliseconds 50
            }
        }
        return $lines
    }

    $dbg = New-Object 'System.Collections.Generic.List[string]'

    try {
        _Log "Telnet DEBUG: connect ${Host}:${Port} cmd='$Command'" -Level DEBUG
        if ($CaptureDebug) { $dbg.Add("connect ${Host}:${Port}") | Out-Null }

        $tcp = [System.Net.Sockets.TcpClient]::new()
        $tcp.ReceiveTimeout = $TimeoutMs
        $tcp.SendTimeout    = $TimeoutMs

        $ar        = $tcp.BeginConnect($Host, $Port, $null, $null)
        $connected = $ar.AsyncWaitHandle.WaitOne($TimeoutMs)
        if (-not $connected -or -not $tcp.Connected) {
            throw "Could not connect to ${Host}:${Port} within ${TimeoutMs}ms"
        }
        $tcp.EndConnect($ar)

        $stream          = $tcp.GetStream()
        $reader          = [System.IO.StreamReader]::new($stream, [System.Text.Encoding]::UTF8)
        $writer          = [System.IO.StreamWriter]::new($stream, [System.Text.Encoding]::UTF8)
        $writer.AutoFlush = $true
        $writer.NewLine   = "`r`n"

        $deadline = (Get-Date).AddMilliseconds($TimeoutMs)

        # Step 1: Kick the server with a blank line so it prints a prompt (some builds require this)
        try { $writer.WriteLine('') } catch { }
        if ($CaptureDebug) { $dbg.Add("sent initial newline") | Out-Null }

        # Step 2: Wait for password prompt (case-insensitive)
        $prompt = _TelnetRead -R $reader -StopPattern '(?i)password' -Until $deadline
        _Log "Telnet DEBUG: prompt lines=$($prompt.Count)" -Level DEBUG
        if ($CaptureDebug -and $prompt.Count -gt 0) {
            $dbg.Add("prompt: " + (($prompt -join ' | ') -replace '\s+',' ')) | Out-Null
        }

        # Step 3: Authenticate (even if prompt was not seen)
        $writer.WriteLine($Password)
        $deadline  = (Get-Date).AddMilliseconds($TimeoutMs)
        $authLines = _TelnetRead -R $reader -StopPattern '(?i)successful|incorrect|failed' -Until $deadline
        $authText  = $authLines -join ' '
        _Log "Telnet DEBUG: auth response: $authText" -Level DEBUG
        if ($CaptureDebug) { $dbg.Add("auth response: $authText") | Out-Null }

        if ($authText -match 'incorrect|failed|wrong') {
            throw "Telnet authentication failed for ${Host}:${Port} - check TelnetPassword in profile and serverconfig.xml"
        }
        if ($authLines.Count -eq 0 -and $prompt.Count -eq 0) {
            throw "No telnet prompt/auth response from ${Host}:${Port} (server may not be listening on this port)"
        }

        # Step 3: Send command
        $writer.WriteLine($Command)
        $deadline  = (Get-Date).AddMilliseconds($TimeoutMs)
        if ($CaptureDebug) { $dbg.Add("sent command") | Out-Null }

        # Read until the bare ">" prompt appears - that signals end of output
        $respLines = _TelnetRead -R $reader -StopPattern '^>\s*$' -Until $deadline

        # Step 4: Clean disconnect
        try { $writer.WriteLine('exit') } catch { }

        # Strip the trailing prompt and blanks
        $cleanLines = $respLines | Where-Object { $_ -notmatch '^>\s*$' -and $_.Trim() -ne '' }
        $response   = ($cleanLines -join "`n").Trim()
        if ($CaptureDebug) { $dbg.Add("response lines: $($respLines.Count)") | Out-Null }

        _Log "Telnet DEBUG: response lines=$($respLines.Count) cleaned='$response'" -Level DEBUG

        return @{ Success = $true; Response = $response; Debug = @($dbg) }
    }
    catch {
        _Log "Telnet error: $_" -Level WARN
        return @{ Success = $false; Response = ''; Error = "$_"; Debug = @($dbg) }
    }
    finally {
        if ($reader) { try { $reader.Close() } catch {} }
        if ($writer) { try { $writer.Close() } catch {} }
        if ($stream) { try { $stream.Close() } catch {} }
        if ($tcp)    { try { $tcp.Close()    } catch {} }
    }
}

# -----------------------------------------------------------------------------
#  INVOKE RAW COMMAND STRING  -  UI helper for Command Catalog
#  Auto-selects best transport (Telnet > RCON > stdin) based on profile config
# -----------------------------------------------------------------------------
function Invoke-ServerCommandText {
    param(
        [string]$Prefix,
        [string]$Command,
        [switch]$ForceRest,
        [switch]$ForceSatisfactoryApi,
        [switch]$VerboseDebug,
        [hashtable]$SharedState = $null
    )
    $debug = New-Object 'System.Collections.Generic.List[string]'
    try {
        $prefix = $Prefix.ToUpper()
        $state  = if ($null -ne $SharedState) { $SharedState } else { $script:State }

    if ($null -eq $state -or -not $state.ContainsKey('Profiles')) {
        return @{ Success = $false; Message = "[ERROR] No profiles available" }
    }

    if ([string]::IsNullOrWhiteSpace($Command)) {
        return @{ Success = $false; Message = "[ERROR] Command is empty" }
    }

    if (-not $state.Profiles.ContainsKey($prefix)) {
        return @{ Success = $false; Message = "[ERROR] Profile '$prefix' not found" }
    }

    $profile = $state.Profiles[$prefix]

    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    $status = Get-ServerStatus -Prefix $prefix
    if (-not $status.Running) {
        return @{ Success = $false; Message = "[OFFLINE] $($profile.GameName) is not running."; Debug = @($debug) }
    }

    $cmd = $Command.Trim()
    if ($VerboseDebug) { $debug.Add("raw cmd: $Command") | Out-Null }

    # If the user typed a REST-style verb, force REST routing immediately
    if (-not $ForceRest -and $cmd -match '^(GET|POST)\s+') {
        $ForceRest = $true
        if ($VerboseDebug) { $debug.Add("auto: detected REST verb, forcing REST") | Out-Null }
    }

    # Games that expect console commands WITHOUT a leading "/" (only when not REST)
    $noSlash = @('7DaysToDie', 'ProjectZomboid', 'Satisfactory', 'Valheim', 'Palworld')
    if (-not $ForceRest -and $cmd.StartsWith('/')) {
        foreach ($name in $noSlash) {
            if (_TestProfileGame -Profile $profile -KnownGame $name) {
                $cmd = $cmd.TrimStart('/')
                break
            }
        }
    }
    if ($VerboseDebug -and $cmd -ne $Command.Trim()) {
        $debug.Add("normalized cmd: $cmd") | Out-Null
    }

    # Satisfactory: map common console-style/admin commands to HTTPS API.
    # The dedicated server accepts these reliably via API, while stdin often
    # appears to succeed locally without affecting the live server.
    if ((_TestProfileGame -Profile $profile -KnownGame 'Satisfactory') -and -not $ForceRest -and -not $ForceSatisfactoryApi) {
        $rawSf = ($cmd -replace '^/','').Trim()
        $sfName = $rawSf
        $sfArgText = ''
        if ($rawSf -match '^([A-Za-z0-9_.]+)(?:\s+(.+))?$') {
            $sfName = $Matches[1]
            $sfArgText = if ($Matches.Count -ge 3) { "$($Matches[2])".Trim() } else { '' }
        }

        $sfCanonical = $sfName.ToLowerInvariant()
        $sfApiCmd = $null

        switch ($sfCanonical) {
            'save' { $sfApiCmd = if ($sfArgText) { "SaveGame $sfArgText" } else { 'SaveGame' } }
            'savegame' { $sfApiCmd = if ($sfArgText) { "SaveGame $sfArgText" } else { 'SaveGame' } }
            'server.savegame' { $sfApiCmd = if ($sfArgText) { "SaveGame $sfArgText" } else { 'SaveGame' } }
            'players' { $sfApiCmd = 'QueryServerState' }
            'status' { $sfApiCmd = 'QueryServerState' }
            'querystate' { $sfApiCmd = 'QueryServerState' }
            'queryserverstate' { $sfApiCmd = 'QueryServerState' }
            'server.queryserverstate' { $sfApiCmd = 'QueryServerState' }
            'stop' { $sfApiCmd = 'Shutdown' }
            'quit' { $sfApiCmd = 'Shutdown' }
            'exit' { $sfApiCmd = 'Shutdown' }
            'shutdown' { $sfApiCmd = 'Shutdown' }
            'server.shutdown' { $sfApiCmd = 'Shutdown' }
        }

        if ($sfApiCmd) {
            $cmd = $sfApiCmd
            $ForceSatisfactoryApi = $true
            if ($VerboseDebug) { $debug.Add("satisfactory: mapped $sfName -> API $cmd") | Out-Null }
        }
    }

    # Palworld: map console-style commands to REST (stdin is unreliable)
    if (_TestProfileGame -Profile $profile -KnownGame 'Palworld') {
        $alreadyVerb = ($cmd -match '^(GET|POST)\s+/')
        function _SplitArgs([string]$text) {
            $args = @()
            if ([string]::IsNullOrWhiteSpace($text)) { return $args }
            $sb = New-Object System.Text.StringBuilder
            $inQuote = $false
            foreach ($ch in $text.ToCharArray()) {
                if ($ch -eq '"') { $inQuote = -not $inQuote; continue }
                if (-not $inQuote -and [char]::IsWhiteSpace($ch)) {
                    if ($sb.Length -gt 0) { $args += $sb.ToString(); $sb.Clear() | Out-Null }
                    continue
                }
                $null = $sb.Append($ch)
            }
            if ($sb.Length -gt 0) { $args += $sb.ToString() }
            return $args
        }

        $raw = ($cmd -replace '^/','').Trim()
        $parts = _SplitArgs $raw
        if ($parts.Count -gt 0 -and -not $alreadyVerb) {
            $name = $parts[0].ToLowerInvariant()
            $restCmd = $null

            switch ($name) {
                'save'        { $restCmd = 'POST /v1/api/save' }
                'info'        { $restCmd = 'GET /v1/api/info' }
                'showplayers' { $restCmd = 'GET /v1/api/players' }
                'players'     { $restCmd = 'GET /v1/api/players' }
                'settings'    { $restCmd = 'GET /v1/api/settings' }
                'metrics'     { $restCmd = 'GET /v1/api/metrics' }
                'stop'        { $restCmd = 'POST /v1/api/stop' }
                'doexit'      { $restCmd = 'POST /v1/api/stop' }
                'shutdown' {
                    $wait = 30
                    $msg = 'Server shutting down'
                    if ($parts.Count -ge 2 -and [int]::TryParse($parts[1], [ref]$null)) {
                        $wait = [int]$parts[1]
                        if ($parts.Count -ge 3) { $msg = ($parts[2..($parts.Count-1)] -join ' ') }
                    } elseif ($parts.Count -ge 2) {
                        $msg = ($parts[1..($parts.Count-1)] -join ' ')
                    }
                    $body = @{ waittime = $wait; message = $msg } | ConvertTo-Json -Compress -Depth 4
                    $restCmd = "POST /v1/api/shutdown $body"
                }
                'broadcast' {
                    if ($parts.Count -ge 2) {
                        $msg = ($parts[1..($parts.Count-1)] -join ' ')
                        $body = @{ message = $msg } | ConvertTo-Json -Compress -Depth 4
                        $restCmd = "POST /v1/api/announce $body"
                    }
                }
                'kickplayer' {
                    if ($parts.Count -ge 2) {
                        $uid = $parts[1]
                        $msg = if ($parts.Count -ge 3) { ($parts[2..($parts.Count-1)] -join ' ') } else { 'You have been kicked' }
                        $body = @{ userid = $uid; message = $msg } | ConvertTo-Json -Compress -Depth 4
                        $restCmd = "POST /v1/api/kick $body"
                    }
                }
                'banplayer' {
                    if ($parts.Count -ge 2) {
                        $uid = $parts[1]
                        $msg = if ($parts.Count -ge 3) { ($parts[2..($parts.Count-1)] -join ' ') } else { 'You have been banned' }
                        $body = @{ userid = $uid; message = $msg } | ConvertTo-Json -Compress -Depth 4
                        $restCmd = "POST /v1/api/ban $body"
                    }
                }
                'unbanplayer' {
                    if ($parts.Count -ge 2) {
                        $uid = $parts[1]
                        $body = @{ userid = $uid } | ConvertTo-Json -Compress -Depth 4
                        $restCmd = "POST /v1/api/unban $body"
                    }
                }
            }

            if ($restCmd) {
                $cmd = $restCmd
                $ForceRest = $true
                if ($VerboseDebug) { $debug.Add("palworld: mapped $name -> REST $cmd") | Out-Null }
            }
        }
    }

    # Allow explicit REST prefix in command text
    if ($cmd -match '^(?i)rest\s*[:\s]\s*(.+)$') {
        $cmd = $Matches[1].Trim()
        $ForceRest = $true
    }

    # Allow explicit Satisfactory API prefix
    if ($cmd -match '^(?i)api\s*[:\s]\s*(.+)$') {
        $cmd = $Matches[1].Trim()
        $ForceSatisfactoryApi = $true
    }

    if ($ForceRest) {
        $restHost = if ($null -ne $profile.RestHost)     { $profile.RestHost }     else { '127.0.0.1' }
        $restPort = if ($null -ne $profile.RestPort)     { [int]$profile.RestPort } else { 8212 }
        $restPass = if ($null -ne $profile.RestPassword) { $profile.RestPassword } else { '' }

        if ($restPort -le 0) {
            return @{ Success = $false; Message = "[ERROR] REST port not configured for $($profile.GameName)."; Debug = @($debug) }
        }

        $method = 'GET'
        $endpoint = $cmd
        $bodyObj = $null

        if ($cmd -match '^(GET|POST)\s+(\S+)(?:\s+(.+))?$') {
            $method = $Matches[1].ToUpper()
            $endpoint = $Matches[2]
            $bodyText = $Matches[3]
            if ($method -eq 'POST' -and -not [string]::IsNullOrWhiteSpace($bodyText)) {
                $rawBody = $bodyText.Trim()
                try {
                    $bodyObj = $rawBody | ConvertFrom-Json
                } catch {
                    $bodyObj = $rawBody
                }
            }
        }

        # Palworld: normalize endpoint/method (REST is case-sensitive)
        if (_TestProfileGame -Profile $profile -KnownGame 'Palworld') {
            # Normalize leading / and lowercase known v1 endpoints
            if ($endpoint -match '^/v1/api/') {
                $endpoint = $endpoint.ToLowerInvariant()
            } elseif ($endpoint -match '^/') {
                $endpoint = $endpoint.ToLowerInvariant()
            } else {
                $endpoint = $endpoint.ToLowerInvariant()
            }

            # Enforce correct method for POST-only endpoints
            if ($endpoint -match '^/v1/api/(save|shutdown|stop|announce|kick|ban|unban)$') {
                $method = 'POST'
            }
        }

        # Allow full URL in command text (e.g., POST http://localhost:8212/v1/api/save)
        $explicitUrl = $null
        if ($endpoint -match '^(?i)https?://') {
            try { $explicitUrl = [System.Uri]$endpoint } catch { $explicitUrl = $null }
            if ($explicitUrl) {
                $restHost = $explicitUrl.Host
                if ($explicitUrl.Port -gt 0) { $restPort = $explicitUrl.Port }
                $endpoint = $explicitUrl.AbsolutePath
                if ($explicitUrl.Query) { $endpoint = "$endpoint$($explicitUrl.Query)" }
            }
        }

        if (-not $endpoint.StartsWith('/')) {
            return @{ Success = $false; Message = "[ERROR] REST endpoint must start with '/'. Example: /v1/api/players"; Debug = @($debug) }
        }

        # Palworld REST API lives under /v1/api. Auto-prefix if missing.
        if ((_TestProfileGame -Profile $profile -KnownGame 'Palworld') -and $endpoint -notmatch '^/v1/api/') {
            $endpoint = "/v1/api$endpoint"
            if ($VerboseDebug) { $debug.Add("endpoint normalized: $endpoint") | Out-Null }
        }
        if (_TestProfileGame -Profile $profile -KnownGame 'Palworld') {
            $endpoint = $endpoint.ToLowerInvariant()
            if ($endpoint -match '^/v1/api/(save|shutdown|stop|announce|kick|ban|unban)$') { $method = 'POST' }
        }

        _Log "[$($profile.GameName)] REST (UI) -> $restHost`:$restPort $method $endpoint" -Level DEBUG
        if ($VerboseDebug) {
            $debug.Add("transport: REST") | Out-Null
            $debug.Add("host: $restHost") | Out-Null
            $debug.Add("port: $restPort") | Out-Null
            $debug.Add("method: $method") | Out-Null
            $debug.Add("endpoint: $endpoint") | Out-Null
        }

        $restResult = Invoke-PalworldRestRequest `
            -RestHost $restHost `
            -Port $restPort `
            -Password $restPass `
            -Endpoint $endpoint `
            -Method $method `
            -Profile $profile `
            -Body $bodyObj `
            -ReturnMeta:($VerboseDebug -or $true)

        if ($restResult -and $restResult.Success) {
            if ($VerboseDebug) {
                $debug.Add("=== REST API REQUEST ===") | Out-Null
                $debug.Add("Method      : $($restResult.Method)") | Out-Null
                $debug.Add("URL         : $($restResult.Url)") | Out-Null
                $debug.Add("Auth        : $($restResult.AuthUsed)") | Out-Null
                $bodySent = if ($restResult.BodySent) { "$($restResult.BodySent)" } else { '' }
                if ($bodySent.Length -gt 400) { $bodySent = $bodySent.Substring(0, 400) + '...' }
                $debug.Add("Body        : $bodySent") | Out-Null
                $debug.Add("=======================") | Out-Null
                $debug.Add("=== REST API RESPONSE ===") | Out-Null
                $debug.Add("Status      : $($restResult.Status)") | Out-Null
                $bodyTrim = if ($restResult.Body) { "$($restResult.Body)" } else { '' }
                if ($bodyTrim.Length -gt 400) { $bodyTrim = $bodyTrim.Substring(0, 400) + '...' }
                $debug.Add("Response    : $bodyTrim") | Out-Null
                $debug.Add("========================") | Out-Null
            }
            # Push REST summary into game log queue (so it shows in server log tab)
            try {
                if ($state -and $state.ContainsKey('GameLogQueue')) {
                    $line = "[REST] $($restResult.Method) $($restResult.Url) -> $($restResult.Status)"
                    if ($restResult.Body) {
                        $bodyLine = "$($restResult.Body)"
                        if ($bodyLine.Length -gt 300) { $bodyLine = $bodyLine.Substring(0,300) + '...' }
                        $line = "$line | $bodyLine"
                    }
                    $state.GameLogQueue.Enqueue([pscustomobject]@{ Prefix = $prefix; Line = $line; Path = '' })
                }
            } catch { }
            $sw.Stop()
            if ($VerboseDebug) { $debug.Add("duration_ms: $($sw.ElapsedMilliseconds)") | Out-Null }
            return @{ Success = $true; Message = "[REST] $($profile.GameName): OK"; Debug = @($debug) }
        }
        if ($VerboseDebug -and $restResult) {
            $debug.Add("=== REST API REQUEST ===") | Out-Null
            $debug.Add("Method      : $($restResult.Method)") | Out-Null
            $debug.Add("URL         : $($restResult.Url)") | Out-Null
            $debug.Add("Auth        : $($restResult.AuthUsed)") | Out-Null
            $bodySent = if ($restResult.BodySent) { "$($restResult.BodySent)" } else { '' }
            if ($bodySent.Length -gt 400) { $bodySent = $bodySent.Substring(0, 400) + '...' }
            $debug.Add("Body        : $bodySent") | Out-Null
            $debug.Add("=======================") | Out-Null
            $debug.Add("=== REST API RESPONSE ===") | Out-Null
            $debug.Add("Status      : $($restResult.Status)") | Out-Null
            $bodyTrim = if ($restResult.Body) { "$($restResult.Body)" } else { '' }
            if ($bodyTrim.Length -gt 400) { $bodyTrim = $bodyTrim.Substring(0, 400) + '...' }
            $debug.Add("Response    : $bodyTrim") | Out-Null
            if ($restResult.Error) { $debug.Add("Error       : $($restResult.Error)") | Out-Null }
            $debug.Add("========================") | Out-Null
        }
        # Push REST failure into game log queue
        try {
            if ($state -and $state.ContainsKey('GameLogQueue')) {
                $line = "[REST][ERROR] $($profile.GameName): $($restResult.Status) $($restResult.Error)"
                $state.GameLogQueue.Enqueue([pscustomobject]@{ Prefix = $prefix; Line = $line; Path = '' })
            }
        } catch { }
        $sw.Stop()
        if ($VerboseDebug) { $debug.Add("duration_ms: $($sw.ElapsedMilliseconds)") | Out-Null }
        return @{ Success = $false; Message = "[ERROR] REST command failed for $($profile.GameName)."; Debug = @($debug) }
    }

    if ($ForceSatisfactoryApi) {
        $sfHost  = if ($null -ne $profile.SatisfactoryApiHost) { $profile.SatisfactoryApiHost } else { '127.0.0.1' }
        $sfPort  = if ($null -ne $profile.SatisfactoryApiPort) { [int]$profile.SatisfactoryApiPort } else { 7777 }
        $sfToken = if ($null -ne $profile.SatisfactoryApiToken) { $profile.SatisfactoryApiToken } else { '' }

        if (-not $sfPort -or $sfPort -le 0) {
            return @{ Success = $false; Message = "[ERROR] Satisfactory API port not configured."; Debug = @($debug) }
        }

        $func = $null
        $data = $null

        if ($cmd -match '^([A-Za-z0-9_.]+)(?:\s+(.+))?$') {
            $func = $Matches[1]
            $rest = $Matches[2]
            if (-not [string]::IsNullOrWhiteSpace($rest)) {
                $raw = $rest.Trim()
                # If JSON provided, parse it. Otherwise treat as SaveGame name.
                if ($raw.StartsWith('{') -and $raw.EndsWith('}')) {
                    try { $data = $raw | ConvertFrom-Json } catch { $data = $null }
                } elseif ($func -eq 'SaveGame') {
                    $data = @{ SaveName = $raw }
                }
            }
        }

        # Normalize legacy/console-style "server.SaveGame" -> "SaveGame"
        if ($func -and $func -match '^(?i)server\.(.+)$') {
            $func = $Matches[1]
        }

        if (-not $func) {
            return @{ Success = $false; Message = "[ERROR] API function missing. Example: SaveGame or QueryServerState"; Debug = @($debug) }
        }

        if ($func -eq 'SaveGame' -and -not $data) {
            $data = @{ SaveName = 'ecc_ui' }
        }

        _Log "[$($profile.GameName)] SF API (UI) -> $func" -Level DEBUG
        if ($VerboseDebug) {
            $debug.Add("transport: SatisfactoryApi") | Out-Null
            $debug.Add("host: $sfHost") | Out-Null
            $debug.Add("port: $sfPort") | Out-Null
            $debug.Add("function: $func") | Out-Null
        }

        $r = Invoke-SatisfactoryApiRequest -Host $sfHost -Port $sfPort -Token $sfToken -Function $func -Data $data
        if ($r.Success) {
            if ($VerboseDebug) {
                if ($r.Status) { $debug.Add("status: $($r.Status)") | Out-Null }
                if ($r.Raw) {
                    $raw = $r.Raw
                    if ($raw.Length -gt 400) { $raw = $raw.Substring(0, 400) + '...' }
                    $debug.Add("body: $raw") | Out-Null
                }
            }
            $sw.Stop()
            if ($VerboseDebug) { $debug.Add("duration_ms: $($sw.ElapsedMilliseconds)") | Out-Null }
            if ($func -eq 'QueryServerState') {
                $players = 0
                $sessionName = 'Unknown'
                try { $players     = [int]$r.Data.serverGameState.numConnectedPlayers } catch { }
                try { $sessionName = $r.Data.serverGameState.activeSessionName        } catch { }
                return @{ Success = $true; Message = "[API] $($profile.GameName): Players=$players Session=$sessionName"; Debug = @($debug) }
            }
            return @{ Success = $true; Message = "[API] $($profile.GameName): $func OK"; Debug = @($debug) }
        }

        $hint = if ([string]::IsNullOrWhiteSpace($sfToken)) { " (Token missing?)" } else { "" }
        if ($VerboseDebug) {
            if ($r.Status) { $debug.Add("status: $($r.Status)") | Out-Null }
            if ($r.Raw) {
                $raw = $r.Raw
                if ($raw.Length -gt 400) { $raw = $raw.Substring(0, 400) + '...' }
                $debug.Add("body: $raw") | Out-Null
            }
            if ($r.Error) { $debug.Add("error: $($r.Error)") | Out-Null }
        }
        $sw.Stop()
        if ($VerboseDebug) { $debug.Add("duration_ms: $($sw.ElapsedMilliseconds)") | Out-Null }
        return @{ Success = $false; Message = "[ERROR] Satisfactory API failed: $($r.Error)$hint"; Debug = @($debug) }
    }

    # Prefer Telnet if configured (7 Days to Die)
    $telnetHost = if ($null -ne $profile.TelnetHost)     { $profile.TelnetHost }     else { '127.0.0.1' }
    $telnetPort = if ($null -ne $profile.TelnetPort)     { [int]$profile.TelnetPort } else { 8081 }
    $telnetPass = if ($null -ne $profile.TelnetPassword) { $profile.TelnetPassword } else { '' }

    if (-not [string]::IsNullOrWhiteSpace($telnetPass) -and $telnetPort -gt 0) {
        _Log "[$($profile.GameName)] Telnet (UI) -> ${telnetHost}:${telnetPort} cmd='$cmd'" -Level DEBUG
        if ($VerboseDebug) {
            $debug.Add("transport: Telnet") | Out-Null
            $debug.Add("host: $telnetHost") | Out-Null
            $debug.Add("port: $telnetPort") | Out-Null
            $debug.Add("cmd: $cmd") | Out-Null
        }

        $t = Invoke-TelnetCommand -Host $telnetHost -Port $telnetPort -Password $telnetPass -Command $cmd -CaptureDebug:($VerboseDebug -eq $true)
        if ($t.Success) {
            $resp = if ($t.Response) { $t.Response } else { 'command sent' }
            if ($VerboseDebug -and $t.Debug) {
                foreach ($d in $t.Debug) { $debug.Add([string]$d) | Out-Null }
            }
            # Push telnet response lines into the game log queue so they appear in the server log tab
            try {
                if ($state -and $state.ContainsKey('GameLogQueue') -and $t.Response) {
                    $lines = ($t.Response -split "(`r`n|`n|`r)")
                    foreach ($ln in $lines) {
                        if ([string]::IsNullOrWhiteSpace($ln)) { continue }
                        $state.GameLogQueue.Enqueue(
                            [pscustomobject]@{
                                Prefix = $prefix
                                Line   = "[TELNET] $ln"
                                Path   = ''
                            }
                        )
                    }
                }
            } catch { }
            $sw.Stop()
            if ($VerboseDebug) { $debug.Add("duration_ms: $($sw.ElapsedMilliseconds)") | Out-Null }
            return @{ Success = $true; Message = "[TELNET] $($profile.GameName): $resp"; Debug = @($debug) }
        }
        if ($VerboseDebug -and $t.Debug) {
            foreach ($d in $t.Debug) { $debug.Add([string]$d) | Out-Null }
        }
        $sw.Stop()
        if ($VerboseDebug) { $debug.Add("duration_ms: $($sw.ElapsedMilliseconds)") | Out-Null }
        return @{ Success = $false; Message = "[ERROR] Telnet failed for $($profile.GameName): $($t.Error)"; Debug = @($debug) }
    }

    # Next, try RCON if configured (Palworld, etc.)
    $rconHost = if ($null -ne $profile.RconHost)     { $profile.RconHost }     else { '127.0.0.1' }
    $rconPort = if ($null -ne $profile.RconPort)     { [int]$profile.RconPort } else { 25575 }
    $rconPass = if ($null -ne $profile.RconPassword) { $profile.RconPassword } else { '' }

    if (-not [string]::IsNullOrWhiteSpace($rconPass) -and $rconPort -gt 0) {
        _Log "[$($profile.GameName)] RCON (UI) -> ${rconHost}:${rconPort} cmd='$cmd'" -Level DEBUG
        if ($VerboseDebug) {
            $debug.Add("transport: RCON") | Out-Null
            $debug.Add("host: $rconHost") | Out-Null
            $debug.Add("port: $rconPort") | Out-Null
            $debug.Add("cmd: $cmd") | Out-Null
        }
        $r = Invoke-RconCommand -Host $rconHost -Port $rconPort -Password $rconPass -Command $cmd -CaptureDebug:($VerboseDebug -eq $true)
        if ($r.Success) {
            $resp = if ($r.Response) { $r.Response } else { 'command sent' }
            if ($VerboseDebug -and $r.Debug) {
                foreach ($d in $r.Debug) { $debug.Add([string]$d) | Out-Null }
            }
            $sw.Stop()
            if ($VerboseDebug) { $debug.Add("duration_ms: $($sw.ElapsedMilliseconds)") | Out-Null }
            return @{ Success = $true; Message = "[RCON] $($profile.GameName): $resp"; Debug = @($debug) }
        }
        if ($VerboseDebug -and $r.Debug) {
            foreach ($d in $r.Debug) { $debug.Add([string]$d) | Out-Null }
        }
        $sw.Stop()
        if ($VerboseDebug) { $debug.Add("duration_ms: $($sw.ElapsedMilliseconds)") | Out-Null }
        return @{ Success = $false; Message = "[ERROR] RCON failed for $($profile.GameName): $($r.Error)"; Debug = @($debug) }
    }

    # Fallback: stdin (best for local console)
    if ($VerboseDebug) {
        $debug.Add("transport: STDIN") | Out-Null
        $debug.Add("cmd: $cmd") | Out-Null
    }
    $stdinResult = Send-ServerStdin -Prefix $prefix -Command $cmd -ReturnDetail:($VerboseDebug -eq $true)
    if ($stdinResult -is [System.Collections.IDictionary]) {
        if ($VerboseDebug) {
            $debug.Add("stdin method: $($stdinResult.Method)") | Out-Null
            if ($stdinResult.ProcessName) { $debug.Add("process: $($stdinResult.ProcessName)") | Out-Null }
            if ($stdinResult.TargetPid -and [int]$stdinResult.TargetPid -gt 0) { $debug.Add("pid: $($stdinResult.TargetPid)") | Out-Null }
            if ($stdinResult.Error) { $debug.Add("stdin error: $($stdinResult.Error)") | Out-Null }
        }
        $sw.Stop()
        if ($VerboseDebug) { $debug.Add("duration_ms: $($sw.ElapsedMilliseconds)") | Out-Null }
        if ($stdinResult.Success) {
            return @{ Success = $true; Message = "[STDIN] $($profile.GameName): command sent"; Debug = @($debug) }
        }
        return @{ Success = $false; Message = "[ERROR] Failed to send command to $($profile.GameName)"; Debug = @($debug) }
    } else {
        $sw.Stop()
        if ($VerboseDebug) { $debug.Add("duration_ms: $($sw.ElapsedMilliseconds)") | Out-Null }
        if ($stdinResult) {
            return @{ Success = $true; Message = "[STDIN] $($profile.GameName): command sent"; Debug = @($debug) }
        }
        return @{ Success = $false; Message = "[ERROR] Failed to send command to $($profile.GameName)"; Debug = @($debug) }
    }
    }
    catch {
        try { _Log "[Invoke-ServerCommandText] Unhandled error: $_" -Level ERROR } catch { }
        return @{ Success = $false; Message = "[ERROR] Command failed: $_"; Debug = @($debug) }
    }
}

# -----------------------------------------------------------------------------
#  TRANSPORT TEST HELPERS (no side effects where possible)
# -----------------------------------------------------------------------------
function Test-TelnetConnection {
    param(
        [string]$Prefix,
        [int]$TimeoutMs = 5000,
        [switch]$VerboseDebug,
        [hashtable]$SharedState = $null
    )

    $prefix = $Prefix.ToUpper()
    $state  = if ($null -ne $SharedState) { $SharedState } else { $script:State }
    $debug = New-Object 'System.Collections.Generic.List[string]'

    if (-not $state -or -not $state.ContainsKey('Profiles') -or -not $state.Profiles.ContainsKey($prefix)) {
        return @{ Success = $false; Message = "[ERROR] Profile '$prefix' not found"; Debug = @($debug) }
    }

    $profile = $state.Profiles[$prefix]
    $status = Get-ServerStatus -Prefix $prefix
    if (-not $status.Running) {
        return @{ Success = $false; Message = "[OFFLINE] $($profile.GameName) is not running."; Debug = @($debug) }
    }

    $telnetHost = if ($null -ne $profile.TelnetHost)     { $profile.TelnetHost }     else { '127.0.0.1' }
    $telnetPort = if ($null -ne $profile.TelnetPort)     { [int]$profile.TelnetPort } else { 8081 }
    $telnetPass = if ($null -ne $profile.TelnetPassword) { $profile.TelnetPassword } else { '' }

    if ([string]::IsNullOrWhiteSpace($telnetPass)) {
        return @{ Success = $false; Message = "[ERROR] TelnetPassword not set for $($profile.GameName)."; Debug = @($debug) }
    }

    # Use a safe, read-only command for telnet test
    $testCmd = 'version'
    if ($VerboseDebug) {
        $debug.Add("transport: Telnet") | Out-Null
        $debug.Add("host: $telnetHost") | Out-Null
        $debug.Add("port: $telnetPort") | Out-Null
        $debug.Add("cmd: $testCmd") | Out-Null
    }

    $t = Invoke-TelnetCommand -Host $telnetHost -Port $telnetPort -Password $telnetPass -Command $testCmd -TimeoutMs $TimeoutMs -CaptureDebug:($VerboseDebug -eq $true)
    if ($VerboseDebug -and $t.Debug) {
        foreach ($d in $t.Debug) { $debug.Add([string]$d) | Out-Null }
    }
    if ($t.Success) {
        return @{ Success = $true; Message = "[TELNET] $($profile.GameName): OK"; Debug = @($debug) }
    }
    return @{ Success = $false; Message = "[ERROR] Telnet test failed for $($profile.GameName): $($t.Error)"; Debug = @($debug) }
}

function Test-RconConnection {
    param(
        [string]$Prefix,
        [int]$TimeoutMs = 5000,
        [switch]$VerboseDebug,
        [hashtable]$SharedState = $null
    )

    $prefix = $Prefix.ToUpper()
    $state  = if ($null -ne $SharedState) { $SharedState } else { $script:State }
    $debug = New-Object 'System.Collections.Generic.List[string]'

    if (-not $state -or -not $state.ContainsKey('Profiles') -or -not $state.Profiles.ContainsKey($prefix)) {
        return @{ Success = $false; Message = "[ERROR] Profile '$prefix' not found"; Debug = @($debug) }
    }

    $profile = $state.Profiles[$prefix]
    $status = Get-ServerStatus -Prefix $prefix
    if (-not $status.Running) {
        return @{ Success = $false; Message = "[OFFLINE] $($profile.GameName) is not running."; Debug = @($debug) }
    }

    $rconHost = if ($null -ne $profile.RconHost)     { $profile.RconHost }     else { '127.0.0.1' }
    $rconPort = if ($null -ne $profile.RconPort)     { [int]$profile.RconPort } else { 25575 }
    $rconPass = if ($null -ne $profile.RconPassword) { $profile.RconPassword } else { '' }

    if ([string]::IsNullOrWhiteSpace($rconPass)) {
        return @{ Success = $false; Message = "[ERROR] RCON password not set for $($profile.GameName)."; Debug = @($debug) }
    }

    $testCmd = if (_TestProfileGame -Profile $profile -KnownGame 'Palworld') { 'Info' } else { 'help' }
    if ($VerboseDebug) {
        $debug.Add("transport: RCON") | Out-Null
        $debug.Add("host: $rconHost") | Out-Null
        $debug.Add("port: $rconPort") | Out-Null
        $debug.Add("cmd: $testCmd") | Out-Null
    }

    $r = Invoke-RconCommand -Host $rconHost -Port $rconPort -Password $rconPass -Command $testCmd -TimeoutMs $TimeoutMs -CaptureDebug:($VerboseDebug -eq $true)
    if ($VerboseDebug -and $r.Debug) {
        foreach ($d in $r.Debug) { $debug.Add([string]$d) | Out-Null }
    }
    if ($r.Success) {
        return @{ Success = $true; Message = "[RCON] $($profile.GameName): OK"; Debug = @($debug) }
    }
    return @{ Success = $false; Message = "[ERROR] RCON test failed for $($profile.GameName): $($r.Error)"; Debug = @($debug) }
}

function Test-SatisfactoryApiConnection {
    param(
        [string]$Prefix,
        [int]$TimeoutMs = 8000,
        [switch]$VerboseDebug,
        [hashtable]$SharedState = $null
    )

    $prefix = $Prefix.ToUpper()
    $state  = if ($null -ne $SharedState) { $SharedState } else { $script:State }
    $debug = New-Object 'System.Collections.Generic.List[string]'

    if (-not $state -or -not $state.ContainsKey('Profiles') -or -not $state.Profiles.ContainsKey($prefix)) {
        return @{ Success = $false; Message = "[ERROR] Profile '$prefix' not found"; Debug = @($debug) }
    }

    $profile = $state.Profiles[$prefix]
    $status = Get-ServerStatus -Prefix $prefix
    if (-not $status.Running) {
        return @{ Success = $false; Message = "[OFFLINE] $($profile.GameName) is not running."; Debug = @($debug) }
    }

    $sfHost  = if ($null -ne $profile.SatisfactoryApiHost) { $profile.SatisfactoryApiHost } else { '127.0.0.1' }
    $sfPort  = if ($null -ne $profile.SatisfactoryApiPort) { [int]$profile.SatisfactoryApiPort } else { 7777 }
    $sfToken = if ($null -ne $profile.SatisfactoryApiToken) { $profile.SatisfactoryApiToken } else { '' }

    if (-not $sfPort -or $sfPort -le 0) {
        return @{ Success = $false; Message = "[ERROR] Satisfactory API port not configured."; Debug = @($debug) }
    }

    if ($VerboseDebug) {
        $debug.Add("transport: SatisfactoryApi") | Out-Null
        $debug.Add("host: $sfHost") | Out-Null
        $debug.Add("port: $sfPort") | Out-Null
        $debug.Add("function: QueryServerState") | Out-Null
    }

    $r = Invoke-SatisfactoryApiRequest -Host $sfHost -Port $sfPort -Token $sfToken -Function 'QueryServerState'
    if ($r.Success) {
        return @{ Success = $true; Message = "[API] $($profile.GameName): OK"; Debug = @($debug) }
    }
    return @{ Success = $false; Message = "[ERROR] Satisfactory API test failed: $($r.Error)"; Debug = @($debug) }
}

function Test-StdInConnection {
    param(
        [string]$Prefix,
        [switch]$VerboseDebug,
        [hashtable]$SharedState = $null
    )

    $prefix = $Prefix.ToUpper()
    $state  = if ($null -ne $SharedState) { $SharedState } else { $script:State }
    $debug = New-Object 'System.Collections.Generic.List[string]'

    if (-not $state -or -not $state.ContainsKey('Profiles') -or -not $state.Profiles.ContainsKey($prefix)) {
        return @{ Success = $false; Message = "[ERROR] Profile '$prefix' not found"; Debug = @($debug) }
    }

    $profile = $state.Profiles[$prefix]
    $status = Get-ServerStatus -Prefix $prefix
    if (-not $status.Running) {
        return @{ Success = $false; Message = "[OFFLINE] $($profile.GameName) is not running."; Debug = @($debug) }
    }

    if ($script:State.StdinHandles.ContainsKey($prefix)) {
        if ($VerboseDebug) { $debug.Add("stdin handle: available") | Out-Null }
        return @{ Success = $true; Message = "[STDIN] $($profile.GameName): handle OK"; Debug = @($debug) }
    }

    $procName = _CoalesceStr $profile.ProcessName ''
    if (-not $procName) {
        return @{ Success = $false; Message = "[ERROR] No ProcessName configured for stdin fallback."; Debug = @($debug) }
    }

    try {
        $targets = Get-Process -Name $procName -ErrorAction SilentlyContinue
        if (-not $targets -or $targets.Count -eq 0) {
            return @{ Success = $false; Message = "[ERROR] Process '$procName' not found for stdin fallback."; Debug = @($debug) }
        }
        $target = $targets | Select-Object -First 1
        if ($target.MainWindowHandle -eq [IntPtr]::Zero) {
            return @{ Success = $false; Message = "[ERROR] Process '$procName' has no main window handle."; Debug = @($debug) }
        }
        if ($VerboseDebug) { $debug.Add("stdin window: ok pid=$($target.Id)") | Out-Null }
        return @{ Success = $true; Message = "[STDIN] $($profile.GameName): window OK (pid $($target.Id))"; Debug = @($debug) }
    } catch {
        return @{ Success = $false; Message = "[ERROR] STDIN test failed: $_"; Debug = @($debug) }
    }
}

# -----------------------------------------------------------------------------
#  INVOKE PROFILE COMMAND  -  Routes Discord/GUI commands to the right action
#  Lives here in ServerManager so it has direct access to all server functions
# -----------------------------------------------------------------------------
function Invoke-ProfileCommand {
    param(
        [string]$Prefix,
        [string]$CommandName,
        [hashtable]$SharedState = $null
    )

    $prefix = $Prefix.ToUpper()

    # Use passed SharedState or fall back to module-level state
    $state = if ($null -ne $SharedState) { $SharedState } else { $script:State }

    if ($null -eq $state -or -not $state.ContainsKey('Profiles')) {
        return @{ Message = "[ERROR] No profiles available"; Success = $false }
    }

    $profiles = $state.Profiles

    if ($null -eq $profiles[$prefix]) {
        $available = @($profiles.Keys) -join ', '
        return @{ Message = "[ERROR] Profile '$prefix' not found. Available: $available"; Success = $false }
    }

    $profile = $profiles[$prefix]

    function _GetCommandStatusDetail {
        param(
            [string]$PrefixKey,
            [hashtable]$StateTable
        )

        if (-not $StateTable -or -not $StateTable.ContainsKey('ServerRuntimeState')) { return '' }
        if (-not $StateTable.ServerRuntimeState.ContainsKey($PrefixKey)) { return '' }

        try {
            $entry = $StateTable.ServerRuntimeState[$PrefixKey]
            if ($entry -and $entry.Detail -and "$($entry.Detail)".Trim() -ne '') {
                return "$($entry.Detail)".Trim()
            }
        } catch { }

        return ''
    }

    # ------------------------------------------------------------------
    # Robust command lookup
    # $profile.Commands may be a Hashtable, OrderedDictionary, or (if
    # ConvertTo-Json round-tripped badly) a PSCustomObject.  We try each
    # access method in turn so the command is never missed due to type
    # or case mismatches.
    # ------------------------------------------------------------------
    function _LookupCommand {
        param([object]$CmdsObj, [string]$Name)
        if ($null -eq $CmdsObj) { return $null }

        # 1. Standard dictionary indexer (Hashtable / OrderedDictionary)
        if ($CmdsObj -is [System.Collections.IDictionary]) {
            # Try exact key first
            if ($CmdsObj.Contains($Name)) { return $CmdsObj[$Name] }
            # Case-insensitive scan as fallback (handles OrderedDictionary case issues)
            foreach ($k in $CmdsObj.Keys) {
                if ([string]::Equals($k, $Name, [System.StringComparison]::OrdinalIgnoreCase)) {
                    return $CmdsObj[$k]
                }
            }
        }

        # 2. PSCustomObject (Commands was not converted by ConvertTo-Hashtable)
        if ($CmdsObj -is [psobject]) {
            foreach ($prop in $CmdsObj.PSObject.Properties) {
                if ([string]::Equals($prop.Name, $Name, [System.StringComparison]::OrdinalIgnoreCase)) {
                    return $prop.Value
                }
            }
        }

        return $null
    }

    function _ListCommandKeys {
        param([object]$CmdsObj)
        if ($null -eq $CmdsObj) { return @() }
        if ($CmdsObj -is [System.Collections.IDictionary]) { return @($CmdsObj.Keys) }
        if ($CmdsObj -is [psobject]) { return @($CmdsObj.PSObject.Properties.Name) }
        return @()
    }

    if ($null -eq $profile.Commands) {
        return @{ Message = "[ERROR] Profile '$prefix' has no Commands section."; Success = $false }
    }

    $cmdDef = _LookupCommand -CmdsObj $profile.Commands -Name $CommandName
    if ($null -eq $cmdDef) {
        $available = (_ListCommandKeys -CmdsObj $profile.Commands) -join ', '
        return @{ Message = "[ERROR] Command '$CommandName' not found. Available: $available"; Success = $false }
    }

    # Extract Type from cmdDef regardless of whether it is a hashtable or PSCustomObject
    $cmdType = $null
    if ($cmdDef -is [System.Collections.IDictionary]) {
        $cmdType = $cmdDef['Type']
        if ($null -eq $cmdType) {
            foreach ($k in $cmdDef.Keys) {
                if ([string]::Equals($k, 'Type', [System.StringComparison]::OrdinalIgnoreCase)) {
                    $cmdType = $cmdDef[$k]; break
                }
            }
        }
    } elseif ($cmdDef -is [psobject]) {
        $tp = $cmdDef.PSObject.Properties['Type']
        if ($tp) { $cmdType = $tp.Value }
    }
    if ($null -eq $cmdType -or "$cmdType".Trim() -eq '') { $cmdType = 'stdin' }

    _Log "[$($profile.GameName)] Command debug: name='$CommandName' type='$cmdType'" -Level DEBUG

    $result    = $false
    $resultMsg = ''

    switch ($cmdType) {
        'Start' {
            $result = Start-GameServer -Prefix $prefix
            $resultMsg = switch ($result) {
                'already_running' {
                    $status = Get-ServerStatus -Prefix $prefix
                    $uptime = if ($status.Running) { "$([Math]::Round($status.Uptime.TotalMinutes,1)) min" } else { 'active' }
                    New-DiscordGameMessage -Profile $profile -Prefix $prefix -Event 'already_running' -Values @{ Uptime = $uptime }
                }
                $true {
                    New-DiscordGameMessage -Profile $profile -Prefix $prefix -Event 'starting'
                }
                default {
                    New-DiscordGameMessage -Profile $profile -Prefix $prefix -Event 'error_start' -Values @{ Reason = 'Check the Program Log for the launch failure details.' }
                }
            }
        }
        'Stop' {
            $result    = Invoke-SafeShutdown -Prefix $prefix
            $resultMsg = if ($result) {
                New-DiscordGameMessage -Profile $profile -Prefix $prefix -Event 'waiting' -Values @{ WaitSeconds = if ($null -ne $profile.SaveWaitSeconds -and "$($profile.SaveWaitSeconds)".Trim() -ne '') { [int]$profile.SaveWaitSeconds } else { 15 } }
            } else {
                New-DiscordGameMessage -Profile $profile -Prefix $prefix -Event 'error_restart' -Values @{ Reason = 'ECC could not begin the safe shutdown sequence.' }
            }
        }
        'Restart' {
            $result    = Restart-GameServer -Prefix $prefix
            $resultMsg = if ($result) {
                New-DiscordGameMessage -Profile $profile -Prefix $prefix -Event 'restarting'
            } else {
                New-DiscordGameMessage -Profile $profile -Prefix $prefix -Event 'error_restart' -Values @{ Reason = 'ECC could not begin the restart sequence.' }
            }
        }
        'Status' {
            $status    = Get-ServerStatus -Prefix $prefix
            $result    = $true
            $uptime    = if ($status.Running) { "$([Math]::Round($status.Uptime.TotalMinutes,1)) min" } else { 'N/A' }
            $stateDetail = _GetCommandStatusDetail -PrefixKey $prefix -StateTable $state
            $resultMsg = if ($status.Running) {
                New-DiscordGameMessage -Profile $profile -Prefix $prefix -Event 'status_online' -Values @{
                    Pid = $status.Pid
                    Uptime = $uptime
                    StatusDetail = if ([string]::IsNullOrWhiteSpace($stateDetail)) { '' } else { "Status: $stateDetail" }
                }
            } else {
                New-DiscordGameMessage -Profile $profile -Prefix $prefix -Event 'status_offline' -Values @{
                    StatusDetail = if ([string]::IsNullOrWhiteSpace($stateDetail)) { '' } else { "Last state: $stateDetail" }
                }
            }
        }
        'SendCommand' {
            $cmd       = if ($null -ne $cmdDef.Command) { $cmdDef.Command } else { $CommandName }
            $result    = Send-ServerStdin -Prefix $prefix -Command $cmd
            $parsedPlayers = $null
            if ($result -and (_TestProfileGame -Profile $profile -KnownGame 'Hytale') -and $cmd -eq 'who') {
                Start-Sleep -Milliseconds 900
                $recentLogText = _ReadRecentProfileLogText -Profile $profile -Prefix $prefix -TailLines 140
                $parsedPlayers = _ParseHytaleWhoText -Text $recentLogText
                if ($parsedPlayers.Available) {
                    _SetLatestPlayersSnapshot -Prefix $prefix -Names @($parsedPlayers.Names) -Count ([int]$parsedPlayers.Count) -SharedState $state
                    if ($state -and $state.ContainsKey('PlayerQueryState')) {
                        $state.PlayerQueryState[$prefix] = [ordered]@{
                            LastAttemptAt = Get-Date
                            LastSuccessAt = Get-Date
                            Source        = 'HytaleWhoLog'
                            Note          = $parsedPlayers.Note
                        }
                    }
                }
            }
            $resultMsg = if ($result) {
                if ((_TestProfileGame -Profile $profile -KnownGame 'Hytale') -and $cmd -eq 'who' -and $parsedPlayers -and $parsedPlayers.Available) {
                    if ([int]$parsedPlayers.Count -le 0) {
                        "[PLAYERS] $($profile.GameName) - no players online."
                    } elseif (@($parsedPlayers.Names).Count -gt 0) {
                        "[PLAYERS] $($profile.GameName) - online: $((@($parsedPlayers.Names)) -join ', ')"
                    } else {
                        "[PLAYERS] $($profile.GameName) - players online: $([int]$parsedPlayers.Count)"
                    }
                } elseif ($CommandName -eq 'save') {
                    New-DiscordGameMessage -Profile $profile -Prefix $prefix -Event 'save_sent'
                } else {
                    New-DiscordGameMessage -Profile $profile -Prefix $prefix -Event 'command_ok' -Values @{ Command = "'$cmd'" }
                }
            } else {
                New-DiscordGameMessage -Profile $profile -Prefix $prefix -Event 'command_error' -Values @{
                    Command = "'$cmd'"
                    Reason  = 'Check the Program Log if the server ignored the command.'
                }
            }
        }
        'stdin' {
            $stdinCmd  = if ($null -ne $cmdDef.StdinCommand) { $cmdDef.StdinCommand } else { $CommandName }
            $result    = Send-ServerStdin -Prefix $prefix -Command $stdinCmd
            $resultMsg = if ($result) {
                if ($CommandName -eq 'save') {
                    New-DiscordGameMessage -Profile $profile -Prefix $prefix -Event 'save_sent'
                } else {
                    New-DiscordGameMessage -Profile $profile -Prefix $prefix -Event 'command_ok' -Values @{ Command = "'$stdinCmd'" }
                }
            } else {
                New-DiscordGameMessage -Profile $profile -Prefix $prefix -Event 'command_error' -Values @{
                    Command = "'$stdinCmd'"
                    Reason  = 'The command did not make it through the server input path.'
                }
            }
        }
        'Rcon' {
            $rconHost = if ($null -ne $profile.RconHost)     { $profile.RconHost     } else { '127.0.0.1' }
            $rconPort = if ($null -ne $profile.RconPort)     { [int]$profile.RconPort } else { 25575       }
            $rconPass = if ($null -ne $profile.RconPassword) { $profile.RconPassword } else { ''          }
            $rconCmd  = if ($null -ne $cmdDef.Command)       { $cmdDef.Command       } else { $CommandName }

            if (-not $rconPass) {
                $resultMsg = "[ERROR] RCON password not configured in profile for $($profile.GameName). Add RconPassword to the profile JSON."
            } else {
                _Log "[$($profile.GameName)] RCON -> ${rconHost}:${rconPort} cmd='$rconCmd'" -Level DEBUG
                $rconResult = Invoke-RconCommand -Host $rconHost -Port $rconPort -Password $rconPass -Command $rconCmd
                if ($rconResult.Success) {
                    $result    = $true
                    $resp      = $rconResult.Response
                    $resultMsg = if ($resp) { "[RCON] $($profile.GameName): $resp" } else { "[RCON] $($profile.GameName): command sent (no response)" }
                } else {
                    $resultMsg = "[ERROR] RCON failed for $($profile.GameName): $($rconResult.Error)"
                }
            }
        }
        'http' {
            $url = if ($null -ne $cmdDef.Url) { $cmdDef.Url } else { '' }
            if ($url) {
                $result    = $null -ne (Invoke-ServerHttp -Prefix $prefix -Url $url)
                $resultMsg = if ($result) { "[OK] HTTP command executed" } else { "[ERROR] HTTP command failed" }
            } else {
                $resultMsg = "[ERROR] No URL configured for HTTP command"
            }
        }
        'Rest' {
            # Generic REST command (used by Palworld ExtraCommands)
            $Resthost     = if ($null -ne $profile.RestHost)     { $profile.RestHost }     else { '127.0.0.1' }
            $port     = if ($null -ne $profile.RestPort)     { [int]$profile.RestPort } else { 8212        }
            $password = if ($null -ne $profile.RestPassword) { $profile.RestPassword } else { ''           }

            $endpoint = if ($null -ne $cmdDef.Endpoint) { $cmdDef.Endpoint } else { '' }
            $method   = if ($null -ne $cmdDef.Method)   { $cmdDef.Method   } else { 'GET' }
            $body     = if ($null -ne $cmdDef.Body)     { $cmdDef.Body     } else { $null }

            if (-not $endpoint) {
                $resultMsg = "[ERROR] No Endpoint configured for REST command"
            } else {
                _Log "[$($profile.GameName)] REST -> $resthost`:$port $method $endpoint" -Level DEBUG
                $restResult = Invoke-PalworldRestRequest `
                    -RestHost $restHost `
                    -Port $port `
                    -Password $password `
                    -Endpoint $endpoint `
                    -Method $method `
                    -Profile $Profile `
                    -Body $body

                if ($restResult) {
                    $result    = $true
                    $resultMsg = "[REST] $($profile.GameName): $restResult"
                } else {
                    $resultMsg = "[ERROR] REST command failed for $($profile.GameName). Check RESTAPIKey in PalWorldSettings.ini or set RestPassword in the profile."
                }
            }
        }
        'script' {
            $scriptPath = if ($null -ne $cmdDef.ScriptPath) { $cmdDef.ScriptPath } else { '' }
            if ($scriptPath) {
                $result    = Invoke-CustomScript -Prefix $prefix -ScriptPath $scriptPath
                $resultMsg = if ($result) { "[OK] Script executed" } else { "[ERROR] Script execution failed" }
            } else {
                $resultMsg = "[ERROR] No script path configured"
            }
        }
        'Telnet' {
            $telnetHost = if ($null -ne $profile.TelnetHost)     { $profile.TelnetHost }     else { '127.0.0.1' }
            $telnetPort = if ($null -ne $profile.TelnetPort)     { [int]$profile.TelnetPort } else { 8081        }
            $telnetPass = if ($null -ne $profile.TelnetPassword) { $profile.TelnetPassword } else { ''          }
            $telnetCmd  = if ($null -ne $cmdDef.Command)         { $cmdDef.Command         } else { $CommandName }

            if ([string]::IsNullOrWhiteSpace($telnetPass)) {
                $resultMsg = "[ERROR] TelnetPassword not set in profile for $($profile.GameName). Add TelnetPassword to the profile JSON and enable Telnet in serverconfig.xml."
            } else {
                _Log "[$($profile.GameName)] Telnet -> ${telnetHost}:${telnetPort} cmd='$telnetCmd'" -Level DEBUG
                $telnetResult = Invoke-TelnetCommand `
                    -Host      $telnetHost `
                    -Port      $telnetPort `
                    -Password  $telnetPass `
                    -Command   $telnetCmd

                if ($telnetResult.Success) {
                    $result = $true
                    # Log the full server response for debug - do not send raw output to Discord
                    _Log "[$($profile.GameName)] Telnet response: $($telnetResult.Response)" -Level DEBUG
                    if ((_TestProfileGame -Profile $profile -KnownGame '7DaysToDie') -and $telnetCmd -eq 'listplayers') {
                        $parsedPlayers = _Parse7DaysToDiePlayersResponse -ResponseText $telnetResult.Response
                        if (-not $parsedPlayers.Available) {
                            Start-Sleep -Milliseconds 700
                            $recentLogText = _ReadRecentProfileLogText -Profile $profile -Prefix $prefix -TailLines 200
                            $logParsedPlayers = _Parse7DaysToDiePlayersLogText -Text $recentLogText
                            if ($logParsedPlayers.Available) {
                                $parsedPlayers = $logParsedPlayers
                            }
                        }
                        if ($parsedPlayers.Available) {
                            _SetLatestPlayersSnapshot -Prefix $prefix -Names @($parsedPlayers.Names) -Count ([int]$parsedPlayers.Count) -SharedState $state
                            if ($state -and $state.ContainsKey('PlayerQueryState')) {
                                $state.PlayerQueryState[$prefix] = [ordered]@{
                                    LastAttemptAt = Get-Date
                                    LastSuccessAt = Get-Date
                                    Source        = '7DaysToDieTelnet'
                                    Note          = $parsedPlayers.Note
                                }
                            }
                        }
                    }
                    # Send a clean human-readable message to Discord matching the style of other commands
                    $resultMsg = switch ($telnetCmd) {
                        'saveworld'   { "[OK] $($profile.GameName) - world saved successfully."      }
                        'listplayers' {
                            if (($parsedPlayers) -and $parsedPlayers.Available) {
                                if ([int]$parsedPlayers.Count -le 0) {
                                    "[PLAYERS] $($profile.GameName) - no players online."
                                } elseif (@($parsedPlayers.Names).Count -gt 0) {
                                    "[PLAYERS] $($profile.GameName) - online: $((@($parsedPlayers.Names)) -join ', ')"
                                } else {
                                    "[PLAYERS] $($profile.GameName) - players online: $([int]$parsedPlayers.Count)"
                                }
                            } else {
                                "[OK] $($profile.GameName) - player query completed."
                            }
                        }
                        'version'     { "[OK] $($profile.GameName) - $($telnetResult.Response)"      }
                        default       { "[OK] $($profile.GameName) - command '$telnetCmd' completed." }
                    }
                } else {
                    $resultMsg = "[ERROR] $($profile.GameName) - Telnet command failed. Check the Program Log for details."
                    _Log "[$($profile.GameName)] Telnet error detail: $($telnetResult.Error)" -Level WARN
                }
            }
        }
        'SatisfactoryApi' {
            $sfHost  = if ($null -ne $profile.SatisfactoryApiHost) { $profile.SatisfactoryApiHost } else { '127.0.0.1' }
            $sfPort  = if ($null -ne $profile.SatisfactoryApiPort) { [int]$profile.SatisfactoryApiPort } else { 7777 }
            $sfToken = if ($null -ne $profile.SatisfactoryApiToken) { $profile.SatisfactoryApiToken } else { '' }
            $sfFunc  = if ($null -ne $cmdDef.Function) { $cmdDef.Function } else { '' }

            if (-not $sfFunc) {
                $resultMsg = "[ERROR] No Function specified for SatisfactoryApi command"
            } else {
                _Log "[$($profile.GameName)] SF API -> $sfFunc" -Level DEBUG

                # SaveGame requires a SaveName parameter
                $sfData = $null
                if ($sfFunc -eq 'SaveGame') {
                    $sfData = @{ SaveName = 'ecc_autosave' }
                }

                $r = Invoke-SatisfactoryApiRequest `
                    -Host     $sfHost `
                    -Port     $sfPort `
                    -Token    $sfToken `
                    -Function $sfFunc `
                    -Data     $sfData

                if ($r.Success) {
                    $result = $true
                    $resultMsg = switch ($sfFunc) {
                        'SaveGame' {
                            "[OK] $($profile.GameName) - world saved successfully."
                        }
                        'QueryServerState' {
                            $players = 0
                            $sessionName = 'Unknown'
                            try { $players     = [int]$r.Data.serverGameState.numConnectedPlayers } catch { }
                            try { $sessionName = $r.Data.serverGameState.activeSessionName        } catch { }
                            "[STATUS] $($profile.GameName) - Players: $players   Session: $sessionName"
                        }
                        default {
                            "[OK] $($profile.GameName) - '$sfFunc' completed."
                        }
                    }
                } else {
                    $hint = ''
                    if ([string]::IsNullOrWhiteSpace($sfToken)) {
                        $hint = " No API token set - add it to the SatisfactoryApiToken field in the Profile Editor."
                    }
                    $resultMsg = "[ERROR] $($profile.GameName) - API call '$sfFunc' failed.$hint"
                    _Log "[$($profile.GameName)] SF API error: $($r.Error)" -Level WARN
                }
            }
        }
        default {
            $resultMsg = "[ERROR] Unknown command type '$cmdType'. Valid: Start, Stop, Restart, Status, SendCommand, stdin, Rcon, Telnet, Rest, SatisfactoryApi, http, script"
        }
    }

    return @{ Message = $resultMsg; Success = $result }
}

Export-ModuleMember -Function `
    Initialize-ServerManager, Sync-RunningServersFromProcesses, Get-ServerStatus, `
    Start-GameServer, Stop-GameServer, Restart-GameServer, `
    Get-ServerRuntimeState, Set-ServerRuntimeState, Clear-ServerRuntimeState, `
    Restore-ServerRuntimeStateFromSharedState, `
    Set-JoinableServerRuntimeState, Set-ObservedPlayersServerRuntimeState, Set-LatestPlayersSnapshot, Get-ProfileHealthSnapshot, `
    New-DiscordGameMessage, New-DiscordSystemMessage, Send-DiscordGameEvent, `
    Invoke-SafeShutdown, Invoke-ProfileCommand, Invoke-ServerCommandText, `
    Send-ServerStdin, Invoke-ServerHttp, Invoke-CustomScript, `
    Invoke-RconCommand, Invoke-TelnetCommand, Start-ServerMonitor, `
    Test-TelnetConnection, Test-RconConnection, Test-SatisfactoryApiConnection, Test-StdInConnection, `
    Invoke-HytaleDownloaderCommand, Get-HytaleRequiredFilesStatus, Get-HytaleServerUpdateStatus, `
    Update-HytaleDownloader, Update-HytaleServerFiles
