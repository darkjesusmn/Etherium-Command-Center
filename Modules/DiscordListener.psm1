# =============================================================================
# DiscordListener.psm1  -  Discord REST polling + webhook delivery
# =============================================================================

$script:State      = $null
$script:ModuleRoot = $PSScriptRoot
$script:DefaultHttpTimeoutMs = 10000
 $script:MaxLoggedHttpBodyChars = 300

function Initialize-DiscordListener {
    param([hashtable]$SharedState)
    $script:State = $SharedState
}

function _GetDiscordHttpTimeoutMs {
    if ($script:State -and $script:State.Settings -and $script:State.Settings.ContainsKey('DiscordHttpTimeoutMs')) {
        try {
            $value = [int]$script:State.Settings.DiscordHttpTimeoutMs
            if ($value -gt 0) { return $value }
        } catch { }
    }
    return $script:DefaultHttpTimeoutMs
}

function _TrimHttpLogBody {
    param([string]$Body)

    if (-not $Body) { return '' }
    $clean = ($Body -replace '\s+', ' ').Trim()
    if ($clean.Length -gt $script:MaxLoggedHttpBodyChars) {
        return $clean.Substring(0, $script:MaxLoggedHttpBodyChars) + '...'
    }
    return $clean
}

function _GetHttpErrorDetails {
    param([Parameter(Mandatory = $true)]$ErrorRecord)

    $details = [ordered]@{
        StatusCode        = ''
        StatusDescription = ''
        Body              = ''
        ErrorType         = ''
        IsTimeout         = $false
    }

    try {
        if ($ErrorRecord.Exception) {
            $details.ErrorType = $ErrorRecord.Exception.GetType().FullName
        }
    } catch { }

    try {
        $status = [System.Net.WebExceptionStatus]::UnknownError
        if ($ErrorRecord.Exception -and $ErrorRecord.Exception.Status) {
            $status = $ErrorRecord.Exception.Status
        }
        $details.IsTimeout = ($status -eq [System.Net.WebExceptionStatus]::Timeout)
    } catch { }

    try {
        $resp = $ErrorRecord.Exception.Response
        if ($resp) {
            try { $details.StatusCode = [int]$resp.StatusCode } catch { $details.StatusCode = '' }
            try { $details.StatusDescription = $resp.StatusDescription } catch { $details.StatusDescription = '' }
            try {
                $stream = $resp.GetResponseStream()
                if ($stream) {
                    $reader = [System.IO.StreamReader]::new($stream)
                    $details.Body = _TrimHttpLogBody -Body ($reader.ReadToEnd())
                    $reader.Close()
                }
            } catch { $details.Body = '' }
        }
    } catch { }

    return $details
}

# -----------------------------------------------------------------------------
#  Low-level HTTP helper
# -----------------------------------------------------------------------------
function _DiscordRequest {
    param(
        [string]$Path,
        [string]$Method = 'GET',
        [string]$Body   = '',
        [string]$Token  = '',
        [int]$TimeoutMs = 0
    )

    $url = "https://discord.com/api/v10$Path"
    $effectiveTimeoutMs = if ($TimeoutMs -gt 0) { $TimeoutMs } else { $script:DefaultHttpTimeoutMs }
    $startedAt = Get-Date

    try {
        $req = [System.Net.HttpWebRequest]::Create($url)
        $req.Method      = $Method
        $req.ContentType = 'application/json'
        $req.UserAgent   = 'DiscordBot (Etherium Command Center, 1.0)'
        $req.Timeout = $effectiveTimeoutMs
        $req.ReadWriteTimeout = $effectiveTimeoutMs

        if ($Token) {
            $req.Headers['Authorization'] = "Bot $Token"
        }

        if ($Body -and $Method -ne 'GET') {
            $bytes = [System.Text.Encoding]::UTF8.GetBytes($Body)
            $req.ContentLength = $bytes.Length
            $s = $req.GetRequestStream()
            $s.Write($bytes, 0, $bytes.Length)
            $s.Close()
        }

        $resp   = $req.GetResponse()
        $reader = [System.IO.StreamReader]::new($resp.GetResponseStream())
        $text   = $reader.ReadToEnd()
        $statusCode = ''
        $statusDesc = ''
        try { $statusCode = [int]$resp.StatusCode } catch { $statusCode = '' }
        try { $statusDesc = $resp.StatusDescription } catch { $statusDesc = '' }
        $reader.Close()
        $resp.Close()
        $elapsedMs = [int][Math]::Round(((Get-Date) - $startedAt).TotalMilliseconds)
        _DLog "Discord API OK [$Method $Path] ${elapsedMs}ms HTTP $statusCode $statusDesc" -Level DEBUG
        return $text
    }
    catch {
        $err = $_
        $elapsedMs = [int][Math]::Round(((Get-Date) - $startedAt).TotalMilliseconds)
        $details = _GetHttpErrorDetails -ErrorRecord $err
        $timeoutLabel = if ($details.IsTimeout) { ' timeout' } else { '' }
        if ($details.StatusCode -or $details.StatusDescription -or $details.Body) {
            _DLog "Discord API${timeoutLabel} error [$Method $url] after ${elapsedMs}ms (timeout=${effectiveTimeoutMs}ms) type=$($details.ErrorType) HTTP $($details.StatusCode) $($details.StatusDescription) Body: $($details.Body)" -Level ERROR
        } else {
            _DLog "Discord API${timeoutLabel} error [$Method $url] after ${elapsedMs}ms (timeout=${effectiveTimeoutMs}ms) type=$($details.ErrorType): $err" -Level ERROR
        }
        return $null
    }
}

# -----------------------------------------------------------------------------
#  Send a message as the bot to a channel
# -----------------------------------------------------------------------------
function _SendBotMessage {
    param(
        [string]$Token,
        [string]$ChannelId,
        [string]$Content,
        [int]$TimeoutMs = 0
    )

    if (-not $Token -or -not $ChannelId -or -not $Content) { return $false }

    try {
        _DLog "Sending bot message -> channel=$ChannelId len=$($Content.Length)" -Level DEBUG
        $body = @{ content = $Content } | ConvertTo-Json -Compress
        $resp = _DiscordRequest -Path "/channels/$ChannelId/messages" -Method 'POST' -Body $body -Token $Token -TimeoutMs $TimeoutMs
        return ($null -ne $resp)
    } catch {
        _DLog "Bot message send failed: $_" -Level WARN
        return $false
    }
}

# -----------------------------------------------------------------------------
#  Logging helper
# -----------------------------------------------------------------------------
function _DLog {
    param([string]$Msg, [string]$Level = 'INFO')

    if ($Level -eq 'DEBUG') {
        $dbg = $false
        if ($script:State -and $script:State.Settings -and $script:State.Settings.ContainsKey('EnableDebugLogging')) {
            $dbg = [bool]$script:State.Settings.EnableDebugLogging
        }
        if (-not $dbg) { return }
    }

    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry = "[$ts][$Level][Discord] $Msg"

    if ($script:State -and $script:State.ContainsKey('LogQueue')) {
        $script:State.LogQueue.Enqueue($entry)
    }

    $col = switch ($Level) {
        'ERROR' { 'Red' }
        'WARN'  { 'Yellow' }
        default { 'DarkCyan' }
    }

    try { Write-Host $entry -ForegroundColor $col } catch {}
}

function _TraceDiscordDeliveryDecision {
    param(
        [string]$Action,
        [string]$Method = '',
        [string]$Reason = '',
        [string]$Content = ''
    )

    $preview = '""'
    if (-not [string]::IsNullOrWhiteSpace($Content)) {
        $safePreview = $Content -replace "(`r`n|`n|`r)", ' '
        $safePreview = $safePreview -replace '\s+', ' '
        $safePreview = $safePreview.Trim()
        if ($safePreview.Length -gt 140) {
            $safePreview = $safePreview.Substring(0, 140) + '...'
        }
        $safePreview = '"' + $safePreview.Replace('"', "'") + '"'
        $preview = $safePreview
    }

    $safeReason = '""'
    if (-not [string]::IsNullOrWhiteSpace($Reason)) {
        $safeReason = '"' + $Reason.Trim().Replace('"', "'") + '"'
    }

    $message =
        'DISCORD delivery ' +
        'action=' + $(if ([string]::IsNullOrWhiteSpace($Action)) { 'observe' } else { $Action.Trim() }) + ' ' +
        'method=' + $(if ([string]::IsNullOrWhiteSpace($Method)) { '<none>' } else { $Method.Trim() }) + ' ' +
        'preview=' + $preview + ' ' +
        'reason=' + $safeReason

    _DLog $message -Level DEBUG
}

# -----------------------------------------------------------------------------
#  Webhook sender
# -----------------------------------------------------------------------------
function Send-WebhookMessage {
    param(
        [string]$WebhookUrl,
        [string]$Content,
        [int]$TimeoutMs = 0
    )

    if (-not $WebhookUrl) { return }
    $effectiveTimeoutMs = if ($TimeoutMs -gt 0) { $TimeoutMs } else { $script:DefaultHttpTimeoutMs }
    $startedAt = Get-Date

    try {
        $payload = @{ content = $Content } | ConvertTo-Json -Compress
        $req = [System.Net.HttpWebRequest]::Create($WebhookUrl)
        $req.Method      = 'POST'
        $req.ContentType = 'application/json'
        $req.UserAgent   = 'DiscordBot (Etherium Command Center, 1.0)'
        $req.Timeout = $effectiveTimeoutMs
        $req.ReadWriteTimeout = $effectiveTimeoutMs

        $bytes = [System.Text.Encoding]::UTF8.GetBytes($payload)
        $req.ContentLength = $bytes.Length

        $s = $req.GetRequestStream()
        $s.Write($bytes, 0, $bytes.Length)
        $s.Close()

        $resp = $req.GetResponse()
        $statusCode = ''
        $statusDesc = ''
        try { $statusCode = [int]$resp.StatusCode } catch { $statusCode = '' }
        try { $statusDesc = $resp.StatusDescription } catch { $statusDesc = '' }
        $resp.Close()
        $elapsedMs = [int][Math]::Round(((Get-Date) - $startedAt).TotalMilliseconds)
        _DLog "Webhook delivered ${elapsedMs}ms HTTP $statusCode $statusDesc len=$($Content.Length)" -Level DEBUG
    }
    catch {
        $details = _GetHttpErrorDetails -ErrorRecord $_
        $elapsedMs = [int][Math]::Round(((Get-Date) - $startedAt).TotalMilliseconds)
        $timeoutLabel = if ($details.IsTimeout) { ' timeout' } else { '' }
        if ($details.StatusCode -or $details.StatusDescription -or $details.Body) {
            _DLog "Webhook${timeoutLabel} delivery failed after ${elapsedMs}ms (timeout=${effectiveTimeoutMs}ms) type=$($details.ErrorType) HTTP $($details.StatusCode) $($details.StatusDescription) Body: $($details.Body)" -Level WARN
        } else {
            _DLog "Webhook${timeoutLabel} delivery failed after ${elapsedMs}ms (timeout=${effectiveTimeoutMs}ms) type=$($details.ErrorType): $_" -Level WARN
        }
    }
}

function _SendOutboundDiscordMessage {
    param(
        [string]$Token,
        [string]$ChannelId,
        [string]$WebhookUrl,
        [string]$Content,
        [int]$TimeoutMs = 0
    )

    if ([string]::IsNullOrWhiteSpace($Content)) {
        _TraceDiscordDeliveryDecision -Action 'skip' -Reason 'blank outbound Discord message was ignored' -Content $Content
        return [ordered]@{ Success = $false; Method = '' }
    }

    if (-not [string]::IsNullOrWhiteSpace($Token) -and -not [string]::IsNullOrWhiteSpace($ChannelId)) {
        _TraceDiscordDeliveryDecision -Action 'attempt' -Method 'bot' -Reason 'attempting Discord bot-channel delivery first' -Content $Content
        $sentViaBot = $false
        try {
            $sentViaBot = _SendBotMessage -Token $Token -ChannelId $ChannelId -Content $Content -TimeoutMs $TimeoutMs
        } catch {
            $sentViaBot = $false
        }

        if ($sentViaBot) {
            _TraceDiscordDeliveryDecision -Action 'deliver' -Method 'bot' -Reason 'Discord message delivered through the bot channel' -Content $Content
            return [ordered]@{ Success = $true; Method = 'bot' }
        }

        _DLog "Bot-channel delivery failed for outbound message. Falling back to webhook." -Level WARN
        _TraceDiscordDeliveryDecision -Action 'fallback' -Method 'webhook' -Reason 'bot-channel delivery failed; falling back to webhook delivery' -Content $Content
    }

    if (-not [string]::IsNullOrWhiteSpace($WebhookUrl)) {
        _TraceDiscordDeliveryDecision -Action 'deliver' -Method 'webhook' -Reason 'sending Discord message through webhook delivery' -Content $Content
        Send-WebhookMessage -WebhookUrl $WebhookUrl -Content $Content -TimeoutMs $TimeoutMs
        return [ordered]@{ Success = $true; Method = 'webhook' }
    }

    _TraceDiscordDeliveryDecision -Action 'drop' -Reason 'no Discord bot channel or webhook destination was configured' -Content $Content
    return [ordered]@{ Success = $false; Method = '' }
}

# -----------------------------------------------------------------------------
#  Flush webhook queue
# -----------------------------------------------------------------------------
function _FlushWebhooks {
    param(
        [string]$Token = '',
        [string]$ChannelId = '',
        [string]$WebhookUrl,
        [int]$TimeoutMs = 0
    )

    if (-not $script:State.ContainsKey('WebhookQueue')) { return }

    $item = $null
    while ($script:State.WebhookQueue.TryDequeue([ref]$item)) {
        _TraceDiscordDeliveryDecision -Action 'dequeue' -Method 'queue' -Reason 'pulled a Discord message from the outbound queue for delivery' -Content $item
        $delivery = _SendOutboundDiscordMessage -Token $Token -ChannelId $ChannelId -WebhookUrl $WebhookUrl -Content $item -TimeoutMs $TimeoutMs
        if ($delivery.Method -eq 'bot') {
            Start-Sleep -Milliseconds 1100
        } else {
            Start-Sleep -Milliseconds 300
        }
    }
}

# -----------------------------------------------------------------------------
#  Parse commands like: !PZstart or !PZ start
# -----------------------------------------------------------------------------
function _ParseCommand {
    param([string]$Content, [string]$BotPrefix)

    # Case-insensitive prefix check
    if (-not $Content.ToUpper().StartsWith($BotPrefix.ToUpper())) { return $null }

    $rest = $Content.Substring($BotPrefix.Length).Trim()

    # Get list of available game prefixes
    if ($null -eq $script:State -or $null -eq $script:State.Profiles) {
        return $null
    }

    $availablePrefixes = @($script:State.Profiles.Keys)

    # Try to match against known prefixes (longest match first) - all comparisons case-insensitive
    foreach ($knownPrefix in ($availablePrefixes | Sort-Object -Property Length -Descending)) {

        # Case 1: "!PZ restart" or "!pz RESTART" (with space between prefix and command)
        # Use (?i) for case-insensitive regex
        if ($rest -match "(?i)^$([regex]::Escape($knownPrefix))\s+([A-Za-z]+)$") {
            return @{
                GamePrefix = $knownPrefix      # always stored uppercase in Profiles dict
                Command    = $Matches[1].ToLower()
            }
        }

        # Case 2: "!PZrestart" or "!pzRESTART" (no space - prefix glued to command)
        if ($rest.ToUpper().StartsWith($knownPrefix.ToUpper())) {
            $afterPrefix = $rest.Substring($knownPrefix.Length).Trim()
            if ($afterPrefix -match '^([A-Za-z]+)$') {
                return @{
                    GamePrefix = $knownPrefix
                    Command    = $afterPrefix.ToLower()
                }
            }
        }
    }

    return $null
}

# -----------------------------------------------------------------------------
#  Process a single Discord message
# -----------------------------------------------------------------------------
# -----------------------------------------------------------------------------
#  Command cooldown table
#  Key: "GAMEPREFIX_commandname"  Value: DateTime of last execution
# -----------------------------------------------------------------------------
$script:CooldownTable = [hashtable]::Synchronized(@{})

$script:CooldownSeconds = @{
    'start'   = 60
    'stop'    = 60
    'restart' = 60
    'save'    = 30
    'status'  = 5
}
$script:CooldownDefault = 10   # fallback for any custom command

function _GetCooldown {
    param([string]$CmdName)
    if ($script:CooldownSeconds.ContainsKey($CmdName)) {
        return $script:CooldownSeconds[$CmdName]
    }
    return $script:CooldownDefault
}

function _CheckCooldown {
    param([string]$GamePrefix, [string]$CmdName)
    $key      = "${GamePrefix}_${CmdName}"
    $cooldown = _GetCooldown -CmdName $CmdName
    $now      = Get-Date

    if ($script:CooldownTable.ContainsKey($key)) {
        $elapsed = ($now - $script:CooldownTable[$key]).TotalSeconds
        if ($elapsed -lt $cooldown) {
            $remaining = [Math]::Ceiling($cooldown - $elapsed)
            return @{ OnCooldown = $true; Remaining = $remaining; Cooldown = $cooldown }
        }
    }
    return @{ OnCooldown = $false; Remaining = 0; Cooldown = $cooldown }
}

function _SetCooldown {
    param([string]$GamePrefix, [string]$CmdName)
    $key = "${GamePrefix}_${CmdName}"
    $script:CooldownTable[$key] = Get-Date
}

function _GetRequesterLabel {
    param([object]$Msg)

    try {
        if ($Msg) {
            $displayName = ''
            try {
                if ($Msg.member -and $Msg.member.nick) {
                    $displayName = "$($Msg.member.nick)".Trim()
                }
            } catch { $displayName = '' }

            if ([string]::IsNullOrWhiteSpace($displayName)) {
                try {
                    if ($Msg.author -and $Msg.author.global_name) {
                        $displayName = "$($Msg.author.global_name)".Trim()
                    }
                } catch { $displayName = '' }
            }

            if ([string]::IsNullOrWhiteSpace($displayName)) {
                try {
                    if ($Msg.author -and $Msg.author.username) {
                        $displayName = "$($Msg.author.username)".Trim()
                    }
                } catch { $displayName = '' }
            }

            if (-not [string]::IsNullOrWhiteSpace($displayName)) {
                return "@$displayName"
            }
        }
    } catch { }

    return 'there'
}

function _SetDiscordCommandContext {
    param(
        [string]$GamePrefix,
        [string]$CommandName
    )

    if (-not $script:State -or [string]::IsNullOrWhiteSpace($GamePrefix) -or [string]::IsNullOrWhiteSpace($CommandName)) {
        return
    }

    if (-not $script:State.ContainsKey('DiscordCommandContext') -or $null -eq $script:State.DiscordCommandContext) {
        $script:State['DiscordCommandContext'] = [hashtable]::Synchronized(@{})
    }

    $key = $GamePrefix.ToUpperInvariant()
    $ttlSeconds = switch ($CommandName.ToLowerInvariant()) {
        'start'   { 180 }
        'stop'    { 180 }
        'restart' { 240 }
        default   { 60 }
    }

    $script:State.DiscordCommandContext[$key] = [ordered]@{
        Command   = $CommandName.ToLowerInvariant()
        ExpiresAt = (Get-Date).AddSeconds($ttlSeconds)
        Source    = 'Discord'
    }
}

function _ProcessMessage {
    param([object]$Msg, [string]$BotUserId)

    if ($Msg.author.bot -or $Msg.author.id -eq $BotUserId) { 
        $user = if ($Msg.author.username) { $Msg.author.username } else { 'Bot' }
        $content = if ($Msg.content) { $Msg.content } else { '' }
        _DLog "Bot message from ${user}: $content"
        return 
    }

    $settings = $script:State.Settings
    $botPrefix = if ($settings.CommandPrefix) { $settings.CommandPrefix } else { '!' }
    $token = if ($settings.BotToken) { $settings.BotToken } else { '' }
    $channelId = if ($settings.MonitorChannelId) { $settings.MonitorChannelId } else { '' }

    _DLog "Message received: '$($Msg.content)' from $($Msg.author.username), looking for prefix: '$botPrefix'"

    $parsed = _ParseCommand -Content $Msg.content -BotPrefix $botPrefix
    if (-not $parsed) { 
        _DLog "Command parse failed for: '$($Msg.content)'"
        return 
    }

    $gamePrefix = $parsed.GamePrefix
    $cmdName    = $parsed.Command

    _DLog "Command parsed! GamePrefix=$gamePrefix, Command=$cmdName"

    # Safety enforcement
    if ($cmdName -match '^(kick|ban|tp|teleport|op|deop|whitelist|give|take)$') {
        _DLog "Blocked unsafe command: $cmdName" -Level WARN
        $script:State.WebhookQueue.Enqueue("[BLOCKED] Unsafe command '$cmdName' is not allowed.")
        return
    }

    if (-not $script:State.Profiles.ContainsKey($gamePrefix)) {
        $availablePrefixes = @($script:State.Profiles.Keys) -join ', '
        _DLog "Unknown prefix '$gamePrefix'. Available: $availablePrefixes" -Level WARN
        $script:State.WebhookQueue.Enqueue("[UNKNOWN] No game found with prefix '$gamePrefix'. Available: $availablePrefixes")
        return
    }

    _DLog "Profile found for $gamePrefix, invoking command: $cmdName"
    $requesterLabel = _GetRequesterLabel -Msg $Msg

    # Cooldown check
    $cooldownResult = _CheckCooldown -GamePrefix $gamePrefix -CmdName $cmdName
    if ($cooldownResult.OnCooldown) {
        $msg = "[COOLDOWN] !${gamePrefix} ${cmdName} is on cooldown. Please wait $($cooldownResult.Remaining)s before using it again."
        _DLog $msg -Level WARN
        $script:State.WebhookQueue.Enqueue($msg)
        return
    }

    # Mark a pending "players" request so the log tailer can respond with names
    if ($cmdName -eq 'players') {
        if (-not $script:State.ContainsKey('PlayersRequests')) {
            $script:State['PlayersRequests'] = [hashtable]::Synchronized(@{})
        }
        $script:State.PlayersRequests[$gamePrefix] = @{
            Source = 'Discord'
            RequestedAt = Get-Date
        }
    }

    # ── Immediate acknowledgement ─────────────────────────────────────────────
    # Send this BEFORE the command runs so users aren't staring at silence
    $profile = $script:State.Profiles[$gamePrefix]
    $gameName = $profile.GameName
    $ackMsg = switch ($cmdName) {
        'start'   { New-DiscordGameMessage -Profile $profile -Prefix $gamePrefix -Event 'received_start'   -Values @{ Requester = $requesterLabel } }
        'stop'    { New-DiscordGameMessage -Profile $profile -Prefix $gamePrefix -Event 'received_stop'    -Values @{ Requester = $requesterLabel } }
        'restart' { New-DiscordGameMessage -Profile $profile -Prefix $gamePrefix -Event 'received_restart' -Values @{ Requester = $requesterLabel } }
        'save'    { New-DiscordGameMessage -Profile $profile -Prefix $gamePrefix -Event 'received_save'    -Values @{ Requester = $requesterLabel } }
        'status'  { New-DiscordGameMessage -Profile $profile -Prefix $gamePrefix -Event 'received_status'  -Values @{ Requester = $requesterLabel } }
        default   { New-DiscordGameMessage -Profile $profile -Prefix $gamePrefix -Event 'received_command' -Values @{ Requester = $requesterLabel; Command = "'$cmdName'" } }
    }
    $script:State.WebhookQueue.Enqueue($ackMsg)
    # Flush immediately so it hits Discord right now, not on the next poll tick
    $webhookUrl = $script:State.Settings.WebhookUrl
    $httpTimeoutMs = _GetDiscordHttpTimeoutMs
    if ($webhookUrl) {
        _FlushWebhooks -Token $token -ChannelId $channelId -WebhookUrl $webhookUrl -TimeoutMs $httpTimeoutMs
    }

    try {
        if ($cmdName -in @('start', 'stop', 'restart')) {
            _SetDiscordCommandContext -GamePrefix $gamePrefix -CommandName $cmdName
        }
        # Pass SharedState to Invoke-ProfileCommand
        $result = Invoke-ProfileCommand -Prefix $gamePrefix -CommandName $cmdName -SharedState $script:State
        _DLog "Command executed: $($result.Message)"
        $sendResultMessage = $true
        if ($cmdName -in @('start', 'stop', 'restart')) {
            $sendResultMessage = ($result.Success -ne $true)
        }
        if ($sendResultMessage) {
            $script:State.WebhookQueue.Enqueue($result.Message)
        }
        # Flush immediately so the result appears without waiting for next poll
        if ($webhookUrl) {
            _FlushWebhooks -Token $token -ChannelId $channelId -WebhookUrl $webhookUrl -TimeoutMs $httpTimeoutMs
        }
        if ($null -ne $result.Success -and [bool]$result.Success) {
            _SetCooldown -GamePrefix $gamePrefix -CmdName $cmdName
        }
    }
    catch {
        _DLog "Routing error: $_" -Level ERROR
        $script:State.WebhookQueue.Enqueue("[ERROR] Failed to process command: $_")
        if ($webhookUrl) {
            _FlushWebhooks -Token $token -ChannelId $channelId -WebhookUrl $webhookUrl -TimeoutMs $httpTimeoutMs
        }
    }
}

# -----------------------------------------------------------------------------
#  Fetch new messages
# -----------------------------------------------------------------------------
function _GetNewMessages {
    param(
        [string]$Token,
        [string]$ChannelId,
        [string]$AfterMessageId,
        [int]$TimeoutMs = 0
    )

    $path = "/channels/$ChannelId/messages?limit=10"
    if ($AfterMessageId) { $path += "&after=$AfterMessageId" }

    $json = _DiscordRequest -Path $path -Token $Token -TimeoutMs $TimeoutMs
    if (-not $json) { return @() }

    try {
        $msgs = $json | ConvertFrom-Json
        if ($msgs -is [array]) { return $msgs }
        return @($msgs)
    }
    catch { return @() }
}

# -----------------------------------------------------------------------------
#  Get bot user ID
# -----------------------------------------------------------------------------
function _GetBotUserId {
    param(
        [string]$Token,
        [int]$TimeoutMs = 0
    )

    $json = _DiscordRequest -Path '/users/@me' -Token $Token -TimeoutMs $TimeoutMs
    if (-not $json) { return '' }

    try {
        return ($json | ConvertFrom-Json).id
    }
    catch { return '' }
}

# -----------------------------------------------------------------------------
#  MAIN LISTENER LOOP
# -----------------------------------------------------------------------------
function Start-DiscordListener {
    param([hashtable]$SharedState)

    Initialize-DiscordListener -SharedState $SharedState

    Import-Module (Join-Path $script:ModuleRoot 'Logging.psm1')         -Force
    Import-Module (Join-Path $script:ModuleRoot 'ProfileManager.psm1') -Force
    Import-Module (Join-Path $script:ModuleRoot 'ServerManager.psm1')  -Force
    Initialize-ServerManager -SharedState $SharedState

    $settings  = $SharedState.Settings
    $token     = $settings.BotToken
    $channelId = $settings.MonitorChannelId
    $webhook   = $settings.WebhookUrl
    $interval  = if ($settings.PollIntervalSeconds) { [int]$settings.PollIntervalSeconds } else { 2 }
    $httpTimeoutMs = _GetDiscordHttpTimeoutMs

    if (-not $token)     { _DLog 'No BotToken configured.' -Level WARN; return }
    if (-not $channelId) { _DLog 'No MonitorChannelId configured.' -Level WARN; return }

    _DLog "Listener starting (channel=$channelId, interval=${interval}s, timeout=${httpTimeoutMs}ms)"

    $botUserId = _GetBotUserId -Token $token -TimeoutMs $httpTimeoutMs
    _DLog "Bot user ID: $botUserId"

    # Seed last message ID
    $seed = _GetNewMessages -Token $token -ChannelId $channelId -AfterMessageId '' -TimeoutMs $httpTimeoutMs
    $lastId = if ($seed.Count -gt 0) {
        ($seed | Sort-Object { [decimal]$_.id } | Select-Object -Last 1).id
    } else { '' }

    try { $SharedState['DiscordOfflineNoticeSent'] = $false } catch { }
    $onlineDelivery = _SendOutboundDiscordMessage -Token $token -ChannelId $channelId -WebhookUrl $webhook -Content (New-DiscordSystemMessage -Event 'online') -TimeoutMs $httpTimeoutMs
    if (-not $onlineDelivery.Success) {
        _DLog 'Failed to send online notification.' -Level WARN
    }
    $SharedState['ListenerRunning'] = $true

    while (-not ($SharedState.ContainsKey('StopListener') -and $SharedState['StopListener'] -eq $true)) {
        Start-Sleep -Seconds $interval

        # Send GUI-originated messages queued by the ECC UI.
        if ($SharedState.ContainsKey('DiscordOutbox') -and $SharedState.DiscordOutbox) {
            $out = $null
            while ($SharedState.DiscordOutbox.TryDequeue([ref]$out)) {
                if (-not [string]::IsNullOrWhiteSpace($out)) {
                    $ok = _SendBotMessage -Token $token -ChannelId $channelId -Content $out -TimeoutMs $httpTimeoutMs
                    if (-not $ok) { _DLog "Failed to send GUI message." -Level WARN }
                    # Basic rate limiting: Discord allows ~5 messages / 5 seconds per channel
                    Start-Sleep -Milliseconds 1100
                }
            }
        }

        _FlushWebhooks -Token $token -ChannelId $channelId -WebhookUrl $webhook -TimeoutMs $httpTimeoutMs

        try {
            $msgs = _GetNewMessages -Token $token -ChannelId $channelId -AfterMessageId $lastId -TimeoutMs $httpTimeoutMs
            if ($msgs.Count -gt 0) {
                $sorted = $msgs | Sort-Object { [decimal]$_.id }
                foreach ($msg in $sorted) {
                    _ProcessMessage -Msg $msg -BotUserId $botUserId
                }
                $lastId = ($sorted | Select-Object -Last 1).id
            }
        }
        catch {
            _DLog "Poll error: $_" -Level ERROR
        }
    }

    # Flush any remaining queued webhooks before going offline
    _FlushWebhooks -Token $token -ChannelId $channelId -WebhookUrl $webhook -TimeoutMs $httpTimeoutMs

    $SharedState['ListenerRunning'] = $false
    # Send offline notification
    $offlineAlreadySent = $false
    try {
        if ($SharedState.ContainsKey('DiscordOfflineNoticeSent')) {
            $offlineAlreadySent = [bool]$SharedState['DiscordOfflineNoticeSent']
        }
    } catch { $offlineAlreadySent = $false }

    if (-not $offlineAlreadySent) {
        try { $SharedState['DiscordOfflineNoticeSent'] = $true } catch { }
        $offlineDelivery = _SendOutboundDiscordMessage -Token $token -ChannelId $channelId -WebhookUrl $webhook -Content (New-DiscordSystemMessage -Event 'offline') -TimeoutMs $httpTimeoutMs
        if ($offlineDelivery.Success) {
            _DLog ("Offline message sent via {0}." -f $offlineDelivery.Method)
        } else {
            _DLog 'Failed to send offline message.' -Level WARN
        }
    } else {
        _DLog 'Offline message skipped because it was already sent.' -Level DEBUG
    }
    _DLog 'Listener stopped.'
}

Export-ModuleMember -Function Start-DiscordListener, Send-WebhookMessage, Initialize-DiscordListener
