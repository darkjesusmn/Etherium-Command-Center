# =============================================================================
# DiscordListener.psm1  -  Discord REST polling + webhook delivery
# =============================================================================

$script:State      = $null
$script:ModuleRoot = $PSScriptRoot

function Initialize-DiscordListener {
    param([hashtable]$SharedState)
    $script:State = $SharedState
}

# -----------------------------------------------------------------------------
#  Low-level HTTP helper
# -----------------------------------------------------------------------------
function _DiscordRequest {
    param(
        [string]$Path,
        [string]$Method = 'GET',
        [string]$Body   = '',
        [string]$Token  = ''
    )

    $url = "https://discord.com/api/v10$Path"

    try {
        $req = [System.Net.HttpWebRequest]::Create($url)
        $req.Method      = $Method
        $req.ContentType = 'application/json'
        $req.UserAgent   = 'DiscordBot (Etherium Command Center, 1.0)'

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
        $reader.Close()
        $resp.Close()
        return $text
    }
    catch {
        $err = $_
        $status = ''
        $desc = ''
        $body = ''
        try {
            $resp = $err.Exception.Response
            if ($resp) {
                try { $status = [int]$resp.StatusCode } catch { $status = '' }
                try { $desc = $resp.StatusDescription } catch { $desc = '' }
                try {
                    $reader = New-Object System.IO.StreamReader($resp.GetResponseStream())
                    $body = $reader.ReadToEnd()
                    $reader.Close()
                } catch { $body = '' }
            }
        } catch { }

        if ($body -and $body.Length -gt 300) { $body = $body.Substring(0,300) + '...' }
        if ($status -or $desc -or $body) {
            _DLog "Discord API error [$Method $url] HTTP $status $desc Body: $body" -Level ERROR
        } else {
            _DLog "Discord API error [$Method $url]: $err" -Level ERROR
        }
        return $null
    }
}

# -----------------------------------------------------------------------------
#  Send a message as the bot to a channel
# -----------------------------------------------------------------------------
function _SendBotMessage {
    param([string]$Token, [string]$ChannelId, [string]$Content)

    if (-not $Token -or -not $ChannelId -or -not $Content) { return $false }

    try {
        _DLog "Sending bot message -> channel=$ChannelId len=$($Content.Length)" -Level DEBUG
        $body = @{ content = $Content } | ConvertTo-Json -Compress
        $resp = _DiscordRequest -Path "/channels/$ChannelId/messages" -Method 'POST' -Body $body -Token $Token
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

    if ($script:State -and $script:State.ContainsKey('LogQueue')) {
        $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $script:State.LogQueue.Enqueue("[$ts][$Level][Discord] $Msg")
    }

    $col = switch ($Level) {
        'ERROR' { 'Red' }
        'WARN'  { 'Yellow' }
        default { 'DarkCyan' }
    }

    try { Write-Host "[Discord][$Level] $Msg" -ForegroundColor $col } catch {}
}

# -----------------------------------------------------------------------------
#  Webhook sender
# -----------------------------------------------------------------------------
function Send-WebhookMessage {
    param([string]$WebhookUrl, [string]$Content)

    if (-not $WebhookUrl) { return }

    try {
        $payload = @{ content = $Content } | ConvertTo-Json -Compress
        $req = [System.Net.HttpWebRequest]::Create($WebhookUrl)
        $req.Method      = 'POST'
        $req.ContentType = 'application/json'
        $req.UserAgent   = 'DiscordBot (Etherium Command Center, 1.0)'

        $bytes = [System.Text.Encoding]::UTF8.GetBytes($payload)
        $req.ContentLength = $bytes.Length

        $s = $req.GetRequestStream()
        $s.Write($bytes, 0, $bytes.Length)
        $s.Close()

        $resp = $req.GetResponse()
        $resp.Close()
    }
    catch {
        _DLog "Webhook delivery failed: $_" -Level WARN
    }
}

# -----------------------------------------------------------------------------
#  Flush webhook queue
# -----------------------------------------------------------------------------
function _FlushWebhooks {
    param([string]$WebhookUrl)

    if (-not $script:State.ContainsKey('WebhookQueue')) { return }

    $item = $null
    while ($script:State.WebhookQueue.TryDequeue([ref]$item)) {
        Send-WebhookMessage -WebhookUrl $WebhookUrl -Content $item
        Start-Sleep -Milliseconds 300
    }
}

# -----------------------------------------------------------------------------
#  Parse commands like: !PZstart or !PZ start
# -----------------------------------------------------------------------------
function _ParseCommand {
    param([string]$Content, [string]$BotPrefix)

    if (-not $Content.StartsWith($BotPrefix)) { return $null }

    $rest = $Content.Substring($BotPrefix.Length).Trim()

    # Get list of available game prefixes
    if ($null -eq $script:State -or $null -eq $script:State.Profiles) {
        return $null
    }

    $availablePrefixes = @($script:State.Profiles.Keys)
    
    # Try to match against known prefixes first (longest match first)
    foreach ($knownPrefix in ($availablePrefixes | Sort-Object -Property Length -Descending)) {
        # Case 1: "!PZDS start" (with space) - PREFERRED
        if ($rest -match "^$knownPrefix\s+([A-Za-z]+)$") {
            return @{
                GamePrefix = $knownPrefix
                Command    = $Matches[1].ToLower()
            }
        }
        
        # Case 2: "!PZDSstart" (without space) - fallback
        if ($rest.ToUpper().StartsWith($knownPrefix)) {
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

    # Cooldown check
    $cooldownResult = _CheckCooldown -GamePrefix $gamePrefix -CmdName $cmdName
    if ($cooldownResult.OnCooldown) {
        $msg = "[COOLDOWN] !${gamePrefix} ${cmdName} is on cooldown. Please wait $($cooldownResult.Remaining)s before using it again."
        _DLog $msg -Level WARN
        $script:State.WebhookQueue.Enqueue($msg)
        return
    }

    # Record the cooldown timestamp BEFORE executing so rapid-fire duplicates are blocked
    _SetCooldown -GamePrefix $gamePrefix -CmdName $cmdName

    # Mark a pending "players" request so the log tailer can respond with names
    if ($cmdName -eq 'players') {
        if (-not $script:State.ContainsKey('PlayersRequests')) {
            $script:State['PlayersRequests'] = [hashtable]::Synchronized(@{})
        }
        $script:State.PlayersRequests[$gamePrefix] = Get-Date
    }

    # ── Immediate acknowledgement ─────────────────────────────────────────────
    # Send this BEFORE the command runs so users aren't staring at silence
    $gameName = $script:State.Profiles[$gamePrefix].GameName
    $ackMsg = switch ($cmdName) {
        'start'   { "[RECEIVED] $gameName - start command received. Starting server..." }
        'stop'    { "[RECEIVED] $gameName - stop command received. Saving and shutting down..." }
        'restart' { "[RECEIVED] $gameName - restart command received. Saving and restarting..." }
        'save'    { "[RECEIVED] $gameName - save command received. Saving now..." }
        'status'  { "[RECEIVED] $gameName - checking status..." }
        default   { "[RECEIVED] $gameName - running command: $cmdName..." }
    }
    $script:State.WebhookQueue.Enqueue($ackMsg)
    # Flush immediately so it hits Discord right now, not on the next poll tick
    $webhookUrl = $script:State.Settings.WebhookUrl
    if ($webhookUrl) {
        _FlushWebhooks -WebhookUrl $webhookUrl
    }

    try {
        # Pass SharedState to Invoke-ProfileCommand
        $result = Invoke-ProfileCommand -Prefix $gamePrefix -CommandName $cmdName -SharedState $script:State
        _DLog "Command executed: $($result.Message)"
        $script:State.WebhookQueue.Enqueue($result.Message)
        # Flush immediately so the result appears without waiting for next poll
        if ($webhookUrl) {
            _FlushWebhooks -WebhookUrl $webhookUrl
        }
    }
    catch {
        _DLog "Routing error: $_" -Level ERROR
        $script:State.WebhookQueue.Enqueue("[ERROR] Failed to process command: $_")
        if ($webhookUrl) {
            _FlushWebhooks -WebhookUrl $webhookUrl
        }
    }
}

# -----------------------------------------------------------------------------
#  Fetch new messages
# -----------------------------------------------------------------------------
function _GetNewMessages {
    param([string]$Token, [string]$ChannelId, [string]$AfterMessageId)

    $path = "/channels/$ChannelId/messages?limit=10"
    if ($AfterMessageId) { $path += "&after=$AfterMessageId" }

    $json = _DiscordRequest -Path $path -Token $Token
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
    param([string]$Token)

    $json = _DiscordRequest -Path '/users/@me' -Token $Token
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

    if (-not $token)     { _DLog 'No BotToken configured.' -Level WARN; return }
    if (-not $channelId) { _DLog 'No MonitorChannelId configured.' -Level WARN; return }

    _DLog "Listener starting (channel=$channelId, interval=${interval}s)"

    $botUserId = _GetBotUserId -Token $token
    _DLog "Bot user ID: $botUserId"

    # Seed last message ID
    $seed = _GetNewMessages -Token $token -ChannelId $channelId -AfterMessageId ''
    $lastId = if ($seed.Count -gt 0) {
        ($seed | Sort-Object { [decimal]$_.id } | Select-Object -Last 1).id
    } else { '' }

    Send-WebhookMessage -WebhookUrl $webhook -Content '[ONLINE] Etherium Command Center is now online.'
    $SharedState['ListenerRunning'] = $true

    while (-not ($SharedState.ContainsKey('StopListener') -and $SharedState['StopListener'] -eq $true)) {
        Start-Sleep -Seconds $interval

        # Send GUI-originated messages (DiscordOutbox or legacy SendDiscordMessage)
        if ($SharedState.ContainsKey('DiscordOutbox') -and $SharedState.DiscordOutbox) {
            $out = $null
            while ($SharedState.DiscordOutbox.TryDequeue([ref]$out)) {
                if (-not [string]::IsNullOrWhiteSpace($out)) {
                    $ok = _SendBotMessage -Token $token -ChannelId $channelId -Content $out
                    if (-not $ok) { _DLog "Failed to send GUI message." -Level WARN }
                }
            }
        } elseif ($SharedState.ContainsKey('SendDiscordMessage')) {
            $msg = $SharedState['SendDiscordMessage']
            if ($msg) {
                $SharedState['SendDiscordMessage'] = $null
                $ok = _SendBotMessage -Token $token -ChannelId $channelId -Content $msg
                if (-not $ok) { _DLog "Failed to send GUI message." -Level WARN }
            }
        }

        _FlushWebhooks -WebhookUrl $webhook

        try {
            $msgs = _GetNewMessages -Token $token -ChannelId $channelId -AfterMessageId $lastId
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
    _FlushWebhooks -WebhookUrl $webhook

    $SharedState['ListenerRunning'] = $false
    # Send offline notification
    Send-WebhookMessage -WebhookUrl $webhook -Content '[OFFLINE] Etherium Command Center is now offline.'
    _DLog 'Offline message sent.'
    _DLog 'Listener stopped.'
}

Export-ModuleMember -Function Start-DiscordListener, Send-WebhookMessage, Initialize-DiscordListener

