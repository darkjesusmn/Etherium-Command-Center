# =============================================================================
# ServerManager.psm1  -  Start / stop / monitor game server processes
# Compatible with PowerShell 5.1+
# =============================================================================



$script:State       = $null
$script:ModuleRoot  = $PSScriptRoot

function Initialize-ServerManager {
    param([hashtable]$SharedState)
    $script:State = $SharedState
    if (-not $script:State.ContainsKey('RunningServers'))  { $script:State['RunningServers']  = [hashtable]::Synchronized(@{}) }
    if (-not $script:State.ContainsKey('RestartCounters')) { $script:State['RestartCounters'] = [hashtable]::Synchronized(@{}) }
    if (-not $script:State.ContainsKey('RestFailCounts'))  { $script:State['RestFailCounts']  = [hashtable]::Synchronized(@{}) }
    if (-not $script:State.ContainsKey('StdinHandles'))    { $script:State['StdinHandles']    = [hashtable]::Synchronized(@{}) }
    if (-not $script:State.ContainsKey('ShutdownFlags'))   { $script:State['ShutdownFlags']   = [hashtable]::Synchronized(@{}) }
}

# ---------------------------------------------------------------------------
#  Rebuild RunningServers from live processes (used at startup)
# ---------------------------------------------------------------------------
function Sync-RunningServersFromProcesses {
    param([hashtable]$SharedState)

    Initialize-ServerManager -SharedState $SharedState

    $count = 0

    foreach ($pfx in @($SharedState.Profiles.Keys)) {
        $profile = $SharedState.Profiles[$pfx]
        if ($null -eq $profile) { continue }

        # If we already track a live PID, leave it alone
        if ($SharedState.RunningServers.ContainsKey($pfx)) {
            $entry = $SharedState.RunningServers[$pfx]
            $live  = try { Get-Process -Id $entry.Pid -ErrorAction Stop } catch { $null }
            if ($live -and -not $live.HasExited) { continue }
            $SharedState.RunningServers.Remove($pfx)
        }

        $procName = _CoalesceStr $profile.ProcessName ''
        if (-not $procName) { continue }

        $procs = Get-Process -Name $procName -ErrorAction SilentlyContinue
        if (-not $procs) { continue }

        $proc = $procs | Sort-Object StartTime -Descending | Select-Object -First 1

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
        $SharedState.ShutdownFlags[$pfx] = $false
        $count++

        _Log "[$($profile.GameName)] Detected running process '$procName' (PID $($proc.Id))"
    }

    return $count
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
    if ($script:State -and $script:State.ContainsKey('LogQueue')) {
        $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $script:State.LogQueue.Enqueue("[$ts][$Level][ServerManager] $Msg")
    }
    $col = switch ($Level) { 'ERROR'{'Red'} 'WARN'{'Yellow'} default{'Cyan'} }
    try { Write-Host "[ServerManager][$Level] $Msg" -ForegroundColor $col } catch {}
}

function _Webhook {
    param([string]$Message)
    if ($script:State -and $script:State.ContainsKey('WebhookQueue')) {
        $script:State.WebhookQueue.Enqueue($Message)
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
        [object]$Profile
    )

    _Log "REST DEBUG: endpoint='$Endpoint' method='$Method'" -Level DEBUG

    $adminPassword = ''

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

    try {
        $headers = @{}
        # Prefer Basic Auth with AdminPassword when available
        if (-not [string]::IsNullOrWhiteSpace($adminPassword)) {
            $basic = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes("admin:$adminPassword"))
            $headers['Authorization'] = "Basic $basic"
            _Log "REST DEBUG: using Basic auth" -Level DEBUG
        }
        # Also include RESTAPIKey if present (some builds use x-api-key)
        if (-not [string]::IsNullOrWhiteSpace($Password)) {
            $headers['x-api-key'] = $Password
            _Log "REST DEBUG: using x-api-key header" -Level DEBUG
        }

        if ($Method -eq 'GET') {
            $resp = Invoke-RestMethod -Uri $url -Method GET -Headers $headers -TimeoutSec 5
        }
        else {
            $resp = Invoke-RestMethod -Uri $url -Method POST -Headers $headers -Body '{}' -ContentType 'application/json' -TimeoutSec 5
        }
        return $resp
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
        _Log "Palworld REST error: $_" -Level WARN
        return $null
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

    $status = Get-ServerStatus -Prefix $prefix
    if ($status.Running) {
        _Log "[$($profile.GameName)] Already running (PID $($status.Pid))" -Level WARN
        return 'already_running'
    }

    $exe    = $profile.Executable
    $folder = $profile.FolderPath
    $args   = $profile.LaunchArgs

    if (-not (Test-Path $folder)) {
        _Log "[$($profile.GameName)] Folder not found: $folder" -Level ERROR
        return $false
    }

    _Log "[$($profile.GameName)] Starting -> $exe $args"
    _Log "[$($profile.GameName)] Start debug: folder='$folder' exe='$exe' args='$args'" -Level DEBUG
    _Webhook "[STARTING] $($profile.GameName) is starting..."

    try {
        [System.Environment]::SetEnvironmentVariable("SteamAppId", "892970", "Process")
        [System.Environment]::SetEnvironmentVariable("SteamGameId", "892970", "Process")
        $psi = [System.Diagnostics.ProcessStartInfo]::new()
        $psi.WorkingDirectory       = $folder
        $psi.UseShellExecute        = $false
        $psi.RedirectStandardInput  = $true
        $psi.CreateNoWindow         = $false

        # Optional stdout/stderr capture to a log file (for servers that don't write logs)
        $captureOut = $false
        $logPath = ''
        if ($null -ne $profile.CaptureOutput) {
            $captureOut = [bool]$profile.CaptureOutput
        }
        if ($captureOut -and $profile.ServerLogPath) {
            $logPath = [string]$profile.ServerLogPath
            if (-not [string]::IsNullOrWhiteSpace($logPath)) {
                $logDir = Split-Path -Path $logPath -Parent
                if ($logDir -and -not (Test-Path $logDir)) {
                    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
                }
                $psi.RedirectStandardOutput = $true
                $psi.RedirectStandardError  = $true
            } else {
                $captureOut = $false
            }
        }

        # --- BAT FILES ---
        if ($exe -match '\.bat$') {
            $psi.FileName  = 'cmd.exe'
            $psi.Arguments = "/c `"$exe`" $args"
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
            if ($args -and $args.Trim() -ne '') {
                $javaArgs += " $args"
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
            $psi.Arguments = $args
        }

        # --- NORMAL EXE ---
        else {
            $fullExe = Join-Path $folder $exe
            if (-not (Test-Path $fullExe)) {
                _Log "[$($profile.GameName)] Executable not found: $fullExe" -Level ERROR
                return $false
            }

            $psi.FileName  = $fullExe
            $psi.Arguments = $args
        }

        # --- START PROCESS ---
        $proc = [System.Diagnostics.Process]::Start($psi)
        if (-not $proc) {
            _Log "[$($profile.GameName)] Failed to start process (null proc returned)" -Level ERROR
            return $false
        }

        if ($captureOut -and $logPath) {
            try {
                $writer = New-Object System.IO.StreamWriter($logPath, $true, [System.Text.Encoding]::UTF8)
                $writer.AutoFlush = $true
                $proc.add_OutputDataReceived({
                    param($s,$e)
                    if ($e.Data) { $writer.WriteLine($e.Data) }
                })
                $proc.add_ErrorDataReceived({
                    param($s,$e)
                    if ($e.Data) { $writer.WriteLine($e.Data) }
                })
                $proc.BeginOutputReadLine()
                $proc.BeginErrorReadLine()

                if (-not $script:State.ContainsKey('LogWriters')) {
                    $script:State['LogWriters'] = [hashtable]::Synchronized(@{})
                }
                $script:State.LogWriters[$prefix] = $writer
                _Log "[$($profile.GameName)] Capturing stdout/stderr -> $logPath"
            } catch {
                _Log "[$($profile.GameName)] Failed to capture stdout/stderr: $_" -Level WARN
            }
        }

        # --- LOG PATH RESOLUTION ---
        $srvLogPath = ''
        if ($null -ne $profile.ServerLogPath -and "$($profile.ServerLogPath)".Trim() -ne '') {
            $srvLogPath = "$($profile.ServerLogPath)".Trim()
        }

        # --- REGISTER RUNNING SERVER ---
        $script:State.RunningServers[$prefix] = @{
            Pid           = $proc.Id
            StartTime     = Get-Date
            Process       = $proc
            ServerLogPath = $srvLogPath
        }

        if ($proc.StandardInput) {
            $script:State.StdinHandles[$prefix] = $proc.StandardInput
        }

        $script:State.ShutdownFlags[$prefix] = $false

        # --- LOG SUCCESS ---
        if ($IsAutoRestart) {
            _Log "[$($profile.GameName)] Auto-restarted (PID $($proc.Id))"
            _Webhook "[RESTARTED] $($profile.GameName) has been automatically restarted (PID $($proc.Id))."
        } else {
            _Log "[$($profile.GameName)] Started (PID $($proc.Id))"
            _Webhook "[ONLINE] $($profile.GameName) is now running (PID $($proc.Id))."
        }

        return $true
    }
    catch {
        _Log "[$($profile.GameName)] Failed to start: $_" -Level ERROR
        _Webhook "[ERROR] $($profile.GameName) failed to start: $_"
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

    $status = Get-ServerStatus -Prefix $prefix
    if (-not $status.Running) {
        _Log "[$($profile.GameName)] Safe shutdown requested but server is not running." -Level WARN
        return $true
    }

    $script:State.ShutdownFlags[$prefix] = $true

    # Step 1: Determine wait time
    $wait = 15
    if ($null -ne $profile.SaveWaitSeconds -and "$($profile.SaveWaitSeconds)".Trim() -ne '') {
        $wait = [int]$profile.SaveWaitSeconds
    }

    # Step 2: Save
    $saveMethod = _CoalesceStr $profile.SaveMethod 'none'
    if ($saveMethod -ne 'none') {
        _Log "[$($profile.GameName)] Sending save command (method: $saveMethod)..."
        if (-not $Quiet) { _Webhook "[SAVING] $($profile.GameName) - sending save command, then waiting ${wait}s before shutdown..." }
        $saveResult = _ExecuteSave -Prefix $prefix -Profile $profile
        if ($saveResult) {
            _Log "[$($profile.GameName)] Save command sent successfully. Waiting ${wait}s..."
        } else {
            _Log "[$($profile.GameName)] Save command failed or not configured - still waiting ${wait}s." -Level WARN
        }
    } else {
        _Log "[$($profile.GameName)] No save method configured - waiting ${wait}s for graceful shutdown..."
        if (-not $Quiet) { _Webhook "[WAITING] $($profile.GameName) - waiting ${wait}s before shutdown..." }
    }

    # Step 3: Wait
    Start-Sleep -Seconds $wait

    # Step 3: Stop
    _ExecuteStop -Prefix $prefix -Profile $profile

    if (-not $Quiet) { _Webhook "[STOPPED] $($profile.GameName) has been safely stopped." }
    _Log "[$($profile.GameName)] Safe shutdown complete."
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
    _Webhook "[RESTARTING] $($profile.GameName) - restart sequence started..."

    Invoke-SafeShutdown -Prefix $prefix -Quiet
    Start-Sleep -Seconds 2

    $started = Start-GameServer -Prefix $prefix
    if ($started) {
        _Webhook "[RESTARTED] $($profile.GameName) has been safely restarted."
    } else {
        _Webhook "[ERROR] $($profile.GameName) failed to restart."
    }
    return $started
}

# -----------------------------------------------------------------------------
#  SEND STDIN COMMAND
# -----------------------------------------------------------------------------
function Send-ServerStdin {
    param([string]$Prefix, [string]$Command)

    $prefix  = $Prefix.ToUpper()
    $profile = _GetProfile $prefix

    # ── Strategy 1: direct stdin handle (works for java/.exe launched directly) ──
    if ($script:State.StdinHandles.ContainsKey($prefix)) {
        try {
            $handle = $script:State.StdinHandles[$prefix]
            $handle.WriteLine($Command)
            $handle.Flush()
            _Log "[$($profile.GameName)] Sent stdin via handle: $Command"
            return $true
        }
        catch {
            _Log "[$($profile.GameName)] Stdin handle send failed: $_ - trying window method" -Level WARN
        }
    }

    # ── Strategy 2: find the game process window by ProcessName and send keys ──
    # Used when the server was launched via cmd.exe /c bat (stdin goes to cmd, not the game)
    $procName = _CoalesceStr $profile.ProcessName ''
    if (-not $procName) {
        _Log "[$($profile.GameName)] No ProcessName configured for window stdin fallback." -Level WARN
        return $false
    }

    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue

        $targets = Get-Process -Name $procName -ErrorAction SilentlyContinue
        if (-not $targets -or $targets.Count -eq 0) {
            _Log "[$($profile.GameName)] Process '$procName' not found for window stdin." -Level WARN
            return $false
        }

        $target = $targets | Select-Object -First 1

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
            return $false
        }

        [WinAPI]::ShowWindow($hwnd, 9)       | Out-Null  # SW_RESTORE
        [WinAPI]::SetForegroundWindow($hwnd) | Out-Null
        Start-Sleep -Milliseconds 200

        [System.Windows.Forms.SendKeys]::SendWait($Command)
        [System.Windows.Forms.SendKeys]::SendWait('{ENTER}')

        _Log "[$($profile.GameName)] Sent stdin via SendKeys to '$procName': $Command"
        return $true
    }
    catch {
        _Log "[$($profile.GameName)] SendKeys stdin failed: $_" -Level ERROR
        return $false
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

        default { return $false }
    }
}

# -----------------------------------------------------------------------------
#  INTERNAL: Execute stop (raw, no save logic)
# -----------------------------------------------------------------------------
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
    } catch { }
}

function _ExecuteStop {
    param([string]$Prefix, [object]$Profile)
    $prefix     = $Prefix.ToUpper()
    $stopMethod = _CoalesceStr $Profile.StopMethod 'processKill'

    $entry = $script:State.RunningServers[$prefix]
    if (-not $entry) { return }
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
        default {
            # processKill - kill the wrapper PID and its entire child tree
            _KillProcessTree -RootPid $entry.Pid
        }
    }

    $script:State.RunningServers.Remove($prefix)
    $script:State.StdinHandles.Remove($prefix)

    # Close any stdout/stderr capture writer
    try {
        if ($script:State -and $script:State.ContainsKey('LogWriters')) {
            $writer = $script:State.LogWriters[$prefix]
            if ($writer) { $writer.Close() }
            $script:State.LogWriters.Remove($prefix) | Out-Null
        }
    } catch { }
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

    # Query server state
    $stateOk = $false
    $statusEndpoint = ''
    if ($Profile -and $Profile.RestStatusEndpoint) {
        $statusEndpoint = [string]$Profile.RestStatusEndpoint
    } elseif ($Profile -and $Profile.GameName -and $Profile.GameName -notmatch 'Palworld') {
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
                        $SharedState['PalworldState'] = $state
                        $stateOk = $true
                    } else {
                        _Log "Palworld REST info response not JSON: $trim" -Level WARN
                    }
                } else {
                    $SharedState['PalworldState'] = $stateResp
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
                    $SharedState['PalworldPlayers'] = $players
                    $playersOk = $true
                } else {
                    _Log "Palworld REST players response not JSON: $trim" -Level WARN
                }
            } else {
                $SharedState['PalworldPlayers'] = $playersResp
                $playersOk = $true
            }
        } catch {
            _Log "Palworld REST players parse failed: $_" -Level WARN
        }
    }

    if ($stateOk -or $playersOk) {
        $SharedState.RestFailCounts[$pfx] = 0
    } else {
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
function Start-ServerMonitor {
    param([hashtable]$SharedState)
    _Log "Server monitor started."

    while (-not ($SharedState.ContainsKey('StopListener') -and $SharedState['StopListener'] -eq $true)) {
        Start-Sleep -Seconds 15
        $profileCount = 0
        $runningCount = 0
        try { $profileCount = @($SharedState.Profiles.Keys).Count } catch { $profileCount = 0 }
        try { $runningCount = if ($SharedState.RunningServers) { $SharedState.RunningServers.Count } else { 0 } } catch { $runningCount = 0 }
        _Log "Monitor tick: profiles=$profileCount running=$runningCount" -Level DEBUG

        foreach ($prefix in @($SharedState.Profiles.Keys)) {
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

                # Safe property access for OrderedDictionary
                $autoRestart = $true
                if ($null -ne $profile.EnableAutoRestart) {
                    $autoRestart = [bool]$profile.EnableAutoRestart
                }
                if (-not $autoRestart) { continue }

                # Safe check for shutdown flag using property access instead of ContainsKey
                $shutdownFlag = $false
                if ($null -ne $SharedState.ShutdownFlags) {
                    $shutdownFlag = $SharedState.ShutdownFlags[$prefix]
                    if ($shutdownFlag -eq $true) { continue }
                }

                $hasEntry = $false
                if ($SharedState.RunningServers -and $SharedState.RunningServers.ContainsKey($prefix)) { $hasEntry = $true }
                _Log "Monitor check: prefix=$prefix autoRestart=$autoRestart shutdownFlag=$shutdownFlag hasEntry=$hasEntry" -Level DEBUG

                $entry = $SharedState.RunningServers[$prefix]
                if (-not $entry) { continue }

                $proc = try { Get-Process -Id $entry.Pid -ErrorAction Stop } catch { $null }
                if ($null -ne $proc -and -not $proc.HasExited) { continue }

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
                    $SharedState.WebhookQueue.Enqueue("[WARNING] $($profile.GameName) has crashed $count times this hour. Auto-restart suspended.")
                    $SharedState.RunningServers.Remove($prefix)
                    continue
                }

                $SharedState.RestartCounters[$key] = $count + 1
                $delay = 10
                
                if ($null -ne $profile.RestartDelaySeconds) {
                    $delay = [int]$profile.RestartDelaySeconds
                }

                $SharedState.WebhookQueue.Enqueue("[CRASHED] $($profile.GameName) crashed. Restarting in ${delay}s (attempt $($count+1)/$max)...")
                $SharedState.RunningServers.Remove($prefix)

                Start-Sleep -Seconds $delay

                Initialize-ServerManager -SharedState $SharedState
                Start-GameServer -Prefix $prefix -IsAutoRestart
            }
            catch {
                _Log "Error in monitor loop for $prefix : $_" -Level ERROR
            }
        }
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
        [int]   $TimeoutMs = 5000
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

    try {
        _Log "RCON DEBUG: connect ${Host}:${Port} cmd='$Command'" -Level DEBUG
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

        $stream = $tcp.GetStream()
        $reader = [System.IO.BinaryReader]::new($stream)
        $writer = [System.IO.BinaryWriter]::new($stream)

        # Authenticate (type 3, id 1)
        $authPacket = _RconPacket -Id 1 -Type 3 -Body $Password
        $writer.Write($authPacket)
        $writer.Flush()

        # Read auth response - server sends two packets on auth
        $resp1 = _ReadPacket $reader
        # Some servers send an empty type-0 first, then the auth response
        if ($resp1.Type -eq 0) {
            $resp1 = _ReadPacket $reader
        }

        if ($resp1.Id -eq -1) {
            throw "RCON authentication failed - wrong password"
        }

        # Send command (type 2, id 2)
        $cmdPacket = _RconPacket -Id 2 -Type 2 -Body $Command
        $writer.Write($cmdPacket)
        $writer.Flush()

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

        return @{
            Success  = $true
            Response = $responseBody.ToString().Trim()
        }
    }
    catch {
        _Log "RCON DEBUG: error=$_" -Level DEBUG
        return @{
            Success  = $false
            Response = ''
            Error    = "$_"
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

    if ($null -eq $profile.Commands -or $null -eq $profile.Commands[$CommandName]) {
        $available = if ($null -ne $profile.Commands) { @($profile.Commands.Keys) -join ', ' } else { 'none' }
        return @{ Message = "[ERROR] Command '$CommandName' not found. Available: $available"; Success = $false }
    }

    $cmdDef  = $profile.Commands[$CommandName]
    $cmdType = if ($null -ne $cmdDef.Type) { $cmdDef.Type } else { 'stdin' }

    _Log "[$($profile.GameName)] Command debug: name='$CommandName' type='$cmdType'" -Level DEBUG

    $result    = $false
    $resultMsg = ''

    switch ($cmdType) {
        'Start' {
            $result = Start-GameServer -Prefix $prefix
            $resultMsg = switch ($result) {
                'already_running' { "[INFO] $($profile.GameName) is already running."     }
                $true             { "[OK] $($profile.GameName) is starting..."             }
                default           { "[ERROR] Failed to start $($profile.GameName)"         }
            }
        }
        'Stop' {
            $result    = Invoke-SafeShutdown -Prefix $prefix
            $resultMsg = if ($result) { "[OK] $($profile.GameName) stop sequence initiated." } else { "[ERROR] Failed to stop $($profile.GameName)" }
        }
        'Restart' {
            $result    = Restart-GameServer -Prefix $prefix
            $resultMsg = if ($result) { "[OK] $($profile.GameName) restart sequence initiated." } else { "[ERROR] Failed to restart $($profile.GameName)" }
        }
        'Status' {
            $status    = Get-ServerStatus -Prefix $prefix
            $result    = $true
            $uptime    = if ($status.Running) { "$([Math]::Round($status.Uptime.TotalMinutes,1)) min" } else { 'N/A' }
            $resultMsg = if ($status.Running) { "[ONLINE] $($profile.GameName) is running (PID $($status.Pid), uptime $uptime)" } else { "[OFFLINE] $($profile.GameName) is not running." }
        }
        'SendCommand' {
            $cmd       = if ($null -ne $cmdDef.Command) { $cmdDef.Command } else { $CommandName }
            $result    = Send-ServerStdin -Prefix $prefix -Command $cmd
            $resultMsg = if ($result) { "[OK] Sent '$cmd' to $($profile.GameName)" } else { "[ERROR] Failed to send '$cmd' to $($profile.GameName)" }
        }
        'stdin' {
            $stdinCmd  = if ($null -ne $cmdDef.StdinCommand) { $cmdDef.StdinCommand } else { $CommandName }
            $result    = Send-ServerStdin -Prefix $prefix -Command $stdinCmd
            $resultMsg = if ($result) { "[OK] Sent '$stdinCmd' to $($profile.GameName)" } else { "[ERROR] Failed to send to $($profile.GameName)" }
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
                    -Profile $Profile

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
        default {
            $resultMsg = "[ERROR] Unknown command type '$cmdType'. Valid: Start, Stop, Restart, Status, SendCommand, stdin, http, script"
        }
    }

    return @{ Message = $resultMsg; Success = $result }
}

Export-ModuleMember -Function `
    Initialize-ServerManager, Sync-RunningServersFromProcesses, Get-ServerStatus, `
    Start-GameServer, Stop-GameServer, Restart-GameServer, `
    Invoke-SafeShutdown, Invoke-ProfileCommand, `
    Send-ServerStdin, Invoke-ServerHttp, Invoke-CustomScript, `
    Invoke-RconCommand, Start-ServerMonitor
