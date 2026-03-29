# =============================================================================
# Launch.ps1  -  Etherium Command Center Entry Point
# Double-click, or right-click -> Run with PowerShell.
# Handles elevation, unblocking, and execution policy automatically.
# =============================================================================
#Requires -Version 5.1

Set-StrictMode -Off
$ErrorActionPreference = 'Stop'

# Resolve own path whether run directly, dot-sourced, or via Start-Process
$_self = if ($PSCommandPath) { $PSCommandPath } `
         elseif ($MyInvocation.MyCommand.Path) { $MyInvocation.MyCommand.Path } `
         else { $null }

# Everything inside try/catch so errors are always visible before window closes
try {

    # ── Step 1: Elevate to Administrator if needed ────────────────────────────
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
               ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if (-not $isAdmin) {
        if (-not $_self) {
            throw "Cannot determine script path. Please run Launch.ps1 directly (do not dot-source it)."
        }
        # If debug logging is enabled in settings, keep the elevated window open
        $debugHold = $false
        try {
            $cfgPath = Join-Path (Split-Path -Path $_self -Parent) 'Config\Settings.json'
            if (Test-Path $cfgPath) {
                $raw = Get-Content -Path $cfgPath -Raw -Encoding UTF8 | ConvertFrom-Json
                if ($raw -and $raw.PSObject.Properties.Name -contains 'EnableDebugLogging') {
                    $debugHold = [bool]$raw.EnableDebugLogging
                }
            }
        } catch { }
        Write-Host "Requesting Administrator privileges..." -ForegroundColor Yellow
        Start-Process powershell.exe `
            -ArgumentList "-ExecutionPolicy Bypass $(if ($debugHold) { '-NoExit ' } else { '' })-File `"$_self`"" `
            -Verb RunAs
        exit
    }

    # ── Step 2: Unblock all files ─────────────────────────────────────────────
    Write-Host "Unblocking Etherium Command Center files..." -ForegroundColor Cyan
    Get-ChildItem -Path $PSScriptRoot -Recurse -Include '*.ps1','*.psm1','*.json' |
        ForEach-Object { Unblock-File -Path $_.FullName -ErrorAction SilentlyContinue }
    Write-Host "Files unblocked." -ForegroundColor Green

    # ── Step 3: Set execution policy ─────────────────────────────────────────
    try {
        Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
        Write-Host "Execution policy set to RemoteSigned." -ForegroundColor Green
    } catch {
        Write-Host "Could not set execution policy (non-fatal): $_" -ForegroundColor Yellow
    }

    # ── Resolve paths ─────────────────────────────────────────────────────────
    $Root        = $PSScriptRoot
    $ModulesDir  = Join-Path $Root 'Modules'
    $ProfilesDir = Join-Path $Root 'Profiles'
    $ConfigDir   = Join-Path $Root 'Config'
    $LogsDir     = Join-Path $Root 'Logs'
    $ConfigPath  = Join-Path $ConfigDir 'Settings.json'

    foreach ($d in @($ProfilesDir, $ConfigDir, $LogsDir)) {
        if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
    }

    # ── Start transcript - captures ALL console output into a file the GUI can tail
    $TranscriptPath = Join-Path $LogsDir "console.log"
    try {
        # Clear the log file on every launch so the GUI only shows this session
        [System.IO.File]::WriteAllText($TranscriptPath, '')
        Start-Transcript -Path $TranscriptPath -Force | Out-Null
    } catch { }

    # ── Load modules ──────────────────────────────────────────────────────────
    $modulePaths = @(
        (Join-Path $ModulesDir 'Logging.psm1'),
        (Join-Path $ModulesDir 'ProfileManager.psm1'),
        (Join-Path $ModulesDir 'ServerManager.psm1'),
        (Join-Path $ModulesDir 'DiscordListener.psm1'),
        (Join-Path $ModulesDir 'GUI.psm1')
    )
    foreach ($mp in $modulePaths) {
        if (-not (Test-Path $mp)) { throw "Module not found: $mp" }
        Import-Module $mp -Force
    }

    function Assert-BootstrapCommand {
        param(
            [Parameter(Mandatory = $true)][string]$Name,
            [Parameter(Mandatory = $true)][string]$ExpectedModule,
            [scriptblock]$Recovery
        )

        $cmd = Get-Command -Name $Name -ErrorAction SilentlyContinue
        $resolvedModule = ''
        if ($cmd) {
            $resolvedModule = if ($cmd.Source) { $cmd.Source } else { $cmd.ModuleName }
        }

        $matchesExpected = ($cmd -and ($resolvedModule -eq $ExpectedModule))
        if (-not $matchesExpected -and $Recovery) {
            & $Recovery
            $cmd = Get-Command -Name $Name -ErrorAction SilentlyContinue
            $resolvedModule = ''
            if ($cmd) {
                $resolvedModule = if ($cmd.Source) { $cmd.Source } else { $cmd.ModuleName }
            }
            $matchesExpected = ($cmd -and ($resolvedModule -eq $ExpectedModule))
        }

        if (-not $cmd) {
            throw "Startup bootstrap validation failed: required command '$Name' is unavailable."
        }

        if (-not $matchesExpected) {
            throw "Startup bootstrap validation failed: command '$Name' resolved from '$resolvedModule' instead of '$ExpectedModule'."
        }
    }

    function Import-GuiModule {
        param([string]$Path)
        if (-not (Test-Path $Path)) { throw "GUI module not found: $Path" }
        try {
            $existingGui = Get-Module -Name 'GUI' -ErrorAction SilentlyContinue
            if ($existingGui) {
                Remove-Module -Name 'GUI' -Force -ErrorAction SilentlyContinue
            }
        } catch { }
        Import-Module $Path -Force
    }

    Assert-BootstrapCommand -Name 'Initialize-Logging' -ExpectedModule 'Logging' -Recovery {
        Import-Module (Join-Path $ModulesDir 'Logging.psm1') -Force
    }
    Assert-BootstrapCommand -Name 'Write-Log' -ExpectedModule 'Logging' -Recovery {
        Import-Module (Join-Path $ModulesDir 'Logging.psm1') -Force
    }

    # ── Initialise logging ────────────────────────────────────────────────────
    Initialize-Logging -LogDir $LogsDir
    Write-Log "Etherium Command Center starting..."
    Write-Log "Root: $Root"

    # ── Load settings ─────────────────────────────────────────────────────────
    $defaultSettings = @{
        BotToken            = ''
        WebhookUrl          = ''
        MonitorChannelId    = ''
        CommandPrefix       = '!'
        PollIntervalSeconds = 2
        EnableDebugLogging  = $false
        EnablePerformanceDebugMode = $false
        AutoSaveEnabled           = $false
        AutoSaveIntervalMinutes   = 30
        ScheduledRestartEnabled   = $false
        ScheduledRestartHours     = 0
    }
    $settings = $defaultSettings.Clone()
    if (Test-Path $ConfigPath) {
        try {
            $raw = Get-Content -Path $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($key in $raw.PSObject.Properties.Name) { $settings[$key] = $raw.$key }
            Write-Log "Settings loaded from $ConfigPath"
            Write-Log "Settings debug: EnableDebugLogging=$($settings.EnableDebugLogging) EnablePerformanceDebugMode=$($settings.EnablePerformanceDebugMode)" -Level DEBUG
        } catch {
            Write-Log "Could not parse settings file - using defaults. Error: $_" -Level WARN
        }
    } else {
        Write-Log "No settings file found - will be created when you save settings in the GUI."
    }

    # Apply debug logging toggle (if present)
    try {
        if (Get-Command Set-DebugLoggingEnabled -ErrorAction SilentlyContinue) {
            Set-DebugLoggingEnabled -Enabled ([bool]$settings.EnableDebugLogging)
        }
    } catch { }

    # ── Load game profiles ────────────────────────────────────────────────────
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "PROFILE LOADER: Looking in: $ProfilesDir" -ForegroundColor Cyan
    Write-Host "PROFILE LOADER: Folder exists: $(Test-Path $ProfilesDir)" -ForegroundColor Cyan

    $profiles = [hashtable]::Synchronized(@{})

    $profileFiles = @(Get-ChildItem -Path $ProfilesDir -Filter '*.json' -ErrorAction SilentlyContinue)
    Write-Host "PROFILE LOADER: Found $($profileFiles.Count) JSON file(s)" -ForegroundColor Cyan

    foreach ($file in $profileFiles) {
        Write-Host "PROFILE LOADER: Reading $($file.Name)..." -ForegroundColor Yellow
        try {
            $json = Get-Content $file.FullName -Raw -Encoding UTF8
            Write-Host "  JSON length: $($json.Length) chars" -ForegroundColor Gray

            $raw = $json | ConvertFrom-Json
            Write-Host "  ConvertFrom-Json OK. Type: $($raw.GetType().Name)" -ForegroundColor Gray

            $profile = ConvertTo-Hashtable -Object $raw
            if (-not ($profile.ContainsKey('BlockStartIfRamPercentUsed'))) {
                $profile['BlockStartIfRamPercentUsed'] = 0
            }
            if (-not ($profile.ContainsKey('BlockStartIfFreeRamBelowGB'))) {
                $profile['BlockStartIfFreeRamBelowGB'] = 0
            }
            if (-not ($profile.ContainsKey('StartupTimeoutSeconds'))) {
                $profile['StartupTimeoutSeconds'] = 300
            }
            if (-not ($profile.ContainsKey('ShutdownIfNoPlayersAfterStartupMinutes'))) {
                $profile['ShutdownIfNoPlayersAfterStartupMinutes'] = 0
            }
            if (-not ($profile.ContainsKey('ShutdownIfEmptyAfterLastPlayerLeavesMinutes'))) {
                $profile['ShutdownIfEmptyAfterLastPlayerLeavesMinutes'] = 0
            }
            Write-Host "  ConvertTo-Hashtable OK. Type: $($profile.GetType().Name)" -ForegroundColor Gray
            Write-Host "  Prefix value: '$($profile.Prefix)'" -ForegroundColor Gray
            Write-Host "  GameName value: '$($profile.GameName)'" -ForegroundColor Gray
            Write-Log "Profile debug: File=$($file.Name) Prefix=$($profile.Prefix) GameName=$($profile.GameName)" -Level DEBUG
            Write-Log "Profile debug: FolderPath=$($profile.FolderPath) LogRoot=$($profile.ServerLogRoot) ConfigRoot=$($profile.ConfigRoot)" -Level DEBUG

            if ($null -ne $profile.Prefix -and $profile.Prefix -ne '') {
                $pfx = $profile.Prefix.ToUpper()
                $profiles[$pfx] = $profile
                Write-Host "  SUCCESS: Loaded [$pfx] $($profile.GameName)" -ForegroundColor Green
                Write-Log "  Loaded: [$pfx] $($profile.GameName)"
            } else {
                Write-Host "  SKIPPED: Prefix is null or empty" -ForegroundColor Red
            }
        } catch {
            Write-Host "  ERROR loading $($file.Name): $_" -ForegroundColor Red
            Write-Log "  Error loading $($file.Name): $_" -Level WARN
        }
    }

    Write-Host "PROFILE LOADER: Total loaded: $($profiles.Count)" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Log "Loaded $($profiles.Count) profile(s)."


    # ── Build shared state ────────────────────────────────────────────────────
    $SharedState = [hashtable]::Synchronized(@{
        Settings        = $settings
        Profiles        = $profiles
        RunningServers  = [hashtable]::Synchronized(@{})
        StdinHandles    = [hashtable]::Synchronized(@{})
        ShutdownFlags   = [hashtable]::Synchronized(@{})
        RestartCounters = [hashtable]::Synchronized(@{})
        PendingScheduledRestarts = [hashtable]::Synchronized(@{})
        LogQueue        = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
        WebhookQueue    = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
        GameLogQueue    = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
        PlayersRequests = [hashtable]::Synchronized(@{})
        LatestPlayers   = [hashtable]::Synchronized(@{})
        PlayerActivityState = [hashtable]::Synchronized(@{})
        RestartProgram  = $false
        DiscordOutbox   = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
        LastAutoSave    = [hashtable]::Synchronized(@{})
        SentWarnings    = [hashtable]::Synchronized(@{})
        ServerRuntimeState = [hashtable]::Synchronized(@{})
        ReloadUI        = $false
        StopMonitor     = $false
        DiscordOfflineNoticeSent = $false
    })

    # ── Initialise ServerManager ──────────────────────────────────────────────
    Initialize-ServerManager -SharedState $SharedState
    $detected = Sync-RunningServersFromProcesses -SharedState $SharedState
    if ($detected -gt 0) {
        Write-Log "Detected $detected already-running server(s)."
    }
    try {
        $rehydrated = Restore-ServerRuntimeStateFromSharedState -SharedState $SharedState
        if ($rehydrated -gt 0) {
            Write-Log "Restored runtime/timer state for $rehydrated server slot(s) from shared state."
        }
    } catch {
        Write-Log "Initial runtime/timer rehydration failed: $_" -Level WARN
    }

    # ── Register engine exit hook (catches window X / Ctrl+C / taskkill) ───────
    $null = Register-EngineEvent -SourceIdentifier PowerShell.Exiting -Action {
        try {
            $webhook = $SharedState.Settings.WebhookUrl
            $alreadySent = $false
            try {
                if ($SharedState.ContainsKey('DiscordOfflineNoticeSent')) {
                    $alreadySent = [bool]$SharedState['DiscordOfflineNoticeSent']
                }
            } catch { $alreadySent = $false }

            if ($webhook -and -not $alreadySent) {
                try { $SharedState['DiscordOfflineNoticeSent'] = $true } catch { }
                # Best-effort direct send - no queuing, we are dying right now
                $content = if (Get-Command -Name New-DiscordSystemMessage -ErrorAction SilentlyContinue) { New-DiscordSystemMessage -Event 'offline' } else { '[OFFLINE] Etherium Command Center is now offline.' }
                $payload = (@{ content = $content } | ConvertTo-Json -Compress)
                $req = [System.Net.HttpWebRequest]::Create($webhook)
                $req.Method      = 'POST'
                $req.ContentType = 'application/json'
                $req.UserAgent   = 'DiscordBot (Etherium Command Center, 1.0)'
                $req.Timeout     = 4000
                $bytes = [System.Text.Encoding]::UTF8.GetBytes($payload)
                $req.ContentLength = $bytes.Length
                $s = $req.GetRequestStream()
                $s.Write($bytes, 0, $bytes.Length)
                $s.Close()
                $req.GetResponse().Close()
            }
        } catch { }
    }

    # ── Discord Listener runspace ─────────────────────────────────────────────
    Write-Log "Starting Discord listener runspace..."
    $listenerRunspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $listenerRunspace.ApartmentState = 'STA'
    $listenerRunspace.ThreadOptions  = 'ReuseThread'
    $listenerRunspace.Open()
    $listenerRunspace.SessionStateProxy.SetVariable('ModulesDir',  $ModulesDir)
    $listenerRunspace.SessionStateProxy.SetVariable('SharedState', $SharedState)

    $listenerPS = [System.Management.Automation.PowerShell]::Create()
    $listenerPS.Runspace = $listenerRunspace
    $listenerPS.AddScript({
        Set-StrictMode -Off
        $ErrorActionPreference = 'Continue'
        try {
            Import-Module (Join-Path $ModulesDir 'Logging.psm1')         -Force
            Import-Module (Join-Path $ModulesDir 'ProfileManager.psm1')  -Force
            Import-Module (Join-Path $ModulesDir 'ServerManager.psm1')   -Force
            Import-Module (Join-Path $ModulesDir 'DiscordListener.psm1') -Force
            Start-DiscordListener -SharedState $SharedState
        } catch {
            $SharedState.LogQueue.Enqueue(
                "[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][ERROR][ListenerRunspace] $_")
        }
    }) | Out-Null
    $listenerHandle = $listenerPS.BeginInvoke()
    Write-Log "Discord listener started."
    # Expose listener handles so the GUI can restart/stop without duplicates
    $SharedState['ListenerRunspace'] = $listenerRunspace
    $SharedState['ListenerPS']       = $listenerPS
    $SharedState['ListenerHandle']   = $listenerHandle

    # ── Server Monitor runspace ───────────────────────────────────────────────
    Write-Log "Starting server monitor runspace..."
    $monitorRunspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $monitorRunspace.ApartmentState = 'STA'
    $monitorRunspace.ThreadOptions  = 'ReuseThread'
    $monitorRunspace.Open()
    $monitorRunspace.SessionStateProxy.SetVariable('ModulesDir',  $ModulesDir)
    $monitorRunspace.SessionStateProxy.SetVariable('SharedState', $SharedState)

    $monitorPS = [System.Management.Automation.PowerShell]::Create()
    $monitorPS.Runspace = $monitorRunspace
    $monitorPS.AddScript({
        Set-StrictMode -Off
        $ErrorActionPreference = 'Continue'
        try {
            Import-Module (Join-Path $ModulesDir 'Logging.psm1')        -Force
            Import-Module (Join-Path $ModulesDir 'ProfileManager.psm1') -Force
            Import-Module (Join-Path $ModulesDir 'ServerManager.psm1')  -Force
            Initialize-ServerManager -SharedState $SharedState
            Start-ServerMonitor      -SharedState $SharedState
        } catch {
            $SharedState.LogQueue.Enqueue(
                "[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][ERROR][MonitorRunspace] $_")
        }
    }) | Out-Null
    $monitorHandle = $monitorPS.BeginInvoke()
    Write-Log "Server monitor started."

    # ── Launch GUI (blocks until closed) ─────────────────────────────────────
    $pendingUiReloadRelaunch = $false
    $fatalGuiReloadFailure = $false
    do {
        $SharedState['ReloadUI'] = $false
        try {
            Import-GuiModule -Path (Join-Path $ModulesDir 'GUI.psm1')
            Assert-BootstrapCommand -Name 'Start-GUI' -ExpectedModule 'GUI' -Recovery {
                Import-GuiModule -Path (Join-Path $ModulesDir 'GUI.psm1')
            }
            Assert-BootstrapCommand -Name 'Write-Log' -ExpectedModule 'Logging' -Recovery {
                Import-Module (Join-Path $ModulesDir 'Logging.psm1') -Force
            }
            Write-Log "Launching GUI..."
Start-GUI -SharedState $SharedState -ConfigPath $ConfigPath -ProfilesDir $ProfilesDir -TranscriptPath $TranscriptPath
        } catch {
            $stage = if ($pendingUiReloadRelaunch) { 'UI reload relaunch' } else { 'GUI launch' }
            Write-Log "Fatal failure during ${stage}: $_" -Level ERROR

            if ($pendingUiReloadRelaunch) {
                $fatalGuiReloadFailure = $true
                $SharedState['FatalGuiFailure'] = $true
                $SharedState['FatalGuiFailureReason'] = 'UI reload failed to relaunch the GUI.'
                break
            }

            throw
        }

        if ($SharedState.ReloadUI -eq $true) {
            Write-Log "UI reload requested. Rebuilding GUI without stopping running servers."
            try {
                $detected = Sync-RunningServersFromProcesses -SharedState $SharedState
                Write-Log "UI reload sync complete. Running servers tracked: $($SharedState.RunningServers.Count). Newly detected: $detected."
                try {
                    $rehydrated = Restore-ServerRuntimeStateFromSharedState -SharedState $SharedState
                    Write-Log "UI reload runtime/timer rehydration complete. Restored states: $rehydrated."
                } catch {
                    Write-Log "UI reload runtime/timer rehydration failed: $_" -Level WARN
                }
            } catch {
                Write-Log "UI reload sync failed: $_" -Level WARN
            }
            $pendingUiReloadRelaunch = $true
            continue
        }

        $pendingUiReloadRelaunch = $false
        break
    } while ($true)

    # ── Cleanup ───────────────────────────────────────────────────────────────
    Write-Log "GUI closed - shutting down..." 

    $stopServersBeforeExit = ($fatalGuiReloadFailure -eq $true -or $SharedState.RestartProgram -eq $true)
    if ($stopServersBeforeExit) {
        if ($fatalGuiReloadFailure) {
            Write-Log "Fatal GUI reload failure detected. Gracefully stopping all running servers before exit." -Level WARN
        } else {
            Write-Log "Restart requested. Stopping all servers before relaunch..."
        }
        try {
            foreach ($pfx in @($SharedState.RunningServers.Keys)) {
                try { Invoke-SafeShutdown -Prefix $pfx -Quiet | Out-Null } catch { }
            }
        } catch { }
    }

    # Signal background loops to exit during full program shutdown.
    $SharedState['StopListener'] = $true
    $SharedState['StopMonitor']  = $true

    # Give the ACTIVE listener enough time to send its offline webhook and exit.
    # This must be at least as long as the Discord/webhook HTTP timeout or the
    # fallback sender below can race the listener and produce duplicate OFFLINE messages.
    $listenerTimeoutMs = 10000
    try {
        if ($SharedState.Settings -and $SharedState.Settings.ContainsKey('DiscordHttpTimeoutMs')) {
            $cfgTimeout = [int]$SharedState.Settings.DiscordHttpTimeoutMs
            if ($cfgTimeout -gt 0) { $listenerTimeoutMs = $cfgTimeout }
        }
    } catch { }
    $listenerWaitSeconds = [Math]::Max(8, [int][Math]::Ceiling(($listenerTimeoutMs + 3000) / 1000.0))
    $deadline = (Get-Date).AddSeconds($listenerWaitSeconds)
    while ((Get-Date) -lt $deadline) {
        $currentListenerHandle = $null
        if ($SharedState.ContainsKey('ListenerHandle')) {
            $currentListenerHandle = $SharedState['ListenerHandle']
        }
        $listenerDone = $false
        try { $listenerDone = ($null -ne $currentListenerHandle -and $currentListenerHandle.IsCompleted) } catch { $listenerDone = $false }
        if ($listenerDone -and $monitorHandle.IsCompleted) { break }
        Start-Sleep -Milliseconds 200
    }

    # Fallback: if listener didn't finish in time, send the offline message directly
    $currentListenerHandle = $null
    if ($SharedState.ContainsKey('ListenerHandle')) {
        $currentListenerHandle = $SharedState['ListenerHandle']
    }
    $listenerStillRunning = $true
    try { $listenerStillRunning = ($null -eq $currentListenerHandle -or -not $currentListenerHandle.IsCompleted) } catch { $listenerStillRunning = $true }
    if ($listenerStillRunning) {
        try {
            $webhook = $SharedState.Settings.WebhookUrl
            $alreadySent = $false
            try {
                if ($SharedState.ContainsKey('DiscordOfflineNoticeSent')) {
                    $alreadySent = [bool]$SharedState['DiscordOfflineNoticeSent']
                }
            } catch { $alreadySent = $false }

            if ($webhook -and -not $alreadySent) {
                try { $SharedState['DiscordOfflineNoticeSent'] = $true } catch { }
                $content = if (Get-Command -Name New-DiscordSystemMessage -ErrorAction SilentlyContinue) { New-DiscordSystemMessage -Event 'offline' } else { '[OFFLINE] Etherium Command Center is now offline.' }
                $payload = (@{ content = $content } | ConvertTo-Json -Compress)
                $req = [System.Net.HttpWebRequest]::Create($webhook)
                $req.Method      = 'POST'
                $req.ContentType = 'application/json'
                $req.UserAgent   = 'DiscordBot (Etherium Command Center, 1.0)'
                $req.Timeout     = 4000
                $bytes = [System.Text.Encoding]::UTF8.GetBytes($payload)
                $req.ContentLength = $bytes.Length
                $s = $req.GetRequestStream()
                $s.Write($bytes, 0, $bytes.Length)
                $s.Close()
                $req.GetResponse().Close()
                Write-Log "Offline message sent (fallback)."
            }
        } catch {
            Write-Log "Could not send offline message: $_" -Level WARN
        }
    }

    $currentListenerPS = $null
    $currentListenerRunspace = $null
    if ($SharedState.ContainsKey('ListenerPS')) { $currentListenerPS = $SharedState['ListenerPS'] }
    if ($SharedState.ContainsKey('ListenerRunspace')) { $currentListenerRunspace = $SharedState['ListenerRunspace'] }
    try { if ($currentListenerPS) { $currentListenerPS.Stop() } } catch { }
    try { if ($monitorPS) { $monitorPS.Stop() } } catch { }
    try { if ($currentListenerRunspace) { $currentListenerRunspace.Close() } } catch { }
    try { if ($monitorRunspace) { $monitorRunspace.Close() } } catch { }
    Write-Log "Etherium Command Center stopped cleanly."
    try { Stop-Transcript | Out-Null } catch { }

    if ($SharedState.RestartProgram -eq $true) {
        $self = $_self
        if ($self) {
            $hold = $false
            try { $hold = [bool]$settings.EnableDebugLogging } catch { }
            $args = "-ExecutionPolicy Bypass " + $(if ($hold) { "-NoExit " } else { "" }) + "-File `"$self`""
            Start-Process powershell.exe -ArgumentList $args -Verb RunAs
        }
        exit 0
    }

    if ($settings.EnableDebugLogging -eq $true) {
        Write-Host "Debug logging is enabled. Press Enter to close." -ForegroundColor Yellow
        Read-Host | Out-Null
        return
    }

    exit 0


} catch {
    Write-Host ""
    Write-Host "========== ERROR ==========" -ForegroundColor Red
    Write-Host $_ -ForegroundColor Red
    Write-Host "===========================" -ForegroundColor Red
    Write-Host ""
    Read-Host "Press Enter to close"
}
