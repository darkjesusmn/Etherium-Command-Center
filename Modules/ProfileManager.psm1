# =============================================================================
# ProfileManager.psm1  -  Game profile auto-generation, loading & safety checks
# PowerShell 5.1 compatible. No external dependencies.
# =============================================================================


# -----------------------------------------------------------------------------
#  SAFETY RULES  -  enforced at profile load AND command add time
# -----------------------------------------------------------------------------
# NOTE: 'player' is intentionally NOT in this list.
# The word 'players' (list online players) is a safe, read-only status command.
# We only block targeting commands that act ON a specific player.
$script:ForbiddenKeywords    = @('kick','ban','tp','teleport','op','deop',
                                  'whitelist','give','take','pardon','unban')
$script:ForbiddenPlaceholder = '\{player\}'

# -----------------------------------------------------------------------------
#  BUILT-IN GAME TEMPLATES
#
#  Fields reference:
#    Prefix          - short Discord command prefix (e.g. PZ, MC, HY)
#    StopMethod      - processKill | processName | stdin
#                        processKill  = kill the launcher PID + full child tree
#                        processName  = kill all processes matching ProcessName
#                        stdin        = send StdinStopCommand then wait/kill
#    ProcessName     - name of the real game process (used by processName stop
#                      and the fallback PID-tracker in ServerManager)
#    StdinStop       - command sent to stdin when StopMethod = stdin
#    SaveMethod      - none | stdin | http
#    StdinSave       - command sent to stdin for save
#    SaveWaitSeconds - seconds to wait after save before kill
#    RconPort        - default RCON port (0 = not supported)
#    RconPassword    - always blank; user fills via GUI
#    ExeHints        - ordered list of filenames to look for as the launcher
#    MinRamGB        - JVM -Xms value (Java games only)
#    MaxRamGB        - JVM -Xmx value (Java games only)
#    ExtraCommands   - hashtable of additional Discord commands beyond the
#                      four base ones (start/stop/restart/status)
#
#  Every game template MUST have at least:
#    save    - trigger a world save
#    players - list online players (read-only, always safe)
#
#  Games that support RCON should expose players via RCON where possible.
#  Games without RCON use SendCommand (stdin) or Status fallbacks.
# -----------------------------------------------------------------------------
$script:KnownGames = @{

    # -------------------------------------------------------------------------
    #  PROJECT ZOMBOID
    #  - Launched via StartServer64.bat (cmd.exe /c wrapper)
    #  - The real game process is ZuluPlatformx64Architecture
    #  - StopMethod = processKill kills the cmd wrapper and its child tree,
    #    which includes the Zulu JVM.  processName is unreliable here because
    #    multiple Java-based games on the same machine could share a JVM name.
    #  - Save is sent over stdin before the kill.
    #  - RCON is available on port 27015 (Steam RCON protocol).
    # -------------------------------------------------------------------------
    'ProjectZomboid' = @{
        Prefix          = 'PZ'
        StopMethod      = 'processKill'
        ProcessName     = 'ZuluPlatformx64Architecture'
        StdinStop       = ''
        SaveMethod      = 'stdin'
        StdinSave       = 'save'
        SaveWaitSeconds = 20
        RconPort        = 27015
        RconPassword    = ''
        ConfigRoot      = '%USERPROFILE%\Zomboid\Server'
        # PZ log structure (confirmed from live server logs):
        #
        #   %USERPROFILE%\Zomboid\Logs\
        #       <YYYY-MM-DD_HH-MM>\              <- one subfolder per server session
        #           <date>_DebugLog-server.txt    <- startup, save, crash, world events  ** PRIMARY **
        #           <date>_user.txt               <- player join, disconnect, auth fail  ** PRIMARY **
        #           <date>_chat.txt               <- chat messages and faction events
        #           <date>_cmd.txt                <- admin commands run on the server
        #           <date>_pvp.txt                <- combat and safety zone events
        #           <date>_PerkLog.txt            <- player skill snapshot on login
        #           <date>_ClientActionLog.txt    <- per-player vehicle/item actions (too noisy)
        #
        # LogStrategy = 'PZSessionFolder':
        #   The log reader finds %USERPROFILE%\Zomboid\Logs, gets the newest subfolder
        #   (subfolders are named by date so sort descending = newest first), then tails
        #   *_DebugLog-server.txt AND *_user.txt from inside that folder.
        #   When a new session starts the folder name changes; the reader must reopen.
        #   The install folder Logs\ only contains useless Steam overlay logs - ignore it.
        LogStrategy        = 'PZSessionFolder'
        ServerLogRoot      = ''   # populated at profile-gen time: %USERPROFILE%\Zomboid\Logs
        ServerLogSubDir    = ''   # not fixed - reader resolves newest session subfolder
        ServerLogFile      = ''   # not fixed - reader uses wildcard patterns below
        ServerLogFileDebug = '*_DebugLog-server.txt'
        ServerLogFileUser  = '*_user.txt'
        ServerLogFileChat  = '*_chat.txt'
        ExeHints        = @(
            'StartServer64.bat',
            'StartServer64_nosteam.bat',
            'StartServer32.bat',
            'ProjectZomboidServer.bat',
            'start-server.bat'
        )
        MinRamGB        = 2
        MaxRamGB        = 16
        ExtraCommands   = @{
            save    = @{ Type = 'SendCommand'; Command = 'save'    }
            players = @{ Type = 'SendCommand'; Command = 'players' }
        }
    }

    # -------------------------------------------------------------------------
    #  HYTALE
    #  - Pre-release / private beta server. Launcher is a JAR file.
    #  - No RCON protocol; all control is via stdin.
    #  - Requires AOT cache and asset zip passed via launch args.
    #  - Requires --backup-dir for the backup command to work.
    # -------------------------------------------------------------------------
    'Hytale' = @{
        Prefix          = 'HY'
        StopMethod      = 'processKill'
        ProcessName     = 'java'

        # stdin control
        StdinStop       = 'stop'
        SaveMethod      = 'stdin'
        StdinSave       = 'backup'
        SaveWaitSeconds = 60

        # No RCON
        RconPort        = 0
        RconPassword    = ''

        # Logging
        LogStrategy     = 'NewestFile'
        ServerLogSubDir = 'logs'
        ServerLogFile   = '*_server.log'

        # Executable + memory
        ExeHints        = @('HytaleServer.jar')
        Executable      = 'HytaleServer.jar'
        MinRamGB        = 4
        MaxRamGB        = 8

        # NEW REQUIRED FIELDS
        AOTCache        = 'HytaleServer.aot'
        AssetFile       = 'Assets.zip'
        BackupDir       = 'backup'

        # Extra commands
        ExtraCommands   = @{
            save    = @{ Type = 'SendCommand'; Command = 'backup' }
            players = @{ Type = 'SendCommand'; Command = 'who'    }
        }
    }

    # -------------------------------------------------------------------------
    #  MINECRAFT  (Vanilla / Paper / Forge / Spigot)
    #  - Launched as a JAR directly via java.exe.
    #  - stdin stop ('stop') is clean and saves the world automatically.
    #  - Explicit save-all before stop is still sent for safety.
    #  - RCON on 25575 is standard; players/list routed via RCON.
    # -------------------------------------------------------------------------
    'Minecraft' = @{
        Prefix          = 'MC'
        StopMethod      = 'stdin'
        ProcessName     = 'java'
        StdinStop       = 'stop'
        SaveMethod      = 'stdin'
        StdinSave       = 'save-all'
        SaveWaitSeconds = 15
        RconPort        = 25575
        RconPassword    = ''
        ServerLogSubDir = 'logs'
        ServerLogFile   = 'latest.log'
        ExeHints        = @(
            'server.jar',
            'minecraft_server*.jar',
            'forge*.jar',
            'paper*.jar',
            'spigot*.jar',
            'purpur*.jar'
        )
        MinRamGB        = 2
        MaxRamGB        = 8
        ExtraCommands   = @{
            save    = @{ Type = 'SendCommand'; Command = 'save-all' }
            players = @{ Type = 'Rcon';        Command = 'list'     }
        }
    }

    # -------------------------------------------------------------------------
    #  PALWORLD (2026 Modern Profile)
    #  - Launches via PalServer.exe (UE5 dedicated server)
    #  - REST API required for save/status/players/info
    #  - REST password auto-loaded from PalWorldSettings.ini (RESTAPIKey)
    #  - Stop = processKill (no stdin pipe)
    # -------------------------------------------------------------------------
    'Palworld' = @{
        Prefix          = 'PW'

        # --- PROCESS / EXECUTABLE ---
        Executable      = 'PalServer.exe'
        ProcessName     = 'PalServer-Win64-Shipping'
        FolderPath      = ''

        # LaunchArgs are UI‑generated; this is only a fallback
        LaunchArgs      = '-port=8211 -players=32 -useperfthreads -NoAsyncLoadingThread -UseMultithreadForDS -restapiport=8212'

        # --- REST API (AUTO‑LOADED FROM INI) ---
        RestEnabled     = $true
        RestHost        = '127.0.0.1'
        RestPort        = 8212
        RestPassword    = ''          # Manager fills this from INI: RESTAPIKey=
        RestProtocol    = 'http'

        # --- STOP / SAVE BEHAVIOR ---
        StopMethod      = 'processKill'
        SaveMethod      = 'rest'
        SaveWaitSeconds = 10

        # --- LOGGING ---
        LogStrategy     = 'NewestFile'
        ServerLogSubDir = 'Pal\Saved\Logs'
        ServerLogFile   = '*.log'

        # --- RCON (OPTIONAL FALLBACK) ---
        RconHost        = '127.0.0.1'
        RconPort        = 25575
        RconPassword    = ''          # Auto-loaded from INI if present

        # --- RAM LIMITS ---
        MinRamGB        = 4
        MaxRamGB        = 16

        # --- REST COMMANDS ---
        ExtraCommands   = @{
            save    = @{ Type = 'Rest'; Endpoint = '/v1/api/save';        Method = 'POST' }
            players = @{ Type = 'Rest'; Endpoint = '/v1/api/players';     Method = 'GET'  }
            info    = @{ Type = 'Rest'; Endpoint = '/v1/api/server/info'; Method = 'GET'  }
            status  = @{ Type = 'Rest'; Endpoint = '/v1/api/server/state';Method = 'GET'  }
        }

        # --- EXECUTABLE HINTS ---
        ExeHints = @(
            'PalServer.exe',
            'PalServer-Win64-Shipping.exe',
            'PalServer-Win64-Test-Cmd.exe'
        )
    }

    # -------------------------------------------------------------------------
    #  7 DAYS TO DIE
    #  - Launched via 7DaysToDieServer.exe or a bat wrapper.
    #  - No stdin pipe in headless mode; save = 'saveworld' via stdin window.
    #  - RCON on 25575 (Telnet-based, but mapped via RCON in this system).
    #  - listplayers is a safe read-only RCON command.
    # -------------------------------------------------------------------------
    '7DaysToDie' = @{
        Prefix          = 'DZ'
        StopMethod      = 'processKill'
        ProcessName     = '7DaysToDieServer'
        StdinStop       = ''
        SaveMethod      = 'stdin'
        StdinSave       = 'saveworld'
        SaveWaitSeconds = 20
        RconPort        = 25575
        RconPassword    = ''
        ServerLogSubDir = '7DaysToDieServer_Data'
        ServerLogFile   = 'output_log.txt'
        ExeHints        = @(
            '7DaysToDieServer.exe',
            'startserver.bat',
            'startdedicated.bat'
        )
        MinRamGB        = 2
        MaxRamGB        = 8
        ExtraCommands   = @{
            save    = @{ Type = 'SendCommand'; Command = 'saveworld'   }
            players = @{ Type = 'Rcon';        Command = 'listplayers' }
        }
    }

    # -------------------------------------------------------------------------
    #  VALHEIM
    #  - Launched via valheim_server.exe or a bat wrapper.
    #  - No RCON support in vanilla; process kill is the only stop method.
    #  - The game auto-saves on a timer; no explicit save command available.
    #  - Players: no native command; Status fallback reports running state.
    # -------------------------------------------------------------------------
    'Valheim' = @{
        Prefix          = 'VH'
        StopMethod      = 'processKill'
        ProcessName     = 'valheim_server'
        StdinStop       = ''
        SaveMethod      = 'none'
        StdinSave       = ''
        SaveWaitSeconds = 15
        RconPort        = 0
        RconPassword    = ''
        # Valheim logs go to %AppData%\..\LocalLow\IronGate\Valheim\, NOT the install folder.
        # LogStrategy = 'ValheimUserFolder' tells the reader to resolve this path at runtime.
        # Actual path: C:\Users\<n>\AppData\LocalLow\IronGate\Valheim\valheim_server.log
        LogStrategy     = 'ValheimUserFolder'
        ServerLogRoot   = ''   # resolved at profile-gen time to %AppData%\..\LocalLow\IronGate\Valheim
        ServerLogSubDir = ''
        ServerLogFile   = 'valheim_server.log'
        ExeHints        = @(
            'valheim_server.exe',
            'start_headless_server.bat',
            'start_server.bat'
        )
        MinRamGB        = 2
        MaxRamGB        = 8
        ExtraCommands   = @{
            save    = @{ Type = 'Status'; Command = '' }
            players = @{ Type = 'Status'; Command = '' }
        }
    }

    # -------------------------------------------------------------------------
    #  TERRARIA  (TShock or vanilla dedicated server)
    #  - Launched via TerrariaServer.exe or a bat.
    #  - stdin 'save' and 'exit' are supported in both vanilla and TShock.
    #  - TShock exposes a REST API but we default to stdin for compatibility.
    #  - 'playing' lists online players in vanilla; TShock uses the same.
    # -------------------------------------------------------------------------
    'Terraria' = @{
        Prefix          = 'TR'
        StopMethod      = 'stdin'
        ProcessName     = 'TerrariaServer'
        StdinStop       = 'exit'
        SaveMethod      = 'stdin'
        StdinSave       = 'save'
        SaveWaitSeconds = 15
        RconPort        = 0
        RconPassword    = ''
        # Vanilla TerrariaServer.exe has no log file - output is stdout only.
        # TShock (the popular server mod) writes to logs\server.log inside the install folder.
        # If running vanilla, ServerLogPath will exist but file won't be created - that is expected.
        ServerLogSubDir = 'logs'
        ServerLogFile   = 'server.log'
        ExeHints        = @(
            'TerrariaServer.exe',
            'TerrariaServer.bat',
            'start-server.bat'
        )
        MinRamGB        = 1
        MaxRamGB        = 4
        ExtraCommands   = @{
            save    = @{ Type = 'SendCommand'; Command = 'save'    }
            players = @{ Type = 'SendCommand'; Command = 'playing' }
        }
    }

    # -------------------------------------------------------------------------
    #  RUST
    #  - Launched via RustDedicated.exe.
    #  - RCON on 28016 (WebSocket RCON - not supported by our TCP RCON client).
    #    We use stdin fallback for save/players.
    #  - StopMethod = processKill; 'quit' can also be sent via stdin.
    # -------------------------------------------------------------------------
    'Rust' = @{
        Prefix          = 'RS'
        StopMethod      = 'processKill'
        ProcessName     = 'RustDedicated'
        StdinStop       = ''
        SaveMethod      = 'stdin'
        StdinSave       = 'server.save'
        SaveWaitSeconds = 15
        RconPort        = 0
        RconPassword    = ''
        # Rust does not write a log file by default.
        # The -logFile launch argument tells Unity where to write output.
        # We default LaunchArgs to include -logFile so the file exists after first start.
        # Actual path: <FolderPath>\logs\RustDedicated.log  (created by Unity on start)
        ServerLogSubDir = 'logs'
        ServerLogFile   = 'RustDedicated.log'
        # Default LaunchArgs injects -logFile so logs\RustDedicated.log is created automatically.
        # The GUI pre-populates LaunchArgs with this value when generating a Rust profile.
        DefaultLaunchArgs = '-logFile logs/RustDedicated.log'
        ExeHints        = @(
            'RustDedicated.exe',
            'start_rust.bat',
            'start.bat'
        )
        MinRamGB        = 4
        MaxRamGB        = 16
        ExtraCommands   = @{
            save    = @{ Type = 'SendCommand'; Command = 'server.save'       }
            players = @{ Type = 'SendCommand'; Command = 'playerlist'        }
        }
    }

    # -------------------------------------------------------------------------
    #  ARK: SURVIVAL EVOLVED / ASA
    #  - Launched via ShooterGameServer.exe or ArkAscendedServer.exe.
    #  - RCON on 32330 is standard.
    #  - 'saveworld' and 'listplayers' are valid RCON commands.
    # -------------------------------------------------------------------------
    'Ark' = @{
        Prefix          = 'ARK'
        StopMethod      = 'processKill'
        ProcessName     = 'ShooterGameServer'
        StdinStop       = ''
        SaveMethod      = 'none'
        StdinSave       = ''
        SaveWaitSeconds = 20
        RconPort        = 32330
        RconPassword    = ''
        ServerLogSubDir = 'ShooterGame\Saved\Logs'
        ServerLogFile   = 'ShooterGame.log'
        ExeHints        = @(
            'ShooterGameServer.exe',
            'ArkAscendedServer.exe',
            'start_ark.bat',
            'start.bat'
        )
        MinRamGB        = 4
        MaxRamGB        = 16
        ExtraCommands   = @{
            save    = @{ Type = 'Rcon'; Command = 'saveworld'   }
            players = @{ Type = 'Rcon'; Command = 'listplayers' }
        }
    }

    # -------------------------------------------------------------------------
    #  SATISFACTORY
    #  - Launched via FactoryServer.exe (Unreal Engine).
    #  - No RCON; no stdin pipe in headless mode.
    #  - Game auto-saves; no explicit save command available.
    # -------------------------------------------------------------------------
    'Satisfactory' = @{
        Prefix          = 'SF'
        StopMethod      = 'processKill'
        ProcessName     = 'FactoryServer-Win64-Shipping'
        StdinStop       = ''
        SaveMethod      = 'none'
        StdinSave       = ''
        SaveWaitSeconds = 15
        RconPort        = 0
        RconPassword    = ''
        # Primary log: FactoryGame\Saved\Logs\FactoryGame.log
        # Some dedicated server builds write Server.log instead.
        # LogStrategy = 'NewestFile' opens the newest *.log as a fallback.
        LogStrategy     = 'NewestFile'
        ServerLogSubDir = 'FactoryGame\Saved\Logs'
        ServerLogFile   = 'FactoryGame.log'
        ExeHints        = @(
            'FactoryServer.exe',
            'Start_FactoryServer.bat',
            'start.bat'
        )
        MinRamGB        = 4
        MaxRamGB        = 16
        ExtraCommands   = @{
            save    = @{ Type = 'Status'; Command = '' }
            players = @{ Type = 'Status'; Command = '' }
        }
    }

}

# =============================================================================
#  HELPER: Derive a short prefix from any game name
#  "Project Zomboid" -> "PZ"
#  "7DaysToDie"      -> "7DT" (first char of each word)
#  "Hytale"          -> "HY"
# =============================================================================
function Get-AutoPrefix {
    param([string]$GameName)
    $words = ($GameName -split '\s+') | Where-Object { $_ -ne '' }
    if ($words.Count -ge 2) {
        return ($words | ForEach-Object { $_[0].ToString().ToUpper() }) -join ''
    }
    # Single word: take first 2-3 uppercase chars; fall back to first 3 chars
    $upper = ($GameName.ToCharArray() | Where-Object { [char]::IsUpper($_) }) -join ''
    if ($upper.Length -ge 2) { return $upper.Substring(0, [Math]::Min(3, $upper.Length)) }
    return $GameName.Substring(0, [Math]::Min(3, $GameName.Length)).ToUpper()
}

# =============================================================================
#  HELPER: Detect the best executable / launch file in a folder
# =============================================================================
function Find-ServerExecutable {
    param(
        [string]$FolderPath,
        [string[]]$Hints = @()
    )
    if (-not (Test-Path $FolderPath)) { return $null }

    # Try hints first (in order given by template)
    foreach ($hint in $Hints) {
        $found = Get-ChildItem -Path $FolderPath -Filter $hint -ErrorAction SilentlyContinue |
                 Select-Object -First 1
        if ($found) { return $found.Name }
    }

    # Fallback: .bat before .exe (Java/script launchers are usually .bat)
    $bat = Get-ChildItem -Path $FolderPath -Filter '*.bat' -ErrorAction SilentlyContinue |
           Where-Object { $_.Name -notmatch 'install|update|uninstall|setup' } |
           Select-Object -First 1
    if ($bat) { return $bat.Name }

    $exe = Get-ChildItem -Path $FolderPath -Filter '*.exe' -ErrorAction SilentlyContinue |
           Where-Object { $_.Name -notmatch 'unins|setup|install|update|redist' } |
           Select-Object -First 1
    if ($exe) { return $exe.Name }

    return $null
}

# =============================================================================
#  HELPER: Match a folder/game name against the KnownGames dictionary
#  Returns the dictionary key (e.g. 'ProjectZomboid') or $null
# =============================================================================
function Resolve-KnownGame {
    param([string]$FolderName)
    $lower      = $FolderName.ToLower()
    $lowerNoSpc = $lower -replace '[\s\-_]+', ''

    # Additional alias table so variant spellings hit the right template
    $aliases = @{
        'zomboid'       = 'ProjectZomboid'
        'projectzomboid'= 'ProjectZomboid'
        'hytale'        = 'Hytale'
        'minecraft'     = 'Minecraft'
        'palworld'      = 'Palworld'
        '7days'         = '7DaysToDie'
        '7daystodieserver' = '7DaysToDie'
        'sevendays'     = '7DaysToDie'
        'valheim'       = 'Valheim'
        'terraria'      = 'Terraria'
        'tshock'        = 'Terraria'
        'rust'          = 'Rust'
        'rustdedicated' = 'Rust'
        'ark'           = 'Ark'
        'arksurvival'   = 'Ark'
        'arkascended'   = 'Ark'
        'satisfactory'  = 'Satisfactory'
        'factoryserver' = 'Satisfactory'
    }

    # Check alias table first (fastest, most specific)
    foreach ($alias in $aliases.Keys) {
        if ($lowerNoSpc -match $alias -or $lower -match $alias) {
            return $aliases[$alias]
        }
    }

    # Fall back to checking all known game keys and their prefixes
    foreach ($key in $script:KnownGames.Keys) {
        $keyLower    = $key.ToLower()
        $prefixLower = $script:KnownGames[$key].Prefix.ToLower()
        if ($lower      -match $keyLower    -or
            $lowerNoSpc -match $keyLower    -or
            $lower      -match $prefixLower -or
            $lowerNoSpc -match $prefixLower) {
            return $key
        }
    }

    return $null
}

# =============================================================================
#  Auto-generate a profile object from a folder path (does NOT save to disk)
# =============================================================================
function New-GameProfile {
    param(
        [string]$FolderPath,
        [string]$GameName   = '',
        [string]$Prefix     = '',
        [string]$Executable = '',
        [string]$LaunchArgs = ''
    )

    if (-not (Test-Path $FolderPath)) {
        throw "Folder not found: $FolderPath"
    }

    # Derive game name from folder name if not supplied
    if (-not $GameName) {
        $GameName = (Split-Path $FolderPath -Leaf)
    }

    # Try to match a known game template
    $knownKey = Resolve-KnownGame -FolderName $GameName
    $template = if ($knownKey) { $script:KnownGames[$knownKey] } else { $null }

    Write-Host "[ProfileManager] New profile: '$GameName' -> template: $(if ($knownKey) { $knownKey } else { 'generic' })" -ForegroundColor Cyan

    # --- Prefix ---
    if (-not $Prefix) {
        $Prefix = if ($template) { $template.Prefix } else { Get-AutoPrefix -GameName $GameName }
    }

    # --- Executable ---
    if (-not $Executable) {
        $hints      = if ($template -and $template.ExeHints) { $template.ExeHints } else { @() }
        $Executable = Find-ServerExecutable -FolderPath $FolderPath -Hints $hints
        if (-not $Executable) { $Executable = 'server.exe' }
    }

    # --- Stop / Save methods ---
    $stopMethod = if ($template) { $template.StopMethod  } else { 'processKill' }
    $saveMethod = if ($template) { $template.SaveMethod  } else { 'none' }
    $stdinStop  = if ($template -and $template.StdinStop) { $template.StdinStop } else { '' }
    $stdinSave  = if ($template -and $template.StdinSave) { $template.StdinSave } else { '' }

    # --- Base commands (present for every game) ---
    $commands = [ordered]@{
        start   = @{ Type = 'Start'   }
        stop    = @{ Type = 'Stop'    }
        restart = @{ Type = 'Restart' }
        status  = @{ Type = 'Status'  }
    }

    # --- Merge game-specific extra commands ---
    if ($template -and $template.ExtraCommands -and $template.ExtraCommands.Count -gt 0) {
        foreach ($cmd in $template.ExtraCommands.Keys) {
            $commands[$cmd] = $template.ExtraCommands[$cmd]
        }
    } else {
        # Generic fallback: safe stdin save and players status
        $commands['save']    = @{ Type = 'Status'; Command = '' }
        $commands['players'] = @{ Type = 'Status'; Command = '' }
    }

    # --- Numeric / string fields ---
    $processName  = if ($template -and $template.ProcessName)    { $template.ProcessName    } else { '' }
    $minRamGB     = if ($template -and $template.MinRamGB)       { $template.MinRamGB       } else { 2  }
    $maxRamGB     = if ($template -and $template.MaxRamGB)       { $template.MaxRamGB       } else { 4  }
    $saveWaitSecs = if ($template -and $template.SaveWaitSeconds) { $template.SaveWaitSeconds } else { 15 }
    $rconPort     = if ($template -and $template.RconPort)       { $template.RconPort       } else { 0  }
    $configRoot   = if ($template -and $template.ConfigRoot)     {
        [Environment]::ExpandEnvironmentVariables("$($template.ConfigRoot)")
    } else { '' }
    if ([string]::IsNullOrWhiteSpace($configRoot)) { $configRoot = $FolderPath }

    # Build ServerLogPath from template sub-dir + filename, rooted in FolderPath.
    # For PZ the log lives outside the install folder entirely, so we handle it specially.
    $serverLogPath  = ''
    $serverLogRoot  = ''
    $logStrategy    = if ($template -and $template.LogStrategy) { $template.LogStrategy } else { 'SingleFile' }

    switch ($logStrategy) {

        'PZSessionFolder' {
            # Root = %USERPROFILE%\Zomboid\Logs
            # Reader finds newest date-named subfolder, tails *_DebugLog-server.txt + *_user.txt
            $serverLogRoot = Join-Path $env:USERPROFILE 'Zomboid\Logs'
            $serverLogPath = ''
        }

        'ValheimUserFolder' {
            # Root = %AppData%\..\LocalLow\IronGate\Valheim
            # Valheim writes outside the install folder entirely
            $localLow      = Join-Path $env:APPDATA '..\..\..' | Resolve-Path -ErrorAction SilentlyContinue
            if ($localLow) {
                $serverLogRoot = Join-Path $localLow 'LocalLow\IronGate\Valheim'
            } else {
                $serverLogRoot = "C:\Users\$env:USERNAME\AppData\LocalLow\IronGate\Valheim"
            }
            $serverLogPath = Join-Path $serverLogRoot 'valheim_server.log'
        }

        'NewestFile' {
            # Log filename is dynamic (timestamped). Reader opens newest *.log in the folder.
            if ($template -and $template.ServerLogSubDir -and $template.ServerLogSubDir -ne '') {
                $serverLogRoot = Join-Path $FolderPath $template.ServerLogSubDir
            } else {
                $serverLogRoot = $FolderPath
            }
            # Store the known preferred filename - reader falls back to newest *.log if missing
            $preferredFile = if ($template -and $template.ServerLogFile -and
                                 $template.ServerLogFile -ne '*.log') { $template.ServerLogFile } else { '' }
            $serverLogPath = if ($preferredFile) { Join-Path $serverLogRoot $preferredFile } else { '' }
        }

        default {
            # SingleFile - fixed path relative to install folder
            if ($template -and $template.ServerLogFile -and $template.ServerLogFile -ne '') {
                if ($template.ServerLogSubDir -and $template.ServerLogSubDir -ne '') {
                    $serverLogPath = Join-Path $FolderPath (Join-Path $template.ServerLogSubDir $template.ServerLogFile)
                } else {
                    $serverLogPath = Join-Path $FolderPath $template.ServerLogFile
                }
                $serverLogRoot = Split-Path $serverLogPath -Parent
            }
        }
    }

    # Inject DefaultLaunchArgs from template if LaunchArgs not already set
    if ([string]::IsNullOrWhiteSpace($LaunchArgs) -and
        $template -and $template.DefaultLaunchArgs -and $template.DefaultLaunchArgs -ne '') {
        $LaunchArgs = $template.DefaultLaunchArgs
    }

    # Build a human-readable note for the GUI
    $logPathNote = switch ($logStrategy) {
        'PZSessionFolder'    { "PZ logs are under: $serverLogRoot`nReader auto-finds newest session subfolder and tails *_DebugLog-server.txt + *_user.txt." }
        'ValheimUserFolder'  { "Valheim logs are at: $serverLogPath`n(Outside install folder - written by the game engine)" }
        'NewestFile'         { "Logs folder: $serverLogRoot`nReader opens the newest *.log file found there." }
        default              { if ([string]::IsNullOrWhiteSpace($serverLogPath)) { "Log path could not be auto-detected. Set ServerLogPath manually." } else { '' } }
    }

    $profile = [ordered]@{
        GameName            = $GameName
        Prefix              = $Prefix.ToUpper()
        Executable          = $Executable
        FolderPath          = $FolderPath
        LaunchArgs          = $LaunchArgs
        StopMethod          = $stopMethod
        ProcessName         = $processName
        StdinStopCommand    = $stdinStop
        SaveMethod          = $saveMethod
        StdinSaveCommand    = $stdinSave
        SaveWaitSeconds     = $saveWaitSecs
        RconHost            = '127.0.0.1'
        RconPort            = $rconPort
        RconPassword        = ''
        MinRamGB            = $minRamGB
        MaxRamGB            = $maxRamGB
        EnableAutoRestart   = $true
        RestartDelaySeconds = 10
        MaxRestartsPerHour  = 5
        LogStrategy         = $logStrategy
        ServerLogRoot       = $serverLogRoot
        ServerLogPath       = $serverLogPath
        ServerLogNote       = $logPathNote
        ConfigRoot          = $configRoot
        Commands            = $commands
        AOTCache            = if ($template -and $template.AOTCache)   { $template.AOTCache }   else { '' }
        AssetFile           = if ($template -and $template.AssetFile)  { $template.AssetFile }  else { '' }
        BackupDir           = if ($template -and $template.BackupDir)  { $template.BackupDir }  else { '' }
        # --- REST API ---
        RestEnabled     = if ($template) { $template.RestEnabled } else { $false }
        RestHost        = if ($template) { $template.RestHost    } else { '' }
        RestPort        = if ($template) { $template.RestPort    } else { 0  }
        RestPassword    = ''   # always blank; filled from INI at runtime
        RestProtocol    = if ($template) { $template.RestProtocol } else { 'http' }
        RestPollOnlyWhenRunning = $true

        # --- LOGGING (extra fields) ---
        ServerLogSubDir = if ($template -and $template.ServerLogSubDir) { $template.ServerLogSubDir } else { '' }
        ServerLogFile   = if ($template -and $template.ServerLogFile)   { $template.ServerLogFile   } else { '' }

        # --- EXTRA COMMANDS (already merged into Commands, but store raw template too) ---
        ExtraCommands   = if ($template -and $template.ExtraCommands) { $template.ExtraCommands } else { @{} }

        # --- EXECUTABLE HINTS ---
        ExeHints        = if ($template -and $template.ExeHints) { $template.ExeHints } else { @() }

    }

    Write-Host "[ProfileManager] Profile created: [$($profile.Prefix)] $($profile.GameName)" -ForegroundColor Green
    return $profile
}

# =============================================================================
#  Safety validator - raises if a single command definition is unsafe
# =============================================================================
function Test-CommandSafety {
    param([string]$CmdName, [hashtable]$CmdDef)

    # Build a combined text string of all values to scan
    $allValues = @($CmdName)
    if ($CmdDef -is [System.Collections.IDictionary]) {
        foreach ($v in $CmdDef.Values) {
            if ($null -ne $v) { $allValues += "$v" }
        }
    }
    $allText = ($allValues -join ' ').ToLower()

    # Reject {player} placeholder
    if ($allText -match $script:ForbiddenPlaceholder) {
        throw "UNSAFE: Command '$CmdName' contains a {player} placeholder."
    }

    # Reject forbidden keywords (whole-word match to avoid false positives)
    foreach ($kw in $script:ForbiddenKeywords) {
        if ($allText -match "\b$([regex]::Escape($kw))\b") {
            throw "UNSAFE: Command '$CmdName' contains forbidden keyword '$kw'."
        }
    }

    return $true
}

# =============================================================================
#  Validate an entire profile's command set
# =============================================================================
function Test-ProfileSafety {
    param([hashtable]$Profile)

    if ($null -eq $Profile -or $null -eq $Profile.Commands) { return $true }

    $errors = [System.Collections.Generic.List[string]]::new()

    foreach ($cmdName in $Profile.Commands.Keys) {
        $def = $Profile.Commands[$cmdName]
        # ConvertTo-Hashtable may return an ordered dict; cast for Test-CommandSafety
        $defHt = if ($def -is [System.Collections.IDictionary]) { $def } else { @{} }
        try   { Test-CommandSafety -CmdName $cmdName -CmdDef $defHt | Out-Null }
        catch { $errors.Add($_.Exception.Message) }
    }

    if ($errors.Count -gt 0) {
        throw "Profile '$($Profile.GameName)' failed safety check:`n$($errors -join "`n")"
    }
    return $true
}

# =============================================================================
#  Save a profile hashtable to the Profiles folder as JSON
# =============================================================================
function Save-GameProfile {
    param(
        [hashtable]$Profile,
        [string]$ProfilesDir = '.\Profiles'
    )

    # Safety gate
    Test-ProfileSafety -Profile $Profile | Out-Null

    # Resolve to absolute path
    if (-not [System.IO.Path]::IsPathRooted($ProfilesDir)) {
        $ProfilesDir = Join-Path (Get-Location) $ProfilesDir
    }

    if (-not (Test-Path $ProfilesDir)) {
        New-Item -ItemType Directory -Path $ProfilesDir -Force | Out-Null
    }

    # Sanitize filename
    $safeName = ($Profile.GameName -replace '[\\/:*?"<>|]', '_') -replace '\s+', '_'
    $outPath  = Join-Path $ProfilesDir "$safeName.json"

    try {
        $Profile | ConvertTo-Json -Depth 10 | Set-Content -Path $outPath -Encoding UTF8 -Force
        Write-Host "[ProfileManager] Saved profile -> $outPath" -ForegroundColor Green
    } catch {
        Write-Host "[ProfileManager] ERROR saving profile: $_" -ForegroundColor Red
        throw $_
    }

    return $outPath
}

# =============================================================================
#  Load all profiles from disk and return a prefix-keyed hashtable
# =============================================================================
function Get-AllProfiles {
    param([string]$ProfilesDir = '.\Profiles')
    $profiles = @{}
    if (-not (Test-Path $ProfilesDir)) { return $profiles }

    Get-ChildItem -Path $ProfilesDir -Filter '*.json' -ErrorAction SilentlyContinue |
    ForEach-Object {
        try {
            $raw = Get-Content -Path $_.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
            $p   = ConvertTo-Hashtable -Object $raw
            Test-ProfileSafety -Profile $p | Out-Null
            $profiles[$p.Prefix.ToUpper()] = $p
            Write-Host "[ProfileManager] Loaded: [$($p.Prefix)] $($p.GameName)" -ForegroundColor Green
        } catch {
            Write-Host "[ProfileManager] WARN - Skipped $($_.Name): $_" -ForegroundColor Yellow
        }
    }
    return $profiles
}

# =============================================================================
#  Helper: Recursively convert PSCustomObject (from ConvertFrom-Json) to
#  ordered hashtable so all profile access is dictionary-style throughout.
# =============================================================================
function ConvertTo-Hashtable {
    param([object]$Object)

    if ($null -eq $Object) { return $null }

    if ($Object -is [System.Collections.IDictionary]) {
        $hash = @{}
        foreach ($key in $Object.Keys) {
            $hash[$key] = ConvertTo-Hashtable -Object $Object[$key]
        }
        return $hash
    }
    elseif ($Object -is [psobject]) {
        # Convert PSCustomObject to hashtable so downstream code can use dictionary access.
        $hash = [ordered]@{}
        foreach ($prop in $Object.PSObject.Properties) {
            $hash[$prop.Name] = ConvertTo-Hashtable -Object $prop.Value
        }
        return $hash
    }
    elseif ($Object -is [System.Collections.IEnumerable] -and
            -not ($Object -is [string])) {
        return @($Object | ForEach-Object { ConvertTo-Hashtable -Object $_ })
    }
    else {
        return $Object
    }
}

# =============================================================================
#  Return the list of all known game keys (used by GUI to show template info)
# =============================================================================
function Get-KnownGameKeys {
    return @($script:KnownGames.Keys)
}

Export-ModuleMember -Function `
    New-GameProfile, Save-GameProfile, Get-AllProfiles, `
    Test-ProfileSafety, Test-CommandSafety, `
    Get-AutoPrefix, Find-ServerExecutable, `
    Resolve-KnownGame, ConvertTo-Hashtable, Get-KnownGameKeys
