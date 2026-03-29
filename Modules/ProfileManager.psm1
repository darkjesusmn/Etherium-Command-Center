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

function _GetWindowsLocalLowPath {
    try {
        $localAppData = [Environment]::GetFolderPath('LocalApplicationData')
        if (-not [string]::IsNullOrWhiteSpace($localAppData)) {
            $appDataRoot = Split-Path -Parent $localAppData
            if (-not [string]::IsNullOrWhiteSpace($appDataRoot)) {
                return (Join-Path $appDataRoot 'LocalLow')
            }
        }
    } catch { }

    if (-not [string]::IsNullOrWhiteSpace($env:USERPROFILE)) {
        return (Join-Path $env:USERPROFILE 'AppData\LocalLow')
    }

    return ''
}

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
        StartupTimeoutSeconds = 600
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
    #  7 DAYS TO DIE  (Alpha 21 / 1.x dedicated server)
    #
    #  Launch details:
    #    - 7DaysToDieServer.exe is launched directly with -dedicated and a
    #      -configfile pointing at serverconfig.xml in the install folder.
    #    - RedirectStandardInput is enabled so stdin commands work correctly.
    #    - The Unity engine writes timestamped logs to the install folder root:
    #      output_log_dedi__YYYY-MM-DD__HH-MM-SS.txt
    #    - ServerManager expands the $LOGFILE token in LaunchArgs at launch time.
    #    - SteamAppId 251570 is injected into the process environment.
    #
    #  Stop / Save:
    #    - StopMethod = 'stdin' sends 'shutdown' to the console, waits 25s,
    #      then force-kills if needed.  processKill alone can corrupt saves.
    #    - SaveMethod = 'stdin' sends 'saveworld' before the stop sequence.
    #      On large worlds this can take several seconds - hence 25s wait.
    #
    #  Telnet (preferred for save/players/version commands):
    #    - 7DTD uses its own plain-text Telnet protocol on port 8081.
    #      Our Invoke-RconCommand speaks Source RCON (binary), which is
    #      incompatible.  Use Type = 'Telnet' for all console commands.
    #    - Enable in serverconfig.xml:
    #        <property name="TelnetEnabled"  value="true"/>
    #        <property name="TelnetPort"     value="8081"/>
    #        <property name="TelnetPassword" value="yourpassword"/>
    #    - TelnetPassword must not be empty - the server rejects passwordless
    #      connections when TelnetEnabled is true.
    #
    #  Log strategy:
    #    - LogStrategy = 'NewestFile' - log filename is timestamped each run.
    #      GUI tails the newest output_log_dedi__*.txt automatically.
    # -------------------------------------------------------------------------
    '7DaysToDie' = @{
        Prefix          = 'DZ'

        # --- PROCESS / EXECUTABLE ---
        Executable      = '7DaysToDieServer.exe'
        ProcessName     = '7DaysToDieServer'
        HideConsoleWindow = $true
        FolderPath      = ''

        # $LOGFILE is expanded at launch time to a timestamped filename.
        # -dedicated MUST be last per the official startserver.bat.
        LaunchArgs      = '-logfile "$LOGFILE" -quit -batchmode -nographics -configfile=serverconfig.xml -dedicated'

        # SteamAppId injected into process environment before launch
        SteamAppId      = '251570'

        # --- STOP / SAVE BEHAVIOR ---
        StopMethod       = 'stdin'
        StdinStopCommand = 'shutdown'
        SaveMethod       = 'stdin'
        StdinSaveCommand = 'saveworld'
        SaveWaitSeconds  = 25

        # --- RCON (NOT SUPPORTED - 7DTD uses Telnet, not Source RCON) ---
        RconHost        = '127.0.0.1'
        RconPort        = 0
        RconPassword    = ''

        # --- TELNET ---
        TelnetHost      = '127.0.0.1'
        TelnetPort      = 8081
        TelnetPassword  = ''   # must match TelnetPassword in serverconfig.xml

        # --- LOGGING ---
        LogStrategy     = 'NewestFile'
        ServerLogSubDir = ''
        ServerLogFile   = 'output_log_dedi__*.txt'

        # --- RAM (display only for .exe servers) ---
        MinRamGB        = 4
        MaxRamGB        = 12

        # --- EXECUTABLE HINTS ---
        ExeHints        = @(
            '7DaysToDieServer.exe',
            'startserver.bat',
            'startdedicated.bat',
            'start.bat'
        )

        # --- EXTRA COMMANDS ---
        # All routed via Telnet - much more reliable than stdin on 7DTD.
        ExtraCommands   = @{
            save    = @{ Type = 'Telnet'; Command = 'saveworld'   }
            players = @{ Type = 'Telnet'; Command = 'listplayers' }
            version = @{ Type = 'Telnet'; Command = 'version'     }
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
        StopMethod      = 'ProcessKill'
        ProcessName     = 'valheim_server'
        SaveMethod      = 'none'
        SaveWaitSeconds = 15

        LogStrategy     = 'SingleFile'
        ServerLogSubDir = ''
        ServerLogFile   = 'valheim_output.log'

        ExeHints        = @(
            'valheim_server.exe',
            'start_headless_server.bat',
            'start_server.bat'
        )

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
    #  SATISFACTORY  (1.0+)
    #
    #  Control method: HTTPS API on the game port (default 7777).
    #  The API uses self-signed TLS and Bearer token auth.
    #
    #  Setup (one-time, in the in-game Server Manager console tab):
    #    server.GenerateAPIToken
    #  Copy the printed token into the profile's SatisfactoryApiToken field
    #  via the GUI Profile Editor, or set it directly in the JSON.
    #
    #  Alternatively, launch the server with:
    #    -ini:Engine:[SystemSettings]:FG.DedicatedServer.AllowInsecureLocalAccess=1
    #  and leave SatisfactoryApiToken blank — local requests will be
    #  accepted without a token.
    #
    #  Save:    API function 'SaveGame'         (Admin privilege)
    #  Status:  API function 'QueryServerState' (no auth needed)
    #  Players: API function 'QueryServerState' returns numConnectedPlayers
    #  Stop:    API function 'Shutdown'         (Admin privilege) clean save+quit
    #
    #  Because the API handles a clean shutdown (including a final save),
    #  StopMethod is 'SatisfactoryApi' which calls Shutdown before any kill.
    #  SaveMethod is 'SatisfactoryApi' which calls SaveGame.
    #
    #  Log path:    <InstallFolder>\FactoryGame\Saved\Logs\FactoryGame.log
    #               Relative to the install folder - works on any machine.
    #
    #  Config path: %LOCALAPPDATA%\FactoryGame\Saved\Config\Windows
    #               NOTE: configs are stored in the USER profile, NOT the
    #               install folder. %LOCALAPPDATA% expands at profile-gen time
    #               to the correct path on any machine.
    #
    #  Ports (1.1+):  -Port=7777  -ReliablePort=8888
    #  (Ports 15000 and 15777 are no longer used as of 1.0.)
    # -------------------------------------------------------------------------
    'Satisfactory' = @{
        Prefix          = 'SF'

        # --- PROCESS / EXECUTABLE ---
        Executable      = 'FactoryServer.exe'
        ProcessName     = 'FactoryServer-Win64-Shipping'
        FolderPath      = ''

        # Standard launch args for 1.1+.  -log writes FactoryGame.log.
        # -unattended suppresses interactive prompts on a headless box.
        LaunchArgs      = '-Port=7777 -ReliablePort=8888 -log -unattended'

        # --- SATISFACTORY HTTPS API ---
        SatisfactoryApiHost  = '127.0.0.1'
        SatisfactoryApiPort  = 7777
        # Paste the output of server.GenerateAPIToken here.
        # Leave blank if launching with AllowInsecureLocalAccess=1.
        SatisfactoryApiToken = ''

        # --- STOP / SAVE ---
        StopMethod       = 'SatisfactoryApi'
        StdinStopCommand = ''
        SaveMethod       = 'SatisfactoryApi'
        StdinSaveCommand = ''
        SaveWaitSeconds  = 20

        # --- RCON (not supported) ---
        RconHost        = '127.0.0.1'
        RconPort        = 0
        RconPassword    = ''

        # --- LOGGING ---
        # FactoryGame.log lives INSIDE the install folder under FactoryGame\Saved\Logs.
        # ServerLogSubDir is relative to FolderPath — New-GameProfile joins them at
        # profile-gen time, so this works correctly on any machine.
        # LogStrategy NewestFile: FactoryGame.log is always the current-session log.
        LogStrategy     = 'NewestFile'
        ServerLogSubDir = 'FactoryGame\Saved\Logs'
        ServerLogFile   = 'FactoryGame.log'

        # --- CONFIG ---
        # On Windows the dedicated server writes its ini files to TWO locations:
        #
        # Primary (confirmed):  %LOCALAPPDATA%\FactoryGame\Saved\Config\Windows
        #   Game.ini, GameUserSettings.ini, Engine.ini live here.
        #   %LOCALAPPDATA% is expanded at profile-gen time by ExpandEnvironmentVariables
        #   so it resolves correctly on every machine automatically.
        #
        # Secondary (install folder):  FactoryGame\Saved\Config\WindowsServer
        #   Only created after first graceful shutdown. May or may not exist.
        #   The GUI Config Editor checks the profile's FolderPath as a fallback
        #   so these files are accessible even without being listed in ConfigRoot.
        #
        # We point ConfigRoot at the primary AppData location since that is where
        # the real settings are. The secondary path is discovered automatically.
        ConfigRoot      = '%LOCALAPPDATA%\FactoryGame\Saved\Config\Windows'

        # --- MEMORY ---
        MinRamGB        = 8
        MaxRamGB        = 16

        # --- EXECUTABLE HINTS ---
        ExeHints        = @(
            'FactoryServer.exe',
            'Start_FactoryServer.bat',
            'start.bat'
        )

        # --- EXTRA COMMANDS ---
        ExtraCommands   = @{
            save    = @{ Type = 'SatisfactoryApi'; Function = 'SaveGame'         }
            players = @{ Type = 'SatisfactoryApi'; Function = 'QueryServerState' }
            status  = @{ Type = 'SatisfactoryApi'; Function = 'QueryServerState' }
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

    # Strip leading digits/punctuation from each word before taking initials
    # so "7 Days to Die" -> words [Days, to, Die] -> DTD, not 7DT.
    $words = ($GameName -split '\s+') |
             Where-Object { $_ -ne '' } |
             ForEach-Object { $_ -replace '^[^A-Za-z]+', '' } |
             Where-Object { $_ -ne '' }

    if ($words.Count -ge 2) {
        $letters = ($words | Select-Object -First 4 |
                    ForEach-Object { $_[0].ToString().ToUpper() }) -join ''
        if ($letters.Length -ge 2) { return $letters }
    }

    if ($words.Count -eq 1) {
        $word  = $words[0]
        $upper = ($word.ToCharArray() | Where-Object { [char]::IsUpper($_) }) -join ''
        if ($upper.Length -ge 2) { return $upper.Substring(0, [Math]::Min(3, $upper.Length)) }
        return $word.Substring(0, [Math]::Min(3, $word.Length)).ToUpper()
    }

    # Final fallback: strip non-letters and take first 3
    $clean = ($GameName -replace '[^A-Za-z]', '')
    if ($clean.Length -ge 2) { return $clean.Substring(0, [Math]::Min(3, $clean.Length)).ToUpper() }
    return $GameName.Substring(0, [Math]::Min(3, $GameName.Length)).ToUpper()
}

# =============================================================================
#  HELPER: Detect the best executable / launch file in a folder
# =============================================================================
function _GetRelativeExecutablePath {
    param(
        [string]$BaseFolder,
        [string]$TargetPath
    )

    if ([string]::IsNullOrWhiteSpace($BaseFolder) -or [string]::IsNullOrWhiteSpace($TargetPath)) { return '' }

    try {
        $baseFull = [System.IO.Path]::GetFullPath($BaseFolder).TrimEnd('\')
        $targetFull = [System.IO.Path]::GetFullPath($TargetPath)
        $needle = "$baseFull\"
        if ($targetFull.StartsWith($needle, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $targetFull.Substring($needle.Length)
        }
        if ($targetFull.Equals($baseFull, [System.StringComparison]::OrdinalIgnoreCase)) {
            return [System.IO.Path]::GetFileName($targetFull)
        }
    } catch { }

    try { return [System.IO.Path]::GetFileName($TargetPath) } catch { return '' }
}

function _GetExecutableCandidateScore {
    param(
        [System.IO.FileInfo]$File,
        [string]$BaseFolder,
        [string]$Hint = '',
        [int]$HintOrder = 0
    )

    if ($null -eq $File) { return [int]::MinValue }

    $relativePath = _GetRelativeExecutablePath -BaseFolder $BaseFolder -TargetPath $File.FullName
    $depth = 0
    if (-not [string]::IsNullOrWhiteSpace($relativePath)) {
        $depth = @($relativePath -split '[\\/]' | Where-Object { $_ -ne '' }).Count - 1
    }

    $score = 0
    $nameLower = $File.Name.ToLowerInvariant()
    $relativeLower = if ($relativePath) { $relativePath.ToLowerInvariant() } else { '' }
    $dirLower = if ($File.DirectoryName) { $File.DirectoryName.ToLowerInvariant() } else { '' }

    if (-not [string]::IsNullOrWhiteSpace($Hint)) {
        $score += 5000 - ($HintOrder * 100)
        if ($relativePath -ieq $Hint) { $score += 500 }
        elseif ($File.Name -like $Hint) { $score += 350 }
    }

    switch ($File.Extension.ToLowerInvariant()) {
        '.bat' { $score += 250 }
        '.cmd' { $score += 225 }
        '.exe' { $score += 200 }
        '.jar' { $score += 175 }
        default { $score += 50 }
    }

    if ($nameLower -match 'server|dedicated|headless') { $score += 250 }
    if ($nameLower -match '^start') { $score += 125 }
    if ($nameLower -match 'shipping') { $score += 100 }
    if ($relativeLower -match 'binaries\\win64|bin\\|server\\|dedicated\\') { $score += 90 }

    if ($nameLower -match 'steamcmd|steamservice|steamerrorreporter|steamwebhelper') {
        $score -= 6000
    }
    if ($nameLower -match 'install|update|uninstall|setup|redist|prereq|crashhandler|crashreport|easyanticheat|eac') {
        $score -= 5000
    }
    if ($dirLower -match 'redist|prereq|easyanticheat') {
        $score -= 3000
    }

    $score -= ($depth * 40)
    return $score
}

function Find-ServerExecutable {
    param(
        [string]$FolderPath,
        [string[]]$Hints = @()
    )
    if (-not (Test-Path $FolderPath)) { return $null }
    $hintList = @($Hints)

    $candidates = New-Object 'System.Collections.Generic.List[object]'
    $seenPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $addCandidate = {
        param(
            [System.IO.FileInfo]$File,
            [string]$HintValue = '',
            [int]$HintIndex = 0
        )

        if ($null -eq $File) { return }
        $relativePath = _GetRelativeExecutablePath -BaseFolder $FolderPath -TargetPath $File.FullName
        if ([string]::IsNullOrWhiteSpace($relativePath)) { return }
        if (-not $seenPaths.Add($relativePath)) { return }

        $score = _GetExecutableCandidateScore -File $File -BaseFolder $FolderPath -Hint $HintValue -HintOrder $HintIndex
        $candidates.Add([pscustomobject]@{
            RelativePath = $relativePath
            Score = $score
            Depth = @($relativePath -split '[\\/]' | Where-Object { $_ -ne '' }).Count - 1
            Name = $File.Name
        }) | Out-Null
    }

    # Try hints first (in order given by template)
    for ($i = 0; $i -lt $hintList.Count; $i++) {
        $hint = $hintList[$i]
        foreach ($found in @(Get-ChildItem -LiteralPath $FolderPath -File -Filter $hint -ErrorAction SilentlyContinue)) {
            & $addCandidate -File $found -HintValue $hint -HintIndex $i
        }
    }
    if ($candidates.Count -gt 0) {
        return ($candidates | Sort-Object @{ Expression = 'Score'; Descending = $true }, @{ Expression = 'Depth'; Descending = $false }, 'Name' | Select-Object -First 1).RelativePath
    }

    for ($i = 0; $i -lt $hintList.Count; $i++) {
        $hint = $hintList[$i]
        foreach ($found in @(Get-ChildItem -LiteralPath $FolderPath -Recurse -File -Filter $hint -ErrorAction SilentlyContinue)) {
            & $addCandidate -File $found -HintValue $hint -HintIndex $i
        }
    }
    if ($candidates.Count -gt 0) {
        return ($candidates | Sort-Object @{ Expression = 'Score'; Descending = $true }, @{ Expression = 'Depth'; Descending = $false }, 'Name' | Select-Object -First 1).RelativePath
    }

    # Fallback: gather both top-level and nested candidates, then choose the
    # best overall match so wrapper roots (for example SteamCMD) do not win
    # over an actual server executable stored one folder deeper.
    foreach ($pattern in @('*.bat','*.cmd','*.exe','*.jar')) {
        foreach ($found in @(Get-ChildItem -LiteralPath $FolderPath -File -Filter $pattern -ErrorAction SilentlyContinue)) {
            & $addCandidate -File $found
        }
    }
    foreach ($pattern in @('*.bat','*.cmd','*.exe','*.jar')) {
        foreach ($found in @(Get-ChildItem -LiteralPath $FolderPath -Recurse -File -Filter $pattern -ErrorAction SilentlyContinue)) {
            & $addCandidate -File $found
        }
    }
    if ($candidates.Count -gt 0) {
        return ($candidates | Sort-Object @{ Expression = 'Score'; Descending = $true }, @{ Expression = 'Depth'; Descending = $false }, 'Name' | Select-Object -First 1).RelativePath
    }

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
        '7days'            = '7DaysToDie'
        '7daystodieserver' = '7DaysToDie'
        '7daystodie'       = '7DaysToDie'
        'sevendays'        = '7DaysToDie'
        'sdtd'             = '7DaysToDie'
        '7dtd'             = '7DaysToDie'
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

function _GetKnownGameFromProfile {
    param([System.Collections.IDictionary]$Profile)

    if (-not $Profile) { return $null }

    $candidateNames = New-Object 'System.Collections.Generic.List[string]'
    foreach ($field in @('KnownGame','GameName')) {
        if (($Profile.Keys -contains $field) -and -not [string]::IsNullOrWhiteSpace("$($Profile[$field])")) {
            $candidateNames.Add("$($Profile[$field])") | Out-Null
        }
    }

    foreach ($candidateName in $candidateNames) {
        if ($script:KnownGames.ContainsKey($candidateName)) {
            return $candidateName
        }
        $resolved = Resolve-KnownGame -FolderName $candidateName
        if ($resolved) { return $resolved }
    }

    $prefix = if (($Profile.Keys -contains 'Prefix') -and -not [string]::IsNullOrWhiteSpace("$($Profile.Prefix)")) {
        "$($Profile.Prefix)".ToUpperInvariant()
    } else {
        ''
    }

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

    return $null
}

function _CollectStringValues {
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
                _CollectStringValues -Value $Value[$key] -List $List -Seen $Seen
                return
            }
        }
        return
    }

    if ($Value -is [psobject]) {
        foreach ($propName in @('value','Value','values','Values','items','Items')) {
            $prop = $Value.PSObject.Properties[$propName]
            if ($prop) {
                _CollectStringValues -Value $prop.Value -List $List -Seen $Seen
                return
            }
        }
    }

    if ($Value -is [System.Collections.IEnumerable]) {
        foreach ($item in $Value) {
            _CollectStringValues -Value $item -List $List -Seen $Seen
        }
    }
}

function _GetNormalizedExeHints {
    param([System.Collections.IDictionary]$Profile)

    $list = New-Object 'System.Collections.Generic.List[string]'
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

    if ($Profile -and ($Profile.Keys -contains 'ExeHints')) {
        _CollectStringValues -Value $Profile.ExeHints -List $list -Seen $seen
    }

    if ($list.Count -eq 0 -and $Profile -and ($Profile.Keys -contains 'Executable')) {
        $exeName = [System.IO.Path]::GetFileName("$($Profile.Executable)")
        if (-not [string]::IsNullOrWhiteSpace($exeName) -and $seen.Add($exeName)) {
            $list.Add($exeName) | Out-Null
        }
    }

    $knownGame = _GetKnownGameFromProfile -Profile $Profile
    if ($knownGame -and $script:KnownGames.ContainsKey($knownGame)) {
        foreach ($hint in @($script:KnownGames[$knownGame].ExeHints)) {
            $candidate = "$hint".Trim()
            if (-not [string]::IsNullOrWhiteSpace($candidate) -and $seen.Add($candidate)) {
                $list.Add($candidate) | Out-Null
            }
        }
    }

    return ,([string[]]$list.ToArray())
}

function _NormalizeProfileSchema {
    param([System.Collections.IDictionary]$Profile)

    if (-not $Profile) { return }

    $Profile['ExeHints'] = _GetNormalizedExeHints -Profile $Profile
    if (-not ($Profile.Keys -contains 'BlockStartIfRamPercentUsed')) {
        $Profile['BlockStartIfRamPercentUsed'] = 0
    }
    if (-not ($Profile.Keys -contains 'BlockStartIfFreeRamBelowGB')) {
        $Profile['BlockStartIfFreeRamBelowGB'] = 0
    }
    if (-not ($Profile.Keys -contains 'StartupTimeoutSeconds')) {
        $knownGame = ''
        try { $knownGame = _GetKnownGameFromProfile -Profile $Profile } catch { $knownGame = '' }
        if ($knownGame -eq 'ProjectZomboid') {
            $Profile['StartupTimeoutSeconds'] = 600
        } else {
            $Profile['StartupTimeoutSeconds'] = 300
        }
    }
    if (-not ($Profile.Keys -contains 'ShutdownIfNoPlayersAfterStartupMinutes')) {
        $Profile['ShutdownIfNoPlayersAfterStartupMinutes'] = 0
    }
    if (-not ($Profile.Keys -contains 'ShutdownIfEmptyAfterLastPlayerLeavesMinutes')) {
        $Profile['ShutdownIfEmptyAfterLastPlayerLeavesMinutes'] = 0
    }
}

function _PLog {
    param(
        [string]$Msg,
        [ValidateSet('INFO','WARN','ERROR','DEBUG')][string]$Level = 'INFO'
    )

    $writeLog = Get-Command -Name 'Write-Log' -ErrorAction SilentlyContinue
    if ($writeLog) {
        try {
            Write-Log -Message $Msg -Level $Level -Source 'ProfileManager'
            return
        } catch { }
    }

    if ($Level -eq 'DEBUG') { return }

    $color = switch ($Level) {
        'ERROR' { 'Red' }
        'WARN'  { 'Yellow' }
        default { 'Cyan' }
    }
    try { Write-Host "[ProfileManager][$Level] $Msg" -ForegroundColor $color } catch { }
}

function _FormatProfileDiffValue {
    param([object]$Value)

    if ($null -eq $Value) { return '<null>' }

    $serializable = ConvertTo-SerializableObject -Obj $Value
    if ($serializable -is [string] -or $serializable -is [ValueType]) {
        $text = "$serializable"
    } else {
        try {
            $jsonText = @($serializable | ConvertTo-Json -Depth 6 -Compress)
            $text = ($jsonText -join '')
        } catch {
            $text = "$serializable"
        }
    }

    $text = ($text -replace '\s+', ' ').Trim()
    if ($text.Length -gt 160) {
        return $text.Substring(0, 160) + '...'
    }
    return $text
}

function _FlattenProfileForDiff {
    param(
        [object]$Value,
        [string]$Path = ''
    )

    $result = @{}
    $serializable = ConvertTo-SerializableObject -Obj $Value

    if ($null -eq $serializable) {
        $result[$Path] = '<null>'
        return $result
    }

    if ($serializable -is [System.Collections.IDictionary]) {
        foreach ($key in @($serializable.Keys | Sort-Object)) {
            $childPath = if ([string]::IsNullOrWhiteSpace($Path)) { [string]$key } else { "$Path.$key" }
            $childMap = _FlattenProfileForDiff -Value $serializable[$key] -Path $childPath
            foreach ($childKey in $childMap.Keys) { $result[$childKey] = $childMap[$childKey] }
        }
        return $result
    }

    if ($serializable -is [System.Collections.IEnumerable] -and $serializable -isnot [string]) {
        $result[$Path] = _FormatProfileDiffValue -Value $serializable
        return $result
    }

    $result[$Path] = _FormatProfileDiffValue -Value $serializable
    return $result
}

function _GetProfileDiffSummary {
    param(
        [object]$Before,
        [object]$After,
        [int]$MaxEntries = 25
    )

    $beforeMap = _FlattenProfileForDiff -Value $Before
    $afterMap = _FlattenProfileForDiff -Value $After
    $allKeys = @((@($beforeMap.Keys) + @($afterMap.Keys)) | Sort-Object -Unique)
    $diffs = New-Object 'System.Collections.Generic.List[string]'

    foreach ($key in $allKeys) {
        $hasBefore = $beforeMap.ContainsKey($key)
        $hasAfter = $afterMap.ContainsKey($key)
        $beforeValue = if ($hasBefore) { $beforeMap[$key] } else { '<missing>' }
        $afterValue = if ($hasAfter) { $afterMap[$key] } else { '<missing>' }

        if ($beforeValue -ne $afterValue) {
            $diffs.Add("${key}: $beforeValue -> $afterValue") | Out-Null
        }
    }

    if ($diffs.Count -le $MaxEntries) {
        return ,($diffs.ToArray())
    }

    $truncated = New-Object 'System.Collections.Generic.List[string]'
    foreach ($item in ($diffs | Select-Object -First $MaxEntries)) {
        $truncated.Add($item) | Out-Null
    }
    $truncated.Add("... $($diffs.Count - $MaxEntries) more change(s)") | Out-Null
    return ,($truncated.ToArray())
}

function Resolve-KnownGameFromFolderDetailed {
    param(
        [string]$FolderPath,
        [string]$PreferredGameName = '',
        [string]$DisplayName = ''
    )

    $fallbackKeys = New-Object 'System.Collections.Generic.List[string]'
    foreach ($candidateName in @($PreferredGameName, $DisplayName, (Split-Path $FolderPath -Leaf))) {
        if ([string]::IsNullOrWhiteSpace($candidateName)) { continue }
        $candidateKey = Resolve-KnownGame -FolderName $candidateName
        if ($candidateKey -and -not $fallbackKeys.Contains($candidateKey)) {
            $fallbackKeys.Add($candidateKey) | Out-Null
        }
    }

    $result = [ordered]@{
        KnownGame  = $null
        Resolution = 'generic'
        Detail     = ''
        BestScore  = 0
        Candidates = @()
    }

    if (-not (Test-Path -LiteralPath $FolderPath)) {
        if ($fallbackKeys.Count -gt 0) {
            $result.KnownGame = $fallbackKeys[0]
            $result.Resolution = 'fallback'
            $result.Detail = "Folder missing; fallback candidates=$($fallbackKeys -join ', ')"
        } else {
            $result.Detail = 'Folder missing and no fallback candidate was resolved.'
        }
        return [pscustomobject]$result
    }

    $genericHintPatterns = @(
        '^start\.bat$',
        '^start-server\.bat$',
        '^startdedicated\.bat$',
        '^server\.exe$'
    )

    $signatureMap = @{
        'ProjectZomboid' = @('ProjectZomboid64.json', 'ProjectZomboid32.json', 'media\lua', 'media\maps')
        'Hytale'         = @('HytaleServer.aot', 'Assets.zip')
        'Minecraft'      = @('server.properties', 'eula.txt')
        'Palworld'       = @('Pal\Saved\Config\WindowsServer\PalWorldSettings.ini', 'Saved\Config\WindowsServer\PalWorldSettings.ini', 'Pal\Binaries\Win64\PalServer-Win64-Shipping.exe', 'Binaries\Win64\PalServer-Win64-Shipping.exe')
        '7DaysToDie'     = @('serverconfig.xml', '7DaysToDie_Data\boot.config')
        'Valheim'        = @('valheim_server_Data\boot.config', 'BepInEx\plugins')
        'Terraria'       = @('TerrariaServer.exe.config', 'tshock')
        'Rust'           = @('RustDedicated_Data', 'Bundles')
        'Ark'            = @('ShooterGame\Binaries\Win64', 'ShooterGame\Saved\Config')
        'Satisfactory'   = @('FactoryGame\Binaries\Win64\FactoryServer-Win64-Shipping.exe', 'FactoryGame\Saved\Config\WindowsServer')
    }

    $bestKey = $null
    $bestScore = 0
    $bestReasons = @()
    $candidateSummaries = New-Object 'System.Collections.Generic.List[string]'

    foreach ($key in $script:KnownGames.Keys) {
        $template = $script:KnownGames[$key]
        $score = 0
        $reasons = New-Object 'System.Collections.Generic.List[string]'

        foreach ($hint in @($template.ExeHints)) {
            if ([string]::IsNullOrWhiteSpace($hint)) { continue }
            $hintName = [System.IO.Path]::GetFileName($hint).ToLowerInvariant()
            $hintScore = 120
            $nestedHintScore = 100
            if ($genericHintPatterns | Where-Object { $hintName -match $_ }) {
                $hintScore = 15
                $nestedHintScore = 10
            }
            $match = Get-ChildItem -LiteralPath $FolderPath -File -Filter $hint -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($match) {
                $score += $hintScore
                $reasons.Add("exe(top)=$hint") | Out-Null
                break
            }
            $nestedMatch = Get-ChildItem -LiteralPath $FolderPath -Recurse -File -Filter $hint -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($nestedMatch) {
                $score += $nestedHintScore
                $reasons.Add("exe(nested)=$hint") | Out-Null
                break
            }
        }

        if ($template.ContainsKey('Executable') -and -not [string]::IsNullOrWhiteSpace("$($template.Executable)")) {
            $exeCandidate = Join-Path $FolderPath "$($template.Executable)"
            if (Test-Path -LiteralPath $exeCandidate) {
                $score += 120
                $reasons.Add("template-exe=$($template.Executable)") | Out-Null
            }
        }

        foreach ($relativePath in @($signatureMap[$key])) {
            if ([string]::IsNullOrWhiteSpace($relativePath)) { continue }
            if (Test-Path -LiteralPath (Join-Path $FolderPath $relativePath)) {
                $score += 45
                $reasons.Add("signature=$relativePath") | Out-Null
            }
        }

        if ($fallbackKeys.Contains($key)) {
            $score += 10
            $reasons.Add('name-fallback') | Out-Null
        }

        if ($score -gt 0 -or $reasons.Count -gt 0) {
            $candidateSummaries.Add("${key}: score=$score reasons=$($reasons -join ',')") | Out-Null
        }

        if ($score -gt $bestScore) {
            $bestScore = $score
            $bestKey = $key
            $bestReasons = @($reasons.ToArray())
        }
    }

    $result.Candidates = @($candidateSummaries.ToArray())
    $result.BestScore = $bestScore

    if ($bestScore -ge 40 -and -not [string]::IsNullOrWhiteSpace($bestKey)) {
        $result.KnownGame = $bestKey
        $result.Resolution = 'signature'
        $result.Detail = if ($bestReasons.Count -gt 0) { $bestReasons -join ', ' } else { 'signature score threshold met' }
        return [pscustomobject]$result
    }

    if ($fallbackKeys.Count -gt 0) {
        $result.KnownGame = $fallbackKeys[0]
        $result.Resolution = 'fallback'
        $result.Detail = "fallback candidates=$($fallbackKeys -join ', ')"
        return [pscustomobject]$result
    }

    $result.Detail = 'no supported game signature matched'
    return [pscustomobject]$result
}

function Resolve-KnownGameFromFolder {
    param(
        [string]$FolderPath,
        [string]$PreferredGameName = '',
        [string]$DisplayName = ''
    )
    $details = Resolve-KnownGameFromFolderDetailed -FolderPath $FolderPath -PreferredGameName $PreferredGameName -DisplayName $DisplayName
    return $details.KnownGame
}

function Resolve-ProfileSourceFolder {
    param(
        [string]$FolderPath,
        [string]$PreferredGameName = '',
        [string]$DisplayName = ''
    )

    $result = [ordered]@{
        FolderPath = $FolderPath
        Reason     = 'selected-folder'
        KnownGame  = $null
        Detail     = ''
    }

    if ([string]::IsNullOrWhiteSpace($FolderPath) -or -not (Test-Path -LiteralPath $FolderPath)) {
        return [pscustomobject]$result
    }

    $selectedResolution = Resolve-KnownGameFromFolderDetailed -FolderPath $FolderPath -PreferredGameName $PreferredGameName -DisplayName $DisplayName
    $result.KnownGame = $selectedResolution.KnownGame
    $result.Detail = $selectedResolution.Detail

    $preferredKnownGame = $null
    foreach ($nameCandidate in @($PreferredGameName, $DisplayName, (Split-Path $FolderPath -Leaf))) {
        if ([string]::IsNullOrWhiteSpace($nameCandidate)) { continue }
        $resolved = Resolve-KnownGame -FolderName $nameCandidate
        if (-not [string]::IsNullOrWhiteSpace($resolved)) {
            $preferredKnownGame = $resolved
            break
        }
    }

    $searchRoots = New-Object 'System.Collections.Generic.List[string]'
    $steamCommon = Join-Path $FolderPath 'steamapps\common'
    if (Test-Path -LiteralPath $steamCommon) {
        $searchRoots.Add($steamCommon) | Out-Null
    }
    if ((Split-Path $FolderPath -Leaf).Equals('common', [System.StringComparison]::OrdinalIgnoreCase)) {
        $searchRoots.Add($FolderPath) | Out-Null
    }

    if ($searchRoots.Count -eq 0) {
        return [pscustomobject]$result
    }

    $bestChild = $null
    $bestChildResolution = $null
    $bestChildScore = [int]::MinValue

    foreach ($searchRoot in $searchRoots) {
        foreach ($child in @(Get-ChildItem -LiteralPath $searchRoot -Directory -ErrorAction SilentlyContinue)) {
            $childResolution = Resolve-KnownGameFromFolderDetailed -FolderPath $child.FullName -PreferredGameName $PreferredGameName -DisplayName $child.Name
            $candidateScore = [int]$childResolution.BestScore

            if ($preferredKnownGame -and $childResolution.KnownGame -eq $preferredKnownGame) {
                $candidateScore += 1000
            } elseif (-not [string]::IsNullOrWhiteSpace($childResolution.KnownGame)) {
                $candidateScore += 200
            }

            if ($candidateScore -gt $bestChildScore) {
                $bestChild = $child.FullName
                $bestChildResolution = $childResolution
                $bestChildScore = $candidateScore
            }
        }
    }

    if ([string]::IsNullOrWhiteSpace($bestChild)) {
        return [pscustomobject]$result
    }

    $shouldUseChild = $false
    if ($preferredKnownGame -and $bestChildResolution.KnownGame -eq $preferredKnownGame) {
        $shouldUseChild = $true
    } elseif (($selectedResolution.BestScore -lt 40 -or [string]::IsNullOrWhiteSpace($selectedResolution.KnownGame)) -and $bestChildResolution.BestScore -ge 40) {
        $shouldUseChild = $true
    }

    if (-not $shouldUseChild) {
        return [pscustomobject]$result
    }

    $result.FolderPath = $bestChild
    $result.Reason = 'steam-common-child'
    $result.KnownGame = $bestChildResolution.KnownGame
    $result.Detail = "Selected '$FolderPath' but resolved server folder '$bestChild' ($($bestChildResolution.Detail))"
    return [pscustomobject]$result
}

# =============================================================================
#  Auto-generate a profile object from a folder path (does NOT save to disk)
# =============================================================================
function New-GameProfile {
    param(
        [string]$FolderPath,
        [string]$GameName   = '',
        [string]$KnownGame  = '',
        [string]$Prefix     = '',
        [string]$Executable = '',
        [string]$LaunchArgs = ''
    )

    if (-not (Test-Path $FolderPath)) {
        throw "Folder not found: $FolderPath"
    }

    $originalFolderPath = $FolderPath
    $resolvedSourceFolder = Resolve-ProfileSourceFolder -FolderPath $FolderPath -PreferredGameName $KnownGame -DisplayName $GameName
    if ($resolvedSourceFolder -and -not [string]::IsNullOrWhiteSpace($resolvedSourceFolder.FolderPath)) {
        $FolderPath = $resolvedSourceFolder.FolderPath
        if (-not $GameName) {
            $GameName = (Split-Path $FolderPath -Leaf)
        }
        if (-not $originalFolderPath.Equals($FolderPath, [System.StringComparison]::OrdinalIgnoreCase)) {
            _PLog "Adjusted selected folder from '$originalFolderPath' to '$FolderPath' (reason=$($resolvedSourceFolder.Reason) detail=$($resolvedSourceFolder.Detail))" -Level DEBUG
        }
    }

    # Derive game name from folder name if not supplied
    if (-not $GameName) {
        $GameName = (Split-Path $FolderPath -Leaf)
    }

    # Detect the supported game from folder signatures first so custom display
    # names and renamed install folders still resolve to the correct template.
    $resolution = Resolve-KnownGameFromFolderDetailed -FolderPath $FolderPath -PreferredGameName $KnownGame -DisplayName $GameName
    $knownKey = $resolution.KnownGame
    $template = if ($knownKey) { $script:KnownGames[$knownKey] } else { $null }
    $templateGameName = if ($knownKey) { $knownKey } else { $GameName }

    _PLog "Template selection: display='$GameName' preferred='$KnownGame' folder='$FolderPath' selected='$(if ($knownKey) { $knownKey } else { 'generic' })' mode=$($resolution.Resolution) detail=$($resolution.Detail)" -Level DEBUG
    if ($resolution.Candidates -and @($resolution.Candidates).Count -gt 0) {
        _PLog "Template candidates: $(@($resolution.Candidates) -join ' | ')" -Level DEBUG
    }
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
    $resolvedExecutable = $Executable
    $preserveLaunchArgs = $PSBoundParameters.ContainsKey('LaunchArgs') -and -not [string]::IsNullOrWhiteSpace($LaunchArgs)
    $explicitLaunchArgs = $LaunchArgs

    # --- Stop / Save methods ---
    $stopMethod = if ($template) { $template.StopMethod  } else { 'processKill' }
    $saveMethod = if ($template) { $template.SaveMethod  } else { 'none' }
    # Accept canonical field names (StdinStopCommand / StdinSaveCommand) or legacy short names
    $stdinStop  = ''
    if ($template) {
        if ($template.ContainsKey('StdinStopCommand') -and $template.StdinStopCommand) { $stdinStop = $template.StdinStopCommand }
        elseif ($template.ContainsKey('StdinStop')    -and $template.StdinStop)        { $stdinStop = $template.StdinStop }
    }
    $stdinSave  = ''
    if ($template) {
        if ($template.ContainsKey('StdinSaveCommand') -and $template.StdinSaveCommand) { $stdinSave = $template.StdinSaveCommand }
        elseif ($template.ContainsKey('StdinSave')    -and $template.StdinSave)        { $stdinSave = $template.StdinSave }
    }

    # --- Base commands (present for every game) ---
    # MUST be plain @{} not [ordered]@{} -- PS 5.1 ConvertTo-Json serializes
    # OrderedDictionary as an array of @{Key=x;Value=y} pairs, not a JSON object.
    $commands = @{
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
    $telnetHost   = if ($template -and $template.TelnetHost)     { $template.TelnetHost     } else { '' }
    $telnetPort   = if ($template -and $template.TelnetPort)     { $template.TelnetPort     } else { 0  }
    $telnetPass   = if ($template -and $template.TelnetPassword) { $template.TelnetPassword } else { '' }
    $steamAppId   = if ($template -and $template.SteamAppId)     { $template.SteamAppId     } else { '' }
    $configRoot   = if ($template -and $template.ConfigRoot)     {
        [Environment]::ExpandEnvironmentVariables("$($template.ConfigRoot)")
    } else { '' }
    if ([string]::IsNullOrWhiteSpace($configRoot)) { $configRoot = $FolderPath }

    # For games that store configs in multiple locations (e.g. Satisfactory stores
    # primary configs in %LOCALAPPDATA% AND secondary configs in the install folder)
    # we build a ConfigRoots array holding all valid paths.  The GUI Config Editor
    # iterates this list and shows a dropdown when more than one root exists.
    #
    # Use a typed List then call .ToArray() at the end.
    # PS += on @() produces a decorated PSObject array that ConvertTo-Json
    # serializes as "@{Count=N; value=System.Object[]}" instead of a JSON array.
    $configRootsList = New-Object 'System.Collections.Generic.List[string]'
    if (-not [string]::IsNullOrWhiteSpace($configRoot)) {
        $configRootsList.Add([string]$configRoot)
    }
    # Satisfactory: also expose the WindowsServer subfolder inside the install folder
    if ($knownKey -eq 'Satisfactory') {
        $sfSecondary = [System.IO.Path]::Combine([string]$FolderPath, 'FactoryGame', 'Saved', 'Config', 'WindowsServer')
        if ($sfSecondary -and -not $configRootsList.Contains($sfSecondary)) {
            $configRootsList.Add([string]$sfSecondary)
        }
    }
    # Plain string array - ConvertTo-Json serializes this correctly as a JSON array
    $configRoots = $configRootsList.ToArray()

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
            $localLow = _GetWindowsLocalLowPath
            if (-not [string]::IsNullOrWhiteSpace($localLow)) {
                $serverLogRoot = Join-Path $localLow 'IronGate\Valheim'
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

    # Inject launch args from template when none supplied by caller.
    # Check template.LaunchArgs first (direct), then DefaultLaunchArgs (legacy).
    if ([string]::IsNullOrWhiteSpace($LaunchArgs) -and $template) {
        if ($template.ContainsKey('LaunchArgs') -and -not [string]::IsNullOrWhiteSpace($template.LaunchArgs)) {
            $LaunchArgs = $template.LaunchArgs
        } elseif ($template.ContainsKey('DefaultLaunchArgs') -and -not [string]::IsNullOrWhiteSpace($template.DefaultLaunchArgs)) {
            $LaunchArgs = $template.DefaultLaunchArgs
        }
    }

    # Build LaunchArgState from catalog (if available) so new profiles start with full toggles.
    $launchArgState = $null
    try {
        $launchArgState = Build-LaunchArgState -GameName $templateGameName -LaunchArgs $LaunchArgs
        if ($launchArgState) {
            $LaunchArgs = Build-LaunchArgsFromState -GameName $templateGameName -State $launchArgState
        }
    } catch {
        $launchArgState = $null
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
        KnownGame           = if ($knownKey) { $knownKey } else { '' }
        Executable          = $Executable
        FolderPath          = $FolderPath
        LaunchArgs          = $LaunchArgs
        LaunchArgState      = $launchArgState
        StopMethod          = $stopMethod
        ProcessName         = $processName
        HideConsoleWindow   = if ($template -and $null -ne $template.HideConsoleWindow) { [bool]$template.HideConsoleWindow } else { $false }
        StdinStopCommand    = $stdinStop
        SaveMethod          = $saveMethod
        StdinSaveCommand    = $stdinSave
        SaveWaitSeconds     = $saveWaitSecs
        RconHost            = '127.0.0.1'
        RconPort            = $rconPort
        RconPassword        = ''
        TelnetHost          = $telnetHost
        TelnetPort          = $telnetPort
        TelnetPassword      = $telnetPass
        SteamAppId          = $steamAppId
        MinRamGB            = $minRamGB
        MaxRamGB            = $maxRamGB
        BlockStartIfRamPercentUsed = 0
        BlockStartIfFreeRamBelowGB = 0
        StartupTimeoutSeconds = if ($template -and $null -ne $template.StartupTimeoutSeconds) { [int]$template.StartupTimeoutSeconds } else { 300 }
        ShutdownIfNoPlayersAfterStartupMinutes = 0
        ShutdownIfEmptyAfterLastPlayerLeavesMinutes = 0
        EnableAutoRestart   = $true
        RestartDelaySeconds = 10
        MaxRestartsPerHour  = 5
        LogStrategy         = $logStrategy
        ServerLogRoot       = $serverLogRoot
        ServerLogPath       = $serverLogPath
        ServerLogNote       = $logPathNote
        ConfigRoot          = $configRoot
        ConfigRoots         = $configRoots
        Commands            = $commands
        AOTCache            = if ($template -and $template.AOTCache)   { $template.AOTCache }   else { '' }
        AssetFile           = if ($template -and $template.AssetFile)  { $template.AssetFile }  else { '' }
        BackupDir           = if ($template -and $template.BackupDir)  { $template.BackupDir }  else { '' }
        # --- REST API (Palworld) ---
        RestEnabled     = if ($template) { $template.RestEnabled } else { $false }
        RestHost        = if ($template) { $template.RestHost    } else { '' }
        RestPort        = if ($template) { $template.RestPort    } else { 0  }
        RestPassword    = ''   # always blank; filled from INI at runtime
        RestProtocol    = if ($template) { $template.RestProtocol } else { 'http' }
        RestPollOnlyWhenRunning = $true

        # --- SATISFACTORY HTTPS API ---
        # Token generated via: server.GenerateAPIToken  (in the in-game console tab)
        # Leave blank if launching with -ini:Engine:[SystemSettings]:FG.DedicatedServer.AllowInsecureLocalAccess=1
        SatisfactoryApiHost  = if ($template -and $template.SatisfactoryApiHost) { $template.SatisfactoryApiHost } else { '127.0.0.1' }
        SatisfactoryApiPort  = if ($template -and $template.SatisfactoryApiPort) { $template.SatisfactoryApiPort } else { 7777 }
        SatisfactoryApiToken = ''   # user fills this in the Profile Editor after generating it

        # --- LOGGING (extra fields) ---
        ServerLogSubDir = if ($template -and $template.ServerLogSubDir) { $template.ServerLogSubDir } else { '' }
        ServerLogFile   = if ($template -and $template.ServerLogFile)   { $template.ServerLogFile   } else { '' }

        # --- EXTRA COMMANDS (already merged into Commands, but store raw template too) ---
        # Deep-copy each entry as a plain @{} so we never hold a reference to the
        # in-memory template (which would mutate it on profile edits).
        ExtraCommands   = if ($template -and $template.ExtraCommands) {
                              $ec = @{}
                              foreach ($k in $template.ExtraCommands.Keys) {
                                  $ecDef = $template.ExtraCommands[$k]
                                  $ecCopy = @{}
                                  foreach ($ek in $ecDef.Keys) { $ecCopy[$ek] = $ecDef[$ek] }
                                  $ec[$k] = $ecCopy
                              }
                              $ec
                          } else { @{} }

        # --- EXECUTABLE HINTS ---
        # Force to plain [string[]] - PS arrays built with @() or += can carry
        # PSObject decoration that ConvertTo-Json emits as "@{Count=N; value=...}"
        ExeHints        = if ($template -and $template.ExeHints) {
                              [string[]]@($template.ExeHints)
                          } else { [string[]]@() }

    }

    $profileBeforeDefaults = ConvertTo-SerializableObject -Obj $profile

    # Apply shipped defaults when no user profiles exist yet
    $defaults = _FindDefaultProfileForGame -GameName $templateGameName
    if ($defaults) {
        _ApplyDefaultProfileValues -Profile $profile -Defaults $defaults

        # Keep install-specific executable detection and any explicit launch args
        # supplied by the caller. Defaults should shape generated profiles, but
        # they should not override values that came from the selected folder or
        # an explicit Add Game choice.
        if (-not [string]::IsNullOrWhiteSpace($resolvedExecutable)) {
            $profile.Executable = $resolvedExecutable
        }
        if ($preserveLaunchArgs) {
            $profile.LaunchArgs = $explicitLaunchArgs
        }

        # If defaults provide an explicit LaunchArgs string, keep that exact
        # string and only rebuild the editor state from it.
        if ($preserveLaunchArgs -and -not [string]::IsNullOrWhiteSpace("$($profile.LaunchArgs)")) {
            try {
                $profile.LaunchArgState = Build-LaunchArgState -GameName $templateGameName -LaunchArgs $profile.LaunchArgs
            } catch { }
        } elseif (($defaults.Keys -contains 'LaunchArgs') -and -not [string]::IsNullOrWhiteSpace("$($profile.LaunchArgs)")) {
            try {
                $profile.LaunchArgState = Build-LaunchArgState -GameName $templateGameName -LaunchArgs $profile.LaunchArgs
            } catch { }
        } elseif ($profile.LaunchArgState) {
            try {
                $profile.LaunchArgs = Build-LaunchArgsFromState -GameName $templateGameName -State $profile.LaunchArgState
            } catch { }
        } elseif ($profile.LaunchArgs) {
            try {
                $profile.LaunchArgState = Build-LaunchArgState -GameName $templateGameName -LaunchArgs $profile.LaunchArgs
                $profile.LaunchArgs = Build-LaunchArgsFromState -GameName $templateGameName -State $profile.LaunchArgState
            } catch { }
        }
    }

    _NormalizeProfileSchema -Profile $profile

    $profileAfterGeneration = ConvertTo-SerializableObject -Obj $profile
    $generationDiff = _GetProfileDiffSummary -Before $profileBeforeDefaults -After $profileAfterGeneration -MaxEntries 25
    if ($generationDiff.Count -gt 0) {
        _PLog "Generated profile diff [$($profile.Prefix)][$($profile.GameName)]: $($generationDiff -join '; ')" -Level DEBUG
    } else {
        _PLog "Generated profile diff [$($profile.Prefix)][$($profile.GameName)]: no changes after defaults/normalization." -Level DEBUG
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
#  SERIALIZATION HELPER
#
#  PS 5.1 ConvertTo-Json has two silent corruption bugs that destroy nested
#  objects in profile JSON files:
#
#  1. [ordered]@{} / OrderedDictionary  -> serialized as an array of
#     @{Key=x; Value=y} pairs instead of a JSON object.
#
#  2. Arrays built with @() or += that carry PSObject decoration  ->
#     serialized as "@{Count=N; value=System.Object[]}" string.
#
#  This helper recursively converts the entire object graph to only the
#  plain types that ConvertTo-Json handles correctly:
#    - [hashtable]          for any dictionary / ordered-dict / PSCustomObject
#    - [object[]] / scalar  for arrays and primitives
#
#  Call it on the profile hashtable immediately before ConvertTo-Json.
# =============================================================================
function ConvertTo-SerializableObject {
    param([object]$Obj)

    if ($null -eq $Obj) { return $null }

    # Any dictionary type -> plain [hashtable]
    if ($Obj -is [System.Collections.IDictionary]) {
        $h = @{}
        foreach ($key in $Obj.Keys) {
            $h[[string]$key] = ConvertTo-SerializableObject -Obj $Obj[$key]
        }
        return $h
    }

    # PSCustomObject (e.g. result of ConvertFrom-Json) -> plain [hashtable]
    if ($Obj -is [psobject] -and $Obj -isnot [string]) {
        $props = $Obj.PSObject.Properties
        if ($props -and $props.Count -gt 0) {
            $h = @{}
            foreach ($prop in $props) {
                $h[$prop.Name] = ConvertTo-SerializableObject -Obj $prop.Value
            }
            return $h
        }
    }

    # Any enumerable except string -> plain [object[]]
    if ($Obj -is [System.Collections.IEnumerable] -and $Obj -isnot [string]) {
        $list = New-Object 'System.Collections.Generic.List[object]'
        foreach ($item in $Obj) {
            $list.Add((ConvertTo-SerializableObject -Obj $item))
        }
        return ,($list.ToArray())
    }

    # Scalar (string, int, bool, double, etc.) - return as-is
    return $Obj
}

# =============================================================================
#  Save a profile hashtable to the Profiles folder as JSON
# =============================================================================
function Save-GameProfile {
    param(
        [hashtable]$Profile,
        [string]$ProfilesDir = '.\Profiles'
    )

    _NormalizeProfileSchema -Profile $Profile

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
        # Sanitize the entire object graph before serializing.
        # This eliminates OrderedDictionary and decorated-array corruption.
        $serializable = ConvertTo-SerializableObject -Obj $Profile
        $profileDiff = _GetProfileDiffSummary -Before $Profile -After $serializable -MaxEntries 20
        if ($profileDiff.Count -gt 0) {
            _PLog "Save profile diff [$($Profile.Prefix)][$($Profile.GameName)]: $($profileDiff -join '; ')" -Level DEBUG
        } else {
            _PLog "Save profile diff [$($Profile.Prefix)][$($Profile.GameName)]: no serialization changes." -Level DEBUG
        }
        $serializable | ConvertTo-Json -Depth 10 | Set-Content -Path $outPath -Encoding UTF8 -Force
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
            _NormalizeProfileSchema -Profile $p
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

    # Use a case-insensitive hashtable for ALL dictionary conversions.
    # This ensures $profile.Commands["restart"] works regardless of the
    # casing used in JSON keys, and that ConvertTo-Json serializes the
    # result correctly as a JSON object (not as an array of key-value pairs,
    # which is what PS5.1 does with [ordered]@{} / OrderedDictionary).
    if ($Object -is [System.Collections.IDictionary]) {
        $hash = New-Object 'System.Collections.Hashtable' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($key in $Object.Keys) {
            $hash[[string]$key] = ConvertTo-Hashtable -Object $Object[$key]
        }
        return $hash
    }
    elseif ($Object -is [psobject]) {
        # PSCustomObject from ConvertFrom-Json: convert every property recursively.
        $hash = New-Object 'System.Collections.Hashtable' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($prop in $Object.PSObject.Properties) {
            $hash[$prop.Name] = ConvertTo-Hashtable -Object $prop.Value
        }
        return $hash
    }
    elseif ($Object -is [System.Collections.IEnumerable] -and
            -not ($Object -is [string])) {
        return ,(@($Object | ForEach-Object { ConvertTo-Hashtable -Object $_ }))
    }
    else {
        return $Object
    }
}

# =============================================================================
#  Launch Argument Catalog Helpers
# =============================================================================
$script:LaunchArgCatalog = $null

function Get-LaunchArgCatalog {
    if ($script:LaunchArgCatalog) { return $script:LaunchArgCatalog }
    $base = Split-Path -Path $PSScriptRoot -Parent
    $path = Join-Path $base 'Config\LaunchArgCatalog.json'
    if (-not (Test-Path $path)) { return $null }
    try {
        $raw = Get-Content -Raw -Path $path -Encoding UTF8
        $cat = $raw | ConvertFrom-Json
        $script:LaunchArgCatalog = ConvertTo-Hashtable -Object $cat
        return $script:LaunchArgCatalog
    } catch {
        Write-Host "[ProfileManager] WARN - Failed to load LaunchArgCatalog.json: $_" -ForegroundColor Yellow
        return $null
    }
}

function _NormalizeGameKey([string]$name) {
    if ([string]::IsNullOrWhiteSpace($name)) { return '' }
    return ($name -replace '[^a-zA-Z0-9]', '').ToLowerInvariant()
}

function Resolve-LaunchArgGameKey {
    param([string]$GameName)
    $cat = Get-LaunchArgCatalog
    if (-not $cat -or -not $cat.Games) { return $null }
    $target = _NormalizeGameKey $GameName
    foreach ($k in $cat.Games.Keys) {
        if ((_NormalizeGameKey $k) -eq $target) { return $k }
    }
    return $null
}

function Get-LaunchArgDefinitions {
    param([string]$GameName)
    $cat = Get-LaunchArgCatalog
    if (-not $cat -or -not $cat.Games) { return $null }
    $key = Resolve-LaunchArgGameKey -GameName $GameName
    if ($key -and $cat.Games.ContainsKey($key)) {
        return $cat.Games[$key].Args
    }
    return $null
}

# Return merged launch arg definitions across ALL games
function Get-AllLaunchArgDefinitions {
    $cat = Get-LaunchArgCatalog
    if (-not $cat -or -not $cat.Games) { return $null }

    $merged = @{}
    foreach ($g in $cat.Games.Keys) {
        $args = $cat.Games[$g].Args
        if (-not $args) { continue }
        foreach ($def in @($args)) {
            if (-not $def -or -not $def.Key) { continue }
            $k = $def.Key.ToLowerInvariant()
            if (-not $merged.ContainsKey($k)) {
                $merged[$k] = $def
            }
        }
    }

    return @($merged.Values)
}

# =============================================================================
#  Default Profile Overlay (use existing profile as defaults for new ones)
# =============================================================================
function _FindDefaultProfileForGame {
    param([string]$GameName)
    if ([string]::IsNullOrWhiteSpace($GameName)) { return $null }

    $base = Split-Path -Path $PSScriptRoot -Parent
    $defaultsFile = Join-Path $base 'Config\DefaultProfileTemplates.json'
    if (-not (Test-Path $defaultsFile)) { return $null }

    $target = _NormalizeGameKey $GameName
    try {
        $raw = Get-Content -Raw -Path $defaultsFile -Encoding UTF8 | ConvertFrom-Json
        $ht  = ConvertTo-Hashtable -Object $raw
        foreach ($k in $ht.Keys) {
            $p = $ht[$k]
            if (-not $p) { continue }
            $gn = if ($p.GameName) { "$($p.GameName)" } else { '' }
            if ((_NormalizeGameKey $gn) -eq $target) { return $p }
        }
    } catch { }
    return $null
}

function _ApplyDefaultProfileValues {
    param([System.Collections.IDictionary]$Profile, [System.Collections.IDictionary]$Defaults)
    if (-not $Profile -or -not $Defaults) { return }

    $exclude = @('GameName','FolderPath','Prefix')
    $copy = ConvertTo-SerializableObject -Obj $Defaults
    foreach ($k in $copy.Keys) {
        if ($exclude -contains $k) { continue }
        $value = $copy[$k]
        if ($null -eq $value) { continue }
        if ($value -is [string] -and [string]::IsNullOrWhiteSpace($value)) { continue }
        if ($value -is [System.Collections.IDictionary] -and $value.Count -eq 0) { continue }
        if ($value -is [System.Collections.IEnumerable] -and $value -isnot [string]) {
            if (@($value).Count -eq 0) { continue }
        }
        $Profile[$k] = $copy[$k]
    }

    _RefreshGeneratedProfileDerivedFields -Profile $Profile
}

function _RefreshGeneratedProfileDerivedFields {
    param([System.Collections.IDictionary]$Profile)

    if (-not $Profile) { return }

    $hasGameName = ($Profile.Keys -contains 'GameName')
    $hasFolderPath = ($Profile.Keys -contains 'FolderPath')
    $hasConfigRoot = ($Profile.Keys -contains 'ConfigRoot')
    $hasLogStrategy = ($Profile.Keys -contains 'LogStrategy')
    $hasServerLogSubDir = ($Profile.Keys -contains 'ServerLogSubDir')
    $hasServerLogFile = ($Profile.Keys -contains 'ServerLogFile')
    $hasServerLogRoot = ($Profile.Keys -contains 'ServerLogRoot')

    $gameName = if ($hasGameName) { "$($Profile.GameName)" } else { '' }
    $gameKey = _NormalizeGameKey $gameName
    $folderPath = if ($hasFolderPath) { "$($Profile.FolderPath)" } else { '' }

    $configRoot = ''
    if ($hasConfigRoot -and -not [string]::IsNullOrWhiteSpace("$($Profile.ConfigRoot)")) {
        $configRoot = "$($Profile.ConfigRoot)"
    } elseif (-not [string]::IsNullOrWhiteSpace($folderPath)) {
        $configRoot = $folderPath
    }
    if (-not [string]::IsNullOrWhiteSpace($configRoot)) {
        $configRoot = [Environment]::ExpandEnvironmentVariables($configRoot)
        $Profile['ConfigRoot'] = $configRoot
    }

    $configRootsList = New-Object 'System.Collections.Generic.List[string]'
    if (-not [string]::IsNullOrWhiteSpace($configRoot)) {
        $configRootsList.Add($configRoot)
    }
    if ($gameKey -eq 'satisfactory' -and -not [string]::IsNullOrWhiteSpace($folderPath)) {
        $sfSecondary = [System.IO.Path]::Combine($folderPath, 'FactoryGame', 'Saved', 'Config', 'WindowsServer')
        if (-not [string]::IsNullOrWhiteSpace($sfSecondary) -and -not $configRootsList.Contains($sfSecondary)) {
            $configRootsList.Add($sfSecondary)
        }
    }
    $Profile['ConfigRoots'] = $configRootsList.ToArray()

    $logStrategy = if ($hasLogStrategy -and -not [string]::IsNullOrWhiteSpace("$($Profile.LogStrategy)")) {
        "$($Profile.LogStrategy)"
    } else {
        'SingleFile'
    }
    $serverLogSubDir = if ($hasServerLogSubDir) { "$($Profile.ServerLogSubDir)" } else { '' }
    $serverLogFile = if ($hasServerLogFile) { "$($Profile.ServerLogFile)" } else { '' }
    $serverLogRoot = ''
    $serverLogPath = ''

    switch ($logStrategy) {
        'PZSessionFolder' {
            $serverLogRoot = Join-Path $env:USERPROFILE 'Zomboid\Logs'
            $serverLogPath = ''
        }

        'ValheimUserFolder' {
            $localLow = _GetWindowsLocalLowPath
            if (-not [string]::IsNullOrWhiteSpace($localLow)) {
                $serverLogRoot = Join-Path $localLow 'IronGate\Valheim'
            }
            $serverLogPath = Join-Path $serverLogRoot 'valheim_server.log'
        }

        'NewestFile' {
            if (-not [string]::IsNullOrWhiteSpace($folderPath)) {
                if (-not [string]::IsNullOrWhiteSpace($serverLogSubDir)) {
                    $serverLogRoot = Join-Path $folderPath $serverLogSubDir
                } else {
                    $serverLogRoot = $folderPath
                }
            } elseif ($hasServerLogRoot) {
                $serverLogRoot = "$($Profile.ServerLogRoot)"
            }

            $preferredFile = ''
            if (-not [string]::IsNullOrWhiteSpace($serverLogFile) -and
                -not $serverLogFile.StartsWith('*') -and
                -not $serverLogFile.StartsWith('$')) {
                $preferredFile = $serverLogFile
            }

            if (-not [string]::IsNullOrWhiteSpace($serverLogRoot) -and -not [string]::IsNullOrWhiteSpace($preferredFile)) {
                $serverLogPath = Join-Path $serverLogRoot $preferredFile
            }
        }

        default {
            if (-not [string]::IsNullOrWhiteSpace($folderPath) -and -not [string]::IsNullOrWhiteSpace($serverLogFile)) {
                if (-not [string]::IsNullOrWhiteSpace($serverLogSubDir)) {
                    $serverLogPath = Join-Path $folderPath (Join-Path $serverLogSubDir $serverLogFile)
                } else {
                    $serverLogPath = Join-Path $folderPath $serverLogFile
                }
                $serverLogRoot = Split-Path $serverLogPath -Parent
            }
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($serverLogRoot) -or $logStrategy -eq 'PZSessionFolder') {
        $Profile['ServerLogRoot'] = $serverLogRoot
    }
    $Profile['ServerLogPath'] = $serverLogPath
}

function _TokenizeLaunchArgs {
    param([string]$ArgsLine)
    if ([string]::IsNullOrWhiteSpace($ArgsLine)) { return @() }
    $matches = [regex]::Matches($ArgsLine, '\"[^\"]*\"|\S+')
    return @($matches | ForEach-Object { $_.Value })
}

function _UnquoteArgValue {
    param([string]$Value)
    if ($null -eq $Value) { return $Value }
    $v = "$Value"
    if ($v.Length -ge 2 -and $v.StartsWith('"') -and $v.EndsWith('"')) {
        return $v.Substring(1, $v.Length - 2)
    }
    return $v
}

function Build-LaunchArgState {
    param(
        [string]$GameName,
        [string]$LaunchArgs
    )

    $defs = Get-LaunchArgDefinitions -GameName $GameName
    if (-not $defs) {
        $fallback = if ($null -ne $LaunchArgs) { "$LaunchArgs" } else { '' }
        return @{ Args = @{}; CustomArgs = $fallback }
    }

    $state = @{
        Args        = @{}
        CustomArgs  = ''
        UnknownArgs = @()
    }

    $defMap = @{}
    $groupMap = @{}
    foreach ($def in $defs) {
        if (-not $def.Key) { continue }
        $defMap[$def.Key.ToLowerInvariant()] = $def
        if ($def.Type -eq 'flaggroup' -and $def.Keys) {
            foreach ($k in $def.Keys) {
                $groupMap["$k".ToLowerInvariant()] = $def
            }
        }
    }

    $tokens = _TokenizeLaunchArgs -ArgsLine $LaunchArgs
    $customParts = New-Object 'System.Collections.Generic.List[string]'

    for ($i = 0; $i -lt $tokens.Count; $i++) {
        $tok = $tokens[$i]
        if (-not $tok) { continue }

        if ($tok.StartsWith('-')) {
            $key = $tok
            $val = $null
            $usedEquals = $false

            if ($tok -match '^([^=]+)=(.+)$') {
                $key = $matches[1]
                $val = $matches[2]
                $usedEquals = $true
            } else {
                if ($i + 1 -lt $tokens.Count -and -not $tokens[$i + 1].StartsWith('-')) {
                    $val = $tokens[$i + 1]
                    $i++
                }
            }

            $keyLower = $key.ToLowerInvariant()

            if ($groupMap.ContainsKey($keyLower)) {
                $def = $groupMap[$keyLower]
                $state.Args[$def.Key] = @{ Enabled = $true }
                continue
            }

            if ($defMap.ContainsKey($keyLower)) {
                $def = $defMap[$keyLower]
                $entry = @{ Enabled = $true }
                if ($def.Type -ne 'flag' -and $def.Type -ne 'flaggroup') {
                    $entry.Value = _UnquoteArgValue $val
                }
                $state.Args[$def.Key] = $entry
                continue
            }

            # Unknown argument - keep as toggle
            $entry = @{ Enabled = $true }
            if ($val) { $entry.Value = _UnquoteArgValue $val }
            $state.Args[$key] = $entry
            # Track original format so we can rebuild
            $state.UnknownArgs += @{
                Key        = $key
                Value      = if ($entry.ContainsKey('Value')) { $entry.Value } else { $null }
                UsedEquals = $usedEquals
            }
            continue
        }

        # Non-flag token - preserve as-is
        $customParts.Add($tok) | Out-Null
    }

    foreach ($def in $defs) {
        if (-not $def.Key) { continue }
        if (-not $state.Args.ContainsKey($def.Key)) {
            $entry = @{ Enabled = $false }
            if ($def.Default) { $entry.Value = "$($def.Default)" }
            $state.Args[$def.Key] = $entry
        }
    }

    $state.CustomArgs = ($customParts -join ' ').Trim()
    return $state
}

function Build-LaunchArgsFromState {
    param(
        [string]$GameName,
        [hashtable]$State
    )

    if (-not $State) { return '' }
    $defs = Get-LaunchArgDefinitions -GameName $GameName
    if (-not $defs) {
        if ($State.CustomArgs) { return "$($State.CustomArgs)" }
        return ''
    }

    $parts = New-Object 'System.Collections.Generic.List[string]'

    foreach ($def in $defs) {
        if (-not $def.Key) { continue }
        if (-not $State.Args -or -not $State.Args.ContainsKey($def.Key)) { continue }
        $entry = $State.Args[$def.Key]
        if (-not $entry -or -not $entry.Enabled) { continue }

        if ($def.Type -eq 'flag') {
            $parts.Add($def.Key) | Out-Null
            continue
        }
        if ($def.Type -eq 'flaggroup' -and $def.Keys) {
            foreach ($k in $def.Keys) {
                $parts.Add($k) | Out-Null
            }
            continue
        }

        $val = if ($entry.ContainsKey('Value')) { "$($entry.Value)" } else { '' }
        if ([string]::IsNullOrWhiteSpace($val)) { continue }
        $qMode = if ($def.Quote) { "$($def.Quote)" } else { 'auto' }
        $vOut = $val
        if ($qMode -eq 'always' -or ($qMode -eq 'auto' -and $vOut -match '\s')) {
            $vOut = '"' + ($vOut -replace '"','\"') + '"'
        }

        if ($def.Style -eq 'equals') {
            $parts.Add("$($def.Key)=$vOut") | Out-Null
        } else {
            $parts.Add($def.Key) | Out-Null
            $parts.Add($vOut) | Out-Null
        }
    }

    # Append unknown args that are toggled on
    $known = @{}
    foreach ($def in $defs) {
        if ($def.Key) { $known[$def.Key.ToLowerInvariant()] = $true }
    }

    # Build a map for UnknownArgs formatting
    $unknownMap = @{}
    if ($State.UnknownArgs) {
        foreach ($ua in @($State.UnknownArgs)) {
            if ($ua -and $ua.Key) { $unknownMap[$ua.Key] = $ua }
        }
    }

    if ($State.Args) {
        foreach ($k in $State.Args.Keys) {
            if (-not $k) { continue }
            if ($known.ContainsKey($k.ToLowerInvariant())) { continue }
            $entry = $State.Args[$k]
            if (-not $entry -or -not $entry.Enabled) { continue }

            $val = $null
            if ($entry.ContainsKey('Value')) { $val = $entry.Value }
            $vOut = $val
            if ($null -ne $vOut -and "$vOut" -match '\s') {
                $vOut = '"' + ("$vOut" -replace '"','\"') + '"'
            }

            $usedEquals = $false
            if ($unknownMap.ContainsKey($k) -and $unknownMap[$k].UsedEquals -eq $true) {
                $usedEquals = $true
            }

            if ($null -eq $vOut -or "$vOut" -eq '') {
                $parts.Add($k) | Out-Null
            } elseif ($usedEquals) {
                $parts.Add("$k=$vOut") | Out-Null
            } else {
                $parts.Add($k) | Out-Null
                $parts.Add($vOut) | Out-Null
            }
        }
    }

    $custom = if ($State.CustomArgs) { "$($State.CustomArgs)".Trim() } else { '' }
    if ($custom) { $parts.Add($custom) | Out-Null }

    return ($parts -join ' ').Trim()
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
    Resolve-KnownGame, Resolve-KnownGameFromFolder, ConvertTo-Hashtable, Get-KnownGameKeys, `
    ConvertTo-SerializableObject, `
    Get-LaunchArgCatalog, Get-LaunchArgDefinitions, Get-AllLaunchArgDefinitions, `
    Build-LaunchArgState, Build-LaunchArgsFromState
