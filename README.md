# Etherium Command Center (ECC)

Etherium Command Center is a Windows desktop control panel for managing local dedicated game servers from one place. ECC combines a dashboard, profile system, live logs, a config editor, a Discord control bot, and game-specific helpers like the Project Zomboid item and vehicle spawner.

This README is written to match the current ECC codebase and intended packaged setup flow, including the UI layout, settings flow, command routing, and generated files.

## Table of Contents

- [What ECC Does](#what-ecc-does)
- [Supported Game Profile Types](#supported-game-profile-types)
- [Requirements](#requirements)
- [Install and First Launch](#install-and-first-launch)
- [UI Guide](#ui-guide)
- [Settings Guide](#settings-guide)
- [Discord Bot Guide](#discord-bot-guide)
- [Profiles Guide](#profiles-guide)
- [Per-Game Profile Setup Guide](#per-game-profile-setup-guide)
- [Config Editor Guide](#config-editor-guide)
- [Commands Window Guide](#commands-window-guide)
- [Project Zomboid Item and Vehicle Spawner Guide](#project-zomboid-item-and-vehicle-spawner-guide)
- [Server Control Guide](#server-control-guide)
- [Logs Guide](#logs-guide)
- [Folder Structure](#folder-structure)
- [Required Files and Generated Files](#required-files-and-generated-files)
- [FAQ](#faq)
- [Troubleshooting and Debug Questions](#troubleshooting-and-debug-questions)
- [Security Notes Before You Push a PR](#security-notes-before-you-push-a-pr)

## What ECC Does

ECC is designed for single-machine Windows hosting. It gives you:

- A three-column dashboard for profiles, server cards, and profile editing
- Start, stop, restart, status, and command routing per server profile
- A Discord listener that accepts commands from a monitored channel
- A webhook-based notification stream for online, offline, save, crash, and restart events
- A config editor for common server config file types
- Live Program Log, Discord log, and per-game log tabs
- Auto-save, scheduled restart, and auto-restart crash recovery
- Launch argument editors for supported games
- Project Zomboid item and vehicle command builders with local asset/cache support

ECC is not a remote multi-host orchestration platform. It is built around local server processes running on the same Windows machine as the app.

## Supported Game Profile Types

ECC is currently set up around these game profile types:

| Game | Prefix | Primary Control Path | Notes |
| --- | --- | --- | --- |
| 7 Days to Die | `DZ` | Telnet + stdin | Uses timestamped Unity logs and telnet for players/save/version |
| Hytale | `HY` | stdin | Java/JAR launch with AOT cache, assets, and backups |
| Palworld | `PW` | REST API | REST is the main control path for save/status/players/shutdown |
| Project Zomboid | `PZ` | stdin | PZ logs use dated session folders; commands window includes item/vehicle spawner |
| Satisfactory | `SF` | Satisfactory HTTPS API | Save/status/players/shutdown route through the API |
| Valheim | `VH` | process control | Uses fixed log file and basic process management |

When a profile is created, it is stored in `Profiles\*.json` and can be edited directly in ECC.

## Requirements

- Windows 10 or Windows 11
- PowerShell 5.1
- Administrator rights for process management
- Local access to the game server install folders you want ECC to manage
- Discord bot token, monitored channel ID, and webhook URL if you want full Discord control and notifications

Game-specific requirements:

- 7 Days to Die: telnet must be enabled in `serverconfig.xml` if you want telnet tests and telnet command routing to work
- Palworld: the server must expose its REST API and use a valid REST port
- Satisfactory: you need either a valid API token from `server.GenerateAPIToken` or local insecure API access enabled
- Project Zomboid: the item/vehicle spawner uses ECC cache data and can also fall back to local Project Zomboid server/client files when needed for catalogs or preview assets

## Install and First Launch

### Standard install

1. Download or clone the project to a local Windows folder.
2. Keep the folder structure intact.
3. Run `Start.bat` as Administrator.

`Start.bat` launches `Launch.ps1`, and `Launch.ps1` will:

- Elevate to Administrator if needed
- Unblock bundled `.ps1`, `.psm1`, and `.json` files
- Set PowerShell execution policy to `RemoteSigned` for the current user
- Create missing `Config`, `Profiles`, and `Logs` folders
- Start the ECC modules, listener, monitor, and GUI

### First launch behavior

If you are starting from a fresh packaged install:

- `Config\Settings.json` does not need to exist yet
- `Profiles\*.json` does not need to exist yet
- ECC creates the main folders at launch
- ECC can start with empty/default runtime settings
- The **Settings** window still opens even if the file is missing
- `Config\Settings.json` is created the first time the user clicks `Save Settings`
- Profile JSON files are created when profiles are added or saved through the GUI

### Recommended first steps

1. Open **Settings**
2. Save your Discord values if you want Discord control
3. Create your first profile from `+ Add Game`
4. Review that profile and verify paths, ports, passwords, and tokens match your machine
5. Use the dashboard to start one server at a time and confirm logs appear

## UI Guide

ECC uses a three-column layout with a bottom log strip.

The left, right, and bottom panel headers can be collapsed from the UI if you want more space while working.

### Top Bar

The top bar shows:

- CPU usage
- RAM usage
- Network throughput
- Bot state

It also includes these actions:

- `Reload UI`: rebuilds the WinForms UI without resetting running servers or timers
- `Reload Bot`: restarts the Discord listener only
- `Reload Commands`: reloads `CommandCatalog.json` and profiles from disk without resetting server timers
- `Full Restart`: restarts the whole ECC program and stops running servers
- `Settings`: opens the settings window

### Left Column: Game Profiles

This is the profile list.

Use it to:

- Select a profile for editing
- Add a new game profile with `+ Add Game`
- Remove the currently selected profile with `Remove`

When you add a profile:

1. ECC asks you to select the game server folder
2. ECC asks for a display name
3. ECC optionally asks for a separate config folder
4. ECC tries to match the game to a known template
5. ECC writes a new JSON profile into `Profiles`

### Center Column: Server Dashboard

This is the main control area. Each loaded profile gets a dashboard card.

Each card shows:

- Game name and prefix
- Online/offline badge
- PID and uptime if running
- Auto-save and scheduled restart timer line
- Per-profile `Auto-Restart` checkbox

Each card gives you:

- `Start`
- `Stop`
- `Restart`
- `Commands`
- `Config`

The `Config` button is enabled only if ECC can resolve a valid config root for that profile.

Clicking the card body opens the profile in the Profile Editor.

### Right Column: Profile Editor

This is where you edit the selected profile safely from the UI.

The profile editor is grouped into these sections when fields are present:

- `Basics`
- `Launch`
- `Logs`
- `REST`
- `RCON`
- `Restart/Safety`
- `Config`
- `Commands`
- `Misc`
- `Other`

At the bottom of the editor there is a `Profile Actions` card with:

- `Save Changes`
- `Restart Server`
- `Stop Server`

### Bottom Strip: Logs

The bottom area contains:

- `Discord` tab
- `Program Log` tab
- One game log tab per running or detected server

You can copy or clear most log tabs from their footers.

## Settings Guide

Open **Settings** from the top bar.

ECC currently stores these main settings in `Config\Settings.json`:

| Setting | Purpose |
| --- | --- |
| `BotToken` | Discord bot token used to poll the monitored channel and send bot messages |
| `WebhookUrl` | Discord webhook URL used for ECC notifications and command responses |
| `MonitorChannelId` | Discord channel ECC watches for commands |
| `CommandPrefix` | Global bot prefix, usually `!` |
| `PollIntervalSeconds` | How often ECC polls Discord for new messages |
| `EnableDebugLogging` | Enables verbose logging; changing it restarts ECC |
| `AutoSaveEnabled` | Global auto-save on/off |
| `AutoSaveIntervalMinutes` | Global auto-save interval |
| `ScheduledRestartEnabled` | Global scheduled restart on/off |
| `ScheduledRestartHours` | Global scheduled restart interval |

Important settings behavior:

- Saving settings restarts the Discord listener automatically
- Changing debug mode prompts for a full app restart and stops all servers
- Scheduled restart warnings are sent at 60, 30, 15, 10, 5, 2, and 1 minute before restart

## Discord Bot Guide

### What ECC uses Discord for

ECC has two separate Discord paths:

- Bot token + channel ID: used to read commands and send bot-originated messages
- Webhook URL: used for notifications and command acknowledgements/results

For full Discord control, configure all three:

- `BotToken`
- `MonitorChannelId`
- `WebhookUrl`

### How to create the bot

1. Go to the Discord Developer Portal.
2. Create an application.
3. Add a bot to the application.
4. Copy the bot token.
5. Turn on `Message Content Intent`.
6. Invite the bot to your server with permission to read and send messages.

### How to get the webhook

1. Open your Discord server channel settings.
2. Go to `Integrations`.
3. Create a webhook for your ECC channel.
4. Copy the webhook URL into ECC settings.

### How to get the channel ID

1. Enable Developer Mode in Discord.
2. Right-click the command channel.
3. Copy Channel ID.

### Discord command format

ECC accepts commands in either of these forms:

- `!PZ start`
- `!PZstart`

Parsing is case-insensitive, so these also work:

- `!pz start`
- `!pw status`
- `!sf save`

The format is:

```text
<global prefix><game prefix> <command>
```

Examples:

```text
!PZ start
!PZ save
!PW players
!SF status
!DZ version
```

### Discord bot behavior

ECC sends messages like:

- `[ONLINE] Etherium Command Center is now online.`
- `[OFFLINE] Etherium Command Center is now offline.`
- `[STARTING] GameName is starting...`
- `[ONLINE] GameName is now running ...`
- `[SAVING] GameName ...`
- `[STOPPED] GameName has been safely stopped.`
- `[CRASHED] GameName crashed. Restarting in ...`
- `[RESTART WARNING] ...`
- `[JOINABLE] Palworld Server Can Be Joined`

### Discord cooldowns

ECC applies cooldowns in the listener:

- `start`: 60 seconds
- `stop`: 60 seconds
- `restart`: 60 seconds
- `save`: 30 seconds
- `status`: 5 seconds
- Other commands: 10 seconds

### Safety behavior

Profile-bound Discord commands are safety-checked. Commands containing unsafe keywords like `kick`, `ban`, `teleport`, `give`, or `{player}` placeholders are blocked from being saved into profile command definitions.

Use the **Commands** window for advanced manual console commands instead of trying to turn them into profile-level Discord shortcuts.

## Profiles Guide

Profiles are JSON files in `Profiles\*.json`.

A profile tells ECC:

- What game it is
- Which executable to launch
- Which folder to launch from
- How to save and stop cleanly
- How to find logs
- Where config files live
- How to route commands
- Whether REST, RCON, telnet, or game-specific APIs are available

### How ECC creates profiles

When you use **+ Add Game**, ECC:

1. Reads the selected folder name
2. Tries to match it to a known game template
3. Auto-generates a prefix
4. Searches the folder for a likely server executable
5. Builds base commands (`start`, `stop`, `restart`, `status`)
6. Adds template-specific commands like `save` and `players`
7. Saves the profile as JSON

If a game matches a known template, ECC also applies launch args, log strategy, config paths, and extra commands from the template system.

For first-time users, the flow usually works like this:

1. Click `+ Add Game`
2. Choose the server install folder
3. If ECC recognizes the game, it builds the profile from the matching template
4. If ECC does not fully recognize the game, it asks for a display name
5. For unknown games, ECC can also ask whether you want to choose a config folder now
6. ECC saves the new profile and immediately opens it in the Profile Editor so you can review it

The auto-created profile is a starting point. Always check the generated `Executable`, `FolderPath`, save/stop commands, log paths, and network settings before relying on it in production.

### How to edit a profile

1. Click the profile in the left column.
2. Review each section in the Profile Editor.
3. Change fields as needed.
4. Click `Save Changes`.

ECC saves profile JSON back to `Profiles\<GameName>.json`.

### What each editor section means

#### Basics

Main identity and launch target:

- `GameName`: the display name of this server profile inside ECC and the base name used for the saved profile JSON file
- `Prefix`: the short game code ECC uses for game-specific logic, labels, filters, and feature routing
- `ProcessName`: the process ECC watches when checking whether the server is running
- `Executable`: the main server executable or launch target ECC starts for this profile
- `FolderPath`: the working folder ECC uses when launching the server and locating nearby files

#### Launch

Startup behavior:

- `LaunchArgs`: the raw command-line arguments ECC passes when it starts the server
- `LaunchArgState`: the saved on/off values and text values behind the launch-argument builder UI
- `MinRamGB`: the minimum memory ECC should reserve for Java-based servers that support memory launch arguments
- `MaxRamGB`: the maximum memory ECC should allow for Java-based servers that support memory launch arguments
- `AOTCache`: a game-specific cache path or cache setting used by profiles that need ahead-of-time/runtime cache support

If the game is in `Config\LaunchArgCatalog.json`, ECC shows a launch-arg editor with toggles and text fields, then rebuilds `LaunchArgs` for you.

#### Logs

Log discovery:

- `LogStrategy`: tells ECC how to find the correct server log for this game
- `ServerLogRoot`: the base folder ECC starts from when searching for log files
- `ServerLogSubDir`: an optional subfolder under the log root that narrows where ECC looks
- `ServerLogFile`: the expected log filename when the game writes to a fixed file
- `ServerLogPath`: a full direct log path used when the profile points to one exact file
- `ServerLogNote`: a human-readable note in the profile explaining any game-specific log behavior

#### REST

Used mainly by Palworld:

- `RestEnabled`: turns REST-based controls and API tests on for this profile
- `RestHost`: the hostname or IP ECC should contact for the server REST API
- `RestPort`: the port ECC should use for REST requests
- `RestPassword`: the password or API credential ECC sends to authenticated REST endpoints
- `RestProtocol`: whether ECC should connect with `http` or `https`
- `RestPollOnlyWhenRunning`: limits background REST checks to times when ECC already believes the server is running

#### RCON

Used for games that support source-style RCON:

- `RconHost`: the hostname or IP ECC should use for the RCON connection
- `RconPort`: the network port ECC should use for RCON
- `RconPassword`: the password ECC sends when opening the RCON session

#### Restart/Safety

Crash recovery and shutdown behavior:

- `EnableAutoRestart`: lets ECC automatically start the server again after an unexpected stop or crash
- `RestartDelaySeconds`: how long ECC waits before trying the restart
- `MaxRestartsPerHour`: the restart safety cap that prevents endless crash loops
- `SaveMethod`: the command route ECC uses when it needs to trigger a save before shutdown
- `SaveWaitSeconds`: how long ECC waits after sending a save command before continuing the stop sequence
- `StopMethod`: the command route ECC uses to stop the server cleanly

#### Config

Config editor roots:

- `ConfigRoot`: the main config folder ECC opens in the Config Editor for this profile
- `ConfigRoots`: a list of available config folders ECC can offer in the Config Editor dropdown

If `ConfigRoots` contains more than one valid folder, the Config Editor shows a dropdown.

#### Commands

Advanced command definitions:

- `Commands`: the built-in command definitions ECC uses for actions like start, stop, save, status, and send command
- `ExtraCommands`: additional custom commands that appear in the Commands window for this profile
- `StdinSaveCommand`: the exact text ECC sends to the server console when it performs a stdin-based save
- `StdinStopCommand`: the exact text ECC sends to the server console when it performs a stdin-based stop
- `ExeHints`: extra executable or launch hints ECC can use when deciding how to start or identify the server process

These are advanced fields. Only edit them if you understand how ECC routes command types like `Start`, `Stop`, `Restart`, `Status`, `SendCommand`, `Telnet`, `Rcon`, `Rest`, or `SatisfactoryApi`.

#### Misc

Game-specific extra fields:

- `AssetFile`: an extra game data file path used by features that need packaged or cached asset data
- `BackupDir`: the folder ECC should use for server backups or exported backup data

#### Other

Anything not grouped elsewhere lands in `Other`.

This is normal. In generated or existing profiles, `Other` can contain important fields such as:

- `TelnetHost`, `TelnetPort`, `TelnetPassword`
- `SteamAppId`
- `SatisfactoryApiHost`, `SatisfactoryApiPort`, `SatisfactoryApiToken`
- `StdinPreferWindow`, `StdinWindowProcessName`
- `DisableFileTail`
- `CaptureOutput`

### Manual JSON editing

You can edit profile JSON by hand if needed. ECC expects valid JSON and will skip malformed or unsafe profiles during load.

## Per-Game Profile Setup Guide

This section is aligned with the main game profile setups ECC is currently built around.

### 7 Days to Die (`DZ`)

Use this profile when your server is launched by `7DaysToDieServer.exe`.

Important fields:

- `Executable`: usually `7DaysToDieServer.exe`
- `FolderPath`: install folder
- `LaunchArgs`: includes `-logfile "$LOGFILE"` and `-dedicated`
- `StopMethod`: `stdin`
- `SaveMethod`: `stdin`
- `StdinSaveCommand`: `saveworld`
- `StdinStopCommand`: `shutdown`
- `TelnetHost`, `TelnetPort`, `TelnetPassword`: required for telnet tools
- `SteamAppId`: `251570`

How ECC handles it:

- Expands `$LOGFILE` into a timestamped `output_log_dedi__YYYY-MM-DD__HH-MM-SS.txt` path at launch
- Uses the newest 7DTD log file in the install folder
- Prefers Telnet for `players`, `save`, and `version`

Required game-side setup:

- Enable telnet in `serverconfig.xml`
- Make sure the telnet password in the game config matches the ECC profile

If `Test Telnet` fails, fix that before relying on advanced commands.

### Hytale (`HY`)

Use this profile for a Hytale dedicated server launched from a `.jar`.

Important fields:

- `Executable`: `HytaleServer.jar`
- `FolderPath`: server folder
- `AOTCache`
- `AssetFile`
- `BackupDir`
- `MinRamGB`, `MaxRamGB`
- `StopMethod`: `processKill`
- `SaveMethod`: `stdin`
- `StdinSaveCommand`: `backup`
- `StdinStopCommand`: `stop`

How ECC handles it:

- Builds a Java command line using RAM values
- Injects `AOTCache`, `--assets`, and `--backup-dir` when present
- Reads the newest `*_server.log` from the `logs` folder

If startup fails, verify:

- Java is installed and available on `PATH`
- `Assets.zip` and AOT cache paths are correct for your install

### Palworld (`PW`)

This profile is REST-first.

Important fields:

- `Executable`: `PalServer.exe`
- `FolderPath`: Palworld server folder
- `LaunchArgs`: should include the REST API port
- `RestEnabled`: `true`
- `RestHost`
- `RestPort`
- `RestProtocol`
- `SaveMethod`: `rest`
- `StopMethod`: process-based
- `RconPort`: optional fallback only

How ECC handles it:

- Uses REST for save, status, players, info, shutdown, and other admin calls
- Auto-loads `RESTAPIKey` and `AdminPassword` from `PalWorldSettings.ini` at runtime if needed
- Can run the log tab in activity-feed mode instead of file-tail mode
- Marks the server as joinable once REST starts responding

Commands window notes:

- `players`, `save`, `settings`, `metrics`, `stop`, `shutdown`, `broadcast`, `kickplayer`, `banplayer`, and `unbanplayer` can auto-map to REST
- `Use REST` forces REST routing
- `Test REST` is the fastest health check for this profile

If Palworld commands fail:

- Confirm the launch args expose the REST API port
- Confirm the API is listening on the same host/port as the profile
- Check the Program Log for REST request/response debug lines

### Project Zomboid (`PZ`)

This profile uses local server launch and PZ session log discovery.

Important fields:

- `Executable`: usually `StartServer64.bat`
- `FolderPath`: Project Zomboid dedicated server install folder
- `ConfigRoot`: usually `%USERPROFILE%\Zomboid\Server`
- `ServerLogRoot`: usually `%USERPROFILE%\Zomboid\Logs`
- `LogStrategy`: `PZSessionFolder`
- `SaveMethod`: `stdin`
- `StdinSaveCommand`: `save`
- `RconHost`, `RconPort`: optional if you use RCON

How ECC handles it:

- Launches through the batch file
- Tails the newest session folder in `%USERPROFILE%\Zomboid\Logs`
- Reads `*_DebugLog-server.txt` and `*_user.txt`
- Exposes Project Zomboid item and vehicle spawner tools in the Commands window
- Can use ECC cache data for the spawner, while still falling back to local Project Zomboid paths for some catalog and preview lookups

If logs do not appear:

- Make sure the profile points to the real user `Zomboid\Logs` location, not the Steam install folder

### Satisfactory (`SF`)

This profile is API-first and should be treated that way.

Important fields:

- `Executable`: `FactoryServer.exe`
- `FolderPath`: dedicated server folder
- `LaunchArgs`: usually `-Port=7777 -ReliablePort=8888 -log -unattended`
- `StopMethod`: `SatisfactoryApi`
- `SaveMethod`: `SatisfactoryApi`
- `SatisfactoryApiHost`
- `SatisfactoryApiPort`
- `SatisfactoryApiToken`
- `ConfigRoot`
- `ConfigRoots`

How ECC handles it:

- Uses the Satisfactory HTTPS API for `SaveGame`, `QueryServerState`, and `Shutdown`
- Reads logs from `FactoryGame\Saved\Logs\FactoryGame.log`
- Exposes both the AppData config root and optional install-folder `WindowsServer` config root

Required game-side setup:

- Run `server.GenerateAPIToken` in the Satisfactory server console and paste the token into the profile
- Or enable insecure local API access if that matches your environment

Commands window notes:

- `Use API` forces API routing
- `save`, `players`, `status`, `stop`, `quit`, and `exit` are auto-mapped to the API when appropriate
- `Test API` is the most useful connection test for this profile

### Valheim (`VH`)

This is the simplest built-in setup in the current ECC profile flow.

Important fields:

- `Executable`: `valheim_server.exe`
- `FolderPath`: Valheim dedicated server folder
- `LaunchArgs`: includes name, world, password, public flag, crossplay, and `-logFile`
- `LogStrategy`: `SingleFile`
- `ServerLogPath`: fixed `valheim_output.log`
- `StopMethod`: process-based
- `SaveMethod`: `none`
- `StdinPreferWindow`, `StdinWindowProcessName`: optional advanced fallback fields

How ECC handles it:

- Starts the server from the local install
- Reads a fixed log file if configured
- Uses basic process management instead of a dedicated API

Important limitation:

- Valheim does not have the same clean command/API path as Satisfactory or Palworld in this setup
- The default `players` and `save` profile commands are effectively status placeholders

## Config Editor Guide

Open the Config Editor from a dashboard card by clicking `Config`.

### What it can open

ECC lists files in the active config root with these extensions:

- `.ini`
- `.txt`
- `.cfg`
- `.json`
- `.xml`
- `.yml`
- `.yaml`
- `.properties`
- `.conf`
- `.lua`

### Generated tab vs Raw tab

The Config Editor has two tabs:

- `Generated`
- `Raw`

`Generated` is a structured editor. It currently works best for:

- `.ini`
- `.cfg`
- `.conf`
- `.properties`
- `.txt`
- Simple `.lua` files that contain editable key/value pairs

`Raw` is always available and should be used for:

- Complex Lua
- JSON
- XML
- YAML
- Files with advanced formatting that ECC cannot safely parse into fields

If ECC cannot safely structure the file, the Generated tab tells you why and you can use Raw instead.

### Generated editor features

- Section headers
- Boolean checkboxes
- Plain text fields
- Numeric validation
- Filter/search box
- Change highlighting before save

### Save behavior

When you save:

- ECC validates generated numeric fields
- ECC serializes the updated content
- ECC writes the file as UTF-8 without BOM

If a profile exposes multiple config roots, the Config Editor shows a root selector so you can switch between them without leaving the window.

Config file size limit:

- Files larger than 1 MB are blocked in the built-in editor
- Use an external editor for larger files

## Commands Window Guide

Open the Commands window from a dashboard card with `Commands`.

### What the window contains

The Commands window has:

- A catalog of known commands for the current game
- A command text box at the bottom
- A debug/output box
- Connection test buttons
- Routing toggles
- A send button

For Project Zomboid, it also includes the item and vehicle spawner panel on the right.

### Important rule

Clicking a catalog button does not execute the command immediately. It fills the command box. You still need to press `Send`.

### Routing controls

The command box can route in multiple ways:

- Automatic routing
- `Use API`
- `Use REST`

You can also force routing in the command box itself with prefixes like:

- `api: QueryServerState`
- `rest: GET /v1/api/players`

Automatic routing behavior:

- Satisfactory: API mappings are preferred for core admin actions
- Palworld: REST mappings are preferred for REST-capable actions
- 7 Days to Die: telnet is preferred if configured
- RCON is used if configured and available
- STDIN is the final fallback

### Connection tests

The Commands window includes:

- `Test Telnet`
- `Test RCON`
- `Test API`
- `Test STDIN`
- `Test PID`
- `Test REST`

What they mean:

- `Test Telnet`: attempts a safe telnet command
- `Test RCON`: attempts a safe RCON command
- `Test API`: checks the Satisfactory API
- `Test STDIN`: checks whether ECC can reach a live stdin handle or window fallback
- `Test PID`: confirms whether ECC sees the process as running
- `Test REST`: sends a REST-style info request

### Verbose Debug

Turn on `Verbose Debug` in the Commands window when you want:

- Request/response details
- Transport selection details
- Command mapping details
- Timing information

This is especially useful for REST and API troubleshooting.

### Useful command examples

Palworld REST:

```text
GET /v1/api/players
POST /v1/api/save
POST /v1/api/announce {"message":"Server restarting soon"}
```

Forced API routing:

```text
api: QueryServerState
api: SaveGame {"SaveName":"manual_backup"}
```

Generic routing:

```text
players
save
status
```

## Project Zomboid Item and Vehicle Spawner Guide

The PZ spawner appears only in the Project Zomboid Commands window.

### What it does

It builds ready-to-send admin commands for:

- `/additem`
- `/addvehicle`

### Item spawner workflow

1. Open the `PZ` Commands window.
2. Click `Refresh` beside `Online Player`.
3. Pick an online player or type a username manually.
4. Set `Count`.
5. Search for an item.
6. Select the item.
7. Use one of:
   - `Insert Into Command Box`
   - `Build /additem`
   - Double-click the item
8. Review the command and press `Send`.

ECC builds item commands like:

```text
/additem "PlayerName" "Module.ItemName" 5
```

### Vehicle spawner workflow

1. Open the `Vehicles` tab in the right panel.
2. Refresh players if needed.
3. Search for a vehicle.
4. Select it.
5. Use:
   - `Insert Into Command Box`
   - `Build /addvehicle`
   - Double-click the vehicle
6. Review the command and press `Send`.

ECC builds vehicle commands like:

```text
/addvehicle "Base.VehicleName" "PlayerName"
```

### Where the spawner data comes from

ECC first uses cached Project Zomboid spawner data under:

- `Config\AssetCache\ProjectZomboid\Catalogs`
- `Config\AssetCache\ProjectZomboid\TexturePackItems`
- `Config\AssetCache\ProjectZomboid\VehiclePreviews`

The current ECC code can also fall back to local Project Zomboid paths when cache data is missing or incomplete:

- It may read item script data from the profile `FolderPath`
- It may look at local Project Zomboid roots when resolving item preview assets
- It may rebuild parts of the item asset cache from local files if needed

Because of that, the most accurate setup advice today is:

- Keep the ECC PZ asset cache folders present when packaging ECC
- Set the PZ profile `FolderPath` to a valid Project Zomboid server folder
- Do not assume every missing cache file can be handled without local game data

### Preview image support

Preview images come from imported cache/manifests. If previews are missing but the list still loads, the command builder can still work.

You can refresh preview data with:

- `Scripts\Import-PZTexturePackItemIcons.ps1`
- `Scripts\Import-PZVehiclePreviewTextures.ps1`

## Server Control Guide

### Start a server

Use either:

- Dashboard `Start`
- Discord `!<Prefix> start`

When ECC starts a server it:

- Verifies the profile path and executable
- Expands launch args if needed
- Starts the process
- Tracks the PID
- Opens or updates log tabs

### Stop a server

Use either:

- Dashboard `Stop`
- Profile Editor `Stop Server`
- Discord `!<Prefix> stop`

ECC safe shutdown flow is:

1. Send save command if the profile has one
2. Wait `SaveWaitSeconds`
3. Send stop command or kill process according to `StopMethod`

### Restart a server

Use either:

- Dashboard `Restart`
- Profile Editor `Restart Server`
- Discord `!<Prefix> restart`

ECC restart flow is:

1. Safe shutdown
2. Short pause
3. Start again

### Auto-restart

Each dashboard card has an `Auto-Restart` checkbox.

If enabled and ECC detects an unexpected exit:

- ECC logs the crash
- Waits `RestartDelaySeconds`
- Restarts the server
- Enforces `MaxRestartsPerHour`

If the hourly crash limit is reached, ECC suspends auto-restart for that hour.

### Auto-save

Auto-save is controlled globally in Settings and can also be overridden per profile.

ECC only auto-saves if:

- Global auto-save is enabled
- The profile is running
- The profile has a valid `SaveMethod`
- The profile has not opted out

### Scheduled restarts

Scheduled restarts are controlled globally in Settings and can be opted out per profile.

ECC:

- Tracks server uptime
- Sends warning webhooks before the restart
- Performs a safe shutdown and restart when the interval is reached

## Logs Guide

ECC has three log views:

### Discord tab

Shows Discord listener activity and supports manual outbound bot messages.

You can:

- Type a message into the send box
- Press `Send`
- Clear the tab with `Clear`

### Program Log tab

Shows ECC program activity, module logs, and runtime events.

Use it for:

- Startup errors
- Path problems
- REST/API/telnet/rcon failures
- Profile load issues
- Monitor and auto-restart events

You can:

- Copy text
- Clear the tab

### Per-game log tabs

ECC creates game tabs dynamically for running or detected servers.

How ECC resolves logs depends on `LogStrategy`:

- `SingleFile`: reads a fixed log path
- `NewestFile`: opens the newest file in a folder
- `PZSessionFolder`: opens the newest Project Zomboid session folder and tails the main files
- `ValheimUserFolder`: supported by ECC even if not used by the default Valheim profile right now

Palworld note:

- Some Palworld profiles use activity-feed mode rather than file-tail mode
- In that mode, you still get REST/player/event information in the game tab

### Log files on disk

ECC writes and uses:

- `Logs\console.log`: session transcript, cleared on each launch
- `Logs\bot_YYYYMMDD_HHMMSS.log`: ECC log file created per run

## Folder Structure

```text
Etherium Command Center/
|-- Config/
|   |-- AssetCache/
|   |   `-- ProjectZomboid/
|   |-- CommandCatalog.json
|   |-- DefaultProfileTemplates.json
|   |-- LaunchArgCatalog.json
|   `-- Settings.json (generated after first save)
|-- Logs/
|   |-- console.log
|   `-- bot_*.log
|-- Modules/
|   |-- DiscordListener.psm1
|   |-- GUI.psm1
|   |-- Logging.psm1
|   |-- ProfileManager.psm1
|   `-- ServerManager.psm1
|-- Profiles/
|   `-- *.json (generated when profiles are created)
|-- Scripts/
|   |-- Import-PZTexturePackItemIcons.ps1
|   `-- Import-PZVehiclePreviewTextures.ps1
|-- Launch.ps1
|-- README.md
`-- Start.bat
```

## Required Files and Generated Files

### Files that should ship with ECC

| Path | Required | Purpose |
| --- | --- | --- |
| `Start.bat` | Yes | Easiest launcher for users |
| `Launch.ps1` | Yes | Main app entry point |
| `Modules\Logging.psm1` | Yes | Logging |
| `Modules\ProfileManager.psm1` | Yes | Profile creation, loading, safety, launch arg handling |
| `Modules\ServerManager.psm1` | Yes | Server lifecycle, command routing, APIs, monitoring |
| `Modules\DiscordListener.psm1` | Yes | Discord polling and webhooks |
| `Modules\GUI.psm1` | Yes | WinForms UI |
| `Config\CommandCatalog.json` | Yes for Commands window | Populates the command catalog UI |
| `Config\LaunchArgCatalog.json` | Recommended | Enables structured launch-arg editing |
| `Config\DefaultProfileTemplates.json` | Recommended | Applies bundled defaults to newly generated profiles |

### Files ECC creates automatically

| Path | When created | Notes |
| --- | --- | --- |
| `Config\` | On launch if missing | ECC creates the folder |
| `Profiles\` | On launch if missing | ECC creates the folder |
| `Logs\` | On launch if missing | ECC creates the folder |
| `Config\Settings.json` | When you save settings | ECC can run with defaults before this exists; opening Settings does not create it until save |
| `Logs\console.log` | Every launch | Cleared at start so it reflects the current session |
| `Logs\bot_*.log` | Every launch | Main ECC log file for the run |
| `Profiles\*.json` | When you add/save profiles | One JSON file per profile |

### Files ECC creates for Project Zomboid caching

| Path | How it appears | Purpose |
| --- | --- | --- |
| `Config\AssetCache\ProjectZomboid\Catalogs\*.json` | When ECC builds or saves PZ catalogs | Item and vehicle catalog cache |
| `Config\AssetCache\ProjectZomboid\item-texture-manifest.json` | When item texture import runs | Maps item names to cached preview icons |
| `Config\AssetCache\ProjectZomboid\vehicle-texture-manifest.json` | When vehicle preview import runs | Maps vehicle names to cached preview images |
| `Config\AssetCache\ProjectZomboid\TexturePackItems\*.png` | Generated by item import | Cached item preview images |
| `Config\AssetCache\ProjectZomboid\VehiclePreviews\*.png` | Generated by vehicle import | Cached vehicle preview images |

### If a file is missing

- Missing `Settings.json`: ECC starts with defaults, still opens the Settings window, and creates the file when you click `Save Settings`
- Missing `Profiles\*.json`: ECC still starts; create profiles with `+ Add Game`
- Missing `CommandCatalog.json`: the Commands window will not load command catalog content
- Missing `LaunchArgCatalog.json`: ECC falls back to plain launch args instead of structured launch arg toggles
- Missing PZ asset cache: the PZ spawner windows may be incomplete, and ECC may need local Project Zomboid files from the profile folder or discovered roots to rebuild or fill missing data

## FAQ

### Can I run ECC without Discord?

Yes. The local dashboard, profiles, config editor, commands window, and logs still work. Discord is optional.

### Do I need both a bot token and a webhook URL?

If you want full Discord control and responses, yes. The bot reads commands from the monitored channel, while ECC notifications and command responses use the webhook path.

### Can I add games that are not already pre-created in `Profiles`?

Yes. Use `+ Add Game`, then edit the generated profile manually if ECC does not fully recognize the game.

### Why is the Config button disabled?

ECC could not resolve a valid config root for that profile.

### Why do command tests say the server is offline?

Most transport tests require the target server to already be running.

### Why did ECC restart when I changed debug mode?

That is intentional. Changing `EnableDebugLogging` triggers a full restart and stops running servers.

### Why are some advanced fields in `Other`?

Because the profile editor only groups common sections explicitly. Advanced or game-specific keys still show up under `Other` so you can edit them.

### Can I edit profile JSON outside the UI?

Yes. ECC loads profile JSON from disk at startup and can reload it with `Reload Commands`.

### Does ECC support large config files in the built-in editor?

Not if they exceed 1 MB. Use an external editor for larger files

## Troubleshooting and Debug Questions

### ECC does not launch

Check:

- You launched `Start.bat` or `Launch.ps1`
- PowerShell 5.1 exists
- You allowed elevation
- `Modules\*.psm1` files are present

### A server will not start

Check:

- `FolderPath` is correct
- `Executable` exists inside that folder
- `LaunchArgs` match the game
- Program Log shows a real process start attempt

### Stop or restart is unsafe or too abrupt

Check:

- `SaveMethod`
- `StdinSaveCommand`
- `SaveWaitSeconds`
- `StopMethod`
- Game-specific APIs or telnet credentials

### No logs are appearing

Check:

- `LogStrategy`
- `ServerLogRoot`
- `ServerLogPath`
- `ServerLogSubDir`
- `ServerLogNote`

For Project Zomboid specifically:

- The logs are usually under `%USERPROFILE%\Zomboid\Logs`, not the Steam install folder

### Discord bot sees commands but does not answer correctly

Check:

- `BotToken`
- `WebhookUrl`
- `MonitorChannelId`
- `CommandPrefix`
- `Message Content Intent`
- That you are using a loaded game prefix like `PZ`, `PW`, `SF`, `DZ`, `HY`, or `VH`

### 7 Days to Die telnet fails

Check:

- `TelnetEnabled=true` in `serverconfig.xml`
- `TelnetPort` matches the profile
- `TelnetPassword` matches the profile
- The server is already running before you use `Test Telnet`

### Palworld REST fails

Check:

- `RestEnabled`
- `RestHost`
- `RestPort`
- Launch args include the REST API port
- The REST API is actually listening
- Program Log for REST debug entries

### Satisfactory API fails

Check:

- `SatisfactoryApiHost`
- `SatisfactoryApiPort`
- `SatisfactoryApiToken`
- Whether you generated the token with `server.GenerateAPIToken`
- Whether the server is already running when you press `Test API`

### PZ item or vehicle spawner is empty

Check:

- `Config\AssetCache\ProjectZomboid` exists
- The ECC cache folders include the expected item catalogs, vehicle catalogs, and preview assets
- The PZ profile `FolderPath` points to a valid Project Zomboid server folder
- Local Project Zomboid files needed for fallback catalog or preview loading are available if cache data is incomplete

### The Generated config tab is unavailable

That is expected when:

- The file type is not yet supported for structured editing
- The file uses advanced formatting ECC cannot safely rewrite
- The Lua file contains nested tables or code blocks

Use the Raw tab in those cases.

### What to collect before opening an issue or debugging a broken profile

Have these ready:

- The game prefix and profile name
- The exact action you took
- The exact command you sent
- A screenshot of the relevant Profile Editor sections
- Relevant lines from Program Log
- Results from `Test Telnet`, `Test RCON`, `Test API`, `Test REST`, or `Test STDIN`
- The current `LogStrategy`, config root, and server path values

## Security Notes Before You Push a PR

Before publishing this repo or opening a public PR, scrub or replace:

- Discord bot tokens
- Webhook URLs
- Channel IDs if you do not want them public
- RCON passwords
- Telnet passwords
- Satisfactory API tokens
- Any public IPs or hostnames you do not want exposed

Check these locations especially:

- `Config\Settings.json`
- `Profiles\*.json`

ECC can load blank placeholders for secrets just fine, so do not publish live credentials.

## License

Private use during development. Distribution terms pending.

## Credits

Built by DarkjesusMN and AI contributors.

<img width="1919" height="1077" alt="{B332CCE4-B67F-4803-B2B1-187F2088B4F0}" src="https://github.com/user-attachments/assets/8c8ebe02-c32c-473a-890f-c1c79a860341" />

