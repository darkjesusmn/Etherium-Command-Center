# Etherium Command Center (ECC)

Current GUI version: `v0.9`

Etherium Command Center, or ECC, is a Windows desktop app for running and managing local dedicated game servers from one control panel.

ECC can start servers, stop servers, restart servers, show logs, edit server config files, send commands, talk to Discord, run auto-save, run scheduled restarts, and help with game-specific tools such as the Project Zomboid item and vehicle command builder.

This README is written for a new user. You do not need to know PowerShell to use ECC, but you should be comfortable finding your server install folders and reading simple settings.

## Table of Contents

- [Quick Answer](#quick-answer)
- [What ECC Can Do](#what-ecc-can-do)
- [Supported Games](#supported-games)
- [Requirements](#requirements)
- [Download and Install](#download-and-install)
- [First Launch](#first-launch)
- [First Profile Walkthrough](#first-profile-walkthrough)
- [Main Window Guide](#main-window-guide)
- [Settings Guide](#settings-guide)
- [Discord Setup](#discord-setup)
- [Profiles Guide](#profiles-guide)
- [Complete Profile Field Reference](#complete-profile-field-reference)
- [Per-Game Setup Notes](#per-game-setup-notes)
- [Commands Window Guide](#commands-window-guide)
- [Config Editor Guide](#config-editor-guide)
- [Project Zomboid Spawner Guide](#project-zomboid-spawner-guide)
- [Logs Guide](#logs-guide)
- [Files and Folders](#files-and-folders)
- [What Should and Should Not Be Uploaded](#what-should-and-should-not-be-uploaded)
- [Troubleshooting](#troubleshooting)
- [Security Notes](#security-notes)
- [License](#license)

## Quick Answer

If you just want to try ECC:

1. Download the repo as a ZIP from GitHub.
2. Extract the ZIP to a normal folder, such as `C:\ECC`.
3. Do not run ECC from inside the ZIP file.
4. Right-click `Start.bat`.
5. Click `Run as administrator`.
6. Open `Settings`.
7. Save Discord settings if you want Discord control.
8. Click `+ Add Game`.
9. Pick your game type.
10. Pick the real server folder.
11. Review the profile before trusting it.
12. Start one server and confirm logs, save, stop, and status work.

ECC can run without Discord. Discord is optional.

## What ECC Can Do

ECC is built for one Windows host that runs local dedicated game servers.

It can:

- Start, stop, and restart game servers.
- Show a dashboard card for each server profile.
- Track server process state and uptime.
- Show ECC program logs and per-game logs.
- Edit common server config files.
- Save and load server profiles from JSON.
- Run global auto-save.
- Run scheduled restarts.
- Restart crashed servers when auto-restart is enabled.
- Send Discord notifications through a webhook.
- Read Discord commands from one Discord channel.
- Route commands through stdin, RCON, telnet, REST, or the Satisfactory API depending on the game.
- Show command catalogs for supported games.
- Build Project Zomboid item and vehicle admin commands.
- Reload the UI, bot, commands, and profiles without always closing the whole program.

ECC is not:

- A cloud server host.
- A Linux server manager.
- A multi-machine server cluster tool.
- A replacement for owning or installing the actual game server files.

## Supported Games

ECC has built-in profile templates for these games:

| Game | Prefix | Main control path | Notes |
| --- | --- | --- | --- |
| 7 Days to Die | `DZ` | Telnet and stdin | Telnet is used for players, save, and version checks. |
| Hytale | `HY` | stdin | Java/JAR based. Uses `Assets.zip` and AOT cache settings when available. |
| Minecraft | `MC` | stdin and RCON | Uses `server.jar`, `logs\latest.log`, stdin save/stop, and RCON player list. |
| Palworld | `PW` | REST API | Uses Palworld REST endpoints for save, players, status, and shutdown. |
| Project Zomboid | `PZ` | stdin | Uses dated PZ log folders and includes an item/vehicle command builder. |
| Satisfactory | `SF` | Satisfactory API | Uses the server HTTPS API for save, status, players, and shutdown. |
| Valheim | `VH` | process control | Basic process launch and log tailing. |

You can also make custom profiles, but built-in templates are easier.

## Requirements

You need:

- Windows 10 or Windows 11.
- Windows PowerShell 5.1.
- Administrator rights.
- Local access to the dedicated server folders you want ECC to manage.
- The game server files already installed.

Optional:

- A Discord bot token.
- A Discord channel ID.
- A Discord webhook URL.

Game-specific requirements:

- 7 Days to Die needs telnet enabled in `serverconfig.xml` if you want telnet commands.
- Minecraft needs `eula.txt` accepted and RCON enabled if you want RCON player queries.
- Palworld needs REST enabled and reachable.
- Satisfactory needs a valid API token or local insecure API access.
- Project Zomboid works best when ECC can read local PZ server/client files.
- Hytale support depends on your local Hytale server folder and files.

## Download and Install

### Option A: Download From GitHub

Use this if you are not using Git.

1. Open the GitHub repo page.
2. Click the green `Code` button.
3. Click `Download ZIP`.
4. Extract the ZIP to a normal folder.
5. Example folder:

```text
C:\Etherium Command Center
```

6. Open the extracted folder.
7. Right-click `Start.bat`.
8. Click `Run as administrator`.

Important:

- Do not run ECC from inside the ZIP file.
- Keep the folder structure together.
- Do not move only `Start.bat` by itself.

### Option B: Clone With Git

Use this if you already know Git:

```powershell
git clone https://github.com/darkjesusmn/Etherium-Command-Center.git
cd Etherium-Command-Center
```

Then right-click `Start.bat` and run it as administrator.

### Option C: Packaged Release ZIP

If a release ZIP is provided later, use that instead of the source ZIP. A release ZIP can include a cleaner folder layout and optional add-ons.

## First Launch

`Start.bat` launches `Launch.ps1`.

On launch, ECC will:

- Ask for administrator rights if needed.
- Unblock bundled `.ps1`, `.psm1`, and `.json` files.
- Set the current user's PowerShell execution policy to `RemoteSigned`.
- Create missing `Config`, `Profiles`, and `Logs` folders.
- Load the ECC modules.
- Start the Discord listener runspace.
- Start the server monitor runspace.
- Open the GUI.

Fresh installs do not need:

- `Config\Settings.json`
- `Profiles\*.json`
- `Logs\*.log`

ECC creates these when needed.

If Windows asks whether you trust the script, allow it only if you downloaded ECC from a source you trust.

## First Profile Walkthrough

This is the safest way to set up your first server.

1. Start ECC as administrator.
2. Click `+ Add Game`.
3. Pick the game type.
4. Pick the folder where that dedicated server is installed.
5. Let ECC create the profile.
6. Click the new profile in the left list.
7. Check these fields first:

| Field | What to check |
| --- | --- |
| `GameName` | The name shown in ECC. |
| `Prefix` | The short code used for Discord commands. |
| `FolderPath` | The real server folder. |
| `Executable` | The real server executable or start script. |
| `LaunchArgs` | The command-line options used to start the server. |
| `SaveMethod` | How ECC saves the server. |
| `StopMethod` | How ECC stops the server. |
| `LogStrategy` | How ECC finds logs. |
| `EnableAutoRestart` | Whether ECC should restart this server after a crash. |

8. Click `Save Changes`.
9. Start only this one server.
10. Confirm:

- The server starts.
- ECC sees it as running.
- Logs appear.
- `Save` works if the game supports it.
- `Stop` works safely.
- `Players` or `Status` works if supported.

Do not use `Start All` until each profile has been tested by itself.

## Main Window Guide

ECC uses a three-column layout with logs at the bottom.

### Top Bar

The top bar shows:

- CPU usage.
- RAM usage.
- Network use.
- Disk/free-space status.
- Total trusted player count.
- Bot state.

Top bar buttons:

| Button | What it does |
| --- | --- |
| `Start All` | Starts profiles ECC thinks are ready. |
| `Stop All` | Stops all running servers using each profile's stop method. |
| `Reload UI` | Rebuilds the UI without stopping running servers. |
| `Reload Bot` | Restarts the Discord listener. |
| `Reload Commands` | Reloads command catalog and profiles from disk. |
| `Full Restart` | Stops running servers, closes ECC, and relaunches it. |
| `Settings` | Opens the settings window. |

### Left Column: Game Profiles

The left column shows all loaded profiles.

Use it to:

- Select a profile.
- Add a profile with `+ Add Game`.
- Remove the selected profile.

Profiles are stored in:

```text
Profiles\*.json
```

### Center Column: Server Dashboard

Each profile gets a server card.

Cards can show:

- Game name.
- Prefix.
- Running/stopped state.
- Health state.
- PID.
- Uptime.
- Player/status source.
- Auto-save timer.
- Scheduled restart timer.
- Auto-restart checkbox.

Common card buttons:

| Button | What it does |
| --- | --- |
| `Start` | Starts the selected server. |
| `Stop` | Saves then stops the selected server when possible. |
| `Restart` | Stops then starts the selected server. |
| `Commands` | Opens the command tool for that game. |
| `Config` | Opens the config editor if ECC can find a config folder. |

Clicking a card body opens that profile in the editor.

### Right Column: Profile Editor

The profile editor is where you change how ECC treats a server.

Sections may include:

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

At the bottom, `Profile Actions` can:

- Save profile changes.
- Restart the server.
- Stop the server.

### Bottom Strip: Logs

The bottom strip has:

- `Discord` tab.
- `Program Log` tab.
- One tab per game server log when available.

Use `Program Log` first when something breaks.

## Settings Guide

Open settings from the top bar.

Settings are saved to:

```text
Config\Settings.json
```

This file is created when you click `Save Settings`.

### Settings Fields

| Setting | Simple meaning |
| --- | --- |
| `BotToken` | Discord bot token used to read commands. |
| `WebhookUrl` | Discord webhook used to send notifications and command responses. |
| `MonitorChannelId` | Discord channel ECC watches for commands. |
| `CommandPrefix` | Global Discord command prefix, usually `!`. |
| `PollIntervalSeconds` | How often ECC checks Discord for new messages. |
| `EnableDebugLogging` | Turns on detailed logs. Changing it may require restart. |
| `EnablePerformanceDebugMode` | Turns on performance tracing for lag investigation. |
| `AutoSaveEnabled` | Turns global auto-save on or off. |
| `AutoSaveIntervalMinutes` | Minutes between auto-saves. |
| `ScheduledRestartEnabled` | Turns scheduled restarts on or off. |
| `ScheduledRestartHours` | Hours between scheduled restarts. |
| `WindowWidth` | Saved ECC window width. |
| `WindowHeight` | Saved ECC window height. |
| `WindowX` | Saved window X position. |
| `WindowY` | Saved window Y position. |
| `WindowState` | Saved window state, such as normal or maximized. |

### Important Settings Behavior

- Discord is optional.
- Saving settings restarts the Discord listener.
- Debug logging can make logs much larger.
- Performance Trace Mode is meant for long-session lag testing.
- Scheduled restart warnings are sent before a restart when webhooks are set up.

## Discord Setup

ECC uses two Discord paths:

- Bot token and channel ID: read commands from Discord.
- Webhook URL: send notifications and command results.

For full Discord control, set all three:

- `BotToken`
- `MonitorChannelId`
- `WebhookUrl`

### Create a Discord Bot

1. Go to the Discord Developer Portal.
2. Create an application.
3. Add a bot.
4. Copy the bot token.
5. Turn on `Message Content Intent`.
6. Invite the bot to your server.
7. Give it permission to read and send messages in the command channel.

### Get the Channel ID

1. Open Discord settings.
2. Enable Developer Mode.
3. Right-click the channel ECC should watch.
4. Click `Copy Channel ID`.
5. Paste it into `MonitorChannelId`.

### Create a Webhook

1. Open the Discord channel settings.
2. Go to `Integrations`.
3. Create a webhook.
4. Copy the webhook URL.
5. Paste it into `WebhookUrl`.

### Discord Command Format

ECC accepts commands like:

```text
!PZ start
!PZ save
!PW players
!SF status
!DZ version
```

Format:

```text
<global prefix><game prefix> <command>
```

With the default global prefix `!`, the Project Zomboid profile prefix `PZ` becomes:

```text
!PZ start
```

ECC also accepts compact commands like:

```text
!PZstart
```

Commands are case-insensitive.

### Discord Cooldowns

ECC uses cooldowns so one person cannot spam commands too fast:

| Command type | Cooldown |
| --- | --- |
| `start` | 60 seconds |
| `stop` | 60 seconds |
| `restart` | 60 seconds |
| `save` | 30 seconds |
| `status` | 5 seconds |
| Other commands | 10 seconds |

### Discord Safety

ECC blocks unsafe saved profile commands that include dangerous actions such as:

- `kick`
- `ban`
- `teleport`
- `give`
- `{player}` placeholder commands

Use the Commands window for manual admin commands.

## Profiles Guide

A profile is one JSON file that tells ECC how to control one server.

Profiles live here:

```text
Profiles\*.json
```

A profile tells ECC:

- What game this is.
- What short prefix to use.
- Which file starts the server.
- Which folder to start from.
- How to save.
- How to stop.
- How to find logs.
- Which API, RCON, telnet, or REST settings to use.
- Which commands are allowed.
- Which config folder to edit.
- Whether auto-restart should run.

### How ECC Creates a Profile

When you click `+ Add Game`, ECC:

1. Asks for the game type.
2. Opens a folder picker.
3. Reads the selected folder.
4. Looks for known executable names.
5. Applies a game template if one matches.
6. Creates default commands.
7. Creates or updates `Profiles\*.json`.
8. Opens the new profile for review.

The generated profile is a starting point. You still need to check it.

### Editing a Profile

1. Click a profile on the left.
2. Edit fields on the right.
3. Click `Save Changes`.
4. Use `Reload Commands` if you edited JSON by hand.

Bad JSON will stop that profile from loading.

## Complete Profile Field Reference

Not every profile uses every field. Fields only appear when a profile needs them or when they were saved before.

### Basics

| Field | Meaning |
| --- | --- |
| `GameName` | Display name shown in ECC. Also used for the profile file name. |
| `Prefix` | Short code used for Discord and internal routing, such as `PZ` or `MC`. |
| `ProcessName` | Process name ECC watches to decide if the server is running. |
| `Executable` | File ECC starts, such as `server.jar`, `PalServer.exe`, or `StartServer64.bat`. |
| `FolderPath` | Working folder where ECC starts the server. Usually the server install folder. |

### Launch

| Field | Meaning |
| --- | --- |
| `LaunchArgs` | Raw command-line arguments used when starting the server. |
| `LaunchArgState` | Saved toggle/value state for ECC's launch argument editor. |
| `MinRamGB` | Minimum RAM setting used by Java-style launch handling where supported. |
| `MaxRamGB` | Maximum RAM setting used by Java-style launch handling where supported. |
| `AOTCache` | Hytale-style ahead-of-time cache file or runtime cache setting. |
| `AssetFile` | Extra asset file used by a game profile, such as `Assets.zip`. |
| `BackupDir` | Backup folder used by profiles that support backup commands. |
| `SteamAppId` | Steam app ID set before launch for games that need it. |
| `ExeHints` | Extra executable names ECC can search for when creating or fixing a profile. |

### Logs

| Field | Meaning |
| --- | --- |
| `LogStrategy` | How ECC finds logs. See log strategy table below. |
| `ServerLogRoot` | Base folder for log searching. |
| `ServerLogSubDir` | Subfolder under the main server folder or log root. |
| `ServerLogFile` | Expected log file name or wildcard. |
| `ServerLogPath` | Direct full path to one log file. |
| `ServerLogNote` | Human note explaining this game's log behavior. |
| `DisableFileTail` | If true, ECC does not tail a normal log file for that profile. |
| `CaptureOutput` | If true, ECC tries to capture process stdout/stderr when starting the server. |

Log strategy values:

| Value | Meaning |
| --- | --- |
| `SingleFile` | Tail one fixed log file. |
| `NewestFile` | Find and tail the newest matching file. |
| `PZSessionFolder` | Find newest Project Zomboid session log folder. |
| `ValheimUserFolder` | Valheim-specific log lookup support. |

### REST

Used mostly by Palworld.

| Field | Meaning |
| --- | --- |
| `RestEnabled` | Turns REST support on for this profile. |
| `RestHost` | Hostname or IP for REST requests. |
| `RestPort` | REST API port. |
| `RestPassword` | Password or API credential for REST requests. |
| `RestProtocol` | `http` or `https`. |
| `RestPollOnlyWhenRunning` | Only poll REST while ECC thinks the server is running. |

### RCON

Used by games that support RCON, such as Minecraft.

| Field | Meaning |
| --- | --- |
| `RconHost` | RCON host or IP. |
| `RconPort` | RCON port. |
| `RconPassword` | RCON password. |

### Telnet

Used mainly by 7 Days to Die.

| Field | Meaning |
| --- | --- |
| `TelnetHost` | Telnet host or IP. |
| `TelnetPort` | Telnet port. |
| `TelnetPassword` | Telnet password from the game config. |

### Satisfactory API

Used by Satisfactory.

| Field | Meaning |
| --- | --- |
| `SatisfactoryApiHost` | Hostname or IP for the Satisfactory API. |
| `SatisfactoryApiPort` | Satisfactory API port. |
| `SatisfactoryApiToken` | Token created by the Satisfactory server. |

### Restart and Safety

| Field | Meaning |
| --- | --- |
| `EnableAutoRestart` | If true, ECC can restart the server after an unexpected crash. |
| `RestartDelaySeconds` | Seconds ECC waits before trying auto-restart. |
| `MaxRestartsPerHour` | Safety limit to avoid endless crash loops. |
| `BlockStartIfRamPercentUsed` | Blocks server start if system RAM use is above this percent. `0` disables it. |
| `BlockStartIfFreeRamBelowGB` | Blocks server start if free RAM is below this many GB. `0` disables it. |
| `StartupTimeoutSeconds` | How long ECC waits for a new server to look alive. |
| `ShutdownIfNoPlayersAfterStartupMinutes` | Stops the server if no players join after startup for this many minutes. `0` disables it. |
| `ShutdownIfEmptyAfterLastPlayerLeavesMinutes` | Stops the server after the last player leaves for this many minutes. `0` disables it. |
| `SaveMethod` | How ECC saves the server. |
| `SaveWaitSeconds` | Seconds ECC waits after saving before stopping. |
| `StopMethod` | How ECC stops the server. |

Save method values:

| Value | Meaning |
| --- | --- |
| `none` | No save command is sent. |
| `stdin` | Send text to the server console/stdin. |
| `http` | Use an HTTP-style save path. |
| `rest` | Use REST. |
| `SatisfactoryApi` | Use the Satisfactory API. |

Stop method values:

| Value | Meaning |
| --- | --- |
| `processKill` | Kill the process. Fast but less graceful. |
| `processName` | Stop/kill by process name. |
| `stdin` | Send a stop command to stdin. |
| `ctrlc` | Try Ctrl+C style shutdown. |
| `http` | Use an HTTP-style stop path. |
| `SatisfactoryApi` | Use the Satisfactory API. |

### Stdin and Window Fallback

| Field | Meaning |
| --- | --- |
| `StdinSaveCommand` | Text ECC sends to save through stdin. |
| `StdinStopCommand` | Text ECC sends to stop through stdin. |
| `StdinPreferWindow` | Prefer sending to a server window when normal stdin is not reliable. |
| `StdinWindowProcessName` | Process name used for the window fallback. |

### Config Editor

| Field | Meaning |
| --- | --- |
| `ConfigRoot` | Main folder ECC opens in the Config Editor. |
| `ConfigRoots` | Multiple config folders ECC can offer in a dropdown. |

### Commands and Extra Commands

| Field | Meaning |
| --- | --- |
| `Commands` | Main saved commands used by Discord and ECC. |
| `ExtraCommands` | Extra commands shown in the Commands window. |

Command object fields:

| Field | Meaning |
| --- | --- |
| `Type` | How ECC routes the command. |
| `Command` | Raw command text to send. |
| `Endpoint` | REST endpoint path. |
| `Method` | REST method, such as `GET` or `POST`. |
| `Function` | API function name, mostly for Satisfactory. |

Common command `Type` values:

| Type | Meaning |
| --- | --- |
| `Start` | Start the server. |
| `Stop` | Stop the server. |
| `Restart` | Restart the server. |
| `Status` | Show ECC status. |
| `SendCommand` | Send a raw command through the best route. |
| `Telnet` | Send through telnet. |
| `Rcon` | Send through RCON. |
| `Rest` | Send through REST. |
| `SatisfactoryApi` | Send through the Satisfactory API. |

### Other

If a field does not fit a main section, ECC shows it under `Other`.

That does not mean the field is useless. It only means it is less common or game-specific.

## Per-Game Setup Notes

### 7 Days to Die (`DZ`)

Important fields:

- `Executable`: usually `7DaysToDieServer.exe`.
- `FolderPath`: dedicated server folder.
- `LaunchArgs`: usually includes `-logfile "$LOGFILE"` and `-dedicated`.
- `StopMethod`: usually `stdin`.
- `SaveMethod`: usually `stdin` or telnet command routing.
- `StdinSaveCommand`: `saveworld`.
- `StdinStopCommand`: `shutdown`.
- `TelnetHost`: usually `127.0.0.1`.
- `TelnetPort`: usually `8081`.
- `TelnetPassword`: must match `serverconfig.xml`.
- `SteamAppId`: `251570`.

Check before use:

- Telnet is enabled in `serverconfig.xml`.
- Telnet password matches ECC.
- `Test Telnet` works while the server is running.
- Logs appear after launch.

### Hytale (`HY`)

Important fields:

- `Executable`: usually `HytaleServer.jar`.
- `FolderPath`: Hytale server folder.
- `ProcessName`: `java`.
- `AOTCache`: usually `HytaleServer.aot`.
- `AssetFile`: usually `Assets.zip`.
- `BackupDir`: usually `backup`.
- `SaveMethod`: `stdin`.
- `StdinSaveCommand`: `backup`.
- `StdinStopCommand`: `stop`.
- `LogStrategy`: usually `NewestFile`.
- `ServerLogSubDir`: usually `logs`.

Check before use:

- Java is installed.
- The server JAR exists.
- `Assets.zip` exists if your server requires it.
- The AOT cache setting matches your server setup.
- Logs appear in the configured logs folder.

### Minecraft (`MC`)

Important fields:

- `Executable`: usually `server.jar`.
- `FolderPath`: Minecraft server folder.
- `ProcessName`: `java`.
- `LaunchArgs`: usually `nogui`.
- `SaveMethod`: `stdin`.
- `StdinSaveCommand`: `save-all`.
- `StdinStopCommand`: `stop`.
- `RconHost`: usually `127.0.0.1`.
- `RconPort`: usually `25575`.
- `RconPassword`: must match `server.properties`.
- `LogStrategy`: `SingleFile`.
- `ServerLogSubDir`: `logs`.
- `ServerLogFile`: `latest.log`.

Check before use:

- `eula.txt` is accepted.
- `server.jar` exists.
- RCON is enabled if you want player list commands.
- `server.properties` has the same RCON port and password as ECC.
- `logs\latest.log` updates while the server is running.

### Palworld (`PW`)

Important fields:

- `Executable`: `PalServer.exe`.
- `ProcessName`: `PalServer-Win64-Shipping`.
- `LaunchArgs`: should include the correct game, query, and REST ports.
- `RestEnabled`: true for REST control.
- `RestHost`: usually `127.0.0.1`.
- `RestPort`: usually `8212`.
- `RestProtocol`: usually `http`.
- `RestPassword`: password/API credential if required.
- `SaveMethod`: `rest`.
- `StopMethod`: often REST or process-based depending on setup.
- `DisableFileTail`: often true because ECC can use an activity feed style view.

Check before use:

- REST is enabled by your Palworld server setup.
- `Test REST` works while the server is running.
- `players`, `save`, and `status` work from the Commands window.

### Project Zomboid (`PZ`)

Important fields:

- `Executable`: usually `StartServer64.bat`.
- `FolderPath`: Project Zomboid dedicated server folder.
- `ProcessName`: often `ZuluPlatformx64Architecture`.
- `SaveMethod`: `stdin`.
- `StdinSaveCommand`: `save`.
- `StopMethod`: often `processKill` or another configured stop path.
- `LogStrategy`: `PZSessionFolder`.
- `ConfigRoot`: usually `%USERPROFILE%\Zomboid\Server`.
- `RconPort`: often `27015` if RCON is used.

Check before use:

- The PZ server starts from ECC.
- Logs appear from `%USERPROFILE%\Zomboid\Logs`.
- The `players` and `save` commands work.
- The config root points to the folder with your PZ server config files.

### Satisfactory (`SF`)

Important fields:

- `Executable`: `FactoryServer.exe`.
- `ProcessName`: `FactoryServer-Win64-Shipping`.
- `LaunchArgs`: ports and unattended/log options.
- `StopMethod`: `SatisfactoryApi`.
- `SaveMethod`: `SatisfactoryApi`.
- `SatisfactoryApiHost`: usually `127.0.0.1`.
- `SatisfactoryApiPort`: usually the server API port.
- `SatisfactoryApiToken`: token from the server.
- `LogStrategy`: usually `NewestFile`.
- `ServerLogSubDir`: `FactoryGame\Saved\Logs`.
- `ServerLogFile`: `FactoryGame.log`.

Check before use:

- Start the server once.
- Generate an API token with `server.GenerateAPIToken`.
- Paste the token into the profile.
- Run `Test API`.
- Confirm `save`, `players`, and `status` work.

### Valheim (`VH`)

Important fields:

- `Executable`: `valheim_server.exe`.
- `ProcessName`: `valheim_server`.
- `LaunchArgs`: server name, world, password, port, public flag, crossplay, and log file.
- `SaveMethod`: usually `none`.
- `StopMethod`: usually `processKill`.
- `LogStrategy`: `SingleFile`.
- `ServerLogSubDir`: `logs`.
- `ServerLogFile`: `valheim_output.log`.
- `StdinPreferWindow`: optional fallback.
- `StdinWindowProcessName`: optional fallback.

Check before use:

- Your world name is correct.
- Your password is correct.
- Your port is correct.
- The log file path exists or is created on launch.
- You understand that Valheim support is more basic than API-based games.

## Commands Window Guide

Open it with a server card's `Commands` button.

The Commands window can include:

- A command catalog.
- A command text box.
- Send button.
- Debug/output box.
- Transport test buttons.
- Route toggles.
- PZ spawner panel for Project Zomboid.

Important rule:

Clicking a catalog command usually fills the command box. It does not always send the command immediately. Review it, then click `Send`.

### Routing

ECC can route commands through:

- Automatic route selection.
- stdin.
- RCON.
- telnet.
- REST.
- Satisfactory API.

Forced prefixes can be used in the command box:

```text
api: QueryServerState
rest: GET /v1/api/players
```

### Test Buttons

| Button | What it checks |
| --- | --- |
| `Test Telnet` | Telnet connection and auth. |
| `Test RCON` | RCON connection and auth. |
| `Test API` | Satisfactory API connection. |
| `Test STDIN` | Whether ECC can send to the running server. |
| `Test PID` | Whether ECC sees the process. |
| `Test REST` | REST API connection. |

Tests usually require the server to already be running.

### Verbose Debug

Turn this on when you need details about:

- Which route ECC chose.
- What request was sent.
- What response came back.
- Why a command failed.

## Config Editor Guide

Open it with a server card's `Config` button.

The button is enabled only when ECC can find a valid config folder.

ECC can list these file types:

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

### Generated Tab

The `Generated` tab tries to make a simple field editor.

It works best for:

- INI files.
- CFG files.
- CONF files.
- Properties files.
- Simple text config files.
- Simple Lua key/value files.

### Raw Tab

The `Raw` tab is always the safer option for complex files.

Use Raw for:

- JSON.
- XML.
- YAML.
- Complex Lua.
- Files with comments or special formatting ECC cannot safely rebuild.

### Save Behavior

When saving, ECC:

- Validates simple numeric fields.
- Writes the updated file.
- Uses UTF-8 without BOM.

Files larger than 1 MB are blocked in the built-in editor. Use a normal text editor for larger files.

## Project Zomboid Spawner Guide

The Project Zomboid Commands window includes an item and vehicle command builder.

It builds commands like:

```text
/additem "PlayerName" "Module.ItemName" 5
/addvehicle "Base.VehicleName" "PlayerName"
```

### Item Workflow

1. Open the PZ Commands window.
2. Refresh the online player list.
3. Choose or type a player name.
4. Search for an item.
5. Pick the item.
6. Set the count.
7. Insert or build the command.
8. Review it.
9. Click `Send`.

### Vehicle Workflow

1. Open the PZ Commands window.
2. Go to the vehicle tab.
3. Refresh players if needed.
4. Search for a vehicle.
5. Pick the vehicle.
6. Insert or build the command.
7. Review it.
8. Click `Send`.

### About PZ Assets and Cache

The public GitHub repo may not include cached preview images or import scripts.

ECC can still build commands if it can find enough catalog data from local files, but previews and some lists may be incomplete without cache data.

Possible cache paths in packaged builds:

```text
Config\AssetCache\ProjectZomboid\Catalogs
Config\AssetCache\ProjectZomboid\TexturePackItems
Config\AssetCache\ProjectZomboid\VehiclePreviews
```

If those folders are missing, set the PZ profile `FolderPath` correctly and make sure local Project Zomboid files are available.

## Logs Guide

ECC has three kinds of logs.

### Discord Tab

Shows Discord listener activity and manual outgoing Discord messages.

### Program Log Tab

Use this first when something breaks.

It shows:

- Startup errors.
- Profile load errors.
- Path errors.
- API errors.
- REST errors.
- Telnet/RCON errors.
- Auto-save and restart events.
- Monitor events.

### Game Log Tabs

ECC creates game log tabs when it can find logs for running or detected servers.

Log tabs depend on the profile's `LogStrategy`.

### Log Files on Disk

ECC writes:

```text
Logs\console.log
Logs\bot_YYYYMMDD_HHMMSS.log
```

`console.log` is cleared each launch.

## Files and Folders

Public GitHub layout:

```text
Etherium-Command-Center/
|-- Config/
|   |-- CommandCatalog.json
|   |-- DefaultProfileTemplates.json
|   |-- HytaleVanillaItemCatalog.json
|   |-- LaunchArgCatalog.json
|   `-- MinecraftVanillaItemCatalog.json
|-- Modules/
|   |-- DiscordListener.psm1
|   |-- GUI.psm1
|   |-- Logging.psm1
|   |-- ProfileManager.psm1
|   `-- ServerManager.psm1
|-- Launch.ps1
|-- README.md
`-- Start.bat
```

Folders/files ECC creates:

| Path | When created |
| --- | --- |
| `Config\Settings.json` | When settings are saved. |
| `Profiles\` | On launch if missing. |
| `Profiles\*.json` | When profiles are added or saved. |
| `Logs\` | On launch if missing. |
| `Logs\console.log` | Every launch. |
| `Logs\bot_*.log` | Every launch. |

Optional packaged folders:

| Path | Purpose |
| --- | --- |
| `Config\AssetCache\...` | Optional cached item/preview data. |
| `Scripts\...` | Optional helper scripts if a release package includes them. |

## What Should and Should Not Be Uploaded

Safe to upload:

- `Start.bat`
- `Launch.ps1`
- `Modules\*.psm1`
- Public/default JSON files in `Config`
- `README.md`
- A license file

Do not upload:

- `Config\Settings.json`
- `Profiles\*.json` with real paths or passwords
- `Logs\*.log`
- Discord bot tokens
- Webhook URLs
- RCON passwords
- Telnet passwords
- Satisfactory API tokens
- Private IPs or public server IPs you do not want shared
- Third-party game assets unless you have permission to redistribute them

Asset cache note:

Game images and extracted game data may belong to the game developers. Keep public downloads focused on ECC code and let users generate/import local cache data from their own game installs when possible.

## Troubleshooting

### ECC Does Not Launch

Check:

- You extracted the ZIP first.
- You ran `Start.bat`, not only `Launch.ps1`.
- You allowed administrator elevation.
- Windows PowerShell 5.1 exists.
- `Modules\*.psm1` files are present.
- The folder path is not blocked by antivirus.

### PowerShell Blocks the App

ECC tries to unblock files and set execution policy for the current user.

If it still fails:

- Move ECC to a normal folder such as `C:\ECC`.
- Right-click `Start.bat`.
- Run as administrator.
- Check `Program Log` if the GUI opens.

### A Server Will Not Start

Check:

- `FolderPath` exists.
- `Executable` exists.
- `LaunchArgs` are correct.
- The game server can start outside ECC.
- Program Log shows a real process start attempt.

### Stop or Restart Is Too Abrupt

Check:

- `SaveMethod`
- `StdinSaveCommand`
- `SaveWaitSeconds`
- `StopMethod`
- Game-specific API/RCON/telnet credentials

### Logs Do Not Appear

Check:

- `LogStrategy`
- `ServerLogRoot`
- `ServerLogSubDir`
- `ServerLogFile`
- `ServerLogPath`
- `ServerLogNote`

For Project Zomboid, logs are usually under:

```text
%USERPROFILE%\Zomboid\Logs
```

### Discord Commands Do Not Work

Check:

- `BotToken`
- `WebhookUrl`
- `MonitorChannelId`
- `CommandPrefix`
- Discord Message Content Intent
- Bot permissions
- Correct command format, such as `!PZ start`
- The profile prefix is loaded in ECC

### 7 Days to Die Telnet Fails

Check:

- Telnet is enabled in `serverconfig.xml`.
- `TelnetPort` matches.
- `TelnetPassword` matches.
- The server is already running.

### Palworld REST Fails

Check:

- `RestEnabled`
- `RestHost`
- `RestPort`
- `RestPassword`
- REST launch args
- The server is already running.
- `Test REST` output.

### Satisfactory API Fails

Check:

- `SatisfactoryApiHost`
- `SatisfactoryApiPort`
- `SatisfactoryApiToken`
- Token generated with `server.GenerateAPIToken`
- Server is already running.
- `Test API` output.

### PZ Spawner Is Empty

Check:

- PZ `FolderPath`.
- Local PZ files exist.
- Optional `Config\AssetCache\ProjectZomboid` exists if using a package with cache.
- Program Log for catalog/cache errors.

### Config Button Is Disabled

ECC could not find a valid config root.

Check:

- `ConfigRoot`
- `ConfigRoots`
- Whether the folder exists
- Whether the profile points to the right game install or user config folder

## Security Notes

Before sharing ECC publicly, check for secrets in:

```text
Config\Settings.json
Profiles\*.json
Logs\*.log
```

Rotate any token that was ever uploaded or shown publicly.

Secrets to protect:

- Discord bot token.
- Discord webhook URL.
- RCON password.
- Telnet password.
- REST password.
- Satisfactory API token.
- Server IPs you do not want public.

## License

Distribution terms are pending.

Until a real license is added, users should treat this project as public source that can be viewed, but not automatically as open-source software with permission to redistribute or modify.

## Credits

Built by Darkjesusmn and AI contributors.
