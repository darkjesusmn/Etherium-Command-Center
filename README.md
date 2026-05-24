# Etherium Command Center (ECC)

Current GUI version: `v0.9`

Etherium Command Center, or ECC, is a Windows desktop app for running and managing local dedicated game servers from one control panel.

ECC is open-source software released under the MIT License.

ECC can start servers, stop servers, restart servers, show logs, edit server config files, send commands, talk to Discord, run auto-save, run scheduled restarts, and help with game-specific tools such as the Project Zomboid item and vehicle command builder.

<img width="1913" height="1076" alt="image" src="https://github.com/user-attachments/assets/a8ebee1f-93b1-4f8a-8938-438cd92daa40" />


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
- [Button Reference](#button-reference)
- [Settings Guide](#settings-guide)
- [Discord Setup](#discord-setup)
- [Discord Voice and Message Reference](#discord-voice-and-message-reference)
- [Profiles Guide](#profiles-guide)
- [Complete Profile Field Reference](#complete-profile-field-reference)
- [Per-Game Setup Notes](#per-game-setup-notes)
- [Commands Window Guide](#commands-window-guide)
- [Game Command Reference](#game-command-reference)
- [Config Editor Guide](#config-editor-guide)
- [Item Spawner Guides](#item-spawner-guides)
- [Logs Guide](#logs-guide)
- [Files and Folders](#files-and-folders)
- [Maintainer Packaging Notes](#maintainer-packaging-notes)
- [Known Issues and Current Limits](#known-issues-and-current-limits)
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
- Build Minecraft and Hytale item give commands.
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
| Hytale | `HY` | stdin | Java/JAR based. Includes an item give command builder. |
| Minecraft | `MC` | stdin and RCON | Includes an item give command builder. Uses `server.jar`, `logs\latest.log`, stdin save/stop, and RCON player list. |
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

### Dashboard Server States

Each server card shows a state badge. The badge tells you what ECC thinks is happening right now.

Common states:

| State | What it means |
| --- | --- |
| `Ready` | The profile is loaded and ECC is ready to control it, but the server is not running yet. This is a normal idle/offline profile state. |
| `Offline` | ECC does not see the server process running. |
| `Stopped` | ECC stopped the server or confirmed it is no longer running. |
| `Starting` | ECC started the process and is waiting for the server to become healthy or joinable. |
| `Online` | ECC sees the server as running and ready for normal control. |
| `Restarting` | ECC is in the middle of a restart flow. |
| `Stopping` | ECC is saving/stopping the server. |
| `Waiting` | ECC is waiting before the next step, such as waiting for first player data or waiting to retry a failed startup. |
| `Idle` | ECC sees the server as empty or waiting on idle-shutdown rules. |
| `Blocked` | ECC refused to start the server because a safety rule blocked it, such as RAM limits. |
| `Failed` | Startup or control failed. Check Program Log for the reason. |

The exact card text may include more detail under the badge, such as:

- `Waiting for first player`
- `Server empty. Idle shutdown in 10m`
- `Startup failed. Restarting in 10s`
- `Server process started successfully`

### Ready vs Waiting vs Idle

These three can sound similar, but they mean different things:

| State | Plain meaning |
| --- | --- |
| `Ready` | ECC is ready, but the server is not running. You can press `Start`. |
| `Waiting` | ECC is already doing something and is waiting for a condition or timer. Do not treat this as fully stopped. |
| `Idle` | The server is running or recently active, but ECC believes player activity is empty or waiting on idle-shutdown rules. |

Idle states are controlled by these profile fields:

| Field | What it affects |
| --- | --- |
| `ShutdownIfNoPlayersAfterStartupMinutes` | If nobody joins after startup for this many minutes, ECC can shut the server down. |
| `ShutdownIfEmptyAfterLastPlayerLeavesMinutes` | If the server becomes empty after players leave, ECC can shut it down after this many minutes. |

Use `0` for either field to disable that idle-shutdown rule.

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

## Button Reference

This section explains the buttons and button-like controls ECC can show.

Some buttons only appear for certain games or certain windows.

### Main Window Buttons

| Button | Where it appears | What it does |
| --- | --- | --- |
| `Start All` | Top bar | Starts every profile ECC thinks is ready. Test profiles one at a time before using this. |
| `Stop All` | Top bar | Stops all running servers using each profile's stop method. |
| `Reload UI` | Top bar | Rebuilds the interface without intentionally stopping running servers. |
| `Reload Bot` | Top bar | Restarts the Discord listener. |
| `Reload Cmds` | Top bar | Reloads command catalog data and profiles from disk. |
| `Full Restart` | Top bar | Stops managed servers, closes ECC, and relaunches the full program. |
| `Settings` | Top bar | Opens the Settings window. |
| `-` / `+` | Panel headers | Collapses or expands the Game Profiles, Profile Editor, or Logs panel. |
| `///` | Logs panel edge | Drag handle for resizing the bottom log area. |
| Window minimize | Custom window header | Minimizes ECC. |
| Window maximize/restore | Custom window header | Maximizes or restores ECC. |
| Window close | Custom window header | Closes ECC after shutdown handling. |

### Profile List Buttons

| Button | What it does |
| --- | --- |
| `+ Add Game` | Starts the profile creation flow. |
| `Remove` | Removes the selected profile from ECC. Use carefully. |

### Server Card Buttons

| Button | What it does |
| --- | --- |
| `Start` | Starts that server. |
| `Stop` | Saves when possible, waits, then stops that server. |
| `Restart` | Stops then starts that server. |
| `Commands` | Opens the Commands window for that server. |
| `Config` | Opens the Config Editor if ECC can find a valid config root. |
| `Manager` | Opens the Hytale manager tools. Only appears for Hytale profiles. |
| `Auto-Restart` | Enables or disables crash auto-restart for that profile. |

### Profile Editor Buttons

| Button | What it does |
| --- | --- |
| `Save Changes` | Saves the current profile editor values to the profile JSON file. |
| `Restart Server` | Restarts the selected server from the profile editor. |
| `Stop Server` | Stops the selected server from the profile editor. |
| `...` | Browse field button | Opens a picker for path fields. |
| Toggle checkboxes | Profile fields | Turns boolean profile settings on or off. |

### Settings Window Buttons

| Button/control | What it does |
| --- | --- |
| `Save Settings` | Writes `Config\Settings.json` and restarts the Discord listener. |
| `Enabled` | Debug checkbox | Turns debug logging on or off. |
| `Enable long-run perf tracing` | Performance checkbox | Turns performance tracing on or off. |
| `Enable auto-save for all games` | Auto-save checkbox | Turns global auto-save on or off. |
| `Enable scheduled restarts for all games` | Restart checkbox | Turns global scheduled restarts on or off. |

### Log Buttons

| Button | Where it appears | What it does |
| --- | --- | --- |
| `Send` | Discord log tab | Sends the typed text to Discord. |
| `Clear` | Discord, Program Log, and game log tabs | Clears that visible log box in the UI. |
| `Copy` | Program Log and game log tabs | Copies visible log text to the clipboard. |

### Config Editor Buttons

| Button/control | What it does |
| --- | --- |
| `Search` | Applies the config editor filter/search text. |
| `Clear` | Clears the config editor filter/search text. |
| `Generated` | Opens the structured editor tab when supported. |
| `Raw` | Opens the raw text editor tab. |
| `Save File` | Saves the currently opened config file. |
| `Refresh` | Reloads the config file or config tree from disk. |
| `Enabled` | Generated config checkbox | Turns a boolean config value on or off. |

### Commands Window Buttons and Controls

| Button/control | What it does |
| --- | --- |
| Catalog command buttons | Fill the command box with a known command. They do not always send immediately. |
| `Verbose Debug` | Shows extra routing/request/response details. |
| `Use API` | Forces API routing when supported. |
| `Use REST` | Forces REST routing when supported. |
| `Test Telnet` | Tests telnet connection/auth. |
| `Test RCON` | Tests RCON connection/auth. |
| `Test API` | Tests the Satisfactory API. |
| `Test STDIN` | Tests whether ECC can send text to the running server. |
| `Test PID` | Tests whether ECC sees the server process. |
| `Test REST` | Tests the REST API. |
| `Send` | Sends the command currently in the command box. |

### Item and Vehicle Spawner Buttons

| Button/control | Game(s) | What it does |
| --- | --- | --- |
| `Refresh` | Minecraft, Hytale, PZ | Refreshes the online player list when supported. |
| `Items` | Minecraft, Hytale, PZ | Shows the item search/list tab. |
| `Vehicles` | PZ only | Shows the vehicle search/list tab. |
| `Insert Into Command Box` | Minecraft, Hytale, PZ | Inserts the built item or vehicle command into the command box. |
| `Build /give` | Minecraft, Hytale | Builds an item give command. |
| `Build /additem` | PZ | Builds an item add command. |
| `Build /addvehicle` | PZ | Builds a vehicle spawn command. |

### Add Game and Folder Picker Buttons

| Button | What it does |
| --- | --- |
| Game type buttons | Choose the game template for the new profile. |
| `Use Selected` | Uses a detected folder or detected item in an add-game helper window. |
| `Browse Manually` | Opens manual folder selection instead of using a detected path. |
| `Create` | Creates a profile after naming or confirming the game. |
| `Cancel` | Cancels the current picker/dialog. |
| `Up` | Moves up one folder in ECC's folder picker. |
| `Refresh` | Refreshes the folder picker contents. |
| `Select Folder` | Uses the currently selected folder. |

### Hytale Manager Buttons

These appear in the Hytale `Manager` window.

| Button/control | What it does |
| --- | --- |
| `Open Folder` | Opens or targets the Hytale server folder used by the manager. |
| `Update Server` | Runs the Hytale server update flow. |
| `Auto-restart after update` | Restarts the Hytale server after update when enabled. |
| `Check Server Update` | Checks whether the Hytale server files need an update. |
| `Downloader Version` | Checks the downloader version. |
| `Check Downloader Update` | Checks whether the downloader has an update. |
| `Verify Required Files` | Checks required Hytale files such as server JAR/assets/cache paths. |
| `Update Downloader` | Updates the downloader. |
| `Refresh` | Refreshes installed mod/tool state. |
| `Add Mod...` | Adds a mod. |
| `Toggle Mod` | Enables or disables the selected mod. |
| `Mod Folder` | Opens the mod folder. |
| `Delete Mod` | Deletes the selected mod. |
| `Check Conflicts` | Checks installed mods for possible conflicts. |
| `Open Selected` | Opens the selected mod item. |
| `Browse Configs` | Opens/browses mod or server config folders. |
| `Link Mod CF` | Links a mod to CurseForge metadata. |
| `Check Updates` | Checks selected/installed mods for updates. |
| `Update Mod` | Updates the selected mod. |
| `Open Mod Page` | Opens the selected mod's page when known. |
| `Get More Mods` | Opens the flow/page for finding more mods. |
| `Save Notes` | Saves notes for the selected mod. |

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

### Debug Logging and Performance Trace Mode

ECC has two different troubleshooting modes.

| Mode | Setting | What it is for |
| --- | --- | --- |
| Debug Logging | `EnableDebugLogging` | Full troubleshooting logs when something is broken or confusing. |
| Performance Trace Mode | `EnablePerformanceDebugMode` | Lighter long-session tracing when ECC gets slow, laggy, or delayed over time. |

Debug Logging:

- Turns on extra `DEBUG` log entries.
- Helps troubleshoot startup, profile loading, Discord delivery, command routing, GUI behavior, and server control.
- Can make `Logs\console.log` and bot logs much larger.
- Can include local paths, profile names, command names, connection details, and error details.
- Should normally be turned off after you finish troubleshooting.
- Changing it from the Settings window asks for confirmation because ECC restarts and stops managed servers.

Performance Trace Mode:

- Adds timing logs without turning on every normal debug message.
- Is meant for long-running lag checks.
- Records slow UI and log work using entries such as `UIPERF` and `LOGPERF`.
- Helps find slow dashboard redraws, slow log appends, slow log trimming, and other UI delay points.
- Applies right away after saving settings.
- Should normally be turned off after the lag test is done.

Important difference:

Debug Logging is the heavy mode. Performance Trace Mode is the lighter lag-investigation mode. If Debug Logging is on, ECC also allows performance tracing because full debug mode is already active.

Where to look:

- `Program Log` tab for live troubleshooting.
- `Discord` tab for Discord listener activity.
- `Logs\console.log` for the main saved log from the current launch.
- `Logs\bot_*.log` for Discord listener log files.

Do not post debug logs publicly until you check them for private paths, tokens, webhook URLs, passwords, IP addresses, and server details.

### Verbose Debug in the Commands Window

The Commands window also has a separate `Verbose Debug` checkbox.

That checkbox is not the same thing as app-wide Debug Logging. It only affects command testing and command sending inside that Commands window.

Use `Verbose Debug` when a command does not work and you need to see:

- Which route ECC chose.
- Which host and port ECC tried.
- Whether ECC used stdin, RCON, telnet, REST, or Satisfactory API.
- What response or error came back.

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

## Discord Voice and Message Reference

ECC does not talk like a neutral corporate bot. Its Discord style is intentionally dramatic: it sounds like a nervous bunker operator doing dangerous server work while corporate overlords watch the cameras and wait for him to make one career-ending mistake.

In plain terms, the voice is:

- stressed but useful
- bunker/control-room themed
- afraid of management
- trying very hard not to get blamed
- still focused on telling players what happened

The code uses randomized templates. That means ECC does not always send the exact same sentence for the same event. It chooses from event templates and flavor phrase pools for the game.

Because of that, the exact possible full sentence count is very large. The tables below list every Discord message family ECC can send, the fixed direct messages, and the template shapes used for randomized messages.

### Discord Delivery Paths

| Path | What it sends |
| --- | --- |
| Bot token + channel ID | Reads commands and can send direct bot-channel messages when configured. |
| Webhook URL | Sends most ECC notifications, acknowledgements, and command results. |
| Webhook queue | Holds outbound messages until the listener flushes them. |

### Fixed Direct Discord Messages

These messages are direct strings or direct string shapes.

| Message shape | When it appears |
| --- | --- |
| `[BLOCKED] Unsafe command '<command>' is not allowed.` | A Discord user sends a blocked command such as kick, ban, teleport, op, deop, whitelist, give, or take. |
| `[UNKNOWN] No game found with prefix '<prefix>'. Available: <prefix list>` | Discord command uses a prefix ECC does not know. |
| `[COOLDOWN] !<prefix> <command> is on cooldown. Please wait <seconds>s before using it again.` | Command is valid but still on cooldown. |
| `[ERROR] Failed to process command: <error>` | Command routing threw an exception. |
| `[ERROR] No profiles available` | Command routing ran before profiles were loaded. |
| `[ERROR] Profile '<prefix>' not found. Available: <prefix list>` | Command routing could not find the requested profile. |
| `[ERROR] Profile '<prefix>' has no Commands section.` | Profile JSON is missing `Commands`. |
| `[ERROR] Command '<command>' not found. Available: <command list>` | Profile does not define that command. |
| `[SAVE] <GameName> - no manual save command is available for this server. World saving remains automatic.` | A save command is requested for a profile with no manual save route. |
| `[PLAYERS] <GameName> - live player listing is not supported yet for this server.` | Valheim player list is requested. |
| `[PLAYERS] <GameName> - no players online.` | Minecraft/Hytale player query succeeds and finds nobody. |
| `[PLAYERS] <GameName> - online: <names>` | Minecraft/Hytale player query returns player names. |
| `[PLAYERS] <GameName> - players online: <count>` | Minecraft/Hytale player query returns a count but not names. |
| `[RCON] <GameName>: <response>` | RCON command returns text. |
| `[RCON] <GameName>: command sent (no response)` | RCON command succeeds but returns no text. |
| `[ERROR] RCON password not configured in profile for <GameName>. Add RconPassword to the profile JSON.` | RCON command needs a password but none is set. |
| `[ERROR] RCON failed for <GameName>: <error>` | RCON command fails. |
| `[BOT] <message>` | You type a manual message in ECC's Discord tab and click `Send`. |

### Randomized Discord Event Families

ECC can send these event families through `New-DiscordGameMessage`.

| Event | Main tag | What it means |
| --- | --- | --- |
| `received_start` | `[RECEIVED]` | Discord user requested `start`. |
| `received_stop` | `[RECEIVED]` | Discord user requested `stop`. |
| `received_restart` | `[RECEIVED]` | Discord user requested `restart`. |
| `received_save` | `[RECEIVED]` | Discord user requested `save`. |
| `received_status` | `[RECEIVED]` | Discord user requested `status`. |
| `received_command` | `[RECEIVED]` | Discord user requested another allowed command. |
| `starting` | `[STARTING]` | ECC accepted a server start. |
| `online` | `[ONLINE]` | Server process is online. |
| `joinable` | `[JOINABLE]` | Server looks ready for players. |
| `save_sent` | `[OK]` | Save command was sent. |
| `saving` | `[SAVING]` | ECC is saving before shutdown. |
| `waiting` | `[WAITING]` | ECC is waiting before shutdown. |
| `waiting_no_save` | `[WAITING]` | ECC has no save route and is waiting before shutdown. |
| `stopped` | `[STOPPED]` | Server stopped safely. |
| `restarting` | `[RESTARTING]` | Restart sequence started. |
| `restarted` | `[RESTARTED]` | Restart finished successfully. |
| `restarted_auto` | `[RESTARTED]` | Crash recovery restarted the server. |
| `autosave_started` | `[AUTOSAVE STARTED]` | Auto-save started. |
| `autosave_done` | `[AUTOSAVE COMPLETED]` | Auto-save completed. |
| `blocked` | `[BLOCKED]` | Safety rule blocked a start. |
| `startup_failed` | `[STARTUP FAILED]` | Server process started but never became healthy/joinable. |
| `idle_shutdown` | `[IDLE SHUTDOWN]` | ECC shut down an idle/empty server. |
| `restart_warning` | `[RESTART WARNING]` | Scheduled restart warning. |
| `scheduled_restart_started` | `[SCHEDULED RESTART]` | Scheduled restart began. |
| `scheduled_restart_done` | `[RESTARTED]` | Scheduled restart finished. |
| `scheduled_restart_retry` | `[WARNING]` | Scheduled restart recovery will retry. |
| `scheduled_restart_failed` | `[WARNING]` | Scheduled restart recovery failed. |
| `crashed_retry` | `[CRASHED]` | Server crashed and auto-restart is queued. |
| `crash_suspended` | `[WARNING]` | Crash limit was reached and auto-restart is suspended. |
| `players_none` | `[PLAYERS]` | Player query found nobody online. |
| `players_list` | `[PLAYERS]` | Player query found names. |
| `already_running` | `[ONLINE]` | Start was requested but server was already running. |
| `status_online` | `[ONLINE]` | Status command sees server online. |
| `status_offline` | `[OFFLINE]` | Status command sees server offline. |
| `command_ok` | `[OK]` | A generic command was accepted/sent. |
| `command_error` | `[ERROR]` | A generic command failed. |
| `error_start` | `[ERROR]` | Start failed. |
| `error_restart` | `[ERROR]` | Restart failed. |
| default event | `[INFO]` | Fallback message for an unknown event name. |

### Randomized Discord System Families

ECC can also talk about ECC itself through `New-DiscordSystemMessage`.

| Event | Main tag | What it means |
| --- | --- | --- |
| `online` | `[ONLINE]` | ECC or the Discord listener came online. |
| `offline` | `[OFFLINE]` | ECC or the Discord listener is shutting down or going offline. |
| `reload_ui` | `[SYSTEM]` | `Reload UI` was used. |
| `reload_bot` | `[SYSTEM]` | `Reload Bot` was used. |
| `reload_commands` | `[SYSTEM]` | `Reload Cmds` was used. |
| `full_restart` | `[SYSTEM]` | `Full Restart` was used. |
| default system event | `[SYSTEM]` | Fallback message for an unknown ECC system event. |

### Randomized Message Template Shapes

These are the shapes ECC fills with game name, requester, command, reason, PID, uptime, wait seconds, and game flavor text.

| Event | Template shape |
| --- | --- |
| `received_start` | `[RECEIVED] {Requester} start command received for {GameName}. {StartLead} {OperatorOrderLead}` |
| `received_stop` | `[RECEIVED] {Requester} stop command received for {GameName}. {WaitLead} {OperatorOrderLead}` |
| `received_restart` | `[RECEIVED] {Requester} restart command received for {GameName}. {RestartLead} {OperatorOrderLead}` |
| `received_save` | `[RECEIVED] {Requester} save command received for {GameName}. {SaveLead} {OperatorOrderLead}` |
| `received_status` | `[RECEIVED] {Requester} checking status for {GameName}. {StatusLead} {OperatorStatusLead}` |
| `received_command` | `[RECEIVED] {Requester} running {Command} for {GameName}. {OperatorCommandLead}` |
| `starting` | `[STARTING] {GameName} is starting. {StartLead}` |
| `online` | `[ONLINE] {GameName} is now online (PID {Pid}). {ReadyLead}` |
| `joinable` | `[JOINABLE] {GameName} is ready for players. {JoinableTip}` |
| `save_sent` | `[OK] {GameName} got the save call. {SaveLead}` |
| `saving` | `[SAVING] {GameName} save command sent. Waiting {WaitSeconds}s before shutdown. {SaveLead}` |
| `waiting` | `[WAITING] {GameName} is waiting {WaitSeconds}s before shutdown. {WaitLead}` |
| `waiting_no_save` | `[WAITING] {GameName} has no working save command. Waiting {WaitSeconds}s before shutdown. {WaitLead}` |
| `stopped` | `[STOPPED] {GameName} stopped safely. {StopLead}` |
| `restarting` | `[RESTARTING] {GameName} restart sequence started. {RestartLead}` |
| `restarted` | `[RESTARTED] {GameName} restarted successfully and is back online. {ReadyLead}` |
| `restarted_auto` | `[RESTARTED] {GameName} recovered automatically and is back online (PID {Pid}). {ReadyLead}` |
| `autosave_started` | `[AUTOSAVE STARTED] {GameName} auto-save started. {SaveLead}` |
| `autosave_done` | `[AUTOSAVE COMPLETED] {GameName} auto-save completed successfully. {SaveLead}` |
| `blocked` | `[BLOCKED] {GameName} start was blocked. {Reason} {OperatorBlockedLead} {Caution}` |
| `startup_failed` | `[STARTUP FAILED] {GameName} never fully woke up. {Reason} {Action} {OperatorStartupFailLead}` |
| `idle_shutdown` | `[IDLE SHUTDOWN] {GameName} is shutting down from inactivity. {Reason}` |
| `restart_warning` | `[RESTART WARNING] {GameName} restart in {Minutes} minute(s). {PlayerSummary} {Caution}` |
| `scheduled_restart_started` | `[SCHEDULED RESTART] {GameName} scheduled restart has begun. {PlayerSummary} {RestartLead}` |
| `scheduled_restart_done` | `[RESTARTED] {GameName} completed its scheduled restart and is back online. {ReadyLead}` |
| `scheduled_restart_retry` | `[WARNING] {GameName} needs another scheduled restart recovery try in {DelaySeconds}s (attempt {Attempt}/{MaxAttempts}). {Caution}` |
| `scheduled_restart_failed` | `[WARNING] {GameName} scheduled restart recovery failed after {Attempt} attempt(s). {Caution}` |
| `crashed_retry` | `[CRASHED] {GameName} crashed. Restart in {DelaySeconds}s (attempt {Attempt}/{MaxAttempts}). {CrashLead} {OperatorCrashLead}` |
| `crash_suspended` | `[WARNING] {GameName} crashed {Count} times this hour. Auto-restart is suspended. {OperatorStandDownLead} {Caution}` |
| `players_none` | `[PLAYERS] {GameName} online players: none. {JoinableTip}` |
| `players_list` | `[PLAYERS] {GameName} online players: {Names}.` |
| `already_running` | `[ONLINE] {GameName} is already running. Uptime: {Uptime}. {ReadyLead}` |
| `status_online` | `[ONLINE] {GameName} is running. PID {Pid}, uptime {Uptime}. {StatusDetail}` |
| `status_offline` | `[OFFLINE] {GameName} is offline. {StatusDetail}` |
| `command_ok` | `[OK] {GameName} received {Command}. {StatusLead}` |
| `command_error` | `[ERROR] {GameName} could not run {Command}. {Reason}` |
| `error_start` | `[ERROR] {GameName} failed to start. {Reason} {OperatorStartupFailLead}` |
| `error_restart` | `[ERROR] {GameName} restart failed. {Reason} {OperatorRestartFailLead}` |
| default | `[INFO] {GameName} event {EventName}: {Reason}` |

The code includes multiple alternate lines for most of these shapes, so two messages with the same tag can still read differently.

### System Message Template Shapes

These are the ECC-wide message shapes. Like the game messages, each event has multiple alternate lines with the same meaning.

| Event | Template shape |
| --- | --- |
| `online` | `[ONLINE] {AppName} is online again. {OnlineLead}` |
| `offline` | `[OFFLINE] {AppName} is offline for now. {OfflineLead}` |
| `reload_ui` | `[SYSTEM] Fine. Reloading the dashboard. {ReloadUiLead}` |
| `reload_bot` | `[SYSTEM] Fine. I am kicking the Discord bot again. {ReloadBotLead}` |
| `reload_commands` | `[SYSTEM] Fine. Rebuilding the command table. {ReloadCommandsLead}` |
| `full_restart` | `[SYSTEM] Full restart requested for {AppName}. {FullRestartLead}` |
| default | `[SYSTEM] {AppName} event {EventName}: {Reason}` |

### Bunker Operator Flavor Pools

ECC also mixes in phrase pools. These are the major pools:

| Pool | What it adds |
| --- | --- |
| `StartLead` | Startup flavor. |
| `ReadyLead` | Successful online/ready flavor. |
| `JoinableTip` | Joinable/player-ready flavor. |
| `SaveLead` | Save/autosave flavor. |
| `WaitLead` | Shutdown wait/cooldown flavor. |
| `StopLead` | Stop/offline flavor. |
| `RestartLead` | Restart flavor. |
| `CrashLead` | Crash/recovery flavor. |
| `OperatorCrashLead` | Bunker operator panic after a crash. |
| `OperatorStandDownLead` | Bunker operator refusing more auto-restarts after too many crashes. |
| `OperatorBlockedLead` | Bunker operator refusing a dangerous start. |
| `OperatorStartupFailLead` | Bunker operator reacting to startup failure. |
| `OperatorRestartFailLead` | Bunker operator reacting to restart failure. |
| `OperatorOrderLead` | Bunker operator acknowledging a Discord order. |
| `OperatorStatusLead` | Bunker operator reading gauges/status. |
| `OperatorCommandLead` | Bunker operator pushing a manual command. |
| `Caution` | Follow-up warning to watch logs/dashboard. |
| `StatusLead` | Status report flavor. |
| `OnlineLead` | ECC app/listener came online flavor. |
| `OfflineLead` | ECC app/listener went offline flavor. |
| `ReloadUiLead` | Dashboard reload flavor. |
| `ReloadBotLead` | Discord bot reload flavor. |
| `ReloadCommandsLead` | Command/profile reload flavor. |
| `FullRestartLead` | Full ECC restart flavor. |

Examples of the corporate-overlord bunker voice:

- `I am already under the console kicking relays before the overlords smell smoke.`
- `I am trying to make this look controlled before somebody upstairs starts asking for names.`
- `I am working the recovery board now, mostly because I enjoy continued breathing privileges.`
- `I am holding that line shut before corporate decides I ignored a red light.`
- `I am sweeping up the failed startup before corporate asks who touched what.`
- `I am pushing the command through before corporate mistakes hesitation for rebellion.`

Game-specific flavor changes the nouns. For example:

| Game | Flavor examples |
| --- | --- |
| 7 Days to Die | wasteland, horde, bunker, traders, forge, Navezgane |
| Hytale | shard, portal, realm, wilds, gate |
| Palworld | Palbox, island, ranch, Pals |
| Project Zomboid | Knox County, safehouse, barricades, generator |
| Minecraft | spawn, chunks, overworld, creeper-adjacent server work |
| Satisfactory | factory, production line, conveyor, API |
| Valheim | mead hall, longhouse, world tree, Vikings |

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
- Item Spawner panel for Minecraft, Hytale, and Project Zomboid.
- Vehicle Spawner tab for Project Zomboid only.

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

## Game Command Reference

This section lists the commands that ECC includes in `Config\CommandCatalog.json`.

The `Label` is what ECC shows in the command catalog. The `Command` is the text ECC can place into the command box. Any value inside `<angle brackets>` or `[square brackets]` is something you are expected to replace before sending the command.

Some games require admin permission, operator permission, RCON, REST, telnet, or a live server console before the command will work. ECC helps send the command, but the game server still decides whether the command is allowed.

### 7 Days to Die Commands

| Label | Command | Description |
| --- | --- | --- |
| `help` | `help <command>` | Show what a command does and how to use it. |
| `admin add` | `admin add <name\|entity id\|steam id> <permission level>` | Add a player to the admin list and give them a permission level. |
| `admin remove` | `admin remove <name\|entity id\|steam id>` | Remove a player from the admin list. |
| `ai pathgrid` | `ai pathgrid` | Turn the AI path grid debug view on or off. |
| `ai pathlines` | `ai pathlines` | Turn AI path line debug view on or off. |
| `aiddebug` | `aiddebug` | Turn AI director debug output on or off. |
| `ban add` | `ban add <name\|entity id\|steam id> <duration> <duration unit> [reason]` | Ban a player for a set time, with an optional reason. |
| `ban list` | `ban list` | Show the current ban list. |
| `ban remove` | `ban remove <name\|entity id\|steam id>` | Remove a player from the ban list. |
| `buff` | `buff <buff name>` | Apply a buff to yourself. |
| `buffplayer` | `buffplayer <name\|entity id\|steam id> <buff name>` | Apply a buff to a player. |
| `chunkcache` | `chunkcache` | Show chunk cache information. |
| `clear` | `clear` | Clear the console output. |
| `cp add` | `cp add <command> <level>` | Set the required permission level for a command. |
| `cp remove` | `cp remove <command>` | Remove a custom permission rule from a command. |
| `cp list` | `cp list` | Show the current command permission rules. |
| `creativemenu` | `creativemenu` | Turn the creative menu on or off. |
| `deathscreen` | `deathscreen <on\|off>` | Turn the death screen on or off. |
| `debuff` | `debuff <buff name>` | Remove a buff from yourself. |
| `debuffplayer` | `debuffplayer <name\|entity id\|steam id> <buff name>` | Remove a buff from a player. |
| `debugmenu` | `debugmenu [on\|off]` | Turn the debug menu on or off. |
| `enablescope` | `enablescope <on\|off>` | Turn scoped view on or off. |
| `exhausted` | `exhausted` | Make the player exhausted. |
| `exportcurrentconfigs` | `exportcurrentconfigs` | Export the current game config files. |
| `exportitemicons` | `exportitemicons` | Export the game item icons. |
| `getgamepref` | `getgamepref` | Show the current value of a game preference. |
| `getgamestat` | `getgamestat` | Show the current value of a game stat. |
| `gettime` | `gettime` | Show the current in-game time. |
| `gfx af` | `gfx af <0\|1>` | Set anisotropic filtering to 0 or 1. |
| `gfx dti` | `gfx dti` | Turn texture info debug view on or off. |
| `gfx dtpix` | `gfx dtpix` | Turn texture pixel debug view on or off. |
| `givequest` | `givequest` | Give a quest to the player. |
| `giveself` | `giveself <item name> [quality level]` | Give an item to yourself. |
| `giveselfskillxp` | `giveselfskillxp <skill name> <amount>` | Give skill XP to yourself (legacy). |
| `giveselfxp` | `giveselfxp <amount>` | Give XP to yourself. |
| `kick` | `kick <name\|entity id\|steam id> [reason]` | Kick a player from the server, with an optional reason. |
| `kickall` | `kickall [reason]` | Kick all players with an optional reason. |
| `killall` | `killall` | Kill all entities. |
| `lights` | `lights` | Turn light debug output on or off. |
| `listents` | `listents` | Show the active entities. |
| `listlandclaim` | `listlandclaim` | Show the land claim blocks. |
| `listplayerids` | `listplayerids` | Show the player IDs. |
| `listplayers` | `listplayers` | Show the players who are connected right now. |
| `listthreads` | `listthreads` | Show the running threads. |
| `loggamestate` | `loggamestate <message> [true\|false]` | Write the current game state to the log, with an optional message. |
| `loglevel` | `loglevel <loglevel name> <true\|false>` | Turn one log level on or off. |
| `mem` | `mem` | Show memory information. |
| `memcl` | `memcl` | Show memory information and run garbage collection. |
| `pplist` | `pplist` | List persistent player data. |
| `removequest` | `removequest` | Remove a quest. |
| `repairchunkdensity` | `repairchunkdensity <x> <z> [fix]` | Check chunk density at the given coordinates, and optionally fix it. |
| `saveworld` | `saveworld` | Save the world right now. |
| `say` | `say <message>` | Send a chat message to everyone on the server. |
| `setgamepref` | `setgamepref <preference name> <value>` | Set a game preference to a new value. |
| `setgamestat` | `setgamestat <stat name> <value>` | Set a game stat to a new value. |
| `settempunit` | `settempunit <c\|f>` | Set the temperature unit to C or F. |
| `settime` | `settime <day> <hour> <minute>` | Set the in-game day, hour, and minute. |
| `showalbedo` | `showalbedo` | Turn albedo render debug view on or off. |
| `showchunkdata` | `showchunkdata` | Show chunk data debug information. |
| `showclouds` | `showclouds` | Turn cloud debug view on or off. |
| `shownexthordetime` | `shownexthordetime` | Show how long until the next horde. |
| `shownormals` | `shownormals` | Turn normals render debug view on or off. |
| `showspecular` | `showspecular` | Turn specular render debug view on or off. |
| `shutdown` | `shutdown` | Shut down the server. |
| `sounddebug` | `sounddebug` | Turn sound debug output on or off. |
| `spawnairdrop` | `spawnairdrop` | Spawn an airdrop. |
| `spawnentity` | `spawnentity <playerid> <entityid>` | Spawn an entity by ID at a player. |
| `spawnscouts` | `spawnscouts` | Spawn scout zombies. |
| `spawnscreen` | `spawnscreen` | Open the spawn screen. |
| `spawnsupplycrate` | `spawnsupplycrate` | Spawn a supply crate. |
| `spawnwh` | `spawnwh` | Spawn a wandering horde. |
| `spectrum` | `spectrum <choice>` | Set or switch the lighting spectrum. |
| `starve` | `starve` | Make the player starve. |
| `staticmap` | `staticmap` | Generate a static map image. |
| `switchview` | `switchview` | Switch between first-person and third-person view. |
| `systeminfo` | `systeminfo` | Show system information. |
| `teleport` | `teleport <target>` | Teleport to a player or target location. |
| `thirsty` | `thirsty` | Make the player thirsty. |
| `traderarea` | `traderarea` | Turn trader area debug or visibility on or off. |
| `updatelighton` | `updatelighton <name\|entity id\|steam id>` | Force the lights to update for a player. |
| `version` | `version` | Show the current server or game version. |
| `water` | `water` | Turn water debug view on or off. |
| `weather` | `weather` | Turn weather debug view on or off. |
| `weathersurvival` | `weathersurvival <on\|off>` | Turn weather survival effects on or off. |
| `whitelist add` | `whitelist add <name\|entity id\|steam id>` | Add a player to the whitelist so they are allowed to join. |
| `whitelist remove` | `whitelist remove <name\|entity id\|steam id>` | Remove a player from the whitelist. |
| `whitelist list` | `whitelist list` | Show the players on the whitelist. |

### Hytale Commands

| Label | Command | Description |
| --- | --- | --- |
| `/help` | `/help` | Show the available commands. |
| `/heal` | `/heal` | Restore full health and stamina. |
| `/fillsignature` | `/fillsignature` | Fill the signature meter. |
| `/leave` | `/leave` | Leave the current world or server. |
| `/unstuck` | `/unstuck` | Teleport to a safe spot if you are stuck. |
| `/inventory backpack` | `/inventory backpack --size=[#]` | Show or resize your backpack inventory. |
| `/inventory clear` | `/inventory clear` | Clear your inventory. |
| `/inventory see` | `/inventory see [player]` | Show another player's inventory. |
| `/memories unlockall` | `/memories unlockall` | Unlock all memories. |
| `/memories clear` | `/memories clear` | Clear the unlocked memories. |
| `/kill` | `/kill [player]` | Kill yourself, or the specified player. |
| `/neardeath` | `/neardeath` | Drop health to a near-death level. |
| `/damage` | `/damage [player] --amount=[#]` | Deal a set amount of damage to a player. |
| `/emote` | `/emote [emote]` | Play an emote. |
| `/mount` | `/mount [mount]` | Spawn a mount, or mount one. |
| `/ping` | `/ping` | Show ping or latency. |
| `/spawn` | `/spawn` | Teleport to spawn. |
| `/warp set` | `/warp set [name]` | Create a named warp point. |
| `/warp go` | `/warp go [name]` | Teleport to a named warp point. |
| `/warp remove` | `/warp remove [name]` | Remove a named warp point. |
| `/warp list` | `/warp list` | Show the available warp points. |
| `/warp reload` | `/warp reload` | Reload the saved warp data. |
| `/tp home` | `/tp home` | Teleport to your home location. |
| `/tp top` | `/tp top` | Teleport to the top of the world/structure. |
| `/tp` | `/tp [player]` | Teleport to a player. |
| `/gamemode creative` | `/gamemode creative` | Switch to creative mode. |
| `/gamemode exploration` | `/gamemode exploration` | Switch to exploration mode. |
| `/worldmap reload` | `/worldmap reload` | Reload the world map. |
| `/worldmap discover` | `/worldmap discover` | Reveal the world map. |
| `/worldmap undiscover` | `/worldmap undiscover` | Hide the revealed world map again. |
| `/worldmap clearmarkers` | `/worldmap clearmarkers` | Clear the world map markers. |
| `/time dawn` | `/time dawn` | Set the time to dawn. |
| `/time midday` | `/time midday` | Set the time to midday. |
| `/time dusk` | `/time dusk` | Set the time to dusk. |
| `/time midnight` | `/time midnight` | Set the time to midnight. |
| `/noon` | `/noon` | Set the time to noon and pause time. |
| `/weather reset` | `/weather reset` | Reset the weather back to its normal state. |
| `/give` | `/give [item id] --quantity=[#]` | Give yourself an item. |
| `/give player` | `/give [player] [item id] --quantity=[#]` | Give another player an item. |
| `/op self` | `/op self` | Grant yourself operator permissions. |
| `/op add` | `/op add [player]` | Grant operator permissions to a player. |
| `/op remove` | `/op remove [player]` | Remove operator permissions from a player. |
| `/who` | `/who` | Show the players who are connected right now. |
| `/whoami` | `/whoami [player]` | Show player details for yourself, or for another player if you name them. |
| `/whereami` | `/whereami [player]` | Show your location, or another player's location if you name them. |
| `/kick` | `/kick [player]` | Kick a player from the server. |
| `/ban` | `/ban [player] --reason=[reason]` | Ban a player, with an optional reason. |
| `/unban` | `/unban [player]` | Remove a player from the ban list. |
| `/whitelist list` | `/whitelist list` | Show the players on the whitelist. |
| `/whitelist enable` | `/whitelist enable` | Turn on the whitelist so only approved players can join. |
| `/whitelist disable` | `/whitelist disable` | Turn off the whitelist. |
| `/whitelist add` | `/whitelist add [player]` | Add a player to the whitelist so they are allowed to join. |
| `/whitelist remove` | `/whitelist remove [player]` | Remove a player from the whitelist. |
| `/backup` | `/backup` | Create a backup of the server data. |
| `/stop` | `/stop` | Stop the server. |
| `/server dump` | `/server dump` | Show a detailed server information dump. |
| `/server gc` | `/server gc` | Run server garbage collection to clean up memory. |
| `/server stats` | `/server stats [value]` | Show server stats, or one specific stat if you include a value. |
| `/spawning` | `/spawning` | Turn spawning on or off, or change how it behaves. |
| `/perm` | `/perm` | Manage permissions. |
| `/plugin` | `/plugin` | Manage plugins. |
| `/sudo` | `/sudo [player] <command>` | Run a command as another player. |
| `/block` | `/block` | Run block debug or management commands. |
| `/chunk` | `/chunk` | Run chunk debug or management commands. |
| `/fluid` | `/fluid` | Run fluid debug or management commands. |
| `/lighting` | `/lighting` | Run lighting debug or management commands. |
| `/path` | `/path` | Run pathfinding debug commands. |
| `/world` | `/world` | Run world debug or management commands. |
| `/copy` | `/copy` | Copy the selected blocks. |
| `/cut` | `/cut` | Cut the selected blocks. |
| `/editprefab` | `/editprefab` | Edit a prefab. |
| `/fillblocks` | `/fillblocks` | Fill a selected area with blocks. |
| `/paste` | `/paste` | Paste the copied blocks. |
| `/pos1` | `/pos1` | Set position 1 for a selection. |
| `/pos2` | `/pos2` | Set position 2 for a selection. |
| `/prefab` | `/prefab` | Run prefab management commands. |
| `/replace` | `/replace` | Replace blocks inside a selection. |
| `/undo` | `/undo` | Undo the last edit action. |

### Palworld Commands

| Label | Command | Description |
| --- | --- | --- |
| `AdminPassword` | `/AdminPassword <password>` | Set the admin password for this session so admin commands will work. |
| `Shutdown` | `/Shutdown [seconds] [message]` | Shut down the server, with an optional delay and message to players. |
| `DoExit` | `/DoExit` | Stop the server right away. |
| `Broadcast` | `/Broadcast <message>` | Send a message to all connected players. |
| `KickPlayer` | `/KickPlayer <steamid>` | Kick a player by SteamID. |
| `BanPlayer` | `/BanPlayer <steamid>` | Ban a player by SteamID. |
| `TeleportToPlayer` | `/TeleportToPlayer <steamid>` | Teleport yourself to the specified player. |
| `TeleportToMe` | `/TeleportToMe <steamid>` | Teleport the specified player to you. |
| `ShowPlayers` | `/ShowPlayers` | Show the players who are connected right now. |
| `Info` | `/Info` | Show basic server information. |
| `Save` | `/Save` | Save the world data right now. |
| `UnBanPlayer` | `/UnBanPlayer <steamid>` | Remove a SteamID from the ban list. |
| `ToggleSpectate` | `/ToggleSpectate` | Toggle spectator mode. |
| `REST: Info` | `GET /v1/api/info` | Get basic server information by using the REST API. |
| `REST: Players` | `GET /v1/api/players` | Get the current player list by using the REST API. |
| `REST: Settings` | `GET /v1/api/settings` | Get the current server settings by using the REST API. |
| `REST: Metrics` | `GET /v1/api/metrics` | Get server metrics by using the REST API. |
| `REST: Announce` | `POST /v1/api/announce {"message":"Hello from ECC"}` | Send a message to all players by using the REST API. |
| `REST: Kick` | `POST /v1/api/kick {"userid":"<steamid>","message":"<reason>"}` | Kick a player by using the REST API. |
| `REST: Ban` | `POST /v1/api/ban {"userid":"<steamid>","message":"<reason>"}` | Ban a player by using the REST API. |
| `REST: Unban` | `POST /v1/api/unban {"userid":"<steamid>"}` | Remove a player from the ban list by using the REST API. |
| `REST: Save` | `POST /v1/api/save` | Save the world data by using the REST API. |
| `REST: Shutdown` | `POST /v1/api/shutdown {"waittime":30,"message":"Server shutting down"}` | Shut down the server with a delay by using the REST API. |
| `REST: Stop` | `POST /v1/api/stop` | Stop the server right away by using the REST API. |

### Project Zomboid Commands

| Label | Command | Description |
| --- | --- | --- |
| `addalltowhitelist` | `addalltowhitelist` | Add all currently connected password-using users to the whitelist. |
| `additem` | `additem "<username>" "<module.item>" <count>` | Give an item to a player. |
| `adduser` | `adduser "<username>" "<password>"` | Add a new user to a whitelist-enabled server. |
| `addusertowhitelist` | `addusertowhitelist "<username>"` | Add one connected password-using user to the whitelist. |
| `addxp` | `addxp "<playername>" <perkname>=<xp>` | Give XP to a player. |
| `banid` | `banid <SteamID>` | Ban a SteamID. |
| `banuser` | `banuser "<username>" -ip -r "<reason>"` | Ban a user by name, with optional IP ban and reason. |
| `godmode` | `godmode "<username>" -true\|-false` | Turn invincibility on or off for a player. |
| `grantadmin` | `grantadmin "<username>"` | Give a user admin rights. |
| `invisible` | `invisible "<username>" -true\|-false` | Make a player invisible to zombies. |
| `kickuser` | `kickuser "<username>" -r "<reason>"` | Kick a user from the server. |
| `players` | `players` | Show the players who are connected right now. |
| `releasesafehouse` | `releasesafehouse` | Release a safehouse you own. |
| `removeadmin` | `removeadmin "<username>"` | Remove admin rights from a user. |
| `removeuserfromwhitelist` | `removeuserfromwhitelist "<username>"` | Remove a user from the whitelist. |
| `servermsg` | `servermsg "<message>"` | Send a message to all connected players. |
| `setaccesslevel` | `setaccesslevel "<username>" "<accesslevel>"` | Set a player access level, like admin or moderator. |
| `teleport` | `teleport "<playername>" or teleport "<player1>" "<player2>"` | Teleport yourself to another player, or teleport one player to another. |
| `unbanid` | `unbanid <SteamID>` | Remove a SteamID from the ban list. |
| `unbanuser` | `unbanuser "<username>"` | Remove a user from the ban list. |
| `voiceban` | `voiceban "<username>" -true\|-false` | Block or unblock a player's voice chat. |
| `changeoption` | `changeoption <optionName> "<newValue>"` | Change a server option to a new value. |
| `checkModsNeedUpdate` | `checkModsNeedUpdate` | Check whether installed mods need updates and write the result to the log. |
| `clear` | `clear` | Clear the server console output. |
| `connections` | `connections` | Show the current connection information. |
| `help` | `help` | Show help and list the server commands. |
| `log` | `log <debugType> <severity>` | Set the log level for a debug or logging category. |
| `quit` | `quit` | Save the world and shut down the server. |
| `reloadlua` | `reloadlua "<filename>"` | Reload a Lua file on the server. |
| `reloadoptions` | `reloadoptions` | Reload the server options file and send the new values to clients. |
| `save` | `save` | Save the world right now. |
| `showoptions` | `showoptions` | Show the current server options and their values. |
| `stats` | `stats none\|file\|console\|all <period>` | Turn server statistics output on, off, or to a different target and interval. |
| `addvehicle` | `addvehicle "<script>" "<user or x,y,z>"` | Spawn a vehicle for a player or at a set location. |
| `alarm` | `alarm` | Trigger a building alarm at the current admin location. |
| `chopper` | `chopper` | Trigger a helicopter event. |
| `createhorde` | `createhorde <count> "<username>"` | Spawn a zombie horde near a player. |
| `createhorde2` | `createhorde2` | Use an alternate horde-spawn command. This command is still poorly documented. |
| `gunshot` | `gunshot` | Play a gunshot sound event on a random player. |
| `lightning` | `lightning "<username>"` | Trigger lightning, optionally on a player. |
| `removezombies` | `removezombies` | Remove zombies from the current area. |
| `startrain` | `startrain "<intensity>"` | Start rain with the chosen intensity. |
| `startstorm` | `startstorm "<duration>"` | Start a storm for a set duration. |
| `stoprain` | `stoprain` | Stop the rain. |
| `stopweather` | `stopweather` | Stop the active weather effects. |
| `teleportto` | `teleportto <x,y,z>` | Teleport to a set x, y, z location. |
| `thunder` | `thunder "<username>"` | Trigger thunder, optionally on a player. |
| `debugplayer` | `debugplayer` | Show debug information for a player. |
| `noclip` | `noclip "<username>" -true\|-false` | Turn noclip on or off for a player. |
| `replay` | `replay "<playername>" -record\|-play\|-stop <filename>` | Start or manage a replay for a user. This is a legacy command. |

### Minecraft Commands

| Label | Command | Description |
| --- | --- | --- |
| `help` | `help` | Show the available server commands. |
| `list` | `list` | Show online players and the current player count. |
| `say` | `say <message>` | Broadcast a chat message to everyone on the server. |
| `msg` | `msg <player> <message>` | Send a private message to one player. |
| `save-all` | `save-all` | Force the server to save the world and player data. |
| `save-off` | `save-off` | Turn off automatic world saving. |
| `save-on` | `save-on` | Turn automatic world saving back on. |
| `stop` | `stop` | Stop the Minecraft server cleanly. |
| `kick` | `kick <player> [reason]` | Kick a player from the server, with an optional reason. |
| `ban` | `ban <player> [reason]` | Ban a player from the server. |
| `pardon` | `pardon <player>` | Remove a player from the ban list. |
| `ban-ip` | `ban-ip <address\|player> [reason]` | Ban an IP address directly or by using a player name. |
| `pardon-ip` | `pardon-ip <address>` | Remove an IP address from the banned IP list. |
| `banlist` | `banlist [ips\|players]` | Show the current player bans or IP bans. |
| `whitelist on` | `whitelist on` | Turn whitelist enforcement on. |
| `whitelist off` | `whitelist off` | Turn whitelist enforcement off. |
| `whitelist list` | `whitelist list` | Show the players currently on the whitelist. |
| `whitelist reload` | `whitelist reload` | Reload the whitelist from disk. |
| `whitelist add` | `whitelist add <player>` | Add a player to the whitelist. |
| `whitelist remove` | `whitelist remove <player>` | Remove a player from the whitelist. |
| `op` | `op <player>` | Grant operator permissions to a player. |
| `deop` | `deop <player>` | Remove operator permissions from a player. |
| `difficulty` | `difficulty <peaceful\|easy\|normal\|hard>` | Change the server difficulty. |
| `defaultgamemode` | `defaultgamemode <survival\|creative\|adventure\|spectator>` | Set the default game mode for new players. |
| `gamemode` | `gamemode <survival\|creative\|adventure\|spectator> <player>` | Change one player's game mode. |
| `time set` | `time set <day\|night\|noon\|midnight\|value>` | Set the world time. |
| `weather` | `weather <clear\|rain\|thunder> [duration]` | Change the current weather. |

### Satisfactory Commands

| Label | Command | Description |
| --- | --- | --- |
| `quit` | `quit` | Shut down the server cleanly. |
| `stop` | `stop` | Stop the server. This does the same thing as quit. |
| `exit` | `exit` | Close the server process. |
| `SaveGame` | `SaveGame <saveName>` | Force a save by using the save name you provide. |
| `FG.AutosaveInterval` | `FG.AutosaveInterval <seconds>` | Set how often autosave runs, in seconds. |
| `FG.NetworkQuality` | `FG.NetworkQuality <value>` | Set the network quality preset. |
| `FG.DisableSeasonalEvents` | `FG.DisableSeasonalEvents <0\|1>` | Turn seasonal events on or off. |

### Valheim Commands

| Label | Command | Description |
| --- | --- | --- |
| `ban` | `ban <name\|ip\|userID>` | Ban a player by name, IP, or userID. |
| `banned` | `banned` | Show the current ban list. |
| `kick` | `kick <name\|ip\|userID>` | Kick a player from the server by name, IP, or userID. |
| `resetworldkeys` | `resetworldkeys` | Reset the world progression keys, like boss and world unlock flags. |
| `save` | `save` | Save the world right now. |
| `setworldmodifier` | `setworldmodifier <name> <value>` | Set one world modifier to a new value. |
| `setworldpreset` | `setworldpreset <name>` | Apply a world preset by name. |
| `unban` | `unban <name\|ip\|userID>` | Remove a player from the ban list by name, IP, or userID. |

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

## Item Spawner Guides

ECC has a right-side spawner panel in the Commands window for some games.

Supported spawner tools:

| Game | Item spawner | Vehicle spawner |
| --- | --- | --- |
| Minecraft | Yes | No |
| Hytale | Yes | No |
| Project Zomboid | Yes | Yes |

The spawner does not instantly run the command when you pick an item. It builds the command, lets you review it, and then you press `Send`.

### Common Item Spawner Layout

For Minecraft, Hytale, and Project Zomboid, the Commands window can show an `Item Spawner` panel on the right.

The panel has:

- `Online Player` box.
- `Refresh` button for players.
- `Items` tab.
- `Count` number box.
- `Search Items` box.
- Item list.
- Item name/type preview.
- Optional item preview image when ECC has one.
- `Insert Into Command Box` button.
- `Build /give` or `Build /additem` button.

Use `Refresh` to load online players when the server supports player lookup. You can also type a player name manually.

### Minecraft Item Spawner

The Minecraft item spawner appears in the Minecraft Commands window.

It builds `/give` commands.

Command format:

```text
/give PlayerName minecraft:item_id 5
```

Example:

```text
/give Steve minecraft:diamond 5
```

Workflow:

1. Open the Minecraft server card.
2. Click `Commands`.
3. Find the `Item Spawner` panel on the right.
4. Click `Refresh` beside `Online Player` if the server is running and player lookup works.
5. Pick an online player or type a player name.
6. Set `Count`.
7. Search for an item.
8. Select the item.
9. Click `Insert Into Command Box` or `Build /give`.
10. Review the command.
11. Click `Send`.

Where Minecraft item data comes from:

- ECC first uses `Config\MinecraftVanillaItemCatalog.json`.
- ECC may also inspect local Minecraft server/JAR assets when available.
- Optional cached icons may appear if ECC can build or load them.

If the Minecraft item list is empty:

- Make sure `Config\MinecraftVanillaItemCatalog.json` exists.
- Make sure the profile `FolderPath` points to the Minecraft server folder.
- Make sure `Executable` points to a real server JAR or Java launch target.
- Check Program Log for Minecraft catalog or jar scan errors.

Important Minecraft notes:

- The spawner only builds the command. The server must accept the command route ECC uses.
- RCON is useful for Minecraft command/player features.
- If RCON is not set up, ECC may fall back to another route, but that depends on the running server and profile.
- The player must exist on the server for `/give` to work.
- Item IDs usually look like `minecraft:diamond`, `minecraft:oak_log`, or `minecraft:iron_sword`.

### Hytale Item Spawner

The Hytale item spawner appears in the Hytale Commands window.

It builds `/give` commands with a quantity flag.

Command format:

```text
/give PlayerName item_id --quantity=5
```

Example:

```text
/give PlayerName hytale:example_item --quantity=5
```

Workflow:

1. Open the Hytale server card.
2. Click `Commands`.
3. Find the `Item Spawner` panel on the right.
4. Click `Refresh` beside `Online Player` if player lookup works for your server.
5. Pick an online player or type a player name.
6. Set `Count`.
7. Search for an item.
8. Select the item.
9. Click `Insert Into Command Box` or `Build /give`.
10. Review the command.
11. Click `Send`.

Where Hytale item data comes from:

- ECC first uses `Config\HytaleVanillaItemCatalog.json`.
- ECC may use local Hytale server folders to find item icons when available.
- Hytale icon lookup may use paths under the local server folder such as `Assets\Common\Icons\ItemsGenerated`.
- Optional cached icons may appear if ECC can build or load them.

If the Hytale item list is empty:

- Make sure `Config\HytaleVanillaItemCatalog.json` exists.
- Make sure the profile `FolderPath` points to the Hytale server folder.
- Make sure `AssetFile` and `AOTCache` match your local Hytale setup if your server needs them.
- Check Program Log for Hytale catalog or icon cache errors.

Important Hytale notes:

- Hytale support depends on the current server build and command syntax available to your server.
- The spawner builds the command ECC expects, but the server still decides whether that command is valid.
- If player lookup does not work, type the player name manually.

### Project Zomboid Item and Vehicle Spawner

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

The public GitHub repo may not include cached preview images or generated asset cache files.

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
|-- LICENSE
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

ECC does not currently use a top-level `Scripts\` folder as an add-on or plugin system.

## Maintainer Packaging Notes

This section is for the person building a public release or pushing files to GitHub. New users can skip it.

Safe to include in the public ECC repo or release ZIP:

- `Start.bat`
- `Launch.ps1`
- `Modules\*.psm1`
- Public/default JSON files in `Config`
- `README.md`
- `LICENSE`

Do not include in the public repo or release ZIP:

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

## Known Issues and Current Limits

This section lists things that are known limits in ECC right now. Some are bugs, and some are features that are only partly built.

| Area | Current limit | What to do |
| --- | --- | --- |
| Windows only | ECC is built around Windows PowerShell, Windows Forms, `.bat` launch files, and local Windows paths. | Run ECC on Windows. |
| First launch security | Windows may block downloaded PowerShell files or ask whether you trust them. | Extract the ZIP first, then run `Start.bat` as administrator from a trusted download. |
| Discord commands | ECC uses a bot token/channel polling flow plus webhook output. It is not a Discord slash-command app. | Use text commands like `!PZ start` in the monitored channel. |
| Discord secrets | `Config\Settings.json` stores Discord tokens/webhooks locally after saving settings. | Do not upload that file. Rotate tokens if they were exposed. |
| Profile secrets | Profiles can contain RCON, telnet, REST, API, and path details. | Keep real profiles private unless cleaned first. |
| Asset previews | Public packages may not include cached game images or extracted item preview data. | Use local game files or an optional asset cache package when available. |
| Project Zomboid spawner | PZ item/vehicle lists depend on local files and/or cached catalog data. Missing cache can make previews or lists incomplete. | Check `FolderPath` and optional `Config\AssetCache\ProjectZomboid`. |
| Minecraft spawner | Minecraft has an item spawner only. | Use the item spawner to build `/give` commands. |
| Hytale spawner | Hytale has an item spawner only. | Use the item spawner to build `/give` commands. |
| Project Zomboid vehicle spawner | Project Zomboid is the only supported vehicle spawner right now. | Use the PZ Commands window vehicle tab. |
| Valheim player listing | Live Valheim player listing is not supported yet. | Use server logs or Valheim-side tools for player checks. |
| Command permissions | ECC can send a command, but the game server may reject it if admin/RCON/API permissions are wrong. | Check the game's own permissions and use `Verbose Debug`. |
| Config editing | The `Config` button only works when ECC can find a valid config root for that profile. | Set `ConfigRoot` or `ConfigRoots` correctly. |
| Top-level scripts folder | ECC does not currently support a top-level `Scripts\` add-on or plugin system. | Put server start scripts in the server folder and point the profile `Executable` to them. |
| Debug/perf logs | Debug and performance tracing can create large logs with private paths/details. | Turn them off after testing and clean logs before sharing. |

### Developer-Tracked Known Bugs

These are not always first-time setup problems. They are bugs or risks already noticed during ECC testing and development.

| Bug | What can happen | Current workaround |
| --- | --- | --- |
| 7 Days to Die delayed trusted player data | 7 Days to Die can sit in `waiting for live player data` long enough that a first-player idle shutdown may fire late after trusted telnet/player data finally becomes available. | Watch the dashboard state and logs after startup. If needed, manually stop idle servers until this detection path is hardened. |
| Minecraft delayed trusted player data | Minecraft can show the same kind of delayed trusted player-data behavior as 7 Days to Die right now, which can affect first-player idle shutdown timing. | Watch the dashboard state and logs after startup. If needed, manually stop idle servers until this detection path is hardened. |
| Top bar player count mismatch | A server card may show players online while the top bar still says `PLAYERS: 0`. | Trust the server card/player log over the top bar if they disagree. |
| Bottom-right resize affordance drift | The main-window bottom-right resize grip can drift after some normal drag-resize paths. | Maximize/restore or let the window self-correct on the next layout pass. |
| Rare delayed WinForms event-handler risk | Most button and popup handler paths were hardened, but unusual delayed event paths may still have scope-sensitive failures. | Reopen the affected window or restart ECC if a button/popup acts dead or throws a UI handler error. |
| Program Log UI cost | Heavy Program Log rendering can contribute to UI stutter during noisy sessions. | Turn off debug/perf tracing when not testing, and clear logs when the UI gets heavy. |
| Main GUI timer load | The main GUI timer still owns many update jobs, so one slow branch can delay dashboard/log/header refreshes. | Reduce live log noise and use Performance Trace Mode when investigating lag. |

If something is not listed here, check the Troubleshooting section first, then check `Program Log` and `Logs\console.log`.

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

Dont sell my fucking code, this code is free for anyone, if your busted selling my code, thats fucked up.

## Credits

Built by Darkjesusmn and AI contributors.
