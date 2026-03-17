# Etherium Command Center

Etherium Command Center is a Windows desktop control panel for managing dedicated game servers with a built-in Discord bot, live logs, config editor, and safe shutdown tools. It is designed for single-machine hosting and focuses on clarity, stability, and quick recovery.

---

## Features

- Multi-game profiles with start, stop, restart, and status controls
- Discord bot commands (start/stop/restart/save/status/players/etc.)
- Per-game live log tabs with auto-detection of log files
- Config editor for server INI/CFG/JSON files
- Auto-restart and crash protection with rate limiting
- REST API, RCON, stdin, HTTP, and script command support
- Debug mode with verbose logs and preserved admin console
- Clean UI with collapsible panels and profile editor grouping

---

## Requirements

- Windows 10/11
- PowerShell 5.1
- Administrator privileges (for process management)
- Discord bot token + channel ID (optional but recommended)

---

## Quick Start (New Install)

1. Download or clone the project.
2. Unzip to a folder (example: `C:\EtheriumCommandCenter`).
3. Run `Start.bat` **as Administrator**.
4. Open **Settings** and fill in:
   - Bot Token
   - Monitor Channel ID
   - Command Prefix (default `!`)
   - Webhook URL (optional)
5. Click **Save Settings**.
6. Click **+ Add Game** and choose your server folder.
7. If the game stores configs elsewhere, choose a config folder when prompted.
8. Use the center dashboard to start/stop servers.

---

## Discord Setup

1. Create a Discord Application at discord.com/developers.
2. Add a Bot and copy the Bot Token.
3. Enable **Message Content Intent**.
4. Invite the bot to your server with permissions to read/send messages.
5. Enable Developer Mode in Discord, then copy your channel ID.
6. Paste Token + Channel ID into **Settings** and save.

### Command Format

Commands use: `!<prefix> <command>`

Example:
- `!pz start`
- `!pw status`
- `!pw players`

The prefix per game is defined in each profile (for example, `PZ` for Project Zomboid).

---

## Profiles

Profiles live in `Profiles\*.json` and control how each server starts/stops.
The editor on the right side lets you change any value safely.

Common fields:
- `FolderPath` – server install folder
- `Executable` – server exe
- `LaunchArgs` – startup parameters
- `ConfigRoot` – config folder for the Config Editor
- `LogStrategy` – how log files are found
- `Commands` – supported actions (start/stop/restart/save/etc.)

---

## Log Tabs

When a server starts, a log tab is created automatically. Logs are detected based on the profile’s `LogStrategy`:

- `SingleFile` – `ServerLogPath`
- `NewestFile` – newest `*.log` in `ServerLogRoot`
- `PZSessionFolder` – Project Zomboid session logs

---

## Config Editor

Each profile may define `ConfigRoot` or `ConfigRoots` so the Config Editor can open INI/CFG/JSON/etc.

Supported extensions include:
`.ini .cfg .json .xml .yml .yaml .properties .conf .lua`

---

## Debug Mode

Debug Mode enables verbose logging and keeps the admin PowerShell window open.
It can be toggled in Settings. Changing it will restart the app and stop running servers.

---

## Safety + Shutdown

- Auto-restart is guarded by a max restarts per hour limit.
- Safe shutdown sends save commands before stopping.
- If save fails, a fallback kill is used after a delay.

---

## Troubleshooting

**Discord bot not responding**
- Check token, channel ID, and bot permissions.
- Make sure Message Content Intent is enabled.

**No logs appearing**
- Verify `ServerLogRoot` or `ServerLogPath` in the profile.
- Confirm the game is writing logs to that location.

**REST / RCON errors**
- Confirm ports and passwords in game config files.
- Check that the profile’s REST/RCON settings match your server config.

---

## Folder Structure

```
Config\
Logs\
Modules\
Profiles\
Launch.ps1
Start.bat
```

---

## License

Private use during development. Distribution terms pending.

---

## Credits

Built by DarkJesusMN + AI LLMs Claude ChatGPT Copilot


---

## Known Issues / Roadmap

### Known Issues
- Some games require manual tuning of REST/RCON settings for status and save commands.
- Log detection for uncommon games may require adjusting `LogStrategy` and log paths.
- UI polish is ongoing (tooltips, spacing, and advanced layout behaviors).

### Roadmap
- Profile templates for more dedicated servers (pre-filled defaults).
- Advanced log filters, search, and export.
- Per-game health checks with alerting.
- Optional remote control features for multi-host setups.



