# =============================================================================
# GUI.psm1  -  WinForms GUI for Etherium Command Center
# PowerShell 5.1 compatible. No emoji in source strings.
# =============================================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Collections
$script:LogBuffer = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()

# Console window hide/show helper (used to hide admin console in non-debug mode)
try { [NativeWin]::GetConsoleWindow() | Out-Null } catch {
    Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class NativeWin {
  [DllImport("kernel32.dll")]
  public static extern IntPtr GetConsoleWindow();
  [DllImport("user32.dll")]
  public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@
}

$script:ModuleRoot  = $PSScriptRoot
$script:ProfilesDir = '.\Profiles'   # set properly in Start-GUI

# -------------------------
# Theme, fonts, and helpers
# -------------------------
$clrBg     = [System.Drawing.Color]::FromArgb(17,17,27)
$clrPanel  = [System.Drawing.Color]::FromArgb(49,50,68)
$clrText   = [System.Drawing.Color]::FromArgb(220,220,230)
$clrAccent = [System.Drawing.Color]::FromArgb(100,120,255)
$clrGreen  = [System.Drawing.Color]::FromArgb(80,200,120)
$clrRed    = [System.Drawing.Color]::FromArgb(240,100,100)
$clrYellow = [System.Drawing.Color]::FromArgb(240,200,80)
$clrMuted  = [System.Drawing.Color]::FromArgb(120,120,140)
$clrBtnText = [System.Drawing.Color]::FromArgb(40,40,40)
$fontLabel = New-Object System.Drawing.Font("Segoe UI", 9)
$fontBold  = New-Object System.Drawing.Font("Segoe UI", 9,  [System.Drawing.FontStyle]::Bold)
$fontTitle = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$fontMono  = New-Object System.Drawing.Font("Consolas", 9)

# =============================================================================
#  CONTROL FACTORY HELPERS
# =============================================================================
function _Label {
    param($text, $x, $y, $w, $h, $font = $fontLabel)
    $lbl           = New-Object System.Windows.Forms.Label
    $lbl.Text      = $text
    $lbl.Location  = [System.Drawing.Point]::new($x, $y)
    $lbl.Size      = [System.Drawing.Size]::new($w, $h)
    $lbl.ForeColor = $clrText
    $lbl.BackColor = [System.Drawing.Color]::Transparent
    $lbl.Font      = $font
    return $lbl
}

function _TextBox {
    param($x, $y, $w, $h, $text = '', $pass = $false)
    $tb                       = New-Object System.Windows.Forms.TextBox
    $tb.Location              = [System.Drawing.Point]::new($x, $y)
    $tb.Size                  = [System.Drawing.Size]::new($w, $h)
    $tb.Text                  = $text
    $tb.UseSystemPasswordChar = $pass
    $tb.BackColor             = [System.Drawing.Color]::FromArgb(30,30,40)
    $tb.ForeColor             = $clrText
    $tb.BorderStyle           = 'FixedSingle'
    $tb.Font                  = $fontLabel
    return $tb
}

# All parameters positional: text, x, y, w, h, bg, onClick
function _Button {
    param($text, $x, $y, $w, $h, $bg, $onClick)
    $btn                              = New-Object System.Windows.Forms.Button
    $btn.Text                         = $text
    $btn.Location                     = [System.Drawing.Point]::new($x, $y)
    $btn.Size                         = [System.Drawing.Size]::new($w, $h)
    $btn.FlatStyle                    = 'Flat'
    $btn.FlatAppearance.BorderSize    = 0
    $btn.BackColor                    = $bg
    $btn.ForeColor                    = $clrBtnText
    $btn.Font                         = $fontLabel
    if ($onClick) { $btn.Add_Click($onClick) }
    return $btn
}

function _Panel {
    param([int]$X, [int]$Y, [int]$W, [int]$H, $BG = $null)
    $p             = [System.Windows.Forms.Panel]::new()
    $p.Location    = [System.Drawing.Point]::new($X, $Y)
    $p.Size        = [System.Drawing.Size]::new($W, $H)
    $p.BackColor   = if ($null -ne $BG -and $BG -is [System.Drawing.Color]) { $BG } else { $clrPanel }
    $p.BorderStyle = 'None'
    return $p
}

function _VerticalText {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $Text }
    return ($Text.ToCharArray() -join "`n")
}

# =============================================================================
#  SETTINGS SAVE HELPER
# =============================================================================
function _SaveSettings {
    param([hashtable]$Settings, [string]$ConfigPath)

    if ($null -eq $Settings)                          { throw "Settings hashtable is null" }
    if ([string]::IsNullOrWhiteSpace($ConfigPath))    { throw "ConfigPath is null or empty: '$ConfigPath'" }

    $dir = Split-Path -Path $ConfigPath -Parent
    if (-not (Test-Path -Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }

    $Settings | ConvertTo-Json -Depth 5 | Set-Content -Path $ConfigPath -Encoding UTF8 -Force
}

# =============================================================================
#  SETTINGS PANEL  (opened from the top-bar Settings button)
# =============================================================================
function _BuildSettingsTab {
    param(
        [System.Windows.Forms.TabPage]$Tab,
        [hashtable]$Settings,
        [string]$ConfigPath,
        [hashtable]$SharedState
    )

    if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
        throw "_BuildSettingsTab: ConfigPath is empty"
    }

    # Store in script scope so the save-button closure can reach them
    $script:SettingsConfigPath = $ConfigPath
    $script:SettingsData       = $Settings
    # Always resolve SharedState from script scope
    $SharedState = $script:SharedState

    if ($null -eq $SharedState) { throw "_BuildSettingsTab: SharedState is null" }

    # Ensure all required keys exist
    foreach ($key in @('BotToken','WebhookUrl','MonitorChannelId','CommandPrefix','PollIntervalSeconds','EnableDebugLogging')) {
        if (-not $Settings.ContainsKey($key)) { $Settings[$key] = '' }
    }

    $Tab.BackColor = $clrBg
    $y = 20

    $Tab.Controls.Add((_Label 'Bot Settings' 20 $y 400 28 $fontTitle))
    $y += 44

    $Tab.Controls.Add((_Label 'Bot Token (keep secret)' 20 $y 300 20))
    $y += 22
    $tbToken = _TextBox 20 $y 480 24 ($Settings.BotToken) $true
    $Tab.Controls.Add($tbToken)
    $y += 36

    $Tab.Controls.Add((_Label 'Webhook URL' 20 $y 300 20))
    $y += 22
    $tbWebhook = _TextBox 20 $y 480 24 ($Settings.WebhookUrl)
    $Tab.Controls.Add($tbWebhook)
    $y += 36

    $Tab.Controls.Add((_Label 'Monitor Channel ID' 20 $y 300 20))
    $y += 22
    $tbChannel = _TextBox 20 $y 240 24 ($Settings.MonitorChannelId)
    $Tab.Controls.Add($tbChannel)
    $y += 36

    $Tab.Controls.Add((_Label 'Command Prefix (default: !)' 20 $y 300 20))
    $y += 22
    $tbPrefix = _TextBox 20 $y 80 24 ($Settings.CommandPrefix)
    $Tab.Controls.Add($tbPrefix)
    $y += 36

    $Tab.Controls.Add((_Label 'Poll Interval in seconds (default: 2)' 20 $y 300 20))
    $y += 22
    $tbPoll = _TextBox 20 $y 80 24 ([string]$Settings.PollIntervalSeconds)
    $Tab.Controls.Add($tbPoll)
    $y += 36

    $Tab.Controls.Add((_Label 'Debug Logging (verbose)' 20 $y 300 20))
    $y += 22
    $chkDebug           = New-Object System.Windows.Forms.CheckBox
    $chkDebug.Location  = [System.Drawing.Point]::new(20, $y)
    $chkDebug.Size      = [System.Drawing.Size]::new(200, 20)
    $chkDebug.Text      = 'Enabled'
    $chkDebug.ForeColor = $clrText
    $chkDebug.BackColor = [System.Drawing.Color]::Transparent
    $chkDebug.Font      = $fontLabel
    $chkDebug.Checked   = ($Settings.EnableDebugLogging -eq $true)
    $Tab.Controls.Add($chkDebug)
    $y += 44

    # Store textboxes in script scope so the click closure can access them
    $script:SettingsTabToken   = $tbToken
    $script:SettingsTabWebhook = $tbWebhook
    $script:SettingsTabChannel = $tbChannel
    $script:SettingsTabPrefix  = $tbPrefix
    $script:SettingsTabPoll    = $tbPoll
    $script:SettingsTabDebug   = $chkDebug

    $saveBtn = _Button 'Save Settings' 20 $y 140 32 $clrGreen {
        $cfgPath = $script:SettingsConfigPath
        $stngs   = $script:SettingsData
        $st      = $script:SharedState

        if ($null -eq $st)                             { [System.Windows.Forms.MessageBox]::Show('SharedState is null','Error','OK','Error') | Out-Null; return }
        if ([string]::IsNullOrWhiteSpace($cfgPath))    { [System.Windows.Forms.MessageBox]::Show("ConfigPath is empty: '$cfgPath'",'Error','OK','Error') | Out-Null; return }
        if ($null -eq $stngs)                          { $stngs = @{}; $st['Settings'] = $stngs }

        $wasDebug = $false
        try { $wasDebug = [bool]$stngs['EnableDebugLogging'] } catch { $wasDebug = $false }

        try {
            $stngs['BotToken']         = $script:SettingsTabToken.Text.Trim()
            $stngs['WebhookUrl']       = $script:SettingsTabWebhook.Text.Trim()
            $stngs['MonitorChannelId'] = $script:SettingsTabChannel.Text.Trim()
            $stngs['CommandPrefix']    = if ($script:SettingsTabPrefix.Text.Trim()) { $script:SettingsTabPrefix.Text.Trim() } else { '!' }
            $intVal = 0
            $stngs['PollIntervalSeconds'] = if ([int]::TryParse($script:SettingsTabPoll.Text, [ref]$intVal)) { $intVal } else { 2 }
            $stngs['EnableDebugLogging']   = ($script:SettingsTabDebug.Checked -eq $true)
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error reading values: $_",'Error','OK','Error') | Out-Null
            return
        }

        try {
            _SaveSettings -Settings $stngs -ConfigPath $cfgPath
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to save settings: $_",'Save Error','OK','Error') | Out-Null
            return
        }

        try {
            if (Get-Command Set-DebugLoggingEnabled -ErrorAction SilentlyContinue) {
                Set-DebugLoggingEnabled -Enabled ([bool]$stngs['EnableDebugLogging'])
            }
        } catch { }

        $debugChanged = ($wasDebug -ne [bool]$stngs['EnableDebugLogging'])
        if ($debugChanged) {
            $resp = [System.Windows.Forms.MessageBox]::Show(
                'Changing debug mode will restart the app and stop all servers. Continue?',
                'Confirm Restart','YesNo','Warning')
            if ($resp -ne [System.Windows.Forms.DialogResult]::Yes) {
                # Keep previous debug setting, save everything else
                $stngs['EnableDebugLogging'] = $wasDebug
            } else {
                $st['RestartProgram'] = $true
                [System.Windows.Forms.MessageBox]::Show(
                    'Debug mode changed. The app will now restart and all servers will be stopped.',
                    'Restart Required','OK','Information') | Out-Null
                if ($script:MainForm) { $script:MainForm.Close() }
                return
            }
        }

        $st['RestartListener'] = $true
        [System.Windows.Forms.MessageBox]::Show(
            'Settings saved. The Discord listener will reconnect automatically.',
            'Saved','OK','Information') | Out-Null
    }
    $Tab.Controls.Add($saveBtn)

    $y += 50
    $info = "How to get these values:" + [Environment]::NewLine +
            "  Bot Token   : discord.com/developers -> Your App -> Bot -> Reset Token" + [Environment]::NewLine +
            "                Also enable 'Message Content Intent' on the Bot page." + [Environment]::NewLine +
            "  Webhook URL : Channel Settings -> Integrations -> Webhooks -> New Webhook" + [Environment]::NewLine +
            "  Channel ID  : Discord Settings -> Advanced -> enable Developer Mode" + [Environment]::NewLine +
            "                Then right-click your command channel -> Copy Channel ID"

    $note      = _Label $info 20 $y 600 120
    $note.Font = $fontLabel
    $note.AutoSize = $false
    $note.Anchor = 'Top,Left,Right'
    $note.Width = [Math]::Max(300, $Tab.ClientSize.Width - 40)
    $note.Name = 'SettingsNote'
    $Tab.Controls.Add($note)

    # Keep the note sized to the tab width so text doesn't clip
    $Tab.add_Resize({
        $n = $Tab.Controls['SettingsNote']
        if ($n -and $n.PSObject.Properties.Match('Width').Count -gt 0) {
            $n.Width = [Math]::Max(300, $Tab.ClientSize.Width - 40)
        }
    })
}

# =============================================================================
#  SCRIPT-SCOPE SHARED STATE  (set once in Start-GUI, read everywhere)
# =============================================================================
$script:SharedState = $null

# =============================================================================
#  MAIN GUI ENTRY POINT
# =============================================================================
function Start-GUI {
    param(
        [hashtable]$SharedState,
        [string]$ConfigPath      = '.\Config\Settings.json',
        [string]$ProfilesDir     = '.\Profiles',
        [string]$LogsDir         = '.\Logs',
        [string]$TranscriptPath  = ''
    )

    if ($null -eq $SharedState -or -not ($SharedState -is [hashtable])) {
        throw "Start-GUI was called without a valid SharedState hashtable."
    }

    $script:SharedState = $SharedState
    $script:ProfilesDir = $ProfilesDir
    if (-not $script:_SelectedProfilePrefix) { $script:_SelectedProfilePrefix = $null }

    # Ensure keys that the timer reads always exist
    if (-not $SharedState.ContainsKey('ListenerRunning')) { $SharedState['ListenerRunning'] = $false }
    if (-not $SharedState.ContainsKey('StopListener'))    { $SharedState['StopListener']    = $false }
    if (-not $SharedState.ContainsKey('GameLogQueue'))    { $SharedState['GameLogQueue']    = [System.Collections.Concurrent.ConcurrentQueue[object]]::new() }
    if (-not $SharedState.ContainsKey('PlayersRequests')) { $SharedState['PlayersRequests'] = [hashtable]::Synchronized(@{}) }

    # ── Log file tail state ────────────────────────────────────────────────
    # We tail the transcript file (console.log) which captures all console output.
    $script:LogFilePath    = $null
    $script:LogFilePos     = 0L
    $script:LogsDir        = $LogsDir
    # Game server log tails: key = "PREFIX::filepath", value = byte position
    $script:GameLogTails   = @{}
    # If a transcript path was passed directly, use it immediately
    if (-not [string]::IsNullOrEmpty($TranscriptPath)) {
        $script:LogFilePath = $TranscriptPath
        $script:LogFilePos  = 0L
    }

    # Layout constants
    $leftWidth        = 250
    $rightWidth       = 620
    $sideGap          = 30
    $topBarHeight     = 60
    $bottomLogsHeight = 260
    $defaultWidth     = 1920
    $defaultHeight    = 1080
    $minWidth         = 1600
    $minHeight        = 900
    $tabFont = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)

    # =====================================================================
    # MAIN FORM
    # =====================================================================
    $form               = [System.Windows.Forms.Form]::new()
    $script:MainForm    = $form
    $form.Text          = 'Etherium Command Center Dashboard'
    $form.Size          = [System.Drawing.Size]::new($defaultWidth, $defaultHeight)
    $form.MinimumSize   = [System.Drawing.Size]::new($minWidth, $minHeight)
    $form.BackColor     = $clrBg
    $form.StartPosition = 'CenterScreen'
    $form.Icon          = [System.Drawing.SystemIcons]::Application

    # Hide the admin console after the GUI is visible when debug is off
    $form.Add_Shown({
        try {
            $dbg = $false
            if ($script:SharedState -and $script:SharedState.Settings -and $script:SharedState.Settings.ContainsKey('EnableDebugLogging')) {
                $dbg = [bool]$script:SharedState.Settings.EnableDebugLogging
            }
            if (-not $dbg) {
                $h = [NativeWin]::GetConsoleWindow()
                if ($h -ne [IntPtr]::Zero) { [NativeWin]::ShowWindow($h, 0) | Out-Null }
            }
        } catch { }
    })

    # =====================================================================
    # STATUS BAR
    # =====================================================================
    $statusBar           = [System.Windows.Forms.StatusStrip]::new()
    $statusBar.BackColor = $clrPanel
    $statusLabel         = [System.Windows.Forms.ToolStripStatusLabel]::new()
    $statusLabel.Text      = 'Ready'
    $statusLabel.ForeColor = $clrText
    $statusLabel.Font      = $fontLabel
    $statusBar.Items.Add($statusLabel) | Out-Null
    $statusBar.Dock = 'Bottom'
    $form.Controls.Add($statusBar)

    # =====================================================================
    # TOP BAR
    # =====================================================================
    $topBar        = _Panel 0 0 $defaultWidth $topBarHeight $clrPanel
    $topBar.Anchor = 'Top,Left,Right'
    $form.Controls.Add($topBar)

    $lblCPU = _Label 'CPU: --%'    20  20 100 20
    $lblRAM = _Label 'RAM: --%'    120 20 100 20
    $lblNET = _Label 'NET: -- KB/s' 220 20 160 20
    $lblBot = _Label 'Bot: Unknown' 380 20 150 20
    $topBar.Controls.Add($lblCPU)
    $topBar.Controls.Add($lblRAM)
    $topBar.Controls.Add($lblNET)
    $topBar.Controls.Add($lblBot)

    $btnBotStart = _Button 'Start Bot' 540 15 100 30 $clrGreen { $script:SharedState['RestartListener'] = $true }
    $btnBotStop  = _Button 'Stop Bot'  650 15 100 30 $clrRed   { $script:SharedState['StopListener']    = $true }
    $topBar.Controls.Add($btnBotStart)
    $topBar.Controls.Add($btnBotStop)

    $btnSettings = _Button 'Settings' ($defaultWidth - 170) 15 100 30 $clrPanel {
        $settingsForm                  = New-Object System.Windows.Forms.Form
        $settingsForm.Text             = 'Settings'
        $settingsForm.Size             = [System.Drawing.Size]::new(700, 650)
        $settingsForm.MinimumSize      = [System.Drawing.Size]::new(650, 600)
        $settingsForm.StartPosition    = 'CenterParent'
        $settingsForm.BackColor        = $clrBg
        $settingsForm.FormBorderStyle  = 'FixedDialog'
        $settingsForm.MaximizeBox      = $false
        $settingsForm.MinimizeBox      = $false

        $tc      = New-Object System.Windows.Forms.TabControl
        $tc.Dock = 'Fill'
        $settingsForm.Controls.Add($tc)

        $tab           = New-Object System.Windows.Forms.TabPage
        $tab.Text      = 'Settings'
        $tab.BackColor = $clrBg
        $tab.AutoScroll = $true
        $tc.TabPages.Add($tab) | Out-Null

        _BuildSettingsTab `
            -Tab        $tab `
            -Settings   $script:SharedState.Settings `
            -ConfigPath $ConfigPath `
            -SharedState $script:SharedState

        $settingsForm.ShowDialog() | Out-Null
    }
    $btnSettings.Anchor = 'Top,Right'
    $topBar.Controls.Add($btnSettings)

    # =====================================================================
    # THREE-COLUMN MIDDLE SECTION
    # =====================================================================
    $collapsedSize      = 22
    $headerHeight       = 26
    $bottomHeaderHeight = 24

    $script:_LeftCollapsed   = $false
    $script:_RightCollapsed  = $false
    $script:_BottomCollapsed = $false

    # Left container (Profiles)
    $leftContainer = _Panel 0 $topBarHeight $leftWidth 600 $clrPanel
    $leftContainer.Anchor = 'Top,Left,Bottom'
    $leftContainer.BorderStyle = 'FixedSingle'
    $form.Controls.Add($leftContainer)
    $script:_LeftContainer = $leftContainer

    $leftHeader = _Panel 0 0 $leftWidth $headerHeight $clrPanel
    $leftHeader.BackColor = [System.Drawing.Color]::FromArgb(45,46,62)
    $leftHeader.Anchor = 'Top,Left,Right'
    $leftContainer.Controls.Add($leftHeader)
    $script:_LeftHeader = $leftHeader

    $leftHeaderLabel = _Label 'Game Profiles' 6 4 200 18 $fontBold
    $leftHeaderLabel.ForeColor = $clrAccent
    $leftHeader.Controls.Add($leftHeaderLabel)
    $script:_LeftHeaderLabel = $leftHeaderLabel

    $leftBody = _Panel 0 $headerHeight $leftWidth (600 - $headerHeight) $clrPanel
    $leftBody.Anchor = 'Top,Left,Right,Bottom'
    $leftBody.BackColor = $clrPanel
    $leftContainer.Controls.Add($leftBody)
    $script:_ProfilesPanel = $leftBody

    # Right container (Profile Editor)
    $rightContainer = _Panel ($defaultWidth - $rightWidth) $topBarHeight $rightWidth 600 $clrPanel
    $rightContainer.Anchor = 'Top,Right,Bottom'
    $rightContainer.BorderStyle = 'FixedSingle'
    $form.Controls.Add($rightContainer)
    $script:_RightContainer = $rightContainer

    $rightHeader = _Panel 0 0 $rightWidth $headerHeight $clrPanel
    $rightHeader.BackColor = [System.Drawing.Color]::FromArgb(45,46,62)
    $rightHeader.Anchor = 'Top,Left,Right'
    $rightContainer.Controls.Add($rightHeader)
    $script:_RightHeader = $rightHeader

    $rightHeaderLabel = _Label 'Profile Editor' 6 4 200 18 $fontBold
    $rightHeaderLabel.ForeColor = $clrAccent
    $rightHeader.Controls.Add($rightHeaderLabel)
    $script:_RightHeaderLabel = $rightHeaderLabel

    $rightBody = _Panel 0 $headerHeight $rightWidth (600 - $headerHeight) $clrPanel
    $rightBody.Anchor = 'Top,Left,Right,Bottom'
    $rightBody.BackColor = $clrPanel
    $rightContainer.Controls.Add($rightBody)
    $script:_ProfileEditorPanel = $rightBody

    # Center dashboard
    $centerCol = _Panel ($leftWidth + $sideGap) $topBarHeight ($defaultWidth - $leftWidth - $rightWidth - ($sideGap * 2)) 600 $clrBg
    $centerCol.Anchor = 'Top,Left,Right,Bottom'
    $form.Controls.Add($centerCol)
    $script:_ServerDashboardPanel = $centerCol

    # =====================================================================
    # BOTTOM LOG STRIP
    # =====================================================================
    # =====================================================================
    # BOTTOM LOG STRIP  -  TabControl with Discord / Program / per-game tabs
    # =====================================================================
    $bottomContainer        = _Panel 0 ($defaultHeight - $bottomLogsHeight - 22) $defaultWidth $bottomLogsHeight $clrBg
    $bottomContainer.Anchor = 'Left,Right,Bottom'
    $bottomContainer.BorderStyle = 'FixedSingle'
    $form.Controls.Add($bottomContainer)
    $script:_BottomContainer = $bottomContainer

    $bottomHeader = _Panel 0 0 $defaultWidth $bottomHeaderHeight $clrPanel
    $bottomHeader.BackColor = [System.Drawing.Color]::FromArgb(45,46,62)
    $bottomHeader.Anchor = 'Top,Left,Right'
    $bottomContainer.Controls.Add($bottomHeader)
    $script:_BottomHeader = $bottomHeader

    $bottomHeaderLabel = _Label 'Logs' 8 4 200 16 $fontBold
    $bottomHeaderLabel.ForeColor = $clrAccent
    $bottomHeader.Controls.Add($bottomHeaderLabel)
    $script:_BottomHeaderLabel = $bottomHeaderLabel

    $bottomPanel        = _Panel 0 $bottomHeaderHeight $defaultWidth ($bottomLogsHeight - $bottomHeaderHeight) $clrBg
    $bottomPanel.Anchor = 'Left,Right,Bottom'
    $bottomContainer.Controls.Add($bottomPanel)
    $script:_BottomBody = $bottomPanel

    # Collapse toggle handlers (click header or label)
    $leftHeader.Cursor = 'Hand'
    $leftHeaderLabel.Cursor = 'Hand'
    $leftHeader.Add_Click({
        $script:_LeftCollapsed = -not $script:_LeftCollapsed
        _ReflowLayout
    })
    $leftHeaderLabel.Add_Click({
        $script:_LeftCollapsed = -not $script:_LeftCollapsed
        _ReflowLayout
    })

    $rightHeader.Cursor = 'Hand'
    $rightHeaderLabel.Cursor = 'Hand'
    $rightHeader.Add_Click({
        $script:_RightCollapsed = -not $script:_RightCollapsed
        _ReflowLayout
    })
    $rightHeaderLabel.Add_Click({
        $script:_RightCollapsed = -not $script:_RightCollapsed
        _ReflowLayout
    })

    $bottomHeader.Cursor = 'Hand'
    $bottomHeaderLabel.Cursor = 'Hand'
    $bottomHeader.Add_Click({
        $script:_BottomCollapsed = -not $script:_BottomCollapsed
        _ReflowLayout
    })
    $bottomHeaderLabel.Add_Click({
        $script:_BottomCollapsed = -not $script:_BottomCollapsed
        _ReflowLayout
    })

    # Master TabControl fills the entire bottom panel
    $logTabs            = New-Object System.Windows.Forms.TabControl
    $logTabs.Location   = [System.Drawing.Point]::new(0, 0)
    $logTabs.Size       = [System.Drawing.Size]::new($defaultWidth, $bottomLogsHeight)
    $logTabs.Anchor     = 'Left,Right,Top,Bottom'
    $logTabs.BackColor  = $clrBg
    $logTabs.Font       = $fontLabel
    $logTabs.DrawMode   = [System.Windows.Forms.TabDrawMode]::OwnerDrawFixed
    $logTabs.ItemSize   = [System.Drawing.Size]::new(110, 22)
    $logTabs.Add_DrawItem({
        param($s,$e)

        $tab = $logTabs.TabPages[$e.Index]

        $brush = if ($e.Index -eq $logTabs.SelectedIndex) {
            [System.Drawing.SolidBrush]::new($clrAccent)
        } else {
            [System.Drawing.SolidBrush]::new($clrPanel)
        }

        $e.Graphics.FillRectangle($brush, $e.Bounds)

        $txtBrush = [System.Drawing.SolidBrush]::new($clrText)

        $sf = New-Object System.Drawing.StringFormat
        $sf.Alignment = 'Center'
        $sf.LineAlignment = 'Center'

        # SAFE DrawString overload
        $rectF = [System.Drawing.RectangleF]::new(
            $e.Bounds.X, $e.Bounds.Y,
            $e.Bounds.Width, $e.Bounds.Height
        )

        $e.Graphics.DrawString($tab.Text, $tabFont, $txtBrush, $rectF, $sf)

        $brush.Dispose()
        $txtBrush.Dispose()
        $sf.Dispose()
    })

    $bottomPanel.Controls.Add($logTabs)
    $script:_LogTabControl = $logTabs

    # Helper: make a RichTextBox inside a TabPage
    function _MakeLogRTB {
        param([System.Windows.Forms.TabPage]$page)
        $rt            = [System.Windows.Forms.RichTextBox]::new()
        $rt.Dock       = 'Fill'
        $rt.BackColor  = [System.Drawing.Color]::FromArgb(17,17,27)
        $rt.ForeColor  = $clrText
        $rt.Font       = $fontMono
        $rt.ReadOnly   = $true
        $rt.ScrollBars = 'Vertical'
        $rt.WordWrap   = $false
        $page.Controls.Add($rt)
        return $rt
    }

    # Helper: append a colored line to a RichTextBox, cap at 1000 lines
    function _AppendLog {
        param(
            [System.Windows.Forms.RichTextBox]$rt,
            [string]$line,
            [System.Drawing.Color]$colour
        )

        if ($null -eq $rt) { return }

        # Matrix green for Project Zomboid logs
        try {
            if ($rt.Tag -and $rt.Tag.MatrixGreen) {
                $colour = [System.Drawing.Color]::FromArgb(0, 220, 80)
            }
        } catch { }

        # Max lines allowed in the RTB
        $maxLines = 50

        # Detect if user is already at bottom BEFORE adding text
        $isAtBottom = ($rt.SelectionStart -ge ($rt.TextLength - 200))

        $rt.SuspendLayout()

        # Use a fast append
        $rt.SelectionStart  = $rt.TextLength
        $rt.SelectionLength = 0
        $rt.SelectionColor  = $colour
        $rt.AppendText("$line`n")

        # Trim old lines if buffer too large
        $lineCount = $rt.Lines.Count
        if ($lineCount -gt $maxLines) {
            $removeCount = $lineCount - $maxLines
            $rt.Lines = $rt.Lines[$removeCount..($lineCount - 1)]
        }

        # Auto-scroll ONLY if user was already at bottom
        if ($isAtBottom) {
            $rt.SelectionStart = $rt.TextLength
            $rt.ScrollToCaret()
        }

        $rt.ResumeLayout($true)
    }

    # Helper: pick colour from log line content
    function _LogColour {
        param([string]$line)
        if ($line -match '\[ERROR\]|ERROR|Exception|FATAL')  { return $clrRed    }
        if ($line -match '\[WARN\]|WARN|Warning')             { return $clrYellow }
        if ($line -match 'connected|joined|started|Running')  { return $clrGreen  }
        if ($line -match 'disconnect|stopped|crash|killed')   { return $clrRed    }
        if ($line -match '\[Discord\]|Discord')               { return $clrAccent }
        if ($line -match '\[SERVER\]|SERVER')                 { return $clrAccent }
        return $clrText
    }

    # ── DISCORD TAB ──────────────────────────────────────────────────────────
    $tabDiscord           = New-Object System.Windows.Forms.TabPage
    $tabDiscord.Text      = 'Discord'
    $tabDiscord.BackColor = $clrBg
    $logTabs.TabPages.Add($tabDiscord)

    $discordInner         = New-Object System.Windows.Forms.Panel
    $discordInner.Dock    = 'Fill'
    $discordInner.BackColor = $clrBg
    $tabDiscord.Controls.Add($discordInner)

    $rtDiscord            = [System.Windows.Forms.RichTextBox]::new()
    $rtDiscord.Dock       = 'Fill'
    $rtDiscord.BackColor  = [System.Drawing.Color]::FromArgb(17,17,27)
    $rtDiscord.ForeColor  = $clrText
    $rtDiscord.Font       = $fontMono
    $rtDiscord.ReadOnly   = $true
    $rtDiscord.ScrollBars = 'Vertical'
    $rtDiscord.WordWrap   = $false
    $discordInner.Controls.Add($rtDiscord)
    $script:_DiscordLogBox = $rtDiscord

    # Send bar at bottom of Discord tab
    $discordFooter = New-Object System.Windows.Forms.Panel
    $discordFooter.Dock = 'Bottom'
    $discordFooter.Height = 32
    $discordFooter.BackColor = $clrBg
    $discordInner.Controls.Add($discordFooter)
    $script:_DiscordFooter = $discordFooter

    $tbSend  = _TextBox 10 5 ([Math]::Max(100, $discordFooter.Width - 220)) 22 ''
    $tbSend.Anchor = 'Left,Right,Top'
    $btnSend = _Button 'Send' ([Math]::Max(120, $discordFooter.Width - 205)) 3 90 26 $clrAccent {
        $msg = $tbSend.Text.Trim()
        if (-not $msg) { return }
        $msg = "[BOT] $msg"
        if ($script:SharedState.ContainsKey('DiscordOutbox') -and $script:SharedState.DiscordOutbox) {
            $script:SharedState.DiscordOutbox.Enqueue($msg)
        } else {
            $script:SharedState['SendDiscordMessage'] = $msg
        }
        $tbSend.Text = ''
    }
    $btnClearDisc = _Button 'Clear' ([Math]::Max(200, $discordFooter.Width - 110)) 3 70 26 $clrMuted {
        if ($script:_DiscordLogBox) { $script:_DiscordLogBox.Clear() }
    }
    $discordFooter.Controls.Add($tbSend)
    $discordFooter.Controls.Add($btnSend)
    $discordFooter.Controls.Add($btnClearDisc)

    $discordFooter.add_Resize({
        if ($tbSend) { $tbSend.Width = [Math]::Max(100, $discordFooter.Width - 220) }
        if ($btnSend) { $btnSend.Location = [System.Drawing.Point]::new([Math]::Max(120, $discordFooter.Width - 205), 3) }
        if ($btnClearDisc) { $btnClearDisc.Location = [System.Drawing.Point]::new([Math]::Max(200, $discordFooter.Width - 110), 3) }
    })

    # ── PROGRAM LOG TAB ───────────────────────────────────────────────────────
    $tabProgram           = New-Object System.Windows.Forms.TabPage
    $tabProgram.Text      = 'Program Log'
    $tabProgram.BackColor = $clrBg
    $logTabs.TabPages.Add($tabProgram)

    $programInner         = New-Object System.Windows.Forms.Panel
    $programInner.Dock    = 'Fill'
    $programInner.BackColor = $clrBg
    $tabProgram.Controls.Add($programInner)

    $rtProgram            = [System.Windows.Forms.RichTextBox]::new()
    $rtProgram.Location   = [System.Drawing.Point]::new(0, 0)
    $rtProgram.Size       = [System.Drawing.Size]::new($defaultWidth, $bottomLogsHeight - 52)
    $rtProgram.Anchor     = 'Left,Right,Top,Bottom'
    $rtProgram.BackColor  = [System.Drawing.Color]::FromArgb(17,17,27)
    $rtProgram.ForeColor  = $clrText
    $rtProgram.Font       = $fontMono
    $rtProgram.ReadOnly   = $true
    $rtProgram.ScrollBars = 'Vertical'
    $rtProgram.WordWrap   = $false
    $programInner.Controls.Add($rtProgram)
    $script:_ProgramLogBox = $rtProgram

    $btnClearProg = _Button 'Clear' ($defaultWidth - 200) ($bottomLogsHeight - 50) 70 26 $clrMuted {
        if ($script:_ProgramLogBox) { $script:_ProgramLogBox.Clear() }
        if (-not [string]::IsNullOrEmpty($script:LogFilePath) -and (Test-Path $script:LogFilePath)) {
            $script:LogFilePos = (Get-Item $script:LogFilePath).Length
        }
    }
    $btnClearProg.Anchor = 'Right,Bottom'
    $programInner.Controls.Add($btnClearProg)

    # ── GAME LOG TABS  -  created/removed as servers start and stop ───────────
    # Key = game prefix (e.g. 'PZ'), Value = hashtable with TabPage + RTB + tail state
    $script:_GameLogTabs   = @{}   # prefix -> @{ Tab; RTB; Files=@{path->pos}; LogRoot; Strategy; LastFolder }
    $script:_GameLogTails  = @{}   # "prefix::filepath" -> byte position  (legacy compat key)
    $script:_ServerStartNotified = @{}  # prefix -> $true once "SERVER STARTED" seen
    $script:_PlayersCapture = @{}       # prefix -> @{ Active; Expected; Names; Started }

    # ── PER-GAME LOG RESOLVER ─────────────────────────────────────────────────
    # Given a profile hashtable, returns the list of absolute log file paths
    # to tail right now.  Handles all 4 LogStrategy values.
    function _ResolveGameLogFiles {
        param([hashtable]$Profile)
        $strategy = if ($Profile.LogStrategy) { $Profile.LogStrategy } else { 'SingleFile' }
        $files    = [System.Collections.Generic.List[string]]::new()

        switch ($strategy) {

            'PZSessionFolder' {
                # Force: ONLY read logs directly from the root Logs folder
                $root = _ExpandPathVars ([string]$Profile.ServerLogRoot)
                if ([string]::IsNullOrEmpty($root) -or -not (Test-Path $root)) { break }

                # Only pick the NEWEST file for each pattern in the ROOT folder
                $patterns = @(
                    'DebugLog-server.txt',
                    '*_DebugLog-server.txt',
                    '*_user.txt',
                    '*_chat.txt'
                )

                foreach ($pattern in $patterns) {
                    $match = Get-ChildItem -Path $root -Filter $pattern -File -ErrorAction SilentlyContinue |
                            Sort-Object LastWriteTime -Descending |
                            Select-Object -First 1
                    if ($match) { $files.Add($match.FullName) }
                }

                break
            }

            'ValheimUserFolder' {
                # ServerLogRoot = %AppData%\..\LocalLow\IronGate\Valheim
                $root = _ExpandPathVars ([string]$Profile.ServerLogRoot)
                if ([string]::IsNullOrEmpty($root)) {
                    $root = Join-Path $env:APPDATA '..\..\..\LocalLow\IronGate\Valheim'
                    $root = [System.IO.Path]::GetFullPath($root)
                }
                $logFile = Join-Path $root 'valheim_server.log'
                if (Test-Path $logFile) { $files.Add($logFile) }
            }

            'NewestFile' {
                # ServerLogRoot = folder, ServerLogFile = preferred name or *.log
                $root = _ExpandPathVars ([string]$Profile.ServerLogRoot)
                $folderPath = _ExpandPathVars ([string]$Profile.FolderPath)
                if ([string]::IsNullOrEmpty($root) -and -not [string]::IsNullOrEmpty($folderPath)) {
                    if (-not [string]::IsNullOrEmpty($Profile.ServerLogSubDir)) {
                        $root = Join-Path $folderPath $Profile.ServerLogSubDir
                    } else {
                        $root = $folderPath
                    }
                }
                if ([string]::IsNullOrEmpty($root) -or -not (Test-Path $root)) { break }
                # Try preferred filename first (supports wildcards)
                $preferred = $Profile.ServerLogFile
                if (-not [string]::IsNullOrEmpty($preferred) -and $preferred -ne '*.log') {
                    if ($preferred -match '[\*\?]') {
                        $match = Get-ChildItem -Path $root -Filter $preferred -ErrorAction SilentlyContinue |
                                 Sort-Object LastWriteTime -Descending | Select-Object -First 1
                        if ($match) { $files.Add($match.FullName); break }
                    } else {
                        $pf = Join-Path $root $preferred
                        if (Test-Path $pf) { $files.Add($pf); break }
                    }
                }
                # Fall back to newest *.log in the folder
                $newest = Get-ChildItem -Path $root -Filter '*.log' -ErrorAction SilentlyContinue |
                          Sort-Object LastWriteTime -Descending | Select-Object -First 1
                if ($newest) { $files.Add($newest.FullName) }
            }

            default {
                # SingleFile  -  use ServerLogPath directly
                $lp = _ExpandPathVars ([string]$Profile.ServerLogPath)
                if (-not [string]::IsNullOrEmpty($lp) -and (Test-Path $lp)) {
                    $files.Add($lp)
                }
            }
        }
        return $files
    }

    # ── ENSURE GAME LOG TAB  -  idempotent, called each timer tick ────────────
    function _EnsureGameLogTab {
        param([string]$Prefix, [string]$GameName)
        if ($script:_GameLogTabs.ContainsKey($Prefix)) { return }

        $tabPage           = New-Object System.Windows.Forms.TabPage
        $tabPage.Text      = $GameName
        $tabPage.BackColor = $clrBg
        $script:_LogTabControl.TabPages.Add($tabPage)

        # Inner panel to hold RTB + footer
        $inner            = New-Object System.Windows.Forms.Panel
        $inner.Dock       = 'Fill'
        $inner.BackColor  = $clrBg
        $tabPage.Controls.Add($inner)

        # --- LOG VIEWER (FULL WIDTH, FULL HEIGHT) ---
        $rtGame                  = New-Object System.Windows.Forms.RichTextBox
        $rtGame.Dock             = 'Fill'
        $rtGame.BackColor        = [System.Drawing.Color]::FromArgb(17,17,27)
        $rtGame.ForeColor        = $clrText
        $rtGame.Font             = $fontMono
        $rtGame.ReadOnly         = $true
        $rtGame.ScrollBars       = 'Vertical'
        $rtGame.WordWrap         = $false
        if ($Prefix -eq 'PZ') { $rtGame.Tag = @{ MatrixGreen = $true } }
        $inner.Controls.Add($rtGame)

        # --- FOOTER BAR (BOTTOM DOCKED) ---
        $footer            = New-Object System.Windows.Forms.Panel
        $footer.Dock       = 'Bottom'
        $footer.Height     = 28
        $footer.BackColor  = $clrBg
        $inner.Controls.Add($footer)

        # Source label
        $lblSrc            = _Label 'Waiting for log file...' 4 6 700 18
        $lblSrc.ForeColor  = $clrMuted
        $lblSrc.Font       = New-Object System.Drawing.Font('Consolas', 7)
        $lblSrc.Anchor     = 'Left,Bottom'
        $footer.Controls.Add($lblSrc)

        # Clear button (right aligned)
        $btnClearGame = _Button 'Clear' ($footer.Width - 70) 2 70 24 $clrMuted {
            if ($script:_GameLogTabs.ContainsKey($Prefix)) {
                $entry = $script:_GameLogTabs[$Prefix]
                if ($entry -and $entry.RTB) { $entry.RTB.Clear() }
                if ($entry) { $entry.Files = @{} }
            }
        }
        $btnClearGame.Anchor = 'Right,Top'
        $footer.Controls.Add($btnClearGame)

        # Register tab entry
        $script:_GameLogTabs[$Prefix] = @{
            Tab         = $tabPage
            RTB         = $rtGame
            LblSrc      = $lblSrc
            Files       = @{}
            LastSession = ''
        }
    }


    # ── REMOVE GAME LOG TAB  -  called when a server stops ───────────────────
    function _RemoveGameLogTab {
        param([string]$Prefix)
        if (-not $script:_GameLogTabs.ContainsKey($Prefix)) { return }
        $entry = $script:_GameLogTabs[$Prefix]
        $script:_LogTabControl.TabPages.Remove($entry.Tab)
        $script:_GameLogTabs.Remove($Prefix)
    }

    # =====================================================================
    # BACKGROUND METRICS RUNSPACE
    # =====================================================================
    $script:_cpuSmooth = 0.0
    $script:_ramSmooth = 0.0
    $script:_netSmooth = 0.0

    # Metrics are collected in a dedicated background runspace that loops every 2s
    # and writes results directly into SharedState. The GUI timer just reads them.
    # This avoids the BeginInvoke/IsCompleted race and the Get-Counter 1s block issue.
    $script:MetricsRunspace                = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $script:MetricsRunspace.ApartmentState = 'MTA'
    $script:MetricsRunspace.ThreadOptions  = 'ReuseThread'
    $script:MetricsRunspace.Open()
    $script:MetricsRunspace.SessionStateProxy.SetVariable('SharedState', $SharedState)

    $script:MetricsPS          = [System.Management.Automation.PowerShell]::Create()
    $script:MetricsPS.Runspace = $script:MetricsRunspace
    $script:MetricsPS.AddScript({
        Set-StrictMode -Off
        $ErrorActionPreference = 'SilentlyContinue'

        # Seed previous net counters for delta calculation
        $prevBytes = 0L
        $prevTime  = [datetime]::UtcNow

        while ($true) {
            try {
                Start-Sleep -Seconds 2

                # CPU - Get-Counter blocks for ~1s to get a proper sample; fine in background
                $cpu = (Get-Counter '\Processor(_Total)\% Processor Time' -ErrorAction SilentlyContinue).CounterSamples.CookedValue
                if ($null -eq $cpu) { $cpu = 0.0 }

                # RAM - use WMI to get physical memory usage percentage
                $os        = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue
                $ram       = if ($os -and $os.TotalVisibleMemorySize -gt 0) {
                    (($os.TotalVisibleMemorySize - $os.FreePhysicalMemory) / $os.TotalVisibleMemorySize) * 100
                } else { 0.0 }

                # NET - delta of total bytes (send + receive) across all adapters since last sample
                $nowTime   = [datetime]::UtcNow
                $elapsed   = ($nowTime - $prevTime).TotalSeconds
                if ($elapsed -lt 0.1) { $elapsed = 2.0 }

                $adapters  = Get-NetAdapterStatistics -ErrorAction SilentlyContinue
                $nowBytes  = if ($adapters) {
                    ($adapters | Measure-Object -Property ReceivedBytes    -Sum).Sum +
                    ($adapters | Measure-Object -Property SentBytes        -Sum).Sum
                } else { 0L }

                $deltaKBps = if ($prevBytes -gt 0 -and $nowBytes -ge $prevBytes) {
                    (($nowBytes - $prevBytes) / $elapsed) / 1KB
                } else { 0.0 }

                $prevBytes = $nowBytes
                $prevTime  = $nowTime

                $SharedState['_MetricCPU'] = [double]$cpu
                $SharedState['_MetricRAM'] = [double]$ram
                $SharedState['_MetricNET'] = [double]$deltaKBps
            } catch { }
        }
    }) | Out-Null
    $script:MetricsPS.BeginInvoke() | Out-Null

    # =====================================================================
    # BACKGROUND GAME LOG TAIL RUNSPACE
    # =====================================================================
    $script:LogTailRunspace                = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $script:LogTailRunspace.ApartmentState = 'MTA'
    $script:LogTailRunspace.ThreadOptions  = 'ReuseThread'
    $script:LogTailRunspace.Open()
    $script:LogTailRunspace.SessionStateProxy.SetVariable('SharedState', $SharedState)

    $script:LogTailPS          = [System.Management.Automation.PowerShell]::Create()
    $script:LogTailPS.Runspace = $script:LogTailRunspace
    $script:LogTailPS.AddScript({
        Set-StrictMode -Off
        $ErrorActionPreference = 'SilentlyContinue'

        $filePos   = @{}   # path -> int64 position
        $remainders = @{}  # path -> trailing partial line

        function _ExpandPathVars {
            param([string]$Path)
            if ([string]::IsNullOrWhiteSpace($Path)) { return $Path }
            return [Environment]::ExpandEnvironmentVariables($Path)
        }

        function _ResolveLogFiles {
            param($Profile)
            $strategy = if ($Profile.LogStrategy) { $Profile.LogStrategy } else { 'SingleFile' }
            $files    = [System.Collections.Generic.List[string]]::new()

            switch ($strategy) {
                'PZSessionFolder' {
                    $root = _ExpandPathVars ([string]$Profile.ServerLogRoot)
                    if ([string]::IsNullOrEmpty($root) -or -not (Test-Path $root)) { break }
                    $patterns = @(
                        'DebugLog-server.txt',
                        '*_DebugLog-server.txt',
                        '*_user.txt',
                        '*_chat.txt'
                    )
                    foreach ($pattern in $patterns) {
                        $match = Get-ChildItem -Path $root -Filter $pattern -File -ErrorAction SilentlyContinue |
                                Sort-Object LastWriteTime -Descending |
                                Select-Object -First 1
                        if ($match) { $files.Add($match.FullName) }
                    }
                    break
                }

                'ValheimUserFolder' {
                    $root = _ExpandPathVars ([string]$Profile.ServerLogRoot)
                    if ([string]::IsNullOrEmpty($root)) {
                        $root = Join-Path $env:APPDATA '..\..\..\LocalLow\IronGate\Valheim'
                        $root = [System.IO.Path]::GetFullPath($root)
                    }
                    $logFile = Join-Path $root 'valheim_server.log'
                    if (Test-Path $logFile) { $files.Add($logFile) }
                }

                'NewestFile' {
                    $root = _ExpandPathVars ([string]$Profile.ServerLogRoot)
                    $folderPath = _ExpandPathVars ([string]$Profile.FolderPath)
                    if ([string]::IsNullOrEmpty($root) -and -not [string]::IsNullOrEmpty($folderPath)) {
                        if (-not [string]::IsNullOrEmpty($Profile.ServerLogSubDir)) {
                            $root = Join-Path $folderPath $Profile.ServerLogSubDir
                        } else {
                            $root = $folderPath
                        }
                    }
                    if ([string]::IsNullOrEmpty($root) -or -not (Test-Path $root)) { break }
                    $preferred = $Profile.ServerLogFile
                    if (-not [string]::IsNullOrEmpty($preferred) -and $preferred -ne '*.log') {
                        if ($preferred -match '[\*\?]') {
                            $match = Get-ChildItem -Path $root -Filter $preferred -ErrorAction SilentlyContinue |
                                     Sort-Object LastWriteTime -Descending | Select-Object -First 1
                            if ($match) { $files.Add($match.FullName); break }
                        } else {
                            $pf = Join-Path $root $preferred
                            if (Test-Path $pf) { $files.Add($pf); break }
                        }
                    }
                    $newest = Get-ChildItem -Path $root -Filter '*.log' -ErrorAction SilentlyContinue |
                              Sort-Object LastWriteTime -Descending | Select-Object -First 1
                    if ($newest) { $files.Add($newest.FullName) }
                }

                default {
                    $lp = _ExpandPathVars ([string]$Profile.ServerLogPath)
                    if (-not [string]::IsNullOrEmpty($lp) -and (Test-Path $lp)) {
                        $files.Add($lp)
                    }
                }
            }
            return $files
        }

        function _ReadNewLines {
            param([string]$Path)
            if (-not (Test-Path $Path)) { return @() }

            $len = (Get-Item $Path).Length
            if (-not $filePos.ContainsKey($Path)) {
                # Skip existing content on first open
                $filePos[$Path] = $len
                return @()
            }

            if ($len -lt $filePos[$Path]) { $filePos[$Path] = 0 }
            if ($len -eq $filePos[$Path]) { return @() }

            $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
            try {
                $fs.Seek($filePos[$Path], [System.IO.SeekOrigin]::Begin) | Out-Null
                $sr = New-Object System.IO.StreamReader($fs, [System.Text.Encoding]::UTF8, $true)
                $text = $sr.ReadToEnd()
                $filePos[$Path] = $fs.Position
                $sr.Close()
            } finally {
                $fs.Close()
            }

            if ([string]::IsNullOrEmpty($text)) { return @() }

            $prefix = if ($remainders.ContainsKey($Path)) { $remainders[$Path] } else { '' }
            $text = $prefix + $text
            $lines = $text -split "`r?`n"

            # If the last line is partial (no trailing newline), keep it as remainder
            if (-not ($text.EndsWith("`n") -or $text.EndsWith("`r"))) {
                $remainders[$Path] = $lines[-1]
                $lines = $lines[0..($lines.Count - 2)]
            } else {
                $remainders[$Path] = ''
            }

            return $lines
        }

        while ($true) {
            try {
                Start-Sleep -Milliseconds 800

                if ($null -eq $SharedState -or $null -eq $SharedState.Profiles) { continue }
                if (-not $SharedState.ContainsKey('GameLogQueue')) { continue }

                foreach ($pfx in @($SharedState.RunningServers.Keys)) {
                    $profile = $SharedState.Profiles[$pfx]
                    if ($null -eq $profile) { continue }

                    $files = _ResolveLogFiles -Profile $profile
                    foreach ($path in $files) {
                        $lines = _ReadNewLines -Path $path
                        if ($lines.Count -eq 0) { continue }

                        $maxLines = 200
                        if ($lines.Count -gt $maxLines) {
                            $skip = $lines.Count - $maxLines
                            $lines = $lines[-$maxLines..-1]
                            $SharedState.GameLogQueue.Enqueue([pscustomobject]@{ Prefix=$pfx; Line="... skipped $skip lines ..."; Path=$path })
                        }

                        foreach ($line in $lines) {
                            if ($line -eq '') { continue }
                            $SharedState.GameLogQueue.Enqueue([pscustomobject]@{ Prefix=$pfx; Line=$line; Path=$path })
                        }
                    }
                }
            } catch { }
        }
    }) | Out-Null
    $script:LogTailPS.BeginInvoke() | Out-Null

    # =====================================================================
    # LOCAL HELPER FUNCTIONS  (inside Start-GUI so they close over locals)
    # =====================================================================
    function _Smooth {
        param([double]$old, [double]$new, [double]$factor = 0.2)
        return ($old + ($new - $old) * $factor)
    }

    function _SetMetricColor {
        param([System.Windows.Forms.Label]$label, [double]$value)
        if     ($value -lt 50) { $label.ForeColor = $clrGreen  }
        elseif ($value -lt 80) { $label.ForeColor = $clrYellow }
        else                   { $label.ForeColor = $clrRed    }
    }

    function _WriteProgramLog {
        param([string]$Line)
        $rt = $script:_ProgramLogBox
        if ($null -eq $rt) { return }
        _AppendLog $rt $Line (_LogColour $Line)
    }

    function _WriteDiscordLog {
        param([string]$Line)
        $rt = $script:_DiscordLogBox
        if ($null -eq $rt) { return }
        _AppendLog $rt $Line $clrAccent
    }

    function _ExpandPathVars {
        param([string]$Path)
        if ([string]::IsNullOrWhiteSpace($Path)) { return $Path }
        return [Environment]::ExpandEnvironmentVariables($Path)
    }

    function _GetConfigRootsForProfile {
        param([hashtable]$Profile)

        $roots = @()
        if ($null -eq $Profile) { return @() }

        if (($Profile.Keys -contains 'ConfigRoot') -and $Profile.ConfigRoot) {
            $roots += _ExpandPathVars ([string]$Profile.ConfigRoot)
        }
        if (($Profile.Keys -contains 'ConfigRoots') -and $Profile.ConfigRoots) {
            $roots += @($Profile.ConfigRoots | ForEach-Object { _ExpandPathVars ([string]$_) })
        }

        if ($roots.Count -eq 0) {
            $gn  = "$($Profile.GameName)".ToLower()
            $pfx = "$($Profile.Prefix)".ToUpper()
            if ($pfx -eq 'PZ' -or $gn -match 'project\s*zomboid' -or $gn -match 'zomboid') {
                $roots += (Join-Path $env:USERPROFILE 'Zomboid\Server')
            }
        }

        # Ensure we never enumerate a single string as characters
        $roots = @($roots)
        $roots = $roots | Where-Object { $_ -and (Test-Path $_) } | Select-Object -Unique
        return ,$roots
    }

    function _OpenConfigEditor {
        param([hashtable]$Profile)

        $roots = _GetConfigRootsForProfile -Profile $Profile
        # Defensive: ensure array semantics even if a string sneaks through
        $roots = @($roots)
        if ($roots.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                'No config folder was found for this server.',
                'Config Editor','OK','Information') | Out-Null
            return
        }
        if ($script:SharedState -and $script:SharedState.Settings -and $script:SharedState.Settings.EnableDebugLogging) {
            if ($script:SharedState.ContainsKey('LogQueue')) {
                $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][DEBUG][GUI] Config roots: $($roots -join '; ')")
            }
        }

        $form                 = New-Object System.Windows.Forms.Form
        $form.Text            = "Config Editor - $($Profile.GameName)"
        $form.Size            = [System.Drawing.Size]::new(900, 620)
        $form.MinimumSize     = [System.Drawing.Size]::new(760, 520)
        $form.StartPosition   = 'CenterParent'
        $form.BackColor       = $clrBg
        $form.FormBorderStyle = 'Sizable'

        $header = _Label "Config Editor - $($Profile.GameName)" 10 10 500 22 $fontBold
        $form.Controls.Add($header)

        $lblRoot = _Label "Root: $($roots[0])" 10 34 840 20 $fontLabel
        $lblRoot.Anchor = 'Top,Left,Right'
        $form.Controls.Add($lblRoot)

        $combo = $null
        if ($roots.Count -gt 1) {
            $combo = New-Object System.Windows.Forms.ComboBox
            $combo.Location = [System.Drawing.Point]::new(60, 32)
            $combo.Size     = [System.Drawing.Size]::new(680, 24)
            $combo.DropDownStyle = 'DropDownList'
            foreach ($r in $roots) { [void]$combo.Items.Add($r) }
            $combo.SelectedIndex = 0
            $form.Controls.Add($combo)
            $lblRoot.Visible = $false
        }

        $list = New-Object System.Windows.Forms.ListBox
        $list.Location  = [System.Drawing.Point]::new(10, 60)
        $list.Size      = [System.Drawing.Size]::new(250, 470)
        $list.Anchor    = 'Top,Left,Bottom'
        $list.Font      = $fontMono
        $list.BackColor = [System.Drawing.Color]::FromArgb(30,30,40)
        $list.ForeColor = $clrText
        $form.Controls.Add($list)

        $editor = New-Object System.Windows.Forms.TextBox
        $editor.Location  = [System.Drawing.Point]::new(270, 60)
        $editor.Size      = [System.Drawing.Size]::new(600, 470)
        $editor.Anchor    = 'Top,Left,Right,Bottom'
        $editor.Multiline = $true
        $editor.ScrollBars = 'Both'
        $editor.WordWrap  = $false
        $editor.AcceptsTab = $true
        $editor.Font      = $fontMono
        $editor.BackColor = [System.Drawing.Color]::FromArgb(30,30,40)
        $editor.ForeColor = $clrText
        $form.Controls.Add($editor)

        $btnSave = _Button 'Save File' 270 540 100 30 $clrGreen $null
        $btnSave.Anchor = 'Left,Bottom'
        $form.Controls.Add($btnSave)

        $btnRefresh = _Button 'Refresh' 380 540 100 30 $clrPanel $null
        $btnRefresh.Anchor = 'Left,Bottom'
        $form.Controls.Add($btnRefresh)

        $lblStatus = _Label '' 500 545 360 20 $fontLabel
        $lblStatus.Anchor = 'Left,Right,Bottom'
        $form.Controls.Add($lblStatus)

        $allowedExt = @('.ini','.txt','.cfg','.json','.xml','.yml','.yaml','.properties','.conf','.lua')

        $state = [ordered]@{
            Root        = $roots[0]
            Roots       = $roots
            AllowedExt  = $allowedExt
            CurrentFile = ''
            List        = $list
            Editor      = $editor
            Status      = $lblStatus
            RootLabel   = $lblRoot
        }

        $state.Refresh = {
            param($st)
            $st.List.Items.Clear()
            $root = $st.Root
            if (-not (Test-Path $root)) {
                $st.Status.Text = "Folder missing: $root"
                return
            }
            $files = Get-ChildItem -Path $root -File -ErrorAction SilentlyContinue |
                     Where-Object { $st.AllowedExt -contains $_.Extension.ToLower() } |
                     Sort-Object -Property Name
            foreach ($f in $files) { [void]$st.List.Items.Add($f.Name) }
            $st.Status.Text = "$($files.Count) file(s)"
            if ($script:SharedState -and $script:SharedState.Settings -and $script:SharedState.Settings.EnableDebugLogging) {
                if ($script:SharedState.ContainsKey('LogQueue')) {
                    $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][DEBUG][GUI] Config file count: $($files.Count) in $root")
                }
            }
        }

        $state.LoadFile = {
            param($st, $fileName)
            if ([string]::IsNullOrWhiteSpace($fileName)) { return }
            $path = Join-Path $st.Root $fileName
            if (-not (Test-Path -LiteralPath $path)) {
                [System.Windows.Forms.MessageBox]::Show(
                    "File not found: $path",'Config Editor','OK','Error') | Out-Null
                return
            }
            $info = Get-Item -LiteralPath $path -ErrorAction SilentlyContinue
            if ($info -and $info.Length -gt 1048576) {
                [System.Windows.Forms.MessageBox]::Show(
                    'File is larger than 1 MB. Please edit it with an external editor.',
                    'Config Editor','OK','Information') | Out-Null
                return
            }
            try {
                $st.Editor.Text = Get-Content -LiteralPath $path -Raw -ErrorAction Stop
                $st.CurrentFile = $path
                $st.Status.Text = "Editing: $fileName"
            } catch {
                [System.Windows.Forms.MessageBox]::Show(
                    "Failed to open file: $_",'Config Editor','OK','Error') | Out-Null
            }
        }

        $form.Tag = $state

        $btnRefresh.Add_Click({
            $st = $this.FindForm().Tag
            if ($st) { & $st.Refresh $st }
        })

        $btnSave.Add_Click({
            $st = $this.FindForm().Tag
            if (-not $st -or [string]::IsNullOrWhiteSpace($st.CurrentFile)) {
                [System.Windows.Forms.MessageBox]::Show(
                    'Select a file to save first.','Config Editor','OK','Information') | Out-Null
                return
            }
            try {
                Set-Content -LiteralPath $st.CurrentFile -Value $st.Editor.Text -Encoding UTF8 -Force
                $st.Status.Text = "Saved: $(Split-Path $st.CurrentFile -Leaf)"
                if ($script:SharedState -and $script:SharedState.Settings -and $script:SharedState.Settings.EnableDebugLogging) {
                    if ($script:SharedState.ContainsKey('LogQueue')) {
                        $len = 0
                        try { $len = $st.Editor.Text.Length } catch { $len = 0 }
                        $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][DEBUG][GUI] Config saved: $($st.CurrentFile) len=$len")
                    }
                }
            } catch {
                [System.Windows.Forms.MessageBox]::Show(
                    "Failed to save file: $_",'Config Editor','OK','Error') | Out-Null
            }
        })

        $list.Add_SelectedIndexChanged({
            $st = $this.FindForm().Tag
            if ($st -and $this.SelectedItem) {
                & $st.LoadFile $st $this.SelectedItem.ToString()
            }
        })

        if ($combo) {
            $combo.Add_SelectedIndexChanged({
                $st = $this.FindForm().Tag
                if (-not $st) { return }
                $st.Root = $this.SelectedItem.ToString()
                if ($st.RootLabel) { $st.RootLabel.Text = "Root: $($st.Root)" }
                $st.CurrentFile = ''
                $st.Editor.Text = ''
                & $st.Refresh $st
            })
        }

        & $state.Refresh $state
        $form.ShowDialog() | Out-Null
    }

    # =====================================================================
    # PROFILE EDITOR  (RIGHT COLUMN)
    # =====================================================================
    function _BuildProfileEditor {
        param([object]$Profile)

        $panel = $script:_ProfileEditorPanel
        if ($null -eq $panel) { return }
        $panel.Controls.Clear()

        if ($null -eq $Profile) {
            $panel.Controls.Add((_Label 'Select a server to edit its profile.' 10 10 400 20))
            return
        }

        # Store in script scope so the Save button closure has a stable reference.
        # PowerShell 5.1 closures do not reliably capture function-param variables
        # after the function frame is gone - script scope persists for the session.
        $script:_EditingProfile = $Profile
        $script:_EditingPrefix  = "$($Profile['Prefix'])".ToUpper()

        $scroll            = New-Object System.Windows.Forms.Panel
        $scroll.Location   = [System.Drawing.Point]::new(0, 0)
        $scroll.Size       = [System.Drawing.Size]::new($panel.Width, $panel.Height)
        $scroll.AutoScroll = $true
        $scroll.BackColor  = $clrPanel
        $scroll.Anchor     = 'Top,Left,Right,Bottom'
        $panel.Controls.Add($scroll)

        $lw  = 170
        $tw  = [Math]::Max(300, $scroll.ClientSize.Width - ($lw + 40))
        $th  = 24
        $gap = 32

        function _FormatKeyLabel([string]$key) {
            return $key -replace '([a-z])([A-Z])', '$1 $2'
        }

        function Add-DynamicField ([string]$key, [object]$value) {
            $label = _FormatKeyLabel $key
            $scroll.Controls.Add((_Label $label 10 $script:y $lw 20))

            # Complex values get a JSON text area for full visibility/editing.
            $isDict = $value -is [System.Collections.IDictionary]
            $isList = ($value -is [System.Collections.IEnumerable] -and -not ($value -is [string]))

            if ($value -is [bool]) {
                $chk           = New-Object System.Windows.Forms.CheckBox
                $chk.Location  = [System.Drawing.Point]::new($lw + 20, $script:y)
                $chk.Size      = [System.Drawing.Size]::new(200, 20)
                $chk.Text      = if ($key -eq 'RestPollOnlyWhenRunning') { 'Only when running' } else { 'Enabled' }
                $chk.ForeColor = $clrText
                $chk.BackColor = [System.Drawing.Color]::Transparent
                $chk.Checked   = ($value -eq $true)
                $scroll.Controls.Add($chk)
                $script:_ProfileFields[$key] = @{ Control = $chk; Kind = 'bool' }
                $script:y += $gap
                return
            }

            if ($isDict -or $isList) {
                $json = ''
                try { $json = $value | ConvertTo-Json -Depth 6 } catch { $json = '' }
                $tb = New-Object System.Windows.Forms.TextBox
                $tb.Location    = [System.Drawing.Point]::new($lw + 20, $script:y)
                $tb.Size        = [System.Drawing.Size]::new($tw, 80)
                $tb.Multiline   = $true
                $tb.ScrollBars  = 'Both'
                $tb.WordWrap    = $false
                $tb.Text        = $json
                $tb.BackColor   = [System.Drawing.Color]::FromArgb(30,30,40)
                $tb.ForeColor   = $clrText
                $tb.BorderStyle = 'FixedSingle'
                $tb.Font        = $fontMono
                $tb.Anchor      = 'Top,Left,Right'
                $scroll.Controls.Add($tb)
                $script:_ProfileFields[$key] = @{ Control = $tb; Kind = 'json' }
                $script:y += 92
                return
            }

            $tb        = _TextBox ($lw + 20) $script:y $tw $th ([string]$value) $false
            $tb.Anchor = 'Top,Left,Right'
            $scroll.Controls.Add($tb)
            $script:_ProfileFields[$key] = @{ Control = $tb; Kind = 'text' }
            $script:y += $gap
        }

        $script:_ProfileFields = @{}
        $script:_ProfileSeparators = @()
        $script:y = 10

        function Add-SectionHeader([string]$title) {
            $lblSec = _Label $title 10 $script:y 300 20 $fontBold
            $lblSec.ForeColor = $clrAccent
            $scroll.Controls.Add($lblSec)
            $script:y += 20
            $sep = New-Object System.Windows.Forms.Panel
            $sep.Location = [System.Drawing.Point]::new(10, $script:y)
            $sep.Size     = [System.Drawing.Size]::new([Math]::Max(100, $scroll.ClientSize.Width - 20), 1)
            $sep.BackColor = [System.Drawing.Color]::FromArgb(70,70,90)
            $sep.Anchor   = 'Top,Left,Right'
            $scroll.Controls.Add($sep)
            $script:_ProfileSeparators += $sep
            $script:y += 12
        }

        $sections = [ordered]@{
            'Basics' = @('GameName','Prefix','ProcessName','Executable','FolderPath')
            'Launch' = @('LaunchArgs','MaxRamGB','MinRamGB','AOTCache')
            'Logs'   = @('LogStrategy','ServerLogRoot','ServerLogSubDir','ServerLogFile','ServerLogPath','ServerLogNote')
            'REST'   = @('RestEnabled','RestHost','RestPort','RestPassword','RestProtocol','RestPollOnlyWhenRunning')
            'RCON'   = @('RconHost','RconPort','RconPassword')
            'Restart/Safety' = @('EnableAutoRestart','RestartDelaySeconds','MaxRestartsPerHour','SaveMethod','SaveWaitSeconds','StopMethod')
            'Config' = @('ConfigRoot','ConfigRoots')
            'Commands' = @('Commands','ExtraCommands','StdinSaveCommand','StdinStopCommand','ExeHints')
            'Misc'   = @('AssetFile','BackupDir')
        }

        $used = @{}
        foreach ($sec in $sections.Keys) {
            $keys = $sections[$sec]
            $hasAny = $false
            foreach ($k in $keys) {
                if ($Profile.Keys -contains $k) { $hasAny = $true; break }
            }
            if (-not $hasAny) { continue }
            Add-SectionHeader -title $sec
            foreach ($k in $keys) {
                if ($Profile.Keys -contains $k) {
                    Add-DynamicField -key $k -value $Profile[$k]
                    $used[$k] = $true
                }
            }
        }

        # Any remaining keys not in our sections
        $remaining = @()
        foreach ($k in $Profile.Keys) {
            if (-not $used.ContainsKey($k)) { $remaining += $k }
        }
        if ($remaining.Count -gt 0) {
            Add-SectionHeader -title 'Other'
            foreach ($k in ($remaining | Sort-Object)) {
                Add-DynamicField -key $k -value $Profile[$k]
            }
        }

        $btnSave = _Button 'Save Changes' 10 ($script:y + 10) 150 32 $clrGreen {
            # Use script-scope reference - the param $Profile is gone by click time
            $prof = $script:_EditingProfile
            if ($null -eq $prof -or -not ($prof -is [System.Collections.IDictionary])) {
                [System.Windows.Forms.MessageBox]::Show('No profile loaded for editing.','Error','OK','Error') | Out-Null
                return
            }

            $before = @{}
            foreach ($k in $prof.Keys) { $before[$k] = $prof[$k] }

            foreach ($key in $script:_ProfileFields.Keys) {
                $entry = $script:_ProfileFields[$key]
                $ctrl  = $entry.Control
                $kind  = $entry.Kind

                if ($kind -eq 'bool') {
                    $prof[$key] = $ctrl.Checked
                    continue
                }

                if ($kind -eq 'json') {
                    $raw = $ctrl.Text.Trim()
                    if ($raw -eq '') {
                        $prof[$key] = @{}
                        continue
                    }
                    try {
                        $prof[$key] = $raw | ConvertFrom-Json -ErrorAction Stop
                    } catch {
                        [System.Windows.Forms.MessageBox]::Show("Invalid JSON for '$key'. Fix it before saving.",'JSON Error','OK','Error') | Out-Null
                        return
                    }
                    continue
                }

                $val = $ctrl.Text.Trim()
                # Keep numeric-looking values as ints when possible.
                $n = 0
                if ([int]::TryParse($val, [ref]$n) -and $val -match '^\d+$') {
                    $prof[$key] = $n
                } else {
                    $prof[$key] = $val
                }
            }

            # Handle prefix rename
            $newPfx = "$($prof['Prefix'])".ToUpper()
            $prof['Prefix'] = $newPfx
            if ($newPfx -ne $script:_EditingPrefix -and $script:_EditingPrefix) {
                $script:SharedState.Profiles.Remove($script:_EditingPrefix)
                $script:SharedState.Profiles[$newPfx] = $prof
                $script:_EditingPrefix = $newPfx
            } else {
                $script:SharedState.Profiles[$newPfx] = $prof
            }

            try {
                if ($script:SharedState -and $script:SharedState.Settings -and $script:SharedState.Settings.EnableDebugLogging) {
                    if ($script:SharedState.ContainsKey('LogQueue')) {
                        function _TrimDbg([string]$s) {
                            if ($null -eq $s) { return '<null>' }
                            if ($s.Length -gt 120) { return $s.Substring(0,120) + '...' }
                            return $s
                        }
                        function _DbgVal($v) {
                            if ($null -eq $v) { return '<null>' }
                            if ($v -is [string] -or $v -is [int] -or $v -is [bool] -or $v -is [double]) {
                                return (_TrimDbg "$v")
                            }
                            try {
                                return (_TrimDbg ($v | ConvertTo-Json -Depth 6 -Compress))
                            } catch {
                                return (_TrimDbg "$v")
                            }
                        }

                        $sensitive = '(?i)password|token|secret|webhook|apikey|apiKey|restpassword'
                        $changes = @()
                        foreach ($k in $prof.Keys) {
                            $old = if ($before.ContainsKey($k)) { $before[$k] } else { $null }
                            $new = $prof[$k]
                            $oldStr = _DbgVal $old
                            $newStr = _DbgVal $new
                            if ($oldStr -ne $newStr) {
                                if ($k -match $sensitive) {
                                    $changes += "$k=<redacted>"
                                } else {
                                    $changes += "${k}: '$oldStr' -> '$newStr'"
                                }
                            }
                        }

                        if ($changes.Count -eq 0) {
                            $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][DEBUG][GUI] Profile save: no changes detected for $($prof.GameName)")
                        } else {
                            $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][DEBUG][GUI] Profile save: $($changes.Count) change(s) for $($prof.GameName): $($changes -join '; ')")
                        }
                    }
                }

                $pmPath = Join-Path $script:ModuleRoot 'ProfileManager.psm1'
                Import-Module $pmPath -Force
                Save-GameProfile -Profile $prof -ProfilesDir $script:ProfilesDir | Out-Null
                _BuildProfilesList
                _BuildServerDashboard
                [System.Windows.Forms.MessageBox]::Show('Profile saved successfully.','Saved','OK','Information') | Out-Null
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Failed to save profile:`n$_",'Error','OK','Error') | Out-Null
            }
        }
        $scroll.Controls.Add($btnSave)

        $btnRestart = _Button 'Restart Server' 180 ($script:y + 10) 150 32 $clrAccent {
            try {
                $smPath = Join-Path $script:ModuleRoot 'ServerManager.psm1'
                Import-Module $smPath -Force
                Initialize-ServerManager -SharedState $script:SharedState
                Restart-GameServer -Prefix $Profile.Prefix | Out-Null
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Restart failed: $_",'Error','OK','Error') | Out-Null
            }
        }
        $scroll.Controls.Add($btnRestart)

        $btnStop = _Button 'Stop Server' 340 ($script:y + 10) 150 32 $clrRed {
            try {
                $smPath = Join-Path $script:ModuleRoot 'ServerManager.psm1'
                Import-Module $smPath -Force
                Initialize-ServerManager -SharedState $script:SharedState
                Invoke-SafeShutdown -Prefix $Profile.Prefix | Out-Null
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Stop failed: $_",'Error','OK','Error') | Out-Null
            }
        }
        $scroll.Controls.Add($btnStop)
    }

    # =====================================================================
    # PROFILES LIST  (LEFT COLUMN)
    # =====================================================================
    function _BuildProfilesList {
        $panel = $script:_ProfilesPanel
        if ($null -eq $panel) { return }
        $panel.Controls.Clear()

        $listPanel = New-Object System.Windows.Forms.Panel
        $listPanel.Location   = [System.Drawing.Point]::new(10, 10)
        $listPanel.Size       = [System.Drawing.Size]::new($panel.Width - 20, $panel.Height - 70)
        $listPanel.Anchor     = 'Top,Left,Right,Bottom'
        $listPanel.BackColor  = [System.Drawing.Color]::FromArgb(40,40,55)
        $listPanel.AutoScroll = $true
        $panel.Controls.Add($listPanel)
        $script:_ProfilesListPanel = $listPanel

        $ss = $script:SharedState
        $rowH = 30
        $y = 0
        if ($ss -and $ss.Profiles -and $ss.Profiles.Count -gt 0) {
            foreach ($pfx in ($ss.Profiles.Keys | Sort-Object)) {
                $gn = $ss.Profiles[$pfx].GameName
                if (-not $gn) { $gn = $pfx }

                $row = New-Object System.Windows.Forms.Panel
                $row.Location = [System.Drawing.Point]::new(0, $y)
                $row.Size     = [System.Drawing.Size]::new($listPanel.Width - 2, $rowH)
                $row.Anchor   = 'Top,Left,Right'
                $row.Tag      = $pfx
                $row.Cursor   = 'Hand'
                $row.BackColor = if ($script:_SelectedProfilePrefix -eq $pfx) {
                    [System.Drawing.Color]::FromArgb(65,70,95)
                } else {
                    [System.Drawing.Color]::FromArgb(40,40,55)
                }

                $lbl = _Label "[$pfx] $gn" 8 6 ($row.Width - 16) 18
                $lbl.ForeColor = $clrText
                $lbl.Cursor = 'Hand'
                $lbl.Anchor = 'Left,Right,Top'
                $row.Controls.Add($lbl)

                $sep = New-Object System.Windows.Forms.Panel
                $sep.Location = [System.Drawing.Point]::new(0, $rowH - 1)
                $sep.Size     = [System.Drawing.Size]::new($row.Width, 1)
                $sep.Anchor   = 'Left,Right,Bottom'
                $sep.BackColor = [System.Drawing.Color]::FromArgb(70,70,90)
                $row.Controls.Add($sep)

                $row.Add_Click({
                    $p = $this.Tag
                    $script:_SelectedProfilePrefix = $p
                    $prof = $script:SharedState.Profiles[$p]
                    if ($prof) { _BuildProfileEditor -Profile $prof }
                    _BuildProfilesList
                })
                $lbl.Add_Click({
                    $p = $this.Parent.Tag
                    $script:_SelectedProfilePrefix = $p
                    $prof = $script:SharedState.Profiles[$p]
                    if ($prof) { _BuildProfileEditor -Profile $prof }
                    _BuildProfilesList
                })

                $listPanel.Controls.Add($row)
                $y += $rowH
            }
        } else {
            $panel.Controls.Add((_Label 'No profiles found. Use + Add Game to get started.' 10 10 500 20))
        }

        $btnAdd = _Button '+ Add Game' 10 ($panel.Height - 70) 110 30 $clrGreen {
            $dlg             = New-Object System.Windows.Forms.FolderBrowserDialog
            $dlg.Description = 'Select your game server folder'
            if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

            $folder   = $dlg.SelectedPath
            if ([string]::IsNullOrWhiteSpace($folder)) { return }
            $gameName = [System.IO.Path]::GetFileName($folder)

            $nameForm                  = New-Object System.Windows.Forms.Form
            $nameForm.Text             = 'Game Name'
            $nameForm.Size             = [System.Drawing.Size]::new(380, 170)
            $nameForm.BackColor        = $clrBg
            $nameForm.StartPosition    = 'CenterParent'
            $nameForm.FormBorderStyle  = 'FixedDialog'
            $nameForm.MaximizeBox      = $false
            $nameForm.MinimizeBox      = $false
            $nameForm.Controls.Add((_Label 'Enter a display name:' 10 10 340 22))
            $tbName = _TextBox 10 36 340 24 $gameName
            $nameForm.Controls.Add($tbName)
            $nameForm.Controls.Add((_Button 'Create' 10 74 140 30 $clrGreen {
                $script:_newGameName       = $tbName.Text.Trim()
                $nameForm.DialogResult     = [System.Windows.Forms.DialogResult]::OK
                $nameForm.Close()
            }))
            $nameForm.Controls.Add((_Button 'Cancel' 160 74 80 30 $clrMuted {
                $nameForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                $nameForm.Close()
            }))
            if ($nameForm.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

            $finalName = if ($script:_newGameName) { $script:_newGameName } else { $gameName }
            $script:_newGameName = $null

            # Optional: allow user to select a config folder (for games that store configs outside install dir)
            $configFolder = ''
            $cfgPrompt = [System.Windows.Forms.MessageBox]::Show(
                'Would you like to select a config folder? (Optional)',
                'Config Folder','YesNo','Question')
            if ($cfgPrompt -eq [System.Windows.Forms.DialogResult]::Yes) {
                $cfgDlg             = New-Object System.Windows.Forms.FolderBrowserDialog
                $cfgDlg.Description = 'Select your server config folder (optional)'
                if ($cfgDlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                    $configFolder = $cfgDlg.SelectedPath
                }
            }

            try {
                $pmPath = Join-Path $script:ModuleRoot 'ProfileManager.psm1'
                Import-Module $pmPath -Force
                $newProfile = New-GameProfile -FolderPath $folder -GameName $finalName
                if (-not [string]::IsNullOrWhiteSpace($configFolder)) {
                    $newProfile.ConfigRoot = $configFolder
                }
                $outPath    = Save-GameProfile -Profile $newProfile -ProfilesDir $script:ProfilesDir
                $pfx        = $newProfile.Prefix.ToUpper()
                $script:SharedState.Profiles[$pfx] = $newProfile
                _BuildProfilesList
                _BuildServerDashboard
                [System.Windows.Forms.MessageBox]::Show(
                    "Profile created:`nGame:   $finalName`nPrefix: $pfx`nFile:   $outPath`n`nUse !${pfx}start in Discord.",
                    'Created','OK','Information') | Out-Null
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Failed to create profile:`n$_",'Error','OK','Error') | Out-Null
            }
        }
        $btnAdd.Anchor = 'Left,Bottom'
        $panel.Controls.Add($btnAdd)

        $btnRemove = _Button 'Remove' ($panel.Width - 120) ($panel.Height - 70) 110 30 $clrRed {
            $pfx = $script:_SelectedProfilePrefix
            if (-not $pfx) {
                [System.Windows.Forms.MessageBox]::Show('Select a profile first.','Nothing Selected','OK','Information') | Out-Null
                return
            }
            if ([System.Windows.Forms.MessageBox]::Show(
                    "Remove the [$pfx] profile? This cannot be undone.",
                    'Confirm Remove','YesNo','Warning') -ne [System.Windows.Forms.DialogResult]::Yes) { return }

            try {
                $gn       = $script:SharedState.Profiles[$pfx].GameName
                $safe     = ($gn -replace '[\/:*?"<>|]','_') -replace '\s+','_'
                $jsonPath = Join-Path $script:ProfilesDir "$safe.json"
                if (Test-Path $jsonPath) { Remove-Item $jsonPath -Force }
            } catch {}

            $script:SharedState.Profiles.Remove($pfx)
            if ($script:_SelectedProfilePrefix -eq $pfx) { $script:_SelectedProfilePrefix = $null }
            _BuildProfilesList
            _BuildServerDashboard
            _BuildProfileEditor $null
        }
        $btnRemove.Anchor = 'Right,Bottom'
        $panel.Controls.Add($btnRemove)
    }

    # =====================================================================
    # SERVER DASHBOARD  (CENTER COLUMN)
    # =====================================================================
    function _BuildServerDashboard {
        $panel = $script:_ServerDashboardPanel
        if ($null -eq $panel) { return }
        $panel.Controls.Clear()

        $ss = $script:SharedState
        if ($null -eq $ss -or -not $ss.Profiles -or $ss.Profiles.Count -eq 0) {
            $panel.Controls.Add((_Label 'No profiles found. Use + Add Game to get started.' 10 10 500 20))
            return
        }

        $y = 10
        foreach ($pfx in ($ss.Profiles.Keys | Sort-Object)) {
            $profile = $ss.Profiles[$pfx]

            $running = $false
            $entry   = $null
            if ($ss.RunningServers -and $ss.RunningServers.ContainsKey($pfx)) {
                $entry = $ss.RunningServers[$pfx]
                $proc  = try { Get-Process -Id $entry.Pid -ErrorAction Stop } catch { $null }
                if ($null -ne $proc -and -not $proc.HasExited) {
                    $running = $true
                } else {
                    # Launcher PID may have exited (e.g. cmd.exe /c bat) - check by ProcessName
                    $procName = $profile.ProcessName
                    if ($procName) {
                        $gameProc = Get-Process -Name $procName -ErrorAction SilentlyContinue |
                                    Select-Object -First 1
                        if ($null -ne $gameProc -and -not $gameProc.HasExited) {
                            $running = $true
                        }
                    }
                }
            }

            $card        = _Panel 10 $y ($panel.Width - 20) 110 ([System.Drawing.Color]::FromArgb(49,50,68))
            $card.Anchor = 'Top,Left,Right'

            $lblName = _Label "$($profile.GameName) [$pfx]" 10 8 300 22 $fontBold
            $card.Controls.Add($lblName)

            $statusText  = if ($running) { 'ONLINE'  } else { 'OFFLINE' }
            $statusColor = if ($running) { $clrGreen } else { $clrRed   }
            $lblStatus   = _Label $statusText ($card.Width - 120) 8 80 22 $fontBold
            $lblStatus.Anchor   = 'Top,Right'
            $lblStatus.ForeColor = $statusColor
            $card.Controls.Add($lblStatus)

            if ($running -and $entry) {
                $up = [Math]::Round(((Get-Date) - $entry.StartTime).TotalMinutes, 1)
                $card.Controls.Add((_Label "PID $($entry.Pid) | Uptime: ${up} min" 10 35 260 18))
            } else {
                $card.Controls.Add((_Label 'Server is not running' 10 35 260 18))
            }

            $chk           = [System.Windows.Forms.CheckBox]::new()
            $chk.Text      = 'Auto-Restart'
            $chk.Location  = [System.Drawing.Point]::new(10, 60)
            $chk.Size      = [System.Drawing.Size]::new(120, 20)
            $chk.ForeColor = $clrText
            $chk.BackColor = [System.Drawing.Color]::Transparent
            $chk.Font      = $fontLabel
            $chk.Checked   = ($profile.EnableAutoRestart -eq $true)
            $chk.Tag       = $pfx
            $chk.Add_CheckedChanged({ $script:SharedState.Profiles[$this.Tag].EnableAutoRestart = $this.Checked })
            $card.Controls.Add($chk)

            $btnStart     = _Button 'Start'   200 55 70 30 $clrGreen $null
            $btnStart.Tag = $pfx
            $btnStart.Add_Click({
                $p = $this.Tag
                try {
                    $smPath = Join-Path $script:ModuleRoot 'ServerManager.psm1'
                    Import-Module $smPath -Force
                    Initialize-ServerManager -SharedState $script:SharedState
                    Start-GameServer -Prefix $p | Out-Null
                } catch {
                    [System.Windows.Forms.MessageBox]::Show("Start failed: $_",'Error','OK','Error') | Out-Null
                }
                _BuildServerDashboard
            })
            $card.Controls.Add($btnStart)

            $btnStop     = _Button 'Stop' 280 55 70 30 $clrRed $null
            $btnStop.Tag = $pfx
            $btnStop.Add_Click({
                $p = $this.Tag
                [System.Windows.Forms.MessageBox]::Show(
                    "Stopping $p - safe shutdown started (save + wait). Check logs for progress.",
                    'Stopping','OK','Information') | Out-Null

                $rs                = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
                $rs.ApartmentState = 'STA'
                $rs.ThreadOptions  = 'ReuseThread'
                $rs.Open()
                $rs.SessionStateProxy.SetVariable('ModulesDir',   $script:ModuleRoot)
                $rs.SessionStateProxy.SetVariable('SharedState',  $script:SharedState)
                $rs.SessionStateProxy.SetVariable('TargetPrefix', $p)

                $ps          = [System.Management.Automation.PowerShell]::Create()
                $ps.Runspace = $rs
                $ps.AddScript({
                    Set-StrictMode -Off
                    $ErrorActionPreference = 'Continue'
                    Import-Module (Join-Path $ModulesDir 'ProfileManager.psm1') -Force
                    Import-Module (Join-Path $ModulesDir 'ServerManager.psm1')  -Force
                    Initialize-ServerManager -SharedState $SharedState
                    Invoke-SafeShutdown -Prefix $TargetPrefix
                }) | Out-Null
                $ps.BeginInvoke() | Out-Null
            })
            $card.Controls.Add($btnStop)

            $btnRestart     = _Button 'Restart' 360 55 70 30 $clrAccent $null
            $btnRestart.Tag = $pfx
            $btnRestart.Add_Click({
                $p = $this.Tag
                try {
                    $smPath = Join-Path $script:ModuleRoot 'ServerManager.psm1'
                    Import-Module $smPath -Force
                    Initialize-ServerManager -SharedState $script:SharedState
                    Restart-GameServer -Prefix $p | Out-Null
                } catch {
                    [System.Windows.Forms.MessageBox]::Show("Restart failed: $_",'Error','OK','Error') | Out-Null
                }
                _BuildServerDashboard
            })
            $card.Controls.Add($btnRestart)

            $configRoots = _GetConfigRootsForProfile -Profile $profile
            $hasConfig   = ($configRoots.Count -gt 0)
            $btnConfig   = _Button 'Config' 440 55 70 30 $clrAccent $null
            $btnConfig.Tag = $pfx
            $btnConfig.Enabled = $hasConfig
            if (-not $hasConfig) { $btnConfig.BackColor = $clrPanel }
            $btnConfig.Add_Click({
                $p = $this.Tag
                $prof = $script:SharedState.Profiles[$p]
                if ($prof) { _OpenConfigEditor -Profile $prof }
            })
            $card.Controls.Add($btnConfig)

            # Clicking the card body (not a button) opens the profile editor
            $card.Add_Click({
                $pfxCapture = $this.Controls[0].Text -replace '^.+\[(\w+)\].*$','$1'
                $prof = $script:SharedState.Profiles[$pfxCapture]
                if ($prof) { _BuildProfileEditor -Profile $prof }
            })

            $panel.Controls.Add($card)
            $y += 120
        }
    }

    # =====================================================================
    # LIGHTWEIGHT STATUS UPDATER  (called every timer tick - no full rebuild)
    # =====================================================================
    function _UpdateDashboardStatus {
        $ss = $script:SharedState
        if (-not $ss -or -not $ss.Profiles) { return }

        foreach ($pfx in @($ss.Profiles.Keys)) {
            $entry   = $null
            $running = $false

            if ($ss.RunningServers -and $ss.RunningServers.ContainsKey($pfx)) {
                $entry = $ss.RunningServers[$pfx]
                $proc  = try { Get-Process -Id $entry.Pid -ErrorAction Stop } catch { $null }

                if ($null -ne $proc -and -not $proc.HasExited) {
                    $running = $true
                } else {
                    # Launcher PID may have exited (e.g. cmd.exe /c bat) - check by ProcessName
                    $procName = $ss.Profiles[$pfx].ProcessName
                    if ($procName) {
                        $gameProc = Get-Process -Name $procName -ErrorAction SilentlyContinue |
                                    Select-Object -First 1
                        if ($null -ne $gameProc -and -not $gameProc.HasExited) {
                            $running = $true
                            # Update entry so PID stays current
                            $entry = $ss.RunningServers[$pfx]
                        }
                    }
                }
            }

            foreach ($card in $script:_ServerDashboardPanel.Controls) {
                if ($card.Controls.Count -lt 3) { continue }
                $nameLabel = $card.Controls[0]
                # Match "[PFX]" anywhere in the name label text
                if ($nameLabel.Text -notmatch "\[$pfx\]") { continue }

                $statusLabel = $card.Controls[1]
                $uptimeLabel = $card.Controls[2]

                if ($running -and $entry) {
                    $statusLabel.Text      = 'ONLINE'
                    $statusLabel.ForeColor = $clrGreen
                    $up = [Math]::Round(((Get-Date) - $entry.StartTime).TotalMinutes, 1)
                    $uptimeLabel.Text = "PID $($entry.Pid) | Uptime: ${up} min"
                } else {
                    $statusLabel.Text      = 'OFFLINE'
                    $statusLabel.ForeColor = $clrRed
                    $uptimeLabel.Text      = 'Server is not running'
                }
            }
        }
    }

    # =====================================================================
    # INITIAL BUILDS
    # =====================================================================
    _BuildProfilesList
    _BuildProfileEditor $null
    _BuildServerDashboard

    # =====================================================================
    # LAYOUT REFLOW  (called on form resize)
    # =====================================================================
    function _ReflowLayout {
        $cw       = $form.ClientSize.Width
        $ch       = $form.ClientSize.Height
        $sbHeight = $statusBar.Height

        $topBar.Width  = $cw
        $topBar.Height = $topBarHeight

        # Move Settings button to right edge of top bar
        $btnSettings.Location = [System.Drawing.Point]::new($cw - 170, 15)

        $leftW   = if ($script:_LeftCollapsed)  { $collapsedSize } else { $leftWidth }
        $rightW  = if ($script:_RightCollapsed) { $collapsedSize } else { $rightWidth }
        $bottomH = if ($script:_BottomCollapsed){ $bottomHeaderHeight } else { $bottomLogsHeight }

        $bottomY = $ch - $sbHeight - $bottomH
        if ($bottomY -lt ($topBarHeight + 200)) { $bottomY = $topBarHeight + 200 }

        $bottomContainer.Location = [System.Drawing.Point]::new(0, $bottomY)
        $bottomContainer.Size     = [System.Drawing.Size]::new($cw, $bottomH)

        $leftContainer.Location  = [System.Drawing.Point]::new(0, $topBarHeight)
        $leftContainer.Size      = [System.Drawing.Size]::new($leftW, $bottomY - $topBarHeight)

        $rightContainer.Location = [System.Drawing.Point]::new($cw - $rightW, $topBarHeight)
        $rightContainer.Size     = [System.Drawing.Size]::new($rightW, $bottomY - $topBarHeight)

        $centerX = $leftW + $sideGap
        $centerW = $cw - $leftW - $rightW - ($sideGap * 2)
        if ($centerW -lt 400) { $centerW = 400 }
        $centerCol.Location = [System.Drawing.Point]::new($centerX, $topBarHeight)
        $centerCol.Size     = [System.Drawing.Size]::new($centerW, $bottomY - $topBarHeight)

        # Left header/body sizing
        if ($script:_LeftCollapsed) {
            $leftHeader.Location = [System.Drawing.Point]::new(0, 0)
            $leftHeader.Size     = [System.Drawing.Size]::new($leftW, $leftContainer.Height)
            $leftBody.Visible    = $false
            $leftHeaderLabel.Text = _VerticalText 'Profiles'
            $leftHeaderLabel.Location = [System.Drawing.Point]::new(3, 4)
            $leftHeaderLabel.Size = [System.Drawing.Size]::new([Math]::Max(16, $leftW - 6), $leftHeader.Height - 6)
            $leftHeaderLabel.TextAlign = 'TopCenter'
        } else {
            $leftHeader.Location = [System.Drawing.Point]::new(0, 0)
            $leftHeader.Size     = [System.Drawing.Size]::new($leftW, $headerHeight)
            $leftBody.Location   = [System.Drawing.Point]::new(0, $headerHeight)
            $leftBody.Size       = [System.Drawing.Size]::new($leftW, $leftContainer.Height - $headerHeight)
            $leftBody.Visible    = $true
            $leftHeaderLabel.Text = 'Game Profiles'
            $leftHeaderLabel.Location = [System.Drawing.Point]::new(6, 4)
            $leftHeaderLabel.Size = [System.Drawing.Size]::new($leftW - 12, $headerHeight - 6)
            $leftHeaderLabel.TextAlign = 'MiddleLeft'
        }

        # Right header/body sizing
        if ($script:_RightCollapsed) {
            $rightHeader.Location = [System.Drawing.Point]::new(0, 0)
            $rightHeader.Size     = [System.Drawing.Size]::new($rightW, $rightContainer.Height)
            $rightBody.Visible    = $false
            $rightHeaderLabel.Text = _VerticalText 'Editor'
            $rightHeaderLabel.Location = [System.Drawing.Point]::new(3, 4)
            $rightHeaderLabel.Size = [System.Drawing.Size]::new([Math]::Max(16, $rightW - 6), $rightHeader.Height - 6)
            $rightHeaderLabel.TextAlign = 'TopCenter'
        } else {
            $rightHeader.Location = [System.Drawing.Point]::new(0, 0)
            $rightHeader.Size     = [System.Drawing.Size]::new($rightW, $headerHeight)
            $rightBody.Location   = [System.Drawing.Point]::new(0, $headerHeight)
            $rightBody.Size       = [System.Drawing.Size]::new($rightW, $rightContainer.Height - $headerHeight)
            $rightBody.Visible    = $true
            $rightHeaderLabel.Text = 'Profile Editor'
            $rightHeaderLabel.Location = [System.Drawing.Point]::new(6, 4)
            $rightHeaderLabel.Size = [System.Drawing.Size]::new($rightW - 12, $headerHeight - 6)
            $rightHeaderLabel.TextAlign = 'MiddleLeft'
        }

        # Bottom header/body sizing
        if ($script:_BottomCollapsed) {
            $bottomHeader.Location = [System.Drawing.Point]::new(0, 0)
            $bottomHeader.Size     = [System.Drawing.Size]::new($cw, $bottomContainer.Height)
            $bottomPanel.Visible   = $false
        } else {
            $bottomHeader.Location = [System.Drawing.Point]::new(0, 0)
            $bottomHeader.Size     = [System.Drawing.Size]::new($cw, $bottomHeaderHeight)
            $bottomPanel.Location  = [System.Drawing.Point]::new(0, $bottomHeaderHeight)
            $bottomPanel.Size      = [System.Drawing.Size]::new($cw, $bottomContainer.Height - $bottomHeaderHeight)
            $bottomPanel.Visible   = $true
        }

        # Bottom panel internals (only resize controls that actually exist)
        if ($script:_DiscordFooter -and $tbSend -and $btnSend -and $btnClearDisc) {
            $tbSend.Width = [Math]::Max(100, $script:_DiscordFooter.Width - 220)
            $btnSend.Location = [System.Drawing.Point]::new([Math]::Max(120, $script:_DiscordFooter.Width - 205), 3)
            $btnClearDisc.Location = [System.Drawing.Point]::new([Math]::Max(200, $script:_DiscordFooter.Width - 110), 3)
        }

        # FIX: Log TabControl clipping
        if ($script:_LogTabControl) {
            $tabH = $script:_LogTabControl.ItemSize.Height
            $script:_LogTabControl.Size = [System.Drawing.Size]::new(
                $bottomPanel.Width,
                [Math]::Max(0, $bottomPanel.Height - $tabH - 6)
            )
        }

        # Profiles list panel
        if ($script:_ProfilesListPanel) {
            $script:_ProfilesListPanel.Size = [System.Drawing.Size]::new($leftBody.Width - 20, $leftBody.Height - 70)
        }

        # Profile editor scroll panel
        if ($script:_ProfileEditorPanel) {
            foreach ($ctrl in $script:_ProfileEditorPanel.Controls) {
                if ($ctrl -is [System.Windows.Forms.Panel]) {
                    $ctrl.Size = [System.Drawing.Size]::new($script:_ProfileEditorPanel.Width, $script:_ProfileEditorPanel.Height)
                    if ($script:_ProfileFields) {
                        $twNew = [Math]::Max(300, $ctrl.Width - (170 + 40))
                        foreach ($key in $script:_ProfileFields.Keys) {
                            $field = $script:_ProfileFields[$key]
                            if ($field -is [System.Windows.Forms.TextBox]) { $field.Width = $twNew }
                            if ($field -is [System.Collections.IDictionary] -and $field.Control -is [System.Windows.Forms.TextBox]) {
                                $field.Control.Width = $twNew
                            }
                        }
                        if ($script:_ProfileSeparators) {
                            foreach ($sep in $script:_ProfileSeparators) {
                                if ($sep) { $sep.Width = [Math]::Max(100, $ctrl.Width - 20) }
                            }
                        }
                    }
                }
            }
        }

        # Dashboard card widths
        foreach ($card in $script:_ServerDashboardPanel.Controls) {
            if ($card -is [System.Windows.Forms.Panel]) {
                $card.Width = $script:_ServerDashboardPanel.Width - 20
                foreach ($c in $card.Controls) {
                    if ($c -is [System.Windows.Forms.Label] -and $c.Anchor -eq 'Top,Right') {
                        $c.Location = [System.Drawing.Point]::new($card.Width - 120, $c.Location.Y)
                    }
                }
            }
        }
    }

    $form.add_Resize({ _ReflowLayout })
    _ReflowLayout


    # =====================================================================
    # MAIN TIMER  (2000ms — metrics, bot status, dashboard status, UI only)
    # =====================================================================
    $timer          = [System.Windows.Forms.Timer]::new()
    $timer.Interval = 2000

    # Listener control flags
    $script:_ListenerRestartRequested = $false

    function _ListenerIsRunning {
        $handle = $script:SharedState['ListenerHandle']
        if ($null -eq $handle) { return $false }
        return -not $handle.IsCompleted
    }

    function _StartListenerRunspace {
        if (_ListenerIsRunning) { return }

        $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $rs.ApartmentState = 'STA'
        $rs.ThreadOptions  = 'ReuseThread'
        $rs.Open()
        $rs.SessionStateProxy.SetVariable('ModulesDir',  $script:ModuleRoot)
        $rs.SessionStateProxy.SetVariable('SharedState', $script:SharedState)

        $ps = [System.Management.Automation.PowerShell]::Create()
        $ps.Runspace = $rs
        $ps.AddScript({
            Set-StrictMode -Off
            $ErrorActionPreference = 'Continue'
            try {
                Import-Module (Join-Path $ModulesDir 'Logging.psm1')         -Force
                Import-Module (Join-Path $ModulesDir 'ProfileManager.psm1')  -Force
                Import-Module (Join-Path $ModulesDir 'ServerManager.psm1')   -Force
                Import-Module (Join-Path $ModulesDir 'DiscordListener.psm1') -Force
                Start-DiscordListener -SharedState $SharedState
            } catch {
                if ($SharedState -and $SharedState.LogQueue) {
                    $SharedState.LogQueue.Enqueue(
                        "[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][ERROR][ListenerRunspace] $_")
                }
            }
        }) | Out-Null

        $handle = $ps.BeginInvoke()
        $script:SharedState['ListenerRunspace'] = $rs
        $script:SharedState['ListenerPS']       = $ps
        $script:SharedState['ListenerHandle']   = $handle
    }

    function _StopListenerRunspace {
        if (-not $script:SharedState.ContainsKey('StopListener')) {
            $script:SharedState['StopListener'] = $true
        } else {
            $script:SharedState['StopListener'] = $true
        }
    }

    $timer.add_Tick({
        try {
            # --- Read latest metrics written by background metrics runspace ---
            # (unchanged)

            # --- NO LOG TAILING HERE ANYMORE ---
            # Background runspaces now handle all file reading.
            # This timer ONLY updates UI and flushes SharedState queues.

            # --- Manage per-game log tabs (UI only, no file I/O) ---
            try {
                $ss = $script:SharedState
                if ($ss -and $ss.RunningServers) {

                    # Collect currently running prefixes
                    $runningPfx = @($ss.RunningServers.Keys)

                    # Remove tabs for servers that stopped
                    foreach ($deadPfx in @($script:_GameLogTabs.Keys)) {
                        if ($runningPfx -notcontains $deadPfx) {
                            _RemoveGameLogTab $deadPfx
                        }
                    }

                    foreach ($pfx in $runningPfx) {
                        $srvEntry = $ss.RunningServers[$pfx]
                        if ($null -eq $srvEntry) { continue }

                        # Get full profile
                        $profile = $null
                        if ($ss.Profiles -and $ss.Profiles.ContainsKey($pfx)) {
                            $profile = $ss.Profiles[$pfx]
                        }
                        if ($null -eq $profile) { continue }

                        $gameName = if ($profile.GameName) { $profile.GameName } else { $pfx }

                        # Ensure tab exists
                        _EnsureGameLogTab -Prefix $pfx -GameName $gameName

                        $tabEntry = $script:_GameLogTabs[$pfx]
                        $rtb      = $tabEntry.RTB
                        $lblSrc   = $tabEntry.LblSrc

                        # Resolve which files SHOULD be tailed (background thread handles actual tailing)
                        $logFiles = _ResolveGameLogFiles -Profile $profile

                        # Detect PZ session folder change (UI only)
                        if ($profile.LogStrategy -eq 'PZSessionFolder' -and
                            $profile.ServerLogRoot -and (Test-Path $profile.ServerLogRoot)) {

                            $newest = Get-ChildItem -Path $profile.ServerLogRoot -Directory -ErrorAction SilentlyContinue |
                                    Sort-Object Name -Descending | Select-Object -First 1

                            if ($newest -and $newest.Name -ne $tabEntry.LastSession) {
                                $rtb.Clear()
                                $tabEntry.Files       = @{}
                                $tabEntry.LastSession = $newest.Name
                                _AppendLog $rtb "[SESSION] New PZ session: $($newest.Name)" $clrYellow
                                # Reset start notification for new session
                                $script:_ServerStartNotified.Remove($pfx) | Out-Null
                            }
                        }

                        # If no log file found yet, show note once
                        if ($logFiles.Count -eq 0) {
                            $note = $profile.ServerLogNote
                            if (-not [string]::IsNullOrEmpty($note) -and $rtb.Lines.Count -lt 3) {
                                _AppendLog $rtb "[INFO] $note" $clrMuted
                            }
                            continue
                        }

                        # Update source label (UI only)
                        $srcNames = ($logFiles | ForEach-Object { Split-Path $_ -Leaf }) -join '  |  '
                        $lblSrc.Text = $srcNames
                    }

                } else {
                    # No running servers - remove stale tabs
                    foreach ($deadPfx in @($script:_GameLogTabs.Keys)) {
                        _RemoveGameLogTab $deadPfx
                    }
                }
            } catch { }

            # --- Drain SharedState LogQueue (UI only) ---
            for ($i = 0; $i -lt 20; $i++) {
                $item = $null
                if (-not $script:SharedState.LogQueue.TryDequeue([ref]$item)) { break }
                if ($item -match '\[Discord\]') {
                    _WriteDiscordLog $item
                } else {
                    _WriteProgramLog $item
                }
            }

            # --- Drain GameLogQueue (UI only) ---
            for ($i = 0; $i -lt 200; $i++) {
                $entry = $null
                if (-not $script:SharedState.GameLogQueue.TryDequeue([ref]$entry)) { break }
                if ($null -eq $entry) { continue }
                $pfx = $entry.Prefix
                $line = $entry.Line
                if (-not $pfx -or -not $line) { continue }
                if (-not $script:_GameLogTabs.ContainsKey($pfx)) { continue }
                $tabEntry = $script:_GameLogTabs[$pfx]
                _AppendLog $tabEntry.RTB $line (_LogColour $line)

                # Player list capture for Project Zomboid after !PZ players
                $reqs = $script:SharedState.PlayersRequests
                if ($reqs -and $reqs.ContainsKey($pfx)) {
                    $cap = $script:_PlayersCapture[$pfx]

                    # Start capture when "Players connected (N):" appears
                    if ($line -match 'Players connected\s*\((\d+)\):') {
                        $expected = [int]$Matches[1]
                        $script:_PlayersCapture[$pfx] = @{
                            Active   = $true
                            Expected = $expected
                            Names    = New-Object System.Collections.Generic.List[string]
                            Started  = Get-Date
                        }

                        if ($expected -eq 0) {
                            $gameName = $script:SharedState.Profiles[$pfx].GameName
                            $script:SharedState.WebhookQueue.Enqueue("[PLAYERS] ${gameName}: none")
                            $reqs.Remove($pfx)
                            $script:_PlayersCapture.Remove($pfx)
                        }
                        continue
                    }

                    # If capturing, collect "-Name" lines
                    if ($cap -and $cap.Active -eq $true) {
                        if ($line -match '^\s*-\s*(.+)$') {
                            $name = $Matches[1].Trim()
                            if ($name.EndsWith('.')) { $name = $name.Substring(0, $name.Length - 1) }
                            $cap.Names.Add($name) | Out-Null
                        }

                        $elapsed = ((Get-Date) - $cap.Started).TotalSeconds
                        if ($cap.Names.Count -ge $cap.Expected -or $elapsed -gt 5) {
                            $gameName = $script:SharedState.Profiles[$pfx].GameName
                            $names = if ($cap.Names.Count -gt 0) { $cap.Names -join ', ' } else { 'none' }
                            $script:SharedState.WebhookQueue.Enqueue("[PLAYERS] ${gameName}: $names")
                            $reqs.Remove($pfx)
                            $script:_PlayersCapture.Remove($pfx)
                        } else {
                            # keep updated capture state
                            $script:_PlayersCapture[$pfx] = $cap
                        }
                    }
                }

                # Detect server started markers and notify Discord once per session.
                $profile = $script:SharedState.Profiles[$pfx]
                if ($profile -and -not $script:_ServerStartNotified.ContainsKey($pfx)) {
                    if ($profile.GameName -match 'Project Zomboid' -and $line -match '\*\*\* SERVER STARTED \*\*\*\*\.') {
                        $script:_ServerStartNotified[$pfx] = $true
                        if ($script:SharedState.WebhookQueue) {
                            $script:SharedState.WebhookQueue.Enqueue("[JOINABLE] Project Zombiod Server Can Be Joined")
                        }
                    }
                    elseif ($profile.GameName -match 'Hytale' -and $line -match 'Hytale Server Booted! \[Multiplayer\]') {
                        $script:_ServerStartNotified[$pfx] = $true
                        if ($script:SharedState.WebhookQueue) {
                            $script:SharedState.WebhookQueue.Enqueue("[JOINABLE] Hytale Server Can Be Joined")
                        }
                    }
                }
            }

            # --- Listener control (Start/Stop/Restart) ---
            if ($script:SharedState.ContainsKey('RestartListener') -and $script:SharedState['RestartListener'] -eq $true) {
                $script:SharedState['RestartListener'] = $false
                $script:_ListenerRestartRequested = $true
                _StopListenerRunspace
            }

            if ($script:SharedState.ContainsKey('StopListener') -and $script:SharedState['StopListener'] -eq $true) {
                # If the listener has fully stopped, clear the flag
                if (-not (_ListenerIsRunning)) {
                    $script:SharedState['StopListener'] = $false
                }
            }

            if ($script:_ListenerRestartRequested -and -not (_ListenerIsRunning)) {
                $script:SharedState['StopListener'] = $false
                _StartListenerRunspace
                $script:_ListenerRestartRequested = $false
            }

            # --- Bot status ---
            $isRunning = $script:SharedState.ContainsKey('ListenerRunning') -and $script:SharedState['ListenerRunning'] -eq $true
            if ($isRunning) {
                $lblBot.Text      = 'Bot: Online'
                $lblBot.ForeColor = $clrGreen
            } else {
                $lblBot.Text      = 'Bot: Offline'
                $lblBot.ForeColor = $clrRed
            }

            # --- Lightweight dashboard refresh ---
            _UpdateDashboardStatus

            # --- Status bar ---
            $rc = if ($script:SharedState.RunningServers) { $script:SharedState.RunningServers.Count } else { 0 }
            $tc = if ($script:SharedState.Profiles)       { $script:SharedState.Profiles.Count       } else { 0 }
            $statusLabel.Text = "Profiles: $tc  |  Running: $rc  |  $(Get-Date -Format 'HH:mm:ss')"

            # --- Apply latest metrics from background runspace ---
            if ($script:SharedState.ContainsKey('_MetricCPU')) {
                $script:_cpuSmooth = _Smooth $script:_cpuSmooth ([double]$script:SharedState['_MetricCPU'])
                $script:_ramSmooth = _Smooth $script:_ramSmooth ([double]$script:SharedState['_MetricRAM'])
                $script:_netSmooth = _Smooth $script:_netSmooth ([double]$script:SharedState['_MetricNET'])

                $lblCPU.Text = 'CPU: {0:N0}%' -f $script:_cpuSmooth
                $lblRAM.Text = 'RAM: {0:N0}%' -f $script:_ramSmooth
                $lblNET.Text = if ($script:_netSmooth -gt 1024) {
                    'NET: {0:N1} MB/s' -f ($script:_netSmooth / 1024)
                } else {
                    'NET: {0:N0} KB/s' -f $script:_netSmooth
                }
                _SetMetricColor $lblCPU $script:_cpuSmooth
                _SetMetricColor $lblRAM $script:_ramSmooth
                $lblNET.ForeColor = $clrText
            }

        } catch {
            # Swallow all timer errors to keep the GUI alive
        }
    })


    $timer.Start()

    $form.add_FormClosing({
        $timer.Stop()
        if ($script:MetricsPS)        { $script:MetricsPS.Dispose() }
        if ($script:MetricsRunspace)  { $script:MetricsRunspace.Close(); $script:MetricsRunspace.Dispose() }
        if ($script:LogTailPS)        { $script:LogTailPS.Dispose() }
        if ($script:LogTailRunspace)  { $script:LogTailRunspace.Close(); $script:LogTailRunspace.Dispose() }
        $script:SharedState['StopListener'] = $true
    })

    [System.Windows.Forms.Application]::Run($form)
}

Export-ModuleMember -Function Start-GUI

