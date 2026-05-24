# =============================================================================
# GUI.psm1  -  WinForms GUI for Etherium Command Center
# PowerShell 5.1 compatible. No emoji in source strings.
# =============================================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Collections

# Console window hide/show + borderless resize helpers
try { [NativeWin]::GetConsoleWindow() | Out-Null } catch {
    Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class NativeWin {
  [DllImport("kernel32.dll")]
  public static extern IntPtr GetConsoleWindow();
  [DllImport("user32.dll")]
  public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
  [DllImport("user32.dll")]
  public static extern bool ReleaseCapture();
  [DllImport("user32.dll")]
  public static extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);
}
"@
}

$script:ModuleRoot  = $PSScriptRoot
$script:ProfilesDir = '.\Profiles'   # set properly in Start-GUI
$script:AppVersion  = 'v0.9'
$script:_BulkStartTimer = $null

# -------------------------
# Theme, fonts, and helpers
# -------------------------
$clrBg        = [System.Drawing.Color]::FromArgb(15,18,30)
$clrPanel     = [System.Drawing.Color]::FromArgb(31,36,52)
$clrPanelAlt  = [System.Drawing.Color]::FromArgb(38,44,63)
$clrPanelSoft = [System.Drawing.Color]::FromArgb(24,29,43)
$clrBorder    = [System.Drawing.Color]::FromArgb(64,72,98)
$clrShell     = [System.Drawing.Color]::FromArgb(19,23,36)
$clrEdge      = [System.Drawing.Color]::FromArgb(86,97,130)
$clrEdgeGlow  = [System.Drawing.Color]::FromArgb(110,165,255)
$clrText      = [System.Drawing.Color]::FromArgb(229,233,244)
$clrTextSoft  = [System.Drawing.Color]::FromArgb(160,170,194)
$clrAccent    = [System.Drawing.Color]::FromArgb(88,142,255)
$clrAccentAlt = [System.Drawing.Color]::FromArgb(122,168,255)
$clrGreen     = [System.Drawing.Color]::FromArgb(74,201,138)
$clrRed       = [System.Drawing.Color]::FromArgb(239,104,104)
$clrYellow    = [System.Drawing.Color]::FromArgb(240,198,92)
$clrMuted     = [System.Drawing.Color]::FromArgb(97,106,132)
$clrBtnText   = [System.Drawing.Color]::FromArgb(245,247,252)
$fontLabel = New-Object System.Drawing.Font("Segoe UI", 9)
$fontBold  = New-Object System.Drawing.Font("Segoe UI", 9,  [System.Drawing.FontStyle]::Bold)
$fontTitle = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$fontMono  = New-Object System.Drawing.Font("Consolas", 9)

function _GuiModuleLog {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR','DEBUG')][string]$Level = 'INFO'
    )

    $writeLog = Get-Command -Name 'Write-Log' -ErrorAction SilentlyContinue
    if ($writeLog) {
        try {
            Write-Log -Message $Message -Level $Level -Source 'GUI'
            return
        } catch { }
    }

    if ($script:SharedState -and $script:SharedState.LogQueue) {
        try {
            $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][$Level][GUI] $Message")
            return
        } catch { }
    }

    $color = switch ($Level) {
        'ERROR' { 'Red' }
        'WARN'  { 'Yellow' }
        'DEBUG' { 'Gray' }
        default { 'Cyan' }
    }
    try {
        $entry = "[{0}][{1}][GUI] {2}" -f (Get-Date -f 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
        Write-Host $entry -ForegroundColor $color
    } catch { }
}

function _GuiDirectLog {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR','DEBUG')][string]$Level = 'INFO'
    )

    if ([string]::IsNullOrWhiteSpace($Message)) { return }

    $writeLog = Get-Command -Name 'Write-Log' -ErrorAction SilentlyContinue
    if ($writeLog) {
        try {
            Write-Log -Message $Message -Level $Level -Source 'GUI'
            return
        } catch { }
    }

    $color = switch ($Level) {
        'ERROR' { 'Red' }
        'WARN'  { 'Yellow' }
        'DEBUG' { 'Gray' }
        default { 'Cyan' }
    }
    try {
        $entry = "[{0}][{1}][GUI] {2}" -f (Get-Date -f 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
        Write-Host $entry -ForegroundColor $color
    } catch { }
}

function _GuiCrashBreadcrumb {
    param(
        [string]$Message,
        [string]$Level = 'INFO'
    )

    if ([string]::IsNullOrWhiteSpace($Message)) { return }

    $writeBreadcrumb = Get-Command -Name 'Write-CrashBreadcrumb' -ErrorAction SilentlyContinue
    if ($writeBreadcrumb) {
        try {
            Write-CrashBreadcrumb -Category 'GUI' -Level $Level -Message $Message
            return
        } catch { }
    }

    try {
        $fallbackPath = Join-Path $PSScriptRoot '..\Logs\crash-breadcrumbs.log'
        $entry = "[{0}][{1}][GUI] {2}{3}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'), $Level.ToUpperInvariant(), $Message.Trim(), [Environment]::NewLine
        [System.IO.File]::AppendAllText($fallbackPath, $entry, [System.Text.Encoding]::UTF8)
    } catch { }
}

function _ShouldLogEmptyThreadExceptionEvent {
    $now = Get-Date
    $windowSeconds = 30
    try {
        if ($script:SharedState -and $script:SharedState.ContainsKey('LastEmptyGuiThreadExceptionAt')) {
            $lastAt = $script:SharedState['LastEmptyGuiThreadExceptionAt']
            if ($lastAt -is [datetime] -and (($now - $lastAt).TotalSeconds -lt $windowSeconds)) {
                return $false
            }
        }
        if ($script:SharedState) {
            $script:SharedState['LastEmptyGuiThreadExceptionAt'] = $now
        }
    } catch { }
    return $true
}

function _IsPerformanceDebugEnabled {
    try {
        if (-not $script:SharedState -or -not $script:SharedState.Settings) { return $false }
        if ([bool]$script:SharedState.Settings.EnablePerformanceDebugMode) { return $true }
        if ([bool]$script:SharedState.Settings.EnableDebugLogging) { return $true }
    } catch { }
    return $false
}

function _IsGuiDebugEnabled {
    try {
        if (-not $script:SharedState -or -not $script:SharedState.Settings) { return $false }
        return [bool]$script:SharedState.Settings.EnableDebugLogging
    } catch { }
    return $false
}

if ($null -eq $script:_GuiPerfTraceState) { $script:_GuiPerfTraceState = @{} }
if ($null -eq $script:_GuiTickSequence)   { $script:_GuiTickSequence   = 0 }

function _TraceGuiPerformanceSample {
    param(
        [string]$Area,
        [double]$ElapsedMs,
        [double]$WarnAtMs = 250,
        [double]$DebugAtMs = 100,
        [string]$Detail = ''
    )

    try {
        if ([string]::IsNullOrWhiteSpace($Area)) { return }
        if ($ElapsedMs -lt 0) { return }

        $debugEnabled = _IsPerformanceDebugEnabled

        $level = ''
        if ($ElapsedMs -ge $WarnAtMs) {
            $level = 'WARN'
        } elseif ($debugEnabled -and $ElapsedMs -ge $DebugAtMs) {
            $level = 'DEBUG'
        } else {
            return
        }

        $now = Get-Date
        $bucket = [int][Math]::Floor(([Math]::Max(0.0, [double]$ElapsedMs)) / 25.0)
        $minGapSeconds = if ($level -eq 'WARN') { 10 } else { 30 }
        $shouldLog = $true
        $lastState = $null

        try {
            if ($script:_GuiPerfTraceState.ContainsKey($Area)) {
                $lastState = $script:_GuiPerfTraceState[$Area]
            }
        } catch {
            $lastState = $null
        }

        if ($lastState) {
            $lastAt = $null
            $lastLevel = ''
            $lastBucket = -1
            try { $lastAt = $lastState.At } catch { $lastAt = $null }
            try { $lastLevel = [string]$lastState.Level } catch { $lastLevel = '' }
            try { $lastBucket = [int]$lastState.Bucket } catch { $lastBucket = -1 }

            if ($lastLevel -eq $level -and
                $lastAt -is [datetime] -and
                (($now - $lastAt).TotalSeconds -lt $minGapSeconds) -and
                ([Math]::Abs($bucket - $lastBucket) -lt 2)) {
                $shouldLog = $false
            }
        }

        if (-not $shouldLog) { return }

        $script:_GuiPerfTraceState[$Area] = @{
            At     = $now
            Level  = $level
            Bucket = $bucket
        }

        $message = 'UIPERF area={0} elapsedMs={1:N1} warnAtMs={2:N0}' -f $Area, ([Math]::Round($ElapsedMs, 1)), ([Math]::Round($WarnAtMs, 0))
        if (-not [string]::IsNullOrWhiteSpace($Detail)) {
            $message += ' detail=' + $Detail
        }

        _GuiModuleLog -Message $message -Level $level
    } catch { }
}

# =============================================================================
#  CONTROL FACTORY HELPERS
# =============================================================================
function _ResolveUiColor {
    param(
        $Color,
        [System.Drawing.Color]$Fallback = [System.Drawing.Color]::FromArgb(228, 234, 245)
    )

    if ($Color -is [System.Drawing.Color]) {
        return [System.Drawing.Color]$Color
    }

    return $Fallback
}

function _ResolveUiFont {
    param(
        $Font,
        $Fallback = $fontLabel
    )

    if ($Font -is [System.Drawing.Font]) {
        return [System.Drawing.Font]$Font
    }
    if ($Fallback -is [System.Drawing.Font]) {
        return [System.Drawing.Font]$Fallback
    }

    return New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Regular)
}

function _BindClickHandler {
    param(
        [System.Windows.Forms.Control]$Control,
        [scriptblock]$Handler
    )

    if ($null -eq $Control -or $null -eq $Handler) { return }
    $Control.Add_Click($Handler)
}

function _Label {
    param($text, $x, $y, $w, $h, $font = $fontLabel)
    $lbl           = New-Object System.Windows.Forms.Label
    $lbl.Text      = $text
    $lbl.Location  = [System.Drawing.Point]::new($x, $y)
    $lbl.Size      = [System.Drawing.Size]::new($w, $h)
    $lbl.ForeColor = _ResolveUiColor -Color $clrText -Fallback ([System.Drawing.Color]::FromArgb(228, 234, 245))
    $lbl.BackColor = [System.Drawing.Color]::Transparent
    $lbl.Font      = _ResolveUiFont -Font $font
    return $lbl
}

function _BlendColor {
    param(
        [System.Drawing.Color]$Base,
        [int]$Delta = 0
    )

    $r = [Math]::Max(0, [Math]::Min(255, $Base.R + $Delta))
    $g = [Math]::Max(0, [Math]::Min(255, $Base.G + $Delta))
    $b = [Math]::Max(0, [Math]::Min(255, $Base.B + $Delta))
    return [System.Drawing.Color]::FromArgb($r, $g, $b)
}

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

function _PZSharedNormalizeAssetKey {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    $normalized = [string]$Text
    $normalized = $normalized -replace '\\', '/'
    $normalized = [System.IO.Path]::GetFileNameWithoutExtension($normalized)
    $normalized = $normalized -replace '^Item_', ''
    $normalized = $normalized -replace '^items/', ''
    $normalized = $normalized -replace '[^A-Za-z0-9]+', ''
    return $normalized.ToLowerInvariant()
}

function _PZSharedAssetCacheRoot {
    $workspaceRoot = Split-Path -Parent $script:ModuleRoot
    return (Join-Path $workspaceRoot 'Config\AssetCache\ProjectZomboid')
}

function _PZSharedCatalogCacheDirectory {
    $dir = Join-Path (_PZSharedAssetCacheRoot) 'Catalogs'
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
    }
    return $dir
}

function _PZSharedCatalogCachePath {
    param(
        [string]$CatalogName,
        [string]$CacheKey
    )

    $safeName = if ([string]::IsNullOrWhiteSpace($CatalogName)) { 'catalog' } else { $CatalogName }
    $safeKey = if ([string]::IsNullOrWhiteSpace($CacheKey)) { 'default' } else { ($CacheKey -replace '[^A-Za-z0-9\-_\.]+', '_') }
    return (Join-Path (_PZSharedCatalogCacheDirectory) "$safeName-$safeKey.json")
}

function _PZSharedFindCatalogCacheFallbackPath {
    param(
        [string]$CatalogName,
        [string]$CacheKey
    )

    if ([string]::IsNullOrWhiteSpace($CatalogName) -or [string]::IsNullOrWhiteSpace($CacheKey)) { return $null }

    $cacheDir = _PZSharedCatalogCacheDirectory
    if (-not (Test-Path -LiteralPath $cacheDir)) { return $null }

    $cacheParts = @($CacheKey -split '\|')
    if ($cacheParts.Count -eq 0) { return $null }

    $gameRoot = if (@($cacheParts).Count -gt 0) { [string]$cacheParts[0] } else { '' }
    if ([string]::IsNullOrWhiteSpace($gameRoot)) { return $null }

    $safeRootPrefix = $gameRoot -replace '[^A-Za-z0-9\-_\.]+', '_'
    if ([string]::IsNullOrWhiteSpace($safeRootPrefix)) { return $null }

    $namePrefix = "$CatalogName-$safeRootPrefix"
    $matches = @(
        Get-ChildItem -LiteralPath $cacheDir -Filter "$CatalogName-*.json" -File -ErrorAction SilentlyContinue |
            Where-Object { $_.BaseName -like "$namePrefix*" } |
            Sort-Object LastWriteTimeUtc -Descending
    )

    if ($matches.Count -gt 0) {
    if (@($matches).Count -gt 0 -and $matches[0]) {
        return $matches[0].FullName
    }
    return $null
    }

    return $null
}

function _PZSharedLoadCatalogCache {
    param(
        [string]$CatalogName,
        [string]$CacheKey
    )

    $path = _PZSharedCatalogCachePath -CatalogName $CatalogName -CacheKey $CacheKey
    if (-not (Test-Path -LiteralPath $path)) {
        $fallbackPath = _PZSharedFindCatalogCacheFallbackPath -CatalogName $CatalogName -CacheKey $CacheKey
        if ([string]::IsNullOrWhiteSpace($fallbackPath) -or -not (Test-Path -LiteralPath $fallbackPath)) {
            return $null
        }
        $path = $fallbackPath
    }
    try {
        $raw = Get-Content -LiteralPath $path -Raw -ErrorAction Stop | ConvertFrom-Json
        if ($raw -is [System.Array]) {
            return [object[]]$raw
        }
        if ($null -ne $raw) {
            return [object[]]@($raw)
        }
    } catch {
        _GuiModuleLog -Message "PZ catalog cache load failed for '$path': $($_.Exception.Message)" -Level WARN
    }
    return $null
}

function _PZSharedLoadImportedItemTextureMap {
    $cacheRoot = _PZSharedAssetCacheRoot

    $directManifestPath = Join-Path $cacheRoot 'item-texture-manifest.json'
    if (Test-Path -LiteralPath $directManifestPath) {
        try {
            $directManifest = Get-Content -LiteralPath $directManifestPath -Raw -ErrorAction Stop | ConvertFrom-Json
            if ($directManifest -and $directManifest.ItemTextureByName) {
                return $directManifest
            }
        } catch { }
    }

    $importRoot = Join-Path $cacheRoot 'ImportedClient'
    $importManifestPath = Join-Path $importRoot 'import-manifest.json'
    if (-not (Test-Path -LiteralPath $importManifestPath)) { return $null }

    try {
        $importManifest = Get-Content -LiteralPath $importManifestPath -Raw -ErrorAction Stop | ConvertFrom-Json
        $itemFolder = [string]$importManifest.itemFolder
        $metadataFolder = [string]$importManifest.metadataFolder
        $itemXmlPath = if (-not [string]::IsNullOrWhiteSpace($metadataFolder)) { Join-Path $metadataFolder 'items.xml' } else { '' }
        if ([string]::IsNullOrWhiteSpace($itemFolder) -or -not (Test-Path -LiteralPath $itemFolder)) { return $null }

        $fileByKey = @{}
        foreach ($file in @(Get-ChildItem -LiteralPath $itemFolder -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -in @('.png', '.jpg', '.jpeg') })) {
            foreach ($key in @(
                $file.BaseName.ToLowerInvariant(),
                (_PZSharedNormalizeAssetKey $file.BaseName)
            ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique) {
                if (-not $fileByKey.ContainsKey($key)) {
                    $fileByKey[$key] = $file.FullName
                }
            }
        }

        $map = @{}
        if (-not [string]::IsNullOrWhiteSpace($itemXmlPath) -and (Test-Path -LiteralPath $itemXmlPath)) {
            [xml]$itemXml = Get-Content -LiteralPath $itemXmlPath -ErrorAction Stop
            foreach ($node in @($itemXml.itemManager.m_Items)) {
                $textureRef = [string]$node.m_Texture
                $modelRef = [string]$node.m_Model
                if ([string]::IsNullOrWhiteSpace($textureRef)) { continue }

                $resolvedPath = $null
                foreach ($key in @(
                    ([System.IO.Path]::GetFileName($textureRef)).ToLowerInvariant(),
                    (_PZSharedNormalizeAssetKey $textureRef)
                ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique) {
                    if ($fileByKey.ContainsKey($key)) {
                        $resolvedPath = $fileByKey[$key]
                        break
                    }
                }
                if (-not $resolvedPath) { continue }

                $itemKeys = @(
                    ([System.IO.Path]::GetFileName($textureRef)).ToLowerInvariant(),
                    (_PZSharedNormalizeAssetKey $textureRef)
                )
                if (-not [string]::IsNullOrWhiteSpace($modelRef)) {
                    $itemKeys += ([System.IO.Path]::GetFileName($modelRef)).ToLowerInvariant()
                    $itemKeys += (_PZSharedNormalizeAssetKey $modelRef)
                }

                foreach ($key in ($itemKeys | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)) {
                    if (-not $map.ContainsKey($key)) {
                        $map[$key] = $resolvedPath
                    }
                }
            }
        }

        $manifest = [ordered]@{
            SourceRoot = if ($importManifest.PSObject.Properties.Name -contains 'sourceRoot') { [string]$importManifest.sourceRoot } else { $importRoot }
            SourceTicks = if ($importManifest.PSObject.Properties.Name -contains 'importedAt') { [string]$importManifest.importedAt } else { '' }
            UpdatedAt = if ($importManifest.PSObject.Properties.Name -contains 'importedAt') { [string]$importManifest.importedAt } else { (Get-Date).ToString('o') }
            ItemTextureByName = $map
        }
        return [pscustomobject]$manifest
    } catch {
        return $null
    }
}

function _PZSharedLoadVehicleTextureMap {
    $cacheRoot = _PZSharedAssetCacheRoot
    $manifestPath = Join-Path $cacheRoot 'vehicle-texture-manifest.json'
    if (-not (Test-Path -LiteralPath $manifestPath)) { return $null }

    try {
        $manifest = Get-Content -LiteralPath $manifestPath -Raw -ErrorAction Stop | ConvertFrom-Json
        if ($manifest -and $manifest.VehicleTextureByName) {
            return $manifest
        }
    } catch { }

    return $null
}

function _PZSharedResolveVehiclePreviewPath {
    param(
        [string]$FullType,
        [string]$VehicleName,
        [string]$DisplayName,
        [object]$VehicleManifest
    )

    if ($null -eq $VehicleManifest -or -not $VehicleManifest.VehicleTextureByName) { return $null }

    $candidateKeys = @(
        $FullType,
        $VehicleName,
        $DisplayName,
        (_PZSharedNormalizeAssetKey $FullType),
        (_PZSharedNormalizeAssetKey $VehicleName),
        (_PZSharedNormalizeAssetKey $DisplayName)
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique

    foreach ($key in @($candidateKeys)) {
        try {
            $candidate = $null
            $vehicleTextureMap = $VehicleManifest.VehicleTextureByName
            if ($vehicleTextureMap -is [System.Collections.IDictionary]) {
                if ($vehicleTextureMap.Contains($key)) {
                    $candidate = $vehicleTextureMap[$key]
                }
            } else {
                $prop = $vehicleTextureMap.PSObject.Properties[$key]
                if ($null -ne $prop) {
                    $candidate = $prop.Value
                }
            }
            if (-not [string]::IsNullOrWhiteSpace($candidate) -and (Test-Path -LiteralPath $candidate)) {
                return [string]$candidate
            }
        } catch { }
    }

    return $null
}

function _PZSharedGetImportedItemImageIndex {
    if ($null -eq $script:ProjectZomboidImportedItemImageIndexCache) {
        $script:ProjectZomboidImportedItemImageIndexCache = $null
    }
    if ($script:ProjectZomboidImportedItemImageIndexCache) {
        return $script:ProjectZomboidImportedItemImageIndexCache
    }

    $itemRoot = Join-Path (_PZSharedAssetCacheRoot) 'ImportedClient\Items'
    $index = @{
        Exact = @{}
        Normalized = @{}
    }
    if (-not (Test-Path -LiteralPath $itemRoot)) {
        $script:ProjectZomboidImportedItemImageIndexCache = $index
        return $index
    }

    foreach ($file in @(Get-ChildItem -LiteralPath $itemRoot -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -in @('.png', '.jpg', '.jpeg') })) {
        foreach ($key in @(
            $file.BaseName.ToLowerInvariant(),
            (_PZSharedNormalizeAssetKey $file.BaseName)
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique) {
            if (-not $index.Exact.ContainsKey($key)) { $index.Exact[$key] = $file.FullName }
            if (-not $index.Normalized.ContainsKey($key)) { $index.Normalized[$key] = $file.FullName }
        }
    }

    $script:ProjectZomboidImportedItemImageIndexCache = $index
    return $index
}

function _PZSharedFindImportedItemImageByPattern {
    param([string[]]$Patterns)

    $itemRoot = Join-Path (_PZSharedAssetCacheRoot) 'ImportedClient\Items'
    if (-not (Test-Path -LiteralPath $itemRoot)) { return $null }

    foreach ($pattern in @($Patterns | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)) {
        try {
            $match = Get-ChildItem -LiteralPath $itemRoot -Recurse -File -ErrorAction SilentlyContinue |
                Where-Object { $_.BaseName -like $pattern } |
                Sort-Object FullName |
                Select-Object -First 1
            if ($match) { return $match.FullName }
        } catch { }
    }

    return $null
}

function _PZSharedResolveItemPreviewPath {
    param(
        [string]$ItemName,
        [string]$IconName,
        [object]$AssetManifest
    )

    $candidateKeys = @(
        $IconName,
        $ItemName,
        (_PZSharedNormalizeAssetKey $IconName),
        (_PZSharedNormalizeAssetKey $ItemName),
        ("items/{0}" -f [string]$IconName).ToLowerInvariant(),
        ("items/{0}" -f [string]$ItemName).ToLowerInvariant(),
        (_PZSharedNormalizeAssetKey ("items/{0}" -f [string]$IconName)),
        (_PZSharedNormalizeAssetKey ("items/{0}" -f [string]$ItemName)),
        ("Item_{0}" -f [string]$IconName).ToLowerInvariant(),
        ("Item_{0}" -f [string]$ItemName).ToLowerInvariant(),
        (_PZSharedNormalizeAssetKey ("Item_{0}" -f [string]$IconName)),
        (_PZSharedNormalizeAssetKey ("Item_{0}" -f [string]$ItemName))
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique

    if ($AssetManifest -and $AssetManifest.ItemTextureByName) {
        foreach ($key in @($candidateKeys)) {
            try {
                $candidate = $AssetManifest.ItemTextureByName.$key
                if (-not [string]::IsNullOrWhiteSpace($candidate) -and (Test-Path -LiteralPath $candidate)) {
                    return [string]$candidate
                }
            } catch { }
        }
    }

    $importedIndex = _PZSharedGetImportedItemImageIndex
    foreach ($key in @($candidateKeys)) {
        if ($importedIndex.Exact.ContainsKey($key)) {
            return [string]$importedIndex.Exact[$key]
        }
        if ($importedIndex.Normalized.ContainsKey($key)) {
            return [string]$importedIndex.Normalized[$key]
        }
    }

    foreach ($alt in @(
        ($IconName -replace 'loose$', ''),
        ($IconName -replace '\d+loose$', ''),
        ($IconName -replace '\d+$', ''),
        ($ItemName -replace 'Bullets$', ''),
        ($ItemName -replace 'Carton$', ''),
        ($ItemName -replace 'Box$', '')
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique) {
        foreach ($key in @(
            $alt.ToLowerInvariant(),
            (_PZSharedNormalizeAssetKey $alt)
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique) {
            if ($importedIndex.Exact.ContainsKey($key)) {
                return [string]$importedIndex.Exact[$key]
            }
            if ($importedIndex.Normalized.ContainsKey($key)) {
                return [string]$importedIndex.Normalized[$key]
            }
        }
    }

    $patternPath = _PZSharedFindImportedItemImageByPattern -Patterns @(
        ("{0}_*" -f [string]$IconName),
        ("*{0}" -f [string]$IconName),
        ("{0}_*" -f [string]$ItemName),
        ("*{0}" -f [string]$ItemName),
        ("*{0}*" -f [string]($IconName -replace 'loose$', '')),
        ("*{0}*" -f [string]($ItemName -replace 'Bullets$', ''))
    )
    if (-not [string]::IsNullOrWhiteSpace($patternPath)) {
        return $patternPath
    }

    return $null
}

function _MinecraftSharedAssetCacheRoot {
    $workspaceRoot = Split-Path -Parent $script:ModuleRoot
    $dir = Join-Path $workspaceRoot 'Config\AssetCache\Minecraft'
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
    }
    return $dir
}

function _MinecraftSharedCatalogCacheDirectory {
    $dir = Join-Path (_MinecraftSharedAssetCacheRoot) 'Catalogs'
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
    }
    return $dir
}

function _MinecraftSharedCatalogCachePath {
    param(
        [string]$CatalogName,
        [string]$CacheKey
    )

    $safeName = if ([string]::IsNullOrWhiteSpace($CatalogName)) { 'catalog' } else { $CatalogName }
    $safeKey = if ([string]::IsNullOrWhiteSpace($CacheKey)) { 'default' } else { ($CacheKey -replace '[^A-Za-z0-9\-_\.]+', '_') }
    return (Join-Path (_MinecraftSharedCatalogCacheDirectory) "$safeName-$safeKey.json")
}

function _MinecraftSharedLoadCatalogCache {
    param(
        [string]$CatalogName,
        [string]$CacheKey
    )

    $path = _MinecraftSharedCatalogCachePath -CatalogName $CatalogName -CacheKey $CacheKey
    if (-not (Test-Path -LiteralPath $path)) { return $null }
    try {
        $raw = Get-Content -LiteralPath $path -Raw -ErrorAction Stop | ConvertFrom-Json
        if ($raw -is [System.Collections.IEnumerable]) {
            return @($raw | ForEach-Object { [pscustomobject]$_ })
        }
    } catch {
        _GuiModuleLog -Message "Minecraft catalog cache load failed for '$path': $($_.Exception.Message)" -Level WARN
    }
    return $null
}

function _MinecraftSharedSaveCatalogCache {
    param(
        [string]$CatalogName,
        [string]$CacheKey,
        [object[]]$Entries
    )

    $path = _MinecraftSharedCatalogCachePath -CatalogName $CatalogName -CacheKey $CacheKey
    try {
        $json = @($Entries) | ConvertTo-Json -Depth 6
        [System.IO.File]::WriteAllText($path, $json, (New-Object System.Text.UTF8Encoding($false)))
    } catch {
        _GuiModuleLog -Message "Minecraft catalog cache save failed for '$path': $($_.Exception.Message)" -Level WARN
    }
}

function _GetMinecraftSpawnerCacheKey {
    param([hashtable]$Profile)

    $gameRoot = ''
    try { $gameRoot = [string]$Profile.FolderPath } catch { $gameRoot = '' }
    if ([string]::IsNullOrWhiteSpace($gameRoot)) { $gameRoot = 'default' }
    return $gameRoot
}

function _GetMinecraftSpawnerJarCandidates {
    param([hashtable]$Profile)

    $candidates = New-Object 'System.Collections.Generic.List[string]'
    $seen = @{}

    $addCandidate = {
        param([string]$Path)
        if ([string]::IsNullOrWhiteSpace($Path)) { return }
        try {
            $resolved = [System.IO.Path]::GetFullPath($Path)
        } catch {
            $resolved = $Path
        }
        if (-not (Test-Path -LiteralPath $resolved)) { return }
        $key = $resolved.ToLowerInvariant()
        if ($seen.ContainsKey($key)) { return }
        $seen[$key] = $true
        $candidates.Add($resolved) | Out-Null
    }.GetNewClosure()

    $gameRoot = ''
    try { $gameRoot = [string]$Profile.FolderPath } catch { $gameRoot = '' }
    $executable = ''
    try { $executable = [string]$Profile.Executable } catch { $executable = '' }

    if (-not [string]::IsNullOrWhiteSpace($executable)) {
        if ([System.IO.Path]::IsPathRooted($executable)) {
            & $addCandidate $executable
        } elseif (-not [string]::IsNullOrWhiteSpace($gameRoot)) {
            & $addCandidate (Join-Path $gameRoot $executable)
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($gameRoot) -and (Test-Path -LiteralPath $gameRoot)) {
        try {
            $localJarMatches = @(
                Get-ChildItem -LiteralPath $gameRoot -File -Filter *.jar -ErrorAction SilentlyContinue |
                    Sort-Object LastWriteTimeUtc -Descending
            )
            foreach ($match in $localJarMatches) {
                & $addCandidate $match.FullName
            }
        } catch { }
    }

    $appData = [Environment]::GetFolderPath('ApplicationData')
    if (-not [string]::IsNullOrWhiteSpace($appData)) {
        $versionsRoot = Join-Path $appData '.minecraft\versions'
        if (Test-Path -LiteralPath $versionsRoot) {
            try {
                $clientJarMatches = @(
                    Get-ChildItem -LiteralPath $versionsRoot -Recurse -File -Filter *.jar -ErrorAction SilentlyContinue |
                        Where-Object {
                            $_.DirectoryName -like '*\.minecraft\versions\*' -and
                            $_.BaseName -notmatch '(?i)server'
                        } |
                        Sort-Object LastWriteTimeUtc -Descending
                )
                foreach ($match in $clientJarMatches) {
                    & $addCandidate $match.FullName
                }
            } catch { }
        }
    }

    return @($candidates.ToArray())
}

function _GetMinecraftSpawnerDisplayCategory {
    param(
        [string]$ItemId,
        [string]$ModelPath = ''
    )

    if ([string]::IsNullOrWhiteSpace($ItemId)) { return 'Items' }

    $id = $ItemId.ToLowerInvariant()
    if ($id -like '*_spawn_egg') { return 'Spawn Eggs' }
    if ($id -like 'music_disc_*') { return 'Music Discs' }
    if ($id -like '*_boat' -or $id -like '*_chest_boat' -or $id -like 'raft' -or $id -like '*_chest_raft') { return 'Boats' }
    if ($id -like '*_minecart') { return 'Minecarts' }
    if ($id -match '_(sword|axe|pickaxe|shovel|hoe)$' -or $id -in @('bow','crossbow','trident','mace','fishing_rod','shears','flint_and_steel','compass','recovery_compass','clock')) { return 'Tools and Combat' }
    if ($id -match '_(helmet|chestplate|leggings|boots)$' -or $id -like '*horse_armor' -or $id -eq 'elytra' -or $id -eq 'shield') { return 'Armor and Gear' }
    if ($id -match '_(door|trapdoor|fence|fence_gate|planks|slab|stairs|log|wood|stem|hyphae|pressure_plate|button|sign|hanging_sign|bed|banner)$' -or
        $id -match '^(oak|spruce|birch|jungle|acacia|dark_oak|mangrove|cherry|bamboo|crimson|warped|pale_oak)_' -or
        $id -match '_(bricks|wall|glass|pane|terracotta|concrete|concrete_powder|wool|carpet)$' -or
        $id -match '^(stone|cobblestone|andesite|diorite|granite|deepslate|tuff|calcite|basalt|blackstone|sandstone|red_sandstone|prismarine|purpur|quartz|obsidian|crying_obsidian|netherrack|end_stone)$') { return 'Blocks and Building' }
    if ($id -match '(apple|bread|beef|porkchop|chicken|mutton|rabbit|cod|salmon|tropical_fish|potato|carrot|beetroot|melon_slice|cookie|pumpkin_pie|stew|soup|berries|fruit|cake|pie|golden_apple|enchanted_golden_apple|chorus_fruit|rotten_flesh|dried_kelp|honey_bottle|milk_bucket)$' -or
        $id -match '^(cooked_|baked_)' -or
        $id -match '^suspicious_stew$') { return 'Food' }
    if ($id -match '(redstone|repeater|comparator|observer|hopper|dispenser|dropper|piston|lever|daylight_detector|target|sculk_sensor|calibrated_sculk_sensor|tripwire_hook|detector_rail|activator_rail|powered_rail|rail|lightning_rod)$') { return 'Redstone and Logic' }
    if ($id -match '(bucket|saddle|lead|name_tag|spyglass|brush|goat_horn|totem_of_undying|ender_pearl|ender_eye|firework|fire_charge|wind_charge|snowball|egg)$') { return 'Utility and Travel' }
    if ($id -match '(ingot|diamond|emerald|lapis|redstone|coal|charcoal|quartz|amethyst|shard|scrap|netherite|raw_|nugget)$' -or
        $id -match '(stick|string|feather|leather|rabbit_hide|paper|book|enchanted_book|writable_book|written_book|ink_sac|glow_ink_sac|slime_ball|magma_cream|blaze_rod|blaze_powder|ghast_tear|gunpowder|phantom_membrane|echo_shard|trial_key|ominous_trial_key|heavy_core)$') { return 'Materials' }
    if ($id -match '^(grass_block|dirt|coarse_dirt|rooted_dirt|podzol|mycelium|farmland|sand|red_sand|gravel|clay|mud|soul_sand|soul_soil|moss_block|moss_carpet|short_grass|fern|dead_bush|vine|sugar_cane|bamboo|cactus|kelp|seagrass|torchflower|pitcher_plant|sunflower|lilac|rose_bush|peony|allium|azure_bluet|blue_orchid|cornflower|dandelion|poppy|oxeye_daisy|pink_tulip|red_tulip|orange_tulip|white_tulip|wither_rose|spore_blossom|glow_berries|sweet_berries|wheat_seeds|beetroot_seeds|melon_seeds|pumpkin_seeds|nether_wart|cocoa_beans|lily_pad)$') { return 'Nature and Farming' }

    return 'Items'
}

function _GetMinecraftSpawnerDisplayName {
    param(
        [string]$ItemId,
        [hashtable]$LangMap
    )

    if ([string]::IsNullOrWhiteSpace($ItemId)) { return 'Unknown Item' }

    $langKeys = @(
        "item.minecraft.$ItemId",
        "block.minecraft.$ItemId"
    )

    foreach ($key in $langKeys) {
        if ($LangMap.ContainsKey($key) -and -not [string]::IsNullOrWhiteSpace([string]$LangMap[$key])) {
            return [string]$LangMap[$key]
        }
    }

    return (($ItemId -split '_') | ForEach-Object {
        if ([string]::IsNullOrWhiteSpace($_)) { return '' }
        if ($_.Length -le 1) { return $_.ToUpperInvariant() }
        return ($_.Substring(0,1).ToUpperInvariant() + $_.Substring(1))
    }) -join ' '
}

function _GetMinecraftSpawnerItemsFromJar {
    param([string]$JarPath)

    if ([string]::IsNullOrWhiteSpace($JarPath) -or -not (Test-Path -LiteralPath $JarPath)) { return @() }

    try { Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue } catch { }

    $archive = $null
    $langMap = @{}
    $itemIdMap = @{}
    try {
        $archive = [System.IO.Compression.ZipFile]::OpenRead($JarPath)

        foreach ($entry in @($archive.Entries)) {
            $fullName = [string]$entry.FullName
            if ($fullName -match '^assets/minecraft/models/item/(?<id>[^/]+)\.json$') {
                $itemId = [string]$Matches['id']
                if (-not [string]::IsNullOrWhiteSpace($itemId) -and $itemId -notlike 'template_*') {
                    $itemIdMap[$itemId.ToLowerInvariant()] = $fullName
                }
            }
        }

        $langEntry = @($archive.Entries | Where-Object { [string]$_.FullName -ieq 'assets/minecraft/lang/en_us.json' } | Select-Object -First 1)
        if ($langEntry.Count -gt 0 -and $langEntry[0]) {
            $reader = $null
            try {
                $reader = New-Object System.IO.StreamReader($langEntry[0].Open())
                $langRaw = $reader.ReadToEnd()
                if (-not [string]::IsNullOrWhiteSpace($langRaw)) {
                    $langJson = $langRaw | ConvertFrom-Json -ErrorAction Stop
                    foreach ($prop in $langJson.PSObject.Properties) {
                        $langMap[[string]$prop.Name] = [string]$prop.Value
                    }
                }
            } finally {
                if ($reader) { $reader.Dispose() }
            }
        }
    } catch {
        _GuiModuleLog -Message "Minecraft spawner jar scan failed for '$JarPath': $($_.Exception.Message)" -Level WARN
        return @()
    } finally {
        if ($archive) { $archive.Dispose() }
    }

    $items = New-Object 'System.Collections.Generic.List[object]'
    foreach ($itemId in @($itemIdMap.Keys | Sort-Object)) {
        if ($itemId -in @('air','cave_air','void_air')) { continue }
        $displayName = _GetMinecraftSpawnerDisplayName -ItemId $itemId -LangMap $langMap
        $category = _GetMinecraftSpawnerDisplayCategory -ItemId $itemId -ModelPath $itemIdMap[$itemId]
        $fullType = "minecraft:$itemId"
        $listText = '{0,-34} {1}' -f $displayName, $fullType
        $items.Add([pscustomobject]@{
            DisplayName     = $displayName
            FullType        = $fullType
            ItemId          = $itemId
            DisplayCategory = $category
            ListText        = $listText
            IconPath        = ''
        }) | Out-Null
    }

    return @($items.ToArray())
}

function _ReadBundledMinecraftSpawnerCatalogRaw {
    $workspaceRoot = Split-Path -Parent $script:ModuleRoot
    $catalogPath = Join-Path $workspaceRoot 'Config\MinecraftVanillaItemCatalog.json'
    if (-not ($script:BundledMinecraftSpawnerCatalog -is [System.Array] -and $script:BundledMinecraftSpawnerCatalog.Count -gt 0)) {
        if (-not (Test-Path -LiteralPath $catalogPath)) {
            _GuiModuleLog -Message "Bundled Minecraft item catalog missing at '$catalogPath'." -Level WARN
            return @()
        }

        try {
            $raw = Get-Content -LiteralPath $catalogPath -Raw -ErrorAction Stop | ConvertFrom-Json
            $items = @()
            if ($raw -is [System.Collections.IEnumerable]) {
                $items = @($raw | ForEach-Object { [pscustomobject]$_ })
            } elseif ($null -ne $raw) {
                $items = @([pscustomobject]$raw)
            }
            $script:BundledMinecraftSpawnerCatalog = @($items)
        } catch {
            _GuiModuleLog -Message "Bundled Minecraft item catalog load failed: $($_.Exception.Message)" -Level WARN
            return @()
        }
    }

    return @($script:BundledMinecraftSpawnerCatalog)
}

function _LoadBundledMinecraftSpawnerCatalog {
    param([hashtable]$Profile = $null)

    $rawItems = @(_ReadBundledMinecraftSpawnerCatalogRaw)
    if ($rawItems.Count -le 0) { return @() }

    $manifest = _MinecraftSharedLoadItemTextureManifest
    if ($null -eq $manifest) {
        $syncProfile = if ($null -ne $Profile) { $Profile } else { @{ FolderPath = ''; KnownGame = 'Minecraft' } }
        $manifest = _SyncMinecraftItemAssetCache -Profile $syncProfile
        if ($null -eq $manifest) {
            $manifest = _MinecraftSharedLoadItemTextureManifest
        }
    }

    $itemsWithIcons = New-Object 'System.Collections.Generic.List[object]'
    foreach ($item in @($rawItems)) {
        $itemClone = [pscustomobject]@{}
        foreach ($prop in $item.PSObject.Properties) {
            try {
                $itemClone | Add-Member -NotePropertyName $prop.Name -NotePropertyValue $prop.Value -Force
            } catch { }
        }
        $resolvedIconPath = _MinecraftSharedResolveItemPreviewPath -ItemId ([string]$itemClone.ItemId) -FullType ([string]$itemClone.FullType) -Manifest $manifest
        if ($itemClone.PSObject.Properties.Name -notcontains 'IconPath') {
            $itemClone | Add-Member -NotePropertyName IconPath -NotePropertyValue $resolvedIconPath -Force
        } else {
            $itemClone.IconPath = $resolvedIconPath
        }
        $itemsWithIcons.Add($itemClone) | Out-Null
    }

    return @($itemsWithIcons.ToArray())
}

function _MinecraftSharedItemTextureManifestPath {
    return (Join-Path (_MinecraftSharedAssetCacheRoot) 'item-texture-manifest.json')
}

function _MinecraftSharedItemTextureDirectory {
    $dir = Join-Path (_MinecraftSharedAssetCacheRoot) 'Items'
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
    }
    return $dir
}

function _MinecraftSharedLoadItemTextureManifest {
    $manifestPath = _MinecraftSharedItemTextureManifestPath
    if (-not (Test-Path -LiteralPath $manifestPath)) { return $null }
    try {
        $manifest = Get-Content -LiteralPath $manifestPath -Raw -ErrorAction Stop | ConvertFrom-Json
        if ($manifest -and $manifest.ItemTextureByName) { return $manifest }
    } catch {
        _GuiModuleLog -Message "Minecraft item texture manifest load failed: $($_.Exception.Message)" -Level WARN
    }
    return $null
}

function _MinecraftSharedResolveItemPreviewPath {
    param(
        [string]$ItemId,
        [string]$FullType = '',
        [object]$Manifest = $null
    )

    if ([string]::IsNullOrWhiteSpace($ItemId) -and [string]::IsNullOrWhiteSpace($FullType)) { return $null }
    if ($null -eq $Manifest) {
        $Manifest = _MinecraftSharedLoadItemTextureManifest
    }
    if ($null -eq $Manifest -or -not $Manifest.ItemTextureByName) { return $null }

    $map = $Manifest.ItemTextureByName
    $candidateKeys = @(
        $ItemId,
        $FullType,
        ($FullType -replace '^minecraft:', '')
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Select-Object -Unique

    foreach ($key in @($candidateKeys)) {
        try {
            $candidate = $null
            if ($map -is [System.Collections.IDictionary]) {
                if ($map.Contains($key)) { $candidate = $map[$key] }
            } else {
                $prop = $map.PSObject.Properties[$key]
                if ($null -ne $prop) { $candidate = $prop.Value }
            }
            if (-not [string]::IsNullOrWhiteSpace([string]$candidate) -and (Test-Path -LiteralPath $candidate)) {
                return [string]$candidate
            }
        } catch { }
    }

    return $null
}

function _ReadBundledHytaleSpawnerCatalogRaw {
    $workspaceRoot = Split-Path -Parent $script:ModuleRoot
    $catalogPath = Join-Path $workspaceRoot 'Config\HytaleVanillaItemCatalog.json'
    if (-not ($script:BundledHytaleSpawnerCatalog -is [System.Array] -and $script:BundledHytaleSpawnerCatalog.Count -gt 0)) {
        if (-not (Test-Path -LiteralPath $catalogPath)) {
            _GuiModuleLog -Message "Bundled Hytale item catalog missing at '$catalogPath'." -Level WARN
            return @()
        }

        try {
            $raw = Get-Content -LiteralPath $catalogPath -Raw -ErrorAction Stop | ConvertFrom-Json
            $items = @()
            if ($raw -is [System.Collections.IEnumerable]) {
                $items = @($raw | ForEach-Object { [pscustomobject]$_ })
            } elseif ($null -ne $raw) {
                $items = @([pscustomobject]$raw)
            }
            $script:BundledHytaleSpawnerCatalog = @($items)
        } catch {
            _GuiModuleLog -Message "Bundled Hytale item catalog load failed: $($_.Exception.Message)" -Level WARN
            return @()
        }
    }

    return @($script:BundledHytaleSpawnerCatalog)
}

function _HytaleSharedAssetCacheRoot {
    $workspaceRoot = Split-Path -Parent $script:ModuleRoot
    return (Join-Path $workspaceRoot 'Config\AssetCache\Hytale')
}

function _HytaleSharedItemTextureDirectory {
    $dir = Join-Path (_HytaleSharedAssetCacheRoot) 'Items'
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
    }
    return $dir
}

function _HytaleSharedItemIconManifestPath {
    return (Join-Path (_HytaleSharedAssetCacheRoot) 'item-icon-manifest.json')
}

function _HytaleSharedLoadItemIconManifest {
    $manifestPath = _HytaleSharedItemIconManifestPath
    if (-not (Test-Path -LiteralPath $manifestPath)) { return $null }
    try {
        $manifest = Get-Content -LiteralPath $manifestPath -Raw -ErrorAction Stop | ConvertFrom-Json
        if ($manifest -and $manifest.ItemIconByName) { return $manifest }
    } catch {
        _GuiModuleLog -Message "Hytale item icon manifest load failed: $($_.Exception.Message)" -Level WARN
    }
    return $null
}

function _GetHytaleSpawnerIconRoots {
    param([hashtable]$Profile = $null)

    $roots = New-Object 'System.Collections.Generic.List[string]'
    $candidates = @()
    try {
        if ($Profile) {
            $candidates += [string]$Profile.FolderPath
            $candidates += [string]$Profile.ConfigRoot
        }
    } catch { }

    foreach ($candidate in @($candidates | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Select-Object -Unique)) {
        try {
            $full = [System.IO.Path]::GetFullPath($candidate)
            if (Test-Path -LiteralPath $full) {
                $roots.Add($full) | Out-Null
            }
        } catch { }
    }

    return @($roots)
}

function _GetHytaleSpawnerIconCategoryHints {
    param([string]$Category)

    switch ($Category) {
        'Armors'      { return @('armor') }
        'Weapons'     { return @('weapon') }
        'Tools'       { return @('tool') }
        'Projectiles' { return @('projectile') }
        'Back'        { return @('backpack', 'cape', 'quiver', 'armor') }
        'Consumables' { return @('food', 'bandage', 'potion', 'scroll', 'recipe', 'consumable') }
        'Ingredients' { return @('ingredient') }
        'Torch'       { return @('torch') }
        'Vehicles'    { return @('vehicle', 'cart', 'boat', 'glider') }
        default       { return @() }
    }
}

function _GetHytaleSpawnerIconTokenScore {
    param(
        [string[]]$ItemSegments,
        [string]$Category,
        [string[]]$IconTokens,
        [string]$IconBaseName
    )

    $score = 0
    $segmentTokens = New-Object 'System.Collections.Generic.List[string]'
    foreach ($segment in @($ItemSegments)) {
        foreach ($token in @(([string]$segment).ToLowerInvariant().Split('_') | Where-Object { $_ })) {
            if (-not $segmentTokens.Contains($token)) {
                $segmentTokens.Add($token) | Out-Null
            }
        }
    }
    foreach ($hint in @(_GetHytaleSpawnerIconCategoryHints -Category $Category)) {
        if (-not $segmentTokens.Contains($hint)) {
            $segmentTokens.Add($hint) | Out-Null
        }
    }

    foreach ($token in @($segmentTokens)) {
        if ($IconTokens -contains $token) {
            $score += 10
        }
    }

    switch ($Category) {
        'Armors' {
            if ($IconBaseName -like 'Armor_*' -or $IconBaseName -like 'Amor_*') { $score += 8 }
        }
        'Weapons' {
            if ($IconBaseName -like 'Weapon_*') { $score += 8 }
        }
        'Tools' {
            if ($IconBaseName -like 'Tool_*') { $score += 8 }
        }
        'Consumables' {
            if ($IconBaseName -like 'Food_*' -or $IconBaseName -like 'Bandage_*' -or $IconBaseName -like 'Potion_*' -or $IconBaseName -like 'Scroll*' -or $IconBaseName -like 'Recipe*') { $score += 8 }
        }
        'Back' {
            if ($IconBaseName -like '*Cape*' -or $IconBaseName -like '*Backpack*' -or $IconBaseName -like '*Quiver*') { $score += 8 }
        }
        'Torch' {
            if ($IconBaseName -like '*Torch*') { $score += 8 }
        }
    }

    if ($ItemSegments.Count -gt 0) {
        $lastSegment = [string]$ItemSegments[$ItemSegments.Count - 1]
        foreach ($token in @($lastSegment.ToLowerInvariant().Split('_') | Where-Object { $_ })) {
            if ($IconTokens -contains $token) { $score += 6 }
        }
    }
    if ($ItemSegments.Count -gt 1) {
        $materialSegment = [string]$ItemSegments[$ItemSegments.Count - 2]
        foreach ($token in @($materialSegment.ToLowerInvariant().Split('_') | Where-Object { $_ })) {
            if ($IconTokens -contains $token) { $score += 6 }
        }
    }

    return $score
}

function _SyncHytaleItemAssetCache {
    param([hashtable]$Profile = $null)

    $iconSourceDir = $null
    foreach ($root in @(_GetHytaleSpawnerIconRoots -Profile $Profile)) {
        $candidate = Join-Path $root 'Assets\Common\Icons\ItemsGenerated'
        if (Test-Path -LiteralPath $candidate) {
            $iconSourceDir = $candidate
            break
        }
    }
    if ($null -eq $iconSourceDir) { return $null }

    $catalogItems = @(_ReadBundledHytaleSpawnerCatalogRaw)
    if ($catalogItems.Count -le 0) { return $null }

    $iconFiles = @(Get-ChildItem -LiteralPath $iconSourceDir -File -Filter '*.png' -ErrorAction SilentlyContinue)
    if ($iconFiles.Count -le 0) { return $null }

    $sourceSignature = @(
        try { (Get-Item -LiteralPath $iconSourceDir -ErrorAction Stop).FullName } catch { $iconSourceDir }
        try { (Get-Item -LiteralPath $iconSourceDir -ErrorAction Stop).LastWriteTimeUtc.Ticks } catch { '' }
        $iconFiles.Count
    ) -join '|'

    $existing = _HytaleSharedLoadItemIconManifest
    if ($existing -and [string]$existing.SourceSignature -eq $sourceSignature) {
        return $existing
    }

    $iconRows = foreach ($iconFile in @($iconFiles)) {
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($iconFile.Name)
        [pscustomobject]@{
            BaseName = $baseName
            Path     = $iconFile.FullName
            Tokens   = @($baseName.ToLowerInvariant().Split('_') | Where-Object { $_ })
        }
    }

    $cacheDir = _HytaleSharedItemTextureDirectory
    $iconMap = @{}
    $matchedCount = 0

    foreach ($item in @($catalogItems)) {
        $fullType = ''
        try { $fullType = [string]$item.FullType } catch { $fullType = '' }
        if ([string]::IsNullOrWhiteSpace($fullType)) { continue }

        $segments = @($fullType.Split('/') | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        if ($segments.Count -le 0) { continue }
        $category = [string]$segments[0]

        $best = $null
        $bestScore = -1
        foreach ($iconRow in @($iconRows)) {
            $score = _GetHytaleSpawnerIconTokenScore -ItemSegments $segments -Category $category -IconTokens $iconRow.Tokens -IconBaseName $iconRow.BaseName
            if ($score -gt $bestScore) {
                $bestScore = $score
                $best = $iconRow
            }
        }

        $cacheFileName = (($fullType -replace '[\\/:*?"<>|]+', '_') + '.png')
        $cachePath = Join-Path $cacheDir $cacheFileName

        if ($best -and $bestScore -ge 20) {
            try {
                Copy-Item -LiteralPath $best.Path -Destination $cachePath -Force
                $iconMap[$fullType] = $cachePath
                $iconMap[[string]$item.ItemId] = $cachePath
                $matchedCount++
                continue
            } catch { }
        }

        if (Test-Path -LiteralPath $cachePath) {
            $iconMap[$fullType] = $cachePath
            $iconMap[[string]$item.ItemId] = $cachePath
        }
    }

    $manifest = [ordered]@{
        SourceDirectory = $iconSourceDir
        SourceSignature = $sourceSignature
        UpdatedAt       = (Get-Date).ToString('o')
        MatchedCount    = $matchedCount
        ItemIconByName  = $iconMap
    }

    try {
        $manifestJson = $manifest | ConvertTo-Json -Depth 6
        [System.IO.File]::WriteAllText((_HytaleSharedItemIconManifestPath), $manifestJson, (New-Object System.Text.UTF8Encoding($false)))
    } catch {
        _GuiModuleLog -Message "Hytale item icon manifest save failed: $($_.Exception.Message)" -Level WARN
    }

    return [pscustomobject]$manifest
}

function _HytaleSharedResolveItemPreviewPath {
    param(
        [string]$ItemId,
        [string]$FullType = '',
        [object]$Manifest = $null
    )

    $candidateIds = @($ItemId, $FullType) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Select-Object -Unique
    if ($candidateIds.Count -le 0) { return $null }

    if ($null -eq $Manifest) {
        $Manifest = _HytaleSharedLoadItemIconManifest
    }
    if ($Manifest -and $Manifest.ItemIconByName) {
        $map = $Manifest.ItemIconByName
        foreach ($candidateId in @($candidateIds)) {
            try {
                $candidate = $null
                if ($map -is [System.Collections.IDictionary]) {
                    if ($map.Contains($candidateId)) { $candidate = $map[$candidateId] }
                } else {
                    $prop = $map.PSObject.Properties[$candidateId]
                    if ($null -ne $prop) { $candidate = $prop.Value }
                }
                if (-not [string]::IsNullOrWhiteSpace([string]$candidate) -and (Test-Path -LiteralPath $candidate)) {
                    return [string]$candidate
                }
            } catch { }
        }
    }

    $itemsDir = _HytaleSharedItemTextureDirectory
    foreach ($candidateId in @($candidateIds)) {
        $fileName = (($candidateId -replace '[\\/:*?"<>|]+', '_') + '.png')
        $path = Join-Path $itemsDir $fileName
        if (Test-Path -LiteralPath $path) {
            return $path
        }
    }

    return $null
}

function _LoadBundledHytaleSpawnerCatalog {
    param([hashtable]$Profile = $null)

    $rawItems = @(_ReadBundledHytaleSpawnerCatalogRaw)
    if ($rawItems.Count -le 0) { return @() }

    $manifest = _HytaleSharedLoadItemIconManifest

    $itemsWithIcons = New-Object 'System.Collections.Generic.List[object]'
    foreach ($item in @($rawItems)) {
        $itemClone = [pscustomobject]@{}
        foreach ($prop in $item.PSObject.Properties) {
            try {
                $itemClone | Add-Member -NotePropertyName $prop.Name -NotePropertyValue $prop.Value -Force
            } catch { }
        }

        $resolvedIconPath = _HytaleSharedResolveItemPreviewPath -ItemId ([string]$itemClone.ItemId) -FullType ([string]$itemClone.FullType) -Manifest $manifest
        if ($itemClone.PSObject.Properties.Name -notcontains 'IconPath') {
            $itemClone | Add-Member -NotePropertyName IconPath -NotePropertyValue $resolvedIconPath -Force
        } else {
            $itemClone.IconPath = $resolvedIconPath
        }
        $itemsWithIcons.Add($itemClone) | Out-Null
    }

    return @($itemsWithIcons.ToArray())
}

function Get-HytaleSpawnerCatalogsFromCache {
    param([hashtable]$Profile)

    return [pscustomobject]@{
        Items    = @(_LoadBundledHytaleSpawnerCatalog -Profile $Profile)
        Vehicles = @()
    }
}

function Get-HytaleSpawnerCatalogs {
    param([hashtable]$Profile)

    return [pscustomobject]@{
        Items    = @(_LoadBundledHytaleSpawnerCatalog -Profile $Profile)
        Vehicles = @()
    }
}

function _GetMinecraftSpawnerModelTextureReference {
    param(
        [System.IO.Compression.ZipArchive]$Archive,
        [hashtable]$ArchiveIndex = $null,
        [string]$ModelEntryPath
    )

    if ($null -eq $Archive -or [string]::IsNullOrWhiteSpace($ModelEntryPath)) { return $null }
    $entry = @()
    if ($ArchiveIndex) {
        $key = $ModelEntryPath.ToLowerInvariant()
        if ($ArchiveIndex.ContainsKey($key)) {
            $entry = @($ArchiveIndex[$key])
        }
    } else {
        $entry = @($Archive.Entries | Where-Object { [string]$_.FullName -ieq $ModelEntryPath } | Select-Object -First 1)
    }
    if ($entry.Count -le 0 -or -not $entry[0]) { return $null }

    $reader = $null
    try {
        $reader = New-Object System.IO.StreamReader($entry[0].Open())
        $raw = $reader.ReadToEnd()
        if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
        $json = $raw | ConvertFrom-Json -ErrorAction Stop
        if ($json.textures) {
            foreach ($textureKey in @('layer0','particle','all','top','side')) {
                $prop = $json.textures.PSObject.Properties[$textureKey]
                if ($null -ne $prop -and -not [string]::IsNullOrWhiteSpace([string]$prop.Value)) {
                    return [string]$prop.Value
                }
            }
            foreach ($prop in $json.textures.PSObject.Properties) {
                if ($prop -and -not [string]::IsNullOrWhiteSpace([string]$prop.Value)) {
                    return [string]$prop.Value
                }
            }
        }
    } catch { }
    finally {
        if ($reader) { $reader.Dispose() }
    }

    return $null
}

function _SyncMinecraftItemAssetCache {
    param([hashtable]$Profile)

    $jarCandidates = @(_GetMinecraftSpawnerJarCandidates -Profile $Profile)
    if ($jarCandidates.Count -le 0) {
        return $null
    }

    try { Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue } catch { }

    $itemCacheRoot = _MinecraftSharedItemTextureDirectory
    $catalogItems = @(_ReadBundledMinecraftSpawnerCatalogRaw)
    if ($catalogItems.Count -le 0) { return $null }

    foreach ($jarCandidate in $jarCandidates) {
        if ([string]::IsNullOrWhiteSpace([string]$jarCandidate) -or -not (Test-Path -LiteralPath $jarCandidate)) { continue }

        $jarFile = $null
        try { $jarFile = Get-Item -LiteralPath $jarCandidate -ErrorAction Stop } catch { $jarFile = $null }
        if ($null -eq $jarFile) { continue }

        $existing = _MinecraftSharedLoadItemTextureManifest
        if ($existing -and [string]$existing.SourceJar -eq $jarFile.FullName -and [string]$existing.SourceTicks -eq [string]$jarFile.LastWriteTimeUtc.Ticks) {
            return $existing
        }

        $archive = $null
        $itemTextureByName = @{}
        try {
            $archive = [System.IO.Compression.ZipFile]::OpenRead($jarFile.FullName)
            $archiveIndex = @{}
            foreach ($entry in @($archive.Entries)) {
                $archiveIndex[[string]$entry.FullName.ToLowerInvariant()] = $entry
            }

            $hasAssets = $archiveIndex.ContainsKey('assets/minecraft/lang/en_us.json')
            if (-not $hasAssets) {
                continue
            }

            foreach ($item in $catalogItems) {
                $itemId = ''
                try { $itemId = [string]$item.ItemId } catch { $itemId = '' }
                if ([string]::IsNullOrWhiteSpace($itemId)) {
                    try { $itemId = ([string]$item.FullType -replace '^minecraft:', '') } catch { $itemId = '' }
                }
                if ([string]::IsNullOrWhiteSpace($itemId)) { continue }

                $textureEntryPath = "assets/minecraft/textures/item/$itemId.png"
                $textureEntry = @()
                $textureEntryKey = $textureEntryPath.ToLowerInvariant()
                if ($archiveIndex.ContainsKey($textureEntryKey)) {
                    $textureEntry = @($archiveIndex[$textureEntryKey])
                }

                if ($textureEntry.Count -le 0 -or -not $textureEntry[0]) {
                    $modelEntryPath = "assets/minecraft/models/item/$itemId.json"
                    $textureRef = _GetMinecraftSpawnerModelTextureReference -Archive $archive -ArchiveIndex $archiveIndex -ModelEntryPath $modelEntryPath
                    if (-not [string]::IsNullOrWhiteSpace($textureRef)) {
                        $normalizedTextureRef = [string]$textureRef
                        if ($normalizedTextureRef -match '^[^:]+:(.+)$') {
                            $normalizedTextureRef = $Matches[1]
                        }
                        $normalizedTextureRef = $normalizedTextureRef -replace '^#', ''
                        if (-not $normalizedTextureRef.ToLowerInvariant().EndsWith('.png')) {
                            $normalizedTextureRef = "$normalizedTextureRef.png"
                        }
                        if (-not $normalizedTextureRef.ToLowerInvariant().StartsWith('textures/')) {
                            $normalizedTextureRef = "textures/$normalizedTextureRef"
                        }
                        $resolvedEntryPath = "assets/minecraft/$normalizedTextureRef"
                        $resolvedEntryKey = $resolvedEntryPath.ToLowerInvariant()
                        if ($archiveIndex.ContainsKey($resolvedEntryKey)) {
                            $textureEntry = @($archiveIndex[$resolvedEntryKey])
                        } else {
                            $textureEntry = @()
                        }
                    }
                }

                if ($textureEntry.Count -gt 0 -and $textureEntry[0]) {
                    $outputPath = Join-Path $itemCacheRoot ("{0}.png" -f ($itemId -replace '[^A-Za-z0-9\-_\.]+', '_'))
                    try {
                        $sourceStream = $null
                        $targetStream = $null
                        try {
                            $sourceStream = $textureEntry[0].Open()
                            $targetStream = [System.IO.File]::Open($outputPath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
                            $sourceStream.CopyTo($targetStream)
                        } finally {
                            if ($targetStream) { $targetStream.Dispose() }
                            if ($sourceStream) { $sourceStream.Dispose() }
                        }

                        $itemTextureByName[$itemId] = $outputPath
                        $itemTextureByName["minecraft:$itemId"] = $outputPath
                    } catch { }
                }
            }
        } catch {
            _GuiModuleLog -Message "Minecraft item asset cache sync failed for '$($jarFile.FullName)': $($_.Exception.Message)" -Level WARN
            continue
        } finally {
            if ($archive) { $archive.Dispose() }
        }

        if ($itemTextureByName.Count -gt 0) {
            $manifest = [ordered]@{
                SourceJar         = $jarFile.FullName
                SourceTicks       = [string]$jarFile.LastWriteTimeUtc.Ticks
                UpdatedAt         = (Get-Date).ToString('o')
                ItemTextureByName = $itemTextureByName
            }

            try {
                $manifestJson = $manifest | ConvertTo-Json -Depth 6
                [System.IO.File]::WriteAllText((_MinecraftSharedItemTextureManifestPath), $manifestJson, (New-Object System.Text.UTF8Encoding($false)))
            } catch {
                _GuiModuleLog -Message "Minecraft item texture manifest save failed: $($_.Exception.Message)" -Level WARN
            }

            return [pscustomobject]$manifest
        }
    }

    return $null
}

function Get-MinecraftSpawnerCatalogsFromCache {
    param([hashtable]$Profile)

    return [pscustomobject]@{
        Items    = @(_LoadBundledMinecraftSpawnerCatalog -Profile $Profile)
        Vehicles = @()
    }
}

function Get-MinecraftSpawnerCatalogs {
    param([hashtable]$Profile)

    return [pscustomobject]@{
        Items    = @(_LoadBundledMinecraftSpawnerCatalog -Profile $Profile)
        Vehicles = @()
    }
}

function Get-ProjectZomboidSpawnerCatalogsFromCache {
    param([hashtable]$Profile)

    $gameRoot = ''
    try { $gameRoot = [string]$Profile.FolderPath } catch { $gameRoot = '' }
    if ([string]::IsNullOrWhiteSpace($gameRoot) -or -not (Test-Path -LiteralPath $gameRoot)) {
        return [pscustomobject]@{ Items = @(); Vehicles = @() }
    }

    $assetManifest = _PZSharedLoadImportedItemTextureMap
    $assetSignature = ''
    try {
        if ($assetManifest) {
            $assetSignature = if ($assetManifest.PSObject.Properties.Name -contains 'SourceTicks') { [string]$assetManifest.SourceTicks } else { '' }
            if ([string]::IsNullOrWhiteSpace($assetSignature) -and $assetManifest.PSObject.Properties.Name -contains 'UpdatedAt') {
                $assetSignature = [string]$assetManifest.UpdatedAt
            }
        }
    } catch { $assetSignature = '' }

    $items = @()
    $itemScriptsDir = Join-Path $gameRoot 'media\scripts\generated\items'
    if (-not (Test-Path -LiteralPath $itemScriptsDir)) {
        $itemScriptsDir = Join-Path $gameRoot 'media\scripts\items'
    }
    if (Test-Path -LiteralPath $itemScriptsDir) {
        $itemLatestTicks = 0L
        try {
            $latestFile = Get-ChildItem -LiteralPath $itemScriptsDir -Recurse -File -Include *.txt -ErrorAction SilentlyContinue |
                Sort-Object LastWriteTimeUtc -Descending |
                Select-Object -First 1
            if ($latestFile) { $itemLatestTicks = $latestFile.LastWriteTimeUtc.Ticks }
        } catch { }
        $itemCacheKey = "$gameRoot|$itemLatestTicks|$assetSignature"
        $items = _PZSharedLoadCatalogCache -CatalogName 'pz-item-catalog' -CacheKey $itemCacheKey
    }

    $vehicles = @()
    $vehicleScriptsDir = Join-Path $gameRoot 'media\scripts\generated\vehicles'
    if (-not (Test-Path -LiteralPath $vehicleScriptsDir)) {
        $vehicleScriptsDir = Join-Path $gameRoot 'media\scripts\vehicles'
    }
    if (Test-Path -LiteralPath $vehicleScriptsDir) {
        $vehicleManifest = _PZSharedLoadVehicleTextureMap
        $vehicleAssetSignature = ''
        try {
            if ($vehicleManifest) {
                $vehicleAssetSignature = if ($vehicleManifest.PSObject.Properties.Name -contains 'SourceTicks') { [string]$vehicleManifest.SourceTicks } else { '' }
                if ([string]::IsNullOrWhiteSpace($vehicleAssetSignature) -and $vehicleManifest.PSObject.Properties.Name -contains 'UpdatedAt') {
                    $vehicleAssetSignature = [string]$vehicleManifest.UpdatedAt
                }
            }
        } catch { $vehicleAssetSignature = '' }
        $vehicleLatestTicks = 0L
        try {
            $latestFile = Get-ChildItem -LiteralPath $vehicleScriptsDir -Recurse -File -Include *.txt -ErrorAction SilentlyContinue |
                Sort-Object LastWriteTimeUtc -Descending |
                Select-Object -First 1
            if ($latestFile) { $vehicleLatestTicks = $latestFile.LastWriteTimeUtc.Ticks }
        } catch { }
        $vehicleCacheKey = "$gameRoot|vehicles|$vehicleLatestTicks|$vehicleAssetSignature"
        $vehicles = _PZSharedLoadCatalogCache -CatalogName 'pz-vehicle-catalog' -CacheKey $vehicleCacheKey
        if ($vehicles -and $vehicles.Count -gt 0) {
            foreach ($vehicle in @($vehicles)) {
                try {
                    $resolvedPreviewPath = if ($vehicle.PSObject.Properties.Name -contains 'PreviewPath') { [string]$vehicle.PreviewPath } else { '' }
                } catch { $resolvedPreviewPath = '' }
                if ([string]::IsNullOrWhiteSpace($resolvedPreviewPath) -or -not (Test-Path -LiteralPath $resolvedPreviewPath)) {
                    $resolvedPreviewPath = _PZSharedResolveVehiclePreviewPath -FullType ([string]$vehicle.FullType) -VehicleName ([string]$vehicle.VehicleName) -DisplayName ([string]$vehicle.DisplayName) -VehicleManifest $vehicleManifest
                    if (-not [string]::IsNullOrWhiteSpace($resolvedPreviewPath)) {
                        try {
                            if ($vehicle.PSObject.Properties.Name -notcontains 'PreviewPath') {
                                $vehicle | Add-Member -NotePropertyName PreviewPath -NotePropertyValue $resolvedPreviewPath -Force
                            } else {
                                $vehicle.PreviewPath = $resolvedPreviewPath
                            }
                        } catch { }
                    }
                }
            }
        }
    }

    return [pscustomobject]@{
        Items    = if ($null -ne $items) { $items } else { @() }
        Vehicles = if ($null -ne $vehicles) { $vehicles } else { @() }
    }
}

function _ApplyButtonChrome {
    param(
        [System.Windows.Forms.Button]$Button,
        [System.Drawing.Color]$BaseColor
    )

    if ($null -eq $Button) { return }
    $resolvedBaseColor = _ResolveUiColor -Color $BaseColor -Fallback ([System.Drawing.Color]::FromArgb(64, 93, 154))
    $resolvedBtnText   = _ResolveUiColor -Color $clrBtnText -Fallback ([System.Drawing.Color]::White)
    $resolvedTextSoft  = _ResolveUiColor -Color $clrTextSoft -Fallback ([System.Drawing.Color]::FromArgb(176, 182, 194))
    $resolvedPanelSoft = _ResolveUiColor -Color $clrPanelSoft -Fallback ([System.Drawing.Color]::FromArgb(31, 35, 41))
    $resolvedBorder    = _ResolveUiColor -Color $clrBorder -Fallback ([System.Drawing.Color]::FromArgb(74, 82, 110))
    $Button.FlatStyle                 = 'Flat'
    $Button.FlatAppearance.BorderSize = 1
    $baseBorder = _BlendColor $resolvedBaseColor -12
    $hoverColor = _BlendColor $resolvedBaseColor 18
    $hoverBorder = _BlendColor $hoverColor -8
    $Button.FlatAppearance.BorderColor = $baseBorder
    $Button.BackColor                 = $resolvedBaseColor
    $Button.ForeColor                 = $resolvedBtnText
    $Button.Font                      = _ResolveUiFont -Font $fontBold
    $Button.TextAlign                 = [System.Drawing.ContentAlignment]::MiddleCenter
    $Button.Padding                   = [System.Windows.Forms.Padding]::new(0)
    $Button.UseVisualStyleBackColor   = $false
    $Button.AutoEllipsis              = $true
    $Button.AutoSize                  = $false

    $baseLocal  = $resolvedBaseColor
    $hoverLocal = $hoverColor
    $baseBorderLocal = $baseBorder
    $hoverBorderLocal = $hoverBorder
    $btnTextLocal = $resolvedBtnText
    $textSoftLocal = $resolvedTextSoft
    $panelSoftLocal = $resolvedPanelSoft
    $borderLocal = $resolvedBorder

    $Button.Add_MouseEnter({
        $btn = [System.Windows.Forms.Button]$this
        if ($btn.Enabled) {
            $btn.BackColor = $hoverLocal
            $btn.FlatAppearance.BorderColor = $hoverBorderLocal
        }
    }.GetNewClosure())
    $Button.Add_MouseLeave({
        $btn = [System.Windows.Forms.Button]$this
        if ($btn.Enabled) {
            $btn.BackColor = $baseLocal
            $btn.FlatAppearance.BorderColor = $baseBorderLocal
        }
    }.GetNewClosure())
    $Button.Add_EnabledChanged({
        $btn = [System.Windows.Forms.Button]$this
        if ($btn.Enabled) {
            $btn.ForeColor = $btnTextLocal
            $btn.BackColor = $baseLocal
            $btn.FlatAppearance.BorderColor = $baseBorderLocal
        } else {
            $btn.ForeColor = $textSoftLocal
            $btn.BackColor = $panelSoftLocal
            $btn.FlatAppearance.BorderColor = $borderLocal
        }
    }.GetNewClosure())
}

function _ApplyWindowChromeButton {
    param(
        [System.Windows.Forms.Button]$Button,
        [string]$Text,
        [System.Drawing.Color]$BaseColor,
        [System.Drawing.Color]$HoverColor
    )

    if ($null -eq $Button) { return }
    $resolvedBaseColor = _ResolveUiColor -Color $BaseColor -Fallback ([System.Drawing.Color]::FromArgb(45, 58, 82))
    $resolvedHoverColor = _ResolveUiColor -Color $HoverColor -Fallback ([System.Drawing.Color]::FromArgb(86, 112, 168))
    $resolvedBtnText = _ResolveUiColor -Color $clrBtnText -Fallback ([System.Drawing.Color]::White)
    $Button.Text      = $Text
    $Button.Font      = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    $Button.FlatStyle = 'Flat'
    $Button.FlatAppearance.BorderSize = 1
    $Button.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(74, 82, 110)
    $Button.BackColor = $resolvedBaseColor
    $Button.ForeColor = [System.Drawing.Color]::FromArgb(228, 234, 245)
    $Button.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $Button.Padding   = [System.Windows.Forms.Padding]::new(0)
    $Button.Cursor    = 'Hand'
    $Button.UseVisualStyleBackColor = $false

    $baseChrome = $resolvedBaseColor
    $hoverChrome = $resolvedHoverColor
    $pressChrome = [System.Drawing.Color]::FromArgb(
        [Math]::Max(0, $resolvedHoverColor.R - 18),
        [Math]::Max(0, $resolvedHoverColor.G - 18),
        [Math]::Max(0, $resolvedHoverColor.B - 18)
    )
    $baseText = [System.Drawing.Color]::FromArgb(228, 234, 245)
    $hoverText = $resolvedBtnText
    $Button.Add_MouseEnter({
        $this.BackColor = $hoverChrome
        $this.ForeColor = $hoverText
        $this.FlatAppearance.BorderColor = $hoverChrome
    }.GetNewClosure())
    $Button.Add_MouseLeave({
        $this.BackColor = $baseChrome
        $this.ForeColor = $baseText
        $this.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(74, 82, 110)
    }.GetNewClosure())
    $Button.Add_MouseDown({
        param($s, $e)
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
            $this.BackColor = $pressChrome
            $this.ForeColor = $hoverText
            $this.FlatAppearance.BorderColor = $pressChrome
        }
    }.GetNewClosure())
    $Button.Add_MouseUp({
        param($s, $e)
        $screenPt = [System.Windows.Forms.Control]::MousePosition
        $clientPt = $this.PointToClient($screenPt)
        if ($this.ClientRectangle.Contains($clientPt)) {
            $this.BackColor = $hoverChrome
            $this.ForeColor = $hoverText
            $this.FlatAppearance.BorderColor = $hoverChrome
        } else {
            $this.BackColor = $baseChrome
            $this.ForeColor = $baseText
            $this.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(74, 82, 110)
        }
    }.GetNewClosure())
}

function _TextBox {
    param($x, $y, $w, $h, $text = '', $pass = $false)
    $tb                       = New-Object System.Windows.Forms.TextBox
    $tb.Location              = [System.Drawing.Point]::new($x, $y)
    $tb.Size                  = [System.Drawing.Size]::new($w, $h)
    $tb.Text                  = $text
    $tb.UseSystemPasswordChar = $pass
    $tb.BackColor             = _ResolveUiColor -Color $clrPanelSoft -Fallback ([System.Drawing.Color]::FromArgb(31, 35, 41))
    $tb.ForeColor             = _ResolveUiColor -Color $clrText -Fallback ([System.Drawing.Color]::FromArgb(228, 234, 245))
    $tb.BorderStyle           = 'FixedSingle'
    $tb.Font                  = _ResolveUiFont -Font $fontLabel
    return $tb
}

# All parameters positional: text, x, y, w, h, bg, onClick
function _Button {
    param($text, $x, $y, $w, $h, $bg, $onClick)
    $btn                              = New-Object System.Windows.Forms.Button
    $btn.Text                         = $text
    $btn.Location                     = [System.Drawing.Point]::new($x, $y)
    $btn.Size                         = [System.Drawing.Size]::new($w, $h)
    _ApplyButtonChrome -Button $btn -BaseColor (_ResolveUiColor -Color $bg -Fallback ([System.Drawing.Color]::FromArgb(56, 72, 102)))
    _BindClickHandler -Control $btn -Handler $onClick
    return $btn
}

function _SetDashboardStartButtonState {
    param(
        [System.Windows.Forms.Button]$Button,
        [bool]$Enabled,
        [string]$StateCode = ''
    )

    if ($null -eq $Button) { return }

    $Button.Enabled = $Enabled
    if ($Enabled) {
        $Button.BackColor = $clrGreen
        $Button.ForeColor = $clrBtnText
        $Button.FlatAppearance.BorderColor = (_BlendColor $clrGreen -12)
        return
    }

    $normalizedState = if ([string]::IsNullOrWhiteSpace($StateCode)) { '' } else { $StateCode.ToLowerInvariant() }
    $disabledBackColor = switch ($normalizedState) {
        'starting'   { [System.Drawing.Color]::FromArgb(30, 42, 68) }
        'restarting' { [System.Drawing.Color]::FromArgb(54, 45, 24) }
        default      { [System.Drawing.Color]::FromArgb(31, 35, 41) }
    }

    $Button.BackColor = $disabledBackColor
    $Button.ForeColor = $clrTextSoft
    $Button.FlatAppearance.BorderColor = (_BlendColor $disabledBackColor 10)
}

function _SetDashboardStopButtonState {
    param(
        [System.Windows.Forms.Button]$Button,
        [bool]$Enabled,
        [string]$StateCode = ''
    )

    if ($null -eq $Button) { return }

    $Button.Enabled = $Enabled
    if ($Enabled) {
        $Button.BackColor = $clrRed
        $Button.ForeColor = $clrBtnText
        $Button.FlatAppearance.BorderColor = (_BlendColor $clrRed -12)
        return
    }

    $normalizedState = if ([string]::IsNullOrWhiteSpace($StateCode)) { '' } else { $StateCode.ToLowerInvariant() }
    $disabledBackColor = switch ($normalizedState) {
        'stopped' { [System.Drawing.Color]::FromArgb(46, 30, 34) }
        default   { [System.Drawing.Color]::FromArgb(31, 35, 41) }
    }

    $Button.BackColor = $disabledBackColor
    $Button.ForeColor = $clrTextSoft
    $Button.FlatAppearance.BorderColor = (_BlendColor $disabledBackColor 10)
}

function _Panel {
    param([int]$X, [int]$Y, [int]$W, [int]$H, $BG = $null)
    $p             = [System.Windows.Forms.Panel]::new()
    $p.Location    = [System.Drawing.Point]::new($X, $Y)
    $p.Size        = [System.Drawing.Size]::new($W, $H)
    $p.BackColor   = if ($null -ne $BG -and $BG -is [System.Drawing.Color]) { [System.Drawing.Color]$BG } else { (_ResolveUiColor -Color $clrPanel -Fallback ([System.Drawing.Color]::FromArgb(24, 28, 36))) }
    $p.BorderStyle = 'FixedSingle'
    $p.Padding     = [System.Windows.Forms.Padding]::new(0)
    return $p
}

function _VerticalText {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $Text }
    return ($Text.ToCharArray() -join "`n")
}

function _CopyLogText {
    param([System.Windows.Forms.RichTextBox]$Box)
    if ($null -eq $Box) { return $false }
    $text = ''
    try { $text = $Box.Text } catch { $text = '' }
    if ([string]::IsNullOrWhiteSpace($text)) { return $false }
    try {
        [System.Windows.Forms.Clipboard]::SetText($text)
        return $true
    } catch {
        return $false
    }
}

# Ensure ProfileManager helpers are available when needed
function _EnsureProfileManagerLoaded {
    $pmPath = Join-Path $script:ModuleRoot 'ProfileManager.psm1'
    if (Test-Path $pmPath) { Import-Module $pmPath -Force }
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
    foreach ($key in @('BotToken','WebhookUrl','MonitorChannelId','CommandPrefix','PollIntervalSeconds','EnableDebugLogging','EnablePerformanceDebugMode')) {
        if (-not $Settings.ContainsKey($key)) { $Settings[$key] = '' }
    }

    $Tab.BackColor = $clrBg
    $y = 18
    $settingsMargin = 18
    $getSettingsContentWidth = {
        return [Math]::Max(320, $Tab.ClientSize.Width - ($settingsMargin * 2))
    }.GetNewClosure()
    $measureSettingsLabelHeight = {
        param(
            [string]$Text,
            [int]$Width,
            [System.Drawing.Font]$Font,
            [int]$MinHeight = 20
        )

        $safeWidth = [Math]::Max(120, $Width)
        $measured = [System.Windows.Forms.TextRenderer]::MeasureText(
            [string]$Text,
            $Font,
            [System.Drawing.Size]::new($safeWidth, 0),
            [System.Windows.Forms.TextFormatFlags]::WordBreak
        )
        return [Math]::Max($MinHeight, $measured.Height + 6)
    }.GetNewClosure()

    $settingsToolTip = New-Object System.Windows.Forms.ToolTip
    $settingsToolTip.AutoPopDelay = 12000
    $settingsToolTip.InitialDelay = 350
    $settingsToolTip.ReshowDelay  = 150
    $settingsToolTip.ShowAlways   = $true

    $setSettingsTip = {
        param(
            [string]$Text,
            [System.Windows.Forms.Control[]]$Targets
        )

        if ([string]::IsNullOrWhiteSpace($Text)) { return }
        foreach ($target in @($Targets)) {
            if ($target) {
                try { $settingsToolTip.SetToolTip($target, $Text.Trim()) } catch { }
            }
        }
    }.GetNewClosure()

    $lblSettingsTitle = _Label 'Bot Settings' 18 $y (& $getSettingsContentWidth) 28 $fontTitle
    $lblSettingsTitle.Anchor = 'Top,Left,Right'
    $Tab.Controls.Add($lblSettingsTitle)
    $y += 38

    $lblToken = _Label 'Bot Token (keep secret)' 18 $y (& $getSettingsContentWidth) 20
    $lblToken.Anchor = 'Top,Left,Right'
    $Tab.Controls.Add($lblToken)
    $y += 20
    $tbToken = _TextBox 18 $y 480 24 ($Settings.BotToken) $true
    $tbToken.Anchor = 'Top,Left,Right'
    $Tab.Controls.Add($tbToken)
    & $setSettingsTip 'Discord bot token for this app. ECC uses it to sign in as your bot. Keep it private.' @($lblToken, $tbToken)
    $y += 32

    $lblWebhook = _Label 'Webhook URL' 18 $y (& $getSettingsContentWidth) 20
    $lblWebhook.Anchor = 'Top,Left,Right'
    $Tab.Controls.Add($lblWebhook)
    $y += 20
    $tbWebhook = _TextBox 18 $y 480 24 ($Settings.WebhookUrl)
    $tbWebhook.Anchor = 'Top,Left,Right'
    $Tab.Controls.Add($tbWebhook)
    & $setSettingsTip 'Discord webhook ECC uses for outgoing status posts. Paste the full webhook URL here.' @($lblWebhook, $tbWebhook)
    $y += 32

    $lblChannel = _Label 'Monitor Channel ID' 18 $y (& $getSettingsContentWidth) 20
    $lblChannel.Anchor = 'Top,Left,Right'
    $Tab.Controls.Add($lblChannel)
    $y += 20
    $tbChannel = _TextBox 18 $y 240 24 ($Settings.MonitorChannelId)
    $tbChannel.Anchor = 'Top,Left,Right'
    $Tab.Controls.Add($tbChannel)
    & $setSettingsTip 'Discord channel ID ECC watches for commands. This should be the channel where you talk to the bot.' @($lblChannel, $tbChannel)
    $y += 32

    $lblPrefix = _Label 'Command Prefix (default: !)' 18 $y (& $getSettingsContentWidth) 20
    $lblPrefix.Anchor = 'Top,Left,Right'
    $Tab.Controls.Add($lblPrefix)
    $y += 20
    $tbPrefix = _TextBox 18 $y 80 24 ($Settings.CommandPrefix)
    $Tab.Controls.Add($tbPrefix)
    & $setSettingsTip 'The symbol that starts bot commands, like !status or !pzstart. Keep this short.' @($lblPrefix, $tbPrefix)
    $y += 32

    $lblPoll = _Label 'Poll Interval in seconds (default: 2)' 18 $y (& $getSettingsContentWidth) 20
    $lblPoll.Anchor = 'Top,Left,Right'
    $Tab.Controls.Add($lblPoll)
    $y += 20
    $tbPoll = _TextBox 18 $y 80 24 ([string]$Settings.PollIntervalSeconds)
    $Tab.Controls.Add($tbPoll)
    & $setSettingsTip 'How often ECC checks bot and server state. Lower is faster updates. Higher uses less work.' @($lblPoll, $tbPoll)
    $y += 32

    $lblDebug = _Label 'Debug Logging' 18 $y (& $getSettingsContentWidth) 20
    $lblDebug.Anchor = 'Top,Left,Right'
    $Tab.Controls.Add($lblDebug)
    $y += 20
    $chkDebug           = New-Object System.Windows.Forms.CheckBox
    $chkDebug.Location  = [System.Drawing.Point]::new(18, $y)
    $chkDebug.Size      = [System.Drawing.Size]::new(200, 20)
    $chkDebug.Text      = 'Enabled'
    $chkDebug.ForeColor = $clrText
    $chkDebug.BackColor = [System.Drawing.Color]::Transparent
    $chkDebug.Font      = $fontLabel
    $chkDebug.Checked   = ($Settings.EnableDebugLogging -eq $true)
    $Tab.Controls.Add($chkDebug)
    & $setSettingsTip 'Turns on detailed debug logs for troubleshooting. Saving this change restarts ECC.' @($lblDebug, $chkDebug)
    $y += 30

    $lblPerf = _Label 'Performance Trace Mode' 18 $y (& $getSettingsContentWidth) 20
    $lblPerf.Anchor = 'Top,Left,Right'
    $Tab.Controls.Add($lblPerf)
    $y += 20
    $chkPerfTrace           = New-Object System.Windows.Forms.CheckBox
    $chkPerfTrace.Location  = [System.Drawing.Point]::new(18, $y)
    $chkPerfTrace.Size      = [System.Drawing.Size]::new(260, 20)
    $chkPerfTrace.Text      = 'Enable long-run perf tracing'
    $chkPerfTrace.ForeColor = $clrText
    $chkPerfTrace.BackColor = [System.Drawing.Color]::Transparent
    $chkPerfTrace.Font      = $fontLabel
    $chkPerfTrace.Checked   = ($Settings.EnablePerformanceDebugMode -eq $true)
    $Tab.Controls.Add($chkPerfTrace)
    $y += 22
    $perfHint = _Label 'Keeps normal logs quieter than full debug mode while still enabling detailed UIPERF and LOGPERF traces for long-running lag checks.' 18 $y 620 34
    $perfHint.ForeColor = $clrTextSoft
    $perfHint.AutoSize = $false
    $perfHint.Anchor = 'Top,Left,Right'
    $perfHint.Width = (& $getSettingsContentWidth)
    $perfHint.Height = & $measureSettingsLabelHeight $perfHint.Text $perfHint.Width $perfHint.Font 34
    $Tab.Controls.Add($perfHint)
    & $setSettingsTip 'Keeps normal logging lighter than full debug mode, but still records slow UI and log work so you can track lag over time.' @($lblPerf, $chkPerfTrace, $perfHint)
    $y += ($perfHint.Height + 10)

    # ── Auto-Save section ─────────────────────────────────────────────────────
    $lblAutoSaveHeader = _Label 'Auto-Save Settings' 18 $y (& $getSettingsContentWidth) 22 $fontBold
    $lblAutoSaveHeader.Anchor = 'Top,Left,Right'
    $Tab.Controls.Add($lblAutoSaveHeader)
    $y += 24

    $chkAutoSave           = New-Object System.Windows.Forms.CheckBox
    $chkAutoSave.Location  = [System.Drawing.Point]::new(18, $y)
    $chkAutoSave.Size      = [System.Drawing.Size]::new(200, 20)
    $chkAutoSave.Text      = 'Enable auto-save for all games'
    $chkAutoSave.ForeColor = $clrText
    $chkAutoSave.BackColor = [System.Drawing.Color]::Transparent
    $chkAutoSave.Font      = $fontLabel
    $chkAutoSave.Checked   = ($Settings.AutoSaveEnabled -ne $false)
    $Tab.Controls.Add($chkAutoSave)
    & $setSettingsTip 'Lets ECC run save commands for supported servers on a shared schedule.' @($lblAutoSaveHeader, $chkAutoSave)
    $y += 24

    $lblAutoSave = _Label 'Auto-save interval in minutes (default: 30)' 18 $y (& $getSettingsContentWidth) 20
    $lblAutoSave.Anchor = 'Top,Left,Right'
    $Tab.Controls.Add($lblAutoSave)
    $y += 20
    $tbAutoSave = _TextBox 18 $y 80 24 ([string]$(if ($Settings.AutoSaveIntervalMinutes) { $Settings.AutoSaveIntervalMinutes } else { '30' }))
    $Tab.Controls.Add($tbAutoSave)
    & $setSettingsTip 'Minutes between automatic save passes when auto-save is on for the app.' @($lblAutoSave, $tbAutoSave)
    $y += 30

    # ── Scheduled Restart section ──────────────────────────────────────────────
    $lblSchedHeader = _Label 'Scheduled Restart Settings' 18 $y (& $getSettingsContentWidth) 22 $fontBold
    $lblSchedHeader.Anchor = 'Top,Left,Right'
    $Tab.Controls.Add($lblSchedHeader)
    $y += 24

    $chkSchedRestart           = New-Object System.Windows.Forms.CheckBox
    $chkSchedRestart.Location  = [System.Drawing.Point]::new(18, $y)
    $chkSchedRestart.Size      = [System.Drawing.Size]::new(240, 20)
    $chkSchedRestart.Text      = 'Enable scheduled restarts for all games'
    $chkSchedRestart.ForeColor = $clrText
    $chkSchedRestart.BackColor = [System.Drawing.Color]::Transparent
    $chkSchedRestart.Font      = $fontLabel
    $chkSchedRestart.Checked   = ($Settings.ScheduledRestartEnabled -ne $false)
    $Tab.Controls.Add($chkSchedRestart)
    & $setSettingsTip 'Lets ECC restart supported servers on a shared app-wide schedule.' @($lblSchedHeader, $chkSchedRestart)
    $y += 24

    $lblSchedHours = _Label 'Restart interval in hours (default: 6)' 18 $y (& $getSettingsContentWidth) 20
    $lblSchedHours.Anchor = 'Top,Left,Right'
    $Tab.Controls.Add($lblSchedHours)
    $y += 20
    $tbSchedHours = _TextBox 18 $y 80 24 ([string]$(if ($Settings.ScheduledRestartHours) { $Settings.ScheduledRestartHours } else { '6' }))
    $Tab.Controls.Add($tbSchedHours)
    $lblSchedWarn = _Label 'Restart warnings are sent 60, 30, 15, 10, 5, 2, and 1 minute before restart.' 104 ($y + 4) ([Math]::Max(220, (& $getSettingsContentWidth) - 86)) 20
    $lblSchedWarn.ForeColor = $clrTextSoft
    $lblSchedWarn.AutoSize = $false
    $lblSchedWarn.Anchor = 'Top,Left,Right'
    $lblSchedWarn.Height = & $measureSettingsLabelHeight $lblSchedWarn.Text $lblSchedWarn.Width $lblSchedWarn.Font 20
    $Tab.Controls.Add($lblSchedWarn)
    & $setSettingsTip 'Hours between scheduled restart cycles. ECC warns players before each restart using the times shown below.' @($lblSchedHours, $tbSchedHours, $lblSchedWarn)
    $y += [Math]::Max(44, $lblSchedWarn.Height + 16)

    # Store textboxes in script scope so the click closure can access them
    $script:SettingsTabToken      = $tbToken
    $script:SettingsTabWebhook    = $tbWebhook
    $script:SettingsTabChannel    = $tbChannel
    $script:SettingsTabPrefix     = $tbPrefix
    $script:SettingsTabPoll       = $tbPoll
    $script:SettingsTabDebug      = $chkDebug
    $script:SettingsTabPerfDebug  = $chkPerfTrace
    $script:SettingsTabAutoSaveOn = $chkAutoSave
    $script:SettingsTabAutoSaveMin = $tbAutoSave
    $script:SettingsTabSchedRestartOn    = $chkSchedRestart
    $script:SettingsTabSchedRestartHours = $tbSchedHours

    $saveActionsRow = New-Object System.Windows.Forms.FlowLayoutPanel
    $saveActionsRow.Location = [System.Drawing.Point]::new(18, $y)
    $saveActionsRow.Size = [System.Drawing.Size]::new((& $getSettingsContentWidth), 34)
    $saveActionsRow.Anchor = 'Top,Left,Right'
    $saveActionsRow.WrapContents = $false
    $saveActionsRow.AutoScroll = $true
    $saveActionsRow.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
    $saveActionsRow.BackColor = [System.Drawing.Color]::Transparent
    $Tab.Controls.Add($saveActionsRow)

    $saveBtn = _Button 'Save Settings' 0 0 140 32 $clrGreen {
        $cfgPath = $script:SettingsConfigPath
        $stngs   = $script:SettingsData
        $st      = $script:SharedState

        if ($null -eq $st)                             { [System.Windows.Forms.MessageBox]::Show('SharedState is null','Error','OK','Error') | Out-Null; return }
        if ([string]::IsNullOrWhiteSpace($cfgPath))    { [System.Windows.Forms.MessageBox]::Show("ConfigPath is empty: '$cfgPath'",'Error','OK','Error') | Out-Null; return }
        if ($null -eq $stngs)                          { $stngs = @{}; $st['Settings'] = $stngs }

        $wasDebug = $false
        $wasPerfDebug = $false
        try { $wasDebug = [bool]$stngs['EnableDebugLogging'] } catch { $wasDebug = $false }
        try { $wasPerfDebug = [bool]$stngs['EnablePerformanceDebugMode'] } catch { $wasPerfDebug = $false }

        try {
            $stngs['BotToken']         = $script:SettingsTabToken.Text.Trim()
            $stngs['WebhookUrl']       = $script:SettingsTabWebhook.Text.Trim()
            $stngs['MonitorChannelId'] = $script:SettingsTabChannel.Text.Trim()
            $stngs['CommandPrefix']    = if ($script:SettingsTabPrefix.Text.Trim()) { $script:SettingsTabPrefix.Text.Trim() } else { '!' }
            $intVal = 0
            $stngs['PollIntervalSeconds'] = if ([int]::TryParse($script:SettingsTabPoll.Text, [ref]$intVal)) { $intVal } else { 2 }
            $stngs['EnableDebugLogging']   = ($script:SettingsTabDebug.Checked -eq $true)
            $stngs['EnablePerformanceDebugMode'] = ($script:SettingsTabPerfDebug.Checked -eq $true)

            # Auto-save
            $stngs['AutoSaveEnabled'] = ($script:SettingsTabAutoSaveOn.Checked -eq $true)
            $asMin = 30
            if ([int]::TryParse($script:SettingsTabAutoSaveMin.Text, [ref]$asMin) -and $asMin -gt 0) {
                $stngs['AutoSaveIntervalMinutes'] = $asMin
            } else {
                $stngs['AutoSaveIntervalMinutes'] = 30
            }

            # Scheduled restart
            $stngs['ScheduledRestartEnabled'] = ($script:SettingsTabSchedRestartOn.Checked -eq $true)
            $srHours = 6
            $srDouble = 0.0
            if ([double]::TryParse($script:SettingsTabSchedRestartHours.Text, [ref]$srDouble) -and $srDouble -gt 0) {
                $stngs['ScheduledRestartHours'] = $srDouble
            } else {
                $stngs['ScheduledRestartHours'] = 6
            }
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
        $perfChanged = ($wasPerfDebug -ne [bool]$stngs['EnablePerformanceDebugMode'])
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
        $savedMessage = 'Settings saved. The Discord listener will reconnect automatically.'
        if ($perfChanged -and -not $debugChanged) {
            $savedMessage = 'Settings saved. Performance trace mode applies right away, and the Discord listener will reconnect automatically.'
        }
        [System.Windows.Forms.MessageBox]::Show(
            $savedMessage,
            'Saved','OK','Information') | Out-Null
    }
    $saveBtn.Margin = [System.Windows.Forms.Padding]::new(0)
    $saveActionsRow.Controls.Add($saveBtn)
    & $setSettingsTip 'Save these app settings to disk. The Discord listener reconnects after save. Debug mode changes restart ECC.' @($saveBtn)

    $y += 42
    $info = "How to get these values:" + [Environment]::NewLine +
            "  Bot Token   : discord.com/developers -> Your App -> Bot -> Reset Token" + [Environment]::NewLine +
            "                Also enable 'Message Content Intent' on the Bot page." + [Environment]::NewLine +
            "  Webhook URL : Channel Settings -> Integrations -> Webhooks -> New Webhook" + [Environment]::NewLine +
            "  Channel ID  : Discord Settings -> Advanced -> enable Developer Mode" + [Environment]::NewLine +
            "                Then right-click your command channel -> Copy Channel ID"

    $note      = _Label $info 18 $y 600 120
    $note.Font = $fontLabel
    $note.AutoSize = $false
    $note.Anchor = 'Top,Left,Right'
    $note.Width = (& $getSettingsContentWidth)
    $note.Height = & $measureSettingsLabelHeight $note.Text $note.Width $note.Font 120
    $note.Name = 'SettingsNote'
    $Tab.Controls.Add($note)
    & $setSettingsTip 'Quick lookup help for the Discord values shown above.' @($note)

    # Keep the settings content sized and reflowed to the tab width so text doesn't clip
    $settingsTabLocal = $Tab
    $layoutSettingsTab = {
        $contentWidth = [Math]::Max(340, $settingsTabLocal.ClientSize.Width - 36)

        foreach ($ctrl in @(
            $lblSettingsTitle, $lblToken, $lblWebhook, $lblChannel, $lblPrefix, $lblPoll,
            $lblDebug, $lblPerf, $lblAutoSaveHeader, $lblAutoSave, $lblSchedHeader, $lblSchedHours
        )) {
            if ($ctrl -is [System.Windows.Forms.Control]) {
                $ctrl.Width = $contentWidth
            }
        }

        foreach ($tb in @($tbToken, $tbWebhook, $tbChannel)) {
            if ($tb -is [System.Windows.Forms.Control]) {
                $tb.Width = $contentWidth
            }
        }

        if ($perfHint -is [System.Windows.Forms.Control]) {
            $perfHint.Width = $contentWidth
            $perfHint.Height = [System.Windows.Forms.TextRenderer]::MeasureText(
                [string]$perfHint.Text,
                $perfHint.Font,
                [System.Drawing.Size]::new([Math]::Max(120, $contentWidth), 0),
                [System.Windows.Forms.TextFormatFlags]::WordBreak
            ).Height + 6
            $perfHint.Location = [System.Drawing.Point]::new(18, $chkPerfTrace.Bottom + 4)
        }

        if ($lblAutoSaveHeader -is [System.Windows.Forms.Control]) {
            $lblAutoSaveHeader.Location = [System.Drawing.Point]::new(18, $perfHint.Bottom + 10)
        }
        if ($chkAutoSave -is [System.Windows.Forms.Control]) {
            $chkAutoSave.Location = [System.Drawing.Point]::new(18, $lblAutoSaveHeader.Bottom + 6)
        }
        if ($lblAutoSave -is [System.Windows.Forms.Control]) {
            $lblAutoSave.Location = [System.Drawing.Point]::new(18, $chkAutoSave.Bottom + 6)
        }
        if ($tbAutoSave -is [System.Windows.Forms.Control]) {
            $tbAutoSave.Location = [System.Drawing.Point]::new(18, $lblAutoSave.Bottom + 2)
        }

        if ($lblSchedHeader -is [System.Windows.Forms.Control]) {
            $lblSchedHeader.Location = [System.Drawing.Point]::new(18, $tbAutoSave.Bottom + 10)
        }
        if ($chkSchedRestart -is [System.Windows.Forms.Control]) {
            $chkSchedRestart.Location = [System.Drawing.Point]::new(18, $lblSchedHeader.Bottom + 6)
        }
        if ($lblSchedHours -is [System.Windows.Forms.Control]) {
            $lblSchedHours.Location = [System.Drawing.Point]::new(18, $chkSchedRestart.Bottom + 6)
        }
        if ($tbSchedHours -is [System.Windows.Forms.Control]) {
            $tbSchedHours.Location = [System.Drawing.Point]::new(18, $lblSchedHours.Bottom + 2)
        }
        if ($lblSchedWarn -is [System.Windows.Forms.Control]) {
            $lblSchedWarn.Location = [System.Drawing.Point]::new(104, $tbSchedHours.Top + 4)
            $lblSchedWarn.Width = [Math]::Max(220, $contentWidth - 86)
            $lblSchedWarn.Height = [System.Windows.Forms.TextRenderer]::MeasureText(
                [string]$lblSchedWarn.Text,
                $lblSchedWarn.Font,
                [System.Drawing.Size]::new([Math]::Max(120, $lblSchedWarn.Width), 0),
                [System.Windows.Forms.TextFormatFlags]::WordBreak
            ).Height + 6
        }

        if ($saveActionsRow -is [System.Windows.Forms.Control]) {
            $saveActionsRow.Location = [System.Drawing.Point]::new(18, [Math]::Max($tbSchedHours.Bottom + 18, $lblSchedWarn.Bottom + 14))
            $saveActionsRow.Width = $contentWidth
        }

        if ($note -is [System.Windows.Forms.Control]) {
            $note.Location = [System.Drawing.Point]::new(18, $saveActionsRow.Bottom + 14)
            $note.Width = $contentWidth
            $note.Height = [System.Windows.Forms.TextRenderer]::MeasureText(
                [string]$note.Text,
                $note.Font,
                [System.Drawing.Size]::new([Math]::Max(120, $contentWidth), 0),
                [System.Windows.Forms.TextFormatFlags]::WordBreak
            ).Height + 6
        }

        try {
            $settingsTabLocal.AutoScrollMinSize = [System.Drawing.Size]::new(0, $note.Bottom + 18)
        } catch { }
    }.GetNewClosure()
    $Tab.add_Resize($layoutSettingsTab)
    & $layoutSettingsTab
}

# =============================================================================
#  SCRIPT-SCOPE SHARED STATE  (set once in Start-GUI, read everywhere)
# =============================================================================
$script:SharedState = $null
$script:CommandCatalogPath = $null
$script:CommandCatalog     = $null
$script:_UIReloadRequested = $false
$script:_MainToolTip       = $null

function _ResolveMainControlToolTip {
    param([System.Windows.Forms.Control]$Control)

    if ($null -eq $Control) { return '' }

    $name = ''
    $text = ''
    try { $name = [string]$Control.Name } catch { $name = '' }
    try { $text = [string]$Control.Text } catch { $text = '' }

    switch ($name) {
        'btnWinMin'   { return 'Minimize Etherium Command Center to the taskbar.' }
        'btnWinMax'   { return 'Toggle the main window between normal and maximized size.' }
        'btnWinClose' { return 'Close Etherium Command Center.' }
    }

    switch -Regex ($text) {
        '^\+ Add Game$'     { return 'Create a new server profile and add it to the dashboard.' }
        '^Remove$'          { return 'Remove the selected server profile from ECC.' }
        '^Reload UI$'       { return 'Rebuild the ECC window without resetting running servers or timers.' }
        '^Reload Bot$'      { return 'Reconnect the Discord bot without touching running servers.' }
        '^Reload Commands$' { return 'Reload profiles and command catalog files from disk without restarting ECC.' }
        '^Full Restart$'    { return 'Restart the full app. Running servers will be stopped first.' }
        '^Settings$'        { return 'Open ECC settings for the bot, auto-save, and restart behavior.' }
        '^Send$'            { return 'Send the current message or command.' }
        '^Start$'           { return 'Start this server.' }
        '^Stop$'            { return 'Stop this server using its configured shutdown path.' }
        '^Restart$'         { return 'Restart this server using its configured save and stop rules.' }
        '^Commands$'        { return 'Open the command tools window for this server.' }
        '^Config$'          { return 'Open the detected config files for this server.' }
        '^Manager$'         { return 'Open the Hytale manager for updater, downloader, and mod tools.' }
        '^Auto-Restart$'    { return 'Allow ECC to restart this server automatically after a crash.' }
        '^Save Profile$'    { return 'Save changes made to the selected server profile.' }
        '^Restart Server$'  { return 'Restart the selected server from the profile editor.' }
        '^Stop Server$'     { return 'Stop the selected server from the profile editor.' }
    }

    return ''
}

function _SetMainControlToolTip {
    param(
        [System.Windows.Forms.Control]$Control,
        [string]$Text = ''
    )

    if ($null -eq $Control -or $null -eq $script:_MainToolTip) { return }

    $tipText = if ([string]::IsNullOrWhiteSpace($Text)) {
        _ResolveMainControlToolTip -Control $Control
    } else {
        $Text
    }

    if ([string]::IsNullOrWhiteSpace($tipText)) { return }
    try { $script:_MainToolTip.SetToolTip($Control, $tipText.Trim()) } catch { }
}

# =============================================================================
#  MAIN GUI ENTRY POINT
# =============================================================================
function Start-GUI {
    param(
        [hashtable]$SharedState,
        [string]$ConfigPath      = '.\Config\Settings.json',
        [string]$ProfilesDir     = '.\Profiles',
        [string]$TranscriptPath  = ''
    )

    trap {
        try {
            $trapNow = Get-Date
            $trapMessage = $_.Exception.Message
            $trapPosition = ''
            $trapStack = ''
            try { $trapPosition = [string]$_.InvocationInfo.PositionMessage } catch { $trapPosition = '' }
            try { $trapStack = [string]$_.ScriptStackTrace } catch { $trapStack = '' }
            $trapLines = New-Object System.Collections.Generic.List[string]
            $trapLines.Add("[$($trapNow.ToString('yyyy-MM-dd HH:mm:ss'))][ERROR][GUI] Start-GUI fatal: $trapMessage") | Out-Null
            if (-not [string]::IsNullOrWhiteSpace($trapPosition)) {
                $trapLines.Add("[$($trapNow.ToString('yyyy-MM-dd HH:mm:ss'))][ERROR][GUI] Start-GUI position: $trapPosition") | Out-Null
            }
            if (-not [string]::IsNullOrWhiteSpace($trapStack)) {
                $trapLines.Add("[$($trapNow.ToString('yyyy-MM-dd HH:mm:ss'))][ERROR][GUI] Start-GUI stack: $trapStack") | Out-Null
            }

            foreach ($trapLine in $trapLines) {
                try { _GuiCrashBreadcrumb -Message ($trapLine -replace '^\[[^\]]+\]\[ERROR\]\[GUI\]\s*', '') -Level ERROR } catch { }
                try {
                    $writeLog = Get-Command -Name 'Write-Log' -ErrorAction SilentlyContinue
                    if ($writeLog) {
                        Write-Log -Message ($trapLine -replace '^\[[^\]]+\]\[ERROR\]\[GUI\]\s*', '') -Level ERROR -Source 'GUI'
                    }
                } catch { }
                try {
                    if ($script:SharedState -and $script:SharedState.LogQueue) {
                        $script:SharedState.LogQueue.Enqueue($trapLine)
                    }
                } catch { }
                try {
                    $fallbackPath = Join-Path $PSScriptRoot '..\Logs\gui_startup_trap.log'
                    Add-Content -Path $fallbackPath -Value $trapLine -ErrorAction SilentlyContinue
                } catch { }
            }
        } catch { }
        throw
    }

    if ($null -eq $SharedState -or -not ($SharedState -is [hashtable])) {
        throw "Start-GUI was called without a valid SharedState hashtable."
    }

    $script:SharedState = $SharedState
    $script:ProfilesDir = $ProfilesDir
    $guiStartedAt = Get-Date
    _GuiCrashBreadcrumb -Message 'Start-GUI entered.'
    if (-not $script:_SelectedProfilePrefix) { $script:_SelectedProfilePrefix = $null }
    try {
        $cfgDir = Split-Path -Path $ConfigPath -Parent
        if (-not [string]::IsNullOrWhiteSpace($cfgDir)) {
            $script:CommandCatalogPath = Join-Path $cfgDir 'CommandCatalog.json'
        }
    } catch { }

    # Ensure keys that the timer reads always exist
    if (-not $SharedState.ContainsKey('ListenerRunning')) { $SharedState['ListenerRunning'] = $false }
    if (-not $SharedState.ContainsKey('StopListener'))    { $SharedState['StopListener']    = $false }
    if (-not $SharedState.ContainsKey('StopMonitor'))     { $SharedState['StopMonitor']     = $false }
    if (-not $SharedState.ContainsKey('StopMetricsWorker')) { $SharedState['StopMetricsWorker'] = $false } else { $SharedState['StopMetricsWorker'] = $false }
    if (-not $SharedState.ContainsKey('StopLogTailWorker')) { $SharedState['StopLogTailWorker'] = $false } else { $SharedState['StopLogTailWorker'] = $false }
    if (-not $SharedState.ContainsKey('LastGuiTimerErrorAt')) { $SharedState['LastGuiTimerErrorAt'] = $null }
    if (-not $SharedState.ContainsKey('GameLogQueue'))    { $SharedState['GameLogQueue']    = [System.Collections.Concurrent.ConcurrentQueue[object]]::new() }
if (-not $SharedState.ContainsKey('PlayersRequests')) { $SharedState['PlayersRequests'] = [hashtable]::Synchronized(@{}) }
if (-not $SharedState.ContainsKey('LatestPlayers'))   { $SharedState['LatestPlayers']   = [hashtable]::Synchronized(@{}) }
if (-not $SharedState.ContainsKey('LatestPlayerCounts')) { $SharedState['LatestPlayerCounts'] = [hashtable]::Synchronized(@{}) }
if (-not $SharedState.ContainsKey('LatestPlayerObservedAt')) { $SharedState['LatestPlayerObservedAt'] = [hashtable]::Synchronized(@{}) }
if (-not $SharedState.ContainsKey('PzObservedPlayerIds')) { $SharedState['PzObservedPlayerIds'] = [hashtable]::Synchronized(@{}) }
if (-not $SharedState.ContainsKey('ServerStartNotified')) { $SharedState['ServerStartNotified'] = [hashtable]::Synchronized(@{}) }
    if (-not $SharedState.ContainsKey('LastUserActivityAt')) { $SharedState['LastUserActivityAt'] = Get-Date }
    if (-not $SharedState.ContainsKey('LastDashboardRefreshAt')) { $SharedState['LastDashboardRefreshAt'] = [datetime]::MinValue }
    if (-not $SharedState.ContainsKey('GuiRefreshMode')) { $SharedState['GuiRefreshMode'] = 'active' }
    if (-not $SharedState.ContainsKey('GuiRefreshModeReason')) { $SharedState['GuiRefreshModeReason'] = 'startup' }
    if (-not $SharedState.ContainsKey('GuiIdleAfterSeconds')) { $SharedState['GuiIdleAfterSeconds'] = 30 }
    if (-not $SharedState.ContainsKey('GuiIdleDashboardRefreshSeconds')) { $SharedState['GuiIdleDashboardRefreshSeconds'] = 10 }
    if (-not $SharedState.ContainsKey('GuiActiveDashboardRefreshSeconds')) { $SharedState['GuiActiveDashboardRefreshSeconds'] = 2 }

    if (-not $SharedState.ContainsKey('GuiExceptionHooksInstalled')) {
        $SharedState['GuiExceptionHooksInstalled'] = $false
    }
    if (-not [bool]$SharedState['GuiExceptionHooksInstalled']) {
        try {
            [System.Windows.Forms.Application]::SetUnhandledExceptionMode([System.Windows.Forms.UnhandledExceptionMode]::CatchException)
        } catch { }

        try {
            $threadExceptionHandler = [System.Threading.ThreadExceptionEventHandler]{
                param($sender, $args)
                try {
                    if ($args -and $args.Exception) {
                        $msg = $args.Exception.ToString()
                        _GuiCrashBreadcrumb -Message ("WinForms ThreadException: {0}" -f $msg) -Level ERROR
                        _GuiDirectLog -Message ("WinForms ThreadException: {0}" -f $msg) -Level ERROR
                    } else {
                        if (-not (_ShouldLogEmptyThreadExceptionEvent)) { return }
                        $senderType = ''
                        $argsType = ''
                        try { if ($sender) { $senderType = $sender.GetType().FullName } } catch { $senderType = '' }
                        try { if ($args) { $argsType = $args.GetType().FullName } } catch { $argsType = '' }
                        $msg = "WinForms ThreadException fired with no Exception payload. senderType=$senderType argsType=$argsType"
                        _GuiCrashBreadcrumb -Message $msg -Level DEBUG
                        _GuiDirectLog -Message $msg -Level DEBUG
                    }
                } catch { }
            }
            [System.Windows.Forms.Application]::add_ThreadException($threadExceptionHandler)
            $SharedState['GuiThreadExceptionHandler'] = $threadExceptionHandler
        } catch {
            _GuiCrashBreadcrumb -Message ("Failed to register WinForms ThreadException hook: {0}" -f $_.Exception.Message) -Level WARN
        }

        try {
            $sharedDomainHandler = [System.UnhandledExceptionEventHandler]{
                param($sender, $args)
                try {
                    $ex = $args.ExceptionObject
                    $isTerminating = $false
                    try { $isTerminating = [bool]$args.IsTerminating } catch { }
                    $msg = if ($ex -is [System.Exception]) {
                        $ex.ToString()
                    } else {
                        [string]$ex
                    }
                    _GuiCrashBreadcrumb -Message ("GUI observed AppDomain exception. terminating={0} :: {1}" -f $isTerminating, $msg) -Level ERROR
                } catch { }
            }
            [System.AppDomain]::CurrentDomain.add_UnhandledException($sharedDomainHandler)
            $SharedState['GuiAppDomainExceptionHandler'] = $sharedDomainHandler
        } catch {
            _GuiCrashBreadcrumb -Message ("Failed to register GUI AppDomain hook: {0}" -f $_.Exception.Message) -Level WARN
        }

        $SharedState['GuiExceptionHooksInstalled'] = $true
        _GuiCrashBreadcrumb -Message 'Registered GUI exception hooks.'
    }
if (-not $SharedState.ContainsKey('SatisfactoryConnectionCapture')) { $SharedState['SatisfactoryConnectionCapture'] = [hashtable]::Synchronized(@{}) }
if (-not $SharedState.ContainsKey('ValheimPlayerCapture')) { $SharedState['ValheimPlayerCapture'] = [hashtable]::Synchronized(@{}) }

    function _MarkGuiActivity {
        param([string]$Reason = 'user')

        try {
            $script:SharedState['LastUserActivityAt'] = Get-Date
            if ($script:SharedState.ContainsKey('GuiRefreshMode') -and [string]$script:SharedState['GuiRefreshMode'] -ne 'active') {
                $script:SharedState['GuiRefreshMode'] = 'active'
                $script:SharedState['GuiRefreshModeReason'] = "wakeup:$Reason"
                if (_IsGuiDebugEnabled) {
                    _QueueStatusMessage ("GUI refresh mode -> active ({0})" -f $Reason)
                }
            }
        } catch { }
    }

    function _HasActiveServerWorkflow {
        try {
            if ($script:SharedState -and $script:SharedState.RunningServers -and @($script:SharedState.RunningServers.Keys).Count -gt 0) {
                return $true
            }
        } catch { }

        try {
            $activeStates = @('starting','restarting','stopping','waiting_restart','waiting_first_player','idle_wait','idle_shutdown','online')
            if ($script:SharedState -and $script:SharedState.ServerRuntimeState) {
                foreach ($stateKey in @($script:SharedState.ServerRuntimeState.Keys)) {
                    $entry = $null
                    try { $entry = $script:SharedState.ServerRuntimeState[$stateKey] } catch { $entry = $null }
                    if ($null -eq $entry) { continue }
                    $stateCode = ''
                    try { $stateCode = [string]$entry.State } catch { $stateCode = '' }
                    if ($activeStates -contains $stateCode) { return $true }
                }
            }
        } catch { }

        return $false
    }

    function _GetDashboardRefreshDecision {
        $now = Get-Date
        $idleAfterSeconds = 30
        $idleRefreshSeconds = 10
        $activeRefreshSeconds = 2
        try { $idleAfterSeconds = [Math]::Max(5, [int]$script:SharedState['GuiIdleAfterSeconds']) } catch { }
        try { $idleRefreshSeconds = [Math]::Max(4, [int]$script:SharedState['GuiIdleDashboardRefreshSeconds']) } catch { }
        try { $activeRefreshSeconds = [Math]::Max(1, [int]$script:SharedState['GuiActiveDashboardRefreshSeconds']) } catch { }

        $lastUserActivityAt = $now
        try {
            if ($script:SharedState.ContainsKey('LastUserActivityAt') -and $script:SharedState['LastUserActivityAt'] -is [datetime]) {
                $lastUserActivityAt = [datetime]$script:SharedState['LastUserActivityAt']
            }
        } catch { }

        $activeServerWorkflow = _HasActiveServerWorkflow
        $idleCandidate = $false
        $modeReason = 'server_activity'
        $secondsSinceActivity = 0.0
        try { $secondsSinceActivity = [Math]::Max(0.0, ($now - $lastUserActivityAt).TotalSeconds) } catch { $secondsSinceActivity = 0.0 }
        if (-not $activeServerWorkflow) {
            if ($secondsSinceActivity -ge $idleAfterSeconds) {
                $idleCandidate = $true
                $modeReason = ("idle_after_{0}s" -f [int][Math]::Round($secondsSinceActivity, 0))
            } else {
                $modeReason = ("recent_user_activity_{0}s" -f [int][Math]::Round($secondsSinceActivity, 0))
            }
        }

        $desiredMode = if ($idleCandidate) { 'idle' } else { 'active' }
        $intervalSeconds = if ($idleCandidate) { $idleRefreshSeconds } else { $activeRefreshSeconds }

        $lastRefreshAt = [datetime]::MinValue
        try {
            if ($script:SharedState.ContainsKey('LastDashboardRefreshAt') -and $script:SharedState['LastDashboardRefreshAt'] -is [datetime]) {
                $lastRefreshAt = [datetime]$script:SharedState['LastDashboardRefreshAt']
            }
        } catch { }

        $shouldRun = ($lastRefreshAt -eq [datetime]::MinValue)
        if (-not $shouldRun) {
            try { $shouldRun = (($now - $lastRefreshAt).TotalSeconds -ge $intervalSeconds) } catch { $shouldRun = $true }
        }

        $previousMode = ''
        try { $previousMode = [string]$script:SharedState['GuiRefreshMode'] } catch { $previousMode = '' }
        if ($previousMode -ne $desiredMode) {
            try {
                $script:SharedState['GuiRefreshMode'] = $desiredMode
                $script:SharedState['GuiRefreshModeReason'] = $modeReason
                if (_IsGuiDebugEnabled) {
                    _QueueStatusMessage ("GUI refresh mode -> {0} ({1})" -f $desiredMode, $modeReason)
                }
            } catch { }
        } else {
            try { $script:SharedState['GuiRefreshModeReason'] = $modeReason } catch { }
        }

        return @{
            Mode = $desiredMode
            ShouldRun = $shouldRun
            Reason = $modeReason
            IntervalSeconds = $intervalSeconds
            SecondsSinceActivity = $secondsSinceActivity
        }
    }

    function _RegisterGuiActivityHooks {
        param([System.Windows.Forms.Control]$Root)

        if ($null -eq $Root) { return }

        try { $Root.Add_MouseDown({ _MarkGuiActivity 'mouse' }.GetNewClosure()) } catch { }
        try { $Root.Add_MouseWheel({ _MarkGuiActivity 'wheel' }.GetNewClosure()) } catch { }
        try { $Root.Add_KeyDown({ _MarkGuiActivity 'key' }.GetNewClosure()) } catch { }

        foreach ($child in @($Root.Controls)) {
            if ($child -is [System.Windows.Forms.Control]) {
                _RegisterGuiActivityHooks -Root $child
            }
        }
    }

    # ── Log file tail state ────────────────────────────────────────────────
    # We tail the transcript file (console.log) which captures all console output.
    $script:LogFilePath    = $null
    $script:LogFilePos     = 0L
    # If a transcript path was passed directly, use it immediately
    if (-not [string]::IsNullOrEmpty($TranscriptPath)) {
        $script:LogFilePath = $TranscriptPath
        $script:LogFilePos  = 0L
    }

    # Layout constants
    $windowMargin     = 8
    $leftWidth        = 268
    $rightWidth       = 560
    $sideGap          = 18
    $topBarHeight     = 68
    $bottomLogsHeight = 250
    $defaultWidth     = 1920
    $defaultHeight    = 1080
    $minWidth         = 1600
    $minHeight        = 900
    $tabFont = New-Object System.Drawing.Font("Segoe UI", 7, [System.Drawing.FontStyle]::Regular)

    function _GetWorkingAreaRect {
        param(
            [Nullable[int]]$X = $null,
            [Nullable[int]]$Y = $null,
            [Nullable[int]]$Width = $null,
            [Nullable[int]]$Height = $null
        )

        try {
            $allScreens = @([System.Windows.Forms.Screen]::AllScreens)
            if ($allScreens.Count -gt 0) {
                if ($X.HasValue -and $Y.HasValue) {
                    $probeWidth = if ($Width.HasValue -and $Width.Value -gt 0) { $Width.Value } else { 1 }
                    $probeHeight = if ($Height.HasValue -and $Height.Value -gt 0) { $Height.Value } else { 1 }
                    $probeRect = [System.Drawing.Rectangle]::new($X.Value, $Y.Value, $probeWidth, $probeHeight)
                    $bestScreen = $null
                    $bestArea = -1
                    foreach ($screen in $allScreens) {
                        $intersection = [System.Drawing.Rectangle]::Intersect($screen.WorkingArea, $probeRect)
                        $intersectionArea = [Math]::Max(0, $intersection.Width) * [Math]::Max(0, $intersection.Height)
                        if ($intersectionArea -gt $bestArea) {
                            $bestArea = $intersectionArea
                            $bestScreen = $screen
                        }
                    }
                    if ($bestScreen) { return $bestScreen.WorkingArea }

                    foreach ($screen in $allScreens) {
                        if ($screen.Bounds.Contains([System.Drawing.Point]::new($X.Value, $Y.Value))) {
                            return $screen.WorkingArea
                        }
                    }
                }
                return $allScreens[0].WorkingArea
            }
            return [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
        } catch {
            return [System.Drawing.Rectangle]::new(0, 0, $defaultWidth, $defaultHeight)
        }
    }

    function _ClampWindowRectToWorkingArea {
        param(
            [int]$Width,
            [int]$Height,
            [Nullable[int]]$X = $null,
            [Nullable[int]]$Y = $null
        )

        $wa = _GetWorkingAreaRect -X $X -Y $Y -Width $Width -Height $Height
        $effectiveMinWidth = [Math]::Min($minWidth, $wa.Width)
        $effectiveMinHeight = [Math]::Min($minHeight, $wa.Height)
        $width = [Math]::Max($effectiveMinWidth, [Math]::Min($Width, $wa.Width))
        $height = [Math]::Max($effectiveMinHeight, [Math]::Min($Height, $wa.Height))

        $x = if ($X.HasValue) { $X.Value } else { $wa.X + [int][Math]::Floor(($wa.Width - $width) / 2) }
        $y = if ($Y.HasValue) { $Y.Value } else { $wa.Y + [int][Math]::Floor(($wa.Height - $height) / 2) }

        if ($x -lt $wa.X) { $x = $wa.X }
        if ($y -lt $wa.Y) { $y = $wa.Y }
        if (($x + $width) -gt ($wa.X + $wa.Width)) { $x = $wa.X + $wa.Width - $width }
        if (($y + $height) -gt ($wa.Y + $wa.Height)) { $y = $wa.Y + $wa.Height - $height }

        return @{
            Width = $width
            Height = $height
            X = $x
            Y = $y
        }
    }

    function _GetSavedWindowBounds {
        $settings = $script:SharedState.Settings
        if (-not $settings) { return $null }

        $w = 0; $h = 0; $x = 0; $y = 0
        $hasSize = $settings.ContainsKey('WindowWidth') -and $settings.ContainsKey('WindowHeight') -and
                   [int]::TryParse("$($settings.WindowWidth)", [ref]$w) -and
                   [int]::TryParse("$($settings.WindowHeight)", [ref]$h)
        $hasPos = $settings.ContainsKey('WindowX') -and $settings.ContainsKey('WindowY') -and
                  [int]::TryParse("$($settings.WindowX)", [ref]$x) -and
                  [int]::TryParse("$($settings.WindowY)", [ref]$y)
        if (-not $hasSize) { return $null }

        $clamped = _ClampWindowRectToWorkingArea -Width $w -Height $h -X $(if ($hasPos) { $x } else { $null }) -Y $(if ($hasPos) { $y } else { $null })

        return @{
            Width  = $clamped.Width
            Height = $clamped.Height
            X      = $clamped.X
            Y      = $clamped.Y
            HasPos = $hasPos
            State  = if ($settings.ContainsKey('WindowState')) { "$($settings.WindowState)" } else { 'Normal' }
        }
    }

    function _GetReloadWindowBounds {
        if (-not $script:SharedState) { return $null }
        if (-not $script:SharedState.ContainsKey('ReloadWindowBounds')) { return $null }

        try {
            $saved = $script:SharedState['ReloadWindowBounds']
            if ($saved -isnot [System.Collections.IDictionary]) { return $null }

            $w = 0; $h = 0; $x = 0; $y = 0
            $hasSize = $saved.Contains('Width') -and $saved.Contains('Height') -and
                       [int]::TryParse("$($saved.Width)", [ref]$w) -and
                       [int]::TryParse("$($saved.Height)", [ref]$h)
            $hasPos = $saved.Contains('X') -and $saved.Contains('Y') -and
                      [int]::TryParse("$($saved.X)", [ref]$x) -and
                      [int]::TryParse("$($saved.Y)", [ref]$y)
            if (-not $hasSize) { return $null }

            $clamped = _ClampWindowRectToWorkingArea -Width $w -Height $h -X $(if ($hasPos) { $x } else { $null }) -Y $(if ($hasPos) { $y } else { $null })

            return @{
                Width  = $clamped.Width
                Height = $clamped.Height
                X      = $clamped.X
                Y      = $clamped.Y
                HasPos = $hasPos
                State  = if ($saved.Contains('State')) { "$($saved.State)" } else { 'Normal' }
            }
        } catch {
            return $null
        }
    }

    function _PersistWindowSettings {
        if (-not $script:SharedState -or -not $script:SharedState.Settings) { return }
        if ([string]::IsNullOrWhiteSpace($ConfigPath)) { return }

        try {
            $settings = $script:SharedState.Settings
            $bounds = if ($form.WindowState -eq 'Normal') { $form.Bounds } else { $form.RestoreBounds }
            if ($bounds.Width -lt $minWidth -or $bounds.Height -lt $minHeight) { return }

            $settings['WindowWidth']  = [int]$bounds.Width
            $settings['WindowHeight'] = [int]$bounds.Height
            $settings['WindowX']      = [int]$bounds.X
            $settings['WindowY']      = [int]$bounds.Y
            $settings['WindowState']  = if ($form.WindowState -eq 'Maximized') { 'Maximized' } else { 'Normal' }

            _SaveSettings -Settings $settings -ConfigPath $ConfigPath
        } catch { }
    }

    function _OpenHytaleManagerWindow {
        param(
            [hashtable]$Profile,
            [string]$Prefix
        )

        if ($null -eq $Profile) { return }
        if ((_NormalizeGameIdentity (_GetProfileKnownGame -Profile $Profile)) -ne 'hytale') { return }

        $hytaleRoot = ''
        try { $hytaleRoot = [string]$Profile.FolderPath } catch { $hytaleRoot = '' }
        if ([string]::IsNullOrWhiteSpace($hytaleRoot)) {
            try { $hytaleRoot = [string]$Profile.ConfigRoot } catch { $hytaleRoot = '' }
        }
        if (-not [string]::IsNullOrWhiteSpace($hytaleRoot)) {
            try { $hytaleRoot = [System.IO.Path]::GetFullPath($hytaleRoot) } catch { }
        }

        $modsPath = if ($hytaleRoot) { Join-Path $hytaleRoot 'mods' } else { '' }
        $modsDisabledPath = if ($hytaleRoot) { Join-Path $hytaleRoot 'mods_disabled' } else { '' }
        $cfMetadataPath = if ($hytaleRoot) { Join-Path $hytaleRoot 'cf_mod_metadata.json' } else { '' }
        $modNotesPath = if ($hytaleRoot) { Join-Path $hytaleRoot 'mod_notes.json' } else { '' }
        $modsFeatureReady = -not [string]::IsNullOrWhiteSpace($hytaleRoot)

        $cfApiKey = '$2a$10$OAqNZqZBBGveZ8SnmJ4d6.FeLQMtC0DdkrLOSGA3RJjH1vPjWPKaK'
        $cfApiBase = 'https://api.curseforge.com/v1'
        $cfGameId = 70216
        $cfClassId = 9137
        try { Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction Stop | Out-Null } catch { }
        $modDeleteColor = [System.Drawing.Color]::FromArgb(232, 106, 106)
        $modLinkColor = [System.Drawing.Color]::FromArgb(74, 170, 255)
        $modBrowseColor = [System.Drawing.Color]::FromArgb(102, 196, 152)

        $form = New-Object System.Windows.Forms.Form
        $form.Text = "Hytale Manager - $($Profile.GameName)"
        $form.Size = [System.Drawing.Size]::new(1040, 790)
        $form.MinimumSize = [System.Drawing.Size]::new(960, 700)
        $form.StartPosition = 'CenterParent'
        $form.BackColor = $clrBg

        $toolTip = New-Object System.Windows.Forms.ToolTip
        $toolTip.ShowAlways = $true

        $setHytaleTip = {
            param(
                [string]$Text,
                [System.Windows.Forms.Control[]]$Targets
            )

            if ([string]::IsNullOrWhiteSpace($Text)) { return }
            foreach ($target in @($Targets)) {
                if ($target) {
                    try { $toolTip.SetToolTip($target, $Text.Trim()) } catch { }
                }
            }
        }.GetNewClosure()

        $lblHeader = _Label "Hytale Manager - $($Profile.GameName)" 12 10 520 24 $fontTitle
        $form.Controls.Add($lblHeader)

        $lblHint = _Label 'Use the Updater tab for server tools and the Mod Manager tab to manage local .jar mods.' 12 38 980 18 $fontLabel
        $lblHint.ForeColor = $clrTextSoft
        $lblHint.Anchor = 'Top,Left,Right'
        $form.Controls.Add($lblHint)

        $tabMain = New-Object System.Windows.Forms.TabControl
        $tabMain.Location = [System.Drawing.Point]::new(12, 70)
        $tabMain.Size = [System.Drawing.Size]::new($form.ClientSize.Width - 24, $form.ClientSize.Height - 82)
        $tabMain.Anchor = 'Top,Left,Right,Bottom'
        $form.Controls.Add($tabMain)

        $tabUpdater = New-Object System.Windows.Forms.TabPage
        $tabUpdater.Text = 'Updater'
        $tabUpdater.BackColor = $clrBg
        $tabUpdater.ForeColor = $clrText
        $tabUpdater.AutoScroll = $false
        $tabMain.TabPages.Add($tabUpdater)

        $tabMods = New-Object System.Windows.Forms.TabPage
        $tabMods.Text = 'Mod Manager'
        $tabMods.BackColor = $clrBg
        $tabMods.ForeColor = $clrText
        $tabMods.AutoScroll = $false
        $tabMain.TabPages.Add($tabMods)

        $groupUpdate = New-Object System.Windows.Forms.GroupBox
        $groupUpdate.Text = 'Hytale Server Update'
        $groupUpdate.Location = [System.Drawing.Point]::new(16, 16)
        $groupUpdate.Size = [System.Drawing.Size]::new($tabMain.ClientSize.Width - 48, 168)
        $groupUpdate.Anchor = 'Top,Left,Right'
        $groupUpdate.ForeColor = $clrText
        $tabUpdater.Controls.Add($groupUpdate)

        $lblDownloaderPath = _Label 'Downloader:' 12 24 120 18 $fontBold
        $groupUpdate.Controls.Add($lblDownloaderPath)

        $txtDownloaderPath = New-Object System.Windows.Forms.TextBox
        $txtDownloaderPath.Location = [System.Drawing.Point]::new(12, 46)
        $txtDownloaderPath.Size = [System.Drawing.Size]::new($groupUpdate.ClientSize.Width - 128, 22)
        $txtDownloaderPath.Anchor = 'Top,Left,Right'
        $txtDownloaderPath.ReadOnly = $true
        $txtDownloaderPath.BackColor = $clrPanelSoft
        $txtDownloaderPath.ForeColor = $clrText
        $txtDownloaderPath.BorderStyle = 'FixedSingle'
        $groupUpdate.Controls.Add($txtDownloaderPath)

        $btnOpenFolder = _Button 'Open Folder' ($groupUpdate.ClientSize.Width - 108) 44 96 26 $clrPanelAlt $null
        $btnOpenFolder.Anchor = 'Top,Right'
        $groupUpdate.Controls.Add($btnOpenFolder)
        & $setHytaleTip 'Open the folder that should contain the Hytale downloader and server update files.' @($lblDownloaderPath, $txtDownloaderPath, $btnOpenFolder)

        $flowUpdatePrimaryActions = New-Object System.Windows.Forms.FlowLayoutPanel
        $flowUpdatePrimaryActions.Location = [System.Drawing.Point]::new(12, 80)
        $flowUpdatePrimaryActions.Size = [System.Drawing.Size]::new($groupUpdate.ClientSize.Width - 24, 32)
        $flowUpdatePrimaryActions.Anchor = 'Top,Left,Right'
        $flowUpdatePrimaryActions.WrapContents = $true
        $flowUpdatePrimaryActions.AutoScroll = $false
        $flowUpdatePrimaryActions.BackColor = [System.Drawing.Color]::Transparent
        $groupUpdate.Controls.Add($flowUpdatePrimaryActions)

        $btnUpdateServer = _Button 'Update Server' 0 0 140 28 $clrAccent $null
        $btnUpdateServer.Margin = [System.Windows.Forms.Padding]::new(0, 0, 8, 4)
        $flowUpdatePrimaryActions.Controls.Add($btnUpdateServer)
        & $setHytaleTip 'Download and install the newest Hytale server files for this profile.' @($btnUpdateServer)

        $chkAutoRestart = New-Object System.Windows.Forms.CheckBox
        $chkAutoRestart.Text = 'Auto-restart after update'
        $chkAutoRestart.Size = [System.Drawing.Size]::new(190, 20)
        $chkAutoRestart.Checked = $true
        $chkAutoRestart.ForeColor = $clrText
        $chkAutoRestart.BackColor = [System.Drawing.Color]::Transparent
        $chkAutoRestart.Font = $fontLabel
        $chkAutoRestart.Margin = [System.Windows.Forms.Padding]::new(0, 4, 8, 0)
        $flowUpdatePrimaryActions.Controls.Add($chkAutoRestart)
        & $setHytaleTip 'After the update finishes, start the server again if it was stopped for the update.' @($chkAutoRestart)

        $lblWarn = _Label 'Checking current server state...' 12 116 ($groupUpdate.ClientSize.Width - 24) 36 $fontLabel
        $lblWarn.ForeColor = $clrYellow
        $lblWarn.Anchor = 'Top,Left,Right'
        $groupUpdate.Controls.Add($lblWarn)
        & $setHytaleTip 'Important update warning. This changes based on whether the server is running right now.' @($lblWarn)

        $groupUtils = New-Object System.Windows.Forms.GroupBox
        $groupUtils.Text = 'Version & Downloader Tools'
        $groupUtils.Location = [System.Drawing.Point]::new(16, 162)
        $groupUtils.Size = [System.Drawing.Size]::new($tabMain.ClientSize.Width - 48, 128)
        $groupUtils.Anchor = 'Top,Left,Right'
        $groupUtils.ForeColor = $clrText
        $tabUpdater.Controls.Add($groupUtils)

        $flowUpdateTools = New-Object System.Windows.Forms.FlowLayoutPanel
        $flowUpdateTools.Location = [System.Drawing.Point]::new(12, 28)
        $flowUpdateTools.Size = [System.Drawing.Size]::new($groupUtils.ClientSize.Width - 24, 88)
        $flowUpdateTools.Anchor = 'Top,Left,Right'
        $flowUpdateTools.WrapContents = $true
        $flowUpdateTools.AutoScroll = $false
        $flowUpdateTools.BackColor = [System.Drawing.Color]::Transparent
        $groupUtils.Controls.Add($flowUpdateTools)

        $btnCheckServerUpdate = _Button 'Check Server Update' 0 0 170 28 $clrPanelAlt $null
        $flowUpdateTools.Controls.Add($btnCheckServerUpdate)
        $btnDownloaderVersion = _Button 'Downloader Version' 0 0 150 28 $clrPanelAlt $null
        $flowUpdateTools.Controls.Add($btnDownloaderVersion)
        $btnCheckDownloaderUpdate = _Button 'Check Downloader Update' 0 0 190 28 $clrPanelAlt $null
        $flowUpdateTools.Controls.Add($btnCheckDownloaderUpdate)
        $btnCheckFiles = _Button 'Verify Required Files' 0 0 170 28 $clrPanelAlt $null
        $flowUpdateTools.Controls.Add($btnCheckFiles)
        $btnUpdateDownloader = _Button 'Update Downloader' 0 0 150 28 $clrAccent $null
        $flowUpdateTools.Controls.Add($btnUpdateDownloader)
        & $setHytaleTip 'Ask the Hytale services whether a newer server build is available.' @($btnCheckServerUpdate)
        & $setHytaleTip 'Show the version of the downloader file ECC is using for Hytale updates.' @($btnDownloaderVersion)
        & $setHytaleTip 'Check whether the downloader itself has an update available.' @($btnCheckDownloaderUpdate)
        & $setHytaleTip 'Check whether the required Hytale update files exist in this profile folder.' @($btnCheckFiles)
        & $setHytaleTip 'Download and replace the Hytale downloader file used by ECC.' @($btnUpdateDownloader)

        $groupStatus = New-Object System.Windows.Forms.GroupBox
        $groupStatus.Text = 'Required Files Status'
        $groupStatus.Location = [System.Drawing.Point]::new(16, 300)
        $groupStatus.Size = [System.Drawing.Size]::new($tabMain.ClientSize.Width - 48, 112)
        $groupStatus.Anchor = 'Top,Left,Right'
        $groupStatus.ForeColor = $clrText
        $tabUpdater.Controls.Add($groupStatus)

        $statusRows = @()
        foreach ($spec in @(
            @{ Key = 'HytaleServer.jar'; Label = 'HytaleServer.jar' },
            @{ Key = 'Downloader'; Label = 'hytale-downloader-windows-amd64.exe' },
            @{ Key = 'Assets.zip'; Label = 'Assets.zip' }
        )) {
            $rowY = 26 + ($statusRows.Count * 18)
            $lblName = _Label "$($spec.Label) - Unknown" 12 $rowY 260 18 $fontLabel
            $lblName.ForeColor = $clrTextSoft
            $groupStatus.Controls.Add($lblName)
            $lblState = _Label '[..]' 280 $rowY 90 18 $fontBold
            $lblState.ForeColor = $clrYellow
            $groupStatus.Controls.Add($lblState)
            $statusRows += [pscustomobject]@{ Key = $spec.Key; Name = $lblName; Status = $lblState }
            & $setHytaleTip ("Shows whether {0} is present in the Hytale server folder." -f $spec.Label) @($lblName, $lblState)
        }

        $lblOverallStatus = _Label 'Checking file state...' 12 82 ($groupStatus.ClientSize.Width - 24) 18 $fontBold
        $lblOverallStatus.ForeColor = $clrTextSoft
        $lblOverallStatus.Anchor = 'Top,Left,Right'
        $groupStatus.Controls.Add($lblOverallStatus)
        & $setHytaleTip 'Quick summary of whether all required Hytale update files were found.' @($lblOverallStatus)

        $logGroup = New-Object System.Windows.Forms.GroupBox
        $logGroup.Text = 'Activity Log'
        $logGroup.Location = [System.Drawing.Point]::new(16, 422)
        $logGroup.Size = [System.Drawing.Size]::new($tabMain.ClientSize.Width - 48, [Math]::Max(220, $tabMain.ClientSize.Height - 472))
        $logGroup.Anchor = 'Top,Left,Right,Bottom'
        $logGroup.ForeColor = $clrText
        $tabUpdater.Controls.Add($logGroup)

        $txtLog = New-Object System.Windows.Forms.TextBox
        $txtLog.Location = [System.Drawing.Point]::new(12, 24)
        $txtLog.Size = [System.Drawing.Size]::new($logGroup.ClientSize.Width - 24, $logGroup.ClientSize.Height - 36)
        $txtLog.Anchor = 'Top,Left,Right,Bottom'
        $txtLog.Multiline = $true
        $txtLog.ScrollBars = 'Vertical'
        $txtLog.WordWrap = $false
        $txtLog.ReadOnly = $true
        $txtLog.BackColor = [System.Drawing.Color]::FromArgb(20, 20, 28)
        $txtLog.ForeColor = $clrText
        $txtLog.BorderStyle = 'FixedSingle'
        $txtLog.Font = $fontMono
        $logGroup.Controls.Add($txtLog)
        & $setHytaleTip 'Live results from Hytale updater and mod-manager actions in this window.' @($txtLog)

        $layoutHytaleManager = { }.GetNewClosure()

        $appendLog = {
            param([string]$Line)
            if ([string]::IsNullOrWhiteSpace($Line)) { return }
            try { $txtLog.AppendText($Line + [Environment]::NewLine) } catch { }
        }.GetNewClosure()

        $appendResultLogs = {
            param([object]$Result)
            if ($null -eq $Result) { return }
            try {
                foreach ($line in @($Result.Logs)) {
                    if (-not [string]::IsNullOrWhiteSpace([string]$line)) {
                        & $appendLog ([string]$line)
                    }
                }
            } catch { }
        }.GetNewClosure()

        $showResultMessage = {
            param([string]$Title, [string]$Message, [bool]$Success = $true)
            [System.Windows.Forms.MessageBox]::Show(
                $Message,
                $Title,
                [System.Windows.Forms.MessageBoxButtons]::OK,
                $(if ($Success) { [System.Windows.Forms.MessageBoxIcon]::Information } else { [System.Windows.Forms.MessageBoxIcon]::Warning })
            ) | Out-Null
        }.GetNewClosure()

        $importHytaleModules = {
            try {
                if (-not [string]::IsNullOrWhiteSpace($hytaleLoggingModulePath)) { Import-Module $hytaleLoggingModulePath -Force -Global | Out-Null }
                if (-not [string]::IsNullOrWhiteSpace($hytaleProfileModulePath)) { Import-Module $hytaleProfileModulePath -Force -Global | Out-Null }
                if (-not [string]::IsNullOrWhiteSpace($hytaleServerModulePath)) { Import-Module $hytaleServerModulePath -Force -Global | Out-Null }
                if ($hytaleSharedState -is [hashtable]) {
                    & (Get-Command -Name 'Initialize-ServerManager' -ErrorAction Stop) -SharedState $hytaleSharedState
                }
                return $true
            } catch {
                & $appendLog ("[ERROR] Failed to import Hytale modules: {0}" -f $_.Exception.Message)
                return $false
            }
        }.GetNewClosure()

        $applyFilesStatus = {
            param([object]$Result)

            $mapped = @{}
            try {
                foreach ($item in @($Result.Data.Items)) {
                    if ($item -and $item.Name) { $mapped["$($item.Name)"] = $item }
                }
            } catch { }

            foreach ($row in @($statusRows)) {
                $item = if ($mapped.ContainsKey($row.Key)) { $mapped[$row.Key] } else { $null }
                if ($null -eq $item) {
                    $row.Name.Text = "$($row.Key) - Unknown"
                    $row.Name.ForeColor = [System.Drawing.Color]::LightGray
                    $row.Status.Text = '[..]'
                    $row.Status.ForeColor = [System.Drawing.Color]::Khaki
                    continue
                }

                if ($item.Exists -eq $true) {
                    $row.Name.Text = "$($row.Key) - Found"
                    $row.Name.ForeColor = [System.Drawing.Color]::LightGreen
                    $row.Status.Text = '[OK]'
                    $row.Status.ForeColor = [System.Drawing.Color]::LightGreen
                } else {
                    $row.Name.Text = "$($row.Key) - Missing"
                    $row.Name.ForeColor = [System.Drawing.Color]::Tomato
                    $row.Status.Text = '[!!]'
                    $row.Status.ForeColor = [System.Drawing.Color]::Tomato
                }
            }

            $allPresent = $false
            try { $allPresent = ($Result.Data.AllPresent -eq $true) } catch { $allPresent = $false }
            if ($allPresent) {
                $lblOverallStatus.Text = '[OK] All required files present'
                $lblOverallStatus.ForeColor = [System.Drawing.Color]::LightGreen
            } else {
                $lblOverallStatus.Text = '[WARN] Missing required files'
                $lblOverallStatus.ForeColor = [System.Drawing.Color]::Khaki
            }
        }.GetNewClosure()

        $refreshDownloaderPath = {
            try {
                if (-not [string]::IsNullOrWhiteSpace($hytaleRoot)) {
                    $txtDownloaderPath.Text = (Join-Path $hytaleRoot 'hytale-downloader-windows-amd64.exe')
                } else {
                    $txtDownloaderPath.Text = ''
                }
            } catch {
                $txtDownloaderPath.Text = ''
            }
        }.GetNewClosure()

        $refreshUpdateWarning = {
            try {
                $runningNow = $false
                $statusNow = Get-ServerStatus -Prefix $Prefix
                $runningNow = ($statusNow -and $statusNow.Running)
                if ($runningNow) {
                    $lblWarn.Text = '[!] Server is running and will be stopped during update. ECC may look busy during extraction, so do not click around or close the app.'
                } else {
                    $lblWarn.Text = '[i] Server is currently offline. ECC will update files without stopping anything first. Extraction can make ECC look busy for a moment.'
                }
            } catch {
                $lblWarn.Text = '[i] ECC will check server state before starting the update.'
            }
        }.GetNewClosure()

        $refreshFilesSync = {
            if (-not (& $importHytaleModules)) { return }
            try {
                $result = & (Get-Command -Name 'Get-HytaleRequiredFilesStatus' -ErrorAction Stop) -Prefix $Prefix
                & $applyFilesStatus $result
                & $appendResultLogs $result
            } catch {
                & $appendLog ("[ERROR] Failed to refresh file status: {0}" -f $_.Exception.Message)
            }
        }.GetNewClosure()

        $runUpdaterOp = {
            param([string]$OperationName, [scriptblock]$Script)
            $form.UseWaitCursor = $true
            try {
                & $appendLog ("[INFO] Starting: {0}" -f $OperationName)
                if (-not (& $importHytaleModules)) {
                    throw 'Hytale server tools could not be loaded.'
                }
                $result = & $Script
                & $appendResultLogs $result
                return $result
            } catch {
                & $appendLog ("[ERROR] {0}" -f $_.Exception.Message)
                return @{
                    Success = $false
                    Message = $_.Exception.Message
                    Data    = @{ Error = $_.Exception.Message }
                    Logs    = @("[ERROR] $($_.Exception.Message)")
                }
            } finally {
                $form.UseWaitCursor = $false
            }
        }.GetNewClosure()

        $flowModButtons = New-Object System.Windows.Forms.FlowLayoutPanel
        $flowModButtons.Location = [System.Drawing.Point]::new(16, 16)
        $flowModButtons.Size = [System.Drawing.Size]::new($tabMain.ClientSize.Width - 48, 72)
        $flowModButtons.Anchor = 'Top,Left,Right'
        $flowModButtons.WrapContents = $true
        $flowModButtons.AutoScroll = $false
        $flowModButtons.BackColor = [System.Drawing.Color]::Transparent
        $tabMods.Controls.Add($flowModButtons)

        $btnRefreshMods = _Button 'Refresh' 0 0 90 28 $clrPanelAlt $null
        $btnAddMod = _Button 'Add Mod...' 0 0 96 28 $clrAccent $null
        $btnToggleMod = _Button 'Toggle Mod' 0 0 108 28 $clrPanelAlt $null
        $btnOpenModsFolder = _Button 'Mod Folder' 0 0 98 28 $clrPanelAlt $null
        $btnDeleteMod = _Button 'Delete Mod' 0 0 96 28 $modDeleteColor $null
        $btnCheckConflicts = _Button 'Check Conflicts' 0 0 120 28 $clrPanelAlt $null
        $btnOpenSelectedMod = _Button 'Open Selected' 0 0 114 28 $clrPanelAlt $null
        $btnOpenConfigFolder = _Button 'Browse Configs' 0 0 118 28 $clrPanelAlt $null
        $btnLinkCurseForge = _Button 'Link Mod CF' 0 0 102 28 $modLinkColor $null
        $btnCheckModUpdates = _Button 'Check Updates' 0 0 110 28 $clrPanelAlt $null
        $btnUpdateSelectedMod = _Button 'Update Mod' 0 0 98 28 $clrAccent $null
        $btnOpenModPage = _Button 'Open Mod Page' 0 0 116 28 $clrPanelAlt $null
        $btnGetMoreMods = _Button 'Get More Mods' 0 0 118 28 $modBrowseColor $null
        foreach ($ctrl in @($btnRefreshMods, $btnAddMod, $btnToggleMod, $btnOpenModsFolder, $btnDeleteMod, $btnCheckConflicts, $btnOpenSelectedMod, $btnOpenConfigFolder, $btnLinkCurseForge, $btnCheckModUpdates, $btnUpdateSelectedMod, $btnOpenModPage, $btnGetMoreMods)) {
            $ctrl.Margin = [System.Windows.Forms.Padding]::new(0, 0, 8, 8)
            $flowModButtons.Controls.Add($ctrl)
        }
        & $setHytaleTip 'Reload the installed mod list and current file details from disk.' @($btnRefreshMods)
        & $setHytaleTip 'Pick a .jar file and copy it into this Hytale profile as a mod.' @($btnAddMod)
        & $setHytaleTip 'Move the selected mod between enabled and disabled folders.' @($btnToggleMod)
        & $setHytaleTip 'Open the folder where this profile keeps active Hytale mods.' @($btnOpenModsFolder)
        & $setHytaleTip 'Delete the selected mod file from disk after you confirm.' @($btnDeleteMod)
        & $setHytaleTip 'Look for obvious duplicate mod files that may conflict with each other.' @($btnCheckConflicts)
        & $setHytaleTip 'Open the exact mod file or folder for the selected entry.' @($btnOpenSelectedMod)
        & $setHytaleTip 'Open the Hytale config folder tied to this profile.' @($btnOpenConfigFolder)
        & $setHytaleTip 'Connect the selected mod to its CurseForge project so ECC can track updates.' @($btnLinkCurseForge)
        & $setHytaleTip 'Check CurseForge for updates for linked Hytale mods.' @($btnCheckModUpdates)
        & $setHytaleTip 'Download and replace the selected mod if ECC finds a newer linked CurseForge file.' @($btnUpdateSelectedMod)
        & $setHytaleTip 'Open the CurseForge page for the selected linked mod.' @($btnOpenModPage)
        & $setHytaleTip 'Open the Hytale mods page on CurseForge to find more mods.' @($btnGetMoreMods)

        $groupModList = New-Object System.Windows.Forms.GroupBox
        $groupModList.Text = 'Installed Mods'
        $groupModList.Location = [System.Drawing.Point]::new(16, 98)
        $groupModList.Size = [System.Drawing.Size]::new($tabMain.ClientSize.Width - 48, 360)
        $groupModList.Anchor = 'Top,Left,Right'
        $groupModList.ForeColor = $clrText
        $tabMods.Controls.Add($groupModList)

        $lblDropHint = _Label 'Drag .jar files into the list to install them into this Hytale profile.' 12 22 620 16
        $lblDropHint.ForeColor = $clrTextSoft
        $groupModList.Controls.Add($lblDropHint)
        & $setHytaleTip 'You can drag .jar files here to add them as mods for this profile.' @($lblDropHint)

        $lvMods = New-Object System.Windows.Forms.ListView
        $lvMods.Location = [System.Drawing.Point]::new(12, 44)
        $lvMods.Size = [System.Drawing.Size]::new($groupModList.ClientSize.Width - 24, $groupModList.ClientSize.Height - 56)
        $lvMods.Anchor = 'Top,Left,Right,Bottom'
        $lvMods.View = 'Details'
        $lvMods.FullRowSelect = $true
        $lvMods.GridLines = $true
        $lvMods.MultiSelect = $false
        $lvMods.HideSelection = $false
        $lvMods.AllowDrop = $true
        $lvMods.BackColor = [System.Drawing.Color]::FromArgb(20, 20, 28)
        $lvMods.ForeColor = $clrText
        [void]$lvMods.Columns.Add('Mod Name', 210)
        [void]$lvMods.Columns.Add('Status', 85)
        [void]$lvMods.Columns.Add('Version', 110)
        [void]$lvMods.Columns.Add('CF Status', 120)
        [void]$lvMods.Columns.Add('File Size', 90)
        [void]$lvMods.Columns.Add('File Path', 420)
        $groupModList.Controls.Add($lvMods)
        & $setHytaleTip 'Installed Hytale mods for this profile. Select one to view notes, status, and update actions.' @($lvMods)

        $groupModNotes = New-Object System.Windows.Forms.GroupBox
        $groupModNotes.Text = 'Selected Mod Notes'
        $groupModNotes.Location = [System.Drawing.Point]::new(16, 468)
        $groupModNotes.Size = [System.Drawing.Size]::new($tabMain.ClientSize.Width - 48, [Math]::Max(150, $tabMain.ClientSize.Height - 526))
        $groupModNotes.Anchor = 'Top,Left,Right,Bottom'
        $groupModNotes.ForeColor = $clrText
        $tabMods.Controls.Add($groupModNotes)

        $lblSelectedMod = _Label 'Select a mod to view notes and CurseForge status.' 12 22 ($groupModNotes.ClientSize.Width - 24) 34 $fontBold
        $lblSelectedMod.ForeColor = $clrTextSoft
        $lblSelectedMod.Anchor = 'Top,Left,Right'
        $lblSelectedMod.AutoEllipsis = $true
        $groupModNotes.Controls.Add($lblSelectedMod)
        & $setHytaleTip 'Shows which mod you are editing notes for and whether ECC knows its CurseForge status.' @($lblSelectedMod)

        $txtModNotes = New-Object System.Windows.Forms.TextBox
        $txtModNotes.Location = [System.Drawing.Point]::new(12, 62)
        $txtModNotes.Size = [System.Drawing.Size]::new($groupModNotes.ClientSize.Width - 24, $groupModNotes.ClientSize.Height - 108)
        $txtModNotes.Anchor = 'Top,Left,Right,Bottom'
        $txtModNotes.Multiline = $true
        $txtModNotes.ScrollBars = 'Vertical'
        $txtModNotes.BorderStyle = 'FixedSingle'
        $txtModNotes.BackColor = $clrPanelSoft
        $txtModNotes.ForeColor = $clrText
        $groupModNotes.Controls.Add($txtModNotes)
        & $setHytaleTip 'Private notes for the selected mod. Use this for reminders, install details, or admin notes.' @($txtModNotes)

        $modNotesFooter = New-Object System.Windows.Forms.Panel
        $modNotesFooter.Location = [System.Drawing.Point]::new(12, [Math]::Max(76, $groupModNotes.ClientSize.Height - 42))
        $modNotesFooter.Size = [System.Drawing.Size]::new($groupModNotes.ClientSize.Width - 24, 30)
        $modNotesFooter.Anchor = 'Left,Right,Bottom'
        $modNotesFooter.BackColor = [System.Drawing.Color]::Transparent
        $groupModNotes.Controls.Add($modNotesFooter)

        $flowModNotesActions = New-Object System.Windows.Forms.FlowLayoutPanel
        $flowModNotesActions.Location = [System.Drawing.Point]::new(0, 0)
        $flowModNotesActions.Size = [System.Drawing.Size]::new($modNotesFooter.ClientSize.Width, 30)
        $flowModNotesActions.Anchor = 'Top,Right'
        $flowModNotesActions.FlowDirection = 'RightToLeft'
        $flowModNotesActions.WrapContents = $false
        $flowModNotesActions.AutoSize = $true
        $flowModNotesActions.AutoSizeMode = 'GrowAndShrink'
        $flowModNotesActions.BackColor = [System.Drawing.Color]::Transparent
        $modNotesFooter.Controls.Add($flowModNotesActions)

        $btnSaveModNotes = _Button 'Save Notes' 0 0 92 28 $clrAccent $null
        $btnSaveModNotes.Margin = [System.Windows.Forms.Padding]::new(0)
        $flowModNotesActions.Controls.Add($btnSaveModNotes)
        & $setHytaleTip 'Save the notes shown for the selected mod.' @($btnSaveModNotes)

        $layoutHytaleManager = {
            $margin = 16
            $gap = 10

            if ($groupUpdate -is [System.Windows.Forms.Control]) {
                $updaterWidth = [Math]::Max(320, $tabUpdater.ClientSize.Width - ($margin * 2))
                $groupUpdate.SetBounds($margin, $margin, $updaterWidth, $groupUpdate.Height)
                $txtDownloaderPath.Size = [System.Drawing.Size]::new([Math]::Max(180, $groupUpdate.ClientSize.Width - 128), 22)
                $btnOpenFolder.Location = [System.Drawing.Point]::new([Math]::Max(12, $groupUpdate.ClientSize.Width - 108), 44)
                $flowUpdatePrimaryActions.SetBounds(12, 80, [Math]::Max(220, $groupUpdate.ClientSize.Width - 24), 32)
                $flowUpdatePrimaryActions.PerformLayout()
                $updateActionsBottom = 0
                foreach ($ctrl in @($flowUpdatePrimaryActions.Controls)) {
                    if ($ctrl -is [System.Windows.Forms.Control] -and $ctrl.Visible) {
                        $updateActionsBottom = [Math]::Max($updateActionsBottom, ($ctrl.Bottom + $ctrl.Margin.Bottom))
                    }
                }
                $flowUpdatePrimaryActions.Height = [Math]::Max(28, $updateActionsBottom + 2)

                $warnWidth = [Math]::Max(220, $groupUpdate.ClientSize.Width - 24)
                $warnFlags = [System.Windows.Forms.TextFormatFlags]::WordBreak
                $warnSize = [System.Windows.Forms.TextRenderer]::MeasureText([string]$lblWarn.Text, $lblWarn.Font, [System.Drawing.Size]::new($warnWidth, 0), $warnFlags)
                $warnHeight = [Math]::Max(18, $warnSize.Height)
                $lblWarn.SetBounds(12, $flowUpdatePrimaryActions.Bottom + 8, $warnWidth, $warnHeight)
                $groupUpdate.Height = $lblWarn.Bottom + 12

                $groupUtils.SetBounds($margin, $groupUpdate.Bottom + $gap, $updaterWidth, 128)
                $flowUpdateTools.Location = [System.Drawing.Point]::new(12, 28)
                $flowUpdateTools.Width = [Math]::Max(220, $groupUtils.ClientSize.Width - 24)
                $flowUpdateTools.Height = 80
                $flowUpdateTools.PerformLayout()
                $updateToolsBottom = 0
                foreach ($ctrl in @($flowUpdateTools.Controls)) {
                    if ($ctrl -is [System.Windows.Forms.Control] -and $ctrl.Visible) {
                        $updateToolsBottom = [Math]::Max($updateToolsBottom, ($ctrl.Bottom + $ctrl.Margin.Bottom))
                    }
                }
                $flowUpdateTools.Height = [Math]::Max(36, $updateToolsBottom + 4)
                $groupUtils.Height = $flowUpdateTools.Bottom + 12

                $groupStatus.SetBounds($margin, $groupUtils.Bottom + $gap, $updaterWidth, 112)
                $lblOverallStatus.Size = [System.Drawing.Size]::new([Math]::Max(220, $groupStatus.ClientSize.Width - 24), 18)

                $logTop = $groupStatus.Bottom + $gap
                $logHeight = [Math]::Max(180, $tabUpdater.ClientSize.Height - $logTop - $margin)
                $logGroup.SetBounds($margin, $logTop, $updaterWidth, $logHeight)
                $txtLog.Size = [System.Drawing.Size]::new([Math]::Max(220, $logGroup.ClientSize.Width - 24), [Math]::Max(120, $logGroup.ClientSize.Height - 36))
            }

            if ($flowModButtons -is [System.Windows.Forms.Control] -and
                $groupModList -is [System.Windows.Forms.Control] -and
                $groupModNotes -is [System.Windows.Forms.Control]) {
                $modsWidth = [Math]::Max(320, $tabMods.ClientSize.Width - ($margin * 2))

                $flowModButtons.SetBounds($margin, $margin, $modsWidth, 72)
                $flowModButtons.PerformLayout()
                $modButtonsBottom = 0
                foreach ($ctrl in @($flowModButtons.Controls)) {
                    if ($ctrl -is [System.Windows.Forms.Control] -and $ctrl.Visible) {
                        $modButtonsBottom = [Math]::Max($modButtonsBottom, ($ctrl.Bottom + $ctrl.Margin.Bottom))
                    }
                }
                $flowModButtons.Height = [Math]::Max(42, $modButtonsBottom + 6)

                $modListTop = $flowModButtons.Bottom + $gap + 4
                $notesHeightTarget = [Math]::Max(150, [int]([Math]::Floor($tabMods.ClientSize.Height * 0.28)))
                $notesTop = $tabMods.ClientSize.Height - $margin - $notesHeightTarget
                $notesTop = [Math]::Max($modListTop + 180, $notesTop)
                $notesHeight = [Math]::Max(150, $tabMods.ClientSize.Height - $notesTop - $margin)
                $listHeight = [Math]::Max(220, $notesTop - $modListTop - $gap)

                $groupModList.SetBounds($margin, $modListTop, $modsWidth, $listHeight)
                $groupModNotes.SetBounds($margin, $notesTop, $modsWidth, $notesHeight)

                $lvMods.Size = [System.Drawing.Size]::new([Math]::Max(220, $groupModList.ClientSize.Width - 24), [Math]::Max(140, $groupModList.ClientSize.Height - 56))
                $notesInnerWidth = [Math]::Max(220, $groupModNotes.ClientSize.Width - 24)
                $footerHeight = 30
                $footerBottomMargin = 12
                $notesTextTop = 62
                $notesGap = 8

                $lblSelectedMod.Size = [System.Drawing.Size]::new($notesInnerWidth, 34)
                $modNotesFooter.SetBounds(12, [Math]::Max(76, $groupModNotes.ClientSize.Height - $footerHeight - $footerBottomMargin), $notesInnerWidth, $footerHeight)
                $flowModNotesActions.PerformLayout()
                $flowModNotesActions.Location = [System.Drawing.Point]::new([Math]::Max(0, $modNotesFooter.ClientSize.Width - $flowModNotesActions.PreferredSize.Width), 0)
                $txtModNotes.Size = [System.Drawing.Size]::new($notesInnerWidth, [Math]::Max(70, $modNotesFooter.Top - $notesTextTop - $notesGap))
            }
        }.GetNewClosure()

        $cfMetadata = @{}
        $modNotes = @{}

        $loadJsonMap = {
            param([string]$Path)
            if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) { return @{} }
            try {
                $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
                if (Get-Command -Name 'ConvertTo-Hashtable' -ErrorAction SilentlyContinue) {
                    return (ConvertTo-Hashtable -Object $raw)
                }
                return @{}
            } catch {
                return @{}
            }
        }.GetNewClosure()

        $saveJsonMap = {
            param([string]$Path, [hashtable]$Map)
            if ([string]::IsNullOrWhiteSpace($Path)) { return }
            ($Map | ConvertTo-Json -Depth 8) | Set-Content -LiteralPath $Path -Encoding UTF8
        }.GetNewClosure()

        $ensureModFolders = {
            if (-not $modsFeatureReady) { return $false }
            foreach ($path in @($modsPath, $modsDisabledPath)) {
                if (-not (Test-Path -LiteralPath $path)) { New-Item -ItemType Directory -Path $path -Force | Out-Null }
            }
            return $true
        }.GetNewClosure()

        $normalizeLooseName = { param([string]$Text) if ([string]::IsNullOrWhiteSpace($Text)) { '' } else { (($Text -replace '[^a-zA-Z0-9]+', '').ToLowerInvariant()) } }.GetNewClosure()
        $getModVersionFromFilename = {
            param([string]$FileName)
            $base = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
            $match = [regex]::Match($base, '(?i)(\d+(?:\.\d+)+(?:[-+._a-z0-9]+)?)')
            if ($match.Success) { $match.Groups[1].Value } else { 'Unknown' }
        }.GetNewClosure()
        $getModNameFromFilename = {
            param([string]$FileName)
            $name = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
            $name = [regex]::Replace($name, '(?i)[-_ ]?\d+(?:\.\d+)+(?:[-+._a-z0-9]+)?$', '')
            $name = ($name -replace '[_\.]+', ' ').Trim()
            if ([string]::IsNullOrWhiteSpace($name)) { $name = [System.IO.Path]::GetFileNameWithoutExtension($FileName) }
            $name
        }.GetNewClosure()

        $getCurseForgeProject = {
            param([int]$ProjectId)
            try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }
            $wc = New-Object System.Net.WebClient
            try {
                $wc.Headers['x-api-key'] = $cfApiKey
                $wc.Headers['Accept'] = 'application/json'
                $projectUrl = ('{0}/mods/{1}' -f $cfApiBase, $ProjectId)
                $filesUrl = ('{0}/mods/{1}/files?pageSize=20&sortDescending=true' -f $cfApiBase, $ProjectId)
                $projectJson = $wc.DownloadString($projectUrl)
                $filesJson = $wc.DownloadString($filesUrl)
            } finally {
                $wc.Dispose()
            }
            $project = if (-not [string]::IsNullOrWhiteSpace([string]$projectJson)) { $projectJson | ConvertFrom-Json -ErrorAction Stop } else { $null }
            $files = if (-not [string]::IsNullOrWhiteSpace([string]$filesJson)) { $filesJson | ConvertFrom-Json -ErrorAction Stop } else { $null }
            $latest = @($files.data) | Where-Object { $_ -and ([string]$_.fileName).ToLowerInvariant().EndsWith('.jar') } | Select-Object -First 1
            [pscustomobject]@{
                ProjectId      = $ProjectId
                ProjectName    = [string]$project.data.name
                WebsiteUrl     = [string]$project.data.links.websiteUrl
                LatestFileId   = if ($latest) { [int]$latest.id } else { 0 }
                LatestFileName = if ($latest) { [string]$latest.fileName } else { '' }
                DownloadUrl    = if ($latest) { [string]$latest.downloadUrl } else { '' }
            }
        }.GetNewClosure()

        $searchCurseForgeMods = {
            param(
                [string]$SearchQuery,
                [int]$PageSize = 12
            )

            if ([string]::IsNullOrWhiteSpace($SearchQuery)) { return @() }
            $encodedQuery = [System.Uri]::EscapeDataString($SearchQuery.Trim())
            try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }
            $rawResponse = Invoke-WebRequest -Uri ('{0}/mods/search?gameId={1}&classId={2}&searchFilter={3}&pageSize={4}' -f $cfApiBase, $cfGameId, $cfClassId, $encodedQuery, $PageSize) -Headers @{ 'x-api-key' = $cfApiKey; Accept = 'application/json' } -Method Get -UseBasicParsing -ErrorAction Stop
            $response = if ($rawResponse -and -not [string]::IsNullOrWhiteSpace([string]$rawResponse.Content)) { $rawResponse.Content | ConvertFrom-Json -ErrorAction Stop } else { $null }
            @($response.data)
        }.GetNewClosure()

        $getSelectedMod = { if ($lvMods.SelectedItems.Count -gt 0) { $lvMods.SelectedItems[0].Tag } else { $null } }.GetNewClosure()

        $setModActionState = {
            $selected = & $getSelectedMod
            $hasSelection = ($null -ne $selected)

            foreach ($ctrl in @($btnToggleMod, $btnDeleteMod, $btnOpenSelectedMod, $btnOpenConfigFolder, $btnLinkCurseForge, $btnUpdateSelectedMod, $btnOpenModPage, $btnSaveModNotes)) {
                try {
                    if ($ctrl -is [System.Windows.Forms.Control]) {
                        $ctrl.Enabled = ($modsFeatureReady -and $hasSelection)
                    }
                } catch { }
            }
        }.GetNewClosure()

        $refreshModList = {
            if (-not $modsFeatureReady) { $lvMods.Items.Clear(); return }
            if (-not (& $ensureModFolders)) { return }
            $previousSelection = $null
            try {
                $selected = & $getSelectedMod
                if ($selected) { $previousSelection = [string]$selected.FileName }
            } catch { $previousSelection = $null }

            $lvMods.BeginUpdate()
            try {
                $lvMods.Items.Clear()
                foreach ($entry in @(@{ Path = $modsPath; Status = 'Enabled' }, @{ Path = $modsDisabledPath; Status = 'Disabled' })) {
                    if (-not (Test-Path -LiteralPath $entry.Path)) { continue }
                    foreach ($file in Get-ChildItem -LiteralPath $entry.Path -Filter '*.jar' -File -ErrorAction SilentlyContinue | Sort-Object Name) {
                        $meta = if ($cfMetadata.ContainsKey($file.Name)) { $cfMetadata[$file.Name] } else { $null }
                        $cfState = 'Unlinked'
                        if ($meta) {
                            $cfState = 'Linked'
                            if ($meta.ContainsKey('LatestKnownFileId') -and $meta.ContainsKey('FileId') -and ([string]$meta.LatestKnownFileId -ne [string]$meta.FileId)) { $cfState = 'Update avail' }
                        }
                        $mod = [pscustomobject]@{
                            DisplayName = & $getModNameFromFilename $file.Name
                            FileName    = $file.Name
                            Status      = $entry.Status
                            Version     = & $getModVersionFromFilename $file.Name
                            CFStatus    = $cfState
                            FileSize    = ('{0:N1} MB' -f ($file.Length / 1MB))
                            FilePath    = $file.FullName
                        }
                        $item = New-Object System.Windows.Forms.ListViewItem($mod.DisplayName)
                        [void]$item.SubItems.Add($mod.Status)
                        [void]$item.SubItems.Add($mod.Version)
                        [void]$item.SubItems.Add($mod.CFStatus)
                        [void]$item.SubItems.Add($mod.FileSize)
                        [void]$item.SubItems.Add($mod.FilePath)
                        if ($mod.Status -eq 'Disabled') { $item.ForeColor = [System.Drawing.Color]::Silver }
                        $item.Tag = $mod
                        [void]$lvMods.Items.Add($item)
                    }
                }

                if (-not [string]::IsNullOrWhiteSpace($previousSelection)) {
                    foreach ($item in @($lvMods.Items)) {
                        try {
                            if ($item.Tag -and [string]$item.Tag.FileName -eq $previousSelection) {
                                $item.Selected = $true
                                $item.Focused = $true
                                $item.EnsureVisible()
                                break
                            }
                        } catch { }
                    }
                }
            } finally {
                $lvMods.EndUpdate()
            }

            & $setModActionState
        }.GetNewClosure()

        $updateModSelectionUi = {
            $selected = & $getSelectedMod
            if ($selected) {
                $lblSelectedMod.Text = '{0}  [{1}]' -f $selected.DisplayName, $selected.Status
                $txtModNotes.Text = if ($modNotes.ContainsKey($selected.FileName)) { [string]$modNotes[$selected.FileName] } else { '' }
            } else {
                $lblSelectedMod.Text = 'Select a mod to view notes and CurseForge status.'
                $txtModNotes.Text = ''
            }
            & $setModActionState
        }.GetNewClosure()

        $saveCurseForgeLink = {
            param(
                [psobject]$SelectedMod,
                [psobject]$ProjectInfo
            )

            if ($null -eq $SelectedMod -or $null -eq $ProjectInfo) { return }
            $cfMetadata[$SelectedMod.FileName] = @{
                ProjectId         = $ProjectInfo.ProjectId
                ProjectName       = $ProjectInfo.ProjectName
                WebsiteUrl        = $ProjectInfo.WebsiteUrl
                FileId            = $ProjectInfo.LatestFileId
                LatestKnownFileId = $ProjectInfo.LatestFileId
                LatestKnownName   = $ProjectInfo.LatestFileName
                LatestDownloadUrl = $ProjectInfo.DownloadUrl
            }
            & $saveJsonMap $cfMetadataPath $cfMetadata
            & $appendLog ("[INFO] Linked {0} to CurseForge project {1} ({2})" -f $SelectedMod.DisplayName, $ProjectInfo.ProjectName, $ProjectInfo.ProjectId)
            & $refreshModList
            & $updateModSelectionUi
        }.GetNewClosure()

        $showCurseForgeLinkDialog = {
            $selected = & $getSelectedMod
            if ($null -eq $selected) {
                [System.Windows.Forms.MessageBox]::Show('Select a mod first.', 'Link Mod CF', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
                return
            }

            $dialogCfApiKey = [string]$cfApiKey
            $dialogCfApiBase = [string]$cfApiBase
            $dialogCfGameId = [int]$cfGameId
            $dialogCfClassId = [int]$cfClassId
            $dialogSaveCurseForgeLink = $saveCurseForgeLink
            $dialogAppendLog = $appendLog
            $dialogRefreshModList = $refreshModList
            $dialogUpdateModSelectionUi = $updateModSelectionUi

            $dialogGetCurseForgeProject = {
                param([int]$ProjectId)
                try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }
                $wc = New-Object System.Net.WebClient
                try {
                    $wc.Headers['x-api-key'] = $dialogCfApiKey
                    $wc.Headers['Accept'] = 'application/json'
                    $projectUrl = ('{0}/mods/{1}' -f $dialogCfApiBase, $ProjectId)
                    $filesUrl = ('{0}/mods/{1}/files?pageSize=20&sortDescending=true' -f $dialogCfApiBase, $ProjectId)
                    $projectJson = $wc.DownloadString($projectUrl)
                    $filesJson = $wc.DownloadString($filesUrl)
                } finally {
                    $wc.Dispose()
                }

                $project = if (-not [string]::IsNullOrWhiteSpace([string]$projectJson)) { $projectJson | ConvertFrom-Json -ErrorAction Stop } else { $null }
                $files = if (-not [string]::IsNullOrWhiteSpace([string]$filesJson)) { $filesJson | ConvertFrom-Json -ErrorAction Stop } else { $null }
                $latest = @($files.data) | Where-Object { $_ -and ([string]$_.fileName).ToLowerInvariant().EndsWith('.jar') } | Select-Object -First 1
                [pscustomobject]@{
                    ProjectId      = $ProjectId
                    ProjectName    = [string]$project.data.name
                    WebsiteUrl     = [string]$project.data.links.websiteUrl
                    LatestFileId   = if ($latest) { [int]$latest.id } else { 0 }
                    LatestFileName = if ($latest) { [string]$latest.fileName } else { '' }
                    DownloadUrl    = if ($latest) { [string]$latest.downloadUrl } else { '' }
                }
            }.GetNewClosure()

            $dialogBg = if ($clrBg -is [System.Drawing.Color]) { $clrBg } else { [System.Drawing.Color]::FromArgb(15, 18, 30) }
            $dialogText = if ($clrText -is [System.Drawing.Color]) { $clrText } else { [System.Drawing.Color]::WhiteSmoke }
            $dialogSoftText = if ($clrTextSoft -is [System.Drawing.Color]) { $clrTextSoft } else { [System.Drawing.Color]::Silver }
            $dialogPanel = if ($clrPanelSoft -is [System.Drawing.Color]) { $clrPanelSoft } else { [System.Drawing.Color]::FromArgb(28, 32, 48) }
            $dialogButton = if ($clrPanelAlt -is [System.Drawing.Color]) { $clrPanelAlt } else { [System.Drawing.Color]::FromArgb(45, 52, 78) }
            $dialogAccent = if ($clrAccent -is [System.Drawing.Color]) { $clrAccent } else { [System.Drawing.Color]::FromArgb(88, 137, 255) }
            $dialogListBg = if ($lvMods -and $lvMods.BackColor -is [System.Drawing.Color]) { $lvMods.BackColor } else { [System.Drawing.Color]::FromArgb(20, 20, 28) }
            $dialogFontLabel = if ($fontLabel -is [System.Drawing.Font]) { $fontLabel } else { New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Regular) }
            $dialogFontBold = if ($fontBold -is [System.Drawing.Font]) { $fontBold } else { New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold) }

            $newDialogLabel = {
                param(
                    [string]$Text,
                    [int]$X,
                    [int]$Y,
                    [int]$Width,
                    [int]$Height,
                    [System.Drawing.Font]$Font = $dialogFontLabel,
                    [System.Drawing.Color]$ForeColor = $dialogText
                )
                $label = New-Object System.Windows.Forms.Label
                $label.Text = $Text
                $label.Location = [System.Drawing.Point]::new($X, $Y)
                $label.Size = [System.Drawing.Size]::new($Width, $Height)
                $label.ForeColor = $ForeColor
                $label.BackColor = [System.Drawing.Color]::Transparent
                $label.Font = $Font
                return $label
            }.GetNewClosure()

            $newDialogButton = {
                param(
                    [string]$Text,
                    [int]$X,
                    [int]$Y,
                    [int]$Width,
                    [int]$Height,
                    [System.Drawing.Color]$BackColor
                )
                $button = New-Object System.Windows.Forms.Button
                $button.Text = $Text
                $button.Location = [System.Drawing.Point]::new($X, $Y)
                $button.Size = [System.Drawing.Size]::new($Width, $Height)
                $button.BackColor = $BackColor
                $button.ForeColor = [System.Drawing.Color]::White
                $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                $button.FlatAppearance.BorderSize = 1
                $button.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(82, 92, 128)
                $button.Font = $dialogFontBold
                $button.UseVisualStyleBackColor = $false
                return $button
            }.GetNewClosure()

            $searchForm = New-Object System.Windows.Forms.Form
            $searchForm.Text = "Link $($selected.DisplayName) to CurseForge"
            $searchForm.Size = [System.Drawing.Size]::new(760, 620)
            $searchForm.MinimumSize = [System.Drawing.Size]::new(760, 620)
            $searchForm.StartPosition = 'CenterParent'
            $searchForm.BackColor = $dialogBg
            $searchForm.ForeColor = $dialogText
            $searchForm.FormBorderStyle = 'Sizable'
            $searchForm.MaximizeBox = $false
            $searchForm.MinimizeBox = $false

            $lblModFile = & $newDialogLabel ("Mod File: {0}" -f $selected.FileName) 16 14 712 20 $dialogFontBold $dialogText
            $lblModFile.Anchor = 'Top,Left,Right'
            $searchForm.Controls.Add($lblModFile)

            $lblSearch = & $newDialogLabel 'Search CurseForge:' 16 44 180 18 $dialogFontLabel $dialogText
            $searchForm.Controls.Add($lblSearch)

            $txtSearch = New-Object System.Windows.Forms.TextBox
            $txtSearch.Location = [System.Drawing.Point]::new(16, 66)
            $txtSearch.Size = [System.Drawing.Size]::new(566, 24)
            $txtSearch.Anchor = 'Top,Left,Right'
            $txtSearch.BorderStyle = 'FixedSingle'
            $txtSearch.BackColor = $dialogPanel
            $txtSearch.ForeColor = $dialogText
            $txtSearch.Text = [string]$selected.DisplayName
            $searchForm.Controls.Add($txtSearch)

            $btnSearchCurseForge = & $newDialogButton 'Search' 594 64 136 28 $modLinkColor
            $btnSearchCurseForge.Anchor = 'Top,Right'
            $searchForm.Controls.Add($btnSearchCurseForge)

            $lvResults = New-Object System.Windows.Forms.ListView
            $lvResults.Location = [System.Drawing.Point]::new(16, 106)
            $lvResults.Size = [System.Drawing.Size]::new(714, 388)
            $lvResults.Anchor = 'Top,Left,Right,Bottom'
            $lvResults.View = 'Details'
            $lvResults.FullRowSelect = $true
            $lvResults.GridLines = $true
            $lvResults.MultiSelect = $false
            $lvResults.HideSelection = $false
            $lvResults.BackColor = $dialogListBg
            $lvResults.ForeColor = $dialogText
            [void]$lvResults.Columns.Add('Mod Name', 280)
            [void]$lvResults.Columns.Add('Author', 150)
            [void]$lvResults.Columns.Add('Downloads', 110)
            [void]$lvResults.Columns.Add('Project ID', 110)
            $searchForm.Controls.Add($lvResults)

            $manualFooter = New-Object System.Windows.Forms.Panel
            $manualFooter.Location = [System.Drawing.Point]::new(16, 500)
            $manualFooter.Size = [System.Drawing.Size]::new(714, 30)
            $manualFooter.Anchor = 'Left,Right,Bottom'
            $manualFooter.BackColor = [System.Drawing.Color]::Transparent
            $searchForm.Controls.Add($manualFooter)

            $lblManual = & $newDialogLabel 'Or enter a CurseForge project ID manually:' 0 6 240 18 $dialogFontLabel $dialogText
            $lblManual.Anchor = 'Left,Top'
            $manualFooter.Controls.Add($lblManual)

            $txtProjectId = New-Object System.Windows.Forms.TextBox
            $txtProjectId.Location = [System.Drawing.Point]::new(242, 3)
            $txtProjectId.Size = [System.Drawing.Size]::new(160, 24)
            $txtProjectId.Anchor = 'Left,Right,Top'
            $txtProjectId.BorderStyle = 'FixedSingle'
            $txtProjectId.BackColor = $dialogPanel
            $txtProjectId.ForeColor = $dialogText
            $manualFooter.Controls.Add($txtProjectId)

            $actionFooter = New-Object System.Windows.Forms.FlowLayoutPanel
            $actionFooter.Location = [System.Drawing.Point]::new(440, 536)
            $actionFooter.Size = [System.Drawing.Size]::new(290, 32)
            $actionFooter.Anchor = 'Right,Bottom'
            $actionFooter.WrapContents = $false
            $actionFooter.AutoScroll = $false
            $actionFooter.AutoSize = $true
            $actionFooter.AutoSizeMode = 'GrowAndShrink'
            $actionFooter.FlowDirection = [System.Windows.Forms.FlowDirection]::RightToLeft
            $actionFooter.BackColor = [System.Drawing.Color]::Transparent
            $searchForm.Controls.Add($actionFooter)

            $btnConfirmLink = & $newDialogButton 'Link Selected' 0 0 136 30 $dialogAccent
            $btnConfirmLink.Enabled = $false
            $btnConfirmLink.Margin = [System.Windows.Forms.Padding]::new(0)
            $actionFooter.Controls.Add($btnConfirmLink)

            $btnCancelLink = & $newDialogButton 'Cancel' 0 0 136 30 $dialogButton
            $btnCancelLink.Margin = [System.Windows.Forms.Padding]::new(10, 0, 0, 0)
            $actionFooter.Controls.Add($btnCancelLink)

            $layoutCurseForgeLinkDialog = {
                $margin = 16
                $searchFormWidth = [Math]::Max(420, $searchForm.ClientSize.Width - ($margin * 2))

                $lblModFile.Size = [System.Drawing.Size]::new($searchFormWidth, 20)
                $lblSearch.Location = [System.Drawing.Point]::new($margin, 44)
                $txtSearch.Location = [System.Drawing.Point]::new($margin, 66)
                $btnSearchCurseForge.Location = [System.Drawing.Point]::new([Math]::Max($margin + 220, $searchForm.ClientSize.Width - $margin - 136), 64)
                $txtSearch.Size = [System.Drawing.Size]::new([Math]::Max(220, $btnSearchCurseForge.Left - $margin - 12), 24)

                $manualFooter.SetBounds($margin, [Math]::Max(0, $searchForm.ClientSize.Height - 84), $searchFormWidth, 30)
                $lblManual.Location = [System.Drawing.Point]::new(0, 6)
                $manualInputLeft = $lblManual.Right + 10
                $txtProjectId.Location = [System.Drawing.Point]::new($manualInputLeft, 3)
                $txtProjectId.Size = [System.Drawing.Size]::new([Math]::Max(120, $manualFooter.ClientSize.Width - $manualInputLeft), 24)

                $actionFooter.PerformLayout()
                $actionFooter.Location = [System.Drawing.Point]::new(
                    [Math]::Max($margin, $searchForm.ClientSize.Width - $margin - $actionFooter.PreferredSize.Width),
                    [Math]::Max($manualFooter.Bottom + 6, $searchForm.ClientSize.Height - 48)
                )

                $resultsTop = 106
                $resultsBottom = [Math]::Max($resultsTop + 140, $manualFooter.Top - 6)
                $lvResults.SetBounds($margin, $resultsTop, $searchFormWidth, [Math]::Max(140, $resultsBottom - $resultsTop))
            }.GetNewClosure()

            $syncLinkButtonState = {
                $hasManualId = -not [string]::IsNullOrWhiteSpace($txtProjectId.Text)
                $hasSelection = ($lvResults.SelectedItems.Count -gt 0) -and ($null -ne $lvResults.SelectedItems[0].Tag)
                $btnConfirmLink.Enabled = ($hasManualId -or $hasSelection)
            }.GetNewClosure()

            $performSearch = {
                $query = [string]$txtSearch.Text
                if ([string]::IsNullOrWhiteSpace($query)) {
                    [System.Windows.Forms.MessageBox]::Show('Enter a mod name before searching CurseForge.', 'Link Mod CF', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
                    return
                }

                $searchForm.UseWaitCursor = $true
                $btnSearchCurseForge.Enabled = $false
                $lvResults.BeginUpdate()
                try {
                    $lvResults.Items.Clear()
                    $safeQuery = if ($null -ne $query) { [string]$query } else { '' }
                    $encodedQuery = [System.Uri]::EscapeDataString($safeQuery.Trim())
                    $searchUrl = ''
                    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }
                    $wc = New-Object System.Net.WebClient
                    try {
                        $wc.Headers['x-api-key'] = $dialogCfApiKey
                        $wc.Headers['Accept'] = 'application/json'
                        $searchUrl = ('{0}/mods/search?gameId={1}&classId={2}&searchFilter={3}&pageSize={4}' -f $dialogCfApiBase, $dialogCfGameId, $dialogCfClassId, $encodedQuery, 15)
                        $rawJson = $wc.DownloadString($searchUrl)
                    } finally {
                        $wc.Dispose()
                    }
                    $response = if (-not [string]::IsNullOrWhiteSpace([string]$rawJson)) { $rawJson | ConvertFrom-Json -ErrorAction Stop } else { $null }
                    $results = @()
                    try {
                        if ($null -ne $response -and $null -ne $response.data) {
                            $results = @($response.data)
                        }
                    } catch {
                        $results = @()
                    }
                    foreach ($result in $results) {
                        if ($null -eq $result) { continue }
                        try {
                            $author = 'Unknown'
                            try {
                                $authors = @($result.authors)
                                if ($authors.Count -gt 0 -and $null -ne $authors[0]) {
                                    $authorName = [string]$authors[0].name
                                    if (-not [string]::IsNullOrWhiteSpace($authorName)) { $author = $authorName }
                                }
                            } catch { $author = 'Unknown' }

                            $downloads = '0'
                            try {
                                if ($null -ne $result.downloadCount) { $downloads = ('{0:N0}' -f ([double]$result.downloadCount)) }
                            } catch { $downloads = '0' }

                            $itemText = ''
                            try { $itemText = [string]$result.name } catch { $itemText = '' }
                            if ([string]::IsNullOrWhiteSpace($itemText)) { $itemText = '(Unnamed CurseForge Mod)' }

                            $projectIdText = ''
                            try { $projectIdText = [string]$result.id } catch { $projectIdText = '' }

                            $item = New-Object System.Windows.Forms.ListViewItem($itemText)
                            [void]$item.SubItems.Add($author)
                            [void]$item.SubItems.Add($downloads)
                            [void]$item.SubItems.Add($projectIdText)
                            $item.Tag = $result
                            [void]$lvResults.Items.Add($item)
                        } catch {
                            try { & $appendLog ("[WARN] Skipping malformed CurseForge search result: {0}" -f $_.Exception.Message) } catch { }
                        }
                    }

                    if ($lvResults.Items.Count -le 0) {
                        $empty = New-Object System.Windows.Forms.ListViewItem('No CurseForge matches found for this search')
                        $empty.ForeColor = $dialogSoftText
                        [void]$lvResults.Items.Add($empty)
                    }
                } catch {
                    $cfMsg = ''
                    $cfLine = ''
                    $cfStack = ''
                    try { $cfMsg = [string]$_.Exception.ToString() } catch { $cfMsg = [string]$_.Exception.Message }
                    try { $cfLine = [string]$_.InvocationInfo.PositionMessage } catch { $cfLine = '' }
                    try { $cfStack = [string]$_.ScriptStackTrace } catch { $cfStack = '' }
                    $cfUrl = ''
                    try { $cfUrl = [string]$searchUrl } catch { $cfUrl = '' }
                    try { & $appendLog ("[ERROR] CurseForge search failed for '{0}': {1}" -f $safeQuery, $cfMsg) } catch { }
                    if (-not [string]::IsNullOrWhiteSpace($cfUrl)) {
                        try { & $appendLog ("[ERROR] CurseForge search URL: {0}" -f $cfUrl) } catch { }
                    }
                    if (-not [string]::IsNullOrWhiteSpace($cfLine)) {
                        try { & $appendLog ("[ERROR] CurseForge search location: {0}" -f $cfLine) } catch { }
                    }
                    if (-not [string]::IsNullOrWhiteSpace($cfStack)) {
                        try { & $appendLog ("[ERROR] CurseForge search stack: {0}" -f $cfStack) } catch { }
                    }
                    try {
                        if ($script:_ProgramLogBox) {
                            _WriteProgramLog ("[{0}][ERROR][GUI] CurseForge search failed for '{1}': {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $safeQuery, $cfMsg)
                            if (-not [string]::IsNullOrWhiteSpace($cfUrl)) {
                                _WriteProgramLog ("[{0}][ERROR][GUI] CurseForge search URL: {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $cfUrl)
                            }
                            if (-not [string]::IsNullOrWhiteSpace($cfLine)) {
                                _WriteProgramLog ("[{0}][ERROR][GUI] CurseForge search location: {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $cfLine)
                            }
                            if (-not [string]::IsNullOrWhiteSpace($cfStack)) {
                                _WriteProgramLog ("[{0}][ERROR][GUI] CurseForge search stack: {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $cfStack)
                            }
                        }
                    } catch { }
                    try { _GuiModuleLog -Level ERROR -Message ("CurseForge search failed for '{0}'. Error: {1}" -f $safeQuery, $cfMsg) } catch { }
                    if (-not [string]::IsNullOrWhiteSpace($cfUrl)) {
                        try { _GuiModuleLog -Level ERROR -Message ("CurseForge search URL: {0}" -f $cfUrl) } catch { }
                    }
                    if (-not [string]::IsNullOrWhiteSpace($cfLine)) {
                        try { _GuiModuleLog -Level ERROR -Message ("CurseForge search location: {0}" -f $cfLine) } catch { }
                    }
                    if (-not [string]::IsNullOrWhiteSpace($cfStack)) {
                        try { _GuiModuleLog -Level ERROR -Message ("CurseForge search stack: {0}" -f $cfStack) } catch { }
                    }
                    $popupDetail = [string]$_.Exception.Message
                    if (-not [string]::IsNullOrWhiteSpace($cfUrl)) {
                        $popupDetail = "{0}`r`n`r`nURL:`r`n{1}" -f $popupDetail, $cfUrl
                    }
                    if (-not [string]::IsNullOrWhiteSpace($cfLine)) {
                        $popupDetail = "{0}`r`n`r`nLocation:`r`n{1}" -f $popupDetail, $cfLine
                    }
                    [System.Windows.Forms.MessageBox]::Show(("CurseForge search failed.`r`n`r`n{0}" -f $popupDetail), 'Link Mod CF', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
                } finally {
                    $lvResults.EndUpdate()
                    $btnSearchCurseForge.Enabled = $true
                    $searchForm.UseWaitCursor = $false
                    & $syncLinkButtonState
                }
            }.GetNewClosure()

            $completeLink = {
                $projectId = 0

                if (-not [string]::IsNullOrWhiteSpace($txtProjectId.Text)) {
                    if (-not [int]::TryParse($txtProjectId.Text.Trim(), [ref]$projectId)) {
                        [System.Windows.Forms.MessageBox]::Show('The project ID must be a number.', 'Link Mod CF', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
                        return
                    }
                } elseif ($lvResults.SelectedItems.Count -gt 0 -and $null -ne $lvResults.SelectedItems[0].Tag) {
                    try { $projectId = [int]$lvResults.SelectedItems[0].Tag.id } catch { $projectId = 0 }
                }

                if ($projectId -le 0) {
                    [System.Windows.Forms.MessageBox]::Show('Select a search result or enter a project ID to continue.', 'Link Mod CF', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
                    return
                }

                $searchForm.UseWaitCursor = $true
                try {
                    $info = $dialogGetCurseForgeProject.Invoke($projectId)
                    $dialogSaveCurseForgeLink.Invoke($selected, $info)
                    [System.Windows.Forms.MessageBox]::Show(
                        ("Linked '{0}' to CurseForge project:`r`n`r`n{1} ({2})" -f $selected.DisplayName, $info.ProjectName, $info.ProjectId),
                        'Link Mod CF',
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Information
                    ) | Out-Null
                    $searchForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
                    $searchForm.Close()
                } catch {
                    $cfMsg = ''
                    $cfLine = ''
                    $cfStack = ''
                    try { $cfMsg = [string]$_.Exception.ToString() } catch { $cfMsg = [string]$_.Exception.Message }
                    try { $cfLine = [string]$_.InvocationInfo.PositionMessage } catch { $cfLine = '' }
                    try { $cfStack = [string]$_.ScriptStackTrace } catch { $cfStack = '' }
                    try { & $appendLog ("[ERROR] CurseForge link failed for project {0}: {1}" -f $projectId, $cfMsg) } catch { }
                    if (-not [string]::IsNullOrWhiteSpace($cfLine)) {
                        try { & $appendLog ("[ERROR] CurseForge link location: {0}" -f $cfLine) } catch { }
                    }
                    if (-not [string]::IsNullOrWhiteSpace($cfStack)) {
                        try { & $appendLog ("[ERROR] CurseForge link stack: {0}" -f $cfStack) } catch { }
                    }
                    try { _GuiModuleLog -Level ERROR -Message ("CurseForge link failed for project {0}: {1}" -f $projectId, $cfMsg) } catch { }
                    if (-not [string]::IsNullOrWhiteSpace($cfLine)) {
                        try { _GuiModuleLog -Level ERROR -Message ("CurseForge link location: {0}" -f $cfLine) } catch { }
                    }
                    if (-not [string]::IsNullOrWhiteSpace($cfStack)) {
                        try { _GuiModuleLog -Level ERROR -Message ("CurseForge link stack: {0}" -f $cfStack) } catch { }
                    }
                    $popupDetail = [string]$_.Exception.Message
                    if (-not [string]::IsNullOrWhiteSpace($cfLine)) {
                        $popupDetail = "{0}`r`n`r`nLocation:`r`n{1}" -f $popupDetail, $cfLine
                    }
                    [System.Windows.Forms.MessageBox]::Show(("Failed to link CurseForge mod.`r`n`r`n{0}" -f $popupDetail), 'Link Mod CF', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
                } finally {
                    $searchForm.UseWaitCursor = $false
                }
            }.GetNewClosure()

            $btnSearchCurseForge.Add_Click({ $performSearch.Invoke() }.GetNewClosure())
            $txtSearch.Add_KeyDown({
                param($sender, $e)
                if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
                    $performSearch.Invoke()
                    $e.SuppressKeyPress = $true
                }
            }.GetNewClosure())
            $txtProjectId.Add_TextChanged({ $syncLinkButtonState.Invoke() }.GetNewClosure())
            $lvResults.Add_SelectedIndexChanged({ $syncLinkButtonState.Invoke() }.GetNewClosure())
            $lvResults.Add_DoubleClick({ $completeLink.Invoke() }.GetNewClosure())
            $btnCancelLink.Add_Click({ $searchForm.Close() }.GetNewClosure())
            $btnConfirmLink.Add_Click({ $completeLink.Invoke() }.GetNewClosure())
            $searchForm.AcceptButton = $btnConfirmLink
            $searchForm.CancelButton = $btnCancelLink
            $searchForm.Add_Resize({ $layoutCurseForgeLinkDialog.Invoke() }.GetNewClosure())
            $searchForm.Add_Shown({
                $layoutCurseForgeLinkDialog.Invoke()
                $performSearch.Invoke()
                $txtSearch.SelectAll()
                $txtSearch.Focus()
            }.GetNewClosure())

            $searchForm.ShowDialog($form) | Out-Null
        }.GetNewClosure()

        $linkSelectedModCurseForge = {
            & $showCurseForgeLinkDialog
        }.GetNewClosure()

        $checkAllModUpdates = {
            $updates = 0
            foreach ($key in @($cfMetadata.Keys)) {
                try {
                    $meta = $cfMetadata[$key]
                    $info = $getCurseForgeProject.Invoke(([int]$meta.ProjectId))
                    $meta.LatestKnownFileId = $info.LatestFileId
                    $meta.LatestKnownName = $info.LatestFileName
                    $meta.LatestDownloadUrl = $info.DownloadUrl
                    $meta.WebsiteUrl = $info.WebsiteUrl
                    if ([string]$meta.FileId -ne [string]$info.LatestFileId) { $updates++ }
                } catch {
                    & $appendLog ("[WARN] Could not check updates for {0}: {1}" -f $key, $_.Exception.Message)
                }
            }
            & $saveJsonMap $cfMetadataPath $cfMetadata
            & $refreshModList
            [System.Windows.Forms.MessageBox]::Show(("{0} linked mod(s) appear to have updates." -f $updates), 'Check Updates', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        }.GetNewClosure()

        $updateSelectedMod = {
            $selected = & $getSelectedMod
            if ($null -eq $selected) { return }
            if (-not $cfMetadata.ContainsKey($selected.FileName)) {
                [System.Windows.Forms.MessageBox]::Show('Link this mod to CurseForge before checking for an update.', 'Update Mod', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
                return
            }
            try {
                $meta = $cfMetadata[$selected.FileName]
                $info = $getCurseForgeProject.Invoke(([int]$meta.ProjectId))
                if ([string]$meta.FileId -eq [string]$info.LatestFileId) {
                    [System.Windows.Forms.MessageBox]::Show('This mod already matches the latest CurseForge file.', 'Update Mod', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
                    return
                }
                $targetRoot = if ($selected.Status -eq 'Enabled') { $modsPath } else { $modsDisabledPath }
                $targetPath = Join-Path $targetRoot $info.LatestFileName
                Invoke-WebRequest -Uri $info.DownloadUrl -OutFile $targetPath -ErrorAction Stop
                if ($selected.FilePath -ne $targetPath -and (Test-Path -LiteralPath $selected.FilePath)) {
                    Remove-Item -LiteralPath $selected.FilePath -Force -ErrorAction Stop
                }
                $oldKey = $selected.FileName
                $cfMetadata[$info.LatestFileName] = @{
                    ProjectId         = $info.ProjectId
                    ProjectName       = $info.ProjectName
                    WebsiteUrl        = $info.WebsiteUrl
                    FileId            = $info.LatestFileId
                    LatestKnownFileId = $info.LatestFileId
                    LatestKnownName   = $info.LatestFileName
                    LatestDownloadUrl = $info.DownloadUrl
                }
                if ($oldKey -ne $info.LatestFileName) { $null = $cfMetadata.Remove($oldKey) }
                & $saveJsonMap $cfMetadataPath $cfMetadata
                & $appendLog ("[INFO] Updated Hytale mod {0} to {1}" -f $selected.DisplayName, $info.LatestFileName)
                & $refreshModList
            } catch {
                [System.Windows.Forms.MessageBox]::Show(("Failed to update the selected mod.`r`n`r`n{0}" -f $_.Exception.Message), 'Update Mod', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            }
        }.GetNewClosure()

        $btnOpenFolder.Add_Click({
            try {
                $targetFolder = ''
                if (-not [string]::IsNullOrWhiteSpace($txtDownloaderPath.Text)) {
                    $targetFolder = Split-Path -Path $txtDownloaderPath.Text -Parent
                }
                if ([string]::IsNullOrWhiteSpace($targetFolder) -or -not (Test-Path -LiteralPath $targetFolder)) {
                    & $showResultMessage 'Open Folder' 'ECC could not find the Hytale downloader folder.' $false
                    return
                }
                Start-Process -FilePath 'explorer.exe' -ArgumentList $targetFolder | Out-Null
            } catch {
                & $showResultMessage 'Open Folder' ("Failed to open the Hytale downloader folder.`r`n`r`n{0}" -f $_.Exception.Message) $false
            }
        }.GetNewClosure())
        $btnCheckFiles.Add_Click({
            $result = & $runUpdaterOp 'Verify Required Files' {
                & (Get-Command -Name 'Get-HytaleRequiredFilesStatus' -ErrorAction Stop) -Prefix $Prefix
            }
            & $applyFilesStatus $result
            if ($result.Success) {
                & $showResultMessage 'Required Files' 'All required Hytale files are present.' $true
            } else {
                & $showResultMessage 'Required Files' 'One or more required Hytale files are missing. See the updater tab for details.' $false
            }
        }.GetNewClosure())
        $btnCheckServerUpdate.Add_Click({
            $result = & $runUpdaterOp 'Check Server Update' {
                & (Get-Command -Name 'Get-HytaleServerUpdateStatus' -ErrorAction Stop) -Prefix $Prefix
            }
            $data = $null
            try { $data = $result.Data } catch { $data = $null }
            if ($data -and $data.UpdateAvailable -eq $true) {
                & $showResultMessage 'Update Available' ("Update Available!`r`n`r`nInstalled (Server): {0}`r`nZIP Package: {1}" -f $data.DownloaderVersion, $data.ZipVersion) $true
            } elseif ($data -and $data.UpdateAvailable -eq $false) {
                & $showResultMessage 'Up To Date' ("No update available.`r`n`r`nVersion: {0}" -f $data.DownloaderVersion) $true
            } else {
                & $showResultMessage 'Version Check' $(if ($result.Message) { [string]$result.Message } else { 'Version check did not return a usable result.' }) $false
            }
        }.GetNewClosure())
        $btnDownloaderVersion.Add_Click({
            $result = & $runUpdaterOp 'Downloader Version' {
                & (Get-Command -Name 'Invoke-HytaleDownloaderCommand' -ErrorAction Stop) -Prefix $Prefix -Arguments '-version' -Description 'Checking downloader version'
            }
            $output = ''
            try { $output = [string]$result.Data.Output } catch { $output = '' }
            if ([string]::IsNullOrWhiteSpace($output)) { $output = [string]$result.Message }
            & $showResultMessage 'Downloader Version' $output ($result.Success -eq $true)
        }.GetNewClosure())
        $btnCheckDownloaderUpdate.Add_Click({
            $result = & $runUpdaterOp 'Check Downloader Update' {
                & (Get-Command -Name 'Invoke-HytaleDownloaderCommand' -ErrorAction Stop) -Prefix $Prefix -Arguments '-check-update' -Description 'Checking for downloader updates'
            }
            $output = ''
            try { $output = [string]$result.Data.Output } catch { $output = '' }
            if ([string]::IsNullOrWhiteSpace($output)) { $output = [string]$result.Message }
            & $showResultMessage 'Downloader Update Check' $output ($result.Success -eq $true)
        }.GetNewClosure())
        $btnUpdateDownloader.Add_Click({
            $result = & $runUpdaterOp 'Update Downloader' {
                & (Get-Command -Name 'Update-HytaleDownloader' -ErrorAction Stop) -Prefix $Prefix
            }
            & $refreshDownloaderPath
            if ($result.Success -eq $true) {
                & $showResultMessage 'Update Downloader' 'Downloader updated successfully.' $true
            } else {
                & $showResultMessage 'Update Downloader' ("Downloader update failed.`r`n`r`n{0}" -f $result.Message) $false
            }
        }.GetNewClosure())
        $btnUpdateServer.Add_Click({
            $confirmText = 'This will download and install the latest Hytale server files.'
            if ($chkAutoRestart.Checked) {
                $confirmText += "`r`n`r`nECC will restart the server after the update if it was running."
            }
            $confirmText += "`r`n`r`nContinue?"
            $dialog = [System.Windows.Forms.MessageBox]::Show(
                $confirmText,
                'Confirm Hytale Update',
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            if ($dialog -ne [System.Windows.Forms.DialogResult]::Yes) { return }

            $result = & $runUpdaterOp 'Update Server' {
                & (Get-Command -Name 'Update-HytaleServerFiles' -ErrorAction Stop) -Prefix $Prefix -AutoRestartAfterUpdate ([bool]$chkAutoRestart.Checked)
            }
            & $refreshFilesSync
            & $refreshUpdateWarning
            if ($result.Success -eq $true) {
                $message = "Update completed successfully.`r`n`r`nNew files: $($result.Data.FilesCopied)`r`nUpdated files: $($result.Data.FilesUpdated)"
                if ($result.Data.Restarted -eq $true) { $message += "`r`n`r`nServer has been restarted." }
                & $showResultMessage 'Update Complete' $message $true
            } else {
                & $showResultMessage 'Update Failed' ("Hytale update failed.`r`n`r`n{0}" -f $result.Message) $false
            }
        }.GetNewClosure())
        $btnRefreshMods.Add_Click({ & $refreshModList; & $updateModSelectionUi }.GetNewClosure())
        $btnAddMod.Add_Click({
            if (-not (& $ensureModFolders)) {
                [System.Windows.Forms.MessageBox]::Show('ECC could not prepare the Hytale mod folders.', 'Add Mod', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
                return
            }

            $dialog = New-Object System.Windows.Forms.OpenFileDialog
            $dialog.Title = 'Add Hytale Mod'
            $dialog.Filter = 'Hytale Mod JAR (*.jar)|*.jar|All Files (*.*)|*.*'
            $dialog.Multiselect = $true
            $dialog.InitialDirectory = if (Test-Path -LiteralPath $modsPath) { $modsPath } else { $hytaleRoot }

            try {
                if ($dialog.ShowDialog($form) -ne [System.Windows.Forms.DialogResult]::OK) { return }
                foreach ($source in @($dialog.FileNames)) {
                    if ([string]::IsNullOrWhiteSpace($source) -or -not (Test-Path -LiteralPath $source -PathType Leaf)) { continue }
                    if (-not ([string]$source).ToLowerInvariant().EndsWith('.jar')) { continue }
                    $fileName = [System.IO.Path]::GetFileName($source)
                    Copy-Item -LiteralPath $source -Destination (Join-Path $modsPath $fileName) -Force
                    & $appendLog ("[INFO] Added Hytale mod {0}" -f $fileName)
                }
                & $refreshModList
                & $updateModSelectionUi
            } catch {
                [System.Windows.Forms.MessageBox]::Show(("Failed to add the selected mod file(s).`r`n`r`n{0}" -f $_.Exception.Message), 'Add Mod', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            } finally {
                try { $dialog.Dispose() } catch { }
            }
        }.GetNewClosure())
        $btnToggleMod.Add_Click({
            $selected = & $getSelectedMod
            if ($null -eq $selected) { return }
            & $ensureModFolders | Out-Null
            $destinationRoot = if ($selected.Status -eq 'Enabled') { $modsDisabledPath } else { $modsPath }
            Move-Item -LiteralPath $selected.FilePath -Destination (Join-Path $destinationRoot $selected.FileName) -Force
            & $appendLog ("[INFO] Toggled Hytale mod {0}" -f $selected.FileName)
            & $refreshModList
        }.GetNewClosure())
        $btnOpenModsFolder.Add_Click({ if (& $ensureModFolders) { Start-Process -FilePath 'explorer.exe' -ArgumentList $modsPath | Out-Null } }.GetNewClosure())
        $btnDeleteMod.Add_Click({
            $selected = & $getSelectedMod
            if ($null -eq $selected) { return }
            Remove-Item -LiteralPath $selected.FilePath -Force
            $null = $cfMetadata.Remove($selected.FileName)
            $null = $modNotes.Remove($selected.FileName)
            & $saveJsonMap $cfMetadataPath $cfMetadata
            & $saveJsonMap $modNotesPath $modNotes
            & $appendLog ("[INFO] Deleted Hytale mod {0}" -f $selected.FileName)
            & $refreshModList
            & $updateModSelectionUi
        }.GetNewClosure())
        $btnCheckConflicts.Add_Click({
            $groups = @{}
            foreach ($item in @($lvMods.Items)) {
                $tag = $item.Tag
                if ($null -eq $tag) { continue }
                $key = & $normalizeLooseName $tag.DisplayName
                if (-not $groups.ContainsKey($key)) { $groups[$key] = @() }
                $groups[$key] += $tag.FileName
            }
            $conflicts = @($groups.GetEnumerator() | Where-Object { $_.Value.Count -gt 1 } | ForEach-Object { ($_.Value -join '; ') })
            if ($conflicts.Count -le 0) {
                [System.Windows.Forms.MessageBox]::Show('No obvious duplicate Hytale mods were detected.', 'Check Conflicts', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            } else {
                [System.Windows.Forms.MessageBox]::Show(("Potential mod conflicts found:`r`n`r`n{0}" -f ($conflicts -join "`r`n")), 'Check Conflicts', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            }
        }.GetNewClosure())
        $btnOpenSelectedMod.Add_Click({
            $selected = & $getSelectedMod
            if ($selected) { Start-Process -FilePath 'explorer.exe' -ArgumentList ('/select,"{0}"' -f $selected.FilePath) | Out-Null }
        }.GetNewClosure())
        $btnOpenConfigFolder.Add_Click({
            $selected = & $getSelectedMod
            if ($null -eq $selected) { return }
            $needle = & $normalizeLooseName $selected.DisplayName
            foreach ($base in @($modsPath, $modsDisabledPath, $hytaleRoot)) {
                if (-not (Test-Path -LiteralPath $base)) { continue }
                $match = Get-ChildItem -LiteralPath $base -Directory -ErrorAction SilentlyContinue | Where-Object { (& $normalizeLooseName $_.Name) -like "*$needle*" } | Select-Object -First 1
                if ($match) { Start-Process -FilePath 'explorer.exe' -ArgumentList $match.FullName | Out-Null; return }
            }
            [System.Windows.Forms.MessageBox]::Show('ECC could not find a likely config folder for the selected mod.', 'Browse Configs', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        }.GetNewClosure())
        $btnLinkCurseForge.Add_Click({ & $linkSelectedModCurseForge }.GetNewClosure())
        $btnCheckModUpdates.Add_Click({ & $checkAllModUpdates }.GetNewClosure())
        $btnUpdateSelectedMod.Add_Click({ & $updateSelectedMod }.GetNewClosure())
        $btnOpenModPage.Add_Click({
            $selected = & $getSelectedMod
            if ($selected -and $cfMetadata.ContainsKey($selected.FileName)) { Start-Process $cfMetadata[$selected.FileName].WebsiteUrl | Out-Null }
        }.GetNewClosure())
        $btnGetMoreMods.Add_Click({ Start-Process 'https://www.curseforge.com/hytale/server-mods' | Out-Null }.GetNewClosure())
        $btnSaveModNotes.Add_Click({
            $selected = & $getSelectedMod
            if ($selected) {
                $modNotes[$selected.FileName] = [string]$txtModNotes.Text
                & $saveJsonMap $modNotesPath $modNotes
                & $appendLog ("[INFO] Saved notes for {0}" -f $selected.DisplayName)
            }
        }.GetNewClosure())
        $lvMods.Add_SelectedIndexChanged({ & $updateModSelectionUi }.GetNewClosure())
        $lvMods.Add_DragEnter({
            if ($_.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
                $_.Effect = [System.Windows.Forms.DragDropEffects]::Copy
            } else {
                $_.Effect = [System.Windows.Forms.DragDropEffects]::None
            }
        }.GetNewClosure())
        $lvMods.Add_DragDrop({
            if (-not (& $ensureModFolders)) { return }
            foreach ($source in @($_.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop))) {
                if ([string]$source -match '\.jar$') {
                    Copy-Item -LiteralPath $source -Destination (Join-Path $modsPath ([System.IO.Path]::GetFileName($source))) -Force
                    & $appendLog ("[INFO] Added Hytale mod {0}" -f ([System.IO.Path]::GetFileName($source)))
                }
            }
            & $refreshModList
        }.GetNewClosure())

        $form.Add_Resize({ & $layoutHytaleManager }.GetNewClosure())
        $tabMain.Add_SelectedIndexChanged({ & $layoutHytaleManager }.GetNewClosure())

        $form.Add_Shown({
            & $layoutHytaleManager
            & $refreshDownloaderPath
            & $refreshUpdateWarning
            & $refreshFilesSync
            if ($modsFeatureReady) {
                & $ensureModFolders | Out-Null
                $loadedCfMetadata = & $loadJsonMap $cfMetadataPath
                $loadedModNotes = & $loadJsonMap $modNotesPath

                $cfMetadata.Clear()
                foreach ($key in @($loadedCfMetadata.Keys)) {
                    $cfMetadata[$key] = $loadedCfMetadata[$key]
                }

                $modNotes.Clear()
                foreach ($key in @($loadedModNotes.Keys)) {
                    $modNotes[$key] = $loadedModNotes[$key]
                }
                & $refreshModList
                & $updateModSelectionUi
            } else {
                foreach ($ctrl in @($btnRefreshMods, $btnAddMod, $btnToggleMod, $btnOpenModsFolder, $btnDeleteMod, $btnCheckConflicts, $btnOpenSelectedMod, $btnOpenConfigFolder, $btnLinkCurseForge, $btnCheckModUpdates, $btnUpdateSelectedMod, $btnOpenModPage, $btnGetMoreMods, $btnSaveModNotes, $lvMods, $txtModNotes)) {
                    try { $ctrl.Enabled = $false } catch { }
                }
                & $appendLog '[WARN] Hytale mod-manager tools are unavailable because ECC could not resolve the profile root folder.'
            }
        }.GetNewClosure())

        $form.ShowDialog() | Out-Null
    }

    # =====================================================================
    # MAIN FORM
    # =====================================================================
    $form               = [System.Windows.Forms.Form]::new()
    $script:MainForm    = $form
    $form.Text          = "Etherium Command Center $($script:AppVersion)"
    $initialWindowRect  = _ClampWindowRectToWorkingArea -Width $defaultWidth -Height $defaultHeight
    $workingAreaRect    = _GetWorkingAreaRect
    $form.Size          = [System.Drawing.Size]::new($initialWindowRect.Width, $initialWindowRect.Height)
    $form.MinimumSize   = [System.Drawing.Size]::new([Math]::Min($minWidth, $workingAreaRect.Width), [Math]::Min($minHeight, $workingAreaRect.Height))
    $form.BackColor     = $clrBg
    $form.StartPosition = 'Manual'
    $form.Location      = [System.Drawing.Point]::new($initialWindowRect.X, $initialWindowRect.Y)
    $form.Icon          = [System.Drawing.SystemIcons]::Application
    $form.KeyPreview    = $true
    # Remove the native title bar - we draw our own chrome in the top bar
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
    $form.Padding         = [System.Windows.Forms.Padding]::new($windowMargin)

    $savedWindow = _GetReloadWindowBounds
    if (-not $savedWindow) {
        $savedWindow = _GetSavedWindowBounds
    }
    if ($savedWindow) {
        $form.Size = [System.Drawing.Size]::new($savedWindow.Width, $savedWindow.Height)
        if ($savedWindow.HasPos) {
            $form.Location = [System.Drawing.Point]::new($savedWindow.X, $savedWindow.Y)
        }
        if ($savedWindow.State -eq 'Maximized') {
            $form.WindowState = 'Maximized'
        }
    }

    $script:_MainToolTip = New-Object System.Windows.Forms.ToolTip
    $script:_MainToolTip.AutoPopDelay = 12000
    $script:_MainToolTip.InitialDelay = 350
    $script:_MainToolTip.ReshowDelay  = 150
    $script:_MainToolTip.ShowAlways   = $true

    # ── Borderless resize via WM_NCHITTEST intercept ──────────────────────────
    # Child controls cover the whole form so $form.MouseDown never fires.
    # The correct fix is to intercept WM_NCHITTEST in WndProc and return the
    # right HTXXX value — Windows then handles the resize drag natively,
    # including MinimumSize enforcement and snap-to-edge.
    # We compile a NativeWindow subclass once (guarded by type check) and
    # assign it to the form after the form handle is created.
    try { [BorderlessResizeHook] | Out-Null } catch {
        Add-Type -ReferencedAssemblies 'System.Windows.Forms','System.Drawing' @"
using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

public class BorderlessResizeHook : NativeWindow {
    private const int WM_NCHITTEST   = 0x0084;
    private const int WM_NCLBUTTONDOWN = 0x00A1;
    private const int HTCLIENT       = 1;
    private const int HTLEFT         = 10;
    private const int HTRIGHT        = 11;
    private const int HTTOP          = 12;
    private const int HTTOPLEFT      = 13;
    private const int HTTOPRIGHT     = 14;
    private const int HTBOTTOM       = 15;
    private const int HTBOTTOMLEFT   = 16;
    private const int HTBOTTOMRIGHT  = 17;

    private readonly Form _form;
    private readonly int  _grip;

    public BorderlessResizeHook(Form form, int gripSize) {
        _form = form;
        _grip = gripSize;
        AssignHandle(form.Handle);
    }

    protected override void WndProc(ref Message m) {
        if (m.Msg == WM_NCHITTEST) {
            base.WndProc(ref m);
            // Only intercept if Windows said HTCLIENT — leave caption/menu alone
            if (m.Result == (IntPtr)HTCLIENT) {
                Point screen = new Point(m.LParam.ToInt32());
                Point client = _form.PointToClient(screen);
                int x = client.X, y = client.Y;
                int w = _form.ClientSize.Width, h = _form.ClientSize.Height;
                int g = _grip;

                bool onL = x <= g, onR = x >= w - g;
                bool onT = y <= g, onB = y >= h - g;

                if      (onL && onT) m.Result = (IntPtr)HTTOPLEFT;
                else if (onR && onT) m.Result = (IntPtr)HTTOPRIGHT;
                else if (onL && onB) m.Result = (IntPtr)HTBOTTOMLEFT;
                else if (onR && onB) m.Result = (IntPtr)HTBOTTOMRIGHT;
                else if (onT)        m.Result = (IntPtr)HTTOP;
                else if (onB)        m.Result = (IntPtr)HTBOTTOM;
                else if (onL)        m.Result = (IntPtr)HTLEFT;
                else if (onR)        m.Result = (IntPtr)HTRIGHT;
            }
            return;
        }
        base.WndProc(ref m);
    }
}
"@
    }

    # Attach the hook once the form handle exists (HandleCreated fires before Shown)
    $form.Add_HandleCreated({
        try {
            $script:_ResizeHook = New-Object BorderlessResizeHook($form, $windowMargin)
        } catch { }
    })

    # Hide the admin console after the GUI is visible when debug is off
    $form.Add_Shown({
        try {
            _MarkGuiActivity 'shown'
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
    $form.Add_Activated({ _MarkGuiActivity 'activated' }.GetNewClosure())
    $form.Add_ResizeBegin({ _MarkGuiActivity 'resize_begin' }.GetNewClosure())
    $form.Add_ResizeEnd({ _MarkGuiActivity 'resize_end' }.GetNewClosure())

    # =====================================================================
    # STATUS BAR
    # =====================================================================
    $footerResizeGutter  = 10
    $rightResizeGutter   = 12
    $statusBar           = [System.Windows.Forms.StatusStrip]::new()
    $statusBar.BackColor = $clrPanel
    $statusBar.SizingGrip = $false
    $statusBar.AutoSize = $false
    $statusBar.Height = 22
    $statusLabel         = [System.Windows.Forms.ToolStripStatusLabel]::new()
    $statusLabel.Text      = 'Ready'
    $statusLabel.ForeColor = $clrText
    $statusLabel.Font      = $fontLabel
    $statusBar.Items.Add($statusLabel) | Out-Null
    $statusBar.Dock = 'None'
    $statusBar.Anchor = 'Left,Right,Bottom'
    $statusBar.Location = [System.Drawing.Point]::new(0, $defaultHeight - $statusBar.Height)
    $statusBar.Width = $defaultWidth
    $form.Controls.Add($statusBar)

    function _QueueStatusMessage {
        param([string]$Message)
        if ($script:SharedState -and $script:SharedState.ContainsKey('LogQueue')) {
            $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][INFO][GUI] $Message")
        }
    }

    function _SendDiscordNotice {
        param([string]$Message)

        if ([string]::IsNullOrWhiteSpace($Message) -or -not $script:SharedState) { return }

        if ($script:SharedState.ContainsKey('DiscordOutbox') -and $script:SharedState.DiscordOutbox) {
            $script:SharedState.DiscordOutbox.Enqueue($Message)
            return
        }
        if ($script:SharedState.ContainsKey('WebhookQueue') -and $script:SharedState.WebhookQueue) {
            $script:SharedState.WebhookQueue.Enqueue($Message)
        }
    }

    function _ReloadProfilesFromDisk {
        $loaded = Get-AllProfiles -ProfilesDir $script:ProfilesDir
        if ($null -eq $loaded) { $loaded = @{} }

        if (-not $script:SharedState.ContainsKey('Profiles') -or $null -eq $script:SharedState.Profiles) {
            $script:SharedState['Profiles'] = [hashtable]::Synchronized(@{})
        }

        $profilesTable = $script:SharedState.Profiles
        foreach ($key in @($profilesTable.Keys)) { $profilesTable.Remove($key) }
        foreach ($key in @($loaded.Keys)) { $profilesTable[$key] = $loaded[$key] }

        foreach ($pfx in @($script:SharedState.RunningServers.Keys)) {
            if (-not $profilesTable.ContainsKey($pfx)) {
                $script:SharedState.RunningServers.Remove($pfx)
            }
        }

        return $profilesTable.Count
    }

    function _ReloadCommandsAndProfiles {
        $script:CommandCatalog = $null
        $catalog = _LoadCommandCatalog -Path $script:CommandCatalogPath
        $profileCount = _ReloadProfilesFromDisk

        if ([string]::IsNullOrWhiteSpace($script:_SelectedProfilePrefix) -or
            -not $script:SharedState.Profiles.ContainsKey($script:_SelectedProfilePrefix)) {
            $script:_SelectedProfilePrefix = $null
        }

        _BuildProfilesList
        if ($script:_SelectedProfilePrefix -and $script:SharedState.Profiles.ContainsKey($script:_SelectedProfilePrefix)) {
            _BuildProfileEditor -Profile $script:SharedState.Profiles[$script:_SelectedProfilePrefix]
        } else {
            _BuildProfileEditor $null
        }
        _BuildServerDashboard

        $cmdCount = 0
        try {
            if ($catalog -and $catalog.Games) { $cmdCount = @($catalog.Games.Keys).Count }
        } catch { $cmdCount = 0 }

        $statusLabel.Text = "Commands reloaded. Profiles: $profileCount  |  Catalog entries: $cmdCount  |  $(Get-Date -Format 'HH:mm:ss')"
        _QueueStatusMessage "Commands and profiles reloaded from disk without resetting running server timers."
    }

    # =====================================================================
    # TOP BAR  (also acts as the drag handle for moving the borderless form)
    # =====================================================================
    $shellInset = [Math]::Max(3, $windowMargin - 4)

    $windowShell = _Panel $shellInset $shellInset ($defaultWidth - ($shellInset * 2)) ($defaultHeight - $statusBar.Height - ($shellInset * 2)) $clrShell
    $windowShell.Anchor = 'Top,Left,Right,Bottom'
    $windowShell.BorderStyle = 'None'
    $form.Controls.Add($windowShell)
    $windowShell.SendToBack()
    $script:_WindowShellPanel = $windowShell

    $windowEdgeTop = _Panel 0 0 $defaultWidth 2 $clrEdgeGlow
    $windowEdgeTop.Anchor = 'Top,Left'
    $windowEdgeTop.BorderStyle = 'None'
    $form.Controls.Add($windowEdgeTop)
    $script:_WindowEdgeTop = $windowEdgeTop

    $windowEdgeLeft = _Panel 0 0 2 $defaultHeight $clrEdge
    $windowEdgeLeft.Anchor = 'Top,Left'
    $windowEdgeLeft.BorderStyle = 'None'
    $form.Controls.Add($windowEdgeLeft)
    $script:_WindowEdgeLeft = $windowEdgeLeft

    $windowEdgeRight = _Panel ($defaultWidth - 3) 0 3 $defaultHeight $clrEdgeGlow
    $windowEdgeRight.Anchor = 'Top,Left'
    $windowEdgeRight.BorderStyle = 'None'
    $form.Controls.Add($windowEdgeRight)
    $script:_WindowEdgeRight = $windowEdgeRight

    $windowEdgeBottom = _Panel 0 ($defaultHeight - 2) $defaultWidth 2 $clrEdge
    $windowEdgeBottom.Anchor = 'Top,Left'
    $windowEdgeBottom.BorderStyle = 'None'
    $form.Controls.Add($windowEdgeBottom)
    $script:_WindowEdgeBottom = $windowEdgeBottom

    $resizeHandleThickness = [Math]::Max(8, $windowMargin)
    $resizeDebugBottomColor = [System.Drawing.Color]::FromArgb(1, $clrShell)
    $resizeDebugRightColor = $clrPanelSoft
    $resizeDebugCornerColor = [System.Drawing.Color]::Transparent

    $resizeBottomGrip = _Panel 0 ($defaultHeight - $statusBar.Height - $footerResizeGutter) $defaultWidth $footerResizeGutter $resizeDebugBottomColor
    $resizeBottomGrip.Anchor = 'Top,Left'
    $resizeBottomGrip.BorderStyle = 'None'
    $resizeBottomGrip.Cursor = [System.Windows.Forms.Cursors]::SizeNS
    $form.Controls.Add($resizeBottomGrip)
    $script:_ResizeBottomGrip = $resizeBottomGrip
    $resizeBottomMarker = _Panel ([Math]::Max(20, [int](($defaultWidth - 160) / 2))) ([Math]::Max(1, [int](($footerResizeGutter - 4) / 2))) 160 4 $clrEdgeGlow
    $resizeBottomMarker.Anchor = 'Top'
    $resizeBottomMarker.BorderStyle = 'None'
    $resizeBottomGrip.Controls.Add($resizeBottomMarker)
    $script:_ResizeBottomMarker = $resizeBottomMarker
    $resizeBottomMarker.Visible = $true

    $resizeLeftGrip = _Panel 0 0 $resizeHandleThickness ($defaultHeight - $statusBar.Height) ([System.Drawing.Color]::FromArgb(1, $clrShell))
    $resizeLeftGrip.Anchor = 'Top,Left'
    $resizeLeftGrip.BorderStyle = 'None'
    $resizeLeftGrip.Cursor = [System.Windows.Forms.Cursors]::SizeWE
    $form.Controls.Add($resizeLeftGrip)
    $script:_ResizeLeftGrip = $resizeLeftGrip

    $resizeRightGrip = _Panel ($defaultWidth - $rightResizeGutter) 0 $rightResizeGutter ($defaultHeight - $statusBar.Height) $resizeDebugRightColor
    $resizeRightGrip.Anchor = 'Top,Left'
    $resizeRightGrip.BorderStyle = 'None'
    $resizeRightGrip.Cursor = [System.Windows.Forms.Cursors]::SizeWE
    $form.Controls.Add($resizeRightGrip)
    $script:_ResizeRightGrip = $resizeRightGrip
    $resizeRightEdgeMarker = _Panel ([Math]::Max(0, $rightResizeGutter - 3)) 0 3 ($defaultHeight - $statusBar.Height) $clrEdgeGlow
    $resizeRightEdgeMarker.BorderStyle = 'None'
    $resizeRightGrip.Controls.Add($resizeRightEdgeMarker)
    $script:_ResizeRightEdgeMarker = $resizeRightEdgeMarker
    $initialRightMarkerX = [Math]::Max(0, [int](($rightResizeGutter - 4) / 2))
    $resizeRightMarkerTop = _Panel $initialRightMarkerX 120 4 80 $clrEdgeGlow
    $resizeRightMarkerTop.BorderStyle = 'None'
    $resizeRightGrip.Controls.Add($resizeRightMarkerTop)
    $resizeRightMarkerMid = _Panel $initialRightMarkerX 220 4 80 $clrEdgeGlow
    $resizeRightMarkerMid.BorderStyle = 'None'
    $resizeRightGrip.Controls.Add($resizeRightMarkerMid)
    $resizeRightMarkerBot = _Panel $initialRightMarkerX 320 4 80 $clrEdgeGlow
    $resizeRightMarkerBot.BorderStyle = 'None'
    $resizeRightGrip.Controls.Add($resizeRightMarkerBot)
    $script:_ResizeRightMarkers = @($resizeRightMarkerTop, $resizeRightMarkerMid, $resizeRightMarkerBot)

    $resizeBottomLeftGrip = _Panel 0 ($defaultHeight - $statusBar.Height - $footerResizeGutter) ($resizeHandleThickness * 2) $footerResizeGutter $resizeDebugCornerColor
    $resizeBottomLeftGrip.Anchor = 'Top,Left'
    $resizeBottomLeftGrip.BorderStyle = 'None'
    $resizeBottomLeftGrip.Cursor = [System.Windows.Forms.Cursors]::SizeNESW
    $form.Controls.Add($resizeBottomLeftGrip)
    $script:_ResizeBottomLeftGrip = $resizeBottomLeftGrip

    $resizeBottomRightGrip = _Panel ($defaultWidth - $rightResizeGutter) ($defaultHeight - $statusBar.Height - $footerResizeGutter) $rightResizeGutter $footerResizeGutter $resizeDebugCornerColor
    $resizeBottomRightGrip.Anchor = 'Top,Left'
    $resizeBottomRightGrip.BorderStyle = 'None'
    $resizeBottomRightGrip.Cursor = [System.Windows.Forms.Cursors]::SizeNWSE
    $form.Controls.Add($resizeBottomRightGrip)
    $script:_ResizeBottomRightGrip = $resizeBottomRightGrip
    $resizeCornerLabel = _Label '//' 0 0 12 12 $fontBold
    $resizeCornerLabel.ForeColor = $clrYellow
    $resizeCornerLabel.TextAlign = 'MiddleCenter'
    $resizeCornerLabel.BackColor = [System.Drawing.Color]::Transparent
    $resizeCornerLabel.Cursor = [System.Windows.Forms.Cursors]::SizeNWSE
    $resizeBottomRightGrip.Controls.Add($resizeCornerLabel)
    $script:_ResizeCornerLabel = $resizeCornerLabel
    $resizeBottomLabel = _Label '' 0 0 1 1 $fontBold
    $resizeBottomLabel.ForeColor = $clrBtnText
    $resizeBottomLabel.BackColor = [System.Drawing.Color]::Transparent
    $resizeBottomLabel.TextAlign = 'MiddleLeft'
    $resizeBottomLabel.Cursor = [System.Windows.Forms.Cursors]::SizeNS
    $resizeBottomGrip.Controls.Add($resizeBottomLabel)
    $script:_ResizeBottomLabel = $resizeBottomLabel
    $resizeRightLabel = _Label '' 0 0 1 1 $fontBold
    $resizeRightLabel.ForeColor = $clrBtnText
    $resizeRightLabel.BackColor = [System.Drawing.Color]::Transparent
    $resizeRightLabel.TextAlign = 'MiddleCenter'
    $resizeRightLabel.Cursor = [System.Windows.Forms.Cursors]::SizeWE
    $resizeRightGrip.Controls.Add($resizeRightLabel)
    $script:_ResizeRightLabel = $resizeRightLabel

    $layoutResizeChromeElements = {
        param(
            [int]$ClientWidth,
            [int]$ClientHeight,
            [int]$StatusBarHeight = 0
        )

        $cwCurrent = [Math]::Max(0, $ClientWidth)
        $chCurrent = [Math]::Max(0, $ClientHeight)
        $sbhCurrent = [Math]::Max(0, $StatusBarHeight)
        $bottomGripHeight = [Math]::Max(4, $footerResizeGutter)
        $rightGripWidth = [Math]::Max(6, $rightResizeGutter)
        $bottomTop = [Math]::Max(0, $chCurrent - $sbhCurrent - $bottomGripHeight)
        $rightLeft = [Math]::Max(0, $cwCurrent - $rightGripWidth)
        $verticalHeight = [Math]::Max(0, $chCurrent - $sbhCurrent)
        $cornerHeight = [Math]::Max($bottomGripHeight, $sbhCurrent)
        $cornerTop = [Math]::Max(0, $chCurrent - $cornerHeight)
        $gripSquare = [Math]::Max($resizeHandleThickness * 2, $rightGripWidth)

        if ($script:_WindowEdgeTop) {
            $script:_WindowEdgeTop.SetBounds(0, 0, $cwCurrent, 2)
        }
        if ($script:_WindowEdgeLeft) {
            $script:_WindowEdgeLeft.SetBounds(0, 0, 2, $chCurrent)
        }
        if ($script:_WindowEdgeRight) {
            $script:_WindowEdgeRight.SetBounds([Math]::Max(0, $cwCurrent - 3), 0, 3, $chCurrent)
        }
        if ($script:_WindowEdgeBottom) {
            $script:_WindowEdgeBottom.SetBounds(0, [Math]::Max(0, $chCurrent - 2), $cwCurrent, 2)
        }
        if ($script:_ResizeBottomGrip) {
            $script:_ResizeBottomGrip.SetBounds(0, $bottomTop, $cwCurrent, $bottomGripHeight)
            if ($script:_ResizeBottomMarker) {
                $markerWidth = [Math]::Min(160, [Math]::Max(56, $script:_ResizeBottomGrip.Width - 120))
                $markerX = [Math]::Max(12, [int](($script:_ResizeBottomGrip.Width - $markerWidth) / 2))
                $markerY = [Math]::Max(1, [int](($script:_ResizeBottomGrip.Height - 4) / 2))
                $script:_ResizeBottomMarker.SetBounds($markerX, $markerY, $markerWidth, 4)
                $script:_ResizeBottomMarker.Visible = $true
            }
            if ($script:_ResizeBottomLabel) {
                $script:_ResizeBottomLabel.Location = [System.Drawing.Point]::new(0, 0)
                $script:_ResizeBottomLabel.Size = [System.Drawing.Size]::new(1, 1)
            }
        }
        if ($script:_ResizeLeftGrip) {
            $script:_ResizeLeftGrip.SetBounds(0, 0, $resizeHandleThickness, $chCurrent)
        }
        if ($script:_ResizeRightGrip) {
            $script:_ResizeRightGrip.SetBounds($rightLeft, 0, $rightGripWidth, $verticalHeight)
            if ($script:_ResizeRightEdgeMarker) {
                $script:_ResizeRightEdgeMarker.SetBounds([Math]::Max(0, $script:_ResizeRightGrip.Width - 3), 0, 3, $script:_ResizeRightGrip.Height)
            }
            if ($script:_ResizeRightMarkers) {
                $markerX = [Math]::Max(0, [int](($script:_ResizeRightGrip.Width - 4) / 2))
                $markerCount = [Math]::Max(1, $script:_ResizeRightMarkers.Count)
                $availableHeight = [Math]::Max(96, $script:_ResizeRightGrip.Height - 72)
                $markerHeight = [Math]::Max(28, [Math]::Min(72, [int]($availableHeight / (($markerCount * 2) + 1))))
                $gapY = [Math]::Max(10, [int](($script:_ResizeRightGrip.Height - ($markerHeight * $markerCount)) / ($markerCount + 1)))
                if (($gapY * ($markerCount + 1)) + ($markerHeight * $markerCount) -gt $script:_ResizeRightGrip.Height) {
                    $gapY = [Math]::Max(6, [int](($script:_ResizeRightGrip.Height - ($markerHeight * $markerCount)) / ($markerCount + 1)))
                }
                for ($markerIndex = 0; $markerIndex -lt $script:_ResizeRightMarkers.Count; $markerIndex++) {
                    $marker = $script:_ResizeRightMarkers[$markerIndex]
                    if ($marker) {
                        $markerTop = $gapY + ($markerIndex * ($markerHeight + $gapY))
                        $marker.SetBounds($markerX, [Math]::Max(0, $markerTop), 4, [Math]::Max(18, $markerHeight))
                    }
                }
            }
            if ($script:_ResizeRightLabel) {
                $script:_ResizeRightLabel.Location = [System.Drawing.Point]::new(0, 0)
                $script:_ResizeRightLabel.Size = [System.Drawing.Size]::new(1, 1)
            }
        }
        if ($script:_ResizeBottomLeftGrip) {
            $script:_ResizeBottomLeftGrip.SetBounds(0, $cornerTop, $gripSquare, $cornerHeight)
        }
        if ($script:_ResizeBottomRightGrip) {
            $script:_ResizeBottomRightGrip.SetBounds($rightLeft, $cornerTop, $rightGripWidth, $cornerHeight)
            if ($script:_ResizeCornerLabel) {
                $script:_ResizeCornerLabel.Location = [System.Drawing.Point]::new(0, [Math]::Max(0, $script:_ResizeBottomRightGrip.Height - 12))
                $script:_ResizeCornerLabel.Size = [System.Drawing.Size]::new([Math]::Max(8, $script:_ResizeBottomRightGrip.Width), 12)
            }
        }
    }.GetNewClosure()

    $topBar        = _Panel $windowMargin $windowMargin ($defaultWidth - ($windowMargin * 2)) $topBarHeight $clrPanel
    $topBar.Anchor = 'Top,Left,Right'
    $form.Controls.Add($topBar)

    # -- Drag-to-move state --
    $script:_DragActive = $false
    $script:_DragOrigin = [System.Drawing.Point]::new(0, 0)

    $topBar.Add_MouseDown({
        param($s, $e)
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
            $edge = [Math]::Max(10, $resizeHandleThickness)
            $corner = $edge * 2
            if ($e.Y -le $edge -or $e.X -ge ($topBar.ClientSize.Width - $edge) -or ($e.X -ge ($topBar.ClientSize.Width - $corner) -and $e.Y -le $corner)) {
                return
            }
            $script:_DragActive = $true
            $screenPt = $topBar.PointToScreen($e.Location)
            $script:_DragOrigin = [System.Drawing.Point]::new(
                $screenPt.X - $form.Location.X,
                $screenPt.Y - $form.Location.Y
            )
        }
    })
    $topBar.Add_MouseMove({
        param($s, $e)
        if ($script:_DragActive) {
            $screenPt = $topBar.PointToScreen($e.Location)
            $form.Location = [System.Drawing.Point]::new(
                $screenPt.X - $script:_DragOrigin.X,
                $screenPt.Y - $script:_DragOrigin.Y
            )
        }
    })
    $topBar.Add_MouseUp({
        param($s, $e)
        $script:_DragActive = $false
    })

    $beginBorderlessResize = {
        param([int]$hitTarget)
        try {
            [NativeWin]::ReleaseCapture() | Out-Null
            [NativeWin]::SendMessage($form.Handle, 0x00A1, [IntPtr]$hitTarget, [IntPtr]::Zero) | Out-Null
        } catch { }
    }.GetNewClosure()

    $writeResizeTrace = {
        param(
            [string]$Area,
            [string]$Detail = ''
        )
        try {
            if (-not (_IsGuiDebugEnabled)) { return }
            if (-not $script:SharedState -or -not $script:SharedState.LogQueue) { return }
            $key = "RESIZE_TRACE::{0}" -f $Area
            $now = [DateTime]::UtcNow
            if (-not $script:SharedState.ContainsKey($key) -or (($now - [DateTime]$script:SharedState[$key]).TotalSeconds -ge 2)) {
                $script:SharedState[$key] = $now
                $msg = if ([string]::IsNullOrWhiteSpace($Detail)) {
                    "RESIZE $Area"
                } else {
                    "RESIZE $Area :: $Detail"
                }
                $script:SharedState.LogQueue.Enqueue("[{0}][INFO][GUI] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $msg)
            }
        } catch { }
    }.GetNewClosure()

    $normalizeResizeChrome = {
        try {
            $cwLive = [Math]::Max(0, $form.ClientSize.Width)
            $chLive = [Math]::Max(0, $form.ClientSize.Height)
            $sbhLive = 0
            try { if ($statusBar) { $sbhLive = [Math]::Max(0, $statusBar.Height) } } catch { $sbhLive = 0 }
            & $layoutResizeChromeElements $cwLive $chLive $sbhLive
            try { if ($statusBar) { $statusBar.BringToFront() } } catch { }
            try { if ($script:_ResizeBottomGrip) { $script:_ResizeBottomGrip.BringToFront() } } catch { }
            try { if ($script:_ResizeRightGrip) { $script:_ResizeRightGrip.BringToFront() } } catch { }
            try { if ($script:_ResizeBottomLeftGrip) { $script:_ResizeBottomLeftGrip.BringToFront() } } catch { }
            try { if ($script:_ResizeBottomRightGrip) { $script:_ResizeBottomRightGrip.BringToFront() } } catch { }
            try { if ($script:_ResizeCornerLabel) { $script:_ResizeCornerLabel.BringToFront() } } catch { }
        } catch { }
    }.GetNewClosure()

    $handleProxyResizeMouseDown = {
        param(
            $control,
            $mouseEvent,
            [bool]$allowTop = $false,
            [bool]$allowRight = $false,
            [bool]$allowBottom = $false,
            [bool]$allowLeft = $false
        )
        if (-not $control -or -not $mouseEvent) { return }
        if ($mouseEvent.Button -ne [System.Windows.Forms.MouseButtons]::Left) { return }

        $edge = [Math]::Max(10, $resizeHandleThickness)
        $corner = $edge * 2
        $x = $mouseEvent.X
        $y = $mouseEvent.Y
        $w = $control.ClientSize.Width
        $h = $control.ClientSize.Height

        $onLeft = $allowLeft -and ($x -le $edge)
        $onRight = $allowRight -and ($x -ge ($w - $edge))
        $onTop = $allowTop -and ($y -le $edge)
        $onBottom = $allowBottom -and ($y -ge ($h - $edge))

        if ($allowTop -and $allowRight -and ($x -ge ($w - $corner)) -and ($y -le $corner)) {
            & $beginBorderlessResize 14
            return
        }
        if ($allowBottom -and $allowRight -and ($x -ge ($w - $corner)) -and ($y -ge ($h - $corner))) {
            & $beginBorderlessResize 17
            return
        }
        if ($allowBottom -and $allowLeft -and ($x -le $corner) -and ($y -ge ($h - $corner))) {
            & $beginBorderlessResize 16
            return
        }
        if ($onTop) {
            & $beginBorderlessResize 12
            return
        }
        if ($onRight) {
            & $beginBorderlessResize 11
            return
        }
        if ($onBottom) {
            & $beginBorderlessResize 15
            return
        }
        if ($onLeft) {
            & $beginBorderlessResize 10
            return
        }
    }.GetNewClosure()

    $resizeBottomGrip.Add_MouseDown({
        param($s, $e)
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
            & $writeResizeTrace 'BottomGripMouseDown' ("x={0};y={1};w={2};h={3}" -f $e.X, $e.Y, $resizeBottomGrip.Width, $resizeBottomGrip.Height)
            & $beginBorderlessResize 15
        }
    }.GetNewClosure())
    $resizeLeftGrip.Add_MouseDown({
        param($s, $e)
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
            & $writeResizeTrace 'LeftGripMouseDown' ("x={0};y={1};w={2};h={3}" -f $e.X, $e.Y, $resizeLeftGrip.Width, $resizeLeftGrip.Height)
            & $beginBorderlessResize 10
        }
    }.GetNewClosure())
    $resizeRightGrip.Add_MouseDown({
        param($s, $e)
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
            & $writeResizeTrace 'RightGripMouseDown' ("x={0};y={1};w={2};h={3}" -f $e.X, $e.Y, $resizeRightGrip.Width, $resizeRightGrip.Height)
            & $beginBorderlessResize 11
        }
    }.GetNewClosure())
    $resizeBottomLeftGrip.Add_MouseDown({
        param($s, $e)
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
            & $writeResizeTrace 'BottomLeftGripMouseDown' ("x={0};y={1};w={2};h={3}" -f $e.X, $e.Y, $resizeBottomLeftGrip.Width, $resizeBottomLeftGrip.Height)
            & $beginBorderlessResize 16
        }
    }.GetNewClosure())
    $resizeBottomRightGrip.Add_MouseDown({
        param($s, $e)
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
            & $writeResizeTrace 'BottomRightGripMouseDown' ("x={0};y={1};w={2};h={3}" -f $e.X, $e.Y, $resizeBottomRightGrip.Width, $resizeBottomRightGrip.Height)
            & $beginBorderlessResize 17
        }
    }.GetNewClosure())
    $resizeCornerLabel.Add_MouseDown({
        param($s, $e)
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
            & $writeResizeTrace 'CornerLabelMouseDown' ("x={0};y={1};w={2};h={3}" -f $e.X, $e.Y, $resizeBottomRightGrip.Width, $resizeBottomRightGrip.Height)
            & $beginBorderlessResize 17
        }
    }.GetNewClosure())
    $statusBar.Add_MouseDown({
        param($s, $e)
        & $writeResizeTrace 'StatusBarMouseDown' ("x={0};y={1};w={2};h={3}" -f $e.X, $e.Y, $statusBar.Width, $statusBar.Height)
        & $handleProxyResizeMouseDown $statusBar $e $false $true $true $false
    }.GetNewClosure())
    $windowShell.Add_MouseDown({
        param($s, $e)
        & $handleProxyResizeMouseDown $windowShell $e $false $true $true $false
    }.GetNewClosure())
    $topBar.Add_MouseDown({
        param($s, $e)
        & $handleProxyResizeMouseDown $topBar $e $true $true $false $false
    }.GetNewClosure())

    $lblAppTitle = _Label "Etherium Command Center $($script:AppVersion)" 18 9 340 20 $fontBold
    $lblAppTitle.ForeColor = $clrText
    $lblAppTitle.Cursor = 'Hand'
    $metricsPanel = _Panel 18 30 744 28 $clrPanelSoft
    $metricsPanel.BorderStyle = 'FixedSingle'
    $metricsPanel.Cursor = 'Hand'
    $lblCPU = _Label 'CPU: --%'          10  4 86 18 $fontBold
    $lblRAM = _Label 'RAM: --%'         120 4 88 18 $fontBold
$lblNET = _Label 'NET: -- Mbps'     232 4 120 18 $fontBold
    $lblDisk = _Label 'DISK: --'        376 4 170 18 $fontBold
    $lblPlayers = _Label 'PLAYERS: --'  570 4 100 18 $fontBold
    $lblBot = _Label 'Bot: Unknown'     694 4 78 18 $fontBold
    foreach ($metricLabel in @($lblCPU, $lblRAM, $lblNET, $lblDisk, $lblPlayers, $lblBot)) {
        $metricLabel.AutoEllipsis = $true
    }
    $lblCPU.ForeColor = $clrGreen
    $lblRAM.ForeColor = $clrAccentAlt
    $lblNET.ForeColor = $clrText
    $lblDisk.ForeColor = $clrYellow
    $lblPlayers.ForeColor = $clrAccentAlt
    $lblBot.ForeColor = $clrGreen
    $sepCPU = _Panel 104 5 1 16 $clrBorder
    $sepCPU.BorderStyle = 'None'
    $sepRAM = _Panel 216 5 1 16 $clrBorder
    $sepRAM.BorderStyle = 'None'
    $sepNET = _Panel 360 5 1 16 $clrBorder
    $sepNET.BorderStyle = 'None'
    $sepDisk = _Panel 554 5 1 16 $clrBorder
    $sepDisk.BorderStyle = 'None'
    $sepPlayers = _Panel 678 5 1 16 $clrBorder
    $sepPlayers.BorderStyle = 'None'
    $metricsPanel.Controls.Add($lblCPU)
    $metricsPanel.Controls.Add($sepCPU)
    $metricsPanel.Controls.Add($lblRAM)
    $metricsPanel.Controls.Add($sepRAM)
    $metricsPanel.Controls.Add($lblNET)
    $metricsPanel.Controls.Add($sepNET)
    $metricsPanel.Controls.Add($lblDisk)
    $metricsPanel.Controls.Add($sepDisk)
    $metricsPanel.Controls.Add($lblPlayers)
    $metricsPanel.Controls.Add($sepPlayers)
    $metricsPanel.Controls.Add($lblBot)

    $layoutTopMetrics = {
        if ($null -eq $metricsPanel) { return }

        $leftInset = 10
        $topInset = 4
        $labelHeight = 18
        $sepWidth = 1
        $sepHeight = 16
        $sepTop = 5
        $gapBeforeSep = 2
        $gapAfterSep = 4
        $rightInset = 10

        $metricDefs = @(
            @{ Label = $lblCPU;     Min = 64; Want = 80 },
            @{ Label = $lblRAM;     Min = 68; Want = 82 },
            @{ Label = $lblNET;     Min = 90; Want = 110 },
            @{ Label = $lblDisk;    Min = 120; Want = 160 },
            @{ Label = $lblPlayers; Min = 84; Want = 94 },
            @{ Label = $lblBot;     Min = 70; Want = 82 }
        )
        $separators = @($sepCPU, $sepRAM, $sepNET, $sepDisk, $sepPlayers)

        $availableWidth = [Math]::Max(320, $metricsPanel.ClientSize.Width - $leftInset - $rightInset)
        $separatorFootprint = $separators.Count * ($gapBeforeSep + $sepWidth + $gapAfterSep)

        foreach ($def in $metricDefs) {
            $want = [int]$def.Want
            try {
                $measured = [System.Windows.Forms.TextRenderer]::MeasureText([string]$def.Label.Text, $def.Label.Font).Width + 6
                if ($measured -gt $want) { $want = $measured }
            } catch { }
            $def.Width = [Math]::Max([int]$def.Min, $want)
        }

        $targetContentWidth = 0
        foreach ($def in $metricDefs) {
            $targetContentWidth += [int]$def.Width
        }
        $targetContentWidth += $separatorFootprint
        if ($targetContentWidth -gt $availableWidth) {
            $overflow = $targetContentWidth - $availableWidth
            foreach ($def in @($metricDefs[3], $metricDefs[2], $metricDefs[4], $metricDefs[5], $metricDefs[1], $metricDefs[0])) {
                if ($overflow -le 0) { break }
                $shrinkable = [Math]::Max(0, [int]$def.Width - [int]$def.Min)
                if ($shrinkable -le 0) { continue }
                $take = [Math]::Min($overflow, $shrinkable)
                $def.Width = [int]$def.Width - [int]$take
                $overflow -= $take
            }
        }

        $x = $leftInset
        for ($i = 0; $i -lt $metricDefs.Count; $i++) {
            $label = $metricDefs[$i].Label
            $width = [int]$metricDefs[$i].Width
            $label.Location = [System.Drawing.Point]::new($x, $topInset)
            $label.Size = [System.Drawing.Size]::new($width, $labelHeight)
            $x += $width

            if ($i -lt $separators.Count) {
                $x += $gapBeforeSep
                $separators[$i].Location = [System.Drawing.Point]::new($x, $sepTop)
                $separators[$i].Size = [System.Drawing.Size]::new($sepWidth, $sepHeight)
                $x += $sepWidth + $gapAfterSep
            }
        }

        $contentWidth = [Math]::Max(220, $x + $rightInset)
        $metricsPanel.Width = [Math]::Min([Math]::Max($contentWidth, 220), [Math]::Max(220, $topBar.ClientSize.Width - $metricsPanel.Left - 16))
        $lblAppTitle.Width = 340
    }.GetNewClosure()

    # Propagate drag events from labels up to the top bar so clicking on
    # any label still lets the user drag the window
    foreach ($dragLbl in @($lblAppTitle, $metricsPanel, $lblCPU, $lblRAM, $lblNET, $lblDisk, $lblPlayers, $lblBot)) {
        $dragLbl.Add_MouseDown({
            param($s, $e)
            if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
                $script:_DragActive = $true
                # Offset relative to the form, not just the label
                $screenPt = $this.PointToScreen($e.Location)
                $script:_DragOrigin = [System.Drawing.Point]::new(
                    $screenPt.X - $form.Location.X,
                    $screenPt.Y - $form.Location.Y
                )
            }
        })
        $dragLbl.Add_MouseMove({
            param($s, $e)
            if ($script:_DragActive) {
                $screenPt = $this.PointToScreen($e.Location)
                $form.Location = [System.Drawing.Point]::new(
                    $screenPt.X - $script:_DragOrigin.X,
                    $screenPt.Y - $script:_DragOrigin.Y
                )
            }
        })
        $dragLbl.Add_MouseUp({
            param($s, $e)
            $script:_DragActive = $false
        })
    }

    $topBar.Controls.Add($lblAppTitle)
    $topBar.Controls.Add($metricsPanel)
    & $layoutTopMetrics
    _SetMainControlToolTip -Control $lblCPU -Text 'Shows current CPU usage for the ECC host machine.'
    _SetMainControlToolTip -Control $lblRAM -Text 'Shows current RAM usage for the ECC host machine.'
    _SetMainControlToolTip -Control $lblNET -Text 'Shows current network traffic seen by ECC on this machine.'
    _SetMainControlToolTip -Control $lblDisk -Text 'Shows the main server drive ECC is tracking. Hover to see every tracked drive.'
    _SetMainControlToolTip -Control $lblPlayers -Text 'Shows the total trusted active player count across all running servers.'
    _SetMainControlToolTip -Control $lblBot -Text 'Shows whether the Discord bot listener is online.'

    $windowChromeButtonWidth = 36
    $windowChromeGap = 2
    $windowChromeMinGlyph = '-'
    $windowChromeMaxGlyph = '[]'
    $windowChromeCloseGlyph = 'X'
    $windowChromeMinBase = [System.Drawing.Color]::FromArgb(38, 50, 74)
    $windowChromeMaxBase = [System.Drawing.Color]::FromArgb(45, 58, 82)
    $windowChromeCloseBase = [System.Drawing.Color]::FromArgb(66, 42, 52)
    $windowChromeMinHover = [System.Drawing.Color]::FromArgb(86, 112, 168)
    $windowChromeMaxHover = [System.Drawing.Color]::FromArgb(88, 126, 102)
    $windowChromeCloseHover = [System.Drawing.Color]::FromArgb(186, 72, 90)
    $chromeY = 0
    $chromeH = $topBarHeight - 1
    $windowChromeButtonHeight = [Math]::Max(38, $chromeH - 14)
    $windowChromeButtonY = [Math]::Max(0, [int][Math]::Floor(($chromeH - $windowChromeButtonHeight) / 2))
    $chromeClusterWidth = ($windowChromeButtonWidth * 3) + ($windowChromeGap * 2)
    $actionPanelDesiredWidth = 650
    $actionPanelMinWidth = 470
    $actionPanelToChromeGap = 8
    $actionPanelRightOffset = $chromeClusterWidth + $actionPanelToChromeGap + $actionPanelDesiredWidth
    $actionPanel = _Panel ($defaultWidth - $actionPanelRightOffset) 14 $actionPanelDesiredWidth 40 $clrPanelSoft
    $actionPanel.BorderStyle = 'FixedSingle'
    $topBar.Controls.Add($actionPanel)

    $topStartAllColor = [System.Drawing.Color]::FromArgb(58, 128, 94)
    $topStopAllColor = [System.Drawing.Color]::FromArgb(156, 84, 66)
    $topReloadUiColor = [System.Drawing.Color]::FromArgb(66, 112, 214)
    $topReloadBotColor = [System.Drawing.Color]::FromArgb(182, 140, 56)
    $topReloadCommandsColor = [System.Drawing.Color]::FromArgb(154, 82, 168)
    $topFullRestartColor = [System.Drawing.Color]::FromArgb(184, 72, 92)
    $topSettingsColor = [System.Drawing.Color]::FromArgb(78, 92, 122)

    $btnStartAll = _Button 'Start All' 10 4 74 30 $topStartAllColor {
        $bulk = _GetBulkServerOperationTargets -Operation 'Start'
        if (-not $bulk -or $bulk.Count -le 0) {
            [System.Windows.Forms.MessageBox]::Show(
                'There are no offline/startable profiles to queue right now.',
                'Start All','OK','Information') | Out-Null
            return
        }

        $resp = [System.Windows.Forms.MessageBox]::Show(
            "Queue staggered start for $($bulk.Count) profile(s)? ECC will start them one by one with a short delay between each launch.",
            'Confirm Start All','YesNo','Question')
        if ($resp -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        _RunBulkServerOperations -Operation 'Start' -Prefixes @($bulk.Prefixes)
        $statusLabel.Text = "Queued staggered start for $($bulk.Count) profile(s)."
        _QueueStatusMessage ("Bulk staggered start requested for {0} profile(s): {1}" -f $bulk.Count, (_FormatBulkProfileSummary -Names @($bulk.Names)))
    }
    _SetMainControlToolTip -Control $btnStartAll -Text 'Start every offline server profile that ECC says is ready to launch, one at a time.'
    $actionPanel.Controls.Add($btnStartAll)

    $btnStopAll = _Button 'Stop All' 90 4 72 30 $topStopAllColor {
        $bulk = _GetBulkServerOperationTargets -Operation 'Stop'
        if (-not $bulk -or $bulk.Count -le 0) {
            [System.Windows.Forms.MessageBox]::Show(
                'There are no running profiles to stop right now.',
                'Stop All','OK','Information') | Out-Null
            return
        }

        $resp = [System.Windows.Forms.MessageBox]::Show(
            "Queue a safe stop for $($bulk.Count) running profile(s)?",
            'Confirm Stop All','YesNo','Warning')
        if ($resp -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        _RunBulkServerOperations -Operation 'Stop' -Prefixes @($bulk.Prefixes)
        $statusLabel.Text = "Queued stop for $($bulk.Count) running profile(s)."
        _QueueStatusMessage ("Bulk stop requested for {0} profile(s): {1}" -f $bulk.Count, (_FormatBulkProfileSummary -Names @($bulk.Names)))
    }
    _SetMainControlToolTip -Control $btnStopAll -Text 'Send a safe stop request to every server that is currently running.'
    $actionPanel.Controls.Add($btnStopAll)

    $btnReloadUI = _Button 'Reload UI' 168 4 84 30 $topReloadUiColor {
        $statusLabel.Text = "Reloading UI only. Running servers and timers will be preserved."
        _QueueStatusMessage 'Reload UI requested. Preserving running servers, auto-save timers, and restart timers.'
        _SendDiscordNotice (New-DiscordSystemMessage -Event 'reload_ui')
        try {
            if ($script:SharedState) {
                $script:SharedState['ReloadUI'] = $true
                try {
                    $reloadBounds = if ($form.WindowState -eq 'Normal') { $form.Bounds } else { $form.RestoreBounds }
                    $script:SharedState['ReloadWindowBounds'] = @{
                        Width = [int]$reloadBounds.Width
                        Height = [int]$reloadBounds.Height
                        X = [int]$reloadBounds.X
                        Y = [int]$reloadBounds.Y
                        State = if ($form.WindowState -eq 'Maximized') { 'Maximized' } else { 'Normal' }
                    }
                } catch { }
            }
            _PersistWindowSettings
            $script:_UIReloadRequested = $true
            try { $form.DialogResult = [System.Windows.Forms.DialogResult]::Retry } catch { }
            $form.Close()
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "The UI reload handoff failed.`r`n`r`n$($_.Exception.Message)",
                'Reload UI','OK','Error') | Out-Null
        }
    }
    _SetMainControlToolTip -Control $btnReloadUI
    $actionPanel.Controls.Add($btnReloadUI)

    $btnReloadBot = _Button 'Reload Bot' 258 4 84 30 $topReloadBotColor {
        $script:SharedState['RestartListener'] = $true
        $statusLabel.Text = "Reloading Discord bot. Running servers and timers are unchanged."
        _QueueStatusMessage 'Discord bot reload requested. Server timers remain intact.'
        _SendDiscordNotice (New-DiscordSystemMessage -Event 'reload_bot')
    }
    _SetMainControlToolTip -Control $btnReloadBot
    $actionPanel.Controls.Add($btnReloadBot)

    $btnReloadCommands = _Button 'Reload Cmds' 348 4 106 30 $topReloadCommandsColor {
        try {
            _ReloadCommandsAndProfiles
            _SendDiscordNotice (New-DiscordSystemMessage -Event 'reload_commands')
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to reload commands/profiles: $_",
                'Reload Commands','OK','Error') | Out-Null
        }
    }
    _SetMainControlToolTip -Control $btnReloadCommands
    $actionPanel.Controls.Add($btnReloadCommands)

    $btnFullRestart = _Button 'Full Restart' 460 4 92 30 $topFullRestartColor {
        $resp = [System.Windows.Forms.MessageBox]::Show(
            'This will stop all running servers, close ECC, and relaunch the full program. Continue?',
            'Confirm Full Restart','YesNo','Warning')
        if ($resp -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $script:SharedState['RestartProgram'] = $true
        $statusLabel.Text = 'Full program restart requested. ECC will stop running servers before relaunch.'
        _QueueStatusMessage 'Full program restart requested. ECC will stop running servers, close, and relaunch.'
        _SendDiscordNotice (New-DiscordSystemMessage -Event 'full_restart')
        $form.Close()
    }
    _SetMainControlToolTip -Control $btnFullRestart
    $actionPanel.Controls.Add($btnFullRestart)

    $btnSettings = _Button 'Settings' 558 4 82 30 $topSettingsColor {
        $settingsForm                  = New-Object System.Windows.Forms.Form
        $settingsForm.Text             = 'Settings'
        $settingsForm.Size             = [System.Drawing.Size]::new(760, 860)
        $settingsForm.MinimumSize      = [System.Drawing.Size]::new(700, 800)
        $settingsForm.StartPosition    = 'CenterParent'
        $settingsForm.BackColor        = $clrBg
        $settingsForm.FormBorderStyle  = 'Sizable'
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
    _SetMainControlToolTip -Control $btnSettings
    $actionPanel.Controls.Add($btnSettings)

    $layoutTopActions = {
        if ($null -eq $actionPanel) { return }

        $buttons = @(
            @{ Control = $btnStartAll;       Min = 68; Want = 78 },
            @{ Control = $btnStopAll;        Min = 66; Want = 76 },
            @{ Control = $btnReloadUI;       Min = 78; Want = 88 },
            @{ Control = $btnReloadBot;      Min = 80; Want = 90 },
            @{ Control = $btnReloadCommands; Min = 98; Want = 108 },
            @{ Control = $btnFullRestart;    Min = 92; Want = 102 },
            @{ Control = $btnSettings;       Min = 76; Want = 84 }
        )

        $leftInset = 8
        $rightInset = 8
        $topInset = 4
        $buttonHeight = 30
        $buttonGap = 6
        $availableWidth = [Math]::Max(360, $actionPanel.ClientSize.Width - $leftInset - $rightInset)

        foreach ($def in $buttons) {
            $want = [int]$def.Want
            try {
                $measured = [System.Windows.Forms.TextRenderer]::MeasureText([string]$def.Control.Text, $def.Control.Font).Width + 18
                if ($measured -gt $want) { $want = $measured }
            } catch { }
            $def.Width = [Math]::Max([int]$def.Min, $want)
        }

        $targetWidth = (($buttons.Count - 1) * $buttonGap)
        foreach ($def in $buttons) { $targetWidth += [int]$def.Width }

        if ($targetWidth -gt $availableWidth) {
            $overflow = $targetWidth - $availableWidth
            foreach ($def in @($buttons[4], $buttons[5], $buttons[3], $buttons[2], $buttons[6], $buttons[1], $buttons[0])) {
                if ($overflow -le 0) { break }
                $shrinkable = [Math]::Max(0, [int]$def.Width - [int]$def.Min)
                if ($shrinkable -le 0) { continue }
                $take = [Math]::Min($overflow, $shrinkable)
                $def.Width = [int]$def.Width - [int]$take
                $overflow -= $take
            }
        }

        $x = $leftInset
        foreach ($def in $buttons) {
            $control = $def.Control
            $control.Location = [System.Drawing.Point]::new($x, $topInset)
            $control.Size = [System.Drawing.Size]::new([int]$def.Width, $buttonHeight)
            $x += [int]$def.Width + $buttonGap
        }
    }.GetNewClosure()

    # -- Custom window chrome buttons (right edge of top bar) --

    # Minimize
    $btnWinMin = New-Object System.Windows.Forms.Button
    $btnWinMin.Size      = [System.Drawing.Size]::new($windowChromeButtonWidth, $windowChromeButtonHeight)
    $btnWinMin.Location  = [System.Drawing.Point]::new($defaultWidth - (($windowChromeButtonWidth * 3) + ($windowChromeGap * 2)), $windowChromeButtonY)
    $btnWinMin.Anchor    = 'Top,Right'
    $btnWinMin.Name      = 'btnWinMin'
    _ApplyWindowChromeButton -Button $btnWinMin -Text $windowChromeMinGlyph -BaseColor $windowChromeMinBase -HoverColor $windowChromeMinHover
    _BindClickHandler -Control $btnWinMin -Handler { $form.WindowState = 'Minimized' }
    _SetMainControlToolTip -Control $btnWinMin
    $topBar.Controls.Add($btnWinMin)

    # Maximize/Restore
    $btnWinMax = New-Object System.Windows.Forms.Button
    $btnWinMax.Size      = [System.Drawing.Size]::new($windowChromeButtonWidth, $windowChromeButtonHeight)
    $btnWinMax.Location  = [System.Drawing.Point]::new($defaultWidth - (($windowChromeButtonWidth * 2) + $windowChromeGap), $windowChromeButtonY)
    $btnWinMax.Anchor    = 'Top,Right'
    $btnWinMax.Name      = 'btnWinMax'
    _ApplyWindowChromeButton -Button $btnWinMax -Text $windowChromeMaxGlyph -BaseColor $windowChromeMaxBase -HoverColor $windowChromeMaxHover
    _BindClickHandler -Control $btnWinMax -Handler {
        if ($form.WindowState -eq 'Maximized') {
            $form.WindowState = 'Normal'
        } else {
            $form.WindowState = 'Maximized'
        }
        try {
            $form.BeginInvoke([System.Windows.Forms.MethodInvoker]{
                try { & $normalizeResizeChrome } catch { }
                if ($script:_ReflowLayoutHandler) { & $script:_ReflowLayoutHandler }
            }) | Out-Null
        } catch { }
    }
    _SetMainControlToolTip -Control $btnWinMax
    $topBar.Controls.Add($btnWinMax)

    # Close
    $btnWinClose = New-Object System.Windows.Forms.Button
    $btnWinClose.Size      = [System.Drawing.Size]::new($windowChromeButtonWidth, $windowChromeButtonHeight)
    $btnWinClose.Location  = [System.Drawing.Point]::new($defaultWidth - $windowChromeButtonWidth, $windowChromeButtonY)
    $btnWinClose.Anchor    = 'Top,Right'
    $btnWinClose.Name      = 'btnWinClose'
    _ApplyWindowChromeButton -Button $btnWinClose -Text $windowChromeCloseGlyph -BaseColor $windowChromeCloseBase -HoverColor $windowChromeCloseHover
    _BindClickHandler -Control $btnWinClose -Handler { $form.Close() }
    _SetMainControlToolTip -Control $btnWinClose
    $topBar.Controls.Add($btnWinClose)
    & $layoutTopActions

    # =====================================================================
    # THREE-COLUMN MIDDLE SECTION
    # =====================================================================
    $collapsedSize      = 28
    $headerHeight       = 34
    $bottomHeaderHeight = 30

    $script:_LeftCollapsed   = $false
    $script:_RightCollapsed  = $false
    $script:_BottomCollapsed = $false
    $script:_ReflowLayoutHandler = $null
    $script:_PaneToggleDebounce = @{}

    # Left container (Profiles)
    $leftContainer = _Panel $windowMargin ($topBarHeight + $windowMargin + 10) $leftWidth 600 $clrPanel
    $leftContainer.Anchor = 'Top,Left'
    $form.Controls.Add($leftContainer)

    $leftHeader = _Panel 0 0 $leftWidth $headerHeight $clrPanelAlt
    $leftHeader.Anchor = 'Top,Left,Right'
    $leftContainer.Controls.Add($leftHeader)
    $script:_LeftHeader = $leftHeader

    $leftHeaderLabel = _Label 'Game Profiles' 10 8 200 18 $fontBold
    $leftHeaderLabel.ForeColor = $clrAccentAlt
    $leftHeader.Controls.Add($leftHeaderLabel)
    $leftHeaderToggle = New-Object System.Windows.Forms.Button
    $leftHeaderToggle.Size = [System.Drawing.Size]::new(26, 22)
    $leftHeaderToggle.Location = [System.Drawing.Point]::new([Math]::Max(6, $leftWidth - 34), 6)
    $leftHeaderToggle.Anchor = 'Top,Right'
    $leftHeaderToggle.FlatStyle = 'Flat'
    $leftHeaderToggle.FlatAppearance.BorderSize = 1
    $leftHeaderToggle.FlatAppearance.MouseOverBackColor = $clrPanelSoft
    $leftHeaderToggle.FlatAppearance.MouseDownBackColor = $clrPanel
    $leftHeaderToggle.BackColor = $clrPanelSoft
    $leftHeaderToggle.ForeColor = $clrAccentAlt
    $leftHeaderToggle.Font = _ResolveUiFont -Font $fontBold
    $leftHeaderToggle.Text = '-'
    $leftHeaderToggle.Cursor = 'Hand'
    $leftHeaderToggle.Tag = 'Left'
    $leftHeaderToggle.Name = 'LeftHeaderToggle'
    $leftHeader.Controls.Add($leftHeaderToggle)
    $leftHeaderAccent = _Panel 0 0 4 $headerHeight $clrAccent
    $leftHeaderAccent.BorderStyle = 'None'
    $leftHeader.Controls.Add($leftHeaderAccent)

    $leftBody = _Panel 0 $headerHeight $leftWidth (600 - $headerHeight) $clrPanel
    $leftBody.Anchor = 'Top,Left,Right,Bottom'
    $leftBody.BackColor = $clrPanel
    $leftContainer.Controls.Add($leftBody)
    $script:_ProfilesPanel = $leftBody

    # Right container (Profile Editor)
    $rightContainer = _Panel ($defaultWidth - $rightWidth - $windowMargin) ($topBarHeight + $windowMargin + 10) $rightWidth 600 $clrPanel
    $rightContainer.Anchor = 'Top,Left'
    $form.Controls.Add($rightContainer)

    $rightHeader = _Panel 0 0 $rightWidth $headerHeight $clrPanelAlt
    $rightHeader.Anchor = 'Top,Left,Right'
    $rightContainer.Controls.Add($rightHeader)
    $script:_RightHeader = $rightHeader

    $rightHeaderLabel = _Label 'Profile Editor' 10 8 200 18 $fontBold
    $rightHeaderLabel.ForeColor = $clrAccentAlt
    $rightHeader.Controls.Add($rightHeaderLabel)
    $rightHeaderToggle = New-Object System.Windows.Forms.Button
    $rightHeaderToggle.Size = [System.Drawing.Size]::new(26, 22)
    $rightHeaderToggle.Location = [System.Drawing.Point]::new([Math]::Max(6, $rightWidth - 34), 6)
    $rightHeaderToggle.Anchor = 'Top,Right'
    $rightHeaderToggle.FlatStyle = 'Flat'
    $rightHeaderToggle.FlatAppearance.BorderSize = 1
    $rightHeaderToggle.FlatAppearance.MouseOverBackColor = $clrPanelSoft
    $rightHeaderToggle.FlatAppearance.MouseDownBackColor = $clrPanel
    $rightHeaderToggle.BackColor = $clrPanelSoft
    $rightHeaderToggle.ForeColor = $clrAccentAlt
    $rightHeaderToggle.Font = _ResolveUiFont -Font $fontBold
    $rightHeaderToggle.Text = '-'
    $rightHeaderToggle.Cursor = 'Hand'
    $rightHeaderToggle.Tag = 'Right'
    $rightHeaderToggle.Name = 'RightHeaderToggle'
    $rightHeader.Controls.Add($rightHeaderToggle)
    $rightHeaderAccent = _Panel 0 0 4 $headerHeight $clrAccent
    $rightHeaderAccent.BorderStyle = 'None'
    $rightHeader.Controls.Add($rightHeaderAccent)

    $rightBody = _Panel 0 $headerHeight $rightWidth (600 - $headerHeight) $clrPanel
    $rightBody.Anchor = 'Top,Left,Right,Bottom'
    $rightBody.BackColor = $clrPanel
    $rightContainer.Controls.Add($rightBody)
    $script:_ProfileEditorPanel = $rightBody
    $rightContainer.Add_MouseDown({ param($s, $e) & $handleProxyResizeMouseDown $rightContainer $e $false $true $false $false }.GetNewClosure())
    $rightHeader.Add_MouseDown({ param($s, $e) & $handleProxyResizeMouseDown $rightHeader $e $false $true $false $false }.GetNewClosure())
    $rightBody.Add_MouseDown({ param($s, $e) & $handleProxyResizeMouseDown $rightBody $e $false $true $false $false }.GetNewClosure())

    # Center dashboard
    $centerCol = _Panel ($windowMargin + $leftWidth + $sideGap) ($topBarHeight + $windowMargin + 10) ($defaultWidth - $leftWidth - $rightWidth - ($sideGap * 2) - ($windowMargin * 2)) 600 $clrPanel
    $centerCol.Anchor = 'Top,Left'
    $centerCol.BorderStyle = 'FixedSingle'
    $form.Controls.Add($centerCol)
    $script:_ServerDashboardPanel = $centerCol

    # =====================================================================
    # BOTTOM LOG STRIP
    # =====================================================================
    # =====================================================================
    # BOTTOM LOG STRIP  -  TabControl with Discord / Program / per-game tabs
    # =====================================================================
    $bottomContainer        = _Panel $windowMargin ($defaultHeight - $bottomLogsHeight - 22 - $windowMargin) ($defaultWidth - ($windowMargin * 2)) $bottomLogsHeight $clrPanel
    $bottomContainer.Anchor = 'Top,Left'
    $form.Controls.Add($bottomContainer)

    # IMPORTANT: Dock=Fill must be added BEFORE Dock=Top controls.
    # WinForms processes docked controls in reverse add order, so Fill
    # added first means it correctly receives all space not claimed by Top.
    $bottomPanel        = New-Object System.Windows.Forms.Panel
    $bottomPanel.Dock   = 'Fill'
    $bottomPanel.BackColor = $clrBg
    $bottomContainer.Controls.Add($bottomPanel)

    $bottomHeader = New-Object System.Windows.Forms.Panel
    $bottomHeader.Dock      = 'Top'
    $bottomHeader.Height    = $bottomHeaderHeight
    $bottomHeader.BackColor = $clrPanelAlt
    $bottomContainer.Controls.Add($bottomHeader)
    $script:_BottomHeader = $bottomHeader

    $bottomHeaderLabel = _Label 'Logs' 10 7 200 16 $fontBold
    $bottomHeaderLabel.ForeColor = $clrAccentAlt
    $bottomHeader.Controls.Add($bottomHeaderLabel)
    $bottomHeaderToggle = New-Object System.Windows.Forms.Button
    $bottomHeaderToggle.Size = [System.Drawing.Size]::new(26, 22)
    $bottomHeaderToggle.Location = [System.Drawing.Point]::new([Math]::Max(6, $bottomContainer.Width - 34), 4)
    $bottomHeaderToggle.Anchor = 'Top,Right'
    $bottomHeaderToggle.FlatStyle = 'Flat'
    $bottomHeaderToggle.FlatAppearance.BorderSize = 1
    $bottomHeaderToggle.FlatAppearance.MouseOverBackColor = $clrPanelSoft
    $bottomHeaderToggle.FlatAppearance.MouseDownBackColor = $clrPanel
    $bottomHeaderToggle.BackColor = $clrPanelSoft
    $bottomHeaderToggle.ForeColor = $clrAccentAlt
    $bottomHeaderToggle.Font = _ResolveUiFont -Font $fontBold
    $bottomHeaderToggle.Text = '-'
    $bottomHeaderToggle.Cursor = 'Hand'
    $bottomHeaderToggle.Tag = 'Bottom'
    $bottomHeaderToggle.Name = 'BottomHeaderToggle'
    $bottomHeader.Controls.Add($bottomHeaderToggle)
    $bottomResizeHandle = New-Object System.Windows.Forms.Button
    $bottomResizeHandle.Size = [System.Drawing.Size]::new(26, 22)
    $bottomResizeHandle.Location = [System.Drawing.Point]::new([Math]::Max(6, $bottomContainer.Width - 66), 4)
    $bottomResizeHandle.Anchor = 'Top,Right'
    $bottomResizeHandle.FlatStyle = 'Flat'
    $bottomResizeHandle.FlatAppearance.BorderSize = 1
    $bottomResizeHandle.FlatAppearance.MouseOverBackColor = $clrPanelSoft
    $bottomResizeHandle.FlatAppearance.MouseDownBackColor = $clrPanel
    $bottomResizeHandle.BackColor = $clrPanelSoft
    $bottomResizeHandle.ForeColor = $clrAccentAlt
    $bottomResizeHandle.Font = _ResolveUiFont -Font $fontBold
    $bottomResizeHandle.Text = '///'
    $bottomResizeHandle.Cursor = [System.Windows.Forms.Cursors]::SizeNWSE
    $bottomResizeHandle.Tag = 'BottomRight'
    $bottomResizeHandle.Name = 'BottomResizeHandle'
    $bottomHeader.Controls.Add($bottomResizeHandle)
    $bottomResizeHandle.Add_MouseDown({
        param($s, $e)
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) { & $beginBorderlessResize 17 }
    }.GetNewClosure())
    $bottomHeaderAccent = _Panel 0 0 4 $bottomHeaderHeight $clrAccent
    $bottomHeaderAccent.BorderStyle = 'None'
    $bottomHeader.Controls.Add($bottomHeaderAccent)
    $bottomContainer.Add_MouseDown({ param($s, $e) & $handleProxyResizeMouseDown $bottomContainer $e $false $true $true $false }.GetNewClosure())
    $bottomHeader.Add_MouseDown({ param($s, $e) & $handleProxyResizeMouseDown $bottomHeader $e $false $true $true $false }.GetNewClosure())
    $bottomPanel.Add_MouseDown({ param($s, $e) & $handleProxyResizeMouseDown $bottomPanel $e $false $true $true $false }.GetNewClosure())

    # Master TabControl - Dock Fill so it always occupies all of bottomPanel
    $logTabs            = New-Object System.Windows.Forms.TabControl
    $logTabs.Dock       = 'Fill'
    $logTabs.BackColor  = $clrPanelSoft
    $logTabs.Font       = $tabFont
    $logTabs.DrawMode   = [System.Windows.Forms.TabDrawMode]::OwnerDrawFixed
    $logTabs.SizeMode   = [System.Windows.Forms.TabSizeMode]::Fixed
    $logTabs.ItemSize   = [System.Drawing.Size]::new(120, 24)
    $logTabs.Add_DrawItem({
        param($s,$e)

        $tab = $logTabs.TabPages[$e.Index]
        try {
            $stripBottom = [Math]::Max($e.Bounds.Bottom, $logTabs.DisplayRectangle.Top)
            if ($stripBottom -gt 0) {
                $fillBrush = [System.Drawing.SolidBrush]::new($clrPanelAlt)
                $borderPen = New-Object System.Drawing.Pen($clrBorder)
                $originalClip = $e.Graphics.Clip
                try {
                    for ($tabIndex = 0; $tabIndex -lt $logTabs.TabPages.Count; $tabIndex++) {
                        try {
                            $tabRect = $logTabs.GetTabRect($tabIndex)
                            if ($tabRect.Width -gt 0 -and $tabRect.Height -gt 0) {
                                $e.Graphics.ExcludeClip($tabRect)
                            }
                        } catch { }
                    }
                    $e.Graphics.FillRectangle($fillBrush, [System.Drawing.Rectangle]::new(0, 0, $logTabs.ClientSize.Width, $stripBottom))
                } finally {
                    $e.Graphics.Clip = $originalClip
                }
                $e.Graphics.DrawLine($borderPen, 0, $stripBottom - 1, $logTabs.ClientSize.Width, $stripBottom - 1)
                $fillBrush.Dispose()
                $borderPen.Dispose()
            }
        } catch { }

        $brush = if ($e.Index -eq $logTabs.SelectedIndex) {
            [System.Drawing.SolidBrush]::new($clrAccent)
        } else {
            [System.Drawing.SolidBrush]::new($clrPanelSoft)
        }

        $e.Graphics.FillRectangle($brush, $e.Bounds)

        $txtBrush = [System.Drawing.SolidBrush]::new($clrText)

        $sf = New-Object System.Drawing.StringFormat
        $sf.Alignment     = 'Center'
        $sf.LineAlignment = 'Center'
        $sf.Trimming      = [System.Drawing.StringTrimming]::EllipsisCharacter
        $sf.FormatFlags   = [System.Drawing.StringFormatFlags]::NoWrap

        $padding = 6
        $rectF = [System.Drawing.RectangleF]::new(
            $e.Bounds.X + $padding,
            $e.Bounds.Y,
            $e.Bounds.Width - ($padding * 2),
            $e.Bounds.Height
        )

        $e.Graphics.DrawString($tab.Text, $tabFont, $txtBrush, $rectF, $sf)
        $tabPen = New-Object System.Drawing.Pen($clrBorder)
        $e.Graphics.DrawLine($tabPen, $e.Bounds.Left, $e.Bounds.Bottom - 1, $e.Bounds.Right, $e.Bounds.Bottom - 1)

        $brush.Dispose()
        $txtBrush.Dispose()
        $tabPen.Dispose()
        $sf.Dispose()
    })

    $bottomPanel.Controls.Add($logTabs)
    $script:_LogTabControl = $logTabs

    # ── TAB WIDTH RECALCULATOR  -  measures every tab label and sets ItemSize width
    # to fit the longest one.  Min 70px, max 160px.  Called after every add/remove.
    function _RecalcTabWidths {
        $tc = $script:_LogTabControl
        if ($null -eq $tc -or $tc.TabPages.Count -eq 0) { return }
        try {
            $g    = $tc.CreateGraphics()
            $maxW = 70
            foreach ($tp in $tc.TabPages) {
                $sz = $g.MeasureString($tp.Text, $tabFont)
                $w  = [int]$sz.Width + 16   # 8px padding each side
                if ($w -gt $maxW) { $maxW = $w }
            }
            $g.Dispose()
            if ($maxW -gt 160) { $maxW = 160 }
            $tc.ItemSize = [System.Drawing.Size]::new($maxW, 24)
        } catch { }
    }
    # Per-RTB line counters: avoids calling $rt.Lines.Count (O(n) rescan) on every append.
    # Key = RuntimeHelpers.GetHashCode(rtb), Value = int line count
    $script:_RtbLineCounts = @{}
    $maxRtbLines = 50    # hard cap; keep UI log windows intentionally small to avoid long-session lag

    # Helper: append a colored line to a RichTextBox, cap at $maxRtbLines lines
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

        # Use RuntimeHelpers.GetHashCode as a stable cheap identity key per RTB instance.
        # This avoids touching any RTB property (which can trigger layout/repaint).
        $rtKey = [System.Runtime.CompilerServices.RuntimeHelpers]::GetHashCode($rt)
        if (-not $script:_RtbLineCounts.ContainsKey($rtKey)) { $script:_RtbLineCounts[$rtKey] = 0 }
        $script:_RtbLineCounts[$rtKey]++
        $currentCount = $script:_RtbLineCounts[$rtKey]
        $debugAppendLog = _IsGuiDebugEnabled
        $appendMs = 0.0
        $trimMs = 0.0
        $scrollMs = 0.0
        $finalTextLength = 0
        $visibleLineCount = $currentCount

        # Keep the visible RTB history very small. We still avoid touching
        # $rt.Lines on every append, but the trim window is now intentionally
        # tight so long-running sessions do not accumulate expensive UI text.
        $trimThreshold = $maxRtbLines + 5

        try {
            $rt.SelectionStart  = $rt.TextLength
            $rt.SelectionLength = 0
            $rt.SelectionColor  = $colour
            $appendPerf = $null
            if ($debugAppendLog) { $appendPerf = [System.Diagnostics.Stopwatch]::StartNew() }
            $rt.AppendText("$line`n")
            if ($appendPerf) {
                $appendPerf.Stop()
                $appendMs = $appendPerf.Elapsed.TotalMilliseconds
            }

            if ($currentCount -ge $trimThreshold) {
                # Now it is worth paying the Lines.Count cost
                $trimPerf = $null
                if ($debugAppendLog) { $trimPerf = [System.Diagnostics.Stopwatch]::StartNew() }
                $actualLines = $rt.Lines.Count
                if ($actualLines -gt $maxRtbLines) {
                    $removeCount = $actualLines - $maxRtbLines
                    $cutEnd = $rt.GetFirstCharIndexFromLine($removeCount)
                    if ($cutEnd -gt 0) {
                        $restoreReadOnly = $false
                        try { $restoreReadOnly = [bool]$rt.ReadOnly } catch { $restoreReadOnly = $false }
                        try {
                            if ($restoreReadOnly) { $rt.ReadOnly = $false }
                            $rt.SelectionStart  = 0
                            $rt.SelectionLength = $cutEnd
                            $rt.SelectedText    = ''
                        } finally {
                            if ($restoreReadOnly) {
                                try { $rt.ReadOnly = $true } catch { }
                            }
                        }
                    }
                }
                # Sync counter to reality after the trim
                $script:_RtbLineCounts[$rtKey] = $rt.Lines.Count
                $visibleLineCount = $script:_RtbLineCounts[$rtKey]
                if ($trimPerf) {
                    $trimPerf.Stop()
                    $trimMs = $trimPerf.Elapsed.TotalMilliseconds
                }
            } else {
                $visibleLineCount = $currentCount
            }

            # Move caret to end and scroll (must come after any trim)
            $scrollPerf = $null
            if ($debugAppendLog) { $scrollPerf = [System.Diagnostics.Stopwatch]::StartNew() }
            $rt.SelectionStart  = $rt.TextLength
            $rt.SelectionLength = 0
            $rt.ScrollToCaret()
            if ($scrollPerf) {
                $scrollPerf.Stop()
                $scrollMs = $scrollPerf.Elapsed.TotalMilliseconds
            }
            $finalTextLength = $rt.TextLength

        } catch { }

        if ($debugAppendLog) {
            try {
                _GuiDirectLog -Level DEBUG -Message ('APPENDLOG rtb={0} appendMs={1:N2} trimMs={2:N2} scrollMs={3:N2} textLength={4} visibleLines={5} lineChars={6}' -f `
                    $rtKey, $appendMs, $trimMs, $scrollMs, $finalTextLength, $visibleLineCount, $(if ($line) { $line.Length } else { 0 }))
            } catch { }
        }
    }

    # Helper: pick colour from log line content
    function _LogColour {
        param([string]$line)
        if ($line -match 'Shutdown requested|Shutdown phase [1-4]/4|Shutdown complete|Phase [1-4] detail:') { return $clrAccentAlt }
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
    $discordInner.BackColor = $clrPanel
    $tabDiscord.Controls.Add($discordInner)

    $rtDiscord            = [System.Windows.Forms.RichTextBox]::new()
    $rtDiscord.Dock       = 'Fill'
    $rtDiscord.BackColor  = [System.Drawing.Color]::FromArgb(12,14,24)
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
    $discordFooter.Height = 38
    $discordFooter.BackColor = $clrPanelAlt
    $discordInner.Controls.Add($discordFooter)
    $script:_DiscordFooter = $discordFooter

    $tbSend  = _TextBox 0 0 100 24 ''
    $tbSend.Dock = 'Fill'

    $btnSend = _Button 'Send' 0 0 90 28 $clrAccentAlt {
        $msg = $tbSend.Text.Trim()
        if (-not $msg) { return }
        $msg = "[BOT] $msg"
        _SendDiscordNotice -Message $msg
        $tbSend.Text = ''
    }
    _SetMainControlToolTip -Control $btnSend -Text 'Send the current message from the text box to the Discord output channel.'
    $btnClearDisc = _Button 'Clear' ([Math]::Max(206, $discordFooter.Width - 117)) 5 72 28 $clrMuted {
        if ($script:_DiscordLogBox) {
            $script:_DiscordLogBox.Clear()
            _ResetLogBoxState $script:_DiscordLogBox
        }
    }
    _SetMainControlToolTip -Control $btnClearDisc -Text 'Clear the visible Discord log in the ECC window.'
    $btnSend.Margin = [System.Windows.Forms.Padding]::new(0, 0, 8, 0)
    $btnClearDisc.Margin = [System.Windows.Forms.Padding]::new(0)

    $discordInputHost = New-Object System.Windows.Forms.Panel
    $discordInputHost.Dock = 'Fill'
    $discordInputHost.Padding = [System.Windows.Forms.Padding]::new(12, 7, 8, 7)
    $discordInputHost.BackColor = [System.Drawing.Color]::Transparent
    $discordInputHost.Controls.Add($tbSend)

    $discordActions = New-Object System.Windows.Forms.FlowLayoutPanel
    $discordActions.Dock = 'Right'
    $discordActions.Width = 182
    $discordActions.WrapContents = $false
    $discordActions.FlowDirection = 'LeftToRight'
    $discordActions.Padding = [System.Windows.Forms.Padding]::new(0, 5, 8, 5)
    $discordActions.Margin = [System.Windows.Forms.Padding]::new(0)
    $discordActions.BackColor = [System.Drawing.Color]::Transparent
    [void]$discordActions.Controls.Add($btnSend)
    [void]$discordActions.Controls.Add($btnClearDisc)

    $discordFooter.Controls.Add($discordInputHost)
    $discordFooter.Controls.Add($discordActions)

    # ── PROGRAM LOG TAB ───────────────────────────────────────────────────────
    $tabProgram           = New-Object System.Windows.Forms.TabPage
    $tabProgram.Text      = 'Program Log'
    $tabProgram.BackColor = $clrBg
    $logTabs.TabPages.Add($tabProgram)

    $programInner         = New-Object System.Windows.Forms.Panel
    $programInner.Dock    = 'Fill'
    $programInner.BackColor = $clrPanel
    $tabProgram.Controls.Add($programInner)

    # RTB added FIRST (Fill) then footer (Bottom) - same pattern as Discord tab
    $rtProgram            = [System.Windows.Forms.RichTextBox]::new()
    $rtProgram.Dock       = 'Fill'
    $rtProgram.BackColor  = [System.Drawing.Color]::FromArgb(12,14,24)
    $rtProgram.ForeColor  = $clrText
    $rtProgram.Font       = $fontMono
    $rtProgram.ReadOnly   = $true
    $rtProgram.ScrollBars = 'Vertical'
    $rtProgram.WordWrap   = $false
    $programInner.Controls.Add($rtProgram)
    $script:_ProgramLogBox = $rtProgram

    # Footer with Clear button - added AFTER RTB so Dock=Bottom works correctly
    $programFooter        = New-Object System.Windows.Forms.Panel
    $programFooter.Dock   = 'Bottom'
    $programFooter.Height = 36
    $programFooter.BackColor = $clrPanelAlt
    $programInner.Controls.Add($programFooter)

    $btnClearProg = _Button 'Clear' 0 4 72 28 $clrMuted {
        if ($script:_ProgramLogBox) {
            $script:_ProgramLogBox.Clear()
            _ResetLogBoxState $script:_ProgramLogBox
        }
        if (-not [string]::IsNullOrEmpty($script:LogFilePath) -and (Test-Path $script:LogFilePath)) {
            $script:LogFilePos = (Get-Item $script:LogFilePath).Length
        }
    }
    _SetMainControlToolTip -Control $btnClearProg -Text 'Clear the visible Program Log in the ECC window without deleting log files.'
    $btnCopyProg = _Button 'Copy' 0 4 72 28 $clrPanel {
        $ok = _CopyLogText -Box $script:_ProgramLogBox
        if (-not $ok) {
            [System.Windows.Forms.MessageBox]::Show(
                'There is no Program Log text to copy yet.',
                'Copy Program Log','OK','Information') | Out-Null
        }
    }
    _SetMainControlToolTip -Control $btnCopyProg -Text 'Copy the visible Program Log text to the clipboard.'
    $btnCopyProg.Margin = [System.Windows.Forms.Padding]::new(0, 0, 8, 0)
    $btnClearProg.Margin = [System.Windows.Forms.Padding]::new(0)

    $programFooterActions = New-Object System.Windows.Forms.FlowLayoutPanel
    $programFooterActions.Dock = 'Right'
    $programFooterActions.Width = 160
    $programFooterActions.WrapContents = $false
    $programFooterActions.FlowDirection = 'LeftToRight'
    $programFooterActions.Padding = [System.Windows.Forms.Padding]::new(0, 4, 8, 4)
    $programFooterActions.Margin = [System.Windows.Forms.Padding]::new(0)
    $programFooterActions.BackColor = [System.Drawing.Color]::Transparent
    [void]$programFooterActions.Controls.Add($btnCopyProg)
    [void]$programFooterActions.Controls.Add($btnClearProg)
    $programFooter.Controls.Add($programFooterActions)

    # Prefer Program Log as the first visible static tab on startup.
    $logTabs.TabPages.Clear()
    $logTabs.TabPages.Add($tabProgram)
    $logTabs.TabPages.Add($tabDiscord)
    $logTabs.SelectedTab = $tabProgram

    # Set initial tab widths now that both static tabs exist
    _RecalcTabWidths

    # ── GAME LOG TABS  -  created/removed as servers start and stop ───────────
    # Key = game prefix (e.g. 'PZ'), Value = hashtable with TabPage + RTB + tail state
    $script:_GameLogTabs   = @{}   # prefix -> @{ Tab; RTB; Files=@{path->pos}; LogRoot; Strategy; LastFolder }
    $script:_ServerStartNotified = $script:SharedState.ServerStartNotified  # prefix -> $true once "SERVER STARTED" seen
    $script:_PlayersCapture = @{}       # prefix -> @{ Active; Expected; Names; Started }
    $script:_HytaleWhoCapture = @{}     # prefix -> @{ Active; Count; Names; Started }
    $script:_SatisfactoryConnectionCapture = $script:SharedState.SatisfactoryConnectionCapture # prefix -> @{ RecentRemoteAddr=''; Connections=@{} }
    $script:_ValheimPlayerCapture = $script:SharedState.ValheimPlayerCapture

    function _LoadRecentGameLogHistory {
        param(
            [string]$Prefix,
            [hashtable]$Profile,
            [switch]$Force
        )

        if ([string]::IsNullOrWhiteSpace($Prefix)) { return }
        if (-not $script:_GameLogTabs.ContainsKey($Prefix)) { return }
        if ($null -eq $Profile) { return }
        $commandSharedState = $script:SharedState

        $entry = $script:_GameLogTabs[$Prefix]
        if ($null -eq $entry) { return }
        if (-not $Force -and $entry.HistoryLoaded -eq $true) { return }

        $files = _ResolveGameLogFiles -Profile $Profile
        if ($files.Count -eq 0) { return }

        $entry.RTB.Clear()
        $entry.NoteShown = $false
        $entry.Files = @{}

        $loadedAny = $false
        foreach ($path in $files) {
            if ([string]::IsNullOrWhiteSpace($path) -or -not (Test-Path $path)) { continue }
            $leaf = Split-Path $path -Leaf
            $lines = @()
            try {
                $lines = @(Get-Content -Path $path -Tail 10 -ErrorAction Stop)
            } catch {
                continue
            }

            if ($lines.Count -eq 0) { continue }
            _AppendLog $entry.RTB "[HISTORY] Last $($lines.Count) line(s) from $leaf" $clrMuted
            foreach ($line in $lines) {
                if ($null -ne $line -and "$line".Trim().Length -gt 0) {
                    _AppendLog $entry.RTB "$line" (_LogColour "$line")
                }
            }
            $loadedAny = $true
        }

        if ($loadedAny) {
            $entry.HistoryLoaded = $true
        }
    }

    function _FlushPendingGameLogLines {
        param([string]$Prefix)

        if ([string]::IsNullOrWhiteSpace($Prefix)) { return }
        if (-not $script:_GameLogTabs.ContainsKey($Prefix)) { return }

        $entry = $script:_GameLogTabs[$Prefix]
        if ($null -eq $entry -or $null -eq $entry.PendingLines) { return }

        foreach ($queued in @($entry.PendingLines)) {
            if ($null -ne $queued -and "$queued".Length -gt 0) {
                _AppendLog $entry.RTB "$queued" (_LogColour "$queued")
            }
        }
        $entry.PendingLines.Clear()
    }

    function _GetProjectZomboidLatestSessionFolder {
        param([string]$RootPath)

        if ([string]::IsNullOrWhiteSpace($RootPath)) { return $null }
        if (-not (Test-Path -LiteralPath $RootPath)) { return $null }

        $dirs = @(Get-ChildItem -Path $RootPath -Directory -ErrorAction SilentlyContinue)
        if ($dirs.Count -eq 0) { return $null }

        $dated = @($dirs | Where-Object { $_.Name -match '^logs_\d{4}-\d{2}-\d{2}$' -or $_.Name -match '^\d{4}-\d{2}-\d{2}_\d{2}-\d{2}' } | Sort-Object Name -Descending)
        if ($dated.Count -gt 0) { return $dated[0] }

        return ($dirs | Sort-Object LastWriteTime -Descending | Select-Object -First 1)
    }

    function _GetProjectZomboidLogDirectories {
        param([string]$RootPath)

        $dirs = New-Object 'System.Collections.Generic.List[string]'
        if ([string]::IsNullOrWhiteSpace($RootPath)) { return $dirs }
        if (-not (Test-Path -LiteralPath $RootPath)) { return $dirs }

        $dirs.Add($RootPath) | Out-Null
        $latestFolder = _GetProjectZomboidLatestSessionFolder -RootPath $RootPath
        if ($latestFolder -and -not $dirs.Contains($latestFolder.FullName)) {
            $dirs.Add($latestFolder.FullName) | Out-Null
        }

        return $dirs
    }

    function _GetProjectZomboidSessionKeyFromPath {
        param([string]$Path)

        if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
        $leaf = ''
        try { $leaf = [System.IO.Path]::GetFileName($Path) } catch { $leaf = $Path }
        if ($leaf -match '^(\d{4}-\d{2}-\d{2}_\d{2}-\d{2})_') {
            return $Matches[1]
        }
        return ''
    }

    function _ResolveProjectZomboidSessionLogFiles {
        param([hashtable]$Profile)

        $files = [System.Collections.Generic.List[string]]::new()
        $root = _ExpandPathVars ([string]$Profile.ServerLogRoot)
        $rootUsed = $root
        $searchDirs = @(_GetProjectZomboidLogDirectories -RootPath $root)
        if ($searchDirs.Count -eq 0) {
            return [pscustomobject]@{
                Files      = $files
                RootUsed   = $rootUsed
                SessionDir = $null
                SessionKey = ''
                SearchDirs = @()
            }
        }
        $patterns = @(
            $(if ($Profile.ServerLogFileDebug) { [string]$Profile.ServerLogFileDebug } else { '' }),
            $(if ($Profile.ServerLogFileUser)  { [string]$Profile.ServerLogFileUser } else { '' }),
            $(if ($Profile.ServerLogFileChat)  { [string]$Profile.ServerLogFileChat } else { '' }),
            'DebugLog-server.txt',
            '*_DebugLog-server.txt',
            '*_user.txt',
            '*_chat.txt'
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique

        $candidateFiles = New-Object 'System.Collections.Generic.List[object]'
        foreach ($dir in $searchDirs) {
            foreach ($pattern in $patterns) {
                foreach ($match in @(Get-ChildItem -Path $dir -Filter $pattern -File -ErrorAction SilentlyContinue)) {
                    $candidateFiles.Add([pscustomobject]@{
                        File       = $match
                        SessionKey = (_GetProjectZomboidSessionKeyFromPath -Path $match.Name)
                    }) | Out-Null
                }
            }
        }

        $bestSession = ''
        $bestSessionDir = $null
        $sessionCandidates = @($candidateFiles | Where-Object { -not [string]::IsNullOrWhiteSpace($_.SessionKey) } | Sort-Object { $_.File.LastWriteTime } -Descending)
        if ($sessionCandidates.Count -gt 0) {
            $bestSession = [string]$sessionCandidates[0].SessionKey
            $bestSessionDir = $sessionCandidates[0].File.Directory
            $rootUsed = $bestSessionDir.FullName
            foreach ($candidate in $sessionCandidates) {
                if ($candidate.SessionKey -ne $bestSession) { continue }
                if (-not $files.Contains($candidate.File.FullName)) {
                    $files.Add($candidate.File.FullName) | Out-Null
                }
            }
        } else {
            foreach ($pattern in $patterns) {
                $match = $candidateFiles |
                    Where-Object { $_.File.Name -like $pattern } |
                    Sort-Object { $_.File.LastWriteTime } -Descending |
                    Select-Object -First 1
                if ($match -and -not $files.Contains($match.File.FullName)) {
                    $files.Add($match.File.FullName) | Out-Null
                    $rootUsed = $match.File.DirectoryName
                }
            }
        }

        return [pscustomobject]@{
            Files      = $files
            RootUsed   = $rootUsed
            SessionDir = $bestSessionDir
            SessionKey = $bestSession
            SearchDirs = @($searchDirs)
        }
    }

    # ── PER-GAME LOG RESOLVER ─────────────────────────────────────────────────
    # Given a profile hashtable, returns the list of absolute log file paths
    # to tail right now.  Handles all 4 LogStrategy values.
    function _ResolveGameLogFiles {
        param([hashtable]$Profile)
        if ($Profile.DisableFileTail -eq $true) { return [System.Collections.Generic.List[string]]::new() }
        $strategy = if ($Profile.LogStrategy) { $Profile.LogStrategy } else { 'SingleFile' }
        $files    = [System.Collections.Generic.List[string]]::new()
        $rootUsed = ''
        $sourceDetail = ''

        switch ($strategy) {

            'PZSessionFolder' {
                $resolvedPz = _ResolveProjectZomboidSessionLogFiles -Profile $Profile
                $rootUsed = $resolvedPz.RootUsed
                $searchDirText = if ($resolvedPz.SearchDirs -and @($resolvedPz.SearchDirs).Count -gt 0) { @($resolvedPz.SearchDirs) -join ', ' } else { '<none>' }
                foreach ($path in @($resolvedPz.Files)) {
                    if ($path -and -not $files.Contains($path)) { $files.Add($path) }
                }
                if (-not [string]::IsNullOrWhiteSpace([string]$resolvedPz.SessionKey)) {
                    $sessionDirText = if ($resolvedPz.SessionDir) { $resolvedPz.SessionDir.FullName } else { $rootUsed }
                    $sourceDetail = "pz-session key=$($resolvedPz.SessionKey) dir=$sessionDirText searched=$searchDirText"
                } elseif ($files.Count -gt 0) {
                    $sourceDetail = "pz-fallback searched=$searchDirText"
                } else {
                    $sourceDetail = "pz-none searched=$searchDirText"
                }
                break
            }

            'ValheimUserFolder' {
                # ServerLogRoot = %AppData%\..\LocalLow\IronGate\Valheim
                $root = _ExpandPathVars ([string]$Profile.ServerLogRoot)
                $rootUsed = $root
                if ([string]::IsNullOrEmpty($root)) {
                    $localLow = _GetWindowsLocalLowPath
                    if (-not [string]::IsNullOrWhiteSpace($localLow)) {
                        $root = Join-Path $localLow 'IronGate\Valheim'
                    }
                    $rootUsed = $root
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
                $rootUsed = $root
                if ([string]::IsNullOrEmpty($root) -or -not (Test-Path $root)) {
                    $sourceDetail = 'newestfile-root-missing'
                    break
                }
                # Try preferred filename first (supports wildcards) - skip unexpanded tokens
                $preferred = $Profile.ServerLogFile
                if (-not [string]::IsNullOrEmpty($preferred) -and
                    $preferred -ne '*.log' -and $preferred -ne '$LOGFILE') {
                    if ($preferred -match '[\*\?]') {
                        $match = Get-ChildItem -Path $root -Filter $preferred -ErrorAction SilentlyContinue |
                                 Sort-Object LastWriteTime -Descending | Select-Object -First 1
                        if ($match) {
                            $files.Add($match.FullName)
                            $sourceDetail = "preferred-pattern=$preferred matched=$($match.Name)"
                            break
                        }
                    } else {
                        $pf = Join-Path $root $preferred
                        if (Test-Path $pf) {
                            $files.Add($pf)
                            $sourceDetail = "preferred-file=$preferred"
                            break
                        }
                    }
                }
                # Search *.log first, then *.txt (covers 7DTD output_log_dedi__*.txt)
                $newest = Get-ChildItem -Path $root -Filter '*.log' -ErrorAction SilentlyContinue |
                          Sort-Object LastWriteTime -Descending | Select-Object -First 1
                if ($newest) {
                    $files.Add($newest.FullName)
                    $sourceDetail = "newest-log=$($newest.Name)"
                    break
                }
                $newest = Get-ChildItem -Path $root -Filter '*.txt' -ErrorAction SilentlyContinue |
                          Where-Object { $_.Name -match 'output_log' } |
                          Sort-Object LastWriteTime -Descending | Select-Object -First 1
                if ($newest) {
                    $files.Add($newest.FullName)
                    $sourceDetail = "output-log=$($newest.Name)"
                }
                if ($files.Count -eq 0) {
                    $newest = Get-ChildItem -Path $root -File -ErrorAction SilentlyContinue |
                              Sort-Object LastWriteTime -Descending | Select-Object -First 1
                    if ($newest) {
                        $files.Add($newest.FullName)
                        $sourceDetail = "newest-any=$($newest.Name)"
                    }
                }
                if ([string]::IsNullOrWhiteSpace($sourceDetail) -and $files.Count -eq 0) {
                    $sourceDetail = 'newestfile-no-match'
                }
            }

            default {
                # SingleFile  -  use ServerLogPath directly
                $lp = _ExpandPathVars ([string]$Profile.ServerLogPath)
                $rootUsed = $lp
                if (-not [string]::IsNullOrEmpty($lp) -and (Test-Path $lp)) {
                    $files.Add($lp)
                }
                $knownGame = ''
                try { $knownGame = (_NormalizeGameIdentity (_GetProfileKnownGame -Profile $Profile)) } catch { $knownGame = '' }
                if ($knownGame -eq 'valheim') {
                    $logRoot = _ExpandPathVars ([string]$Profile.ServerLogRoot)
                    if ([string]::IsNullOrWhiteSpace($logRoot) -and -not [string]::IsNullOrWhiteSpace($lp)) {
                        try { $logRoot = Split-Path -Path $lp -Parent } catch { $logRoot = '' }
                    }
                    if (-not [string]::IsNullOrWhiteSpace($logRoot) -and (Test-Path $logRoot)) {
                        $connLog = Get-ChildItem -Path $logRoot -Filter 'connection_log_*.txt' -File -ErrorAction SilentlyContinue |
                                   Sort-Object LastWriteTime -Descending | Select-Object -First 1
                        if ($connLog -and -not $files.Contains($connLog.FullName)) {
                            $files.Add($connLog.FullName)
                            if ([string]::IsNullOrWhiteSpace($sourceDetail)) {
                                $sourceDetail = "valheim-singlefile+connection=$($connLog.Name)"
                            } else {
                                $sourceDetail += " connection=$($connLog.Name)"
                            }
                        }
                    }
                }
            }
        }
        try {
            if ($script:SharedState -and $script:SharedState.Settings -and [bool]$script:SharedState.Settings.EnableDebugLogging) {
                $resolved = if ($files.Count -gt 0) { $files -join '; ' } else { '<none>' }
                $gameName = if ($Profile.GameName) { [string]$Profile.GameName } else { '<unknown>' }
                $prefix = if ($Profile.Prefix) { [string]$Profile.Prefix } else { '' }
                if ($script:SharedState.LogQueue) {
                    $detail = if ([string]::IsNullOrWhiteSpace($sourceDetail)) { '<n/a>' } else { $sourceDetail }
                    $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][DEBUG][GUI] Log resolve [$prefix][$gameName] strategy=$strategy root=$rootUsed source=$detail files=$resolved")
                }
            }
        } catch { }
        return $files
    }

    # ── ENSURE GAME LOG TAB  -  idempotent, called each timer tick ────────────
    function _EnsureGameLogTab {
        param([string]$Prefix, [string]$GameName, [hashtable]$Profile = $null)
        if ($script:_GameLogTabs.ContainsKey($Prefix)) { return }

        $tabPage           = New-Object System.Windows.Forms.TabPage
        $tabPage.Text      = $GameName
        $tabPage.BackColor = $clrBg
        $script:_LogTabControl.TabPages.Add($tabPage)
        _RecalcTabWidths

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
        $footer.BackColor  = $clrPanelAlt
        $inner.Controls.Add($footer)

        # Source label
        $lblSrc            = _Label 'Waiting for log file...' 0 1 700 18
        $lblSrc.ForeColor  = $clrMuted
        $lblSrc.Font       = New-Object System.Drawing.Font('Consolas', 7)
        $lblSrc.Dock       = 'Fill'
        $lblSrc.TextAlign  = 'MiddleLeft'

        $footerLabelHost = New-Object System.Windows.Forms.Panel
        $footerLabelHost.Dock = 'Fill'
        $footerLabelHost.Padding = [System.Windows.Forms.Padding]::new(10, 5, 8, 5)
        $footerLabelHost.BackColor = [System.Drawing.Color]::Transparent
        $footerLabelHost.Controls.Add($lblSrc)
        $footer.Controls.Add($footerLabelHost)

        $gameLogPrefixLocal = [string]$Prefix
        $gameLogBoxLocal = $rtGame

        # Clear button (right aligned)
        $btnCopyGame = _Button 'Copy' 0 0 72 24 $clrPanel {
            if ([string]::IsNullOrWhiteSpace($gameLogPrefixLocal)) { return }
            $ok = $false
            if ($gameLogBoxLocal -is [System.Windows.Forms.RichTextBox]) {
                $text = ''
                try { $text = [string]$gameLogBoxLocal.Text } catch { $text = '' }
                if (-not [string]::IsNullOrWhiteSpace($text)) {
                    try {
                        [System.Windows.Forms.Clipboard]::SetText($text)
                        $ok = $true
                    } catch {
                        $ok = $false
                    }
                }
            }

            if (-not $ok) {
                [System.Windows.Forms.MessageBox]::Show(
                    "There is no visible log text to copy for $gameLogPrefixLocal yet.",
                    'Copy Game Log','OK','Information') | Out-Null
            }
        }.GetNewClosure()
        $btnCopyGame.Margin = [System.Windows.Forms.Padding]::new(0, 0, 8, 0)
        _SetMainControlToolTip -Control $btnCopyGame -Text "Copy the visible $gameLogPrefixLocal log text to the clipboard."

        $btnClearGame = _Button 'Clear' 0 0 72 24 $clrMuted {
            if ([string]::IsNullOrWhiteSpace($gameLogPrefixLocal)) { return }
            if ($gameLogBoxLocal -is [System.Windows.Forms.RichTextBox]) {
                $gameLogBoxLocal.Clear()
            }
            if ($script:_GameLogTabs -and ($script:_GameLogTabs -is [System.Collections.IDictionary]) -and $script:_GameLogTabs.ContainsKey($gameLogPrefixLocal)) {
                $entry = $script:_GameLogTabs[$gameLogPrefixLocal]
                if ($entry -and ($entry.PSObject.Properties.Name -contains 'Files')) {
                    $entry.Files = @{}
                }
            }
        }.GetNewClosure()
        $btnClearGame.Margin = [System.Windows.Forms.Padding]::new(0)
        _SetMainControlToolTip -Control $btnClearGame -Text "Clear the visible $gameLogPrefixLocal log in the ECC window."

        $footerActions = New-Object System.Windows.Forms.FlowLayoutPanel
        $footerActions.Dock = 'Right'
        $footerActions.Width = 160
        $footerActions.WrapContents = $false
        $footerActions.FlowDirection = 'LeftToRight'
        $footerActions.Padding = [System.Windows.Forms.Padding]::new(0, 2, 8, 2)
        $footerActions.Margin = [System.Windows.Forms.Padding]::new(0)
        $footerActions.BackColor = [System.Drawing.Color]::Transparent
        [void]$footerActions.Controls.Add($btnCopyGame)
        [void]$footerActions.Controls.Add($btnClearGame)
        $footer.Controls.Add($footerActions)

        # Register tab entry
        $script:_GameLogTabs[$Prefix] = @{
            Tab         = $tabPage
            RTB         = $rtGame
            LblSrc      = $lblSrc
            Files       = @{}
            LastSession = ''
            NoteShown   = $false
            HistoryLoaded = $false
            PendingLines  = New-Object 'System.Collections.Generic.List[string]'
        }

        if ($Profile) {
            _LoadRecentGameLogHistory -Prefix $Prefix -Profile $Profile
        }
    }


    # ── REMOVE GAME LOG TAB  -  called when a server stops ───────────────────
    function _RemoveGameLogTab {
        param([string]$Prefix)
        if (-not $script:_GameLogTabs.ContainsKey($Prefix)) { return }
        $entry = $script:_GameLogTabs[$Prefix]
        try {
            if ($null -ne $script:_LogTabControl -and
                $null -ne $entry -and
                $null -ne $entry.Tab -and
                $script:_LogTabControl.SelectedTab -eq $entry.Tab -and
                $script:_LogTabControl.TabPages.Count -gt 0) {
                $script:_LogTabControl.SelectedIndex = 0
            }
        } catch { }

        try { if ($entry.PendingLines) { $entry.PendingLines.Clear() } } catch { }
        try {
            if ($null -ne $script:_LogTabControl -and $null -ne $entry -and $null -ne $entry.Tab) {
                $script:_LogTabControl.TabPages.Remove($entry.Tab)
            }
        } catch { }
        $script:_GameLogTabs.Remove($Prefix)
        $script:_ServerStartNotified.Remove($Prefix) | Out-Null
        $script:_PlayersCapture.Remove($Prefix) | Out-Null
        $script:_HytaleWhoCapture.Remove($Prefix) | Out-Null
        $script:_SatisfactoryConnectionCapture.Remove($Prefix) | Out-Null
        try { $script:_ValheimPlayerCapture.Remove($Prefix) | Out-Null } catch { }
        try {
            if ($script:SharedState -and $script:SharedState.ContainsKey('PzObservedPlayerIds') -and $script:SharedState.PzObservedPlayerIds) {
                $script:SharedState.PzObservedPlayerIds.Remove($Prefix) | Out-Null
            }
        } catch { }
        _RecalcTabWidths
    }

    $logTabs.Add_SelectedIndexChanged({
        try {
            $selected = $script:_LogTabControl.SelectedTab
            if ($null -eq $selected) { return }

            foreach ($pfx in @($script:_GameLogTabs.Keys)) {
                $entry = $script:_GameLogTabs[$pfx]
                if ($null -eq $entry -or $entry.Tab -ne $selected) { continue }

                $profile = $null
                if ($script:SharedState -and $script:SharedState.Profiles -and $script:SharedState.Profiles.ContainsKey($pfx)) {
                    $profile = $script:SharedState.Profiles[$pfx]
                }

                if ($profile) {
                    _LoadRecentGameLogHistory -Prefix $pfx -Profile $profile
                }
                _FlushPendingGameLogLines -Prefix $pfx
                break
            }
        } catch { }
    })

    # =====================================================================
    # BACKGROUND METRICS RUNSPACE
    # =====================================================================
    $script:_cpuSmooth = 0.0
    $script:_ramSmooth = 0.0
    $script:_netSmooth = 0.0
    $script:_diskMetricSnapshot = $null
    $script:_diskMetricSnapshotAt = $null

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
        $lastCpu   = 0.0
        $lastMetricsErrorAt = [datetime]::MinValue

        while (-not ($SharedState -and $SharedState.ContainsKey('StopMetricsWorker') -and $SharedState['StopMetricsWorker'] -eq $true)) {
            try {
                Start-Sleep -Seconds 2
                if ($SharedState -and $SharedState.ContainsKey('StopMetricsWorker') -and $SharedState['StopMetricsWorker'] -eq $true) { break }

                # CPU - Get-Counter can occasionally throw transient negative-denominator
                # perf-counter errors even on healthy systems. Reuse the last sample quietly.
                $cpu = $lastCpu
                try {
                    $cpuSample = (Get-Counter '\Processor(_Total)\% Processor Time' -ErrorAction Stop).CounterSamples.CookedValue
                    if ($null -ne $cpuSample) {
                        $cpu = [double]$cpuSample
                    }
                } catch {
                    $cpuMessage = "$($_.Exception.Message)"
                    $knownCounterGlitch =
                        ($cpuMessage -match 'negative denominator value') -or
                        ($cpuMessage -match 'negative denominator')
                    if (-not $knownCounterGlitch) {
                        throw
                    }
                }
                $lastCpu = [double]$cpu

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
            } catch {
                try {
                    $now = Get-Date
                    if (($now - $lastMetricsErrorAt).TotalSeconds -ge 30 -and $SharedState -and $SharedState.LogQueue) {
                        $lastMetricsErrorAt = $now
                        $SharedState.LogQueue.Enqueue("[$($now.ToString('yyyy-MM-dd HH:mm:ss'))][WARN][GUI] Metrics worker sample failed: $($_.Exception.Message)")
                    }
                } catch { }
            }
        }
    }) | Out-Null
    $script:MetricsHandle = $script:MetricsPS.BeginInvoke()

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
        $lastLogTailErrorAt = [datetime]::MinValue
        $logPerfState = @{}

        function _TraceLogTailPerformance {
            param(
                [string]$Area,
                [double]$ElapsedMs,
                [double]$WarnAtMs = 120,
                [double]$DebugAtMs = 45,
                [string]$Detail = ''
            )

            try {
                if ([string]::IsNullOrWhiteSpace($Area)) { return }
                if ($ElapsedMs -lt 0) { return }
                if (-not ($SharedState -and $SharedState.LogQueue)) { return }

                $debugEnabled = $false
                try {
                    $debugEnabled = (($SharedState.Settings -and [bool]$SharedState.Settings.EnablePerformanceDebugMode) -or
                                     ($SharedState.Settings -and [bool]$SharedState.Settings.EnableDebugLogging))
                } catch {
                    $debugEnabled = $false
                }

                $level = ''
                if ($ElapsedMs -ge $WarnAtMs) {
                    $level = 'WARN'
                } elseif ($debugEnabled -and $ElapsedMs -ge $DebugAtMs) {
                    $level = 'DEBUG'
                } else {
                    return
                }

                $now = Get-Date
                $bucket = [int][Math]::Floor(([Math]::Max(0.0, [double]$ElapsedMs)) / 25.0)
                $minGapSeconds = if ($level -eq 'WARN') { 10 } else { 30 }
                $shouldLog = $true
                $last = $null
                try { if ($logPerfState.ContainsKey($Area)) { $last = $logPerfState[$Area] } } catch { $last = $null }

                if ($last) {
                    $lastAt = $null
                    $lastLevel = ''
                    $lastBucket = -1
                    try { $lastAt = $last.At } catch { $lastAt = $null }
                    try { $lastLevel = [string]$last.Level } catch { $lastLevel = '' }
                    try { $lastBucket = [int]$last.Bucket } catch { $lastBucket = -1 }

                    if ($lastLevel -eq $level -and
                        $lastAt -is [datetime] -and
                        (($now - $lastAt).TotalSeconds -lt $minGapSeconds) -and
                        ([Math]::Abs($bucket - $lastBucket) -lt 2)) {
                        $shouldLog = $false
                    }
                }

                if (-not $shouldLog) { return }

                $logPerfState[$Area] = @{
                    At     = $now
                    Level  = $level
                    Bucket = $bucket
                }

                $message = '[{0}][{1}][GUI] LOGPERF area={2} elapsedMs={3:N1} warnAtMs={4:N0}' -f `
                    $now.ToString('yyyy-MM-dd HH:mm:ss'),
                    $level,
                    $Area,
                    ([Math]::Round($ElapsedMs, 1)),
                    ([Math]::Round($WarnAtMs, 0))
                if (-not [string]::IsNullOrWhiteSpace($Detail)) {
                    $message += ' detail=' + $Detail
                }
                $SharedState.LogQueue.Enqueue($message)
            } catch { }
        }

        function _ExpandPathVars {
            param([string]$Path)
            if ([string]::IsNullOrWhiteSpace($Path)) { return $Path }
            return [Environment]::ExpandEnvironmentVariables($Path)
        }

        function _GetProjectZomboidLatestSessionFolder {
            param([string]$RootPath)

            if ([string]::IsNullOrWhiteSpace($RootPath)) { return $null }
            if (-not (Test-Path -LiteralPath $RootPath)) { return $null }

            $dirs = @(Get-ChildItem -Path $RootPath -Directory -ErrorAction SilentlyContinue)
            if ($dirs.Count -eq 0) { return $null }

            $dated = @($dirs | Where-Object { $_.Name -match '^logs_\d{4}-\d{2}-\d{2}$' -or $_.Name -match '^\d{4}-\d{2}-\d{2}_\d{2}-\d{2}' } | Sort-Object Name -Descending)
            if ($dated.Count -gt 0) { return $dated[0] }

            return ($dirs | Sort-Object LastWriteTime -Descending | Select-Object -First 1)
        }

        function _GetProjectZomboidLogDirectories {
            param([string]$RootPath)

            $dirs = New-Object 'System.Collections.Generic.List[string]'
            if ([string]::IsNullOrWhiteSpace($RootPath)) { return $dirs }
            if (-not (Test-Path -LiteralPath $RootPath)) { return $dirs }

            $dirs.Add($RootPath) | Out-Null
            $latestFolder = _GetProjectZomboidLatestSessionFolder -RootPath $RootPath
            if ($latestFolder -and -not $dirs.Contains($latestFolder.FullName)) {
                $dirs.Add($latestFolder.FullName) | Out-Null
            }

            return $dirs
        }

        function _GetProjectZomboidSessionKeyFromPath {
            param([string]$Path)

            if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
            $leaf = ''
            try { $leaf = [System.IO.Path]::GetFileName($Path) } catch { $leaf = $Path }
            if ($leaf -match '^(\d{4}-\d{2}-\d{2}_\d{2}-\d{2})_') {
                return $Matches[1]
            }
            return ''
        }

        function _ResolveProjectZomboidSessionLogFiles {
            param($Profile)

            $files = [System.Collections.Generic.List[string]]::new()
            $root = _ExpandPathVars ([string]$Profile.ServerLogRoot)
            $rootUsed = $root
            $searchDirs = @(_GetProjectZomboidLogDirectories -RootPath $root)
            if ($searchDirs.Count -eq 0) {
                return [pscustomobject]@{
                    Files      = $files
                    RootUsed   = $rootUsed
                    SessionDir = $null
                    SessionKey = ''
                    SearchDirs = @()
                }
            }
            $patterns = @(
                $(if ($Profile.ServerLogFileDebug) { [string]$Profile.ServerLogFileDebug } else { '' }),
                $(if ($Profile.ServerLogFileUser)  { [string]$Profile.ServerLogFileUser } else { '' }),
                $(if ($Profile.ServerLogFileChat)  { [string]$Profile.ServerLogFileChat } else { '' }),
                'DebugLog-server.txt',
                '*_DebugLog-server.txt',
                '*_user.txt',
                '*_chat.txt'
            ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique

            $candidateFiles = New-Object 'System.Collections.Generic.List[object]'
            foreach ($dir in $searchDirs) {
                foreach ($pattern in $patterns) {
                    foreach ($match in @(Get-ChildItem -Path $dir -Filter $pattern -File -ErrorAction SilentlyContinue)) {
                        $candidateFiles.Add([pscustomobject]@{
                            File       = $match
                            SessionKey = (_GetProjectZomboidSessionKeyFromPath -Path $match.Name)
                        }) | Out-Null
                    }
                }
            }

            $bestSession = ''
            $bestSessionDir = $null
            $sessionCandidates = @($candidateFiles | Where-Object { -not [string]::IsNullOrWhiteSpace($_.SessionKey) } | Sort-Object { $_.File.LastWriteTime } -Descending)
            if ($sessionCandidates.Count -gt 0) {
                $bestSession = [string]$sessionCandidates[0].SessionKey
                $bestSessionDir = $sessionCandidates[0].File.Directory
                $rootUsed = $bestSessionDir.FullName
                foreach ($candidate in $sessionCandidates) {
                    if ($candidate.SessionKey -ne $bestSession) { continue }
                    if (-not $files.Contains($candidate.File.FullName)) {
                        $files.Add($candidate.File.FullName) | Out-Null
                    }
                }
            } else {
                foreach ($pattern in $patterns) {
                    $match = $candidateFiles |
                        Where-Object { $_.File.Name -like $pattern } |
                        Sort-Object { $_.File.LastWriteTime } -Descending |
                        Select-Object -First 1
                    if ($match -and -not $files.Contains($match.File.FullName)) {
                        $files.Add($match.File.FullName) | Out-Null
                        $rootUsed = $match.File.DirectoryName
                    }
                }
            }

            return [pscustomobject]@{
                Files      = $files
                RootUsed   = $rootUsed
                SessionDir = $bestSessionDir
                SessionKey = $bestSession
                SearchDirs = @($searchDirs)
            }
        }

        function _ResolveLogFiles {
            param($Profile)
            if ($Profile.DisableFileTail -eq $true) { return [System.Collections.Generic.List[string]]::new() }
            $strategy = if ($Profile.LogStrategy) { $Profile.LogStrategy } else { 'SingleFile' }
            $files    = [System.Collections.Generic.List[string]]::new()
            $rootUsed = ''
            $sourceDetail = ''

            switch ($strategy) {
                'PZSessionFolder' {
                    $resolvedPz = _ResolveProjectZomboidSessionLogFiles -Profile $Profile
                    $rootUsed = $resolvedPz.RootUsed
                    $searchDirText = if ($resolvedPz.SearchDirs -and @($resolvedPz.SearchDirs).Count -gt 0) { @($resolvedPz.SearchDirs) -join ', ' } else { '<none>' }
                    foreach ($path in @($resolvedPz.Files)) {
                        if ($path -and -not $files.Contains($path)) { $files.Add($path) }
                    }
                    if (-not [string]::IsNullOrWhiteSpace([string]$resolvedPz.SessionKey)) {
                        $sessionDirText = if ($resolvedPz.SessionDir) { $resolvedPz.SessionDir.FullName } else { $rootUsed }
                        $sourceDetail = "pz-session key=$($resolvedPz.SessionKey) dir=$sessionDirText searched=$searchDirText"
                    } elseif ($files.Count -gt 0) {
                        $sourceDetail = "pz-fallback searched=$searchDirText"
                    } else {
                        $sourceDetail = "pz-none searched=$searchDirText"
                    }
                    break
                }

                'ValheimUserFolder' {
                    $root = _ExpandPathVars ([string]$Profile.ServerLogRoot)
                    $rootUsed = $root
                    if ([string]::IsNullOrEmpty($root)) {
                        $localLow = _GetWindowsLocalLowPath
                        if (-not [string]::IsNullOrWhiteSpace($localLow)) {
                            $root = Join-Path $localLow 'IronGate\Valheim'
                        }
                        $rootUsed = $root
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
                    $rootUsed = $root
                    if ([string]::IsNullOrEmpty($root) -or -not (Test-Path $root)) {
                        $sourceDetail = 'newestfile-root-missing'
                        break
                    }
                    $preferred = $Profile.ServerLogFile
                    if (-not [string]::IsNullOrEmpty($preferred) -and
                        $preferred -ne '*.log' -and $preferred -ne '$LOGFILE') {
                        if ($preferred -match '[\*\?]') {
                            $match = Get-ChildItem -Path $root -Filter $preferred -ErrorAction SilentlyContinue |
                                     Sort-Object LastWriteTime -Descending | Select-Object -First 1
                            if ($match) {
                                $files.Add($match.FullName)
                                $sourceDetail = "preferred-pattern=$preferred matched=$($match.Name)"
                                break
                            }
                        } else {
                            $pf = Join-Path $root $preferred
                            if (Test-Path $pf) {
                                $files.Add($pf)
                                $sourceDetail = "preferred-file=$preferred"
                                break
                            }
                        }
                    }
                    # Search *.log first, then *.txt (covers 7DTD output_log_dedi__*.txt)
                    $newest = Get-ChildItem -Path $root -Filter '*.log' -ErrorAction SilentlyContinue |
                              Sort-Object LastWriteTime -Descending | Select-Object -First 1
                    if ($newest) {
                        $files.Add($newest.FullName)
                        $sourceDetail = "newest-log=$($newest.Name)"
                        break
                    }
                    $newest = Get-ChildItem -Path $root -Filter '*.txt' -ErrorAction SilentlyContinue |
                              Where-Object { $_.Name -match 'output_log' } |
                              Sort-Object LastWriteTime -Descending | Select-Object -First 1
                    if ($newest) {
                        $files.Add($newest.FullName)
                        $sourceDetail = "output-log=$($newest.Name)"
                    }
                    if ($files.Count -eq 0) {
                        $newest = Get-ChildItem -Path $root -File -ErrorAction SilentlyContinue |
                                  Sort-Object LastWriteTime -Descending | Select-Object -First 1
                        if ($newest) {
                            $files.Add($newest.FullName)
                            $sourceDetail = "newest-any=$($newest.Name)"
                        }
                    }
                    if ([string]::IsNullOrWhiteSpace($sourceDetail) -and $files.Count -eq 0) {
                        $sourceDetail = 'newestfile-no-match'
                    }
                }

                default {
                    $lp = _ExpandPathVars ([string]$Profile.ServerLogPath)
                    $rootUsed = $lp
                    if (-not [string]::IsNullOrEmpty($lp)) {
                        if ($lp -match '[\*\?]') {
                            $root    = Split-Path $lp -Parent
                            $pattern = Split-Path $lp -Leaf
                            if ($root -and (Test-Path $root)) {
                                $match = Get-ChildItem -Path $root -Filter $pattern -ErrorAction SilentlyContinue |
                                         Sort-Object LastWriteTime -Descending | Select-Object -First 1
                                if ($match) { $files.Add($match.FullName) }
                            }
                        } elseif (Test-Path $lp) {
                            $files.Add($lp)
                        }
                    }
                    $knownGame = ''
                    try { $knownGame = (_NormalizeGameIdentity (_GetProfileKnownGame -Profile $Profile)) } catch { $knownGame = '' }
                    if ($knownGame -eq 'valheim') {
                        $logRoot = _ExpandPathVars ([string]$Profile.ServerLogRoot)
                        if ([string]::IsNullOrWhiteSpace($logRoot) -and -not [string]::IsNullOrWhiteSpace($lp)) {
                            try { $logRoot = Split-Path -Path $lp -Parent } catch { $logRoot = '' }
                        }
                        if (-not [string]::IsNullOrWhiteSpace($logRoot) -and (Test-Path $logRoot)) {
                            $connLog = Get-ChildItem -Path $logRoot -Filter 'connection_log_*.txt' -File -ErrorAction SilentlyContinue |
                                       Sort-Object LastWriteTime -Descending | Select-Object -First 1
                            if ($connLog -and -not $files.Contains($connLog.FullName)) {
                                $files.Add($connLog.FullName)
                                if ([string]::IsNullOrWhiteSpace($sourceDetail)) {
                                    $sourceDetail = "valheim-singlefile+connection=$($connLog.Name)"
                                } else {
                                    $sourceDetail += " connection=$($connLog.Name)"
                                }
                            }
                        }
                    }
                }
            }
            try {
                if ($SharedState -and $SharedState.Settings -and [bool]$SharedState.Settings.EnableDebugLogging -and $SharedState.LogQueue) {
                    $resolved = if ($files.Count -gt 0) { $files -join '; ' } else { '<none>' }
                    $gameName = if ($Profile.GameName) { [string]$Profile.GameName } else { '<unknown>' }
                    $prefix = if ($Profile.Prefix) { [string]$Profile.Prefix } else { '' }
                    $detail = if ([string]::IsNullOrWhiteSpace($sourceDetail)) { '<n/a>' } else { $sourceDetail }
                    $SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][DEBUG][LogTail] Resolve [$prefix][$gameName] strategy=$strategy root=$rootUsed source=$detail files=$resolved")
                }
            } catch { }
            return $files
        }

        # Maximum bytes to read from a single file in one tick.
        # Reading more than this per tick means the file is writing faster than
        # we can display — in that case we jump forward and show a skip notice.
        # Keep this intentionally small because ECC may tail multiple busy servers
        # at once and the UI only needs enough recent text to detect readiness and
        # show a small live window.
        $maxReadBytes = 4096L   # 4 KB per file per tick
        $prioritySignalScanBytes = 65536L # up to 64 KB of skipped text scanned for readiness markers only

        function _GetPriorityLogSignalLines {
            param(
                [string]$Prefix = '',
                [string]$Text = ''
            )

            $signals = New-Object 'System.Collections.Generic.List[string]'
            if ([string]::IsNullOrWhiteSpace($Text)) { return @($signals) }

            $normalizedPrefix = if ([string]::IsNullOrWhiteSpace($Prefix)) { '' } else { $Prefix.ToUpperInvariant() }
            $lines = @($Text -split "(`r`n|`n|`r)")
            foreach ($rawLine in $lines) {
                $line = [string]$rawLine
                if ([string]::IsNullOrWhiteSpace($line)) { continue }

                $isPriority = $false
                switch ($normalizedPrefix) {
                    'MC' {
                        $isPriority = (
                            $line -match 'For help, type' -or
                            $line -match 'RCON running on' -or
                            $line -match 'Thread RCON Listener started' -or
                            $line -match '(?i)\bthere are \d+ of a max of \d+ players online\b'
                        )
                    }
                    'PZ' { $isPriority = ($line -match 'SERVER STARTED') }
                    'HY' { $isPriority = ($line -match 'Hytale Server Booted' -or $line -match 'Server Booted') }
                    'VH' { $isPriority = ($line -match 'Game server connected') }
                    'DZ' { $isPriority = ($line -match 'INF StartGame done') }
                    'SF' {
                        $isPriority = (
                            $line -match 'Server startup time elapsed and saving/level loading is done' -and
                            $line -match 'WorldTimeSeconds\s*=\s*([0-9]+(?:\.[0-9]+)?)'
                        )
                    }
                }

                if ($isPriority) {
                    $signals.Add($line) | Out-Null
                }
            }

            return @($signals)
        }

        function _ReadNewLines {
            param(
                [string]$Path,
                [string]$Prefix = ''
            )
            $readPerf = [System.Diagnostics.Stopwatch]::StartNew()
            if (-not (Test-Path $Path)) { return @() }

            $len = (Get-Item $Path).Length

            if (-not $filePos.ContainsKey($Path)) {
                # First time we see this file: anchor to the current end.
                # We never want to dump the entire existing log into the UI.
                $filePos[$Path]    = $len
                $remainders[$Path] = ''
                try {
                    $readPerf.Stop()
                    _TraceLogTailPerformance -Area 'FileAnchor' -ElapsedMs $readPerf.Elapsed.TotalMilliseconds -WarnAtMs 80 -DebugAtMs 30 -Detail ('prefix={0};file={1};len={2}' -f $Prefix, ([System.IO.Path]::GetFileName($Path)), $len)
                } catch { }
                return @()
            }

            # File was truncated or rotated — reset cursor to beginning.
            if ($len -lt $filePos[$Path]) {
                $filePos[$Path]    = 0
                $remainders[$Path] = ''
            }

            if ($len -eq $filePos[$Path]) {
                try { $readPerf.Stop() } catch { }
                return @()
            }

            # --- Always-tail logic ---
            # If more than $maxReadBytes of new content has appeared since last
            # tick, skip ahead so we read only the LAST $maxReadBytes.
            # This guarantees the display stays glued to the bottom of the log
            # regardless of how fast the server is writing.
            $available    = $len - $filePos[$Path]
            $skippedBytes = 0L
            $prioritySignalLines = @()
            if ($available -gt $maxReadBytes) {
                $scanStart = [Math]::Max([int64]$filePos[$Path], [int64]($len - $maxReadBytes - $prioritySignalScanBytes))
                $scanLength = [int64](($len - $maxReadBytes) - $scanStart)
                if ($scanLength -gt 0) {
                    $scanFs = [System.IO.File]::Open(
                        $Path,
                        [System.IO.FileMode]::Open,
                        [System.IO.FileAccess]::Read,
                        [System.IO.FileShare]::ReadWrite
                    )
                    try {
                        $scanFs.Seek($scanStart, [System.IO.SeekOrigin]::Begin) | Out-Null
                        $scanBuffer = New-Object byte[] $scanLength
                        $bytesRead = $scanFs.Read($scanBuffer, 0, [int]$scanLength)
                        if ($bytesRead -gt 0) {
                            $scanText = [System.Text.Encoding]::UTF8.GetString($scanBuffer, 0, $bytesRead)
                            $prioritySignalLines = @(_GetPriorityLogSignalLines -Prefix $Prefix -Text $scanText)
                        }
                    } finally {
                        $scanFs.Close()
                    }
                }

                $skippedBytes           = $available - $maxReadBytes
                $filePos[$Path]         = $len - $maxReadBytes
                $remainders[$Path]      = ''   # discard partial line before the jump
            }

            $fs = [System.IO.File]::Open(
                $Path,
                [System.IO.FileMode]::Open,
                [System.IO.FileAccess]::Read,
                [System.IO.FileShare]::ReadWrite
            )
            try {
                $fs.Seek($filePos[$Path], [System.IO.SeekOrigin]::Begin) | Out-Null
                $sr   = New-Object System.IO.StreamReader($fs, [System.Text.Encoding]::UTF8, $true)
                $text = $sr.ReadToEnd()
                $filePos[$Path] = $fs.Position
                $sr.Close()
            } finally {
                $fs.Close()
            }

            if ([string]::IsNullOrEmpty($text)) { return @() }

            # Prepend any partial line buffered from the previous tick
            $prior = if ($remainders.ContainsKey($Path)) { $remainders[$Path] } else { '' }
            $text  = $prior + $text
            $lines = $text -split "`r?`n"

            # If the last element has no trailing newline, hold it for next tick
            if (-not ($text.EndsWith("`n") -or $text.EndsWith("`r"))) {
                $remainders[$Path] = $lines[-1]
                $lines = $lines[0..($lines.Count - 2)]
            } else {
                $remainders[$Path] = ''
            }

            # Insert a single skip notice when we jumped forward
            if ($skippedBytes -gt 0) {
                $kb     = [Math]::Round($skippedBytes / 1KB, 1)
                $notice = "... [skipped ${kb} KB - showing latest output] ..."
                $signalNotices = @()
                foreach ($signalLine in @($prioritySignalLines | Select-Object -Unique)) {
                    if ([string]::IsNullOrWhiteSpace($signalLine)) { continue }
                    $signalNotices += "[PRIORITY SIGNAL] $signalLine"
                }
                $lines  = @($notice) + @($signalNotices) + @($lines | Where-Object { $_ -ne '' })
            }

            try {
                $readPerf.Stop()
                $readBytes = [Math]::Max(0L, [int64]$text.Length)
                $lineCount = @($lines).Count
                _TraceLogTailPerformance -Area 'FileRead' `
                    -ElapsedMs $readPerf.Elapsed.TotalMilliseconds `
                    -WarnAtMs 90 `
                    -DebugAtMs 35 `
                    -Detail ('prefix={0};file={1};lines={2};textChars={3};skippedBytes={4}' -f `
                        $Prefix, ([System.IO.Path]::GetFileName($Path)), $lineCount, $readBytes, $skippedBytes)
            } catch { }

            return $lines
        }

        while (-not ($SharedState -and $SharedState.ContainsKey('StopLogTailWorker') -and $SharedState['StopLogTailWorker'] -eq $true)) {
            try {
                $iterationFiles = 0
                $iterationLines = 0
                Start-Sleep -Milliseconds 800
                if ($SharedState -and $SharedState.ContainsKey('StopLogTailWorker') -and $SharedState['StopLogTailWorker'] -eq $true) { break }
                $iterationPerf = [System.Diagnostics.Stopwatch]::StartNew()

                if ($null -eq $SharedState -or $null -eq $SharedState.Profiles) { continue }
                if (-not $SharedState.ContainsKey('GameLogQueue')) { continue }

                foreach ($pfx in @($SharedState.RunningServers.Keys)) {
                    $profilePerf = [System.Diagnostics.Stopwatch]::StartNew()
                    $profileFileCount = 0
                    $profileLineCount = 0
                    $profile = $SharedState.Profiles[$pfx]
                    if ($null -eq $profile) { continue }

                    # Prefer the runtime log path stored at launch (e.g. 7DTD timestamped file)
                    # over the static profile value which may be empty or contain unexpanded tokens.
                    $runtimeLogPath = ''
                    $srvEntry = $SharedState.RunningServers[$pfx]
                    if ($null -ne $srvEntry -and
                        $null -ne $srvEntry.ServerLogPath -and
                        "$($srvEntry.ServerLogPath)".Trim() -ne '') {
                        $runtimeLogPath = "$($srvEntry.ServerLogPath)".Trim()
                    }

                    $files = if ($runtimeLogPath -ne '' -and
                                 $runtimeLogPath -notmatch '[\*\?]' -and
                                 (Test-Path $runtimeLogPath)) {
                        @($runtimeLogPath)
                    } else {
                        _ResolveLogFiles -Profile $profile
                    }
                    foreach ($path in $files) {
                        $profileFileCount++
                        $iterationFiles++
                        $lines = _ReadNewLines -Path $path -Prefix $pfx
                        if ($lines.Count -eq 0) { continue }
                        $profileLineCount += @($lines).Count
                        $iterationLines += @($lines).Count

                        # Hard cap: never enqueue more than 20 lines per file per tick.
                        # _ReadNewLines already limits the read window, but even a small
                        # chunk can still yield a burst of short lines on chatty servers.
                        # This final cap keeps the GameLogQueue shallow.
                        if ($lines.Count -gt 20) {
                            $dropped = $lines.Count - 20
                            $lines   = @("... [+$dropped lines not shown this tick] ...") + $lines[-20..-1]
                        }

                        foreach ($line in $lines) {
                            if ($line -eq '') { continue }
                            $SharedState.GameLogQueue.Enqueue(
                                [pscustomobject]@{ Prefix = $pfx; Line = $line; Path = $path }
                            )
                        }
                    }

                    $profilePerf.Stop()
                    _TraceLogTailPerformance -Area 'ProfileBatch' `
                        -ElapsedMs $profilePerf.Elapsed.TotalMilliseconds `
                        -WarnAtMs 140 `
                        -DebugAtMs 55 `
                        -Detail ('prefix={0};files={1};lines={2}' -f $pfx, $profileFileCount, $profileLineCount)
                }

                $iterationPerf.Stop()
                $runningServerCount = 0
                try { $runningServerCount = @($SharedState.RunningServers.Keys).Count } catch { $runningServerCount = 0 }
                if ($runningServerCount -gt 0 -or $iterationFiles -gt 0 -or $iterationLines -gt 0) {
                    _TraceLogTailPerformance -Area 'WorkerIteration' `
                        -ElapsedMs $iterationPerf.Elapsed.TotalMilliseconds `
                        -WarnAtMs 220 `
                        -DebugAtMs 90 `
                        -Detail ('running={0};files={1};lines={2}' -f $runningServerCount, $iterationFiles, $iterationLines)
                }
            } catch {
                try {
                    $now = Get-Date
                    if (($now - $lastLogTailErrorAt).TotalSeconds -ge 15 -and $SharedState -and $SharedState.LogQueue) {
                        $lastLogTailErrorAt = $now
                        $SharedState.LogQueue.Enqueue("[$($now.ToString('yyyy-MM-dd HH:mm:ss'))][WARN][GUI] Game log tail worker iteration failed: $($_.Exception.Message)")
                    }
                } catch { }
            }
        }
    }) | Out-Null
    $script:LogTailHandle = $script:LogTailPS.BeginInvoke()

    # =====================================================================
    # LOCAL HELPER FUNCTIONS  (inside Start-GUI so they close over locals)
    # =====================================================================
    function _Smooth {
        param([double]$old, [double]$new, [double]$factor = 0.2)
        return ($old + ($new - $old) * $factor)
    }

    function _CaptureScrollPosition([System.Windows.Forms.Control]$Control) {
        $x = 0
        $y = 0
        try {
            if ($Control -is [System.Windows.Forms.ScrollableControl]) {
                if ($Control.HorizontalScroll.Visible) { $x = $Control.HorizontalScroll.Value }
                if ($Control.VerticalScroll.Visible) { $y = $Control.VerticalScroll.Value }
            }
        } catch { }
        return [pscustomobject]@{ X = $x; Y = $y }
    }

    function _RestoreScrollPosition([System.Windows.Forms.Control]$Control, [object]$Position) {
        if ($null -eq $Control -or $null -eq $Position) { return }
        try {
            if ($Control -is [System.Windows.Forms.ScrollableControl]) {
                $Control.AutoScrollPosition = [System.Drawing.Point]::new([int]$Position.X, [int]$Position.Y)
            }
        } catch { }
    }

    function _ClearDashboardFocus([System.Windows.Forms.Control]$FallbackControl = $null) {
        try {
            if ($script:_MainTabs -is [System.Windows.Forms.Control] -and $script:_MainTabs.CanFocus) {
                $script:_MainTabs.Focus() | Out-Null
            } elseif ($FallbackControl -is [System.Windows.Forms.Control] -and $FallbackControl.CanFocus) {
                $FallbackControl.Focus() | Out-Null
            } elseif ($script:_DashboardScrollPanel -is [System.Windows.Forms.Control] -and $script:_DashboardScrollPanel.CanFocus) {
                $script:_DashboardScrollPanel.Focus() | Out-Null
            }
        } catch { }
    }

    function _SetMetricColor {
        param([System.Windows.Forms.Label]$label, [double]$value)
        if     ($value -lt 50) { $label.ForeColor = $clrGreen  }
        elseif ($value -lt 80) { $label.ForeColor = $clrYellow }
        else                   { $label.ForeColor = $clrRed    }
    }

    function _FormatFreeSpaceText {
        param([double]$Bytes)

        if ($Bytes -lt 0) { return 'unavailable' }

        $kb = 1024.0
        $mb = $kb * 1024.0
        $gb = $mb * 1024.0
        $tb = $gb * 1024.0

        if ($Bytes -ge $tb) { return ('{0:N1} TB free' -f ($Bytes / $tb)) }
        if ($Bytes -ge $gb) {
            $gbFree = $Bytes / $gb
            if ($gbFree -lt 10) { return ('{0:N1} GB free' -f $gbFree) }
            return ('{0:N0} GB free' -f $gbFree)
        }
        if ($Bytes -ge $mb) { return ('{0:N0} MB free' -f ($Bytes / $mb)) }
        if ($Bytes -ge $kb) { return ('{0:N0} KB free' -f ($Bytes / $kb)) }
        return ('{0:N0} B free' -f $Bytes)
    }

    function _GetDriveRootFromPath {
        param([string]$Path)

        if ([string]::IsNullOrWhiteSpace($Path)) { return '' }

        $candidate = (_ExpandPathVars $Path)
        if ([string]::IsNullOrWhiteSpace($candidate)) { return '' }
        $candidate = $candidate.Trim().Trim('"')

        try {
            if (-not [System.IO.Path]::IsPathRooted($candidate)) {
                $candidate = [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $candidate))
            }
        } catch { }

        $root = ''
        try { $root = [System.IO.Path]::GetPathRoot($candidate) } catch { $root = '' }
        if ([string]::IsNullOrWhiteSpace($root)) { return '' }

        try { return ([System.IO.Path]::GetPathRoot($root)).ToUpperInvariant() } catch { return $root.ToUpperInvariant() }
    }

    function _GetTrackedDiskMetricSnapshot {
        param([hashtable]$SharedState)

        $profiles = $null
        try {
            if ($SharedState -and $SharedState.ContainsKey('Profiles')) {
                $profiles = $SharedState.Profiles
            }
        } catch { $profiles = $null }

        $eccDrive = _GetDriveRootFromPath -Path (Get-Location).Path
        $profileDriveCounts = @{}
        $profileDriveMap = @{}

        $profileValues = if ($profiles) { @($profiles.Values) } else { @() }
        foreach ($profileEntry in $profileValues) {
            if ($null -eq $profileEntry) { continue }
            $profileRoot = ''
            try { $profileRoot = _GetDriveRootFromPath -Path ([string]$profileEntry.FolderPath) } catch { $profileRoot = '' }
            if ([string]::IsNullOrWhiteSpace($profileRoot)) { continue }
            if (-not $profileDriveCounts.ContainsKey($profileRoot)) {
                $profileDriveCounts[$profileRoot] = 0
                $profileDriveMap[$profileRoot] = New-Object System.Collections.Generic.List[string]
            }
            $profileDriveCounts[$profileRoot] = [int]$profileDriveCounts[$profileRoot] + 1
            try {
                $profileName = [string]$profileEntry.GameName
                if ([string]::IsNullOrWhiteSpace($profileName)) { $profileName = [string]$profileEntry.Prefix }
                if (-not [string]::IsNullOrWhiteSpace($profileName)) {
                    $profileDriveMap[$profileRoot].Add($profileName) | Out-Null
                }
            } catch { }
        }

        $primaryRoot = ''
        if ($profileDriveCounts.Count -gt 0) {
            $primaryRootCandidate = @(
                $profileDriveCounts.GetEnumerator() |
                Sort-Object -Property @{ Expression = { [int]$_.Value }; Descending = $true }, @{ Expression = { [string]$_.Key }; Descending = $false } |
                Select-Object -ExpandProperty Key -First 1
            )
            if ($primaryRootCandidate.Count -gt 0) {
                $primaryRoot = [string]$primaryRootCandidate[0]
            }
        }
        if ([string]::IsNullOrWhiteSpace($primaryRoot)) { $primaryRoot = $eccDrive }

        $orderedRoots = New-Object System.Collections.Generic.List[string]
        foreach ($root in @($primaryRoot, $eccDrive) + @($profileDriveCounts.Keys | Sort-Object)) {
            if ([string]::IsNullOrWhiteSpace($root)) { continue }
            if ($orderedRoots.Contains($root)) { continue }
            $orderedRoots.Add($root) | Out-Null
        }

        $records = New-Object System.Collections.Generic.List[object]
        foreach ($root in $orderedRoots) {
            try {
                $drive = [System.IO.DriveInfo]::new($root)
                $label = $root.TrimEnd('\')
                if (-not $drive.IsReady) {
                    $records.Add([pscustomobject]@{
                        Root = $root
                        Label = $label
                        FreeBytes = -1
                        TotalBytes = -1
                        PercentFree = -1
                        Text = "$label unavailable"
                    }) | Out-Null
                    continue
                }

                $freeBytes = [double]$drive.AvailableFreeSpace
                $totalBytes = [double]$drive.TotalSize
                $percentFree = if ($totalBytes -gt 0) { ($freeBytes / $totalBytes) * 100.0 } else { -1 }
                $records.Add([pscustomobject]@{
                    Root = $root
                    Label = $label
                    FreeBytes = $freeBytes
                    TotalBytes = $totalBytes
                    PercentFree = $percentFree
                    Text = "$label $(_FormatFreeSpaceText -Bytes $freeBytes)"
                }) | Out-Null
            } catch { }
        }

        $primaryRecord = $null
        foreach ($record in $records) {
            if ($record.Root -eq $primaryRoot) {
                $primaryRecord = $record
                break
            }
        }
        if ($null -eq $primaryRecord -and $records.Count -gt 0) { $primaryRecord = $records[0] }

        $summary = 'DISK: --'
        $color = $clrYellow
        $tooltipLines = New-Object System.Collections.Generic.List[string]
        if ($null -ne $primaryRecord) {
            $summary = 'DISK: {0} {1}' -f $primaryRecord.Label, (_FormatFreeSpaceText -Bytes $primaryRecord.FreeBytes)
            if ($primaryRecord.PercentFree -lt 0) {
                $color = $clrYellow
            } elseif ($primaryRecord.PercentFree -lt 10) {
                $color = $clrRed
            } elseif ($primaryRecord.PercentFree -lt 20) {
                $color = $clrYellow
            } else {
                $color = $clrGreen
            }

            $primaryCount = 0
            if ($profileDriveCounts.ContainsKey($primaryRoot)) {
                $primaryCount = [int]$profileDriveCounts[$primaryRoot]
            }
            $primaryLine = if ($primaryCount -gt 0) {
                'Primary tracked drive: {0} ({1} profile{2})' -f $primaryRecord.Label, $primaryCount, $(if ($primaryCount -eq 1) { '' } else { 's' })
            } else {
                'Primary tracked drive: {0}' -f $primaryRecord.Label
            }
            $tooltipLines.Add($primaryLine) | Out-Null
        } else {
            $tooltipLines.Add('No tracked drives are available yet.') | Out-Null
        }

        if (-not [string]::IsNullOrWhiteSpace($eccDrive)) {
            $tooltipLines.Add(('ECC host drive: {0}' -f $eccDrive.TrimEnd('\'))) | Out-Null
        }

        if ($records.Count -gt 0) {
            $tooltipLines.Add('Tracked drives:') | Out-Null
            foreach ($record in $records) {
                $prefix = if ($record.Root -eq $primaryRoot) { 'Primary' } elseif ($record.Root -eq $eccDrive) { 'ECC' } else { 'Drive' }
                $detail = $record.Text
                if ($record.PercentFree -ge 0) {
                    $detail = '{0} ({1:N0}% free)' -f $detail, $record.PercentFree
                }
                $tooltipLines.Add(('{0}: {1}' -f $prefix, $detail)) | Out-Null
                $profilesOnDrive = @()
                try {
                    if ($profileDriveMap.ContainsKey($record.Root)) {
                        $profilesOnDrive = @($profileDriveMap[$record.Root] | Sort-Object -Unique)
                    }
                } catch { $profilesOnDrive = @() }
                if (@($profilesOnDrive).Count -gt 0) {
                    $tooltipLines.Add(('  Profiles: {0}' -f (@($profilesOnDrive) -join ', '))) | Out-Null
                }
            }
        }

        return [pscustomobject]@{
            Summary = $summary
            Tooltip = ($tooltipLines -join [Environment]::NewLine).Trim()
            Color = $color
        }
    }

    function _GetActivePlayersMetricSnapshot {
        param([hashtable]$SharedState)

        $profiles = $null
        try {
            if ($SharedState -and $SharedState.ContainsKey('Profiles')) {
                $profiles = $SharedState.Profiles
            }
        } catch { $profiles = $null }
        $breakdown = @()
        $totalPlayers = 0

        $profileEntries = if ($profiles) { @($profiles.GetEnumerator()) } else { @() }
        foreach ($entry in $profileEntries) {
            $prefix = ''
            $prefixKey = ''
            $profile = $null
            try { $prefix = [string]$entry.Key } catch { $prefix = '' }
            try { $prefixKey = $prefix.ToUpperInvariant() } catch { $prefixKey = $prefix }
            try { $profile = $entry.Value } catch { $profile = $null }
            if ([string]::IsNullOrWhiteSpace($prefixKey) -or $null -eq $profile) { continue }

            $observedCounts = @()
            try {
                if ($SharedState.PlayerActivityState -and $SharedState.PlayerActivityState.ContainsKey($prefixKey)) {
                    $activity = $SharedState.PlayerActivityState[$prefixKey]
                    if ($activity -and [bool]$activity.DetectionSupported -and [bool]$activity.DetectionAvailable) {
                        $observedCounts += [Math]::Max(0, [int]$activity.CurrentCount)
                    }
                    $activityNote = ''
                    try { $activityNote = [string]$activity.Note } catch { $activityNote = '' }
                    if ($activityNote -match '^\s*(\d+)\s+player\(s\)\s+online\.?\s*$') {
                        $observedCounts += [int]$Matches[1]
                    }
                }
            } catch { }
            try {
                if ($SharedState.LatestPlayerCounts -and $SharedState.LatestPlayerCounts.ContainsKey($prefixKey)) {
                    $observedCounts += [Math]::Max(0, [int]$SharedState.LatestPlayerCounts[$prefixKey])
                }
            } catch { }
            try {
                if ($SharedState.LatestPlayers -and $SharedState.LatestPlayers.ContainsKey($prefixKey)) {
                    $observedCounts += @($SharedState.LatestPlayers[$prefixKey] | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }).Count
                }
            } catch { }
            try {
                if ($SharedState.ServerRuntimeState -and $SharedState.ServerRuntimeState.ContainsKey($prefixKey)) {
                    $runtimeEntry = $SharedState.ServerRuntimeState[$prefixKey]
                    $runtimeDetail = ''
                    try { $runtimeDetail = [string]$runtimeEntry.Detail } catch { $runtimeDetail = '' }
                    if ($runtimeDetail -match '^\s*(\d+)\s+player\(s\)\s+online\.?\s*$') {
                        $observedCounts += [int]$Matches[1]
                    }
                }
            } catch { }

            $normalizedKnownGame = ''
            try { $normalizedKnownGame = (_NormalizeGameIdentity (_GetProfileKnownGame -Profile $profile)) } catch { $normalizedKnownGame = '' }
            if ($normalizedKnownGame -eq 'minecraft' -and $SharedState.RunningServers -and $SharedState.RunningServers.ContainsKey($prefixKey) -and $SharedState.LatestPlayerCounts -and $SharedState.LatestPlayerCounts.ContainsKey($prefixKey)) {
                $count = 0
                try { $count = [Math]::Max(0, [int]$SharedState.LatestPlayerCounts[$prefixKey]) } catch { $count = 0 }
                if ($count -le 0) { continue }

                $label = ''
                try { $label = [string]$profile.GameName } catch { $label = '' }
                if ([string]::IsNullOrWhiteSpace($label)) { $label = $prefixKey }

                $breakdown += [pscustomobject]@{
                    Prefix = $prefixKey
                    Name   = $label
                    Count  = $count
                }
                $totalPlayers += $count
                continue
            }

            $count = -1
            foreach ($candidate in @($observedCounts)) {
                $candidateValue = -1
                try { $candidateValue = [int]$candidate } catch { $candidateValue = -1 }
                if ($candidateValue -gt $count) {
                    $count = $candidateValue
                }
            }
            if ($count -le 0) { continue }

            $label = ''
            try { $label = [string]$profile.GameName } catch { $label = '' }
            if ([string]::IsNullOrWhiteSpace($label)) { $label = $prefixKey }

            $breakdown += [pscustomobject]@{
                Prefix = $prefixKey
                Name   = $label
                Count  = $count
            }
            $totalPlayers += $count
        }

        $tooltipLines = @()
        if (@($breakdown).Count -gt 0) {
            $tooltipLines += 'Trusted active players by server:'
            foreach ($item in @($breakdown)) {
                $serverLabel = if (-not [string]::IsNullOrWhiteSpace([string]$item.Name) -and [string]$item.Name -ne [string]$item.Prefix) {
                    '{0} [{1}]' -f $item.Name, $item.Prefix
                } else {
                    [string]$item.Prefix
                }
                $tooltipLines += ('{0}: {1}' -f $serverLabel, $item.Count)
            }
            $tooltipLines += ('Total active players: {0}' -f $totalPlayers)
        } else {
            $tooltipLines += 'No trusted active players are detected right now.'
        }

        return [pscustomobject]@{
            Summary      = ('PLAYERS: {0}' -f $totalPlayers)
            Tooltip      = ((@($tooltipLines) -join [Environment]::NewLine).Trim())
            Color        = $(if ($totalPlayers -gt 0) { $clrGreen } else { $clrAccentAlt })
            TotalPlayers = $totalPlayers
            Breakdown    = @($breakdown)
        }
    }

    function _GetDashboardCardOnlineCounts {
        $cardCounts = @{}

        try {
            if ($script:_DashboardCardPlayerCounts -is [hashtable] -and $script:_DashboardCardPlayerCounts.Count -gt 0) {
                foreach ($key in @($script:_DashboardCardPlayerCounts.Keys)) {
                    try {
                        $cardCounts[[string]$key] = [Math]::Max(0, [int]$script:_DashboardCardPlayerCounts[$key])
                    } catch { }
                }
                if ($cardCounts.Count -gt 0) {
                    return $cardCounts
                }
            }
            if (-not $script:_DashboardScrollPanel) { return $cardCounts }
            foreach ($card in @($script:_DashboardScrollPanel.Controls)) {
                if ($null -eq $card) { continue }
                $cardPrefix = ''
                try { $cardPrefix = [string]$card.Tag } catch { $cardPrefix = '' }
                if ([string]::IsNullOrWhiteSpace($cardPrefix)) { continue }

                $subtitleLabel = $null
                try { $subtitleLabel = ($card.Controls.Find('lblSubtitle', $true) | Select-Object -First 1) } catch { $subtitleLabel = $null }
                if (-not $subtitleLabel) { continue }

                $subtitleText = ''
                try { $subtitleText = [string]$subtitleLabel.Text } catch { $subtitleText = '' }
                if ($subtitleText -notmatch '^\s*(\d+)\s+player\(s\)\s+online\.?\s*$') { continue }

                $cardCount = 0
                try { $cardCount = [Math]::Max(0, [int]$Matches[1]) } catch { $cardCount = 0 }
                if ($cardCount -le 0) { continue }

                $cardCounts[$cardPrefix.ToUpperInvariant()] = $cardCount
            }
        } catch { }

        return $cardCounts
    }

    function _GetVisibleDashboardPlayersMetricSnapshot {
        $cardCounts = _GetDashboardCardOnlineCounts
        $totalPlayers = 0
        $breakdown = New-Object System.Collections.Generic.List[object]

        foreach ($prefixKey in @($cardCounts.Keys | Sort-Object)) {
            $count = 0
            try { $count = [Math]::Max(0, [int]$cardCounts[$prefixKey]) } catch { $count = 0 }
            if ($count -le 0) { continue }

            $label = $prefixKey
            try {
                if ($script:SharedState -and $script:SharedState.Profiles -and $script:SharedState.Profiles.ContainsKey($prefixKey)) {
                    $candidateName = [string]$script:SharedState.Profiles[$prefixKey].GameName
                    if (-not [string]::IsNullOrWhiteSpace($candidateName)) {
                        $label = $candidateName
                    }
                }
            } catch { }

            $totalPlayers += $count
            $breakdown.Add([pscustomobject]@{
                Prefix = [string]$prefixKey
                Name = [string]$label
                Count = [int]$count
            }) | Out-Null
        }

        if ($totalPlayers -le 0) { return $null }

        $tooltipLines = New-Object System.Collections.Generic.List[string]
        $tooltipLines.Add('Active players shown on the dashboard cards:') | Out-Null
        foreach ($item in @($breakdown)) {
            $serverLabel = if (-not [string]::IsNullOrWhiteSpace([string]$item.Name) -and [string]$item.Name -ne [string]$item.Prefix) {
                '{0} [{1}]' -f $item.Name, $item.Prefix
            } else {
                [string]$item.Prefix
            }
            $tooltipLines.Add(('{0}: {1}' -f $serverLabel, $item.Count)) | Out-Null
        }
        $tooltipLines.Add(('Total active players: {0}' -f $totalPlayers)) | Out-Null

        return [pscustomobject]@{
            Summary = 'PLAYERS: {0}' -f $totalPlayers
            Tooltip = ($tooltipLines -join [Environment]::NewLine).Trim()
            Color = $clrGreen
            TotalPlayers = $totalPlayers
            Breakdown = @($breakdown)
        }
    }

    function _MaybeTracePlayersMetricSnapshot {
        param(
            [hashtable]$SharedState,
            [object]$MetricSnapshot
        )

        try {
            if (-not $SharedState -or -not $MetricSnapshot) { return }
            if (-not $SharedState.LogQueue) { return }

            $debugEnabled = $false
            try { $debugEnabled = ($SharedState.Settings -and [bool]$SharedState.Settings.EnableDebugLogging) } catch { $debugEnabled = $false }

            $now = Get-Date
            $lastAt = $null
            $lastSignature = ''
            try { if ($SharedState.ContainsKey('LastPlayersMetricDebugAt')) { $lastAt = $SharedState['LastPlayersMetricDebugAt'] } } catch { $lastAt = $null }
            try { if ($SharedState.ContainsKey('LastPlayersMetricDebugSignature')) { $lastSignature = [string]$SharedState['LastPlayersMetricDebugSignature'] } } catch { $lastSignature = '' }

            $cardCounts = _GetDashboardCardOnlineCounts
            $profiles = $null
            try { if ($SharedState.ContainsKey('Profiles')) { $profiles = $SharedState.Profiles } } catch { $profiles = $null }

            $chosenByPrefix = @{}
            try {
                foreach ($item in @($MetricSnapshot.Breakdown)) {
                    if ($null -eq $item) { continue }
                    $chosenByPrefix[[string]$item.Prefix] = [int]$item.Count
                }
            } catch { }

            $segments = New-Object System.Collections.Generic.List[string]
            $anyPositiveSource = $false
            $runningProfileCount = 0
            foreach ($entry in @($profiles.GetEnumerator())) {
                $prefixKey = ''
                $profile = $null
                try { $prefixKey = ([string]$entry.Key).ToUpperInvariant() } catch { $prefixKey = '' }
                try { $profile = $entry.Value } catch { $profile = $null }
                if ([string]::IsNullOrWhiteSpace($prefixKey) -or $null -eq $profile) { continue }

                $running = $false
                try { $running = ($SharedState.RunningServers -and $SharedState.RunningServers.ContainsKey($prefixKey)) } catch { $running = $false }
                if ($running) { $runningProfileCount++ }

                $runtimeCode = ''
                $runtimeDetail = ''
                try {
                    $runtime = _GetRuntimeStateEntry -Prefix $prefixKey -SharedState $SharedState
                    if ($runtime) {
                        try { $runtimeCode = [string]$runtime.Code } catch { $runtimeCode = '' }
                        try { $runtimeDetail = [string]$runtime.Detail } catch { $runtimeDetail = '' }
                    }
                } catch { }

                $activityCount = ''
                $activityNoteCount = ''
                try {
                    if ($SharedState.PlayerActivityState -and $SharedState.PlayerActivityState.ContainsKey($prefixKey)) {
                        $activity = $SharedState.PlayerActivityState[$prefixKey]
                        if ($activity -and [bool]$activity.DetectionSupported -and [bool]$activity.DetectionAvailable) {
                            $activityCount = [string]([Math]::Max(0, [int]$activity.CurrentCount))
                        }
                        $activityNote = ''
                        try { $activityNote = [string]$activity.Note } catch { $activityNote = '' }
                        if ($activityNote -match '^\s*(\d+)\s+player\(s\)\s+online\.?\s*$') {
                            $activityNoteCount = [string][int]$Matches[1]
                        }
                    }
                } catch { }

                $latestCount = ''
                try {
                    if ($SharedState.LatestPlayerCounts -and $SharedState.LatestPlayerCounts.ContainsKey($prefixKey)) {
                        $latestCount = [string]([Math]::Max(0, [int]$SharedState.LatestPlayerCounts[$prefixKey]))
                    }
                } catch { }

                $rosterCount = ''
                try {
                    if ($SharedState.LatestPlayers -and $SharedState.LatestPlayers.ContainsKey($prefixKey)) {
                        $rosterCount = [string](@($SharedState.LatestPlayers[$prefixKey] | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }).Count)
                    }
                } catch { }

                $runtimeCount = ''
                if ($runtimeDetail -match '^\s*(\d+)\s+player\(s\)\s+online\.?\s*$') {
                    $runtimeCount = [string][int]$Matches[1]
                }

                $cardCount = ''
                if ($cardCounts.ContainsKey($prefixKey)) {
                    $cardCount = [string][int]$cardCounts[$prefixKey]
                }

                $chosenCount = ''
                if ($chosenByPrefix.ContainsKey($prefixKey)) {
                    $chosenCount = [string][int]$chosenByPrefix[$prefixKey]
                }

                foreach ($candidate in @($activityCount, $activityNoteCount, $latestCount, $rosterCount, $runtimeCount, $cardCount, $chosenCount)) {
                    $candidateValue = -1
                    try { [void][int]::TryParse([string]$candidate, [ref]$candidateValue) } catch { $candidateValue = -1 }
                    if ($candidateValue -gt 0) {
                        $anyPositiveSource = $true
                        break
                    }
                }

                if (-not $running -and [string]::IsNullOrWhiteSpace($activityCount) -and [string]::IsNullOrWhiteSpace($latestCount) -and [string]::IsNullOrWhiteSpace($rosterCount) -and [string]::IsNullOrWhiteSpace($runtimeCount) -and [string]::IsNullOrWhiteSpace($cardCount) -and [string]::IsNullOrWhiteSpace($chosenCount)) {
                    continue
                }

                $segments.Add(('{0}{{run={1};code={2};activity={3};note={4};latest={5};roster={6};runtime={7};card={8};chosen={9}}}' -f
                        $prefixKey,
                        $(if ($running) { '1' } else { '0' }),
                        $(if ([string]::IsNullOrWhiteSpace($runtimeCode)) { '-' } else { $runtimeCode }),
                        $(if ([string]::IsNullOrWhiteSpace($activityCount)) { '-' } else { $activityCount }),
                        $(if ([string]::IsNullOrWhiteSpace($activityNoteCount)) { '-' } else { $activityNoteCount }),
                        $(if ([string]::IsNullOrWhiteSpace($latestCount)) { '-' } else { $latestCount }),
                        $(if ([string]::IsNullOrWhiteSpace($rosterCount)) { '-' } else { $rosterCount }),
                        $(if ([string]::IsNullOrWhiteSpace($runtimeCount)) { '-' } else { $runtimeCount }),
                        $(if ([string]::IsNullOrWhiteSpace($cardCount)) { '-' } else { $cardCount }),
                        $(if ([string]::IsNullOrWhiteSpace($chosenCount)) { '-' } else { $chosenCount })
                    )) | Out-Null
            }

            $signature = 'total={0};{1}' -f ([int]$MetricSnapshot.TotalPlayers), ($segments -join ' | ')
            $shouldLog = $false
            $cardTotal = 0
            try {
                foreach ($value in @($cardCounts.Values)) {
                    $cardTotal += [Math]::Max(0, [int]$value)
                }
            } catch { $cardTotal = 0 }

            $hasMismatch = ($cardTotal -ne [int]$MetricSnapshot.TotalPlayers)
            $headerMissedPositiveSource = ([int]$MetricSnapshot.TotalPlayers -le 0 -and $anyPositiveSource)
            if ($debugEnabled) {
                if ([string]::IsNullOrWhiteSpace($lastSignature) -or $signature -ne $lastSignature) {
                    $shouldLog = $true
                } elseif ($lastAt -is [datetime] -and (($now - $lastAt).TotalSeconds -ge 15)) {
                    $shouldLog = $true
                } elseif ($segments.Count -gt 0 -and [int]$MetricSnapshot.TotalPlayers -le 0) {
                    $cardPositive = @($cardCounts.Values | Where-Object { [int]$_ -gt 0 }).Count -gt 0
                    if ($cardPositive -and ($null -eq $lastAt -or (($now - $lastAt).TotalSeconds -ge 5))) {
                        $shouldLog = $true
                    }
                }
            } elseif ($hasMismatch -or $headerMissedPositiveSource) {
                if ([string]::IsNullOrWhiteSpace($lastSignature) -or $signature -ne $lastSignature) {
                    $shouldLog = $true
                } elseif ($null -eq $lastAt -or ($lastAt -is [datetime] -and (($now - $lastAt).TotalSeconds -ge 5))) {
                    $shouldLog = $true
                }
            }

            if (-not $shouldLog) { return }

            $SharedState['LastPlayersMetricDebugAt'] = $now
            $SharedState['LastPlayersMetricDebugSignature'] = $signature
            $logLevel = if ($debugEnabled) { 'DEBUG' } else { 'WARN' }
            $logLine = '[{0}][{1}][GUI] PLAYERSMETRIC total={2} cardTotal={3} runningProfiles={4} positiveSource={5} {6}' -f `
                $now.ToString('yyyy-MM-dd HH:mm:ss'),
                $logLevel,
                ([int]$MetricSnapshot.TotalPlayers),
                $cardTotal,
                $runningProfileCount,
                $(if ($anyPositiveSource) { '1' } else { '0' }),
                ($segments -join ' | ')
            try {
                $SharedState.LogQueue.Enqueue($logLine)
            } catch {
                _Log -Level $logLevel -Message ('PLAYERSMETRIC total={0} cardTotal={1} runningProfiles={2} positiveSource={3} {4}' -f ([int]$MetricSnapshot.TotalPlayers), $cardTotal, $runningProfileCount, $(if ($anyPositiveSource) { '1' } else { '0' }), ($segments -join ' | '))
            }
        } catch { }
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

    function _ResetLogBoxState {
        param([System.Windows.Forms.RichTextBox]$Box)

        if ($null -eq $Box) { return }
        try {
            $boxKey = [System.Runtime.CompilerServices.RuntimeHelpers]::GetHashCode($Box)
            if ($script:_RtbLineCounts.ContainsKey($boxKey)) {
                $script:_RtbLineCounts.Remove($boxKey) | Out-Null
            }
        } catch { }
    }

    function _ExpandPathVars {
        param([string]$Path)
        if ([string]::IsNullOrWhiteSpace($Path)) { return $Path }
        return [Environment]::ExpandEnvironmentVariables($Path)
    }

    function _NormalizeGameIdentity {
        param([string]$Name)
        if ([string]::IsNullOrWhiteSpace($Name)) { return '' }
        return (($Name -replace '[^A-Za-z0-9]', '').ToLowerInvariant())
    }

function _GetProfileKnownGame {
        param([hashtable]$Profile)

        if ($null -eq $Profile) { return '' }
        if (($Profile.Keys -contains 'KnownGame') -and -not [string]::IsNullOrWhiteSpace("$($Profile.KnownGame)")) {
            return "$($Profile.KnownGame)"
        }

        $prefix = if (($Profile.Keys -contains 'Prefix') -and $Profile.Prefix) { "$($Profile.Prefix)".ToUpperInvariant() } else { '' }
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

        if (($Profile.Keys -contains 'GameName') -and $Profile.GameName) {
            return "$($Profile.GameName)"
        }

    return ''
}

function _GetSharedLatestPlayers {
    param(
        [string]$Prefix,
        [hashtable]$SharedState
    )

    if ([string]::IsNullOrWhiteSpace($Prefix) -or -not $SharedState) { return @() }
    $key = $Prefix.ToUpperInvariant()
    if ($SharedState.ContainsKey('LatestPlayers') -and $SharedState.LatestPlayers -and $SharedState.LatestPlayers.ContainsKey($key)) {
        return @($SharedState.LatestPlayers[$key] | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    }
    return @()
}

function _AddSharedObservedPlayer {
    param(
        [string]$Prefix,
        [string]$PlayerName,
        [hashtable]$SharedState
    )

    if ([string]::IsNullOrWhiteSpace($PlayerName)) { return }
    $players = @(_GetSharedLatestPlayers -Prefix $Prefix -SharedState $SharedState)
    $players += [string]$PlayerName
    $safePlayers = @(_ToStringArray -Value $players)
    Set-LatestPlayersSnapshot -Prefix $Prefix -Names @($safePlayers) -Count $safePlayers.Count -SharedState $SharedState
}

function _RemoveSharedObservedPlayer {
    param(
        [string]$Prefix,
        [string]$PlayerName,
        [hashtable]$SharedState
    )

    if ([string]::IsNullOrWhiteSpace($PlayerName)) { return }
    $target = $PlayerName.Trim().ToLowerInvariant()
    $players = @(_GetSharedLatestPlayers -Prefix $Prefix -SharedState $SharedState) |
        Where-Object { $_ -and ([string]$_).Trim().ToLowerInvariant() -ne $target }
    $safePlayers = @(_ToStringArray -Value $players)
    Set-LatestPlayersSnapshot -Prefix $Prefix -Names @($safePlayers) -Count $safePlayers.Count -SharedState $SharedState
}

function _ToStringArray {
    param([object]$Value)

    $result = New-Object System.Collections.Generic.List[string]
    foreach ($item in @($Value)) {
        if ($item -is [System.Collections.IEnumerable] -and -not ($item -is [string])) {
            foreach ($nested in $item) {
                $text = [string]$nested
                if (-not [string]::IsNullOrWhiteSpace($text)) {
                    $result.Add($text.Trim()) | Out-Null
                }
            }
        } else {
            $text = [string]$item
            if (-not [string]::IsNullOrWhiteSpace($text)) {
                $result.Add($text.Trim()) | Out-Null
            }
        }
    }

    return @($result.ToArray())
}

function _GetSharedPzObservedPlayerIds {
    param(
        [string]$Prefix,
        [hashtable]$SharedState
    )

    if ([string]::IsNullOrWhiteSpace($Prefix) -or -not $SharedState) { return $null }
    if (-not $SharedState.ContainsKey('PzObservedPlayerIds') -or -not $SharedState.PzObservedPlayerIds) {
        $SharedState['PzObservedPlayerIds'] = [hashtable]::Synchronized(@{})
    }

    $key = $Prefix.ToUpperInvariant()
    if (-not $SharedState.PzObservedPlayerIds.ContainsKey($key) -or $null -eq $SharedState.PzObservedPlayerIds[$key]) {
        $SharedState.PzObservedPlayerIds[$key] = [hashtable]::Synchronized(@{})
    }

    return $SharedState.PzObservedPlayerIds[$key]
}

function _RememberSharedPzObservedPlayer {
    param(
        [string]$Prefix,
        [string]$PlayerId,
        [string]$PlayerName,
        [hashtable]$SharedState
    )

    if ([string]::IsNullOrWhiteSpace($PlayerId) -or [string]::IsNullOrWhiteSpace($PlayerName)) { return }
    $map = _GetSharedPzObservedPlayerIds -Prefix $Prefix -SharedState $SharedState
    if ($null -eq $map) { return }
    $map[[string]$PlayerId] = [string]$PlayerName
}

function _RemoveSharedPzObservedPlayerById {
    param(
        [string]$Prefix,
        [string]$PlayerId,
        [hashtable]$SharedState
    )

    if ([string]::IsNullOrWhiteSpace($PlayerId)) { return }
    $map = _GetSharedPzObservedPlayerIds -Prefix $Prefix -SharedState $SharedState
    if ($null -eq $map) { return }

    $resolvedName = $null
    if ($map.ContainsKey([string]$PlayerId)) {
        $resolvedName = [string]$map[[string]$PlayerId]
        $map.Remove([string]$PlayerId) | Out-Null
    }

    if (-not [string]::IsNullOrWhiteSpace($resolvedName)) {
        _RemoveSharedObservedPlayer -Prefix $Prefix -PlayerName $resolvedName -SharedState $SharedState
    }
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
            $knownGame = _NormalizeGameIdentity (_GetProfileKnownGame -Profile $Profile)
            if ($knownGame -eq 'projectzomboid') {
                $roots += (Join-Path $env:USERPROFILE 'Zomboid\Server')
            }
        }

        # Ensure we never enumerate a single string as characters
        $roots = @($roots)
        $roots = $roots | Where-Object { $_ -and (Test-Path $_) } | Select-Object -Unique
        return ,$roots
    }

    function _GetConfigValueKind {
        param([string]$Value)

        $trimmed = if ($null -eq $Value) { '' } else { $Value.Trim() }
        if ($trimmed -match '^(?i:true|false|yes|no|on|off)$') { return 'bool' }
        $intVal = 0
        if ([int]::TryParse($trimmed, [ref]$intVal)) { return 'int' }
        $doubleVal = 0.0
        if ([double]::TryParse($trimmed, [ref]$doubleVal)) { return 'number' }
        if ($trimmed.Length -gt 90) { return 'multiline' }
        return 'text'
    }

    function _ConvertGeneratedJsonValueToText {
        param([object]$Value)

        if ($null -eq $Value) { return '' }
        if ($Value -is [bool]) {
            if ($Value) { return 'true' }
            return 'false'
        }
        return [string]$Value
    }

    function _AddGeneratedJsonEntries {
        param(
            [object]$Node,
            [string]$Path,
            [string]$Section,
            [System.Collections.ArrayList]$Entries,
            [ref]$EntryIndex
        )

        if ($null -eq $Node) {
            $keyName = if ([string]::IsNullOrWhiteSpace($Path)) { 'value' } else { ($Path -split '\.')[-1] }
            [void]$Entries.Add([pscustomobject][ordered]@{
                Type='entry'; Raw=''; Section=$Section; Key=$keyName; Value=''; OriginalValue=''; Kind='text';
                EntryId="entry_$($EntryIndex.Value)"; Format='json'; JsonPath=$Path; JsonValueType='null'
            })
            $EntryIndex.Value++
            return
        }

        if ($Node -is [System.Collections.IDictionary]) {
            foreach ($key in @($Node.Keys)) {
                $childPath = if ([string]::IsNullOrWhiteSpace($Path)) { [string]$key } else { "$Path.$key" }
                $childSection = if ([string]::IsNullOrWhiteSpace($Path)) { 'General' } else { $Path }
                _AddGeneratedJsonEntries -Node $Node[$key] -Path $childPath -Section $childSection -Entries $Entries -EntryIndex $EntryIndex
            }
            return
        }

        if ($Node -is [System.Management.Automation.PSCustomObject]) {
            foreach ($prop in @($Node.PSObject.Properties)) {
                $childPath = if ([string]::IsNullOrWhiteSpace($Path)) { [string]$prop.Name } else { "$Path.$($prop.Name)" }
                $childSection = if ([string]::IsNullOrWhiteSpace($Path)) { 'General' } else { $Path }
                _AddGeneratedJsonEntries -Node $prop.Value -Path $childPath -Section $childSection -Entries $Entries -EntryIndex $EntryIndex
            }
            return
        }

        if ($Node -is [System.Collections.IList] -and -not ($Node -is [string])) {
            for ($i = 0; $i -lt $Node.Count; $i++) {
                $childPath = if ([string]::IsNullOrWhiteSpace($Path)) { "[$i]" } else { "$Path[$i]" }
                $childSection = if ([string]::IsNullOrWhiteSpace($Path)) { 'General' } else { $Path }
                _AddGeneratedJsonEntries -Node $Node[$i] -Path $childPath -Section $childSection -Entries $Entries -EntryIndex $EntryIndex
            }
            return
        }

        $jsonType = 'text'
        if ($Node -is [bool]) { $jsonType = 'bool' }
        elseif ($Node -is [int] -or $Node -is [long]) { $jsonType = 'int' }
        elseif ($Node -is [double] -or $Node -is [decimal] -or $Node -is [single]) { $jsonType = 'number' }

        $textValue = _ConvertGeneratedJsonValueToText -Value $Node
        $keyName = if ([string]::IsNullOrWhiteSpace($Path)) { 'value' } else { ($Path -split '\.')[-1] -replace '\[\d+\]$','' }
        if ([string]::IsNullOrWhiteSpace($keyName)) { $keyName = $Path }
        [void]$Entries.Add([pscustomobject][ordered]@{
            Type='entry'; Raw=''; Section=$Section; Key=$keyName; Value=$textValue; OriginalValue=$textValue; Kind=$jsonType;
            EntryId="entry_$($EntryIndex.Value)"; Format='json'; JsonPath=$Path; JsonValueType=$jsonType
        })
        $EntryIndex.Value++
    }

    function _GetGeneratedJsonPathSegments {
        param([string]$Path)

        $segments = New-Object System.Collections.Generic.List[object]
        if ([string]::IsNullOrWhiteSpace($Path)) { return ,@() }

        foreach ($part in @($Path -split '\.')) {
            if ([string]::IsNullOrWhiteSpace($part)) { continue }
            $name = ($part -replace '\[\d+\]','')
            if (-not [string]::IsNullOrWhiteSpace($name)) { $segments.Add($name) | Out-Null }
            foreach ($match in [regex]::Matches($part, '\[(\d+)\]')) {
                $segments.Add([int]$match.Groups[1].Value) | Out-Null
            }
        }

        return ,@($segments)
    }

    function _SetGeneratedJsonPathValue {
        param(
            [ref]$Root,
            [object[]]$Segments,
            [object]$Value
        )

        if ($null -eq $Segments -or $Segments.Count -eq 0) {
            $Root.Value = $Value
            return
        }

        if ($null -eq $Root.Value) {
            $Root.Value = if ($Segments[0] -is [int]) { New-Object System.Collections.ArrayList } else { [ordered]@{} }
        }

        $current = $Root.Value
        for ($i = 0; $i -lt $Segments.Count; $i++) {
            $segment = $Segments[$i]
            $isLast = ($i -eq ($Segments.Count - 1))
            $nextSegment = if (-not $isLast) { $Segments[$i + 1] } else { $null }

            if ($segment -is [int]) {
                if (-not ($current -is [System.Collections.IList])) {
                    throw "JSON path segment [$segment] expected an array container."
                }

                while ($current.Count -le $segment) {
                    [void]$current.Add($null)
                }

                if ($isLast) {
                    $current[$segment] = $Value
                } else {
                    if ($null -eq $current[$segment]) {
                        $current[$segment] = if ($nextSegment -is [int]) { New-Object System.Collections.ArrayList } else { [ordered]@{} }
                    }
                    $current = $current[$segment]
                }
                continue
            }

            if ($current -isnot [System.Collections.IDictionary]) {
                throw "JSON path segment '$segment' expected an object container."
            }

            if ($isLast) {
                $current[$segment] = $Value
            } else {
                if (-not $current.Contains($segment) -or $null -eq $current[$segment]) {
                    $current[$segment] = if ($nextSegment -is [int]) { New-Object System.Collections.ArrayList } else { [ordered]@{} }
                }
                $current = $current[$segment]
            }
        }
    }

    function _GetGeneratedXmlSiblingIndex {
        param([System.Xml.XmlNode]$Node)

        if ($null -eq $Node -or $null -eq $Node.ParentNode) { return 1 }
        $index = 0
        foreach ($sibling in @($Node.ParentNode.ChildNodes)) {
            if ($sibling.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
            if ($sibling.Name -ne $Node.Name) { continue }
            $index++
            if ($sibling -eq $Node) { return $index }
        }
        return 1
    }

    function _GetGeneratedXmlNodePath {
        param([System.Xml.XmlNode]$Node)

        if ($null -eq $Node) { return '' }
        $parts = New-Object System.Collections.Generic.List[string]
        $current = $Node
        while ($current -and $current.NodeType -eq [System.Xml.XmlNodeType]::Element) {
            $parts.Insert(0, "/$($current.Name)[$(_GetGeneratedXmlSiblingIndex -Node $current)]")
            $current = $current.ParentNode
            if ($current -and $current.NodeType -eq [System.Xml.XmlNodeType]::Document) { break }
        }
        return ($parts -join '')
    }

    function _AddGeneratedXmlEntries {
        param(
            [System.Xml.XmlNode]$Node,
            [System.Collections.ArrayList]$Entries,
            [ref]$EntryIndex
        )

        if ($null -eq $Node -or $Node.NodeType -ne [System.Xml.XmlNodeType]::Element) { return }

        $nodePath = _GetGeneratedXmlNodePath -Node $Node
        $section = if ($Node.ParentNode -and $Node.ParentNode.NodeType -eq [System.Xml.XmlNodeType]::Element) { $Node.ParentNode.Name } else { 'General' }
        $nameAttr = if ($Node.Attributes) { $Node.Attributes['name'] } else { $null }
        $valueAttr = if ($Node.Attributes) { $Node.Attributes['value'] } else { $null }
        $elementChildren = @($Node.ChildNodes | Where-Object { $_.NodeType -eq [System.Xml.XmlNodeType]::Element })
        $textChildren = @($Node.ChildNodes | Where-Object {
            $_.NodeType -eq [System.Xml.XmlNodeType]::Text -or $_.NodeType -eq [System.Xml.XmlNodeType]::CDATA
        })
        $significantText = @($textChildren | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Value) })

        if ($nameAttr -and $valueAttr -and $elementChildren.Count -eq 0 -and $significantText.Count -eq 0) {
            $displayKey = [string]$nameAttr.Value
            $displayValue = [string]$valueAttr.Value
            if (-not [string]::IsNullOrWhiteSpace($displayKey)) {
                [void]$Entries.Add([pscustomobject][ordered]@{
                    Type='entry'; Raw=''; Section=$section; Key=$displayKey; Value=$displayValue; OriginalValue=$displayValue;
                    Kind=(_GetConfigValueKind -Value $displayValue); EntryId="entry_$($EntryIndex.Value)"; Format='xml';
                    XmlPath=$nodePath; XmlKind='attribute_pair'; XmlName='value'; XmlKeyName='name'; XmlKeyValue=$displayKey
                })
                $EntryIndex.Value++

                foreach ($attr in @($Node.Attributes)) {
                    if ($attr.Name -in @('name','value')) { continue }
                    $value = [string]$attr.Value
                    [void]$Entries.Add([pscustomobject][ordered]@{
                        Type='entry'; Raw=''; Section=$section; Key=("{0}.{1}" -f $displayKey, $attr.Name); Value=$value; OriginalValue=$value;
                        Kind=(_GetConfigValueKind -Value $value); EntryId="entry_$($EntryIndex.Value)"; Format='xml';
                        XmlPath=$nodePath; XmlKind='attribute'; XmlName=$attr.Name
                    })
                    $EntryIndex.Value++
                }
                return
            }
        }

        foreach ($attr in @($Node.Attributes)) {
            $value = [string]$attr.Value
            [void]$Entries.Add([pscustomobject][ordered]@{
                Type='entry'; Raw=''; Section=$section; Key="@$($attr.Name)"; Value=$value; OriginalValue=$value;
                Kind=(_GetConfigValueKind -Value $value); EntryId="entry_$($EntryIndex.Value)"; Format='xml';
                XmlPath=$nodePath; XmlKind='attribute'; XmlName=$attr.Name
            })
            $EntryIndex.Value++
        }

        if ($elementChildren.Count -gt 0 -and $significantText.Count -gt 0) {
            throw "Mixed-content XML at $nodePath needs the raw editor."
        }

        if ($elementChildren.Count -eq 0) {
            $textValue = ''
            if ($textChildren.Count -gt 0) { $textValue = [string]::Join('', @($textChildren | ForEach-Object { [string]$_.Value })) }
            [void]$Entries.Add([pscustomobject][ordered]@{
                Type='entry'; Raw=''; Section=$section; Key=$Node.Name; Value=$textValue; OriginalValue=$textValue;
                Kind=(_GetConfigValueKind -Value $textValue); EntryId="entry_$($EntryIndex.Value)"; Format='xml';
                XmlPath=$nodePath; XmlKind='text'; XmlName=$Node.Name
            })
            $EntryIndex.Value++
        } else {
            foreach ($child in $elementChildren) {
                _AddGeneratedXmlEntries -Node $child -Entries $Entries -EntryIndex $EntryIndex
            }
        }
    }

    function _ConvertGeneratedJsonControlValue {
        param(
            [string]$Value,
            [string]$JsonValueType
        )

        switch ("$JsonValueType".ToLowerInvariant()) {
            'bool' {
                $normalized = [string]$Value
                return @('true','yes','on','1') -contains $normalized.Trim().ToLowerInvariant()
            }
            'int' {
                $intVal = 0
                [void][int]::TryParse([string]$Value, [ref]$intVal)
                return $intVal
            }
            'number' {
                $numVal = 0.0
                [void][double]::TryParse([string]$Value, [ref]$numVal)
                return $numVal
            }
            'null' { return $null }
            default { return [string]$Value }
        }
    }

    function _ParseGeneratedConfig {
        param(
            [string]$Content,
            [string]$Extension
        )

        $supported = @('.ini','.cfg','.conf','.properties','.txt')
        $extLower = "$Extension".ToLowerInvariant()
        if ($extLower -eq '.json') {
            try {
                $jsonRoot = $null
                if (-not [string]::IsNullOrWhiteSpace($Content)) {
                    $jsonRoot = $Content | ConvertFrom-Json -ErrorAction Stop
                }
                $entries = New-Object System.Collections.ArrayList
                $entryIndex = 0
                _AddGeneratedJsonEntries -Node $jsonRoot -Path '' -Section 'General' -Entries $entries -EntryIndex ([ref]$entryIndex)
                $hasEditable = @($entries | Where-Object { $_.Type -eq 'entry' }).Count -gt 0
                if (-not $hasEditable) {
                    return @{ Supported = $false; Reason = 'No editable JSON settings were detected in this file.'; Lines = @(); Entries = @() }
                }
                return @{ Supported = $true; Reason = ''; Lines = @(); Entries = @($entries) }
            } catch {
                return @{ Supported = $false; Reason = "This JSON file could not be parsed for the generated editor: $($_.Exception.Message)"; Lines = @(); Entries = @() }
            }
        }
        if ($extLower -eq '.xml') {
            try {
                $xmlDoc = New-Object System.Xml.XmlDocument
                $xmlDoc.PreserveWhitespace = $true
                $xmlDoc.LoadXml([string]$Content)
                $entries = New-Object System.Collections.ArrayList
                $entryIndex = 0
                if ($xmlDoc.DocumentElement) {
                    _AddGeneratedXmlEntries -Node $xmlDoc.DocumentElement -Entries $entries -EntryIndex ([ref]$entryIndex)
                }
                $hasEditable = @($entries | Where-Object { $_.Type -eq 'entry' }).Count -gt 0
                if (-not $hasEditable) {
                    return @{ Supported = $false; Reason = 'No editable XML settings were detected in this file.'; Lines = @(); Entries = @() }
                }
                return @{ Supported = $true; Reason = ''; Lines = @(); Entries = @($entries) }
            } catch {
                return @{ Supported = $false; Reason = "This XML file needs the raw editor: $($_.Exception.Message)"; Lines = @(); Entries = @() }
            }
        }
        if ($extLower -eq '.lua') {
            $lines = @()
            if ($null -ne $Content) { $lines = $Content -split "\r?\n" }
            $entries = New-Object System.Collections.ArrayList
            $sectionStack = New-Object System.Collections.Generic.List[string]
            $entryIndex = 0
            $sawEditable = $false
            $rootName = ''

            foreach ($line in $lines) {
                if ($line -match '^\s*$') {
                    [void]$entries.Add([pscustomobject][ordered]@{ Type='blank'; Raw=$line; Section=if ($sectionStack.Count -gt 0) { $sectionStack[$sectionStack.Count - 1] } else { '' } })
                    continue
                }
                if ($line -match '^\s*--') {
                    [void]$entries.Add([pscustomobject][ordered]@{ Type='comment'; Raw=$line; Section=if ($sectionStack.Count -gt 0) { $sectionStack[$sectionStack.Count - 1] } else { '' } })
                    continue
                }
                if ($line -match '^\s*([A-Za-z0-9_]+)\s*=\s*\{\s*(--.*)?$') {
                    $name = $Matches[1]
                    $isRoot = ($sectionStack.Count -eq 0)
                    if ($isRoot) {
                        $rootName = $name
                    } else {
                        [void]$entries.Add([pscustomobject][ordered]@{ Type='section'; Raw=$line; Section=$name; Name=$name; IsLuaSection=$true })
                    }
                    $sectionStack.Add($name)
                    [void]$entries.Add([pscustomobject][ordered]@{ Type='lua_table_start'; Raw=$line; Section=$name; Name=$name; IsRoot=$isRoot })
                    continue
                }
                if ($line -match '^\s*\}\s*,?\s*(--.*)?$') {
                    $sectionName = if ($sectionStack.Count -gt 0) { $sectionStack[$sectionStack.Count - 1] } else { '' }
                    [void]$entries.Add([pscustomobject][ordered]@{ Type='lua_table_end'; Raw=$line; Section=$sectionName })
                    if ($sectionStack.Count -gt 0) { $sectionStack.RemoveAt($sectionStack.Count - 1) }
                    continue
                }
                if ($line -match '^\s*([A-Za-z0-9_]+)\s*=\s*(.+?)\s*(,?)\s*$') {
                    $key = $Matches[1]
                    $value = $Matches[2].Trim()
                    $trailing = $Matches[3]
                    if ($value -like '{*' -or $value -match '^function\b') {
                        [void]$entries.Add([pscustomobject][ordered]@{
                            Type='raw'; Raw=$line; Section=if ($sectionStack.Count -gt 0) { $sectionStack[$sectionStack.Count - 1] } else { $rootName }
                        })
                        continue
                    }
                    $kind = _GetConfigValueKind -Value ($value.Trim('"'))
                [void]$entries.Add([pscustomobject][ordered]@{
                    Type='entry'; Raw=$line; Section=if ($sectionStack.Count -gt 0) { $sectionStack[$sectionStack.Count - 1] } else { $rootName };
                    Key=$key; Value=$value; OriginalValue=$value; Kind=$kind; EntryId="entry_$entryIndex"; Format='lua'; TrailingComma=$trailing
                })
                    $entryIndex++
                    $sawEditable = $true
                    continue
                }

                [void]$entries.Add([pscustomobject][ordered]@{
                    Type='raw'; Raw=$line; Section=if ($sectionStack.Count -gt 0) { $sectionStack[$sectionStack.Count - 1] } else { $rootName }
                })
                continue
            }

            if (-not $sawEditable) {
                return @{ Supported = $false; Reason = 'No editable Lua settings were detected in this file.'; Lines = $lines; Entries = @($entries) }
            }

            return @{ Supported = $true; Reason = ''; Lines = $lines; Entries = @($entries) }
        }
        if ($supported -notcontains $extLower) {
            return @{ Supported = $false; Reason = "Generated editor is not enabled for $Extension files yet."; Lines = @(); Entries = @() }
        }

        $lines = @()
        if ($null -ne $Content) { $lines = $Content -split "\r?\n" }
        $entries = New-Object System.Collections.ArrayList
        $currentSection = 'General'
        $entryIndex = 0

        foreach ($line in $lines) {
            if ($line -match '^\s*$') {
                [void]$entries.Add([pscustomobject][ordered]@{ Type='blank'; Raw=$line; Section=$currentSection })
                continue
            }
            if ($line -match '^\s*[#;]') {
                [void]$entries.Add([pscustomobject][ordered]@{ Type='comment'; Raw=$line; Section=$currentSection })
                continue
            }
            if ($line -match '^\s*\[(.+?)\]\s*$') {
                $currentSection = $Matches[1].Trim()
                [void]$entries.Add([pscustomobject][ordered]@{ Type='section'; Raw=$line; Section=$currentSection; Name=$currentSection })
                continue
            }
            if ($line -match '^(\s*)([^=:#]+?)(\s*)(=|:)(\s*)(.*)$') {
                $key = $Matches[2].Trim()
                $value = $Matches[6]
                $kind = _GetConfigValueKind -Value $value
                [void]$entries.Add([pscustomobject][ordered]@{
                    Type='entry'; Raw=$line; Section=$currentSection; Key=$key; Value=$value; OriginalValue=$value; Kind=$kind;
                    Leading=$Matches[1]; SepLeft=$Matches[3]; Separator=$Matches[4]; SepRight=$Matches[5];
                    EntryId="entry_$entryIndex"
                })
                $entryIndex++
                continue
            }

            # Keep unfamiliar lines verbatim instead of rejecting the whole file.
            # The generated editor will still expose any recognized key/value rows,
            # while save serialization preserves passthrough lines unchanged.
            [void]$entries.Add([pscustomobject][ordered]@{
                Type='raw'; Raw=$line; Section=$currentSection
            })
            continue
        }

        $hasEditable = @($entries | Where-Object { $_.Type -eq 'entry' }).Count -gt 0
        if (-not $hasEditable) {
            return @{ Supported = $false; Reason = 'No editable key/value settings were detected in this file.'; Lines = $lines; Entries = @($entries) }
        }

        return @{ Supported = $true; Reason = ''; Lines = $lines; Entries = @($entries) }
    }

    function _SerializeGeneratedConfig {
        param(
            [object[]]$Entries,
            [hashtable]$Controls
        )

        $entryList = @($Entries)
        $structuredEntries = @($entryList | Where-Object { $_.Type -eq 'entry' })
        $firstStructured = $structuredEntries | Select-Object -First 1
        if ($firstStructured -and $firstStructured.PSObject.Properties.Name -contains 'Format') {
            $format = "$($firstStructured.Format)".ToLowerInvariant()
            if ($format -eq 'json') {
                $root = $null
                foreach ($entry in $structuredEntries) {
                    if (-not $Controls.ContainsKey($entry.EntryId)) { continue }
                    $meta = $Controls[$entry.EntryId]
                    $control = if ($meta -is [System.Collections.IDictionary] -and $meta.Contains('Control')) { $meta.Control } else { $meta }
                    $rawValue = _GetGeneratedControlValue -Entry $entry -Control $control
                    $typedValue = _ConvertGeneratedJsonControlValue -Value $rawValue -JsonValueType ([string]$entry.JsonValueType)
                    $segments = _GetGeneratedJsonPathSegments -Path ([string]$entry.JsonPath)
                    _SetGeneratedJsonPathValue -Root ([ref]$root) -Segments $segments -Value $typedValue
                }
                return ($root | ConvertTo-Json -Depth 64)
            }
            if ($format -eq 'xml') {
                $xmlDoc = New-Object System.Xml.XmlDocument
                $xmlDoc.PreserveWhitespace = $true
                $rootPath = ($firstStructured.XmlPath -split '/')[1]
                $rootName = if ($rootPath -match '^([^\[]+)') { $Matches[1] } else { 'root' }
                [void]$xmlDoc.AppendChild($xmlDoc.CreateXmlDeclaration('1.0','utf-8',$null))
                $rootNode = $xmlDoc.CreateElement($rootName)
                [void]$xmlDoc.AppendChild($rootNode)

                foreach ($entry in $structuredEntries) {
                    if (-not $Controls.ContainsKey($entry.EntryId)) { continue }
                    $meta = $Controls[$entry.EntryId]
                    $control = if ($meta -is [System.Collections.IDictionary] -and $meta.Contains('Control')) { $meta.Control } else { $meta }
                    $currentValue = _GetGeneratedControlValue -Entry $entry -Control $control
                    $xmlPath = [string]$entry.XmlPath
                    $pathParts = @($xmlPath -split '/' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
                    $currentNode = $xmlDoc.DocumentElement
                    for ($i = 1; $i -lt $pathParts.Count; $i++) {
                        $part = $pathParts[$i]
                        $childName = if ($part -match '^([^\[]+)') { $Matches[1] } else { $part }
                        $childIndex = 1
                        if ($part -match '\[(\d+)\]') { $childIndex = [int]$Matches[1] }
                        $childNode = $null
                        $matchingChildren = @()
                        foreach ($candidate in @($currentNode.ChildNodes)) {
                            if ($candidate.NodeType -eq [System.Xml.XmlNodeType]::Element -and $candidate.Name -eq $childName) {
                                $matchingChildren += $candidate
                            }
                        }
                        while ($matchingChildren.Count -lt $childIndex) {
                            $newChild = $xmlDoc.CreateElement($childName)
                            [void]$currentNode.AppendChild($newChild)
                            $matchingChildren += $newChild
                        }
                        $childNode = $matchingChildren[$childIndex - 1]
                        $currentNode = $childNode
                    }

                    switch ("$($entry.XmlKind)".ToLowerInvariant()) {
                        'attribute' {
                            $attr = $currentNode.Attributes[[string]$entry.XmlName]
                            if ($null -eq $attr) {
                                $attr = $xmlDoc.CreateAttribute([string]$entry.XmlName)
                                [void]$currentNode.Attributes.Append($attr)
                            }
                            $attr.Value = [string]$currentValue
                        }
                        'attribute_pair' {
                            $keyAttrName = if ($entry.PSObject.Properties.Name -contains 'XmlKeyName' -and -not [string]::IsNullOrWhiteSpace([string]$entry.XmlKeyName)) { [string]$entry.XmlKeyName } else { 'name' }
                            $keyAttr = $currentNode.Attributes[$keyAttrName]
                            if ($null -eq $keyAttr) {
                                $keyAttr = $xmlDoc.CreateAttribute($keyAttrName)
                                [void]$currentNode.Attributes.Append($keyAttr)
                            }
                            $keyAttr.Value = if ($entry.PSObject.Properties.Name -contains 'XmlKeyValue') { [string]$entry.XmlKeyValue } else { [string]$entry.Key }

                            $valueAttrName = if ($entry.PSObject.Properties.Name -contains 'XmlName' -and -not [string]::IsNullOrWhiteSpace([string]$entry.XmlName)) { [string]$entry.XmlName } else { 'value' }
                            $valueAttr = $currentNode.Attributes[$valueAttrName]
                            if ($null -eq $valueAttr) {
                                $valueAttr = $xmlDoc.CreateAttribute($valueAttrName)
                                [void]$currentNode.Attributes.Append($valueAttr)
                            }
                            $valueAttr.Value = [string]$currentValue
                        }
                        default {
                            $currentNode.InnerText = [string]$currentValue
                        }
                    }
                }

                $sw = New-Object System.IO.StringWriter
                $xwSettings = New-Object System.Xml.XmlWriterSettings
                $xwSettings.Indent = $true
                $xwSettings.OmitXmlDeclaration = $false
                $xw = [System.Xml.XmlWriter]::Create($sw, $xwSettings)
                $xmlDoc.Save($xw)
                $xw.Flush()
                $xw.Close()
                return $sw.ToString()
            }
        }

        $outLines = New-Object System.Collections.Generic.List[string]
        foreach ($entry in $entryList) {
            if ($entry.Type -eq 'section' -and
                $entry.PSObject.Properties.Name -contains 'IsLuaSection' -and
                $entry.IsLuaSection) {
                continue
            }

            if ($entry.Type -ne 'entry') {
                $outLines.Add([string]$entry.Raw) | Out-Null
                continue
            }

            $value = [string]$entry.Value

            if ($entry.PSObject.Properties.Name -contains 'Format' -and $entry.Format -eq 'lua') {
                $outLines.Add("    $($entry.Key) = $value$($entry.TrailingComma)") | Out-Null
            } else {
                $outLines.Add("$($entry.Leading)$($entry.Key)$($entry.SepLeft)$($entry.Separator)$($entry.SepRight)$value") | Out-Null
            }
        }

        return ($outLines -join [Environment]::NewLine)
    }

    function _ShowGeneratedConfigPlaceholder {
        param(
            [System.Windows.Forms.Panel]$Host,
            [string]$Title,
            [string]$Message
        )

        if ($null -eq $Host) { return }
        $Host.Controls.Clear()

        $msgCard = _Panel 10 10 ($Host.Width - 20) 84 $clrPanelSoft
        $msgCard.Anchor = 'Top,Left,Right'
        $msgCard.BorderStyle = 'FixedSingle'
        $Host.Controls.Add($msgCard)

        $msgTitle = _Label $Title 14 14 320 22 $fontBold
        $msgTitle.ForeColor = $clrAccentAlt
        $msgCard.Controls.Add($msgTitle)

        $msgText = _Label $Message 14 42 ($msgCard.Width - 28) 24 $fontLabel
        $msgText.ForeColor = $clrTextSoft
        $msgText.Anchor = 'Top,Left,Right'
        $msgCard.Controls.Add($msgText)
    }

    function _GetGeneratedControlValue {
        param(
            [object]$Entry,
            $Control
        )

        if ($null -eq $Entry) { return '' }
        if ($null -eq $Control) { return [string]$Entry.Value }

        if ($Control -is [System.Windows.Forms.CheckBox]) {
            $rawCurrent = [string]$Entry.Value
            $boolTrue  = @('true','yes','on','1')
            $boolFalse = @('false','no','off','0')
            $lowerRaw = $rawCurrent.Trim().ToLowerInvariant()
            if ($boolTrue -contains $lowerRaw -or $boolFalse -contains $lowerRaw) {
                if ($Control.Checked) {
                    if ($lowerRaw -in @('yes','no')) { return 'yes' }
                    if ($lowerRaw -in @('on','off')) { return 'on' }
                    if ($lowerRaw -in @('1','0')) { return '1' }
                    return 'true'
                }
                if ($lowerRaw -in @('yes','no')) { return 'no' }
                if ($lowerRaw -in @('on','off')) { return 'off' }
                if ($lowerRaw -in @('1','0')) { return '0' }
                return 'false'
            }
            if ($Control.Checked) { return 'true' }
            return 'false'
        }
        if ($Control -is [System.Windows.Forms.TextBox]) { return [string]$Control.Text }
        if ($Control -is [System.Windows.Forms.ComboBox]) {
            if ($Control.SelectedItem) { return "$($Control.SelectedItem)" }
            return "$($Control.Text)"
        }
        return [string]$Entry.Value
    }

    function _UpdateGeneratedFieldVisual {
        param([hashtable]$Meta)

        if (-not $Meta -or -not $Meta.Entry -or -not $Meta.Control) { return }
        $current = _GetGeneratedControlValue -Entry $Meta.Entry -Control $Meta.Control
        $original = if ($Meta.Entry.PSObject.Properties.Name -contains 'OriginalValue') { [string]$Meta.Entry.OriginalValue } else { [string]$Meta.Entry.Value }
        $changed = ($current -ne $original)

        if ($Meta.Label -is [System.Windows.Forms.Label]) {
            $Meta.Label.ForeColor = if ($changed) { $clrAccentAlt } else { $clrTextSoft }
        }
        if ($Meta.Control -is [System.Windows.Forms.TextBox]) {
            $Meta.Control.BackColor = if ($changed) { [System.Drawing.Color]::FromArgb(40, 56, 82) } else { $clrPanelSoft }
        } elseif ($Meta.Control -is [System.Windows.Forms.CheckBox]) {
            $Meta.Control.ForeColor = if ($changed) { $clrAccentAlt } else { $clrText }
        }
    }

    function _ValidateGeneratedConfig {
        param([System.Collections.IDictionary]$State)

        $messages = New-Object System.Collections.Generic.List[string]
        if (-not $State -or -not $State.StructuredModel -or -not $State.StructuredControls) { return @() }

        foreach ($entry in @($State.StructuredModel.Entries)) {
            if ($entry.Type -ne 'entry') { continue }
            if (-not $State.StructuredControls.ContainsKey($entry.EntryId)) { continue }
            $meta = $State.StructuredControls[$entry.EntryId]
            $control = if ($meta -is [System.Collections.IDictionary] -and $meta.Contains('Control')) { $meta.Control } else { $meta }
            $value = _GetGeneratedControlValue -Entry $entry -Control $control

            if ($entry.Kind -eq 'int') {
                $tmp = 0
                if (-not [int]::TryParse($value.Trim('"'), [ref]$tmp)) {
                    $messages.Add("$($entry.Key): expected an integer value.") | Out-Null
                }
            } elseif ($entry.Kind -eq 'number') {
                $tmp = 0.0
                if (-not [double]::TryParse($value.Trim('"'), [ref]$tmp)) {
                    $messages.Add("$($entry.Key): expected a numeric value.") | Out-Null
                }
            }
        }

        return @($messages)
    }

    function _BuildGeneratedConfigEditor {
        param(
            [System.Windows.Forms.Panel]$Host,
            [System.Collections.IDictionary]$State,
            [string]$Content,
            [string]$Extension
        )

        if ($null -eq $Host) { return }
        $Host.Controls.Clear()

        $reuseExisting = $false
        if ($State.StructuredModel -and $State.StructuredModel.Supported -and
            $State.Contains('StructuredSource') -and $State.Contains('GeneratedExtension') -and
            $State.StructuredSource -eq $Content -and $State.GeneratedExtension -eq $Extension) {
            $reuseExisting = $true
        }

        $parse = if ($reuseExisting) { $State.StructuredModel } else { _ParseGeneratedConfig -Content $Content -Extension $Extension }
        $State['StructuredModel'] = $parse
        $State['StructuredControls'] = @{}
        $State['_GeneratedBuilder'] = ${function:_BuildGeneratedConfigEditor}
        $State['GeneratedBuilt'] = $true
        if (-not $State.Contains('GeneratedFilter')) { $State['GeneratedFilter'] = '' }
        $State['StructuredSource'] = $Content
        $State['GeneratedExtension'] = $Extension

        if (-not $parse.Supported) {
            _ShowGeneratedConfigPlaceholder -Host $Host -Title 'Generated Editor Unavailable' -Message $parse.Reason
            return
        }

        $scroll = New-Object System.Windows.Forms.Panel
        $scroll.Dock = 'Fill'
        $scroll.AutoScroll = $true
        $scroll.BackColor = $clrPanel
        $Host.Controls.Add($scroll)
        $updateFieldVisualLocal = ${function:_UpdateGeneratedFieldVisual}
        $getGeneratedControlValueLocal = ${function:_GetGeneratedControlValue}

        $header = _Panel 12 12 ($scroll.ClientSize.Width - 24) 100 $clrPanelSoft
        $header.Anchor = 'Top,Left,Right'
        $header.BorderStyle = 'FixedSingle'
        $scroll.Controls.Add($header)
        $headAccent = _Panel 0 0 4 100 $clrAccent
        $headAccent.BorderStyle = 'None'
        $header.Controls.Add($headAccent)
        $headTitle = _Label 'Generated Config Editor' 14 12 280 22 $fontTitle
        $header.Controls.Add($headTitle)
        $headInfo = _Label 'Edit detected settings with checkboxes and fields. Raw remains available for advanced edits.' 14 40 ($header.Width - 28) 18 $fontLabel
        $headInfo.ForeColor = $clrTextSoft
        $headInfo.Anchor = 'Top,Left,Right'
        $header.Controls.Add($headInfo)

        $filterLbl = _Label 'Filter Settings' 14 68 120 18 $fontBold
        $filterLbl.ForeColor = $clrTextSoft
        $filterLbl.Anchor = 'Top,Left'
        $header.Controls.Add($filterLbl)

        $searchBtnW = 78
        $clearBtnW = 70
        $tbFilter = _TextBox 136 64 ([Math]::Max(120, $header.Width - 174 - $searchBtnW - $clearBtnW)) 24 ([string]$State.GeneratedFilter) $false
        $tbFilter.Anchor = 'Top,Left,Right'
        $header.Controls.Add($tbFilter)
        $btnClearFilter = _Button 'Clear' ($header.Width - 18 - $searchBtnW - $clearBtnW) 64 $clearBtnW 24 $clrPanel $null
        $btnClearFilter.Anchor = 'Top,Right'
        $header.Controls.Add($btnClearFilter)
        _SetMainControlToolTip -Control $btnClearFilter -Text 'Clear the current config filter text.'
        $btnApplyFilter = _Button 'Search' ($header.Width - 14 - $searchBtnW) 64 $searchBtnW 24 $clrAccent $null
        $btnApplyFilter.Anchor = 'Top,Right'
        $header.Controls.Add($btnApplyFilter)

        $y = 126
        $labelW = 220
        $fieldW = [Math]::Max(280, $scroll.ClientSize.Width - 280)
        $currentSection = ''
        $filterText = [string]$State.GeneratedFilter
        $filterActive = -not [string]::IsNullOrWhiteSpace($filterText)
        $filterNeedle = $filterText.Trim().ToLowerInvariant()
        $matchingSections = @{}
        if ($filterActive) {
            foreach ($entry in @($parse.Entries)) {
                if ($entry.Type -ne 'entry') { continue }
                $entryText = ("{0} {1}" -f [string]$entry.Key, [string]$entry.Section).ToLowerInvariant()
                if ($entryText.Contains($filterNeedle)) {
                    $matchingSections[[string]$entry.Section] = $true
                }
            }
        }

        foreach ($entry in @($parse.Entries)) {
            if ($entry.Type -eq 'section') {
                if ($filterActive -and -not $matchingSections.ContainsKey([string]$entry.Name)) { continue }
                $currentSection = [string]$entry.Name
                $sec = _Label $currentSection 16 $y 420 22 $fontTitle
                $sec.ForeColor = $clrAccentAlt
                $scroll.Controls.Add($sec)
                $y += 24
                $sep = _Panel 16 $y ($scroll.ClientSize.Width - 32) 2 $clrBorder
                $sep.Anchor = 'Top,Left,Right'
                $sep.BorderStyle = 'None'
                $scroll.Controls.Add($sep)
                $y += 12
                continue
            }
            if ($entry.Type -ne 'entry') { continue }
            if ($filterActive) {
                $entryText = ("{0} {1}" -f [string]$entry.Key, [string]$entry.Section).ToLowerInvariant()
                if (-not $entryText.Contains($filterNeedle)) { continue }
            }
            if ([string]::IsNullOrWhiteSpace($currentSection) -and $y -eq 98) {
                $sec = _Label 'General' 16 $y 420 22 $fontTitle
                $sec.ForeColor = $clrAccentAlt
                $scroll.Controls.Add($sec)
                $y += 24
                $sep = _Panel 16 $y ($scroll.ClientSize.Width - 32) 2 $clrBorder
                $sep.Anchor = 'Top,Left,Right'
                $sep.BorderStyle = 'None'
                $scroll.Controls.Add($sep)
                $y += 12
                $currentSection = 'General'
            }

            $lbl = _Label ([string]$entry.Key) 18 $y $labelW 20 $fontBold
            $lbl.ForeColor = $clrTextSoft
            $scroll.Controls.Add($lbl)

            $valueCtl = $null
            switch ($entry.Kind) {
                'bool' {
                    $chk = New-Object System.Windows.Forms.CheckBox
                    $chk.Location = [System.Drawing.Point]::new(250, $y)
                    $chk.Size = [System.Drawing.Size]::new(160, 20)
                    $chk.Text = 'Enabled'
                    $chk.ForeColor = $clrText
                    $chk.BackColor = [System.Drawing.Color]::Transparent
                    $chk.Checked = @('true','yes','on','1') -contains ([string]$entry.Value).Trim().ToLowerInvariant()
                    $valueCtl = $chk
                }
                'multiline' {
                    $tb = New-Object System.Windows.Forms.TextBox
                    $tb.Location = [System.Drawing.Point]::new(250, $y)
                    $tb.Size = [System.Drawing.Size]::new($fieldW, 54)
                    $tb.Multiline = $true
                    $tb.ScrollBars = 'Vertical'
                    $tb.WordWrap = $false
                    $tb.Text = [string]$entry.Value
                    $tb.BackColor = $clrPanelSoft
                    $tb.ForeColor = $clrText
                    $tb.BorderStyle = 'FixedSingle'
                    $tb.Font = $fontMono
                    $tb.Anchor = 'Top,Left,Right'
                    $valueCtl = $tb
                }
                default {
                    $tb = _TextBox 250 $y $fieldW 24 ([string]$entry.Value) $false
                    $tb.Anchor = 'Top,Left,Right'
                    $valueCtl = $tb
                }
            }

            if ($valueCtl) {
                $scroll.Controls.Add($valueCtl)
                $meta = @{
                    Control = $valueCtl
                    Label   = $lbl
                    Entry   = $entry
                }
                $State.StructuredControls[$entry.EntryId] = $meta
                if ($valueCtl -is [System.Windows.Forms.TextBox]) {
                    $metaLocal = $meta
                    $valueCtl.Add_TextChanged({
                        $metaLocal.Entry.Value = [string]$this.Text
                        & $updateFieldVisualLocal -Meta $metaLocal
                    }.GetNewClosure())
                } elseif ($valueCtl -is [System.Windows.Forms.CheckBox]) {
                    $metaLocal = $meta
                    $valueCtl.Add_CheckedChanged({
                        $metaLocal.Entry.Value = & $getGeneratedControlValueLocal -Entry $metaLocal.Entry -Control $this
                        & $updateFieldVisualLocal -Meta $metaLocal
                    }.GetNewClosure())
                } elseif ($valueCtl -is [System.Windows.Forms.ComboBox]) {
                    $metaLocal = $meta
                    $valueCtl.Add_SelectedIndexChanged({
                        $metaLocal.Entry.Value = & $getGeneratedControlValueLocal -Entry $metaLocal.Entry -Control $this
                        & $updateFieldVisualLocal -Meta $metaLocal
                    }.GetNewClosure())
                }
                & $updateFieldVisualLocal -Meta $meta
            }
            $y += if ($entry.Kind -eq 'multiline') { 64 } else { 34 }
        }

        $hostLocal = $Host
        $stateLocal = $State
        $contentLocal = $Content
        $extensionLocal = $Extension
        $applyGeneratedFilter = {
            if ($stateLocal.Contains('GeneratedFilterTimer') -and $stateLocal.GeneratedFilterTimer) {
                $stateLocal.GeneratedFilterTimer.Stop()
            }
            if ($stateLocal.Contains('_GeneratedBuilder') -and $stateLocal['_GeneratedBuilder']) {
                & $stateLocal['_GeneratedBuilder'] -Host $hostLocal -State $stateLocal -Content $contentLocal -Extension $extensionLocal
            }
        }.GetNewClosure()
        $tbFilter.Add_TextChanged({
            $stateLocal['GeneratedFilter'] = $this.Text
            $stateLocal['FocusGeneratedFilter'] = $true
            $stateLocal['GeneratedFilterSelectionStart'] = $this.SelectionStart
            if ($stateLocal.Contains('GeneratedFilterTimer') -and $stateLocal.GeneratedFilterTimer) {
                $stateLocal.GeneratedFilterTimer.Stop()
                $stateLocal.GeneratedFilterTimer.Start()
            } elseif ($stateLocal.Contains('_GeneratedBuilder') -and $stateLocal['_GeneratedBuilder']) {
                & $stateLocal['_GeneratedBuilder'] -Host $hostLocal -State $stateLocal -Content $contentLocal -Extension $extensionLocal
            }
        }.GetNewClosure())
        $tbFilter.Add_KeyDown({
            param($sender, $e)
            if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
                $stateLocal['GeneratedFilter'] = $this.Text
                $stateLocal['FocusGeneratedFilter'] = $true
                $stateLocal['GeneratedFilterSelectionStart'] = $this.SelectionStart
                $e.SuppressKeyPress = $true
                & $applyGeneratedFilter
            }
        }.GetNewClosure())
        $btnApplyFilter.Add_Click({
            $stateLocal['GeneratedFilter'] = $tbFilter.Text
            $stateLocal['FocusGeneratedFilter'] = $true
            $stateLocal['GeneratedFilterSelectionStart'] = $tbFilter.SelectionStart
            & $applyGeneratedFilter
        }.GetNewClosure())
        $btnClearFilter.Add_Click({
            $tbFilter.Text = ''
            $stateLocal['GeneratedFilter'] = ''
            $stateLocal['FocusGeneratedFilter'] = $true
            $stateLocal['GeneratedFilterSelectionStart'] = 0
            & $applyGeneratedFilter
        }.GetNewClosure())

        if ($State.Contains('FocusGeneratedFilter') -and $State.FocusGeneratedFilter) {
            try {
                $selStart = 0
                if ($State.Contains('GeneratedFilterSelectionStart')) {
                    $selStart = [Math]::Max(0, [Math]::Min([int]$State.GeneratedFilterSelectionStart, $tbFilter.TextLength))
                }
                $tbFilter.Focus() | Out-Null
                $tbFilter.SelectionStart = $selStart
                $tbFilter.SelectionLength = 0
            } catch {}
            $State['FocusGeneratedFilter'] = $false
        }
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
        $form.Size            = [System.Drawing.Size]::new(960, 680)
        $form.MinimumSize     = [System.Drawing.Size]::new(820, 560)
        $form.StartPosition   = 'CenterParent'
        $form.BackColor       = $clrBg
        $form.FormBorderStyle = 'Sizable'

        $headerCard = _Panel 10 10 920 72 $clrPanel
        $headerCard.Anchor = 'Top,Left,Right'
        $headerCard.BorderStyle = 'FixedSingle'
        $form.Controls.Add($headerCard)

        $initialRoot = if (@($roots).Count -gt 0) { [string]$roots[0] } else { '' }
        $headerAccent = _Panel 0 0 4 72 $clrAccentAlt
        $headerAccent.Anchor = 'Top,Left,Bottom'
        $headerAccent.BorderStyle = 'None'
        $headerCard.Controls.Add($headerAccent)

        $header = _Label "Config Editor - $($Profile.GameName)" 14 10 520 22 $fontBold
        $header.Anchor = 'Top,Left,Right'
        $headerCard.Controls.Add($header)

        $lblRoot = _Label "Root: $initialRoot" 14 38 880 20 $fontLabel
        $lblRoot.Anchor = 'Top,Left,Right'
        $headerCard.Controls.Add($lblRoot)

        $combo = $null
        if ($roots.Count -gt 1) {
            $combo = New-Object System.Windows.Forms.ComboBox
            $combo.Location = [System.Drawing.Point]::new(66, 34)
            $combo.Size     = [System.Drawing.Size]::new(820, 24)
            $combo.Anchor   = 'Top,Left,Right'
            $combo.DropDownStyle = 'DropDownList'
            foreach ($r in $roots) { [void]$combo.Items.Add($r) }
            $combo.SelectedIndex = 0
            $headerCard.Controls.Add($combo)
            $lblRoot.Visible = $false
        }

        $leftHost = _Panel 10 92 250 540 $clrPanel
        $leftHost.Anchor = 'Top,Left,Bottom'
        $leftHost.BorderStyle = 'FixedSingle'
        $form.Controls.Add($leftHost)

        $lblFiles = _Label 'Config Files' 12 10 180 20 $fontBold
        $leftHost.Controls.Add($lblFiles)

        $list = New-Object System.Windows.Forms.ListBox
        $list.Location  = [System.Drawing.Point]::new(12, 34)
        $list.Size      = [System.Drawing.Size]::new(224, 492)
        $list.Anchor    = 'Top,Left,Right,Bottom'
        $list.Font      = $fontMono
        $list.BackColor = [System.Drawing.Color]::FromArgb(30,30,40)
        $list.ForeColor = $clrText
        $leftHost.Controls.Add($list)

        $cfgTabAccent = $clrAccent
        $cfgTabPanel = $clrPanelSoft
        $cfgTabText = $clrText
        $cfgTabFont = $tabFont

        $rightHost = _Panel 270 92 660 540 $clrPanel
        $rightHost.Anchor = 'Top,Left,Right,Bottom'
        $rightHost.BorderStyle = 'FixedSingle'
        $form.Controls.Add($rightHost)

        $lblEditor = _Label 'File Editor' 12 10 220 20 $fontBold
        $rightHost.Controls.Add($lblEditor)

        $editorTabs = New-Object System.Windows.Forms.TabControl
        $editorTabs.Location = [System.Drawing.Point]::new(12, 34)
        $editorTabs.Size = [System.Drawing.Size]::new(632, 446)
        $editorTabs.Anchor = 'Top,Left,Right,Bottom'
        $editorTabs.DrawMode = [System.Windows.Forms.TabDrawMode]::OwnerDrawFixed
        $editorTabs.SizeMode = [System.Windows.Forms.TabSizeMode]::Fixed
        $editorTabs.ItemSize = [System.Drawing.Size]::new(110, 26)
        $editorTabs.BackColor = $clrPanelSoft
        $editorTabs.Add_DrawItem({
            param($s,$e)
            $tab = $editorTabs.TabPages[$e.Index]
            $brush = if ($e.Index -eq $editorTabs.SelectedIndex) { [System.Drawing.SolidBrush]::new($cfgTabAccent) } else { [System.Drawing.SolidBrush]::new($cfgTabPanel) }
            $e.Graphics.FillRectangle($brush, $e.Bounds)
            $txtBrush = [System.Drawing.SolidBrush]::new($cfgTabText)
            $fmt = New-Object System.Drawing.StringFormat
            $fmt.Alignment = 'Center'
            $fmt.LineAlignment = 'Center'
            $fmt.FormatFlags = [System.Drawing.StringFormatFlags]::NoWrap
            $e.Graphics.DrawString($tab.Text, $cfgTabFont, $txtBrush, [System.Drawing.RectangleF]::new($e.Bounds.X, $e.Bounds.Y, $e.Bounds.Width, $e.Bounds.Height), $fmt)
            $brush.Dispose(); $txtBrush.Dispose(); $fmt.Dispose()
        }.GetNewClosure())
        $rightHost.Controls.Add($editorTabs)

        $tabGenerated = New-Object System.Windows.Forms.TabPage
        $tabGenerated.Text = 'Generated'
        $tabGenerated.BackColor = $clrPanel
        $editorTabs.TabPages.Add($tabGenerated) | Out-Null

        $generatedHost = New-Object System.Windows.Forms.Panel
        $generatedHost.Dock = 'Fill'
        $generatedHost.BackColor = $clrPanel
        $tabGenerated.Controls.Add($generatedHost)

        $tabRaw = New-Object System.Windows.Forms.TabPage
        $tabRaw.Text = 'Raw'
        $tabRaw.BackColor = $clrPanel
        $editorTabs.TabPages.Add($tabRaw) | Out-Null

        $editor = New-Object System.Windows.Forms.TextBox
        $editor.Dock      = 'Fill'
        $editor.Multiline = $true
        $editor.ScrollBars = 'Both'
        $editor.WordWrap  = $false
        $editor.AcceptsTab = $true
        $editor.Font      = $fontMono
        $editor.BackColor = [System.Drawing.Color]::FromArgb(12,14,24)
        $editor.ForeColor = $clrText
        $tabRaw.Controls.Add($editor)

        $cfgFooter = New-Object System.Windows.Forms.Panel
        $cfgFooter.Location = [System.Drawing.Point]::new(12, 490)
        $cfgFooter.Size = [System.Drawing.Size]::new(632, 38)
        $cfgFooter.Anchor = 'Left,Right,Bottom'
        $cfgFooter.BackColor = [System.Drawing.Color]::Transparent
        $rightHost.Controls.Add($cfgFooter)

        $cfgFooterActions = New-Object System.Windows.Forms.FlowLayoutPanel
        $cfgFooterActions.Dock = 'Left'
        $cfgFooterActions.Size = [System.Drawing.Size]::new(230, 38)
        $cfgFooterActions.WrapContents = $false
        $cfgFooterActions.AutoScroll = $true
        $cfgFooterActions.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
        $cfgFooterActions.BackColor = [System.Drawing.Color]::Transparent
        $cfgFooter.Controls.Add($cfgFooterActions)

        $cfgFooterStatusHost = New-Object System.Windows.Forms.Panel
        $cfgFooterStatusHost.Dock = 'Fill'
        $cfgFooterStatusHost.BackColor = [System.Drawing.Color]::Transparent
        $cfgFooter.Controls.Add($cfgFooterStatusHost)

        $btnSave = _Button 'Save File' 0 0 108 32 $clrGreen $null
        $btnSave.Margin = [System.Windows.Forms.Padding]::new(0, 3, 10, 0)
        $cfgFooterActions.Controls.Add($btnSave)

        $btnRefresh = _Button 'Refresh' 0 0 108 32 $clrPanel $null
        $btnRefresh.Margin = [System.Windows.Forms.Padding]::new(0, 3, 0, 0)
        $cfgFooterActions.Controls.Add($btnRefresh)

        $lblStatus = _Label '' 0 9 ($cfgFooterStatusHost.ClientSize.Width - 6) 20 $fontLabel
        $lblStatus.Anchor = 'Left,Right,Top'
        $cfgFooterStatusHost.Controls.Add($lblStatus)

        $layoutConfigEditor = {
            $clientWidth = [Math]::Max(420, $form.ClientSize.Width)
            $clientHeight = [Math]::Max(320, $form.ClientSize.Height)
            $leftMargin = 10
            $rightMargin = 10
            $topContent = 92
            $bottomMargin = 12
            $gap = 10
            $headerHeight = 72
            $footerHeight = 38
            $contentHeight = [Math]::Max(220, $clientHeight - $topContent - $bottomMargin)
            $usableWidth = $clientWidth - $leftMargin - $rightMargin

            $listWidth = [Math]::Min(250, [Math]::Max(180, [Math]::Floor(($usableWidth - $gap) * 0.34)))
            $editorWidth = [Math]::Max(320, $usableWidth - $listWidth - $gap)
            if (($listWidth + $gap + $editorWidth) -gt $usableWidth) {
                $editorWidth = [Math]::Max(280, $usableWidth - $listWidth - $gap)
            }
            if (($listWidth + $gap + $editorWidth) -gt $usableWidth) {
                $listWidth = [Math]::Max(160, $usableWidth - $gap - $editorWidth)
            }

            $headerCard.Size = [System.Drawing.Size]::new($usableWidth, $headerHeight)

            $leftHost.Location = [System.Drawing.Point]::new($leftMargin, $topContent)
            $leftHost.Size = [System.Drawing.Size]::new($listWidth, $contentHeight)
            $list.Size = [System.Drawing.Size]::new(
                [Math]::Max(120, $leftHost.ClientSize.Width - 24),
                [Math]::Max(120, $leftHost.ClientSize.Height - 46)
            )

            $rightHost.Location = [System.Drawing.Point]::new($leftMargin + $listWidth + $gap, $topContent)
            $rightHost.Size = [System.Drawing.Size]::new($editorWidth, $contentHeight)

            $editorTabs.Size = [System.Drawing.Size]::new(
                [Math]::Max(260, $rightHost.ClientSize.Width - 24),
                [Math]::Max(140, $rightHost.ClientSize.Height - 84)
            )

            $cfgFooter.Location = [System.Drawing.Point]::new(12, [Math]::Max(0, $rightHost.ClientSize.Height - $footerHeight - 12))
            $cfgFooter.Size = [System.Drawing.Size]::new([Math]::Max(240, $rightHost.ClientSize.Width - 24), $footerHeight)

            if ($combo) {
                $combo.Width = [Math]::Max(260, $headerCard.ClientSize.Width - 80)
            } else {
                $lblRoot.Width = [Math]::Max(220, $headerCard.ClientSize.Width - 28)
            }
            $header.Width = [Math]::Max(240, $headerCard.ClientSize.Width - 28)
        }.GetNewClosure()

        $allowedExt = @('.ini','.txt','.cfg','.json','.xml','.yml','.yaml','.properties','.conf','.lua')

        $state = [ordered]@{
            Root        = $initialRoot
            Roots       = $roots
            AllowedExt  = $allowedExt
            CurrentFile = ''
            List        = $list
            Editor      = $editor
            EditorTabs  = $editorTabs
            GeneratedHost = $generatedHost
            GeneratedTab  = $tabGenerated
            RawTab        = $tabRaw
            Status      = $lblStatus
            RootLabel   = $lblRoot
            StructuredModel = $null
            StructuredControls = @{}
            GeneratedBuilt = $false
            StructuredSource = ''
            GeneratedExtension = ''
        }
        $state['_GeneratedBuilder'] = ${function:_BuildGeneratedConfigEditor}

        $generatedFilterTimer = [System.Windows.Forms.Timer]::new()
        $generatedFilterTimer.Interval = 250
        $state['GeneratedFilterTimer'] = $generatedFilterTimer
        $generatedFilterTimer.Add_Tick({
            $generatedFilterTimer.Stop()
            $st = $form.Tag
            if ($st -and $st.Contains('_GeneratedBuilder') -and $st['_GeneratedBuilder'] -and
                $st.GeneratedBuilt -and $st.GeneratedHost) {
                & $st['_GeneratedBuilder'] -Host $st.GeneratedHost -State $st -Content $st.StructuredSource -Extension $st.GeneratedExtension
            }
        }.GetNewClosure())

        $buildGeneratedNow = {
            $st = $form.Tag
            if (-not $st) { return }
            if (-not $st.Contains('_GeneratedBuilder')) { return }
            if (-not $st['_GeneratedBuilder']) { return }
            if ([string]::IsNullOrWhiteSpace($st.CurrentFile) -or [string]::IsNullOrWhiteSpace($st.StructuredSource)) { return }
            if ($st.GeneratedBuilt -eq $true) { return }
            & $st['_GeneratedBuilder'] -Host $st.GeneratedHost -State $st -Content $st.StructuredSource -Extension $st.GeneratedExtension
        }.GetNewClosure()
        $state['BuildGeneratedNow'] = $buildGeneratedNow

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
                $rawText = Get-Content -LiteralPath $path -Raw -ErrorAction Stop
                $st.Editor.Text = $rawText
                $st.CurrentFile = $path
                $st.Status.Text = "Editing: $fileName"
                $ext = [System.IO.Path]::GetExtension($path)
                $st.StructuredModel = $null
                $st.StructuredControls = @{}
                $st.GeneratedBuilt = $false
                $st.StructuredSource = $rawText
                $st.GeneratedExtension = $ext
                $st.GeneratedFilter = ''
                if ($st.GeneratedHost -and $st.Contains('_GeneratedBuilder') -and $st['_GeneratedBuilder']) {
                    & $st['_GeneratedBuilder'] -Host $st.GeneratedHost -State $st -Content $rawText -Extension $ext
                } else {
                    _ShowGeneratedConfigPlaceholder -Host $st.GeneratedHost -Title 'Generated Editor Unavailable' -Message 'Generated editor could not be initialized for this file.'
                }
                if ($st.EditorTabs -and $st.GeneratedTab) {
                    $st.EditorTabs.SelectedTab = $st.GeneratedTab
                }
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
                $saveText = $st.Editor.Text
                $savingGenerated = $false
                if ($st.StructuredModel -and $st.StructuredModel.Supported -and $st.StructuredControls -and $st.StructuredControls.Count -gt 0) {
                    if (-not $st.EditorTabs -or -not $st.RawTab -or $st.EditorTabs.SelectedTab -ne $st.RawTab) {
                        $savingGenerated = $true
                    }
                }
                if ($script:SharedState -and $script:SharedState.Settings -and $script:SharedState.Settings.EnableDebugLogging) {
                    if ($script:SharedState.ContainsKey('LogQueue')) {
                        $selectedTabName = ''
                        try {
                            if ($st.EditorTabs -and $st.EditorTabs.SelectedTab) { $selectedTabName = [string]$st.EditorTabs.SelectedTab.Text }
                        } catch { $selectedTabName = '' }
                        $entryCount = 0
                        try { if ($st.StructuredModel -and $st.StructuredModel.Entries) { $entryCount = @($st.StructuredModel.Entries).Count } } catch { $entryCount = 0 }
                        $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][DEBUG][GUI] Config save begin: file=$($st.CurrentFile) mode=$(if($savingGenerated){'generated'}else{'raw'}) tab=$selectedTabName entries=$entryCount")
                        if ($savingGenerated -and $st.StructuredModel -and $st.StructuredModel.Entries) {
                            $zEntry = @($st.StructuredModel.Entries | Where-Object { $_.Type -eq 'entry' -and "$($_.Key)".Trim().ToLowerInvariant() -eq 'zombies' } | Select-Object -First 1)
                            if ($zEntry.Count -gt 0) {
                                $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][DEBUG][GUI] Generated entry Zombies before serialize: value=$($zEntry[0].Value)")
                            }
                        }
                    }
                }
                if ($savingGenerated) {
                    $validationErrors = _ValidateGeneratedConfig -State $st
                    if ($validationErrors.Count -gt 0) {
                        [System.Windows.Forms.MessageBox]::Show(
                            "Please fix these generated-editor values before saving:`n`n$($validationErrors -join [Environment]::NewLine)",
                            'Generated Config Validation','OK','Warning') | Out-Null
                        return
                    }
                    $saveText = _SerializeGeneratedConfig -Entries $st.StructuredModel.Entries -Controls $st.StructuredControls
                    $st.Editor.Text = $saveText
                    if ($script:SharedState -and $script:SharedState.Settings -and $script:SharedState.Settings.EnableDebugLogging) {
                        if ($script:SharedState.ContainsKey('LogQueue')) {
                            $zLine = (($saveText -split "\r?\n") | Where-Object { $_ -match '^\s*Zombies\s*=' } | Select-Object -First 1)
                            if ($zLine) {
                                $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][DEBUG][GUI] Serialized Zombies line: $zLine")
                            }
                        }
                    }
                }
                $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
                [System.IO.File]::WriteAllText($st.CurrentFile, $saveText, $utf8NoBom)
                $writtenText = ''
                try { $writtenText = Get-Content -LiteralPath $st.CurrentFile -Raw -ErrorAction Stop } catch { $writtenText = '' }
                if ($savingGenerated) {
                    $st.StructuredSource = $saveText
                    foreach ($entry in @($st.StructuredModel.Entries)) {
                        if ($entry.Type -ne 'entry') { continue }
                        $currentValue = [string]$entry.Value
                        $entry.Value = $currentValue
                        if ($entry.PSObject.Properties.Name -contains 'OriginalValue') {
                            $entry.OriginalValue = $currentValue
                        }
                    }
                }
                $st.Status.Text = "Saved: $(Split-Path $st.CurrentFile -Leaf)"
                if ($script:SharedState -and $script:SharedState.Settings -and $script:SharedState.Settings.EnableDebugLogging) {
                    if ($script:SharedState.ContainsKey('LogQueue')) {
                        $len = 0
                        try { $len = $saveText.Length } catch { $len = 0 }
                        $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][DEBUG][GUI] Config saved: $($st.CurrentFile) len=$len")
                        if ($writtenText) {
                            $diskZLine = (($writtenText -split "\r?\n") | Where-Object { $_ -match '^\s*Zombies\s*=' } | Select-Object -First 1)
                            if ($diskZLine) {
                                $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][DEBUG][GUI] Disk Zombies line after save: $diskZLine")
                            }
                        }
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
                $st.GeneratedBuilt = $false
                $st.StructuredModel = $null
                $st.StructuredControls = @{}
                $st.StructuredSource = ''
                $st.GeneratedExtension = ''
                if ($st.GeneratedHost) { $st.GeneratedHost.Controls.Clear() }
                & $st.Refresh $st
            })
        }

        $form.Add_Resize($layoutConfigEditor)
        $form.Add_Shown({ & $layoutConfigEditor }.GetNewClosure())

        & $state.Refresh $state
        $form.ShowDialog() | Out-Null
    }

    # =====================================================================
    # COMMAND CATALOG LOADER + COMMANDS WINDOW
    # =====================================================================
    function _LoadCommandCatalog {
        param([string]$Path)

        $pathToUse = $Path
        if ([string]::IsNullOrWhiteSpace($pathToUse)) { $pathToUse = $script:CommandCatalogPath }
        if ([string]::IsNullOrWhiteSpace($pathToUse)) { return $null }
        if (-not (Test-Path $pathToUse)) { return $null }

        try {
            $raw = Get-Content -Path $pathToUse -Raw -Encoding UTF8 | ConvertFrom-Json
            $script:CommandCatalog = $raw
            return $raw
        } catch {
            try {
                if ($script:SharedState -and $script:SharedState.LogQueue) {
                    $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][WARN][GUI] Failed to load CommandCatalog.json: $_")
                }
            } catch { }
            return $null
        }
    }

    function _FindCommandCatalogForGame {
        param([object]$Catalog, [string]$GameName)

        if ($null -eq $Catalog -or $null -eq $Catalog.Games -or [string]::IsNullOrWhiteSpace($GameName)) {
            return $null
        }

        $gamesObj = $Catalog.Games
        $props = $gamesObj.PSObject.Properties

        # Exact match
        $exact = $props | Where-Object { $_.Name -eq $GameName } | Select-Object -First 1
        if ($exact) { return $exact.Value }

        # Case-insensitive match
        $ci = $props | Where-Object { $_.Name -ieq $GameName } | Select-Object -First 1
        if ($ci) { return $ci.Value }

        # Fuzzy contains
        $fz = $props | Where-Object { $_.Name -like "*$GameName*" -or $GameName -like "*$($_.Name)*" } | Select-Object -First 1
        if ($fz) { return $fz.Value }

        $normalized = _NormalizeGameIdentity $GameName
        if (-not [string]::IsNullOrWhiteSpace($normalized)) {
            $normalizedMatch = $props | Where-Object { (_NormalizeGameIdentity $_.Name) -eq $normalized } | Select-Object -First 1
            if ($normalizedMatch) { return $normalizedMatch.Value }
        }

        return $null
    }

    function _BuildProfileCommandCatalogFallback {
        param([hashtable]$Profile)

        if ($null -eq $Profile) { return $null }

        $rawCommands = New-Object 'System.Collections.Generic.List[object]'
        foreach ($sourceName in @('Commands', 'ExtraCommands')) {
            $source = $null
            try { $source = $Profile[$sourceName] } catch { $source = $null }
            if ($null -eq $source) { continue }

            if ($source -is [System.Collections.IDictionary]) {
                foreach ($entry in $source.GetEnumerator()) {
                    if ($null -eq $entry) { continue }

                    $commandKey = ''
                    try { $commandKey = [string]$entry.Key } catch { $commandKey = '' }
                    if ([string]::IsNullOrWhiteSpace($commandKey)) { continue }

                    $commandValue = $entry.Value
                    $commandType = ''
                    $commandText = ''
                    try { if ($commandValue -and $commandValue.PSObject.Properties['Type']) { $commandType = [string]$commandValue.Type } } catch { $commandType = '' }
                    try { if ($commandValue -and $commandValue.PSObject.Properties['Command']) { $commandText = [string]$commandValue.Command } } catch { $commandText = '' }

                    if ([string]::IsNullOrWhiteSpace($commandText)) { $commandText = $commandKey }

                    $label = (($commandKey -replace '[_\-]+', ' ').Trim())
                    if ([string]::IsNullOrWhiteSpace($label)) { $label = $commandKey }

                    $category = 'Server'
                    switch ($commandType.ToLowerInvariant()) {
                        'rcon'        { $category = 'Remote' }
                        'sendcommand' { $category = 'Console' }
                        'telnet'      { $category = 'Remote' }
                        'rest'        { $category = 'Remote' }
                        'start'       { $category = 'Lifecycle' }
                        'stop'        { $category = 'Lifecycle' }
                        'restart'     { $category = 'Lifecycle' }
                        'status'      { $category = 'Lifecycle' }
                    }

                    $description = if (-not [string]::IsNullOrWhiteSpace($commandType)) {
                        "Runs the profile-defined $commandType command."
                    } else {
                        'Runs the profile-defined command.'
                    }

                    $rawCommands.Add([pscustomobject]@{
                        Id          = $commandKey.ToLowerInvariant()
                        Label       = $label
                        Command     = $commandText
                        Category    = $category
                        Description = $description
                        Source      = 'Profile'
                    }) | Out-Null
                }
            }
        }

        if ($rawCommands.Count -eq 0) { return $null }

        $deduped = @(
            $rawCommands |
            Group-Object Id |
            ForEach-Object { $_.Group | Select-Object -First 1 }
        )

        return [pscustomobject]@{
            SourceNotes = 'Generated from the current profile because the command catalog has no entry for this game.'
            Commands    = $deduped
        }
    }

    function _BuildCommandTooltipText {
        param([object]$Command)

        if ($null -eq $Command) { return '' }

        function _NormalizeCommandSentence {
            param([string]$Text)

            if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
            $result = $Text.Trim()
            if ($result -match '[.!?]$') { return $result }
            return ($result + '.')
        }

        function _ConvertCommandDescriptionToPlainEnglish {
            param(
                [string]$Description,
                [string]$CommandText
            )

            $text = if (-not [string]::IsNullOrWhiteSpace($Description)) { $Description.Trim() } else { '' }
            if ([string]::IsNullOrWhiteSpace($text) -and -not [string]::IsNullOrWhiteSpace($CommandText)) {
                $text = "Runs this command: $($CommandText.Trim())"
            }
            if ([string]::IsNullOrWhiteSpace($text)) { return '' }

            $replacements = [ordered]@{
                '^Show '                               = 'Shows '
                '^Display '                            = 'Shows '
                '^List '                               = 'Lists '
                '^Get '                                = 'Gets '
                '^Set '                                = 'Sets '
                '^Toggle '                             = 'Turns '
                '^Enable or disable '                  = 'Turns '
                '^Enable the '                         = 'Turns on the '
                '^Disable the '                        = 'Turns off the '
                '^Enable '                             = 'Turns on '
                '^Disable '                            = 'Turns off '
                '^Create '                             = 'Creates '
                '^Grant '                              = 'Gives '
                '^Give '                               = 'Gives '
                '^Remove '                             = 'Removes '
                '^Add '                                = 'Adds '
                '^Kick '                               = 'Kicks '
                '^Ban '                                = 'Bans '
                '^Unban '                              = 'Unbans '
                '^Teleport '                           = 'Teleports '
                '^Force '                              = 'Forces '
                '^Save '                               = 'Saves '
                '^Stop '                               = 'Stops '
                '^Shut down '                          = 'Shuts down '
                '^Exit '                               = 'Closes '
                '^Reload '                             = 'Reloads '
                '^Switch to '                          = 'Switches to '
                '^Switch '                             = 'Switches '
                '^Spawn '                              = 'Spawns '
                '^Run '                                = 'Runs '
                '^Check whether '                      = 'Checks whether '
                '^Check '                              = 'Checks '
                '^Apply '                              = 'Applies '
                '^Play '                               = 'Plays '
                '^Clear '                              = 'Clears '
                '^View '                               = 'Shows '
                '^Dump '                               = 'Shows '
                '^Make '                               = 'Makes '
            }

            foreach ($pattern in $replacements.Keys) {
                if ($text -match $pattern) {
                    $text = [regex]::Replace($text, $pattern, [string]$replacements[$pattern], 1)
                    break
                }
            }

            $text = $text -replace ' via REST API', ' by using the REST API'
            $text = $text -replace 'Toggle or control ', 'Turns on, turns off, or controls '
            $text = $text -replace 'Toggle ', 'Turns '
            $text = $text -replace ' \(same effect as quit\)', '. This does the same thing as quit'
            $text = $text -replace '^Legacy:\s*', 'Legacy command: '
            $text = $text -replace '^Legacy/Undocumented:\s*', 'Legacy or undocumented command: '

            return (_NormalizeCommandSentence $text)
        }

        function _BuildCommandUsageHelp {
            param([string]$CommandText)

            if ([string]::IsNullOrWhiteSpace($CommandText)) { return @() }

            $tips = New-Object 'System.Collections.Generic.List[string]'
            $template = $CommandText.Trim()
            $tips.Add("Command to send: $template") | Out-Null

            if ($template -match '<[^>]+>') {
                $tips.Add('Use the values inside < > as required parts you must fill in.') | Out-Null
            }
            if ($template -match '\[[^\]]+\]') {
                $tips.Add('Anything inside [ ] is optional.') | Out-Null
            }

            $requiredMatches = [regex]::Matches($template, '<([^>]+)>')
            if ($requiredMatches.Count -gt 0) {
                $requiredParts = @()
                foreach ($match in @($requiredMatches)) {
                    $part = ''
                    try { $part = [string]$match.Groups[1].Value } catch { $part = '' }
                    if ([string]::IsNullOrWhiteSpace($part)) { continue }
                    $part = $part -replace '\|', ' or '
                    $part = $part -replace '\s+', ' '
                    $requiredParts += $part.Trim()
                }
                $requiredParts = @($requiredParts | Select-Object -Unique)
                if ($requiredParts.Count -gt 0) {
                    $tips.Add(('Required parts: {0}.' -f ($requiredParts -join '; '))) | Out-Null
                }
            }

            return @($tips)
        }

        $lines = New-Object 'System.Collections.Generic.List[string]'

        $label = ''
        $commandText = ''
        $description = ''
        $syntax = ''
        $notes = ''

        try { $label = [string]$Command.Label } catch { $label = '' }
        try { $commandText = [string]$Command.Command } catch { $commandText = '' }
        try { $description = [string]$Command.Description } catch { $description = '' }
        try { $syntax = [string]$Command.Syntax } catch { $syntax = '' }
        try { $notes = [string]$Command.Notes } catch { $notes = '' }

        if (-not [string]::IsNullOrWhiteSpace($label)) {
            $lines.Add($label.Trim()) | Out-Null
        }

        $plainDescription = _ConvertCommandDescriptionToPlainEnglish -Description $description -CommandText $commandText
        if (-not [string]::IsNullOrWhiteSpace($plainDescription)) {
            $lines.Add($plainDescription) | Out-Null
        } elseif (-not [string]::IsNullOrWhiteSpace($commandText)) {
            $lines.Add((_NormalizeCommandSentence $commandText.Trim())) | Out-Null
        }

        foreach ($usageLine in @(_BuildCommandUsageHelp -CommandText $commandText)) {
            if (-not [string]::IsNullOrWhiteSpace([string]$usageLine)) {
                $lines.Add([string]$usageLine) | Out-Null
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($syntax)) {
            $lines.Add("Extra syntax note: $($syntax.Trim())") | Out-Null
        }

        if ($Command.PSObject.Properties['Values'] -and $null -ne $Command.Values) {
            $valueList = @($Command.Values)
            if ($valueList.Count -gt 0) {
                $lines.Add('Values:') | Out-Null
                foreach ($valueEntry in $valueList) {
                    if ($null -eq $valueEntry) { continue }
                    $valueName = ''
                    $valueDesc = ''
                    try { $valueName = [string]$valueEntry.Name } catch { $valueName = '' }
                    try { $valueDesc = [string]$valueEntry.Description } catch { $valueDesc = '' }

                    if ([string]::IsNullOrWhiteSpace($valueName) -and -not [string]::IsNullOrWhiteSpace($valueDesc)) {
                        $lines.Add(" - $($valueDesc.Trim())") | Out-Null
                    } elseif (-not [string]::IsNullOrWhiteSpace($valueName) -and -not [string]::IsNullOrWhiteSpace($valueDesc)) {
                        $lines.Add(" - $($valueName.Trim()): $($valueDesc.Trim())") | Out-Null
                    } elseif (-not [string]::IsNullOrWhiteSpace($valueName)) {
                        $lines.Add(" - $($valueName.Trim())") | Out-Null
                    }
                }
            }
        }

        if ($Command.PSObject.Properties['Examples'] -and $null -ne $Command.Examples) {
            $exampleList = @($Command.Examples | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
            if ($exampleList.Count -gt 0) {
                $lines.Add('Examples:') | Out-Null
                foreach ($example in $exampleList) {
                    $lines.Add(" - $([string]$example)".TrimEnd()) | Out-Null
                }
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($notes)) {
            $lines.Add("Extra note: $($notes.Trim())") | Out-Null
        }

        if ($Command.PSObject.Properties['Deprecated'] -and [bool]$Command.Deprecated) {
            $lines.Add('Deprecated: this entry is kept for reference and may not work on the current branch.') | Out-Null
        }

        if ($Command.PSObject.Properties['Legacy'] -and [bool]$Command.Legacy) {
            $lines.Add('Legacy/undocumented: verify on your server branch before relying on it.') | Out-Null
        }

        if ($Command.PSObject.Properties['Category'] -and $Command.Category) {
            $lines.Add("Category: $([string]$Command.Category)") | Out-Null
        }

        return (($lines | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join [Environment]::NewLine).Trim()
    }

    function _NormalizeProjectZomboidAssetKey {
        param([string]$Text)

        if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
        $normalized = [string]$Text
        $normalized = $normalized -replace '\\', '/'
        $normalized = [System.IO.Path]::GetFileNameWithoutExtension($normalized)
        $normalized = $normalized -replace '^Item_', ''
        $normalized = $normalized -replace '^items/', ''
        $normalized = $normalized -replace '[^A-Za-z0-9]+', ''
        return $normalized.ToLowerInvariant()
    }

    function _GetProjectZomboidAssetCacheRoot {
        $workspaceRoot = Split-Path -Parent $PSScriptRoot
        return (Join-Path $workspaceRoot 'Config\AssetCache\ProjectZomboid')
    }

    function _GetProjectZomboidClientRoots {
        param([hashtable]$Profile)

        $roots = New-Object 'System.Collections.Generic.List[string]'
        $addRoot = {
            param([string]$PathValue)
            if ([string]::IsNullOrWhiteSpace($PathValue)) { return }
            $expanded = ''
            try { $expanded = [Environment]::ExpandEnvironmentVariables($PathValue) } catch { $expanded = $PathValue }
            if ([string]::IsNullOrWhiteSpace($expanded)) { return }
            if (-not (Test-Path -LiteralPath $expanded)) { return }
            $normalized = [System.IO.Path]::GetFullPath($expanded)
            if (-not $roots.Contains($normalized)) { $roots.Add($normalized) | Out-Null }
        }
        $addSteamCandidate = {
            param([string]$BasePath)

            if ([string]::IsNullOrWhiteSpace($BasePath)) { return }
            $expanded = ''
            try { $expanded = [Environment]::ExpandEnvironmentVariables($BasePath) } catch { $expanded = $BasePath }
            if ([string]::IsNullOrWhiteSpace($expanded) -or -not (Test-Path -LiteralPath $expanded)) { return }

            $normalized = ''
            try { $normalized = [System.IO.Path]::GetFullPath($expanded) } catch { $normalized = $expanded }
            if ([string]::IsNullOrWhiteSpace($normalized)) { return }

            $leaf = ''
            try { $leaf = [System.IO.Path]::GetFileName(($normalized.TrimEnd('\'))) } catch { $leaf = '' }
            if ($leaf -ieq 'ProjectZomboid') {
                & $addRoot $normalized
                return
            }

            foreach ($candidate in @(
                (Join-Path $normalized 'steamapps\common\ProjectZomboid'),
                (Join-Path $normalized 'common\ProjectZomboid'),
                (Join-Path $normalized 'ProjectZomboid')
            )) {
                & $addRoot $candidate
            }
        }
        $addSteamLibrariesFromRoot = {
            param([string]$SteamRoot)

            if ([string]::IsNullOrWhiteSpace($SteamRoot)) { return }
            & $addSteamCandidate $SteamRoot

            $libraryVdfPath = Join-Path $SteamRoot 'steamapps\libraryfolders.vdf'
            if (-not (Test-Path -LiteralPath $libraryVdfPath)) { return }

            try {
                $vdfLines = Get-Content -LiteralPath $libraryVdfPath -ErrorAction Stop
                foreach ($line in $vdfLines) {
                    if ($line -match '"path"\s+"([^"]+)"') {
                        $libPath = $Matches[1] -replace '\\\\', '\'
                        & $addSteamCandidate $libPath
                    }
                }
            } catch { }
        }

        try { & $addRoot ([string]$Profile.FolderPath) } catch { }
        try {
            if ($script:SharedState -and $script:SharedState.Settings) {
                & $addRoot ([string]$script:SharedState.Settings.ProjectZomboidClientPath)
            }
        } catch { }
        try {
            $profileRoot = [string]$Profile.FolderPath
            if (-not [string]::IsNullOrWhiteSpace($profileRoot) -and (Test-Path -LiteralPath $profileRoot)) {
                $profileFull = [System.IO.Path]::GetFullPath($profileRoot)
                $commonDir = Split-Path -Parent $profileFull
                if (-not [string]::IsNullOrWhiteSpace($commonDir)) {
                    & $addSteamCandidate (Join-Path $commonDir 'ProjectZomboid')
                }
            }
        } catch { }
        foreach ($steamRoot in @(
            (try { (Get-ItemProperty -Path 'HKCU:\Software\Valve\Steam' -Name 'SteamPath' -ErrorAction Stop).SteamPath } catch { $null }),
            (try { (Get-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\Valve\Steam' -Name 'InstallPath' -ErrorAction Stop).InstallPath } catch { $null }),
            (Join-Path ${env:ProgramFiles(x86)} 'Steam'),
            (Join-Path $env:ProgramFiles 'Steam')
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique) {
            & $addSteamLibrariesFromRoot $steamRoot
        }
        try {
            foreach ($drive in @(Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue)) {
                if (-not $drive.Root) { continue }
                foreach ($basePath in @(
                    (Join-Path $drive.Root 'SteamLibrary'),
                    (Join-Path $drive.Root 'Program Files (x86)\Steam'),
                    (Join-Path $drive.Root 'Program Files\Steam')
                )) {
                    & $addSteamLibrariesFromRoot $basePath
                }
            }
        } catch { }

        return @($roots)
    }

    function _LoadProjectZomboidAssetCacheManifest {
        $cacheRoot = _GetProjectZomboidAssetCacheRoot
        $manifestPath = Join-Path $cacheRoot 'item-texture-manifest.json'
        if (Test-Path -LiteralPath $manifestPath) {
            try {
                $manifest = Get-Content -LiteralPath $manifestPath -Raw -ErrorAction Stop | ConvertFrom-Json
                if ($manifest) { return $manifest }
            } catch { }
        }

        $importRoot = Join-Path $cacheRoot 'ImportedClient'
        $importManifestPath = Join-Path $importRoot 'import-manifest.json'
        if (-not (Test-Path -LiteralPath $importManifestPath)) { return $null }

        try {
            $importManifest = Get-Content -LiteralPath $importManifestPath -Raw -ErrorAction Stop | ConvertFrom-Json
            if (-not $importManifest) { return $null }

            $itemFolder = [string]$importManifest.itemFolder
            $metadataFolder = [string]$importManifest.metadataFolder
            $itemXmlPath = if (-not [string]::IsNullOrWhiteSpace($metadataFolder)) { Join-Path $metadataFolder 'items.xml' } else { '' }
            if ([string]::IsNullOrWhiteSpace($itemFolder) -or -not (Test-Path -LiteralPath $itemFolder)) { return $null }

            $fileByExactKey = @{}
            $fileByNormalizedKey = @{}
            foreach ($file in @(Get-ChildItem -LiteralPath $itemFolder -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -in @('.png', '.jpg', '.jpeg') })) {
                $fullPath = $file.FullName
                $leafName = $file.BaseName
                foreach ($key in @($leafName.ToLowerInvariant(), (_NormalizeProjectZomboidAssetKey $leafName)) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique) {
                    if (-not $fileByExactKey.ContainsKey($key)) { $fileByExactKey[$key] = $fullPath }
                    if (-not $fileByNormalizedKey.ContainsKey($key)) { $fileByNormalizedKey[$key] = $fullPath }
                }
            }

            $map = @{}
            if (-not [string]::IsNullOrWhiteSpace($itemXmlPath) -and (Test-Path -LiteralPath $itemXmlPath)) {
                [xml]$itemXml = Get-Content -LiteralPath $itemXmlPath -ErrorAction Stop
                foreach ($node in @($itemXml.itemManager.m_Items)) {
                    $textureRef = [string]$node.m_Texture
                    $modelRef = [string]$node.m_Model
                    if ([string]::IsNullOrWhiteSpace($textureRef)) { continue }

                    $resolvedPath = $null
                    foreach ($key in @(
                        ([System.IO.Path]::GetFileName($textureRef)).ToLowerInvariant(),
                        (_NormalizeProjectZomboidAssetKey $textureRef)
                    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique) {
                        if ($fileByExactKey.ContainsKey($key)) { $resolvedPath = $fileByExactKey[$key]; break }
                        if ($fileByNormalizedKey.ContainsKey($key)) { $resolvedPath = $fileByNormalizedKey[$key]; break }
                    }
                    if (-not $resolvedPath) { continue }

                    $itemKeys = @(
                        ([System.IO.Path]::GetFileName($textureRef)).ToLowerInvariant(),
                        (_NormalizeProjectZomboidAssetKey $textureRef)
                    )
                    if (-not [string]::IsNullOrWhiteSpace($modelRef)) {
                        $itemKeys += ([System.IO.Path]::GetFileName($modelRef)).ToLowerInvariant()
                        $itemKeys += (_NormalizeProjectZomboidAssetKey $modelRef)
                    }

                    foreach ($key in ($itemKeys | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)) {
                        if (-not $map.ContainsKey($key)) {
                            $map[$key] = $resolvedPath
                        }
                    }
                }
            }

            return [pscustomobject]@{
                SourceRoot = if ($importManifest.PSObject.Properties.Name -contains 'sourceRoot') { [string]$importManifest.sourceRoot } else { $importRoot }
                SourceTicks = if ($importManifest.PSObject.Properties.Name -contains 'importedAt') { [string]$importManifest.importedAt } else { '' }
                UpdatedAt = if ($importManifest.PSObject.Properties.Name -contains 'importedAt') { [string]$importManifest.importedAt } else { (Get-Date).ToString('o') }
                ItemTextureByName = $map
            }
        } catch {
            return $null
        }
    }

    function _GetProjectZomboidCatalogCacheDirectory {
        $cacheRoot = _GetProjectZomboidAssetCacheRoot
        $dir = Join-Path $cacheRoot 'Catalogs'
        if (-not (Test-Path -LiteralPath $dir)) {
            New-Item -ItemType Directory -Force -Path $dir | Out-Null
        }
        return $dir
    }

    function _GetProjectZomboidCatalogCachePath {
        param(
            [string]$CatalogName,
            [string]$CacheKey
        )

        $safeName = if ([string]::IsNullOrWhiteSpace($CatalogName)) { 'catalog' } else { $CatalogName }
        $safeKey = if ([string]::IsNullOrWhiteSpace($CacheKey)) { 'default' } else { ($CacheKey -replace '[^A-Za-z0-9\-_\.]+', '_') }
        return (Join-Path (_GetProjectZomboidCatalogCacheDirectory) "$safeName-$safeKey.json")
    }

    function _LoadProjectZomboidCatalogCache {
        param(
            [string]$CatalogName,
            [string]$CacheKey
        )

        $path = _GetProjectZomboidCatalogCachePath -CatalogName $CatalogName -CacheKey $CacheKey
        if (-not (Test-Path -LiteralPath $path)) { return $null }
        try {
            $raw = Get-Content -LiteralPath $path -Raw -ErrorAction Stop | ConvertFrom-Json
            if ($raw -is [System.Collections.IEnumerable]) {
                return @($raw | ForEach-Object { [pscustomobject]$_ })
            }
        } catch { }
        return $null
    }

    function _SaveProjectZomboidCatalogCache {
        param(
            [string]$CatalogName,
            [string]$CacheKey,
            [object[]]$Entries
        )

        $path = _GetProjectZomboidCatalogCachePath -CatalogName $CatalogName -CacheKey $CacheKey
        try {
            $json = @($Entries) | ConvertTo-Json -Depth 6
            [System.IO.File]::WriteAllText($path, $json, (New-Object System.Text.UTF8Encoding($false)))
        } catch { }
    }

    function _GetProjectZomboidAssetIndex {
        param([string]$GameRoot)

        if ([string]::IsNullOrWhiteSpace($GameRoot) -or -not (Test-Path -LiteralPath $GameRoot)) { return $null }
        if ($null -eq $script:ProjectZomboidAssetIndexCache) {
            $script:ProjectZomboidAssetIndexCache = @{}
        }

        $itemXmlPath = Join-Path $GameRoot 'media\items\items.xml'
        $latestTicks = 0L
        try {
            $candidateFiles = @()
            if (Test-Path -LiteralPath $itemXmlPath) {
                $candidateFiles += Get-Item -LiteralPath $itemXmlPath -ErrorAction SilentlyContinue
            }
            foreach ($dir in @(
                (Join-Path $GameRoot 'media\textures'),
                (Join-Path $GameRoot 'media\inventory')
            )) {
                if (-not (Test-Path -LiteralPath $dir)) { continue }
                $latest = Get-ChildItem -LiteralPath $dir -Recurse -File -ErrorAction SilentlyContinue |
                    Where-Object { $_.Extension -in @('.png', '.jpg', '.jpeg') } |
                    Sort-Object LastWriteTimeUtc -Descending |
                    Select-Object -First 1
                if ($latest) { $candidateFiles += $latest }
            }
            if ($candidateFiles.Count -gt 0) {
                $latestTicks = ($candidateFiles | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1).LastWriteTimeUtc.Ticks
            }
        } catch { }

        $cacheKey = "$GameRoot|assets|$latestTicks"
        if ($script:ProjectZomboidAssetIndexCache.ContainsKey($cacheKey)) {
            return $script:ProjectZomboidAssetIndexCache[$cacheKey]
        }

        $index = [ordered]@{
            ByExactPathKey    = @{}
            ByFileNameKey     = @{}
            ByNormalizedKey   = @{}
            ItemTextureByName = @{}
        }

        $assetRoots = @(
            (Join-Path $GameRoot 'media\textures'),
            (Join-Path $GameRoot 'media\inventory'),
            (Join-Path $GameRoot 'media\ui'),
            (Join-Path $GameRoot 'media\ui\ItemIcons')
        ) | Where-Object { Test-Path -LiteralPath $_ }

        foreach ($dir in $assetRoots) {
            foreach ($file in @(Get-ChildItem -LiteralPath $dir -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -in @('.png', '.jpg', '.jpeg') })) {
                $fullPath = $file.FullName
                $leafName = $file.BaseName
                $relativePath = $fullPath.Substring($GameRoot.Length).TrimStart('\').Replace('\','/')
                $exactKeys = @(
                    $relativePath.ToLowerInvariant(),
                    $leafName.ToLowerInvariant()
                ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
                foreach ($key in $exactKeys) {
                    if (-not $index.ByExactPathKey.ContainsKey($key)) { $index.ByExactPathKey[$key] = $fullPath }
                    if (-not $index.ByFileNameKey.ContainsKey($key)) { $index.ByFileNameKey[$key] = $fullPath }
                }

                $normalizedKeys = @(
                    (_NormalizeProjectZomboidAssetKey $relativePath),
                    (_NormalizeProjectZomboidAssetKey $leafName)
                ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
                foreach ($key in $normalizedKeys) {
                    if (-not $index.ByNormalizedKey.ContainsKey($key)) { $index.ByNormalizedKey[$key] = $fullPath }
                }
            }
        }

        if (Test-Path -LiteralPath $itemXmlPath) {
            try {
                [xml]$itemXml = Get-Content -LiteralPath $itemXmlPath -ErrorAction Stop
                foreach ($node in @($itemXml.itemManager.m_Items)) {
                    $textureRef = [string]$node.m_Texture
                    $modelRef = [string]$node.m_Model
                    if ([string]::IsNullOrWhiteSpace($textureRef)) { continue }

                    $resolvedPath = $null
                    $candidateKeys = @(
                        $textureRef.ToLowerInvariant(),
                        ([System.IO.Path]::GetFileName($textureRef)).ToLowerInvariant(),
                        (_NormalizeProjectZomboidAssetKey $textureRef)
                    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique

                    foreach ($key in $candidateKeys) {
                        if ($index.ByExactPathKey.ContainsKey($key)) { $resolvedPath = $index.ByExactPathKey[$key]; break }
                        if ($index.ByFileNameKey.ContainsKey($key))  { $resolvedPath = $index.ByFileNameKey[$key]; break }
                        if ($index.ByNormalizedKey.ContainsKey($key)) { $resolvedPath = $index.ByNormalizedKey[$key]; break }
                    }
                    if (-not $resolvedPath) { continue }

                    $itemKeys = @(
                        ([System.IO.Path]::GetFileName($textureRef)).ToLowerInvariant(),
                        (_NormalizeProjectZomboidAssetKey $textureRef)
                    )
                    if (-not [string]::IsNullOrWhiteSpace($modelRef)) {
                        $itemKeys += ([System.IO.Path]::GetFileName($modelRef)).ToLowerInvariant()
                        $itemKeys += (_NormalizeProjectZomboidAssetKey $modelRef)
                    }

                    foreach ($key in ($itemKeys | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)) {
                        if (-not $index.ItemTextureByName.ContainsKey($key)) {
                            $index.ItemTextureByName[$key] = $resolvedPath
                        }
                    }
                }
            } catch { }
        }

        $script:ProjectZomboidAssetIndexCache.Clear()
        $script:ProjectZomboidAssetIndexCache[$cacheKey] = [pscustomobject]$index
        return $script:ProjectZomboidAssetIndexCache[$cacheKey]
    }

    function _SyncProjectZomboidItemAssetCache {
        param([hashtable]$Profile)

        $clientRoot = $null
        foreach ($root in @(_GetProjectZomboidClientRoots -Profile $Profile)) {
            if (-not [string]::IsNullOrWhiteSpace($root) -and (Test-Path -LiteralPath (Join-Path $root 'media\items\items.xml'))) {
                $clientRoot = $root
                break
            }
        }
        if ([string]::IsNullOrWhiteSpace($clientRoot)) { return $null }

        $assetIndex = _GetProjectZomboidAssetIndex -GameRoot $clientRoot
        if ($null -eq $assetIndex -or $assetIndex.ItemTextureByName.Count -eq 0) { return $null }

        $cacheRoot = _GetProjectZomboidAssetCacheRoot
        $itemsCacheRoot = Join-Path $cacheRoot 'Items'
        $manifestPath = Join-Path $cacheRoot 'item-texture-manifest.json'

        $latestTicks = 0L
        try {
            $itemXml = Get-Item -LiteralPath (Join-Path $clientRoot 'media\items\items.xml') -ErrorAction SilentlyContinue
            $latestTexture = Get-ChildItem -LiteralPath (Join-Path $clientRoot 'media\textures') -Recurse -File -ErrorAction SilentlyContinue |
                Where-Object { $_.Extension -in @('.png', '.jpg', '.jpeg') } |
                Sort-Object LastWriteTimeUtc -Descending |
                Select-Object -First 1
            $candidates = @($itemXml, $latestTexture) | Where-Object { $null -ne $_ }
            if ($candidates.Count -gt 0) {
                $latestTicks = ($candidates | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1).LastWriteTimeUtc.Ticks
            }
        } catch { }

        try {
            $existing = _LoadProjectZomboidAssetCacheManifest
            if ($existing -and [string]$existing.SourceRoot -eq $clientRoot -and [string]$existing.SourceTicks -eq [string]$latestTicks) {
                return $existing
            }
        } catch { }

        New-Item -ItemType Directory -Force -Path $itemsCacheRoot | Out-Null

        $manifestMap = @{}
        $copiedTargets = @{}
        foreach ($entry in $assetIndex.ItemTextureByName.GetEnumerator()) {
            $key = [string]$entry.Key
            $sourcePath = [string]$entry.Value
            if ([string]::IsNullOrWhiteSpace($key) -or [string]::IsNullOrWhiteSpace($sourcePath)) { continue }
            if (-not (Test-Path -LiteralPath $sourcePath)) { continue }

            $relativePath = ''
            try {
                $relativePath = $sourcePath.Substring($clientRoot.Length).TrimStart('\')
            } catch {
                $relativePath = Split-Path -Leaf $sourcePath
            }
            if ([string]::IsNullOrWhiteSpace($relativePath)) {
                $relativePath = Split-Path -Leaf $sourcePath
            }
            $targetPath = Join-Path $itemsCacheRoot $relativePath
            $targetDir = Split-Path -Parent $targetPath
            if (-not (Test-Path -LiteralPath $targetDir)) {
                New-Item -ItemType Directory -Force -Path $targetDir | Out-Null
            }

            if (-not $copiedTargets.ContainsKey($targetPath)) {
                Copy-Item -LiteralPath $sourcePath -Destination $targetPath -Force
                $copiedTargets[$targetPath] = $true
            }
            $manifestMap[$key] = $targetPath
        }

        $manifest = [ordered]@{
            SourceRoot = $clientRoot
            SourceTicks = "$latestTicks"
            UpdatedAt = (Get-Date).ToString('o')
            ItemTextureByName = $manifestMap
        }
        $manifestJson = $manifest | ConvertTo-Json -Depth 6
        [System.IO.File]::WriteAllText($manifestPath, $manifestJson, (New-Object System.Text.UTF8Encoding($false)))

        return [pscustomobject]$manifest
    }

    function _ResolveProjectZomboidItemPreviewPath {
        param(
            [hashtable]$Profile,
            [string]$ItemName,
            [string]$IconName,
            [object]$CacheManifest = $null
        )

        $candidateKeys = New-Object 'System.Collections.Generic.List[string]'
        foreach ($value in @($IconName, $ItemName)) {
            if ([string]::IsNullOrWhiteSpace($value)) { continue }
            $candidateKeys.Add(([string]$value).ToLowerInvariant()) | Out-Null
            $candidateKeys.Add((_NormalizeProjectZomboidAssetKey $value)) | Out-Null
            $candidateKeys.Add(("items/{0}" -f [string]$value).ToLowerInvariant()) | Out-Null
            $candidateKeys.Add((_NormalizeProjectZomboidAssetKey ("items/{0}" -f [string]$value))) | Out-Null
            $candidateKeys.Add(("Item_{0}" -f [string]$value).ToLowerInvariant()) | Out-Null
            $candidateKeys.Add((_NormalizeProjectZomboidAssetKey ("Item_{0}" -f [string]$value))) | Out-Null
        }

        if ($null -eq $CacheManifest) {
            $CacheManifest = _LoadProjectZomboidAssetCacheManifest
        }
        if ($null -eq $CacheManifest) {
            $CacheManifest = _SyncProjectZomboidItemAssetCache -Profile $Profile
        }

        foreach ($key in @($candidateKeys | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)) {
            try {
                if ($CacheManifest -and $CacheManifest.ItemTextureByName -and $CacheManifest.ItemTextureByName.PSObject.Properties.Name -contains $key) {
                    $cachedPath = [string]$CacheManifest.ItemTextureByName.$key
                    if (-not [string]::IsNullOrWhiteSpace($cachedPath) -and (Test-Path -LiteralPath $cachedPath)) {
                        return $cachedPath
                    }
                }
            } catch { }
        }

        $fallbackRoots = @(_GetProjectZomboidClientRoots -Profile $Profile)
        foreach ($root in $fallbackRoots) {
            $assetIndex = _GetProjectZomboidAssetIndex -GameRoot $root
            if ($null -eq $assetIndex) { continue }
            foreach ($key in @($candidateKeys | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)) {
                if ($assetIndex.ItemTextureByName.ContainsKey($key)) { return $assetIndex.ItemTextureByName[$key] }
                if ($assetIndex.ByExactPathKey.ContainsKey($key))    { return $assetIndex.ByExactPathKey[$key] }
                if ($assetIndex.ByFileNameKey.ContainsKey($key))     { return $assetIndex.ByFileNameKey[$key] }
                if ($assetIndex.ByNormalizedKey.ContainsKey($key))   { return $assetIndex.ByNormalizedKey[$key] }
            }
        }

        return $null
    }

    function _GetProjectZomboidItemCatalog {
        param([hashtable]$Profile)

        $gameRoot = ''
        try { $gameRoot = [string]$Profile.FolderPath } catch { $gameRoot = '' }
        if ([string]::IsNullOrWhiteSpace($gameRoot) -or -not (Test-Path -LiteralPath $gameRoot)) {
            return @()
        }

        $scriptsDir = Join-Path $gameRoot 'media\scripts\generated\items'
        if (-not (Test-Path -LiteralPath $scriptsDir)) {
            $scriptsDir = Join-Path $gameRoot 'media\scripts\items'
        }
        if (-not (Test-Path -LiteralPath $scriptsDir)) {
            return @()
        }

        if ($null -eq $script:ProjectZomboidItemCatalogCache) {
            $script:ProjectZomboidItemCatalogCache = @{}
        }

        $latestTicks = 0L
        try {
            $latestFile = Get-ChildItem -LiteralPath $scriptsDir -Recurse -File -Include *.txt -ErrorAction SilentlyContinue |
                Sort-Object LastWriteTimeUtc -Descending |
                Select-Object -First 1
            if ($latestFile) { $latestTicks = $latestFile.LastWriteTimeUtc.Ticks }
        } catch { }

        $assetManifest = _LoadProjectZomboidAssetCacheManifest
        if ($null -eq $assetManifest) {
            $assetManifest = _SyncProjectZomboidItemAssetCache -Profile $Profile
        }
        $assetSignature = ''
        try {
            if ($assetManifest) {
                $assetSignature = if ($assetManifest.PSObject.Properties.Name -contains 'SourceTicks') { [string]$assetManifest.SourceTicks } else { '' }
                if ([string]::IsNullOrWhiteSpace($assetSignature) -and $assetManifest.PSObject.Properties.Name -contains 'UpdatedAt') {
                    $assetSignature = [string]$assetManifest.UpdatedAt
                }
            }
        } catch { $assetSignature = '' }

        $cacheKey = "$gameRoot|$latestTicks|$assetSignature"
        if ($script:ProjectZomboidItemCatalogCache.ContainsKey($cacheKey)) {
            return @($script:ProjectZomboidItemCatalogCache[$cacheKey])
        }

        $diskCachedCatalog = _LoadProjectZomboidCatalogCache -CatalogName 'pz-item-catalog' -CacheKey $cacheKey
        if ($diskCachedCatalog -and $diskCachedCatalog.Count -gt 0) {
            $script:ProjectZomboidItemCatalogCache.Clear()
            $script:ProjectZomboidItemCatalogCache[$cacheKey] = @($diskCachedCatalog)
            return @($diskCachedCatalog)
        }

        $results = New-Object System.Collections.Generic.List[object]
        $files = @(Get-ChildItem -LiteralPath $scriptsDir -Recurse -File -Include *.txt -ErrorAction SilentlyContinue)

        foreach ($file in $files) {
            $moduleName = 'Base'
            $inItem = $false
            $itemName = ''
            $props = @{}

            foreach ($rawLine in (Get-Content -LiteralPath $file.FullName -ErrorAction SilentlyContinue)) {
                $line = ($rawLine -replace '//.*$', '').Trim()
                if ([string]::IsNullOrWhiteSpace($line)) { continue }

                if (-not $inItem) {
                    if ($line -match '^module\s+([A-Za-z0-9_]+)\s*$') {
                        $moduleName = $Matches[1]
                        continue
                    }
                    if ($line -match '^item\s+([A-Za-z0-9_]+)\s*$') {
                        $inItem = $true
                        $itemName = $Matches[1]
                        $props = @{}
                    }
                    continue
                }

                if ($line -eq '{') { continue }
                if ($line -eq '}') {
                    $isObsolete = $false
                    $isHidden = $false
                    if ($props.ContainsKey('OBSOLETE')) { $isObsolete = ($props['OBSOLETE'] -match '^(?i:true|yes|1)$') }
                    if ($props.ContainsKey('Hidden'))   { $isHidden   = ($props['Hidden']   -match '^(?i:true|yes|1)$') }

                    if (-not $isObsolete -and -not $isHidden -and -not [string]::IsNullOrWhiteSpace($itemName)) {
                        $displayName = if ($props.ContainsKey('DisplayName') -and -not [string]::IsNullOrWhiteSpace($props['DisplayName'])) {
                            [string]$props['DisplayName']
                        } else {
                            $itemName
                        }
                        $iconName = if ($props.ContainsKey('Icon')) { [string]$props['Icon'] } else { '' }

                        $results.Add([pscustomobject]@{
                            FullType        = "$moduleName.$itemName"
                            DisplayName     = $displayName
                            ListText        = "$displayName [$moduleName.$itemName]"
                            Module          = $moduleName
                            ItemName        = $itemName
                            DisplayCategory = if ($props.ContainsKey('DisplayCategory')) { [string]$props['DisplayCategory'] } else { '' }
                            ItemType        = if ($props.ContainsKey('ItemType')) { [string]$props['ItemType'] } else { '' }
                            IconName        = $iconName
                            IconPath        = _ResolveProjectZomboidItemPreviewPath -Profile $Profile -ItemName $itemName -IconName $iconName -CacheManifest $assetManifest
                        }) | Out-Null
                    }

                    $inItem = $false
                    $itemName = ''
                    $props = @{}
                    continue
                }

                if ($line -match '^([A-Za-z0-9_]+)\s*=\s*(.+?)(?:,)?$') {
                    $props[$Matches[1]] = $Matches[2].Trim()
                }
            }
        }

        $catalog = @(
            $results |
            Sort-Object @{ Expression = { if ([string]::IsNullOrWhiteSpace($_.DisplayCategory)) { 'zzz_misc' } else { $_.DisplayCategory } } }, DisplayName, FullType -Unique
        )
        $script:ProjectZomboidItemCatalogCache.Clear()
        $script:ProjectZomboidItemCatalogCache[$cacheKey] = $catalog
        _SaveProjectZomboidCatalogCache -CatalogName 'pz-item-catalog' -CacheKey $cacheKey -Entries $catalog
        return $catalog
    }

    function _FormatProjectZomboidDisplayName {
        param([string]$Name)

        if ([string]::IsNullOrWhiteSpace($Name)) { return '' }

        $formatted = $Name -replace '_', ' '
        $formatted = $formatted -creplace '([a-z0-9])([A-Z])', '$1 $2'
        $formatted = $formatted -replace '\s+', ' '
        return $formatted.Trim()
    }

    function _GetProjectZomboidVehicleCatalog {
        param([hashtable]$Profile)

        $gameRoot = ''
        try { $gameRoot = [string]$Profile.FolderPath } catch { $gameRoot = '' }
        if ([string]::IsNullOrWhiteSpace($gameRoot) -or -not (Test-Path -LiteralPath $gameRoot)) {
            return @()
        }

        $scriptsDir = Join-Path $gameRoot 'media\scripts\generated\vehicles'
        if (-not (Test-Path -LiteralPath $scriptsDir)) {
            $scriptsDir = Join-Path $gameRoot 'media\scripts\vehicles'
        }
        if (-not (Test-Path -LiteralPath $scriptsDir)) {
            return @()
        }

        if ($null -eq $script:ProjectZomboidVehicleCatalogCache) {
            $script:ProjectZomboidVehicleCatalogCache = @{}
        }

        $latestTicks = 0L
        try {
            $latestFile = Get-ChildItem -LiteralPath $scriptsDir -Recurse -File -Include *.txt -ErrorAction SilentlyContinue |
                Sort-Object LastWriteTimeUtc -Descending |
                Select-Object -First 1
            if ($latestFile) { $latestTicks = $latestFile.LastWriteTimeUtc.Ticks }
        } catch { }

        $cacheKey = "$gameRoot|vehicles|$latestTicks"
        if ($script:ProjectZomboidVehicleCatalogCache.ContainsKey($cacheKey)) {
            return @($script:ProjectZomboidVehicleCatalogCache[$cacheKey])
        }

        $diskCachedCatalog = _LoadProjectZomboidCatalogCache -CatalogName 'pz-vehicle-catalog' -CacheKey $cacheKey
        if ($diskCachedCatalog -and $diskCachedCatalog.Count -gt 0) {
            $script:ProjectZomboidVehicleCatalogCache.Clear()
            $script:ProjectZomboidVehicleCatalogCache[$cacheKey] = @($diskCachedCatalog)
            return @($diskCachedCatalog)
        }

        $results = New-Object System.Collections.Generic.List[object]
        $files = @(Get-ChildItem -LiteralPath $scriptsDir -Recurse -File -Include *.txt -ErrorAction SilentlyContinue)

        foreach ($file in $files) {
            $moduleName = 'Base'
            $inVehicle = $false
            $braceDepth = 0
            $vehicleName = ''
            $props = @{}

            foreach ($rawLine in (Get-Content -LiteralPath $file.FullName -ErrorAction SilentlyContinue)) {
                $line = ($rawLine -replace '//.*$', '').Trim()
                if ([string]::IsNullOrWhiteSpace($line)) { continue }

                if (-not $inVehicle) {
                    if ($line -match '^module\s+([A-Za-z0-9_]+)\s*$') {
                        $moduleName = $Matches[1]
                        continue
                    }
                    if ($line -match '^vehicle\s+([A-Za-z0-9_]+)\s*$') {
                        $inVehicle = $true
                        $braceDepth = 0
                        $vehicleName = $Matches[1]
                        $props = @{
                            SourceFile = $file.FullName
                        }
                    }
                    continue
                }

                if ($line -eq '{') {
                    $braceDepth++
                    continue
                }

                if ($line -eq '}') {
                    if ($braceDepth -gt 0) {
                        $braceDepth--
                        if ($braceDepth -eq 0) {
                            $displayName = if ($props.ContainsKey('displayName') -and -not [string]::IsNullOrWhiteSpace($props['displayName'])) {
                                [string]$props['displayName']
                            } else {
                                _FormatProjectZomboidDisplayName -Name $vehicleName
                            }

                            $category = ''
                            try {
                                $category = Split-Path -Path $file.DirectoryName -Leaf
                            } catch { $category = '' }

                            $results.Add([pscustomobject]@{
                                FullType        = "$moduleName.$vehicleName"
                                DisplayName     = $displayName
                                ListText        = "$displayName [$moduleName.$vehicleName]"
                                Module          = $moduleName
                                VehicleName     = $vehicleName
                                DisplayCategory = $category
                                Texture         = if ($props.ContainsKey('texture')) { [string]$props['texture'] } else { '' }
                                SourceFile      = $file.FullName
                            }) | Out-Null

                            $inVehicle = $false
                            $vehicleName = ''
                            $props = @{}
                        }
                    }
                    continue
                }

                if ($braceDepth -eq 1 -and $line -match '^([A-Za-z0-9_]+)\s*=\s*(.+?)(?:,)?$') {
                    $props[$Matches[1]] = $Matches[2].Trim()
                }
            }
        }

        $catalog = @(
            $results |
            Sort-Object @{ Expression = { if ([string]::IsNullOrWhiteSpace($_.DisplayCategory)) { 'zzz_misc' } else { $_.DisplayCategory } } }, DisplayName, FullType -Unique
        )
        $script:ProjectZomboidVehicleCatalogCache.Clear()
        $script:ProjectZomboidVehicleCatalogCache[$cacheKey] = $catalog
        _SaveProjectZomboidCatalogCache -CatalogName 'pz-vehicle-catalog' -CacheKey $cacheKey -Entries $catalog
        return $catalog
    }

    function Get-ProjectZomboidSpawnerCatalogs {
        param([hashtable]$Profile)

        return [pscustomobject]@{
            Items    = @(_GetProjectZomboidItemCatalog -Profile $Profile)
            Vehicles = @(_GetProjectZomboidVehicleCatalog -Profile $Profile)
        }
    }

    function _GetRecentPlayersLogText {
        param(
            [hashtable]$Profile,
            [int]$TailLines = 120
        )

        if ($null -eq $Profile) { return '' }

        $files = @(_ResolveGameLogFiles -Profile $Profile)
        if ($null -eq $files -or $files.Count -eq 0) { return '' }

        $chunks = New-Object 'System.Collections.Generic.List[string]'
        foreach ($path in ($files | Select-Object -Last 2)) {
            if ([string]::IsNullOrWhiteSpace($path) -or -not (Test-Path -LiteralPath $path)) { continue }
            try {
                $lines = @(Get-Content -LiteralPath $path -Tail $TailLines -ErrorAction Stop)
                if ($lines.Count -gt 0) {
                    $chunks.Add(($lines -join [Environment]::NewLine)) | Out-Null
                }
            } catch { }
        }

        return ($chunks -join [Environment]::NewLine)
    }

    function _OpenCommandsWindow {
        param(
            [hashtable]$Profile,
            [string]$Prefix,
            [hashtable]$WindowSharedState = $SharedState
        )

        if ($null -eq $Profile) { return }
        $commandSharedState = $WindowSharedState

        $cmdClrMuted = $clrMuted
        $cmdClrRed   = $clrRed
        $cmdClrGreen = $clrGreen

        $catalog = if ($script:CommandCatalog) { $script:CommandCatalog } else { _LoadCommandCatalog }
        if ($null -eq $catalog) {
            [System.Windows.Forms.MessageBox]::Show(
                'Command catalog not found. Make sure Config\\CommandCatalog.json exists.',
                'Command Catalog Missing','OK','Warning') | Out-Null
            return
        }

        $knownGameName = _GetProfileKnownGame -Profile $Profile
        $gameEntry = _FindCommandCatalogForGame -Catalog $catalog -GameName $knownGameName
        if ($null -eq $gameEntry -or -not $gameEntry.Commands) {
            $gameEntry = _BuildProfileCommandCatalogFallback -Profile $Profile
        }
        if ($null -eq $gameEntry -or -not $gameEntry.Commands) {
            [System.Windows.Forms.MessageBox]::Show(
                "No commands found for '$($Profile.GameName)' in CommandCatalog.json or the profile itself.",
                'No Commands','OK','Information') | Out-Null
            return
        }

        $normalizedCommandGame = (_NormalizeGameIdentity $knownGameName)
        $supportsPzItemSpawner = ($normalizedCommandGame -eq 'projectzomboid')
        $supportsMinecraftItemSpawner = ($normalizedCommandGame -eq 'minecraft')
        $supportsHytaleItemSpawner = ($normalizedCommandGame -eq 'hytale')
        $supportsItemSpawner = ($supportsPzItemSpawner -or $supportsMinecraftItemSpawner -or $supportsHytaleItemSpawner)
        $supportsVehicleSpawner = $supportsPzItemSpawner
        $supportsGiveItemSpawner = ($supportsMinecraftItemSpawner -or $supportsHytaleItemSpawner)
        $itemSpawnerCommandName = if ($supportsGiveItemSpawner) { 'give' } else { 'additem' }
        $itemSpawnerCommandLabel = if ($supportsGiveItemSpawner) { '/give' } else { '/additem' }
        $itemSpawnerLoadHint = if ($supportsMinecraftItemSpawner) {
            'Loading bundled Minecraft item entries...'
        } elseif ($supportsHytaleItemSpawner) {
            'Loading bundled Hytale item entries...'
        } else {
            'Loading local Build 42 item entries...'
        }
        $itemSpawnerMissingHint = if ($supportsMinecraftItemSpawner) {
            'ECC could not load bundled Minecraft item data.'
        } elseif ($supportsHytaleItemSpawner) {
            'ECC could not load bundled Hytale item data.'
        } else {
            'ECC could not load local item data from the Project Zomboid install path.'
        }

        $form                 = New-Object System.Windows.Forms.Form
        $form.Text            = "Commands - $($Profile.GameName)"
        $form.Size            = if ($supportsItemSpawner) { [System.Drawing.Size]::new(1380, 760) } else { [System.Drawing.Size]::new(900, 650) }
        $form.MinimumSize     = if ($supportsItemSpawner) { [System.Drawing.Size]::new(1260, 760) } else { [System.Drawing.Size]::new(780, 560) }
        $form.StartPosition   = 'CenterParent'
        $form.BackColor       = $clrBg
        $form.FormBorderStyle = 'Sizable'

        $lblHeader            = _Label "Commands - $($Profile.GameName)" 10 10 600 22 $fontBold
        $lblHeader.Anchor     = 'Top,Left,Right'
        $form.Controls.Add($lblHeader)

        $commandsPanelGap = if ($supportsItemSpawner) { 6 } else { 10 }
        $rightPanelWidth = if ($supportsItemSpawner) { 500 } else { 0 }
        $commandsFooterHeight = 212
        $commandsFooterBottomMargin = 10
        $commandsContentHeight = [Math]::Max(180, $form.ClientSize.Height - $commandsFooterHeight - $commandsFooterBottomMargin - 50)
        $listPanelWidth = if ($supportsItemSpawner) {
            [Math]::Max(320, $form.ClientSize.Width - $rightPanelWidth - 20 - $commandsPanelGap)
        } else {
            $form.ClientSize.Width - 20
        }

        $toolTip = New-Object System.Windows.Forms.ToolTip
        $toolTip.AutoPopDelay = 12000
        $toolTip.InitialDelay = 500
        $toolTip.ReshowDelay  = 200
        $toolTip.ShowAlways   = $true

        # Scrollable command list
        $listPanel            = New-Object System.Windows.Forms.Panel
        $listPanel.Location   = [System.Drawing.Point]::new(10, 40)
        $listPanel.Size       = [System.Drawing.Size]::new($listPanelWidth, $commandsContentHeight)
        $listPanel.Anchor     = 'Top,Left,Right,Bottom'
        $listPanel.AutoScroll = $true
        $listPanel.BackColor  = $clrPanel
        $form.Controls.Add($listPanel)

        $spawnPanel = $null
        $cmbSpawnPlayer = $null
        $btnRefreshPlayers = $null
        $numSpawnCount = $null
        $tbItemSearch = $null
        $lbItems = $null
        $btnInsertAddItem = $null
        $picItemPreview = $null
        $picVehiclePreview = $null
        $lblItemName = $null
        $lblItemType = $null
        $lblSpawnHint = $null
        $btnBuildAddItem = $null
        $spawnFooterPanel = $null
        $spawnTabControl = $null
        $tabItems = $null
        $tabVehicles = $null
        $vehicleFooterPanel = $null
        $tbVehicleSearch = $null
        $lbVehicles = $null
        $lblVehicleName = $null
        $lblVehicleType = $null
        $lblVehicleHint = $null
        $btnInsertAddVehicle = $null
        $btnBuildAddVehicle = $null
        $pzCatalogState = @{
            ItemSource = @()
            VehicleSource = @()
        }
        $pzItemLookup = @{}
        $pzVehicleLookup = @{}
        $spawnPreviewColumnWidth = 96
        $spawnFooterHeight = 138
        $spawnActionButtonY = 100
        $pzCatalogWorker = $null
        $pzCatalogRunspace = $null
        $pzCatalogAsync = $null
        $pzCatalogTimer = $null
        $pzCatalogLoaded = $false

        if ($supportsItemSpawner) {
            $spawnPanel = New-Object System.Windows.Forms.Panel
            $spawnPanel.Location = [System.Drawing.Point]::new($listPanel.Right + $commandsPanelGap, 40)
            $spawnPanel.Size = [System.Drawing.Size]::new($rightPanelWidth, $commandsContentHeight)
            $spawnPanel.Anchor = 'Top,Right,Bottom'
            $spawnPanel.BackColor = $clrPanel
            $form.Controls.Add($spawnPanel)

            $lblSpawnHeader = _Label 'Item Spawner' 12 12 220 20 $fontBold
            $spawnPanel.Controls.Add($lblSpawnHeader)

            $lblSpawnPlayer = _Label 'Online Player' 12 42 120 18 $fontLabel
            $spawnPanel.Controls.Add($lblSpawnPlayer)

            $cmbSpawnPlayer = New-Object System.Windows.Forms.ComboBox
            $cmbSpawnPlayer.Location = [System.Drawing.Point]::new(12, 62)
            $cmbSpawnPlayer.Size = [System.Drawing.Size]::new(250, 24)
            $cmbSpawnPlayer.Anchor = 'Top,Left,Right'
            $cmbSpawnPlayer.DropDownStyle = 'DropDown'
            $cmbSpawnPlayer.BackColor = [System.Drawing.Color]::FromArgb(30,30,40)
            $cmbSpawnPlayer.ForeColor = $clrText
            $cmbSpawnPlayer.Font = $fontLabel
            $spawnPanel.Controls.Add($cmbSpawnPlayer)

            $btnRefreshPlayers = _Button 'Refresh' 270 61 96 26 $clrPanelAlt $null
            $btnRefreshPlayers.Anchor = 'Top,Right'
            $spawnPanel.Controls.Add($btnRefreshPlayers)

            $spawnTabControl = New-Object System.Windows.Forms.TabControl
            $spawnTabControl.Location = [System.Drawing.Point]::new(12, 96)
            $spawnTabControl.Size = [System.Drawing.Size]::new($spawnPanel.ClientSize.Width - 24, $spawnPanel.ClientSize.Height - 108)
            $spawnTabControl.Anchor = 'Top,Left,Right,Bottom'
            $spawnPanel.Controls.Add($spawnTabControl)

            $tabItems = New-Object System.Windows.Forms.TabPage
            $tabItems.Text = 'Items'
            $tabItems.BackColor = $clrPanel
            $spawnTabControl.TabPages.Add($tabItems) | Out-Null

            if ($supportsVehicleSpawner) {
                $tabVehicles = New-Object System.Windows.Forms.TabPage
                $tabVehicles.Text = 'Vehicles'
                $tabVehicles.BackColor = $clrPanel
                $spawnTabControl.TabPages.Add($tabVehicles) | Out-Null
            }

            $itemsLayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
            $itemsLayoutPanel.Dock = 'Fill'
            $itemsLayoutPanel.Padding = [System.Windows.Forms.Padding]::new(12)
            $itemsLayoutPanel.BackColor = $clrPanel
            $itemsLayoutPanel.ColumnCount = 1
            $itemsLayoutPanel.RowCount = 3
            $itemsLayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
            $itemsLayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 88))) | Out-Null
            $itemsLayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
            $itemsLayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, $spawnFooterHeight))) | Out-Null
            $tabItems.Controls.Add($itemsLayoutPanel)

            $spawnFooterPanel = New-Object System.Windows.Forms.Panel
            $spawnFooterPanel.Dock = 'Fill'
            $spawnFooterPanel.BackColor = [System.Drawing.Color]::FromArgb(28,32,46)
            $itemsLayoutPanel.Controls.Add($spawnFooterPanel, 0, 2)

            $itemsSearchPanel = New-Object System.Windows.Forms.Panel
            $itemsSearchPanel.Dock = 'Fill'
            $itemsSearchPanel.BackColor = $clrPanel
            $itemsLayoutPanel.Controls.Add($itemsSearchPanel, 0, 0)

            $lblSpawnCount = _Label 'Count' 0 0 80 18 $fontLabel
            $itemsSearchPanel.Controls.Add($lblSpawnCount)

            $numSpawnCount = New-Object System.Windows.Forms.NumericUpDown
            $numSpawnCount.Location = [System.Drawing.Point]::new(0, 20)
            $numSpawnCount.Size = [System.Drawing.Size]::new(110, 24)
            $numSpawnCount.Minimum = 1
            $numSpawnCount.Maximum = 10000
            $numSpawnCount.Value = 1
            $numSpawnCount.BackColor = [System.Drawing.Color]::FromArgb(30,30,40)
            $numSpawnCount.ForeColor = $clrText
            $numSpawnCount.Font = $fontLabel
            $itemsSearchPanel.Controls.Add($numSpawnCount)

            $lblSearch = _Label 'Search Items' 0 42 120 18 $fontLabel
            $itemsSearchPanel.Controls.Add($lblSearch)

            $tbItemSearch = New-Object System.Windows.Forms.TextBox
            $tbItemSearch.Location = [System.Drawing.Point]::new(0, 60)
            $tbItemSearch.Multiline = $true
            $tbItemSearch.Size = [System.Drawing.Size]::new(332, 28)
            $tbItemSearch.Anchor = 'Top,Left,Right'
            $tbItemSearch.BackColor = [System.Drawing.Color]::FromArgb(30,30,40)
            $tbItemSearch.ForeColor = $clrText
            $tbItemSearch.BorderStyle = 'FixedSingle'
            $tbItemSearch.Font = $fontLabel
            $itemsSearchPanel.Controls.Add($tbItemSearch)

            $lbItems = New-Object System.Windows.Forms.ListBox
            $lbItems.Dock = 'Fill'
            $lbItems.BackColor = [System.Drawing.Color]::FromArgb(24,24,34)
            $lbItems.ForeColor = $clrText
            $lbItems.BorderStyle = 'FixedSingle'
            $lbItems.Font = $fontMono
            $lbItems.IntegralHeight = $false
            $itemsLayoutPanel.Controls.Add($lbItems, 0, 1)

            $picItemPreview = New-Object System.Windows.Forms.PictureBox
            $picItemPreview.Location = [System.Drawing.Point]::new($spawnFooterPanel.ClientSize.Width - $spawnPreviewColumnWidth + 8, 8)
            $picItemPreview.Size = [System.Drawing.Size]::new($spawnPreviewColumnWidth - 16, 70)
            $picItemPreview.Anchor = 'Top,Right'
            $picItemPreview.BackColor = [System.Drawing.Color]::FromArgb(18,18,28)
            $picItemPreview.BorderStyle = 'FixedSingle'
            $picItemPreview.SizeMode = 'Zoom'
            $picItemPreview.Visible = $false
            $spawnFooterPanel.Controls.Add($picItemPreview)

            $itemTextWidth = [Math]::Max(140, $spawnFooterPanel.ClientSize.Width - $spawnPreviewColumnWidth - 8)

            $lblItemName = _Label 'Select an item' 0 4 $itemTextWidth 20 $fontBold
            $lblItemName.Anchor = 'Left,Right,Top'
            $lblItemName.AutoEllipsis = $true
            $spawnFooterPanel.Controls.Add($lblItemName)

            $lblItemType = _Label 'No item selected yet.' 0 24 $itemTextWidth 30 $fontMono
            $lblItemType.Anchor = 'Left,Right,Top'
            $lblItemType.ForeColor = $clrMuted
            $lblItemType.AutoEllipsis = $true
            $spawnFooterPanel.Controls.Add($lblItemType)

            $lblSpawnHint = _Label '' 0 52 $itemTextWidth 40 $fontLabel
            $lblSpawnHint.Anchor = 'Left,Right,Top'
            $lblSpawnHint.ForeColor = $clrMuted
            $lblSpawnHint.Text = $itemSpawnerLoadHint
            $spawnFooterPanel.Controls.Add($lblSpawnHint)

            $itemActionsRow = New-Object System.Windows.Forms.FlowLayoutPanel
            $itemActionsRow.Location = [System.Drawing.Point]::new(0, $spawnActionButtonY)
            $itemActionsRow.Size = [System.Drawing.Size]::new([Math]::Max(180, $spawnFooterPanel.ClientSize.Width - $spawnPreviewColumnWidth - 8), 30)
            $itemActionsRow.Anchor = 'Left,Right,Bottom'
            $itemActionsRow.WrapContents = $false
            $itemActionsRow.AutoScroll = $true
            $itemActionsRow.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
            $itemActionsRow.BackColor = [System.Drawing.Color]::Transparent
            $spawnFooterPanel.Controls.Add($itemActionsRow)

            $btnInsertAddItem = _Button 'Insert Into Command Box' 0 0 170 28 $clrGreen $null
            $btnInsertAddItem.Margin = [System.Windows.Forms.Padding]::new(0, 0, 12, 0)
            $itemActionsRow.Controls.Add($btnInsertAddItem)

            $btnBuildAddItem = _Button ("Build {0}" -f $itemSpawnerCommandLabel) 0 0 150 28 $clrAccent $null
            $btnBuildAddItem.Margin = [System.Windows.Forms.Padding]::new(0)
            $itemActionsRow.Controls.Add($btnBuildAddItem)

            if ($supportsVehicleSpawner) {
                $vehiclesLayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
                $vehiclesLayoutPanel.Dock = 'Fill'
                $vehiclesLayoutPanel.Padding = [System.Windows.Forms.Padding]::new(12)
                $vehiclesLayoutPanel.BackColor = $clrPanel
                $vehiclesLayoutPanel.ColumnCount = 1
                $vehiclesLayoutPanel.RowCount = 3
                $vehiclesLayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
                $vehiclesLayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 50))) | Out-Null
                $vehiclesLayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
                $vehiclesLayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, $spawnFooterHeight))) | Out-Null
                $tabVehicles.Controls.Add($vehiclesLayoutPanel)

                $vehicleFooterPanel = New-Object System.Windows.Forms.Panel
                $vehicleFooterPanel.Dock = 'Fill'
                $vehicleFooterPanel.BackColor = [System.Drawing.Color]::FromArgb(28,32,46)
                $vehiclesLayoutPanel.Controls.Add($vehicleFooterPanel, 0, 2)

                $vehiclesSearchPanel = New-Object System.Windows.Forms.Panel
                $vehiclesSearchPanel.Dock = 'Fill'
                $vehiclesSearchPanel.BackColor = $clrPanel
                $vehiclesLayoutPanel.Controls.Add($vehiclesSearchPanel, 0, 0)

                $lblVehicleSearch = _Label 'Search Vehicles' 0 0 140 18 $fontLabel
                $vehiclesSearchPanel.Controls.Add($lblVehicleSearch)

                $tbVehicleSearch = New-Object System.Windows.Forms.TextBox
                $tbVehicleSearch.Location = [System.Drawing.Point]::new(0, 20)
                $tbVehicleSearch.Multiline = $true
                $tbVehicleSearch.Size = [System.Drawing.Size]::new(332, 28)
                $tbVehicleSearch.Anchor = 'Top,Left,Right'
                $tbVehicleSearch.BackColor = [System.Drawing.Color]::FromArgb(30,30,40)
                $tbVehicleSearch.ForeColor = $clrText
                $tbVehicleSearch.BorderStyle = 'FixedSingle'
                $tbVehicleSearch.Font = $fontLabel
                $vehiclesSearchPanel.Controls.Add($tbVehicleSearch)

                $lbVehicles = New-Object System.Windows.Forms.ListBox
                $lbVehicles.Dock = 'Fill'
                $lbVehicles.BackColor = [System.Drawing.Color]::FromArgb(24,24,34)
                $lbVehicles.ForeColor = $clrText
                $lbVehicles.BorderStyle = 'FixedSingle'
                $lbVehicles.Font = $fontMono
                $lbVehicles.IntegralHeight = $false
                $vehiclesLayoutPanel.Controls.Add($lbVehicles, 0, 1)

                $picVehiclePreview = New-Object System.Windows.Forms.PictureBox
                $picVehiclePreview.Location = [System.Drawing.Point]::new($vehicleFooterPanel.ClientSize.Width - $spawnPreviewColumnWidth + 8, 8)
                $picVehiclePreview.Size = [System.Drawing.Size]::new($spawnPreviewColumnWidth - 16, 70)
                $picVehiclePreview.Anchor = 'Top,Right'
                $picVehiclePreview.BackColor = [System.Drawing.Color]::FromArgb(18,18,28)
                $picVehiclePreview.BorderStyle = 'FixedSingle'
                $picVehiclePreview.SizeMode = 'Zoom'
                $picVehiclePreview.Visible = $false
                $vehicleFooterPanel.Controls.Add($picVehiclePreview)

                $vehicleTextWidth = [Math]::Max(140, $vehicleFooterPanel.ClientSize.Width - $spawnPreviewColumnWidth - 8)

                $lblVehicleName = _Label 'Select a vehicle' 0 4 $vehicleTextWidth 20 $fontBold
                $lblVehicleName.Anchor = 'Left,Right,Top'
                $lblVehicleName.AutoEllipsis = $true
                $vehicleFooterPanel.Controls.Add($lblVehicleName)

                $lblVehicleType = _Label 'No vehicle selected yet.' 0 24 $vehicleTextWidth 30 $fontMono
                $lblVehicleType.Anchor = 'Left,Right,Top'
                $lblVehicleType.ForeColor = $clrMuted
                $lblVehicleType.AutoEllipsis = $true
                $vehicleFooterPanel.Controls.Add($lblVehicleType)

                $lblVehicleHint = _Label '' 0 52 $vehicleTextWidth 40 $fontLabel
                $lblVehicleHint.Anchor = 'Left,Right,Top'
                $lblVehicleHint.ForeColor = $clrMuted
                $lblVehicleHint.Text = 'Loading local vehicle entries...'
                $vehicleFooterPanel.Controls.Add($lblVehicleHint)

                $vehicleActionsRow = New-Object System.Windows.Forms.FlowLayoutPanel
                $vehicleActionsRow.Location = [System.Drawing.Point]::new(0, $spawnActionButtonY)
                $vehicleActionsRow.Size = [System.Drawing.Size]::new([Math]::Max(180, $vehicleFooterPanel.ClientSize.Width - $spawnPreviewColumnWidth - 8), 30)
                $vehicleActionsRow.Anchor = 'Left,Right,Bottom'
                $vehicleActionsRow.WrapContents = $false
                $vehicleActionsRow.AutoScroll = $true
                $vehicleActionsRow.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
                $vehicleActionsRow.BackColor = [System.Drawing.Color]::Transparent
                $vehicleFooterPanel.Controls.Add($vehicleActionsRow)

                $btnInsertAddVehicle = _Button 'Insert Into Command Box' 0 0 170 28 $clrGreen $null
                $btnInsertAddVehicle.Margin = [System.Windows.Forms.Padding]::new(0, 0, 12, 0)
                $vehicleActionsRow.Controls.Add($btnInsertAddVehicle)

                $btnBuildAddVehicle = _Button 'Build /addvehicle' 0 0 150 28 $clrAccent $null
                $btnBuildAddVehicle.Margin = [System.Windows.Forms.Padding]::new(0)
                $vehicleActionsRow.Controls.Add($btnBuildAddVehicle)
            }

            $spawnFooterPanel.BringToFront()
            if ($vehicleFooterPanel) { $vehicleFooterPanel.BringToFront() }
        }

        $commandFooterPanel = New-Object System.Windows.Forms.Panel
        $commandFooterPanel.Location = [System.Drawing.Point]::new(10, $form.ClientSize.Height - $commandsFooterHeight - $commandsFooterBottomMargin)
        $commandFooterPanel.Size = [System.Drawing.Size]::new($form.ClientSize.Width - 20, $commandsFooterHeight)
        $commandFooterPanel.Anchor = 'Left,Right,Bottom'
        $commandFooterPanel.BackColor = $clrBg
        $form.Controls.Add($commandFooterPanel)

        $commandInputRow = New-Object System.Windows.Forms.Panel
        $commandInputRow.Location = [System.Drawing.Point]::new(0, 154)
        $commandInputRow.Size = [System.Drawing.Size]::new($commandFooterPanel.ClientSize.Width, 32)
        $commandInputRow.Anchor = 'Top,Left,Right'
        $commandInputRow.BackColor = [System.Drawing.Color]::Transparent
        $commandFooterPanel.Controls.Add($commandInputRow)

        # Bottom command area
        $lblCmd = _Label 'Command' 0 0 $commandFooterPanel.ClientSize.Width 18 $fontBold
        $lblCmd.Anchor = 'Top,Left,Right'
        $commandFooterPanel.Controls.Add($lblCmd)

        $tbDebug = New-Object System.Windows.Forms.TextBox
        $tbDebug.Location    = [System.Drawing.Point]::new(0, 24)
        $tbDebug.Size        = [System.Drawing.Size]::new($commandFooterPanel.ClientSize.Width, 70)
        $tbDebug.Anchor      = 'Top,Left,Right'
        $tbDebug.BackColor   = [System.Drawing.Color]::FromArgb(24,24,34)
        $tbDebug.ForeColor   = $clrMuted
        $tbDebug.BorderStyle = 'FixedSingle'
        $tbDebug.Font        = $fontMono
        $tbDebug.Multiline   = $true
        $tbDebug.ReadOnly    = $true
        $tbDebug.ScrollBars  = 'Vertical'
        $commandFooterPanel.Controls.Add($tbDebug)

        $parseOnlinePlayersText = {
            param([string]$Text)

            if ([string]::IsNullOrWhiteSpace($Text)) { return @() }

            $names = New-Object 'System.Collections.Generic.List[string]'
            $seen = @{}

            $quotedMatches = [regex]::Matches($Text, '"([^"]+)"')
            foreach ($match in $quotedMatches) {
                $name = $match.Groups[1].Value.Trim()
                if ([string]::IsNullOrWhiteSpace($name)) { continue }
                $key = $name.ToLowerInvariant()
                if (-not $seen.ContainsKey($key)) {
                    $seen[$key] = $true
                    $names.Add($name) | Out-Null
                }
            }
            if ($names.Count -gt 0) { return @($names) }

            $lines = @($Text -split "(`r`n|`n|`r)")
            $captureDashNames = $false
            foreach ($lineRaw in $lines) {
                $line = $lineRaw.Trim()
                if ([string]::IsNullOrWhiteSpace($line)) { continue }

                if ($captureDashNames) {
                    if ($line -match '^\s*-\s*(.+?)\s*$') {
                        $name = $Matches[1].Trim()
                        if (-not [string]::IsNullOrWhiteSpace($name)) {
                            $key = $name.ToLowerInvariant()
                            if (-not $seen.ContainsKey($key)) {
                                $seen[$key] = $true
                                $names.Add($name) | Out-Null
                            }
                        }
                        continue
                    }

                    if ($line -notmatch '^\[' -and $line -notmatch '^(?i:players?\s+)') {
                        $captureDashNames = $false
                    }
                }

                if ($line -match '(?i)players\s+connected\s*\(\d+\)\s*:\s*$') {
                    $captureDashNames = $true
                    continue
                }

                if ($line -match '(?i)players?.*?:\s*(.+)$') {
                    foreach ($part in ($Matches[1] -split ',')) {
                        $name = $part.Trim(" `t[](){}")
                        if ([string]::IsNullOrWhiteSpace($name)) { continue }
                        $key = $name.ToLowerInvariant()
                        if (-not $seen.ContainsKey($key)) {
                            $seen[$key] = $true
                            $names.Add($name) | Out-Null
                        }
                    }
                }
            }

            return @($names)
        }.GetNewClosure()

        $getGameLogTabTextLocal = {
            param([string]$TargetPrefix)

            if ([string]::IsNullOrWhiteSpace($TargetPrefix)) { return '' }
            if (-not $script:_GameLogTabs -or -not $script:_GameLogTabs.ContainsKey($TargetPrefix)) { return '' }

            try {
                $tabEntry = $script:_GameLogTabs[$TargetPrefix]
                $chunks = New-Object 'System.Collections.Generic.List[string]'
                if ($tabEntry -and $tabEntry.RTB -and -not [string]::IsNullOrWhiteSpace([string]$tabEntry.RTB.Text)) {
                    $chunks.Add([string]$tabEntry.RTB.Text) | Out-Null
                }
                if ($tabEntry -and $tabEntry.PendingLines -and $tabEntry.PendingLines.Count -gt 0) {
                    $chunks.Add((@($tabEntry.PendingLines) -join [Environment]::NewLine)) | Out-Null
                }
                if ($chunks.Count -gt 0) {
                    return ($chunks -join [Environment]::NewLine)
                }
            } catch { }

            return ''
        }.GetNewClosure()

        $getRecentPlayersLogTextLocal = {
            param(
                [hashtable]$TargetProfile,
                [int]$TailLines = 120
            )

            if ($null -eq $TargetProfile) { return '' }

            $files = New-Object 'System.Collections.Generic.List[string]'
            try {
                foreach ($path in @(_ResolveGameLogFiles -Profile $TargetProfile)) {
                    if ($path -and -not $files.Contains($path)) {
                        $files.Add($path) | Out-Null
                    }
                }
            } catch { }

            if ($files.Count -eq 0) { return '' }

            $chunks = New-Object 'System.Collections.Generic.List[string]'
            foreach ($path in ($files | Select-Object -Last 2)) {
                if ([string]::IsNullOrWhiteSpace($path) -or -not (Test-Path -LiteralPath $path)) { continue }
                try {
                    $lines = @(Get-Content -LiteralPath $path -Tail $TailLines -ErrorAction Stop)
                    if ($lines.Count -gt 0) {
                        $chunks.Add(($lines -join [Environment]::NewLine)) | Out-Null
                    }
                } catch { }
            }

            return ($chunks -join [Environment]::NewLine)
        }.GetNewClosure()

        $tbCmd = New-Object System.Windows.Forms.TextBox
        $tbCmd.Location    = [System.Drawing.Point]::new(0, 2)
        $tbCmd.Size        = [System.Drawing.Size]::new($commandInputRow.ClientSize.Width - 240, 26)
        $tbCmd.Anchor      = 'Top,Left'
        $tbCmd.BackColor   = [System.Drawing.Color]::FromArgb(30,30,40)
        $tbCmd.ForeColor   = $clrText
        $tbCmd.BorderStyle = 'FixedSingle'
        $tbCmd.Font        = $fontMono
        $commandInputRow.Controls.Add($tbCmd)

        $defaultVerbose = $false
        try {
            if ($script:SharedState -and $script:SharedState.Settings -and $script:SharedState.Settings.ContainsKey('EnableDebugLogging')) {
                $defaultVerbose = [bool]$script:SharedState.Settings.EnableDebugLogging
            }
        } catch { }

        $chkVerbose = New-Object System.Windows.Forms.CheckBox
        $chkVerbose.Text      = 'Verbose Debug'
        $chkVerbose.Location  = [System.Drawing.Point]::new($form.ClientSize.Width - 450, $form.ClientSize.Height - 82)
        $chkVerbose.Size      = [System.Drawing.Size]::new(120, 20)
        $chkVerbose.Anchor    = 'Right,Bottom'
        $chkVerbose.ForeColor = $clrText
        $chkVerbose.BackColor = [System.Drawing.Color]::Transparent
        $chkVerbose.Font      = $fontLabel
        $chkVerbose.Checked   = $defaultVerbose
        $chkVerbose.Enabled   = $defaultVerbose
        $form.Controls.Add($chkVerbose)

        $isGlobalCommandsDebugEnabled = {
            try {
                if ($script:SharedState -and $script:SharedState.Settings -and $script:SharedState.Settings.ContainsKey('EnableDebugLogging')) {
                    return ([bool]$script:SharedState.Settings.EnableDebugLogging)
                }
            } catch { }
            return $false
        }.GetNewClosure()

        $isCommandsVerboseEnabled = {
            return ((& $isGlobalCommandsDebugEnabled) -and ($chkVerbose -is [System.Windows.Forms.CheckBox]) -and ($chkVerbose.Checked -eq $true))
        }.GetNewClosure()

        $writeCommandsDebugLine = {
            param([string]$Text)

            if ([string]::IsNullOrWhiteSpace($Text)) { return }
            if (-not (& $isCommandsVerboseEnabled)) { return }

            try {
                $ts = Get-Date -Format 'HH:mm:ss'
                $line = "[$ts][PZDBG] $Text"
                if ($tbDebug -is [System.Windows.Forms.TextBoxBase]) {
                    $tbDebug.AppendText($line + [Environment]::NewLine)
                }
                if ($commandSharedState -and $commandSharedState.LogQueue) {
                    $commandSharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][DEBUG][PZREFRESH] $Text")
                }
            } catch { }
        }.GetNewClosure()

        $chkApi = New-Object System.Windows.Forms.CheckBox
        $chkApi.Text      = 'Use API'
        $chkApi.Location  = [System.Drawing.Point]::new($form.ClientSize.Width - 330, $form.ClientSize.Height - 82)
        $chkApi.Size      = [System.Drawing.Size]::new(90, 20)
        $chkApi.Anchor    = 'Right,Bottom'
        $chkApi.ForeColor = $clrText
        $chkApi.BackColor = [System.Drawing.Color]::Transparent
        $chkApi.Font      = $fontLabel
        $chkApi.Enabled   = ($Profile.SatisfactoryApiPort -and [int]$Profile.SatisfactoryApiPort -gt 0)
        $form.Controls.Add($chkApi)

        $chkRest = New-Object System.Windows.Forms.CheckBox
        $chkRest.Text      = 'Use REST'
        $chkRest.Location  = [System.Drawing.Point]::new($form.ClientSize.Width - 230, $form.ClientSize.Height - 82)
        $chkRest.Size      = [System.Drawing.Size]::new(90, 20)
        $chkRest.Anchor    = 'Right,Bottom'
        $chkRest.ForeColor = $clrText
        $chkRest.BackColor = [System.Drawing.Color]::Transparent
        $chkRest.Font      = $fontLabel
        $chkRest.Enabled   = ($Profile.RestEnabled -eq $true -or ($Profile.RestPort -and [int]$Profile.RestPort -gt 0))
        $form.Controls.Add($chkRest)

        $btnTestTelnet = _Button 'Test Telnet' ($form.ClientSize.Width - 560) ($form.ClientSize.Height - 112) 100 28 $clrPanel $null
        $btnTestTelnet.Anchor = 'Right,Bottom'
        $btnTestTelnet.Enabled = ($Profile.TelnetPassword -and $Profile.TelnetPort -and [int]$Profile.TelnetPort -gt 0)
        $form.Controls.Add($btnTestTelnet)

        $btnTestRcon = _Button 'Test RCON' ($form.ClientSize.Width - 450) ($form.ClientSize.Height - 112) 90 28 $clrPanel $null
        $btnTestRcon.Anchor = 'Right,Bottom'
        $btnTestRcon.Enabled = ($Profile.RconPassword -and $Profile.RconPort -and [int]$Profile.RconPort -gt 0)
        $form.Controls.Add($btnTestRcon)

        $btnTestApi = _Button 'Test API' ($form.ClientSize.Width - 350) ($form.ClientSize.Height - 112) 80 28 $clrPanel $null
        $btnTestApi.Anchor = 'Right,Bottom'
        $btnTestApi.Enabled = ($Profile.SatisfactoryApiPort -and [int]$Profile.SatisfactoryApiPort -gt 0)
        $form.Controls.Add($btnTestApi)

        $btnTestStdin = _Button 'Test STDIN' ($form.ClientSize.Width - 260) ($form.ClientSize.Height - 112) 90 28 $clrPanel $null
        $btnTestStdin.Anchor = 'Right,Bottom'
        $btnTestStdin.Enabled = $true
        $form.Controls.Add($btnTestStdin)

        $btnTestPid = _Button 'Test PID' ($form.ClientSize.Width - 160) ($form.ClientSize.Height - 112) 80 28 $clrPanel $null
        $btnTestPid.Anchor = 'Right,Bottom'
        $btnTestPid.Enabled = $true
        $form.Controls.Add($btnTestPid)

        $btnTestRest = _Button 'Test REST' 0 0 90 30 $clrPanel $null
        $btnTestRest.Anchor = 'Right,Bottom'
        $btnTestRest.Enabled = ($Profile.RestEnabled -eq $true -or ($Profile.RestPort -and [int]$Profile.RestPort -gt 0))
        $commandInputRow.Controls.Add($btnTestRest)

        $btnSend = _Button 'Send' 0 0 110 30 $clrGreen $null
        $btnSend.Anchor = 'Right,Bottom'
        $commandInputRow.Controls.Add($btnSend)

        foreach ($bottomControl in @($btnTestTelnet, $btnTestRcon, $btnTestApi, $btnTestStdin, $btnTestPid, $btnTestRest, $btnSend, $chkVerbose, $chkApi, $chkRest)) {
            if ($bottomControl -is [System.Windows.Forms.Control]) {
                $bottomControl.Anchor = 'Left,Top'
                $bottomControl.Margin = [System.Windows.Forms.Padding]::new(0)
            }
        }

        if ($btnTestTelnet -is [System.Windows.Forms.Control]) { $btnTestTelnet.Margin = [System.Windows.Forms.Padding]::new(0, 0, 10, 0) }
        if ($btnTestRcon   -is [System.Windows.Forms.Control]) { $btnTestRcon.Margin   = [System.Windows.Forms.Padding]::new(0, 0, 10, 0) }
        if ($btnTestApi    -is [System.Windows.Forms.Control]) { $btnTestApi.Margin    = [System.Windows.Forms.Padding]::new(0, 0, 10, 0) }
        if ($btnTestStdin  -is [System.Windows.Forms.Control]) { $btnTestStdin.Margin  = [System.Windows.Forms.Padding]::new(0, 0, 10, 0) }
        if ($chkVerbose    -is [System.Windows.Forms.Control]) { $chkVerbose.Margin    = [System.Windows.Forms.Padding]::new(0, 0, 12, 0) }
        if ($chkApi        -is [System.Windows.Forms.Control]) { $chkApi.Margin        = [System.Windows.Forms.Padding]::new(0, 0, 12, 0) }
        if ($btnTestRest   -is [System.Windows.Forms.Control]) { $btnTestRest.Margin   = [System.Windows.Forms.Padding]::new(0, 0, 10, 0) }

        $diagnosticsRow = New-Object System.Windows.Forms.FlowLayoutPanel
        $diagnosticsRow.AutoSize = $true
        $diagnosticsRow.AutoSizeMode = 'GrowAndShrink'
        $diagnosticsRow.WrapContents = $false
        $diagnosticsRow.FlowDirection = 'LeftToRight'
        $diagnosticsRow.Padding = [System.Windows.Forms.Padding]::new(0)
        $diagnosticsRow.Margin = [System.Windows.Forms.Padding]::new(0)
        $diagnosticsRow.BackColor = [System.Drawing.Color]::Transparent
        $diagnosticsRow.Anchor = 'Top,Right'
        foreach ($diagControl in @($btnTestTelnet, $btnTestRcon, $btnTestApi, $btnTestStdin, $btnTestPid)) {
            if ($diagControl -is [System.Windows.Forms.Control]) {
                [void]$diagnosticsRow.Controls.Add($diagControl)
            }
        }
        $commandFooterPanel.Controls.Add($diagnosticsRow)

        $optionsRow = New-Object System.Windows.Forms.FlowLayoutPanel
        $optionsRow.AutoSize = $true
        $optionsRow.AutoSizeMode = 'GrowAndShrink'
        $optionsRow.WrapContents = $false
        $optionsRow.FlowDirection = 'LeftToRight'
        $optionsRow.Padding = [System.Windows.Forms.Padding]::new(0)
        $optionsRow.Margin = [System.Windows.Forms.Padding]::new(0)
        $optionsRow.BackColor = [System.Drawing.Color]::Transparent
        $optionsRow.Anchor = 'Top,Right'
        foreach ($optionControl in @($chkVerbose, $chkApi, $chkRest)) {
            if ($optionControl -is [System.Windows.Forms.Control]) {
                [void]$optionsRow.Controls.Add($optionControl)
            }
        }
        $commandFooterPanel.Controls.Add($optionsRow)

        $commandActionsRow = New-Object System.Windows.Forms.FlowLayoutPanel
        $commandActionsRow.AutoSize = $true
        $commandActionsRow.AutoSizeMode = 'GrowAndShrink'
        $commandActionsRow.WrapContents = $false
        $commandActionsRow.FlowDirection = 'LeftToRight'
        $commandActionsRow.Padding = [System.Windows.Forms.Padding]::new(0)
        $commandActionsRow.Margin = [System.Windows.Forms.Padding]::new(0)
        $commandActionsRow.BackColor = [System.Drawing.Color]::Transparent
        $commandActionsRow.Anchor = 'Top,Right'
        foreach ($actionControl in @($btnTestRest, $btnSend)) {
            if ($actionControl -is [System.Windows.Forms.Control]) {
                [void]$commandActionsRow.Controls.Add($actionControl)
            }
        }
        $commandInputRow.Controls.Add($commandActionsRow)
        $diagnosticsRow.BringToFront()
        $optionsRow.BringToFront()
        $commandActionsRow.BringToFront()

        $lblStatus = _Label '' 0 190 $commandFooterPanel.ClientSize.Width 18
        $lblStatus.Anchor = 'Top,Left,Right'
        $lblStatus.ForeColor = $clrMuted
        $commandFooterPanel.Controls.Add($lblStatus)

        # Async command runner to avoid UI freezes
        $pending = New-Object System.Collections.ArrayList
        $timer = New-Object System.Windows.Forms.Timer
        $timer.Interval = 200
        $logQueuedCommandCallbackFailure = {
            param([string]$FailureMessage)

            $ts = Get-Date -Format 'HH:mm:ss'
            try { $tbDebug.AppendText("[$ts][WARN] $FailureMessage" + [Environment]::NewLine) } catch { }
            try {
                if ($script:SharedState -and $script:SharedState.LogQueue) {
                    $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][WARN][GUI] $FailureMessage")
                }
            } catch { }
        }.GetNewClosure()
        $handleQueuedCommandResult = {
            param([object]$result)

            if ($null -eq $result) {
                $lblStatus.Text = "Failed to send: no result returned."
                $lblStatus.ForeColor = $cmdClrRed
                $ts = Get-Date -Format 'HH:mm:ss'
                $tbDebug.AppendText("[$ts] Failed to send: no result returned." + [Environment]::NewLine)
                return
            }

            $lblStatus.Text = $result.Message
            $lblStatus.ForeColor = if ($result.Success) { $cmdClrGreen } else { $cmdClrRed }

            $ts = Get-Date -Format 'HH:mm:ss'
            $line = "[$ts] $($result.Message)"
            $tbDebug.AppendText($line + [Environment]::NewLine)
            if ((& $isCommandsVerboseEnabled) -and $result.Debug -and $result.Debug.Count -gt 0) {
                foreach ($d in $result.Debug) {
                    $dbgLine = "[$ts][DBG] $d"
                    $tbDebug.AppendText($dbgLine + [Environment]::NewLine)
                    try {
                        if ($script:SharedState -and $script:SharedState.LogQueue) {
                            $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][DEBUG][UI] $dbgLine")
                        }
                    } catch { }
                }
            }
            try {
                if ($script:SharedState -and $script:SharedState.LogQueue) {
                    $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][INFO][UI] $line")
                }
            } catch { }
        }.GetNewClosure()
        $timer.Add_Tick({
            for ($i = $pending.Count - 1; $i -ge 0; $i--) {
                $item = $pending[$i]
                if ($item.Async.IsCompleted) {
                    $result = $null
                    try { $res = $item.PS.EndInvoke($item.Async) } catch { $res = $null }
                    if ($res -is [System.Collections.IEnumerable]) {
                        $result = $res | Select-Object -First 1
                    } else {
                        $result = $res
                    }

                    & $handleQueuedCommandResult $result

                    if ($item.OnComplete -is [scriptblock]) {
                        try { & $item.OnComplete $result } catch {
                            & $logQueuedCommandCallbackFailure "Async command completion callback failed: $($_.Exception.Message)"
                        }
                    }

                    try { $item.PS.Dispose() } catch { }
                    try { $item.Runspace.Close(); $item.Runspace.Dispose() } catch { }
                    $pending.RemoveAt($i)
                }
            }
            if ($pending.Count -eq 0) { $timer.Stop() }
        }.GetNewClosure())

        $resolvedCommandModulesDir = $script:ModuleRoot
        if ([string]::IsNullOrWhiteSpace($resolvedCommandModulesDir) -or -not (Test-Path -LiteralPath $resolvedCommandModulesDir)) {
            try {
                if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot) -and (Test-Path -LiteralPath $PSScriptRoot)) {
                    $resolvedCommandModulesDir = $PSScriptRoot
                }
            } catch { }
        }
        if ([string]::IsNullOrWhiteSpace($resolvedCommandModulesDir) -or -not (Test-Path -LiteralPath $resolvedCommandModulesDir)) {
            try {
                $cwdModulesDir = Join-Path (Get-Location).Path 'Modules'
                if (Test-Path -LiteralPath $cwdModulesDir) {
                    $resolvedCommandModulesDir = $cwdModulesDir
                }
            } catch { }
        }
        if ([string]::IsNullOrWhiteSpace($resolvedCommandModulesDir) -or -not (Test-Path -LiteralPath (Join-Path $resolvedCommandModulesDir 'Logging.psm1'))) {
            throw "Unable to resolve the ECC Modules directory for the Commands window."
        }

        $commandBoxTemplateUpdateInProgress = $false

        $getSelectedCommandsPlayerName = {
            $playerName = ''
            try {
                if ($cmbSpawnPlayer -is [System.Windows.Forms.ComboBox]) {
                    if ($cmbSpawnPlayer.SelectedItem -is [string] -and -not [string]::IsNullOrWhiteSpace($cmbSpawnPlayer.SelectedItem)) {
                        $playerName = $cmbSpawnPlayer.SelectedItem.Trim()
                    } elseif (-not [string]::IsNullOrWhiteSpace($cmbSpawnPlayer.Text)) {
                        $playerName = $cmbSpawnPlayer.Text.Trim()
                    }
                }
            } catch { $playerName = '' }

            return [string]$playerName
        }.GetNewClosure()

        $resolveCommandsPlayerPlaceholders = {
            param([string]$CommandText)

            if ([string]::IsNullOrWhiteSpace($CommandText)) { return [string]$CommandText }

            $playerName = & $getSelectedCommandsPlayerName
            if ([string]::IsNullOrWhiteSpace($playerName)) { return [string]$CommandText }

            $quotedPlayerName = '"' + ($playerName -replace '"', '\"') + '"'
            $quotedReplacement = $quotedPlayerName.Replace('$', '$$')
            $plainReplacement = $playerName.Replace('$', '$$')
            $replacement = if ($supportsPzItemSpawner) { $quotedReplacement } else { $plainReplacement }

            $resolved = $CommandText -replace '(?i)"<user>"', $quotedReplacement
            $resolved = $resolved -replace "(?i)'<user>'", $quotedReplacement
            $resolved = $resolved -replace '(?i)"<player>"', $quotedReplacement
            $resolved = $resolved -replace "(?i)'<player>'", $quotedReplacement
            $resolved = $resolved -replace '(?i)<user>', $replacement
            $resolved = $resolved -replace '(?i)<player>', $replacement
            return [string]$resolved
        }.GetNewClosure()

        $setCommandBoxFromTemplate = {
            param([string]$TemplateText)

            $resolvedText = & $resolveCommandsPlayerPlaceholders $TemplateText
            $commandBoxTemplateUpdateInProgress = $true
            try {
                if ($tbCmd -is [System.Windows.Forms.Control]) {
                    $tbCmd.Tag = if ([string]::IsNullOrWhiteSpace($TemplateText)) { $null } else { [string]$TemplateText }
                    $tbCmd.Text = [string]$resolvedText
                }
            } finally {
                $commandBoxTemplateUpdateInProgress = $false
            }
        }.GetNewClosure()

        if ($tbCmd -is [System.Windows.Forms.Control]) {
            $tbCmd.Add_TextChanged({
                if ($commandBoxTemplateUpdateInProgress) { return }
                try { $tbCmd.Tag = $null } catch { }
            }.GetNewClosure())
        }

        $invokeCommandNowAction = {
            param(
                [string]$CmdText,
                [bool]$ForceRest,
                [bool]$ForceApi,
                [bool]$Verbose
            )

            $lblStatus.Text = 'Sending...'
            $lblStatus.ForeColor = $cmdClrMuted

            try {
                if ($null -eq $commandSharedState -or -not ($commandSharedState -is [hashtable])) {
                    throw 'ECC shared state is unavailable in the Commands window.'
                }

                if (-not (Get-Command -Name Initialize-ServerManager -ErrorAction SilentlyContinue)) {
                    Import-Module (Join-Path $resolvedCommandModulesDir 'ServerManager.psm1') | Out-Null
                }

                Initialize-ServerManager -SharedState $commandSharedState
                $result = Invoke-ServerCommandText -Prefix $Prefix -Command $CmdText -ForceRest:$ForceRest -ForceSatisfactoryApi:$ForceApi -VerboseDebug:$Verbose -SharedState $commandSharedState
                & $handleQueuedCommandResult $result
                return $result
            } catch {
                $result = @{ Success = $false; Message = "[ERROR] Command failed: $_"; Debug = @() }
                & $handleQueuedCommandResult $result
                return $result
            }
        }.GetNewClosure()

        $queueCommandAsyncAction = {
            param(
                [string]$CmdText,
                [bool]$ForceRest,
                [bool]$ForceApi,
                [bool]$Verbose,
                [scriptblock]$OnComplete = $null
            )

            $lblStatus.Text = 'Sending...'
            $lblStatus.ForeColor = $cmdClrMuted

            $capturedModulesDir  = $resolvedCommandModulesDir
            $capturedSharedState = $commandSharedState
            $capturedPrefix      = $Prefix
            $capturedCmd         = $CmdText
            $capturedForceRest   = $ForceRest
            $capturedForceApi    = $ForceApi
            $capturedVerbose     = $Verbose

            $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
            $rs.ApartmentState = 'MTA'
            $rs.ThreadOptions  = 'ReuseThread'
            $rs.Open()
            $rs.SessionStateProxy.SetVariable('ModulesDir',   $capturedModulesDir)
            $rs.SessionStateProxy.SetVariable('SharedState',  $capturedSharedState)
            $rs.SessionStateProxy.SetVariable('TargetPrefix', $capturedPrefix)
            $rs.SessionStateProxy.SetVariable('CmdText',      $capturedCmd)
            $rs.SessionStateProxy.SetVariable('ForceRest',    $capturedForceRest)
            $rs.SessionStateProxy.SetVariable('ForceApi',     $capturedForceApi)
            $rs.SessionStateProxy.SetVariable('Verbose',      $capturedVerbose)

            $ps = [System.Management.Automation.PowerShell]::Create()
            $ps.Runspace = $rs
            $ps.AddScript({
                Set-StrictMode -Off
                $ErrorActionPreference = 'Continue'
                try {
                    Import-Module (Join-Path $ModulesDir 'Logging.psm1')        -Force
                    Import-Module (Join-Path $ModulesDir 'ProfileManager.psm1') -Force
                    Import-Module (Join-Path $ModulesDir 'ServerManager.psm1')  -Force
                    Initialize-ServerManager -SharedState $SharedState
                    return Invoke-ServerCommandText -Prefix $TargetPrefix -Command $CmdText -ForceRest:$ForceRest -ForceSatisfactoryApi:$ForceApi -VerboseDebug:$Verbose -SharedState $SharedState
                } catch {
                    return @{ Success = $false; Message = "[ERROR] Command failed: $_"; Debug = @() }
                }
            }) | Out-Null

            $async = $ps.BeginInvoke()
            $null = $pending.Add(@{ PS = $ps; Runspace = $rs; Async = $async; OnComplete = $OnComplete })
            if (-not $timer.Enabled) { $timer.Start() }
        }.GetNewClosure()

        $queuePlayersRefreshAsync = {
            param(
                [bool]$VerboseDebug = $false,
                [scriptblock]$OnComplete = $null
            )

            $lblStatus.Text = 'Refreshing online players...'
            $lblStatus.ForeColor = $cmdClrMuted

            $capturedModulesDir  = $resolvedCommandModulesDir
            $capturedGuiModulePath = Join-Path $resolvedCommandModulesDir 'GUI.psm1'
            $capturedSharedState = $commandSharedState
            $capturedPrefix      = $Prefix
            $capturedProfile     = $Profile

            $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
            $rs.ApartmentState = 'MTA'
            $rs.ThreadOptions  = 'ReuseThread'
            $rs.Open()
            $rs.SessionStateProxy.SetVariable('ModulesDir',   $capturedModulesDir)
            $rs.SessionStateProxy.SetVariable('GuiModulePath', $capturedGuiModulePath)
            $rs.SessionStateProxy.SetVariable('SharedState',  $capturedSharedState)
            $rs.SessionStateProxy.SetVariable('TargetPrefix', $capturedPrefix)
            $rs.SessionStateProxy.SetVariable('TargetProfile', $capturedProfile)
            $rs.SessionStateProxy.SetVariable('EnableVerboseDebug', ([bool]$VerboseDebug))

            $ps = [System.Management.Automation.PowerShell]::Create()
            $ps.Runspace = $rs
            $ps.AddScript({
                Set-StrictMode -Off
                $ErrorActionPreference = 'Continue'
                try {
                    Import-Module (Join-Path $ModulesDir 'Logging.psm1')        -Force
                    Import-Module (Join-Path $ModulesDir 'ProfileManager.psm1') -Force
                    Import-Module (Join-Path $ModulesDir 'ServerManager.psm1')  -Force
                    Import-Module $GuiModulePath -Force
                    Initialize-ServerManager -SharedState $SharedState

                    try {
                        if ($SharedState -and $SharedState.LatestPlayers) {
                            $SharedState.LatestPlayers.Remove($TargetPrefix) | Out-Null
                        }
                        if ($SharedState -and $SharedState.PlayersRequests) {
                            $SharedState.PlayersRequests[$TargetPrefix] = @{
                                Source = 'UI'
                                RequestedAt = Get-Date
                            }
                        }
                    } catch { }

                    $result = Invoke-ProfileCommand -Prefix $TargetPrefix -CommandName 'players' -SharedState $SharedState

                    $sharedCaptureReady = $false
                    $latestPlayers = @()
                    if ($SharedState -and $SharedState.LatestPlayers) {
                        for ($waitIndex = 0; $waitIndex -lt 25; $waitIndex++) {
                            if ($SharedState.LatestPlayers.ContainsKey($TargetPrefix)) {
                                $latestPlayers = @($SharedState.LatestPlayers[$TargetPrefix])
                                $sharedCaptureReady = $true
                                break
                            }
                            Start-Sleep -Milliseconds 200
                        }
                    }

                    $recentLogText = ''
                    try {
                        $recentLogText = _GetRecentPlayersLogText -Profile $TargetProfile
                    } catch { }

                    return @{
                        Success            = [bool]($result -and $result.Success)
                        Message            = if ($result -and $result.Message) { [string]$result.Message } else { '[INFO] Players refresh finished.' }
                        Debug              = if ($result -and $result.Debug) { @($result.Debug) } else { @() }
                        ResultMessageText  = if ($result -and $result.Message) { [string]$result.Message } else { '' }
                        DebugText          = if ($result -and $result.Debug) { (@($result.Debug) -join [Environment]::NewLine) } else { '' }
                        SharedCaptureReady = $sharedCaptureReady
                        LatestPlayers      = @($latestPlayers)
                        RecentLogText      = [string]$recentLogText
                    }
                } catch {
                    return @{
                        Success            = $false
                        Message            = "[ERROR] Players refresh failed: $_"
                        Debug              = @()
                        ResultMessageText  = ''
                        DebugText          = ''
                        SharedCaptureReady = $false
                        LatestPlayers      = @()
                        RecentLogText      = ''
                    }
                }
            }) | Out-Null

            $async = $ps.BeginInvoke()
            $null = $pending.Add(@{ PS = $ps; Runspace = $rs; Async = $async; OnComplete = $OnComplete })
            if (-not $timer.Enabled) { $timer.Start() }
        }.GetNewClosure()

        $btnSend.Add_Click({
            $templateText = ''
            try {
                if ($tbCmd.Tag -is [string]) { $templateText = [string]$tbCmd.Tag }
            } catch { $templateText = '' }

            $cmdText = if (-not [string]::IsNullOrWhiteSpace($templateText)) {
                & $resolveCommandsPlayerPlaceholders $templateText
            } else {
                & $resolveCommandsPlayerPlaceholders $tbCmd.Text
            }
            $cmdText = $cmdText.Trim()
            if ([string]::IsNullOrWhiteSpace($cmdText)) {
                $lblStatus.Text = 'Command is empty.'
                $lblStatus.ForeColor = $cmdClrRed
                return
            }

            if ($tbCmd.Text -ne $cmdText) {
                $commandBoxTemplateUpdateInProgress = $true
                try { $tbCmd.Text = $cmdText } finally { $commandBoxTemplateUpdateInProgress = $false }
            }

            & $invokeCommandNowAction `
                -CmdText $cmdText `
                -ForceRest:($chkRest.Checked -eq $true) `
                -ForceApi:($chkApi.Checked -eq $true) `
                -Verbose:(& $isCommandsVerboseEnabled)
        }.GetNewClosure())

        $btnTestTelnet.Add_Click({
            try {
                $result = Test-TelnetConnection -Prefix $Prefix -VerboseDebug:(& $isCommandsVerboseEnabled) -SharedState $script:SharedState
                $lblStatus.Text = $result.Message
                $lblStatus.ForeColor = if ($result.Success) { $cmdClrGreen } else { $cmdClrRed }

                $ts = Get-Date -Format 'HH:mm:ss'
                $line = "[$ts] $($result.Message)"
                $tbDebug.AppendText($line + [Environment]::NewLine)
                if ((& $isCommandsVerboseEnabled) -and $result.Debug -and $result.Debug.Count -gt 0) {
                    foreach ($d in $result.Debug) {
                        $dbgLine = "[$ts][DBG] $d"
                        $tbDebug.AppendText($dbgLine + [Environment]::NewLine)
                        try {
                            if ($script:SharedState -and $script:SharedState.LogQueue) {
                                $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][DEBUG][UI] $dbgLine")
                            }
                        } catch { }
                    }
                }
                try {
                    if ($script:SharedState -and $script:SharedState.LogQueue) {
                        $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][INFO][UI] $line")
                    }
                } catch { }
            } catch {
                $lblStatus.Text = "Failed to test Telnet: $_"
                $lblStatus.ForeColor = $cmdClrRed
                $ts = Get-Date -Format 'HH:mm:ss'
                $tbDebug.AppendText("[$ts] Failed to test Telnet: $_" + [Environment]::NewLine)
            }
        }.GetNewClosure())

        $btnTestRcon.Add_Click({
            try {
                $result = Test-RconConnection -Prefix $Prefix -VerboseDebug:(& $isCommandsVerboseEnabled) -SharedState $script:SharedState
                $lblStatus.Text = $result.Message
                $lblStatus.ForeColor = if ($result.Success) { $cmdClrGreen } else { $cmdClrRed }

                $ts = Get-Date -Format 'HH:mm:ss'
                $line = "[$ts] $($result.Message)"
                $tbDebug.AppendText($line + [Environment]::NewLine)
                if ((& $isCommandsVerboseEnabled) -and $result.Debug -and $result.Debug.Count -gt 0) {
                    foreach ($d in $result.Debug) {
                        $dbgLine = "[$ts][DBG] $d"
                        $tbDebug.AppendText($dbgLine + [Environment]::NewLine)
                        try {
                            if ($script:SharedState -and $script:SharedState.LogQueue) {
                                $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][DEBUG][UI] $dbgLine")
                            }
                        } catch { }
                    }
                }
                try {
                    if ($script:SharedState -and $script:SharedState.LogQueue) {
                        $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][INFO][UI] $line")
                    }
                } catch { }
            } catch {
                $lblStatus.Text = "Failed to test RCON: $_"
                $lblStatus.ForeColor = $cmdClrRed
                $ts = Get-Date -Format 'HH:mm:ss'
                $tbDebug.AppendText("[$ts] Failed to test RCON: $_" + [Environment]::NewLine)
            }
        }.GetNewClosure())

        $btnTestApi.Add_Click({
            try {
                $result = Test-SatisfactoryApiConnection -Prefix $Prefix -VerboseDebug:(& $isCommandsVerboseEnabled) -SharedState $script:SharedState
                $lblStatus.Text = $result.Message
                $lblStatus.ForeColor = if ($result.Success) { $cmdClrGreen } else { $cmdClrRed }

                $ts = Get-Date -Format 'HH:mm:ss'
                $line = "[$ts] $($result.Message)"
                $tbDebug.AppendText($line + [Environment]::NewLine)
                if ((& $isCommandsVerboseEnabled) -and $result.Debug -and $result.Debug.Count -gt 0) {
                    foreach ($d in $result.Debug) {
                        $dbgLine = "[$ts][DBG] $d"
                        $tbDebug.AppendText($dbgLine + [Environment]::NewLine)
                        try {
                            if ($script:SharedState -and $script:SharedState.LogQueue) {
                                $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][DEBUG][UI] $dbgLine")
                            }
                        } catch { }
                    }
                }
                try {
                    if ($script:SharedState -and $script:SharedState.LogQueue) {
                        $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][INFO][UI] $line")
                    }
                } catch { }
            } catch {
                $lblStatus.Text = "Failed to test API: $_"
                $lblStatus.ForeColor = $cmdClrRed
                $ts = Get-Date -Format 'HH:mm:ss'
                $tbDebug.AppendText("[$ts] Failed to test API: $_" + [Environment]::NewLine)
            }
        }.GetNewClosure())

        $btnTestStdin.Add_Click({
            try {
                $result = Test-StdInConnection -Prefix $Prefix -VerboseDebug:(& $isCommandsVerboseEnabled) -SharedState $script:SharedState
                $lblStatus.Text = $result.Message
                $lblStatus.ForeColor = if ($result.Success) { $cmdClrGreen } else { $cmdClrRed }

                $ts = Get-Date -Format 'HH:mm:ss'
                $line = "[$ts] $($result.Message)"
                $tbDebug.AppendText($line + [Environment]::NewLine)
                if ((& $isCommandsVerboseEnabled) -and $result.Debug -and $result.Debug.Count -gt 0) {
                    foreach ($d in $result.Debug) {
                        $dbgLine = "[$ts][DBG] $d"
                        $tbDebug.AppendText($dbgLine + [Environment]::NewLine)
                        try {
                            if ($script:SharedState -and $script:SharedState.LogQueue) {
                                $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][DEBUG][UI] $dbgLine")
                            }
                        } catch { }
                    }
                }
                try {
                    if ($script:SharedState -and $script:SharedState.LogQueue) {
                        $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][INFO][UI] $line")
                    }
                } catch { }
            } catch {
                $lblStatus.Text = "Failed to test STDIN: $_"
                $lblStatus.ForeColor = $cmdClrRed
                $ts = Get-Date -Format 'HH:mm:ss'
                $tbDebug.AppendText("[$ts] Failed to test STDIN: $_" + [Environment]::NewLine)
            }
        }.GetNewClosure())

        $btnTestPid.Add_Click({
            try {
                $status = Get-ServerStatus -Prefix $Prefix
                if ($status.Running) {
                    $msg = "[PID] $($Profile.GameName): $($status.Pid)"
                    $lblStatus.Text = $msg
                    $lblStatus.ForeColor = $cmdClrGreen
                } else {
                    $msg = "[OFFLINE] $($Profile.GameName) is not running."
                    $lblStatus.Text = $msg
                    $lblStatus.ForeColor = $cmdClrRed
                }
                $ts = Get-Date -Format 'HH:mm:ss'
                $line = "[$ts] $msg"
                $tbDebug.AppendText($line + [Environment]::NewLine)
                try {
                    if ($script:SharedState -and $script:SharedState.LogQueue) {
                        $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][INFO][UI] $line")
                    }
                } catch { }
            } catch {
                $lblStatus.Text = "Failed to test PID: $_"
                $lblStatus.ForeColor = $cmdClrRed
                $ts = Get-Date -Format 'HH:mm:ss'
                $tbDebug.AppendText("[$ts] Failed to test PID: $_" + [Environment]::NewLine)
            }
        }.GetNewClosure())

        $btnTestRest.Add_Click({
            & $queueCommandAsyncAction `
                -CmdText 'GET /v1/api/info' `
                -ForceRest $true `
                -ForceApi $false `
                -Verbose:(& $isCommandsVerboseEnabled)
        }.GetNewClosure())

        if ($supportsItemSpawner -and $spawnPanel) {
            $pzVehicleTextureManifest = $null
            try { $pzVehicleTextureManifest = _PZSharedLoadVehicleTextureMap } catch { $pzVehicleTextureManifest = $null }
            if ($supportsVehicleSpawner -and $null -eq $pzVehicleTextureManifest) {
                try {
                    $workspaceRootForVehicleManifest = Split-Path -Parent $PSScriptRoot
                    $vehicleManifestPath = Join-Path $workspaceRootForVehicleManifest 'Config\AssetCache\ProjectZomboid\vehicle-texture-manifest.json'
                    if (Test-Path -LiteralPath $vehicleManifestPath) {
                        $directVehicleManifest = Get-Content -LiteralPath $vehicleManifestPath -Raw -ErrorAction Stop | ConvertFrom-Json
                        if ($directVehicleManifest -and $directVehicleManifest.VehicleTextureByName) {
                            $pzVehicleTextureManifest = $directVehicleManifest
                        }
                    }
                } catch { $pzVehicleTextureManifest = $null }
            }

            $loadPzPreviewImage = {
                param([string]$ImagePath)
                if ([string]::IsNullOrWhiteSpace($ImagePath) -or -not (Test-Path -LiteralPath $ImagePath)) { return $null }
                try {
                    $image = [System.Drawing.Image]::FromFile($ImagePath)
                    try {
                        return [System.Drawing.Image]$image.Clone()
                    } finally {
                        $image.Dispose()
                    }
                } catch {
                    return $null
                }
            }.GetNewClosure()

            $resolvePzVehiclePreviewPath = {
                param(
                    [string]$FullType,
                    [string]$VehicleName,
                    [string]$DisplayName
                )

                if ($null -eq $pzVehicleTextureManifest -or -not $pzVehicleTextureManifest.VehicleTextureByName) { return $null }

                $candidateKeys = @(
                    $FullType,
                    $VehicleName,
                    $DisplayName,
                    (_PZSharedNormalizeAssetKey $FullType),
                    (_PZSharedNormalizeAssetKey $VehicleName),
                    (_PZSharedNormalizeAssetKey $DisplayName)
                ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique

                foreach ($key in @($candidateKeys)) {
                    try {
                        $candidate = $null
                        $vehicleTextureMap = $pzVehicleTextureManifest.VehicleTextureByName
                        if ($vehicleTextureMap -is [System.Collections.IDictionary]) {
                            if ($vehicleTextureMap.Contains($key)) {
                                $candidate = $vehicleTextureMap[$key]
                            }
                        } else {
                            $prop = $vehicleTextureMap.PSObject.Properties[$key]
                            if ($null -ne $prop) {
                                $candidate = $prop.Value
                            }
                        }
                        if (-not [string]::IsNullOrWhiteSpace($candidate) -and (Test-Path -LiteralPath $candidate)) {
                            return [string]$candidate
                        }
                    } catch { }
                }

                return $null
            }.GetNewClosure()

            $updatePzItemPreview = {
                param([object]$SelectedItem)

                $clearPreviewImage = {
                    if ($picItemPreview -and $picItemPreview.PSObject.Properties['Image']) {
                        try {
                            if ($picItemPreview.Image) { $picItemPreview.Image.Dispose() }
                        } catch { }
                        $picItemPreview.Image = $null
                        $picItemPreview.Visible = $false
                    }
                }.GetNewClosure()

                if ($null -eq $SelectedItem) {
                    $lblItemName.Text = 'Select an item'
                    $lblItemType.Text = 'No item selected yet.'
                    & $clearPreviewImage
                    return
                }

                $lblItemName.Text = [string]$SelectedItem.DisplayName
                $lblItemType.Text = "$($SelectedItem.FullType)`r`nCategory: $($SelectedItem.DisplayCategory)"
                & $clearPreviewImage

                $resolvedIconPath = ''
                try {
                    $resolvedIconPath = [string]$SelectedItem.IconPath
                } catch { $resolvedIconPath = '' }
                if ([string]::IsNullOrWhiteSpace($resolvedIconPath) -or -not (Test-Path -LiteralPath $resolvedIconPath)) {
                    try {
                        if ($supportsMinecraftItemSpawner) {
                            $resolvedIconPath = _MinecraftSharedResolveItemPreviewPath -ItemId ([string]$SelectedItem.ItemId) -FullType ([string]$SelectedItem.FullType)
                        } elseif ($supportsHytaleItemSpawner) {
                            $resolvedIconPath = _HytaleSharedResolveItemPreviewPath -ItemId ([string]$SelectedItem.ItemId) -FullType ([string]$SelectedItem.FullType)
                        } else {
                            $resolvedIconPath = _PZSharedResolveItemPreviewPath -ItemName ([string]$SelectedItem.ItemName) -IconName ([string]$SelectedItem.IconName) -AssetManifest (_PZSharedLoadImportedItemTextureMap)
                        }
                        if (-not [string]::IsNullOrWhiteSpace($resolvedIconPath)) {
                            try { $SelectedItem.IconPath = $resolvedIconPath } catch { }
                        }
                    } catch { $resolvedIconPath = '' }
                }

                if (-not [string]::IsNullOrWhiteSpace($resolvedIconPath) -and (Test-Path -LiteralPath $resolvedIconPath)) {
                    try {
                        if ($picItemPreview -and $picItemPreview.PSObject.Properties['Image']) {
                            $picItemPreview.Image = & $loadPzPreviewImage $resolvedIconPath
                            $picItemPreview.Visible = ($null -ne $picItemPreview.Image)
                            $picItemPreview.BringToFront()
                        }
                    } catch {
                        & $clearPreviewImage
                    }
                }
            }.GetNewClosure()

            $updatePzVehiclePreview = {
                param([object]$SelectedVehicle)

                $clearVehiclePreviewImage = {
                    if ($picVehiclePreview -and $picVehiclePreview.PSObject.Properties['Image']) {
                        try {
                            if ($picVehiclePreview.Image) { $picVehiclePreview.Image.Dispose() }
                        } catch { }
                        $picVehiclePreview.Image = $null
                        $picVehiclePreview.Visible = $false
                    }
                }.GetNewClosure()

                if ($null -eq $SelectedVehicle) {
                    $lblVehicleName.Text = 'Select a vehicle'
                    $lblVehicleType.Text = 'No vehicle selected yet.'
                    & $clearVehiclePreviewImage
                    return
                }

                $lblVehicleName.Text = [string]$SelectedVehicle.DisplayName
                $categoryText = if ([string]::IsNullOrWhiteSpace($SelectedVehicle.DisplayCategory)) { 'vehicles' } else { $SelectedVehicle.DisplayCategory }
                $lblVehicleType.Text = "$($SelectedVehicle.FullType)`r`nCategory: $categoryText"
                & $clearVehiclePreviewImage

                $resolvedVehiclePreviewPath = ''
                try {
                    if ($SelectedVehicle.PSObject.Properties.Name -contains 'PreviewPath') {
                        $resolvedVehiclePreviewPath = [string]$SelectedVehicle.PreviewPath
                    }
                } catch { $resolvedVehiclePreviewPath = '' }

                if ([string]::IsNullOrWhiteSpace($resolvedVehiclePreviewPath) -or -not (Test-Path -LiteralPath $resolvedVehiclePreviewPath)) {
                    try {
                        $resolvedVehiclePreviewPath = & $resolvePzVehiclePreviewPath ([string]$SelectedVehicle.FullType) ([string]$SelectedVehicle.VehicleName) ([string]$SelectedVehicle.DisplayName)
                        if (-not [string]::IsNullOrWhiteSpace($resolvedVehiclePreviewPath)) {
                            try {
                                if ($SelectedVehicle.PSObject.Properties.Name -notcontains 'PreviewPath') {
                                    $SelectedVehicle | Add-Member -NotePropertyName PreviewPath -NotePropertyValue $resolvedVehiclePreviewPath -Force
                                } else {
                                    $SelectedVehicle.PreviewPath = $resolvedVehiclePreviewPath
                                }
                            } catch { }
                        }
                    } catch { $resolvedVehiclePreviewPath = '' }
                }

                if (-not [string]::IsNullOrWhiteSpace($resolvedVehiclePreviewPath) -and (Test-Path -LiteralPath $resolvedVehiclePreviewPath)) {
                    try {
                        if ($picVehiclePreview -and $picVehiclePreview.PSObject.Properties['Image']) {
                            $picVehiclePreview.Image = & $loadPzPreviewImage $resolvedVehiclePreviewPath
                            $picVehiclePreview.Visible = ($null -ne $picVehiclePreview.Image)
                            $picVehiclePreview.BringToFront()
                        }
                    } catch {
                        & $clearVehiclePreviewImage
                    }
                }
            }.GetNewClosure()

            $buildPzAddItemCommand = {
                $playerName = ''
                if ($cmbSpawnPlayer.SelectedItem -is [string] -and -not [string]::IsNullOrWhiteSpace($cmbSpawnPlayer.SelectedItem)) {
                    $playerName = $cmbSpawnPlayer.SelectedItem.Trim()
                } else {
                    $playerName = $cmbSpawnPlayer.Text.Trim()
                }

                $selectedItem = $null
                if ($lbItems.SelectedItem) {
                    $selectedKey = [string]$lbItems.SelectedItem
                    if ($pzItemLookup.ContainsKey($selectedKey)) {
                        $selectedItem = $pzItemLookup[$selectedKey]
                    }
                }
                if ($null -eq $selectedItem) {
                    return $null
                }
                if ([string]::IsNullOrWhiteSpace($playerName)) { $playerName = '<user>' }

                $count = 1
                try { $count = [int]$numSpawnCount.Value } catch { $count = 1 }
                if ($count -lt 1) { $count = 1 }

                if ($supportsMinecraftItemSpawner) {
                    return "/give $playerName $($selectedItem.FullType) $count"
                } elseif ($supportsHytaleItemSpawner) {
                    return "/give $playerName $($selectedItem.FullType) --quantity=$count"
                }

                return "/additem ""$playerName"" ""$($selectedItem.FullType)"" $count"
            }.GetNewClosure()

            $buildPzAddVehicleCommand = {
                $playerName = ''
                if ($cmbSpawnPlayer.SelectedItem -is [string] -and -not [string]::IsNullOrWhiteSpace($cmbSpawnPlayer.SelectedItem)) {
                    $playerName = $cmbSpawnPlayer.SelectedItem.Trim()
                } else {
                    $playerName = $cmbSpawnPlayer.Text.Trim()
                }

                $selectedVehicle = $null
                if ($lbVehicles.SelectedItem) {
                    $selectedKey = [string]$lbVehicles.SelectedItem
                    if ($pzVehicleLookup.ContainsKey($selectedKey)) {
                        $selectedVehicle = $pzVehicleLookup[$selectedKey]
                    }
                }
                if ($null -eq $selectedVehicle) {
                    return $null
                }
                if ([string]::IsNullOrWhiteSpace($playerName)) { $playerName = '<user>' }

                return "/addvehicle ""$($selectedVehicle.FullType)"" ""$playerName"""
            }.GetNewClosure()

            $refreshPzItemsList = {
                param([string]$SearchText)

                $lbItems.BeginUpdate()
                try {
                    $lbItems.Items.Clear()
                    $pzItemLookup.Clear()
                    $needle = if ([string]::IsNullOrWhiteSpace($SearchText)) { '' } else { $SearchText.Trim().ToLowerInvariant() }
                    $firstSelectableIndex = -1
                    $lastCategoryKey = ''

                    $sortedItems = @(
                        @($pzCatalogState.ItemSource) |
                        Sort-Object @{ Expression = { if ([string]::IsNullOrWhiteSpace($_.DisplayCategory)) { 'zzz_misc' } else { $_.DisplayCategory } } }, DisplayName, FullType
                    )

                    foreach ($item in $sortedItems) {
                        if ($needle) {
                            $haystack = ("{0} {1} {2}" -f $item.DisplayName, $item.FullType, $item.DisplayCategory).ToLowerInvariant()
                            if ($haystack -notlike "*$needle*") { continue }
                        }

                        $categoryKey = if ([string]::IsNullOrWhiteSpace($item.DisplayCategory)) { 'Misc' } else { [string]$item.DisplayCategory }
                        if ($categoryKey -ne $lastCategoryKey) {
                            [void]$lbItems.Items.Add(("=== {0} ===" -f $categoryKey))
                            $lastCategoryKey = $categoryKey
                        }

                        $listText = [string]$item.ListText
                        $pzItemLookup[$listText] = $item
                        [void]$lbItems.Items.Add($listText)
                        if ($firstSelectableIndex -lt 0) {
                            $firstSelectableIndex = $lbItems.Items.Count - 1
                        }
                    }
                } finally {
                    $lbItems.EndUpdate()
                }

                $lblSpawnHint.Text = if ($lbItems.Items.Count -gt 0) {
                    "Showing $($lbItems.Items.Count) item(s). Double-click to build the command."
                } else {
                    'No items matched that search.'
                }

                if ($firstSelectableIndex -ge 0) {
                    $lbItems.SelectedIndex = $firstSelectableIndex
                } else {
                    & $updatePzItemPreview $null
                }
            }.GetNewClosure()

            $refreshPzVehiclesList = {
                param([string]$SearchText)

                $lbVehicles.BeginUpdate()
                try {
                    $lbVehicles.Items.Clear()
                    $pzVehicleLookup.Clear()
                    $needle = if ([string]::IsNullOrWhiteSpace($SearchText)) { '' } else { $SearchText.Trim().ToLowerInvariant() }
                    $firstSelectableIndex = -1
                    $lastCategoryKey = ''

                    $sortedVehicles = @(
                        @($pzCatalogState.VehicleSource) |
                        Sort-Object @{ Expression = { if ([string]::IsNullOrWhiteSpace($_.DisplayCategory)) { 'zzz_misc' } else { $_.DisplayCategory } } }, DisplayName, FullType
                    )

                    foreach ($vehicle in $sortedVehicles) {
                        if ($needle) {
                            $haystack = ("{0} {1} {2}" -f $vehicle.DisplayName, $vehicle.FullType, $vehicle.DisplayCategory).ToLowerInvariant()
                            if ($haystack -notlike "*$needle*") { continue }
                        }

                        $categoryKey = if ([string]::IsNullOrWhiteSpace($vehicle.DisplayCategory)) { 'Misc' } else { [string]$vehicle.DisplayCategory }
                        if ($categoryKey -ne $lastCategoryKey) {
                            [void]$lbVehicles.Items.Add(("=== {0} ===" -f $categoryKey))
                            $lastCategoryKey = $categoryKey
                        }

                        $listText = [string]$vehicle.ListText
                        $pzVehicleLookup[$listText] = $vehicle
                        [void]$lbVehicles.Items.Add($listText)
                        if ($firstSelectableIndex -lt 0) {
                            $firstSelectableIndex = $lbVehicles.Items.Count - 1
                        }
                    }
                } finally {
                    $lbVehicles.EndUpdate()
                }

                $lblVehicleHint.Text = if ($lbVehicles.Items.Count -gt 0) {
                    "Showing $($lbVehicles.Items.Count) vehicle(s). Double-click to build the command."
                } else {
                    'No vehicles matched that search.'
                }

                if ($firstSelectableIndex -ge 0) {
                    $lbVehicles.SelectedIndex = $firstSelectableIndex
                } else {
                    & $updatePzVehiclePreview $null
                }
            }.GetNewClosure()

            $loadPzCatalogsFromCache = {
                param([hashtable]$CatalogProfile)
                $cmdName = if ($supportsMinecraftItemSpawner) {
                    'Get-MinecraftSpawnerCatalogsFromCache'
                } elseif ($supportsHytaleItemSpawner) {
                    'Get-HytaleSpawnerCatalogsFromCache'
                } else {
                    'Get-ProjectZomboidSpawnerCatalogsFromCache'
                }
                $cmd = Get-Command -Name $cmdName -CommandType Function -ErrorAction Stop
                return & $cmd -Profile $CatalogProfile
            }.GetNewClosure()

            $loadPzCatalogsFull = {
                param([hashtable]$CatalogProfile)
                $cmdName = if ($supportsMinecraftItemSpawner) {
                    'Get-MinecraftSpawnerCatalogs'
                } elseif ($supportsHytaleItemSpawner) {
                    'Get-HytaleSpawnerCatalogs'
                } else {
                    'Get-ProjectZomboidSpawnerCatalogs'
                }
                $cmd = Get-Command -Name $cmdName -CommandType Function -ErrorAction Stop
                return & $cmd -Profile $CatalogProfile
            }.GetNewClosure()

            $setPzCatalogUiEnabled = {
                param([bool]$Enabled)

                foreach ($control in @($tbItemSearch, $lbItems, $numSpawnCount, $btnInsertAddItem, $btnBuildAddItem, $tbVehicleSearch, $lbVehicles, $btnInsertAddVehicle, $btnBuildAddVehicle)) {
                    if ($control -is [System.Windows.Forms.Control]) {
                        $control.Enabled = $Enabled
                    }
                }
            }.GetNewClosure()

            $applyPzCatalogsToUi = {
                param([object]$CatalogResult)

                $pzCatalogState.ItemSource = if ($CatalogResult -and $CatalogResult.PSObject.Properties.Name -contains 'Items') { @($CatalogResult.Items) } else { @() }
                $pzCatalogState.VehicleSource = if ($CatalogResult -and $CatalogResult.PSObject.Properties.Name -contains 'Vehicles') { @($CatalogResult.Vehicles) } else { @() }
                $pzCatalogLoaded = $true

                & $setPzCatalogUiEnabled $true
                & $refreshPzItemsList $tbItemSearch.Text
                if ($supportsVehicleSpawner) {
                    & $refreshPzVehiclesList $(if ($tbVehicleSearch) { $tbVehicleSearch.Text } else { '' })
                }

                if (@($pzCatalogState.ItemSource).Count -eq 0) {
                    $lblSpawnHint.Text = $itemSpawnerMissingHint
                }
                if ($supportsVehicleSpawner -and @($pzCatalogState.VehicleSource).Count -eq 0) {
                    $lblVehicleHint.Text = 'ECC could not load local vehicle data from the Project Zomboid install path.'
                }
            }.GetNewClosure()

            $lbItems.DisplayMember = 'ListText'
            if ($lbVehicles) {
                $lbVehicles.DisplayMember = 'ListText'
            }
            & $setPzCatalogUiEnabled $false

            $capturedProfileForPzCatalog = @{}
            foreach ($entry in $Profile.GetEnumerator()) {
                $capturedProfileForPzCatalog[$entry.Key] = $entry.Value
            }

            $hasInitialPzCatalogData = $false
            try {
                $cachedPzCatalogs = & $loadPzCatalogsFromCache $capturedProfileForPzCatalog
                $cachedItemCount = if ($cachedPzCatalogs -and $cachedPzCatalogs.PSObject.Properties.Name -contains 'Items') { @($cachedPzCatalogs.Items).Count } else { 0 }
                $cachedVehicleCount = if ($cachedPzCatalogs -and $cachedPzCatalogs.PSObject.Properties.Name -contains 'Vehicles') { @($cachedPzCatalogs.Vehicles).Count } else { 0 }
                if ($cachedItemCount -gt 0 -or $cachedVehicleCount -gt 0) {
                    $hasInitialPzCatalogData = $true
                    & $applyPzCatalogsToUi $cachedPzCatalogs
                }
            } catch {
                try { $lblSpawnHint.Text = "Cache load error: $($_.Exception.Message)" } catch { }
                try { $lblVehicleHint.Text = "Cache load error: $($_.Exception.Message)" } catch { }
            }

            $tbItemSearch.Add_TextChanged({
                param($sender, $eventArgs)
                $searchText = ''
                try {
                    if ($sender -is [System.Windows.Forms.TextBoxBase]) {
                        $searchText = [string]$sender.Text
                    } else {
                        $searchText = [string]$tbItemSearch.Text
                    }
                } catch { $searchText = '' }
                if (@($pzCatalogState.ItemSource).Count -gt 0) {
                    & $refreshPzItemsList $searchText
                }
            }.GetNewClosure())

            if ($tbVehicleSearch) {
                $tbVehicleSearch.Add_TextChanged({
                    param($sender, $eventArgs)
                    $searchText = ''
                    try {
                        if ($sender -is [System.Windows.Forms.TextBoxBase]) {
                            $searchText = [string]$sender.Text
                        } else {
                            $searchText = [string]$tbVehicleSearch.Text
                        }
                    } catch { $searchText = '' }
                    if (@($pzCatalogState.VehicleSource).Count -gt 0) {
                        & $refreshPzVehiclesList $searchText
                    }
                }.GetNewClosure())
            }

            $lbItems.Add_SelectedIndexChanged({
                $selectedItem = $null
                if ($lbItems.SelectedItem) {
                    $selectedKey = [string]$lbItems.SelectedItem
                    if ($pzItemLookup.ContainsKey($selectedKey)) {
                        $selectedItem = $pzItemLookup[$selectedKey]
                    }
                }
                & $updatePzItemPreview $selectedItem
                $cmdPreview = & $buildPzAddItemCommand
                if ($cmdPreview) { $tbCmd.Text = $cmdPreview }
            }.GetNewClosure())

            $lbItems.Add_DoubleClick({
                $cmdPreview = & $buildPzAddItemCommand
                if ($cmdPreview) {
                    $tbCmd.Text = $cmdPreview
                    $lblStatus.Text = "Inserted $itemSpawnerCommandName command into the command box."
                    $lblStatus.ForeColor = $cmdClrGreen
                }
            }.GetNewClosure())

            $lbItems.Add_KeyDown({
                if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
                    $cmdPreview = & $buildPzAddItemCommand
                    if ($cmdPreview) {
                        $tbCmd.Text = $cmdPreview
                        $lblStatus.Text = "Inserted $itemSpawnerCommandName command into the command box."
                        $lblStatus.ForeColor = $cmdClrGreen
                    }
                    $_.Handled = $true
                }
            }.GetNewClosure())

            if ($lbVehicles) {
                $lbVehicles.Add_SelectedIndexChanged({
                    $selectedVehicle = $null
                    if ($lbVehicles.SelectedItem) {
                        $selectedKey = [string]$lbVehicles.SelectedItem
                        if ($pzVehicleLookup.ContainsKey($selectedKey)) {
                            $selectedVehicle = $pzVehicleLookup[$selectedKey]
                        }
                    }
                    & $updatePzVehiclePreview $selectedVehicle
                    $cmdPreview = & $buildPzAddVehicleCommand
                    if ($cmdPreview) { $tbCmd.Text = $cmdPreview }
                }.GetNewClosure())

                $lbVehicles.Add_DoubleClick({
                    $cmdPreview = & $buildPzAddVehicleCommand
                    if ($cmdPreview) {
                        $tbCmd.Text = $cmdPreview
                        $lblStatus.Text = 'Inserted addvehicle command into the command box.'
                        $lblStatus.ForeColor = $cmdClrGreen
                    }
                }.GetNewClosure())

                $lbVehicles.Add_KeyDown({
                    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
                        $cmdPreview = & $buildPzAddVehicleCommand
                        if ($cmdPreview) {
                            $tbCmd.Text = $cmdPreview
                            $lblStatus.Text = 'Inserted addvehicle command into the command box.'
                            $lblStatus.ForeColor = $cmdClrGreen
                        }
                        $_.Handled = $true
                    }
                }.GetNewClosure())
            }

            $btnInsertAddItem.Add_Click({
                $cmdPreview = & $buildPzAddItemCommand
                if ($cmdPreview) {
                    $tbCmd.Text = $cmdPreview
                    if ([string]::IsNullOrWhiteSpace($cmbSpawnPlayer.Text) -and -not ($cmbSpawnPlayer.SelectedItem -is [string] -and -not [string]::IsNullOrWhiteSpace($cmbSpawnPlayer.SelectedItem))) {
                        $lblStatus.Text = 'Inserted a command with a <user> placeholder. Select a player or edit it manually below.'
                    } else {
                        $lblStatus.Text = "Inserted $itemSpawnerCommandName command into the command box."
                    }
                    $lblStatus.ForeColor = $cmdClrGreen
                } else {
                        $lblStatus.Text = 'Select an item first.'
                    $lblStatus.ForeColor = $cmdClrRed
                }
            }.GetNewClosure())

            if ($btnInsertAddVehicle) {
                $btnInsertAddVehicle.Add_Click({
                    $cmdPreview = & $buildPzAddVehicleCommand
                    if ($cmdPreview) {
                        $tbCmd.Text = $cmdPreview
                        if ([string]::IsNullOrWhiteSpace($cmbSpawnPlayer.Text) -and -not ($cmbSpawnPlayer.SelectedItem -is [string] -and -not [string]::IsNullOrWhiteSpace($cmbSpawnPlayer.SelectedItem))) {
                            $lblStatus.Text = 'Inserted a vehicle command with a <user> placeholder. Select a player or edit it manually below.'
                        } else {
                            $lblStatus.Text = 'Inserted addvehicle command into the command box.'
                        }
                        $lblStatus.ForeColor = $cmdClrGreen
                    } else {
                        $lblStatus.Text = 'Select a vehicle first.'
                        $lblStatus.ForeColor = $cmdClrRed
                    }
                }.GetNewClosure())
            }

            $cmbSpawnPlayer.Add_TextChanged({
                $cmdPreview = if ($spawnTabControl -and $spawnTabControl.SelectedTab -eq $tabVehicles) { & $buildPzAddVehicleCommand } else { & $buildPzAddItemCommand }
                if ($cmdPreview) { $tbCmd.Text = $cmdPreview }
            }.GetNewClosure())

            $cmbSpawnPlayer.Add_SelectedIndexChanged({
                $cmdPreview = if ($spawnTabControl -and $spawnTabControl.SelectedTab -eq $tabVehicles) { & $buildPzAddVehicleCommand } else { & $buildPzAddItemCommand }
                if ($cmdPreview) { $tbCmd.Text = $cmdPreview }
            }.GetNewClosure())

            $numSpawnCount.Add_ValueChanged({
                $cmdPreview = & $buildPzAddItemCommand
                if ($cmdPreview) { $tbCmd.Text = $cmdPreview }
            }.GetNewClosure())

            $btnBuildAddItem.Add_Click({
                $cmdPreview = & $buildPzAddItemCommand
                if ($cmdPreview) {
                    $tbCmd.Text = $cmdPreview
                    if ([string]::IsNullOrWhiteSpace($cmbSpawnPlayer.Text) -and -not ($cmbSpawnPlayer.SelectedItem -is [string] -and -not [string]::IsNullOrWhiteSpace($cmbSpawnPlayer.SelectedItem))) {
                        $lblStatus.Text = 'Built command with <user> placeholder. Review it below, then press Send.'
                    } else {
                        $lblStatus.Text = "Built $itemSpawnerCommandName command. You can review it below, then press Send."
                    }
                    $lblStatus.ForeColor = $cmdClrGreen
                } else {
                    $lblStatus.Text = 'Pick an item first.'
                    $lblStatus.ForeColor = $cmdClrRed
                }
            }.GetNewClosure())

            if ($btnBuildAddVehicle) {
                $btnBuildAddVehicle.Add_Click({
                    $cmdPreview = & $buildPzAddVehicleCommand
                    if ($cmdPreview) {
                        $tbCmd.Text = $cmdPreview
                        if ([string]::IsNullOrWhiteSpace($cmbSpawnPlayer.Text) -and -not ($cmbSpawnPlayer.SelectedItem -is [string] -and -not [string]::IsNullOrWhiteSpace($cmbSpawnPlayer.SelectedItem))) {
                            $lblStatus.Text = 'Built vehicle command with <user> placeholder. Review it below, then press Send.'
                        } else {
                            $lblStatus.Text = 'Built addvehicle command. You can review it below, then press Send.'
                        }
                        $lblStatus.ForeColor = $cmdClrGreen
                    } else {
                        $lblStatus.Text = 'Select a vehicle first.'
                        $lblStatus.ForeColor = $cmdClrRed
                    }
                }.GetNewClosure())
            }

            $playersRefreshInFlight = $false
            $btnRefreshPlayers.Add_Click({
                if ($playersRefreshInFlight) {
                    $lblStatus.Text = 'Player refresh is already running...'
                    $lblStatus.ForeColor = $cmdClrMuted
                    return
                }

                $playersRefreshInFlight = $true
                $btnRefreshPlayers.Enabled = $false

                & $queuePlayersRefreshAsync -VerboseDebug:(& $isCommandsVerboseEnabled) -OnComplete {
                    param([object]$refreshResult)

                    $playersRefreshInFlight = $false
                    $btnRefreshPlayers.Enabled = $true

                    $playerNames = @()
                    $parsedFrom = 'none'
                    $resultMessageText = ''
                    $debugText = ''
                    $gameLogText = ''
                    $recentLogText = ''
                    $sharedCaptureReady = $false

                    try {
                        if ($refreshResult) {
                            $resultMessageText = if ($refreshResult.ResultMessageText) { [string]$refreshResult.ResultMessageText } else { '' }
                            $debugText = if ($refreshResult.DebugText) { [string]$refreshResult.DebugText } else { '' }
                            $recentLogText = if ($refreshResult.RecentLogText) { [string]$refreshResult.RecentLogText } else { '' }
                            $sharedCaptureReady = [bool]$refreshResult.SharedCaptureReady
                        }

                        if ($sharedCaptureReady) {
                            $playerNames = @($refreshResult.LatestPlayers)
                            $parsedFrom = 'shared-capture'
                        }

                        if (($null -eq $playerNames -or $playerNames.Count -eq 0) -and -not [string]::IsNullOrWhiteSpace($resultMessageText)) {
                            $playerNames = @(& $parseOnlinePlayersText $resultMessageText)
                            if ($playerNames.Count -gt 0) { $parsedFrom = 'result-message' }
                        }
                        if (($null -eq $playerNames -or $playerNames.Count -eq 0) -and -not [string]::IsNullOrWhiteSpace($debugText)) {
                            $playerNames = @(& $parseOnlinePlayersText $debugText)
                            if ($playerNames.Count -gt 0) { $parsedFrom = 'command-debug' }
                        }
                        if ($null -eq $playerNames -or $playerNames.Count -eq 0) {
                            $gameLogText = & $getGameLogTabTextLocal $Prefix
                            if (-not [string]::IsNullOrWhiteSpace($gameLogText)) {
                                $playerNames = @(& $parseOnlinePlayersText $gameLogText)
                                if ($playerNames.Count -gt 0) { $parsedFrom = 'live-log-tab' }
                            }
                        }
                        if (($null -eq $playerNames -or $playerNames.Count -eq 0) -and -not [string]::IsNullOrWhiteSpace($recentLogText)) {
                            $playerNames = @(& $parseOnlinePlayersText $recentLogText)
                            if ($playerNames.Count -gt 0) { $parsedFrom = 'log-file-tail' }
                        }

                        & $writeCommandsDebugLine ("refresh result success={0} messageLen={1} debugLines={2}" -f `
                            [bool]($refreshResult -and $refreshResult.Success), `
                            $resultMessageText.Length, `
                            $(if ($refreshResult -and $refreshResult.Debug) { @($refreshResult.Debug).Count } else { 0 }))
                        & $writeCommandsDebugLine ("refresh sources parsedFrom={0} names={1} sharedCaptureReady={2}" -f $parsedFrom, (@($playerNames).Count), $sharedCaptureReady)
                        & $writeCommandsDebugLine ("refresh text lengths result={0} cmdDebug={1} logTab={2} fileTail={3}" -f `
                            $resultMessageText.Length, `
                            $debugText.Length, `
                            $gameLogText.Length, `
                            $recentLogText.Length)
                        if (@($playerNames).Count -gt 0) {
                            & $writeCommandsDebugLine ("refresh players: " + ((@($playerNames) | ForEach-Object { "'$_'" }) -join ', '))
                        }
                    } catch {
                        $playerNames = @()
                        $sharedCaptureReady = $false
                        & $writeCommandsDebugLine ("refresh exception: $_")
                    }

                    $previousText = $cmbSpawnPlayer.Text
                    $cmbSpawnPlayer.Items.Clear()
                    foreach ($name in $playerNames) {
                        [void]$cmbSpawnPlayer.Items.Add($name)
                    }
                    & $writeCommandsDebugLine ("combo population count={0}" -f $cmbSpawnPlayer.Items.Count)

                    if ($cmbSpawnPlayer.Items.Count -gt 0) {
                        $cmbSpawnPlayer.SelectedIndex = 0
                        if (-not [string]::IsNullOrWhiteSpace($previousText)) {
                            $matchIndex = $cmbSpawnPlayer.FindStringExact($previousText)
                            if ($matchIndex -ge 0) { $cmbSpawnPlayer.SelectedIndex = $matchIndex }
                        }
                        $lblSpawnHint.Text = "Found $($cmbSpawnPlayer.Items.Count) online player(s)."
                        $lblStatus.Text = "Players refreshed: $($cmbSpawnPlayer.Items.Count) online."
                        $lblStatus.ForeColor = $cmdClrGreen
                        & $writeCommandsDebugLine ("combo selected='{0}'" -f $cmbSpawnPlayer.Text)
                    } elseif ($sharedCaptureReady) {
                        $cmbSpawnPlayer.Text = ''
                        $lblSpawnHint.Text = 'No online players reported right now. You can still type a username manually.'
                        $lblStatus.Text = 'Players refreshed: no players online.'
                        $lblStatus.ForeColor = $cmdClrMuted
                    } else {
                        if (-not [string]::IsNullOrWhiteSpace($previousText)) { $cmbSpawnPlayer.Text = $previousText }
                        $lblSpawnHint.Text = 'No player names were parsed from the response. You can still type a username manually.'
                    }

                    $cmdPreview = if ($spawnTabControl -and $spawnTabControl.SelectedTab -eq $tabVehicles) { & $buildPzAddVehicleCommand } else { & $buildPzAddItemCommand }
                    if ($cmdPreview) { $tbCmd.Text = $cmdPreview }
                }
            }.GetNewClosure())

            if ($spawnTabControl) {
                $spawnTabControl.Add_SelectedIndexChanged({
                    $cmdPreview = if ($spawnTabControl.SelectedTab -eq $tabVehicles) { & $buildPzAddVehicleCommand } else { & $buildPzAddItemCommand }
                    if ($cmdPreview) { $tbCmd.Text = $cmdPreview }
                }.GetNewClosure())
            }

            $cmbSpawnPlayer.Add_TextChanged({
                $templateText = ''
                try {
                    if ($tbCmd.Tag -is [string]) { $templateText = [string]$tbCmd.Tag }
                } catch { $templateText = '' }
                if (-not [string]::IsNullOrWhiteSpace($templateText) -and $templateText -match '(?i)<(?:user|player)>') {
                    & $setCommandBoxFromTemplate $templateText
                }
            }.GetNewClosure())

            $cmbSpawnPlayer.Add_SelectedIndexChanged({
                $templateText = ''
                try {
                    if ($tbCmd.Tag -is [string]) { $templateText = [string]$tbCmd.Tag }
                } catch { $templateText = '' }
                if (-not [string]::IsNullOrWhiteSpace($templateText) -and $templateText -match '(?i)<(?:user|player)>') {
                    & $setCommandBoxFromTemplate $templateText
                }
            }.GetNewClosure())

            if (-not $hasInitialPzCatalogData) {
                $guiModulePath = Join-Path $PSScriptRoot 'GUI.psm1'

                $pzCatalogRunspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
                $pzCatalogRunspace.ApartmentState = 'MTA'
                $pzCatalogRunspace.ThreadOptions  = 'ReuseThread'
                $pzCatalogRunspace.Open()
                $pzCatalogRunspace.SessionStateProxy.SetVariable('GuiModulePath', $guiModulePath)
                $pzCatalogRunspace.SessionStateProxy.SetVariable('PzProfile', $capturedProfileForPzCatalog)

                $pzCatalogWorker = [System.Management.Automation.PowerShell]::Create()
                $pzCatalogWorker.Runspace = $pzCatalogRunspace
                $pzCatalogWorker.AddScript({
                    Set-StrictMode -Off
                    $ErrorActionPreference = 'Stop'
                    Import-Module $GuiModulePath -Force
                    $normalizedKnownGame = ''
                    try { $normalizedKnownGame = ([string]$PzProfile.KnownGame).ToLowerInvariant() } catch { $normalizedKnownGame = '' }
                    $catalogCacheCommand = if ($normalizedKnownGame -eq 'minecraft') {
                        'Get-MinecraftSpawnerCatalogsFromCache'
                    } elseif ($normalizedKnownGame -eq 'hytale') {
                        'Get-HytaleSpawnerCatalogsFromCache'
                    } else {
                        'Get-ProjectZomboidSpawnerCatalogsFromCache'
                    }
                    $catalogFullCommand = if ($normalizedKnownGame -eq 'minecraft') {
                        'Get-MinecraftSpawnerCatalogs'
                    } elseif ($normalizedKnownGame -eq 'hytale') {
                        'Get-HytaleSpawnerCatalogs'
                    } else {
                        'Get-ProjectZomboidSpawnerCatalogs'
                    }
                    $cached = & (Get-Command -Name $catalogCacheCommand -CommandType Function -ErrorAction Stop) -Profile $PzProfile
                    $cachedItemCount = if ($cached -and $cached.PSObject.Properties.Name -contains 'Items') { @($cached.Items).Count } else { 0 }
                    $cachedVehicleCount = if ($cached -and $cached.PSObject.Properties.Name -contains 'Vehicles') { @($cached.Vehicles).Count } else { 0 }
                    if ($cachedItemCount -gt 0 -or $cachedVehicleCount -gt 0) {
                        return $cached
                    }
                    return & (Get-Command -Name $catalogFullCommand -CommandType Function -ErrorAction Stop) -Profile $PzProfile
                }) | Out-Null
                $pzCatalogAsync = $pzCatalogWorker.BeginInvoke()

                $pzCatalogTimer = New-Object System.Windows.Forms.Timer
                $pzCatalogTimer.Interval = 250
                $pzCatalogTimer.Add_Tick({
                    if ($null -eq $pzCatalogAsync) { return }
                    if (-not $pzCatalogAsync.IsCompleted) { return }

                    $pzCatalogTimer.Stop()
                    try {
                        $catalogResult = $pzCatalogWorker.EndInvoke($pzCatalogAsync)
                        if ($catalogResult -is [System.Array]) {
                            $catalogResult = $catalogResult | Select-Object -Last 1
                        }
                        & $applyPzCatalogsToUi $catalogResult
                    } catch {
                        & $setPzCatalogUiEnabled $true
                        $pzCatalogLoaded = $true
                        $lblSpawnHint.Text = $itemSpawnerMissingHint
                        if ($supportsVehicleSpawner -and $lblVehicleHint) {
                            $lblVehicleHint.Text = 'ECC could not load local vehicle data.'
                        }
                    } finally {
                        try { $pzCatalogWorker.Dispose() } catch { }
                        try { $pzCatalogRunspace.Dispose() } catch { }
                        $pzCatalogWorker = $null
                        $pzCatalogRunspace = $null
                        $pzCatalogAsync = $null
                    }
                }.GetNewClosure())
                $pzCatalogTimer.Start()
            }
        }

        # Build command buttons
        $cmds = @($gameEntry.Commands)
        $cmds = $cmds | Sort-Object Category, Label

        $y = 10
        $currentCat = ''
        foreach ($cmd in $cmds) {
            if ($null -eq $cmd) { continue }
            $cat = if ($cmd.Category) { "$($cmd.Category)" } else { 'General' }

            if ($cat -ne $currentCat) {
                $currentCat = $cat
                $catLbl = _Label $currentCat 10 $y 300 18 $fontBold
                $catLbl.ForeColor = $clrAccent
                $listPanel.Controls.Add($catLbl)
                $y += 22
            }

            $btn = _Button ($cmd.Label) 10 $y 160 24 $clrAccent $null
            $btn.Tag = $cmd.Command
            $tipText = _BuildCommandTooltipText -Command $cmd
            $toolTip.SetToolTip($btn, $tipText)
            $btn.Add_Click({
                $cmdText = [string]$this.Tag
                & $setCommandBoxFromTemplate $cmdText
                if ($cmdText -match '^(GET|POST)\\s+/') {
                    $chkRest.Checked = $true
                } else {
                    $chkRest.Checked = $false
                }
            }.GetNewClosure())
            $listPanel.Controls.Add($btn)

            $lbl = _Label ($cmd.Command) 180 $y ($listPanel.ClientSize.Width - 200) 24 $fontMono
            $lbl.Tag = 'CmdLabel'
            $lbl.Anchor = 'Top,Left,Right'
            if ($tipText) { $toolTip.SetToolTip($lbl, $tipText) }
            $listPanel.Controls.Add($lbl)

            $y += 28
        }

        $commandsListPanelLocal = $listPanel
        $spawnPanelLocal = $spawnPanel
        $spawnTabControlLocal = $spawnTabControl
        $spawnFooterPanelLocal = $spawnFooterPanel
        $vehicleFooterPanelLocal = $vehicleFooterPanel
        $tbItemSearchLocal = $tbItemSearch
        $tbVehicleSearchLocal = $tbVehicleSearch
        $lbItemsLocal = $lbItems
        $lbVehiclesLocal = $lbVehicles
        $btnInsertAddItemLocal = $btnInsertAddItem
        $btnInsertAddVehicleLocal = $btnInsertAddVehicle
        $picItemPreviewLocal = $picItemPreview
        $picVehiclePreviewLocal = $picVehiclePreview
        $lblItemNameLocal = $lblItemName
        $lblItemTypeLocal = $lblItemType
        $lblSpawnHintLocal = $lblSpawnHint
        $lblVehicleNameLocal = $lblVehicleName
        $lblVehicleTypeLocal = $lblVehicleType
        $lblVehicleHintLocal = $lblVehicleHint
        $tabItemsLocal = $tabItems
        $tabVehiclesLocal = $tabVehicles
        $cmbSpawnPlayerLocal = $cmbSpawnPlayer
        $btnRefreshPlayersLocal = $btnRefreshPlayers
        $listPanel.Add_Resize({
            foreach ($c in $commandsListPanelLocal.Controls) {
                if ($c.Tag -eq 'CmdLabel') {
                    $c.Width = [Math]::Max(200, $commandsListPanelLocal.ClientSize.Width - 200)
                }
            }
        }.GetNewClosure())

        # Reflow bottom controls on resize
        $commandsFormLocal = $form
        $commandFooterPanelLocal = $commandFooterPanel
        $commandInputRowLocal = $commandInputRow
        $lblCmdLocal = $lblCmd
        $tbDebugLocal = $tbDebug
        $tbCmdLocal = $tbCmd
        $btnTestTelnetLocal = $btnTestTelnet
        $btnTestRconLocal = $btnTestRcon
        $btnTestApiLocal = $btnTestApi
        $btnTestStdinLocal = $btnTestStdin
        $btnTestPidLocal = $btnTestPid
        $btnSendLocal = $btnSend
        $btnTestRestLocal = $btnTestRest
        $chkVerboseLocal = $chkVerbose
        $chkApiLocal = $chkApi
        $chkRestLocal = $chkRest
        $diagnosticsRowLocal = $diagnosticsRow
        $optionsRowLocal = $optionsRow
        $commandActionsRowLocal = $commandActionsRow
        $lblStatusLocal = $lblStatus
        $layoutCommandsWindow = {
            $footerTop = $commandsFormLocal.ClientSize.Height - $commandsFooterHeight - $commandsFooterBottomMargin
            $contentHeight = [Math]::Max(180, $footerTop - 50)
            if ($commandsListPanelLocal -is [System.Windows.Forms.Control]) {
                $panelWidth = if ($spawnPanelLocal -is [System.Windows.Forms.Control]) {
                    [Math]::Max(320, $commandsFormLocal.ClientSize.Width - $spawnPanelLocal.Width - 20 - $commandsPanelGap)
                } else {
                    $commandsFormLocal.ClientSize.Width - 20
                }
                $commandsListPanelLocal.Size = [System.Drawing.Size]::new($panelWidth, $contentHeight)
            }
            if ($spawnPanelLocal -is [System.Windows.Forms.Control]) {
                $spawnPanelLocal.Location = [System.Drawing.Point]::new($commandsFormLocal.ClientSize.Width - $spawnPanelLocal.Width - 10, 40)
                $spawnPanelLocal.Height = $contentHeight
            }
            if ($commandFooterPanelLocal -is [System.Windows.Forms.Control]) {
                $commandFooterPanelLocal.Location = [System.Drawing.Point]::new(10, $footerTop)
                $commandFooterPanelLocal.Size = [System.Drawing.Size]::new($commandsFormLocal.ClientSize.Width - 20, $commandsFooterHeight)
                $commandFooterPanelLocal.BringToFront()
            }
            if ($spawnTabControlLocal -is [System.Windows.Forms.Control] -and $spawnPanelLocal -is [System.Windows.Forms.Control]) {
                $spawnTabControlLocal.Size = [System.Drawing.Size]::new($spawnPanelLocal.ClientSize.Width - 24, $spawnPanelLocal.ClientSize.Height - 108)
            }
            if ($spawnFooterPanelLocal -is [System.Windows.Forms.Control] -and $tabItemsLocal -is [System.Windows.Forms.Control]) {
                $spawnFooterPanelLocal.BringToFront()
            }
            if ($vehicleFooterPanelLocal -is [System.Windows.Forms.Control] -and $tabVehiclesLocal -is [System.Windows.Forms.Control]) {
                $vehicleFooterPanelLocal.BringToFront()
            }
            if ($btnInsertAddItemLocal -is [System.Windows.Forms.Control] -and $spawnFooterPanelLocal -is [System.Windows.Forms.Control]) {
                $btnInsertAddItemLocal.Location = [System.Drawing.Point]::new(0, $spawnActionButtonY)
            }
            if ($btnInsertAddVehicleLocal -is [System.Windows.Forms.Control] -and $vehicleFooterPanelLocal -is [System.Windows.Forms.Control]) {
                $btnInsertAddVehicleLocal.Location = [System.Drawing.Point]::new(0, $spawnActionButtonY)
            }
            if ($cmbSpawnPlayerLocal -is [System.Windows.Forms.Control] -and $btnRefreshPlayersLocal -is [System.Windows.Forms.Control] -and $spawnPanelLocal -is [System.Windows.Forms.Control]) {
                $btnRefreshPlayersLocal.Location = [System.Drawing.Point]::new($spawnPanelLocal.ClientSize.Width - $btnRefreshPlayersLocal.Width - 12, 61)
                $cmbSpawnPlayerLocal.Width = [Math]::Max(120, $btnRefreshPlayersLocal.Left - 20)
            }
            if ($lblItemNameLocal -is [System.Windows.Forms.Control] -and $spawnFooterPanelLocal -is [System.Windows.Forms.Control]) {
                $lblItemNameLocal.Location = [System.Drawing.Point]::new(0, 4)
                $lblItemNameLocal.Width = [Math]::Max(140, $spawnFooterPanelLocal.ClientSize.Width - $spawnPreviewColumnWidth - 8)
            }
            if ($lblItemTypeLocal -is [System.Windows.Forms.Control] -and $spawnFooterPanelLocal -is [System.Windows.Forms.Control]) {
                $lblItemTypeLocal.Location = [System.Drawing.Point]::new(0, 24)
                $lblItemTypeLocal.Width = [Math]::Max(140, $spawnFooterPanelLocal.ClientSize.Width - $spawnPreviewColumnWidth - 8)
                $lblItemTypeLocal.Height = 30
            }
            if ($lblSpawnHintLocal -is [System.Windows.Forms.Control] -and $spawnFooterPanelLocal -is [System.Windows.Forms.Control]) {
                $lblSpawnHintLocal.Location = [System.Drawing.Point]::new(0, 52)
                $lblSpawnHintLocal.Width = [Math]::Max(140, $spawnFooterPanelLocal.ClientSize.Width - $spawnPreviewColumnWidth - 8)
                $lblSpawnHintLocal.Height = 40
            }
            if ($picItemPreviewLocal -is [System.Windows.Forms.Control] -and $spawnFooterPanelLocal -is [System.Windows.Forms.Control]) {
                $picItemPreviewLocal.Location = [System.Drawing.Point]::new($spawnFooterPanelLocal.ClientSize.Width - $spawnPreviewColumnWidth + 8, 8)
                $picItemPreviewLocal.Size = [System.Drawing.Size]::new($spawnPreviewColumnWidth - 16, 70)
            }
            if ($lblVehicleNameLocal -is [System.Windows.Forms.Control] -and $vehicleFooterPanelLocal -is [System.Windows.Forms.Control]) {
                $lblVehicleNameLocal.Location = [System.Drawing.Point]::new(0, 4)
                $lblVehicleNameLocal.Width = [Math]::Max(140, $vehicleFooterPanelLocal.ClientSize.Width - $spawnPreviewColumnWidth - 8)
            }
            if ($lblVehicleTypeLocal -is [System.Windows.Forms.Control] -and $vehicleFooterPanelLocal -is [System.Windows.Forms.Control]) {
                $lblVehicleTypeLocal.Location = [System.Drawing.Point]::new(0, 24)
                $lblVehicleTypeLocal.Width = [Math]::Max(140, $vehicleFooterPanelLocal.ClientSize.Width - $spawnPreviewColumnWidth - 8)
                $lblVehicleTypeLocal.Height = 30
            }
            if ($lblVehicleHintLocal -is [System.Windows.Forms.Control] -and $vehicleFooterPanelLocal -is [System.Windows.Forms.Control]) {
                $lblVehicleHintLocal.Location = [System.Drawing.Point]::new(0, 52)
                $lblVehicleHintLocal.Width = [Math]::Max(140, $vehicleFooterPanelLocal.ClientSize.Width - $spawnPreviewColumnWidth - 8)
                $lblVehicleHintLocal.Height = 40
            }
            if ($picVehiclePreviewLocal -is [System.Windows.Forms.Control] -and $vehicleFooterPanelLocal -is [System.Windows.Forms.Control]) {
                $picVehiclePreviewLocal.Location = [System.Drawing.Point]::new($vehicleFooterPanelLocal.ClientSize.Width - $spawnPreviewColumnWidth + 8, 8)
                $picVehiclePreviewLocal.Size = [System.Drawing.Size]::new($spawnPreviewColumnWidth - 16, 70)
            }
            if ($commandInputRowLocal -is [System.Windows.Forms.Control] -and $commandFooterPanelLocal -is [System.Windows.Forms.Control]) {
                $commandInputRowLocal.Location = [System.Drawing.Point]::new(0, 154)
                $commandInputRowLocal.Size = [System.Drawing.Size]::new($commandFooterPanelLocal.ClientSize.Width, 32)
            }
            if ($lblCmdLocal -is [System.Windows.Forms.Control]) {
                $lblCmdLocal.Location = [System.Drawing.Point]::new(0, 0)
                if ($commandFooterPanelLocal -is [System.Windows.Forms.Control]) {
                    $lblCmdLocal.Width = $commandFooterPanelLocal.ClientSize.Width
                }
            }
            if ($tbDebugLocal -is [System.Windows.Forms.Control]) {
                $tbDebugLocal.Location = [System.Drawing.Point]::new(0, 24)
                if ($commandFooterPanelLocal -is [System.Windows.Forms.Control]) {
                    $tbDebugLocal.Size = [System.Drawing.Size]::new($commandFooterPanelLocal.ClientSize.Width, 70)
                }
            }
            if ($tbCmdLocal -is [System.Windows.Forms.Control]) {
                $commandActionsWidth = if ($commandActionsRowLocal -is [System.Windows.Forms.Control]) { [Math]::Max($commandActionsRowLocal.Width, $commandActionsRowLocal.PreferredSize.Width) } else { 220 }
                $inputWidth = if ($commandInputRowLocal -is [System.Windows.Forms.Control]) { $commandInputRowLocal.ClientSize.Width } else { $commandsFormLocal.ClientSize.Width - 20 }
                $tbCmdLocal.Location = [System.Drawing.Point]::new(0, 2)
                $tbCmdLocal.Size = [System.Drawing.Size]::new([Math]::Max(160, $inputWidth - $commandActionsWidth - 10), 26)
            }
            if ($diagnosticsRowLocal -is [System.Windows.Forms.Control]) {
                $diagWidth = [Math]::Max($diagnosticsRowLocal.Width, $diagnosticsRowLocal.PreferredSize.Width)
                $footerWidth = if ($commandFooterPanelLocal -is [System.Windows.Forms.Control]) { $commandFooterPanelLocal.ClientSize.Width } else { $commandsFormLocal.ClientSize.Width - 20 }
                $diagnosticsRowLocal.Location = [System.Drawing.Point]::new([Math]::Max(0, $footerWidth - $diagWidth), 104)
            }
            if ($optionsRowLocal -is [System.Windows.Forms.Control]) {
                $optionsWidth = [Math]::Max($optionsRowLocal.Width, $optionsRowLocal.PreferredSize.Width)
                $footerWidth = if ($commandFooterPanelLocal -is [System.Windows.Forms.Control]) { $commandFooterPanelLocal.ClientSize.Width } else { $commandsFormLocal.ClientSize.Width - 20 }
                $optionsRowLocal.Location = [System.Drawing.Point]::new([Math]::Max(0, $footerWidth - $optionsWidth), 132)
            }
            if ($commandActionsRowLocal -is [System.Windows.Forms.Control]) {
                $actionsWidth = [Math]::Max($commandActionsRowLocal.Width, $commandActionsRowLocal.PreferredSize.Width)
                $inputWidth = if ($commandInputRowLocal -is [System.Windows.Forms.Control]) { $commandInputRowLocal.ClientSize.Width } else { $commandsFormLocal.ClientSize.Width - 20 }
                $commandActionsRowLocal.Location = [System.Drawing.Point]::new([Math]::Max(0, $inputWidth - $actionsWidth), 0)
            }
            if ($lblStatusLocal -is [System.Windows.Forms.Control]) {
                $lblStatusLocal.Location = [System.Drawing.Point]::new(0, 190)
                if ($commandFooterPanelLocal -is [System.Windows.Forms.Control]) {
                    $lblStatusLocal.Width = $commandFooterPanelLocal.ClientSize.Width
                }
            }
        }.GetNewClosure()
        $form.add_Resize($layoutCommandsWindow)
        $form.Add_Shown({
            & $layoutCommandsWindow
        }.GetNewClosure())

        $form.Add_FormClosed({
            try {
                if ($pzCatalogTimer) { $pzCatalogTimer.Stop(); $pzCatalogTimer.Dispose() }
            } catch { }
            try {
                if ($pzCatalogWorker) { $pzCatalogWorker.Dispose() }
            } catch { }
            try {
                if ($pzCatalogRunspace) { $pzCatalogRunspace.Dispose() }
            } catch { }
        }.GetNewClosure())

        $form.ShowDialog() | Out-Null
    }

    # =====================================================================
    # BACKGROUND SERVER OPERATION HELPER
    # All server operations (Start / Stop / Restart) that involve any
    # waiting — save wait, process kill, WaitForExit — MUST run here.
    # Calling them on the GUI/STA thread freezes WinForms for the entire
    # duration of the save-wait (15-30 s) because PowerShell 5.1 STA
    # blocks the message pump during Start-Sleep / WaitForExit.
    #
    # Usage:
    #   _RunServerOpInBackground -Prefix 'PZ' -Operation 'Restart'
    #   _RunServerOpInBackground -Prefix 'PW' -Operation 'Stop'
    #   _RunServerOpInBackground -Prefix 'VH' -Operation 'Start'
    # =====================================================================
    function _RunServerOpInBackground {
        param(
            [string]$Prefix,
            [string]$Operation   # 'Start' | 'Stop' | 'Restart'
        )

        # Capture everything the background runspace needs BEFORE it starts.
        # Variables set on the proxy must be set before Open() is called.
        $capturedModulesDir  = $script:ModuleRoot
        $capturedSharedState = $script:SharedState
        $capturedPrefix      = $Prefix
        $capturedOperation   = $Operation

        $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $rs.ApartmentState = 'MTA'   # MTA: no STA message-pump interference
        $rs.ThreadOptions  = 'ReuseThread'
        $rs.Open()
        $rs.SessionStateProxy.SetVariable('ModulesDir',   $capturedModulesDir)
        $rs.SessionStateProxy.SetVariable('SharedState',  $capturedSharedState)
        $rs.SessionStateProxy.SetVariable('TargetPrefix', $capturedPrefix)
        $rs.SessionStateProxy.SetVariable('OpName',       $capturedOperation)

        $ps          = [System.Management.Automation.PowerShell]::Create()
        $ps.Runspace = $rs
        $ps.AddScript({
            Set-StrictMode -Off
            $ErrorActionPreference = 'Continue'
            try {
                Import-Module (Join-Path $ModulesDir 'Logging.psm1')        -Force
                Import-Module (Join-Path $ModulesDir 'ProfileManager.psm1') -Force
                Import-Module (Join-Path $ModulesDir 'ServerManager.psm1')  -Force
                Initialize-ServerManager -SharedState $SharedState

                switch ($OpName) {
                    'Start'   { Start-GameServer      -Prefix $TargetPrefix | Out-Null }
                    'Stop'    { Invoke-SafeShutdown    -Prefix $TargetPrefix | Out-Null }
                    'Restart' { Restart-GameServer     -Prefix $TargetPrefix | Out-Null }
                    default   {
                        $SharedState.LogQueue.Enqueue(
                            "[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][WARN][GUI] Unknown background op: $OpName")
                    }
                }
            } catch {
                if ($SharedState -and $SharedState.LogQueue) {
                    $SharedState.LogQueue.Enqueue(
                        "[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][ERROR][GUI] Background $OpName for $TargetPrefix failed: $_")
                }
            } finally {
                # Always clean up the runspace when done, even on error.
                try { $ps.Dispose() } catch { }
                try { $rs.Close();  $rs.Dispose() } catch { }
            }
        }) | Out-Null

        # Fire-and-forget: BeginInvoke returns immediately; GUI stays responsive.
        $ps.BeginInvoke() | Out-Null
    }

    function _FormatBulkProfileSummary {
        param([string[]]$Names = @())

        $safeNames = @($Names | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
        if ($safeNames.Count -le 0) { return '<none>' }
        if ($safeNames.Count -le 4) { return ($safeNames -join ', ') }

        $preview = @($safeNames | Select-Object -First 4)
        $remaining = [Math]::Max(0, $safeNames.Count - $preview.Count)
        return '{0}, +{1} more' -f ($preview -join ', '), $remaining
    }

    function _GetBulkServerOperationTargets {
        param(
            [ValidateSet('Start','Stop')][string]$Operation,
            [hashtable]$SharedState = $script:SharedState
        )

        $prefixes = New-Object System.Collections.Generic.List[string]
        $names    = New-Object System.Collections.Generic.List[string]

        if (-not $SharedState -or -not $SharedState.Profiles) {
            return [pscustomobject]@{
                Prefixes = @()
                Names    = @()
                Count    = 0
                Skipped  = 0
                Total    = 0
            }
        }

        $entries = @($SharedState.Profiles.GetEnumerator() | Sort-Object {
            if ($_.Value -and $_.Value.GameName) { [string]$_.Value.GameName } else { [string]$_.Key }
        })

        foreach ($entry in $entries) {
            $pfx = [string]$entry.Key
            $profile = $entry.Value
            $running = $false
            $runtimeCode = ''

            try {
                $status = Get-ServerStatus -Prefix $pfx
                $running = ($status -and $status.Running)
            } catch { $running = $false }

            try {
                $runtime = _GetRuntimeStateEntry -Prefix $pfx -SharedState $SharedState
                if ($runtime) { $runtimeCode = [string]$runtime.Code }
            } catch { $runtimeCode = '' }
            $runtimeCode = $runtimeCode.ToLowerInvariant()

            $eligible = $false
            switch ($Operation) {
                'Start' {
                    if (-not $running -and $runtimeCode -notin @('starting','restarting','stopping','waiting_restart','waiting_first_player','idle_wait','idle_shutdown','online','blocked')) {
                        $eligible = $true
                    }
                }
                'Stop' {
                    if ($running) {
                        $eligible = $true
                    }
                }
            }

            if ($eligible) {
                $prefixes.Add($pfx) | Out-Null
                $displayName = if ($profile -and $profile.GameName) { [string]$profile.GameName } else { $pfx }
                $names.Add($displayName) | Out-Null
            }
        }

        return [pscustomobject]@{
            Prefixes = @($prefixes.ToArray())
            Names    = @($names.ToArray())
            Count    = $prefixes.Count
            Skipped  = [Math]::Max(0, $entries.Count - $prefixes.Count)
            Total    = $entries.Count
        }
    }

    function _RunBulkServerOperations {
        param(
            [ValidateSet('Start','Stop')][string]$Operation,
            [string[]]$Prefixes = @()
        )

        $targets = @($Prefixes | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
        if ($targets.Count -le 0) { return }

        $runServerOpCommand = Get-Command -Name '_RunServerOpInBackground' -CommandType Function -ErrorAction SilentlyContinue
        if (-not $runServerOpCommand) {
            _QueueStatusMessage 'Bulk server action is temporarily unavailable because the background launcher helper is missing.'
            try {
                if ($script:SharedState -and $script:SharedState.LogQueue) {
                    $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][WARN][GUI] Bulk server operation skipped because _RunServerOpInBackground could not be resolved.")
                }
            } catch { }
            return
        }

        if ($Operation -eq 'Start') {
            if ($script:_BulkStartTimer -and $script:_BulkStartTimer.Enabled) {
                _QueueStatusMessage 'A staggered Start All batch is already in progress.'
                return
            }

            $startQueue = [System.Collections.Generic.Queue[string]]::new()
            foreach ($targetPrefix in $targets) {
                $startQueue.Enqueue([string]$targetPrefix)
            }

            $dispatchNext = {
                if ($startQueue.Count -le 0) {
                    try {
                        if ($script:_BulkStartTimer) {
                            $script:_BulkStartTimer.Stop()
                            $script:_BulkStartTimer.Dispose()
                        }
                    } catch { }
                    $script:_BulkStartTimer = $null
                    return
                }

                $nextPrefix = [string]$startQueue.Dequeue()
                & $runServerOpCommand -Prefix $nextPrefix -Operation 'Start'

                if ($startQueue.Count -le 0) {
                    try {
                        if ($script:_BulkStartTimer) {
                            $script:_BulkStartTimer.Stop()
                            $script:_BulkStartTimer.Dispose()
                        }
                    } catch { }
                    $script:_BulkStartTimer = $null
                }
            }.GetNewClosure()

            $script:_BulkStartTimer = New-Object System.Windows.Forms.Timer
            $script:_BulkStartTimer.Interval = 1500
            $script:_BulkStartTimer.Add_Tick($dispatchNext)

            & $dispatchNext
            if ($script:_BulkStartTimer) {
                $script:_BulkStartTimer.Start()
            }
            return
        }

        foreach ($targetPrefix in $targets) {
            & $runServerOpCommand -Prefix ([string]$targetPrefix) -Operation $Operation
        }
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
            $emptyCard = _Panel 14 14 ($panel.Width - 28) 100 $clrPanelSoft
            $emptyCard.Anchor = 'Top,Left,Right'
            $emptyCard.BorderStyle = 'FixedSingle'
            $panel.Controls.Add($emptyCard)
            $emptyTitle = _Label 'Profile Editor' 14 14 260 22 $fontTitle
            $emptyTitle.ForeColor = $clrText
            $emptyCard.Controls.Add($emptyTitle)
            $emptySub = _Label 'Select a server from the dashboard to edit its profile, commands, paths, and runtime settings.' 14 44 ($emptyCard.Width - 28) 34 $fontLabel
            $emptySub.ForeColor = $clrTextSoft
            $emptySub.Anchor = 'Top,Left,Right'
            $emptyCard.Controls.Add($emptySub)
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

        $contentX = 18
        $fieldX = 22
        $lw  = 172
        $tw  = [Math]::Max(260, $scroll.ClientSize.Width - ($lw + 104))
        $th  = 24
        $gap = 34
        $helpGap = 22

        $editorToolTip = New-Object System.Windows.Forms.ToolTip
        $editorToolTip.AutoPopDelay = 12000
        $editorToolTip.InitialDelay = 350
        $editorToolTip.ReshowDelay  = 150
        $editorToolTip.ShowAlways   = $true

        function _FormatKeyLabel([string]$key) {
            return $key -replace '([a-z])([A-Z])', '$1 $2'
        }

        function _MeasureProfileHelpHeight([string]$text, [int]$width, [System.Drawing.Font]$font) {
            if ([string]::IsNullOrWhiteSpace($text)) { return 18 }
            $proposed = [System.Drawing.Size]::new([Math]::Max(120, $width), 0)
            try {
                $measured = [System.Windows.Forms.TextRenderer]::MeasureText(
                    $text,
                    $font,
                    $proposed,
                    [System.Windows.Forms.TextFormatFlags]::WordBreak
                )
                return [Math]::Max(18, $measured.Height + 2)
            } catch {
                return 30
            }
        }

        function _GetProfileFieldMeta([string]$key) {
            $map = @{
                GameName = @{
                    Label = 'Profile display name'
                    Help  = 'The name people see for this server in ECC, in Discord, and in the profile list.'
                }
                Prefix = @{
                    Label = 'Command prefix'
                    Help  = 'Short code used in bot commands, like PZ or VH. Keep it short and unique.'
                }
                ProcessName = @{
                    Label = 'Process name to watch'
                    Help  = 'The process name ECC watches to tell whether this server is really running.'
                }
                Executable = @{
                    Label = 'Server executable'
                    Help  = 'The file ECC starts for this server. This can be an exe, bat, cmd, or jar.'
                }
                FolderPath = @{
                    Label = 'Server folder'
                    Help  = 'Main folder where this server is installed.'
                }
                MaxRamGB = @{
                    Label = 'Max memory (GB)'
                    Help  = 'If the launch setup supports memory limits, ECC uses this as the upper RAM target.'
                }
                MinRamGB = @{
                    Label = 'Min memory (GB)'
                    Help  = 'If the launch setup supports memory limits, ECC uses this as the minimum RAM target.'
                }
                AOTCache = @{
                    Label = 'AOT cache folder'
                    Help  = 'Optional cache file or folder used by some games, like Hytale, when that launch mode is turned on.'
                }
                LogStrategy = @{
                    Label = 'Log discovery mode'
                    Help  = 'Tells ECC how to find the live server log, like one fixed file, the newest file, or a game-specific folder rule.'
                    Options = @(
                        @{ Value = 'SingleFile';        Label = 'Single file (fixed path)' }
                        @{ Value = 'NewestFile';        Label = 'Newest file in folder' }
                        @{ Value = 'PZSessionFolder';   Label = 'Project Zomboid session folder' }
                        @{ Value = 'ValheimUserFolder'; Label = 'Valheim user log folder' }
                    )
                }
                ServerLogRoot = @{
                    Label = 'Log root folder'
                    Help  = 'Main folder ECC searches when it needs to find the live server log.'
                }
                ServerLogSubDir = @{
                    Label = 'Log subfolder'
                    Help  = 'Extra folder under the main log folder that ECC should check first.'
                }
                ServerLogFile = @{
                    Label = 'Preferred log filename'
                    Help  = 'Log file name or pattern ECC should try first when it looks for logs.'
                }
                ServerLogPath = @{
                    Label = 'Resolved log path'
                    Help  = 'Full path to the log file when this server uses one fixed log file.'
                }
                ServerLogNote = @{
                    Label = 'Log note'
                    Help  = 'Extra note about unusual log behavior. Most people can leave this alone.'
                }
                RestEnabled = @{
                    Label     = 'REST/API control enabled'
                    Help      = 'Turns on web API control for games that support it.'
                    BoolLabel = 'Use the REST/API control path'
                }
                RestHost = @{
                    Label = 'REST/API host'
                    Help  = 'Host name or IP ECC uses when it calls the server API.'
                }
                RestPort = @{
                    Label = 'REST/API port'
                    Help  = 'Network port ECC uses for REST or HTTP server control.'
                }
                RestPassword = @{
                    Label = 'REST/API password or key'
                    Help  = 'Password, token, or API key ECC sends when it talks to the server API.'
                }
                RestProtocol = @{
                    Label = 'REST/API protocol'
                    Help  = 'Whether ECC connects to the API over http or https.'
                    Options = @(
                        @{ Value = 'http';  Label = 'HTTP' }
                        @{ Value = 'https'; Label = 'HTTPS' }
                    )
                }
                EnableAutoRestart = @{
                    Label     = 'Auto-restart after crash'
                    Help      = 'If the server crashes or closes by accident, ECC can try to start it again for you.'
                    BoolLabel = 'Restart it automatically'
                }
                RestartDelaySeconds = @{
                    Label = 'Crash restart delay (seconds)'
                    Help  = 'How long ECC waits after a crash before it tries to start the server again.'
                }
                MaxRestartsPerHour = @{
                    Label = 'Max crash restarts per hour'
                    Help  = 'Safety limit for crash loops. Lower this if you want ECC to stop retrying sooner.'
                }
                BlockStartIfRamPercentUsed = @{
                    Label = 'Block start if RAM used is above (%)'
                    Help  = 'ECC will not start this server if total system memory use is already above this percent. Use 0 to turn this off.'
                }
                BlockStartIfFreeRamBelowGB = @{
                    Label = 'Block start if free RAM is below (GB)'
                    Help  = 'ECC will not start this server if the machine has less free memory than this amount. Use 0 to turn this off.'
                }
                StartupTimeoutSeconds = @{
                    Label = 'Startup ready timeout (seconds)'
                    Help  = 'How long ECC waits for the server to become ready before it marks startup as failed.'
                }
                ShutdownIfNoPlayersAfterStartupMinutes = @{
                    Label = 'Shut down if nobody joins within (minutes)'
                    Help  = 'After startup, ECC can shut the server down if nobody joins before this timer ends. Use 0 to turn this off.'
                }
                ShutdownIfEmptyAfterLastPlayerLeavesMinutes = @{
                    Label = 'Shut down after last player leaves (minutes)'
                    Help  = 'When the server becomes empty, ECC can wait this many minutes and then shut it down. Use 0 to turn this off.'
                }
                SaveMethod = @{
                    Label = 'Save method'
                    Help  = 'How ECC saves the server before a restart or shutdown.'
                    Options = @(
                        @{ Value = 'none';            Label = 'None (skip save step)' }
                        @{ Value = 'stdin';           Label = 'Server console / STDIN' }
                        @{ Value = 'http';            Label = 'HTTP endpoint' }
                        @{ Value = 'rest';            Label = 'REST/API endpoint' }
                        @{ Value = 'SatisfactoryApi'; Label = 'Satisfactory API' }
                    )
                }
                SaveWaitSeconds = @{
                    Label = 'Save wait (seconds)'
                    Help  = 'How long ECC waits after sending the save command before it keeps going.'
                }
                StopMethod = @{
                    Label = 'Stop method'
                    Help  = 'How ECC shuts the server down when you stop or restart it.'
                    Options = @(
                        @{ Value = 'processKill';     Label = 'Kill launcher process tree' }
                        @{ Value = 'processName';     Label = 'Kill by process name' }
                        @{ Value = 'stdin';           Label = 'Server console / STDIN' }
                        @{ Value = 'ctrlc';           Label = 'Send Ctrl+C' }
                        @{ Value = 'http';            Label = 'HTTP endpoint' }
                        @{ Value = 'SatisfactoryApi'; Label = 'Satisfactory API' }
                    )
                }
                RestPollOnlyWhenRunning = @{
                    Label     = 'Poll REST only while running'
                    Help      = 'Only checks the server API while the server is running.'
                    BoolLabel = 'Only poll while the server is running'
                }
                RconHost = @{
                    Label = 'RCON host'
                    Help  = 'Host name or IP ECC uses when it talks to the server over RCON.'
                }
                RconPort = @{
                    Label = 'RCON port'
                    Help  = 'Network port ECC uses for RCON commands.'
                }
                RconPassword = @{
                    Label = 'RCON password'
                    Help  = 'Password ECC uses when authenticating to the game server over RCON.'
                }
                ConfigRoot = @{
                    Label = 'Primary config folder'
                    Help  = 'Main config folder ECC uses when it reads or writes server config files.'
                }
                ConfigRoots = @{
                    Label = 'Additional config folders'
                    Help  = 'Extra config folders ECC should also check. Leave this alone unless the server really uses more than one config place.'
                }
                Commands = @{
                    Label = 'Base command map'
                    Help  = 'The main command set ECC uses for actions like start, stop, restart, status, and players.'
                }
                ExtraCommands = @{
                    Label = 'Extra command map'
                    Help  = 'Extra custom bot or console commands for this server beyond the normal control actions.'
                }
                StdinSaveCommand = @{
                    Label = 'STDIN save command'
                    Help  = 'Exact text ECC sends into the server console when save method uses STDIN.'
                }
                StdinStopCommand = @{
                    Label = 'STDIN stop command'
                    Help  = 'Exact text ECC sends into the server console when stop method uses STDIN.'
                }
                ExeHints = @{
                    Label = 'Executable hints'
                    Help  = 'Extra file or process names ECC can use to recognize this server during detection.'
                }
                AssetFile = @{
                    Label = 'Asset package file'
                    Help  = 'Optional asset or support file this profile expects to exist with the server.'
                }
                BackupDir = @{
                    Label = 'Backup folder'
                    Help  = 'Folder ECC or the game uses for backups, snapshots, or exported saves.'
                }
                HideConsoleWindow = @{
                    Label     = 'Hide console window'
                    Help      = 'Hide the extra server console window when ECC launches this server. Useful for games that would otherwise leave an extra command window open.'
                    BoolLabel = 'Hide the launcher console window'
                }
                KnownGame = @{
                    Label = 'Known game identity'
                    Help  = 'Internal game type ECC matched this profile to. This controls game-specific features, commands, and detection rules.'
                }
                MinecraftLaunchDetail = @{
                    Label = 'Minecraft launch detail'
                    Help  = 'Extra note about how ECC resolved the Minecraft startup path, like a wrapper script or pack-managed launch.'
                }
                MinecraftLaunchSource = @{
                    Label = 'Minecraft launch source'
                    Help  = 'The file or launcher source ECC resolved for Minecraft startup, such as run.bat or a server jar wrapper.'
                }
                MinecraftLaunchType = @{
                    Label = 'Minecraft launch type'
                    Help  = 'Internal launch mode ECC detected for Minecraft, such as direct jar launch or pack-wrapper startup.'
                }
                SatisfactoryApiHost = @{
                    Label = 'Satisfactory API host'
                    Help  = 'Host name or IP ECC uses when it talks to the Satisfactory server API.'
                }
                SatisfactoryApiPort = @{
                    Label = 'Satisfactory API port'
                    Help  = 'Network port ECC uses for the Satisfactory server API.'
                }
                SatisfactoryApiToken = @{
                    Label = 'Satisfactory API token'
                    Help  = 'API token ECC uses when authenticating to the Satisfactory server API.'
                }
                SteamAppId = @{
                    Label = 'Steam app ID'
                    Help  = 'Steam application ID tied to this dedicated server. ECC can use this for game-specific identification or tooling.'
                }
                TelnetHost = @{
                    Label = 'Telnet host'
                    Help  = 'Host name or IP ECC uses when a game supports telnet control, like 7 Days to Die.'
                }
                TelnetPassword = @{
                    Label = 'Telnet password'
                    Help  = 'Password ECC uses when connecting to a telnet-enabled game server.'
                }
                TelnetPort = @{
                    Label = 'Telnet port'
                    Help  = 'Network port ECC uses for telnet control when the game supports it.'
                }
                StdinPreferWindow = @{
                    Label     = 'Prefer window-based STDIN'
                    Help      = 'Tell ECC to send console input through the server window instead of normal redirected STDIN. Some servers, like Valheim, respond better to this mode.'
                    BoolLabel = 'Use window-based console input'
                }
                StdinWindowProcessName = @{
                    Label = 'STDIN window process name'
                    Help  = 'Process name ECC looks for when it needs to find the server window for window-based console input.'
                }
                CaptureOutput = @{
                    Label     = 'Capture output'
                    Help      = "Capture the server console output into ECC's log file. Use this for servers that do not write a good live log on their own."
                    BoolLabel = 'Capture server console output'
                }
                DisableFileTail = @{
                    Label     = 'Disable file tail'
                    Help      = "Do not tail a live log file for this server. Use this when the game does not keep a stable log file, or when you only want ECC's activity feed."
                    BoolLabel = 'Disable live file tailing'
                }
            }
            if ($map.ContainsKey($key)) { return $map[$key] }
            return $null
        }

        function Add-FieldHelp([string]$helpText, [int]$y, [System.Windows.Forms.Control[]]$targets) {
            if ([string]::IsNullOrWhiteSpace($helpText)) { return }

            $helpHeight = _MeasureProfileHelpHeight -text $helpText -width $tw -font $fontLabel
            $lblHelp = _Label $helpText ($lw + 38) $y $tw $helpHeight $fontLabel
            $lblHelp.ForeColor = $clrTextSoft
            $lblHelp.Anchor = 'Top,Left,Right'
            $scroll.Controls.Add($lblHelp)

            foreach ($target in @($targets)) {
                if ($target) { $editorToolTip.SetToolTip($target, $helpText) }
            }
            $editorToolTip.SetToolTip($lblHelp, $helpText)
            $script:y += [Math]::Max($helpGap, $helpHeight)
        }

        function Add-DynamicField ([string]$key, [object]$value) {
            $meta = _GetProfileFieldMeta $key
            $label = if ($meta -and $meta.Label) { $meta.Label } else { _FormatKeyLabel $key }
            $lblField = _Label $label $fieldX $script:y $lw 20 $fontBold
            $lblField.ForeColor = $clrTextSoft
            $scroll.Controls.Add($lblField)

            # Complex values get a JSON text area for full visibility/editing.
            $isDict = $value -is [System.Collections.IDictionary]
            $isList = ($value -is [System.Collections.IEnumerable] -and -not ($value -is [string]))

            if ($value -is [bool]) {
                $chk           = New-Object System.Windows.Forms.CheckBox
                $chk.Location  = [System.Drawing.Point]::new($lw + 38, $script:y)
                $chk.Size      = [System.Drawing.Size]::new([Math]::Max(220, $tw), 20)
                $chk.Text      = if ($meta -and $meta.BoolLabel) { $meta.BoolLabel } elseif ($key -eq 'RestPollOnlyWhenRunning') { 'Only when running' } else { 'Enabled' }
                $chk.ForeColor = $clrText
                $chk.BackColor = [System.Drawing.Color]::Transparent
                $chk.Checked   = ($value -eq $true)
                $scroll.Controls.Add($chk)
                $script:_ProfileFields[$key] = @{ Control = $chk; Kind = 'bool' }
                $script:y += $gap
                Add-FieldHelp -helpText $meta.Help -y ($script:y - 12) -targets @($lblField, $chk)
                return
            }

            if ($isDict -or $isList) {
                $json = ''
                try { $json = $value | ConvertTo-Json -Depth 6 } catch { $json = '' }
                $tb = New-Object System.Windows.Forms.TextBox
                $tb.Location    = [System.Drawing.Point]::new($lw + 38, $script:y)
                $tb.Size        = [System.Drawing.Size]::new($tw, 80)
                $tb.Multiline   = $true
                $tb.ScrollBars  = 'Both'
                $tb.WordWrap    = $false
                $tb.Text        = $json
                $tb.BackColor   = $clrPanelSoft
                $tb.ForeColor   = $clrText
                $tb.BorderStyle = 'FixedSingle'
                $tb.Font        = $fontMono
                $tb.Anchor      = 'Top,Left,Right'
                $scroll.Controls.Add($tb)
                $script:_ProfileFields[$key] = @{ Control = $tb; Kind = 'json' }
                $script:y += 92
                Add-FieldHelp -helpText $meta.Help -y ($script:y - 10) -targets @($lblField, $tb)
                return
            }

            $tb        = _TextBox ($lw + 38) $script:y $tw $th ([string]$value) $false
            $tb.Anchor = 'Top,Left,Right'
            $scroll.Controls.Add($tb)
            $script:_ProfileFields[$key] = @{ Control = $tb; Kind = 'text' }
            $script:y += $gap
            Add-FieldHelp -helpText $meta.Help -y ($script:y - 12) -targets @($lblField, $tb)
        }

        function Add-SelectField ([string]$key, [object]$value, [string[]]$options) {
            $meta = _GetProfileFieldMeta $key
            $label = if ($meta -and $meta.Label) { $meta.Label } else { _FormatKeyLabel $key }
            $lblField = _Label $label $fieldX $script:y $lw 20 $fontBold
            $lblField.ForeColor = $clrTextSoft
            $scroll.Controls.Add($lblField)

            $cb = New-Object System.Windows.Forms.ComboBox
            $cb.Location  = [System.Drawing.Point]::new($lw + 38, $script:y - 2)
            $cb.Size      = [System.Drawing.Size]::new([Math]::Max(220, $tw), 24)
            $cb.DropDownStyle = 'DropDownList'
            $cb.BackColor = $clrPanelSoft
            $cb.ForeColor = $clrText
            $cb.FlatStyle = 'Popup'
            $optionDefs = @()
            if ($meta -and $meta.Options) {
                foreach ($opt in @($meta.Options)) {
                    $optionDefs += [pscustomobject]@{
                        Value = [string]$opt.Value
                        Label = [string]$opt.Label
                    }
                }
            } else {
                foreach ($opt in $options) {
                    $optionDefs += [pscustomobject]@{
                        Value = [string]$opt
                        Label = [string]$opt
                    }
                }
            }

            $cb.DisplayMember = 'Label'
            $cb.ValueMember   = 'Value'
            foreach ($opt in $optionDefs) { [void]$cb.Items.Add($opt) }

            $selectedValue = if ($null -ne $value) { [string]$value } else { '' }
            if (-not [string]::IsNullOrWhiteSpace($selectedValue)) {
                for ($i = 0; $i -lt $cb.Items.Count; $i++) {
                    $item = $cb.Items[$i]
                    if ($item -and "$($item.Value)" -eq $selectedValue) {
                        $cb.SelectedIndex = $i
                        break
                    }
                }
            }
            if ($cb.SelectedIndex -lt 0 -and $cb.Items.Count -gt 0) {
                $cb.SelectedIndex = 0
            }
            $cb.Anchor = 'Top,Left,Right'
            $scroll.Controls.Add($cb)
            $script:_ProfileFields[$key] = @{ Control = $cb; Kind = 'select' }
            $script:y += $gap
            Add-FieldHelp -helpText $meta.Help -y ($script:y - 12) -targets @($lblField, $cb)
        }

        function Add-LaunchArgsEditor {
            param([hashtable]$Profile)

            _EnsureProfileManagerLoaded
            if (-not $script:_LaunchArgSectionExpanded) { $script:_LaunchArgSectionExpanded = @{} }
            $catalog = $null
            $defs = @()
            $groupedDefs = @()
            try {
                $catalog = Get-LaunchArgCatalog
                if ($catalog -and $catalog.Games) {
                    foreach ($gameName in @($catalog.Games.Keys)) {
                        $gameArgs = @($catalog.Games[$gameName].Args)
                        if ($gameArgs.Count -le 0) { continue }

                        $defs += $gameArgs
                        $groupedDefs += [pscustomobject]@{
                            GameName = $gameName
                            Defs     = $gameArgs
                        }
                    }
                }
            } catch {
                $catalog = $null
                $defs = @()
                $groupedDefs = @()
            }

            if (-not $defs) {
                $fallbackTitle = 'Launch Arguments (catalog not loaded)'
                $scroll.Controls.Add((_Label $fallbackTitle $fieldX $script:y 420 20 $fontBold))
                $script:y += 22
                Add-DynamicField -key 'LaunchArgs' -value $Profile.LaunchArgs
                return
            }

            # Always rebuild LaunchArgState from LaunchArgs to avoid stale CustomArgs
            try {
                $Profile.LaunchArgState = Build-LaunchArgState -GameName (_GetProfileKnownGame -Profile $Profile) -LaunchArgs $Profile.LaunchArgs
            } catch {
                $Profile.LaunchArgState = @{ Args = @{}; CustomArgs = $Profile.LaunchArgs }
            }
            $state = $Profile.LaunchArgState
            $profileKnownGame = _NormalizeGameIdentity (_GetProfileKnownGame -Profile $Profile)
            if (-not $state.Args) { $state.Args = @{} }

            $launchTitle = 'Launch Arguments by Game'
            $scroll.Controls.Add((_Label $launchTitle $fieldX $script:y 420 20 $fontBold))
            $script:y += 24
            $launchHint = _Label 'ECC builds launch arguments from game-aware toggles here, then saves the final command line back into the profile.' $fieldX $script:y 520 30 $fontLabel
            $launchHint.ForeColor = $clrTextSoft
            $launchHint.Anchor = 'Top,Left,Right'
            $scroll.Controls.Add($launchHint)
            $script:y += 34
            # Merge any unknown args found in current profiles so custom options
            # still appear as toggles, even if they are not yet in the catalog.
            $knownKeys = @{}
            foreach ($d in @($defs)) {
                if ($d -and $d.Key) { $knownKeys[$d.Key.ToLowerInvariant()] = $true }
            }

            $extraDefs = @()
            $unknownMap = @{}
            if ($script:SharedState -and $script:SharedState.Profiles) {
                foreach ($pfx in @($script:SharedState.Profiles.Keys)) {
                    $p = $script:SharedState.Profiles[$pfx]
                    if (-not $p -or -not $p.GameName) { continue }
                    $la = if ($p.LaunchArgs) { "$($p.LaunchArgs)" } else { '' }
                    if ([string]::IsNullOrWhiteSpace($la)) { continue }
                    try {
                        $st = Build-LaunchArgState -GameName $p.GameName -LaunchArgs $la
                        if ($st -and $st.Args) {
                            foreach ($k in $st.Args.Keys) {
                                if (-not $k) { continue }
                                if ($knownKeys.ContainsKey($k.ToLowerInvariant())) { continue }
                                $entry = $st.Args[$k]
                                $hasValue = ($entry -and $entry.ContainsKey('Value') -and $entry.Value -ne $null -and "$($entry.Value)" -ne '')
                                if (-not $unknownMap.ContainsKey($k)) {
                                    $unknownMap[$k] = @{ HasValue = $hasValue }
                                } elseif ($hasValue) {
                                    $unknownMap[$k].HasValue = $true
                                }
                            }
                        }
                    } catch { }
                }
            }

            foreach ($k in $unknownMap.Keys) {
                $hasValue = $unknownMap[$k].HasValue -eq $true
                $extraDefs += [pscustomobject]@{
                    Key   = $k
                    Label = "$k (custom)"
                    Type  = if ($hasValue) { 'value' } else { 'flag' }
                    Help  = 'Custom launch argument found in profiles.'
                }
            }

            if ($extraDefs.Count -gt 0) {
                $defs = @($defs) + @($extraDefs)
                $groupedDefs += [pscustomobject]@{
                    GameName = 'Custom / Discovered'
                    Defs     = @($extraDefs | Sort-Object Label, Key)
                }
            }

            # Ensure state has entries for any merged-in defs
            foreach ($d in @($defs)) {
                if (-not $d -or -not $d.Key) { continue }
                if (-not $state.Args.ContainsKey($d.Key)) {
                    $entry = @{ Enabled = $false }
                    if ($d.Default) { $entry.Value = "$($d.Default)" }
                    $state.Args[$d.Key] = $entry
                }
            }

            $count = @($defs).Count
            $groupCount = @($groupedDefs).Count
            $catalogText = "$count launch argument entries grouped across $groupCount game sections. Toggle anything that could apply to this server."
            $catalogLbl = _Label $catalogText $fieldX $script:y 420 18
            $catalogLbl.ForeColor = $clrTextSoft
            $scroll.Controls.Add($catalogLbl)
            $script:y += 22

            $controlMap = @{}
            $controlsByArgKey = @{}
            $syncingSharedArg = $false
            $tip = New-Object System.Windows.Forms.ToolTip
            $tip.InitialDelay = 300
            $tip.ReshowDelay  = 150
            $tip.AutoPopDelay = 10000
            $tip.ShowAlways   = $true

            function _GetLaunchArgDefField {
                param(
                    $Definition,
                    [string]$Name
                )

                if ($null -eq $Definition -or [string]::IsNullOrWhiteSpace($Name)) { return $null }
                if ($Definition -is [System.Collections.IDictionary]) {
                    if ($Definition.ContainsKey($Name)) {
                        return $Definition[$Name]
                    }
                    return $null
                }
                if ($Definition.PSObject -and $Definition.PSObject.Properties.Name -contains $Name) {
                    return $Definition.$Name
                }
                return $null
            }

            function _BuildLaunchArgStateFromControls {
                $tmp = @{ Args = @{}; CustomArgs = '' }
                foreach ($argKey in $controlsByArgKey.Keys) {
                    if ([string]::IsNullOrWhiteSpace($argKey)) { continue }

                    $enabled = $false
                    $bestValue = ''
                    foreach ($c in @($controlsByArgKey[$argKey])) {
                        if (-not $c) { continue }
                        if ($c.Check -and $c.Check.Checked -eq $true) {
                            $enabled = $true
                        }
                        if ($c.Text) {
                            $candidate = $c.Text.Text.Trim()
                            if (-not [string]::IsNullOrWhiteSpace($candidate)) {
                                $bestValue = $candidate
                            }
                        }
                    }

                    $tmp.Args[$argKey] = @{ Enabled = $enabled }
                    if (-not [string]::IsNullOrWhiteSpace($bestValue)) {
                        $tmp.Args[$argKey]['Value'] = $bestValue
                    }
                }
                $tmp.CustomArgs = if ($tbCustom) { $tbCustom.Text.Trim() } else { '' }
                return $tmp
            }

            function _BuildLaunchArgsFromEditorState {
                param([hashtable]$State)

                if (-not $State) { return '' }
                $parts = New-Object 'System.Collections.Generic.List[string]'
                $seenDefs = @{}

                foreach ($def in @($defs)) {
                    $defKeyRaw = _GetLaunchArgDefField -Definition $def -Name 'Key'
                    if ([string]::IsNullOrWhiteSpace("$defKeyRaw")) { continue }
                    $defKey = "$defKeyRaw"
                    $defKeyLower = $defKey.ToLowerInvariant()
                    if ($seenDefs.ContainsKey($defKeyLower)) { continue }
                    $seenDefs[$defKeyLower] = $true

                    if (-not $State.Args -or -not $State.Args.ContainsKey($defKey)) { continue }
                    $entry = $State.Args[$defKey]
                    if (-not $entry -or -not $entry.Enabled) { continue }

                    $defType = [string](_GetLaunchArgDefField -Definition $def -Name 'Type')
                    if ($defType -eq 'flag') {
                        $parts.Add($defKey) | Out-Null
                        continue
                    }
                    $defGroupKeys = _GetLaunchArgDefField -Definition $def -Name 'Keys'
                    if ($defType -eq 'flaggroup' -and $defGroupKeys) {
                        foreach ($k in @($defGroupKeys)) {
                            $parts.Add("$k") | Out-Null
                        }
                        continue
                    }

                    $val = if ($entry.ContainsKey('Value')) { "$($entry.Value)" } else { '' }
                    if ([string]::IsNullOrWhiteSpace($val)) { continue }

                    $quoteMode = _GetLaunchArgDefField -Definition $def -Name 'Quote'
                    $qMode = if ($quoteMode) { "$quoteMode" } else { 'auto' }
                    $vOut = $val
                    if ($qMode -eq 'always' -or ($qMode -eq 'auto' -and $vOut -match '\s')) {
                        $vOut = '"' + ($vOut -replace '"','\"') + '"'
                    }

                    $style = [string](_GetLaunchArgDefField -Definition $def -Name 'Style')
                    if ($style -eq 'equals') {
                        $parts.Add("$defKey=$vOut") | Out-Null
                    } else {
                        $parts.Add($defKey) | Out-Null
                        $parts.Add($vOut) | Out-Null
                    }
                }

                $custom = if ($State.CustomArgs) { "$($State.CustomArgs)".Trim() } else { '' }
                if ($custom) {
                    $parts.Add($custom) | Out-Null
                }

                return ($parts -join ' ').Trim()
            }

            # Live preview
            $previewLbl = _Label 'Generated LaunchArgs' $fieldX $script:y $lw 20 $fontBold
            $previewLbl.ForeColor = $clrTextSoft
            $scroll.Controls.Add($previewLbl)
            $tbPreview = _TextBox ($lw + 34) $script:y $tw $th '' $false
            $tbPreview.Anchor = 'Top,Left,Right'
            $tbPreview.ReadOnly = $true
            $tbPreview.BackColor = [System.Drawing.Color]::FromArgb(21,25,37)
            $scroll.Controls.Add($tbPreview)
            $script:y += $gap

            function _RebuildPreview {
                $tmp = _BuildLaunchArgStateFromControls
                $enabledKeys = @()
                if ($tmp -and $tmp.Args) {
                    foreach ($k in @($tmp.Args.Keys)) {
                        $entry = $tmp.Args[$k]
                        if ($entry -and $entry.Enabled -eq $true) {
                            $enabledKeys += "$k"
                        }
                    }
                }
                try {
                    $previewText = _BuildLaunchArgsFromEditorState -State $tmp
                    if (($enabledKeys.Count -gt 0) -and [string]::IsNullOrWhiteSpace($previewText)) {
                        $tbPreview.Text = "[empty preview with $($enabledKeys.Count) enabled args]"
                    } else {
                        $tbPreview.Text = $previewText
                    }
                    try {
                        $previewLen = if ($null -ne $previewText) { $previewText.Length } else { 0 }
                        Write-Host "[GUI] LaunchArgs preview built. Enabled=$($enabledKeys.Count) Length=$previewLen Value=$previewText" -ForegroundColor Cyan
                    } catch { }

                    if (($enabledKeys.Count -gt 0) -and [string]::IsNullOrWhiteSpace($previewText)) {
                        $sampleKeys = ($enabledKeys | Select-Object -First 8) -join ', '
                        $logLine = "LaunchArgs preview came back empty with $($enabledKeys.Count) enabled args. Sample: $sampleKeys"
                        try {
                            Write-Host "[GUI] $logLine" -ForegroundColor Yellow
                        } catch { }
                        try {
                            if ($script:SharedState -and $script:SharedState.ContainsKey('LogQueue')) {
                                $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][WARN][GUI] $logLine")
                            }
                        } catch { }
                        try { _QueueStatusMessage $logLine } catch { }
                    }
                } catch {
                    $tbPreview.Text = "[preview error] $($_.Exception.Message)"
                    try {
                        Write-Host "[GUI] LaunchArgs preview error: $($_.Exception.Message)" -ForegroundColor Yellow
                    } catch { }
                    try {
                        if ($script:SharedState -and $script:SharedState.ContainsKey('LogQueue')) {
                            $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][ERROR][GUI] LaunchArgs preview error: $($_.Exception.Message)")
                        }
                    } catch { }
                    try { _QueueStatusMessage "LaunchArgs preview error: $($_.Exception.Message)" } catch { }
                }
            }

            function _SyncSharedArgControls {
                param(
                    [string]$ArgKey,
                    [hashtable]$SourceControl
                )

                if ($syncingSharedArg) { return }
                if ([string]::IsNullOrWhiteSpace($ArgKey)) { return }
                if (-not $controlsByArgKey.ContainsKey($ArgKey)) { return }

                $syncingSharedArg = $true
                try {
                    foreach ($peer in @($controlsByArgKey[$ArgKey])) {
                        if ($peer -eq $SourceControl) { continue }
                        if ($peer.Check.Checked -ne $SourceControl.Check.Checked) {
                            $peer.Check.Checked = $SourceControl.Check.Checked
                        }

                        $sourceText = if ($SourceControl.Text) { $SourceControl.Text.Text } else { $null }
                        if ($peer.Text) {
                            $peerText = $peer.Text.Text
                            if ("$peerText" -ne "$sourceText") {
                                $peer.Text.Text = $sourceText
                            }
                        }
                    }
                } finally {
                    $syncingSharedArg = $false
                }
            }

            $rebuildPreviewAction = { _RebuildPreview }.GetNewClosure()
            $syncSharedArgControlsAction = {
                param([string]$ArgKey, [hashtable]$SourceControl)
                _SyncSharedArgControls -ArgKey $ArgKey -SourceControl $SourceControl
            }.GetNewClosure()

            try {
                $sectionPanels = New-Object 'System.Collections.Generic.List[object]'
                $panelWidth = [Math]::Max(500, $scroll.ClientSize.Width - $fieldX - 42)
                $panelLabelX = 12
                $panelInputX = ($lw + 34) - $fieldX
                $panelTextX  = ($lw + 124) - $fieldX
                $controlSequence = 0
                $customAnchorY = $script:y
                $customLbl = $null
                $tbCustom = $null

                function Add-LaunchArgControlRow {
                param(
                    $Parent,
                    $Definition,
                    [string]$SectionName,
                    [int]$RowY,
                    [int]$ContainerWidth
                )

                if (-not $Definition -or -not $Definition.Key) { return $null }

                $script:controlSequence++
                $controlId = '{0}::{1}::{2}' -f $SectionName, $Definition.Key, $script:controlSequence
                $label = "$($Definition.Key)"
                if ($Definition.Label) {
                    $label = "$($Definition.Label) [$($Definition.Key)]"
                }

                $lbl = _Label $label $panelLabelX $RowY ($panelInputX - $panelLabelX - 12) 20 $fontBold
                $lbl.ForeColor = $clrTextSoft
                $Parent.Controls.Add($lbl)

                $chk           = New-Object System.Windows.Forms.CheckBox
                $chk.Location  = [System.Drawing.Point]::new($panelInputX, $RowY)
                $chk.Size      = [System.Drawing.Size]::new(80, 20)
                $chk.Text      = 'On'
                $chk.ForeColor = $clrText
                $chk.BackColor = [System.Drawing.Color]::Transparent

                $entry = $null
                if ($state.Args.ContainsKey($Definition.Key)) {
                    $entry = $state.Args[$Definition.Key]
                }
                $chk.Checked = ($entry -and $entry.Enabled -eq $true)
                $Parent.Controls.Add($chk)

                $tb = $null
                $btnBrowse = $null
                if ($Definition.Type -ne 'flag' -and $Definition.Type -ne 'flaggroup') {
                    $btnW = 28
                    $isPath = ($Definition.Path -eq $true)
                    if ($isPath) {
                        $textW = [Math]::Max(120, $ContainerWidth - $panelTextX - ($btnW + 18))
                    } else {
                        $textW = [Math]::Max(180, $ContainerWidth - $panelTextX - 12)
                    }
                    $tb = _TextBox $panelTextX $RowY $textW $th '' $false
                    $tb.Anchor = 'Top,Left,Right'
                    if ($entry -and $entry.Value) { $tb.Text = "$($entry.Value)" }
                    $Parent.Controls.Add($tb)

                    if ($isPath) {
                        $btnBrowse = _Button '...' ($panelTextX + $textW + 6) ($RowY - 1) $btnW 24 $clrMuted {
                            $targetTb = $this.Tag
                            if (-not $targetTb) { return }

                            $resp = [System.Windows.Forms.MessageBox]::Show(
                                'Select a file path? (Yes = File, No = Folder)',
                                'Select Path','YesNoCancel','Question')
                            if ($resp -eq [System.Windows.Forms.DialogResult]::Cancel) { return }

                            if ($resp -eq [System.Windows.Forms.DialogResult]::Yes) {
                                $dlg = New-Object System.Windows.Forms.OpenFileDialog
                                $dlg.CheckFileExists = $false
                                $dlg.ValidateNames   = $true
                                $dlg.FileName        = 'Select'
                                $current = $targetTb.Text.Trim()
                                if ($current -and (Test-Path (Split-Path $current -Parent))) {
                                    $dlg.InitialDirectory = (Split-Path $current -Parent)
                                }
                                if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                                    $targetTb.Text = $dlg.FileName
                                }
                            } else {
                                $current = $targetTb.Text.Trim()
                                $owner = $this.FindForm()
                                $pickedFolder = _ShowEccFolderPicker -Title 'Select Folder Path' -Prompt 'Choose the folder path for this launch argument.' -InitialPath $current -Owner $owner
                                if (-not [string]::IsNullOrWhiteSpace($pickedFolder)) {
                                    $targetTb.Text = $pickedFolder
                                }
                            }
                        }
                        $btnBrowse.Tag = $tb
                        $btnBrowse.Anchor = 'Top,Right'
                        $Parent.Controls.Add($btnBrowse)
                    }
                }

                $control = @{ Def = $Definition; Check = $chk; Text = $tb; Browse = $btnBrowse }
                $controlMap[$controlId] = $control

                if (-not $controlsByArgKey.ContainsKey($Definition.Key)) {
                    $controlsByArgKey[$Definition.Key] = New-Object 'System.Collections.Generic.List[object]'
                }
                $controlsByArgKey[$Definition.Key].Add($control) | Out-Null

                $tipText = ''
                if ($SectionName -eq 'Master List') {
                    if ($Definition.Help) {
                        $tipText = "$($Definition.Help)"
                    }
                } elseif ($SectionName -eq 'Custom / Discovered') {
                    if ($Definition.Help) {
                        $tipText = "$($Definition.Help)"
                    } else {
                        $tipText = 'Custom launch argument found in profiles.'
                    }
                } else {
                    $baseHelp = ''
                    if ($Definition.Help) {
                        $baseHelp = "$($Definition.Help)`r`n"
                    }
                    $tipText = $baseHelp + "Could be used by $SectionName."
                }
                if ($tipText) {
                    $tip.SetToolTip($lbl, $tipText)
                    $tip.SetToolTip($chk, $tipText)
                    if ($tb) { $tip.SetToolTip($tb, $tipText) }
                    if ($btnBrowse) { $tip.SetToolTip($btnBrowse, $tipText) }
                }

                $chk.Tag = $control
                $chk.Add_CheckedChanged({
                    $bundle = $this.Tag
                    if ($bundle -and $bundle.Def) {
                        if ($syncSharedArgControlsAction -is [scriptblock]) {
                            $null = $syncSharedArgControlsAction.Invoke($bundle.Def.Key, $bundle)
                        }
                    }
                    if ($rebuildPreviewAction -is [scriptblock]) {
                        $null = $rebuildPreviewAction.Invoke()
                    }
                })
                if ($tb) {
                    $tb.Tag = $control
                    $tb.Add_TextChanged({
                        $bundle = $this.Tag
                        if ($bundle -and $bundle.Def) {
                            if ($syncSharedArgControlsAction -is [scriptblock]) {
                                $null = $syncSharedArgControlsAction.Invoke($bundle.Def.Key, $bundle)
                            }
                        }
                        if ($rebuildPreviewAction -is [scriptblock]) {
                            $null = $rebuildPreviewAction.Invoke()
                        }
                    })
                }

                return $control
            }

                function New-LaunchArgSection {
                param(
                    [string]$Title,
                    [string]$Hint,
                    $SectionDefs,
                    [bool]$Collapsible = $false,
                    [bool]$Expanded = $true,
                    $TitleColor = $clrText
                )

                $headerHeight = 32
                $hintHeight = 18
                if ([string]::IsNullOrWhiteSpace($Hint)) {
                    $hintHeight = 0
                }
                $bodyStartY = 26
                if ($hintHeight -gt 0) {
                    $bodyStartY = 42
                }
                $bodyHeight = [Math]::Max(0, (@($SectionDefs).Count * $gap) + 8)
                $expandedHeight = $bodyStartY + $bodyHeight
                $initialPanelHeight = $headerHeight
                if ($Expanded) {
                    $initialPanelHeight = $expandedHeight
                }
                $initialMarker = '+'
                if ($Expanded) {
                    $initialMarker = '-'
                }

                $panel = _Panel $fieldX $script:y $panelWidth $initialPanelHeight $clrPanelSoft
                $panel.Anchor = 'Top,Left,Right'
                $scroll.Controls.Add($panel)

                $titleText = $Title
                if ($Collapsible) {
                    $titleText = ('[{0}] {1}' -f $initialMarker, $Title)
                }
                $titleButton = _Button $titleText 8 4 ($panelWidth - 16) 24 $clrPanel {
                    $meta = $this.Tag
                    if (-not $meta -or -not $meta.Collapsible) { return }
                    $meta.Expanded = -not $meta.Expanded

                    $stateKey = "$($script:_EditingPrefix)|$($meta.Title)"
                    $script:_LaunchArgSectionExpanded[$stateKey] = ($meta.Expanded -eq $true)

                    $savedScroll = _CaptureScrollPosition $scroll

                    if ($script:_EditingProfile) {
                        _BuildProfileEditor -Profile $script:_EditingProfile
                        try {
                            if ($script:_ProfileEditorPanel -and $script:_ProfileEditorPanel.Controls.Count -gt 0) {
                                $newScroll = ($script:_ProfileEditorPanel.Controls | Select-Object -First 1)
                                _RestoreScrollPosition -Control $newScroll -Position $savedScroll
                            }
                        } catch { }
                    }
                    return
                }
                $titleButton.ForeColor = $TitleColor
                $panel.Controls.Add($titleButton)

                if ($hintHeight -gt 0) {
                    $hintLabel = _Label $Hint 12 28 ($panelWidth - 24) 16
                    $hintLabel.ForeColor = $clrTextSoft
                    $panel.Controls.Add($hintLabel)
                }

                $bodyPanel = _Panel 0 $bodyStartY ($panelWidth - 2) $bodyHeight $clrPanelSoft
                $bodyPanel.BorderStyle = 'None'
                $bodyPanel.Anchor = 'Top,Left,Right'
                $bodyPanel.Visible = $Expanded
                $panel.Controls.Add($bodyPanel)

                $rowY = 0
                foreach ($def in @($SectionDefs)) {
                    $null = Add-LaunchArgControlRow -Parent $bodyPanel -Definition $def -SectionName $Title -RowY $rowY -ContainerWidth ($panelWidth - 4)
                    $rowY += $gap
                }

                $meta = [pscustomobject]@{
                    Title          = $Title
                    Panel          = $panel
                    Body           = $bodyPanel
                    Collapsible    = $Collapsible
                    Expanded       = $Expanded
                    HeaderHeight   = $headerHeight
                    ExpandedHeight = $expandedHeight
                }
                $titleButton.Tag = $meta
                $sectionPanels.Add($meta) | Out-Null
                $script:y += $panel.Height + 8
                return $meta
            }

                foreach ($group in @($groupedDefs)) {
                    if (-not $group -or -not $group.Defs) { continue }
                    if ((_NormalizeGameIdentity $group.GameName) -eq $profileKnownGame) {
                        $groupHintText = 'All possible arguments this server profile is most likely to use.'
                    } elseif ($group.GameName -eq 'Custom / Discovered') {
                        $groupHintText = 'Extra arguments found in saved profiles but not in the built-in catalog.'
                    } else {
                        $groupHintText = "All possible arguments that could be used by $($group.GameName)."
                    }
                    $groupColor = $clrText
                    if ((_NormalizeGameIdentity $group.GameName) -eq $profileKnownGame) {
                        $groupColor = $clrAccentAlt
                    }
                    $stateKey = "$($script:_EditingPrefix)|$($group.GameName)"
                    $groupExpanded = ((_NormalizeGameIdentity $group.GameName) -eq $profileKnownGame)
                    if ($script:_LaunchArgSectionExpanded.ContainsKey($stateKey)) {
                        $groupExpanded = ($script:_LaunchArgSectionExpanded[$stateKey] -eq $true)
                    }
                    $groupCollapsible = $true
                    $null = New-LaunchArgSection -Title $group.GameName -Hint $groupHintText -SectionDefs @($group.Defs) -Collapsible $groupCollapsible -Expanded $groupExpanded -TitleColor $groupColor
                }

                $masterExpanded = $false
                $masterCollapsible = $true
                $null = New-LaunchArgSection -Title 'Master List' -Hint 'Everything in the launch-argument catalog, shown in one continuous list.' -SectionDefs @($defs) -Collapsible $masterCollapsible -Expanded $masterExpanded -TitleColor $clrAccentAlt

                $customAnchorY = $script:y

                # Custom args (not in catalog)
                $customLbl = _Label 'Custom Args (unsupported/advanced)' $fieldX $customAnchorY $lw 20 $fontBold
                $customLbl.ForeColor = $clrTextSoft
                $scroll.Controls.Add($customLbl)
                $tbCustom = _TextBox ($lw + 38) $customAnchorY $tw $th '' $false
                $tbCustom.Anchor = 'Top,Left,Right'
                if ($state.CustomArgs) { $tbCustom.Text = "$($state.CustomArgs)" }
                $scroll.Controls.Add($tbCustom)
                $script:y = $customAnchorY + $gap
                $tbCustom.Add_TextChanged({
                    if ($rebuildPreviewAction -is [scriptblock]) {
                        $null = $rebuildPreviewAction.Invoke()
                    }
                })

                $launchArgRelayoutTrace = New-Object 'System.Collections.Generic.List[string]'
                $launchArgRelayoutTrace.Add("Preparing relayout") | Out-Null
                $sectionsSnapshot = $null
                try {
                    $sectionPanelsType = if ($null -eq $sectionPanels) { '<null>' } else { $sectionPanels.GetType().FullName }
                    $sectionPanelsCountHint = 0
                    try { $sectionPanelsCountHint = @($sectionPanels).Count } catch { $sectionPanelsCountHint = -1 }
                    $launchArgRelayoutTrace.Add("SectionPanels type: $sectionPanelsType") | Out-Null
                    $launchArgRelayoutTrace.Add("SectionPanels count hint: $sectionPanelsCountHint") | Out-Null
                    $sectionsSnapshot = @($sectionPanels)
                    $launchArgRelayoutTrace.Add("Section snapshot count: $(@($sectionsSnapshot).Count)") | Out-Null
                } catch {
                    throw [System.ArgumentException]::new(
                        "Launch-arg section relayout failed while materializing section list. SectionPanelsType='$sectionPanelsType'; CountHint='$sectionPanelsCountHint'. Inner: $($_.Exception.Message)"
                    )
                }

                $currentY = $customAnchorY
                foreach ($section in $sectionsSnapshot) {
                    $sectionPanel = $null
                    $sectionTitle = '<unknown>'
                    $launchArgRelayoutTrace.Add("Loop start: currentY=$currentY") | Out-Null
                    if ($section -and $section.PSObject -and $section.PSObject.Properties.Name -contains 'Title') {
                        try { $sectionTitle = [string]$section.Title } catch { $sectionTitle = '<unknown>' }
                    }
                    $launchArgRelayoutTrace.Add("Section title: $sectionTitle") | Out-Null
                    if ($section -and $section.PSObject -and $section.PSObject.Properties.Name -contains 'Panel') {
                        $sectionPanel = $section.Panel
                    }
                    $launchArgRelayoutTrace.Add("Panel resolved: $(if ($sectionPanel) { $sectionPanel.GetType().FullName } else { '<null>' })") | Out-Null
                    if (-not ($sectionPanel -is [System.Windows.Forms.Control])) { continue }
                    $launchArgRelayoutTrace.Add("Before Location: section=$sectionTitle; fieldX=$fieldX; currentY=$currentY") | Out-Null
                    try {
                        $sectionPanel.Location = [System.Drawing.Point]::new($fieldX, $currentY)
                    } catch {
                        throw [System.ArgumentException]::new(
                            "Launch-arg section relayout failed during Location update. Section='$sectionTitle'; PanelType='$($sectionPanel.GetType().FullName)'; fieldX=$fieldX; currentY=$currentY; panelWidth=$panelWidth; panelHeight=$($sectionPanel.Height). Inner: $($_.Exception.Message)"
                        )
                    }
                    $launchArgRelayoutTrace.Add("After Location: section=$sectionTitle; location=$($sectionPanel.Location)") | Out-Null
                    $launchArgRelayoutTrace.Add("Before Width: section=$sectionTitle; panelWidth=$panelWidth") | Out-Null
                    try {
                        $sectionPanel.Width = $panelWidth
                    } catch {
                        throw [System.ArgumentException]::new(
                            "Launch-arg section relayout failed during Width update. Section='$sectionTitle'; PanelType='$($sectionPanel.GetType().FullName)'; fieldX=$fieldX; currentY=$currentY; panelWidth=$panelWidth; panelHeight=$($sectionPanel.Height). Inner: $($_.Exception.Message)"
                        )
                    }
                    $launchArgRelayoutTrace.Add("After Width: section=$sectionTitle; width=$($sectionPanel.Width)") | Out-Null
                    $launchArgRelayoutTrace.Add("Before Height read: section=$sectionTitle") | Out-Null
                    try {
                        $sectionPanelHeight = [int]$sectionPanel.Height
                    } catch {
                        throw [System.ArgumentException]::new(
                            "Launch-arg section relayout failed while reading Height. Section='$sectionTitle'; PanelType='$($sectionPanel.GetType().FullName)'; fieldX=$fieldX; currentY=$currentY; panelWidth=$panelWidth. Inner: $($_.Exception.Message)"
                        )
                    }
                    $launchArgRelayoutTrace.Add("After Height read: section=$sectionTitle; height=$sectionPanelHeight") | Out-Null
                    $currentY += $sectionPanelHeight + 8
                    $launchArgRelayoutTrace.Add("Loop end: section=$sectionTitle; nextY=$currentY") | Out-Null
                }
                if ($customLbl) {
                    $customLbl.Location = [System.Drawing.Point]::new($fieldX, $currentY)
                }
                if ($tbCustom) {
                    $tbCustom.Location = [System.Drawing.Point]::new(($lw + 38), $currentY)
                }

                # Initial preview
                if ($rebuildPreviewAction -is [scriptblock]) {
                    $null = $rebuildPreviewAction.Invoke()
                }
            } catch {
                $fallbackMessage = "Launch args editor fallback active: $($_.Exception.Message)"
                $isKnownFallback = ($_.Exception.Message -eq 'Argument types do not match')
                try {
                    Write-Host "[GUI] $fallbackMessage" -ForegroundColor Yellow
                } catch { }
                if (-not $isKnownFallback) {
                    try {
                        if ($script:SharedState -and $script:SharedState.ContainsKey('LogQueue')) {
                            $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][WARN][GUI] $fallbackMessage")
                        }
                    } catch { }
                    try { _QueueStatusMessage $fallbackMessage } catch { }
                } else {
                    try {
                        if ($script:SharedState -and $script:SharedState.ContainsKey('LogQueue') -and $script:SharedState.Settings -and [bool]$script:SharedState.Settings.EnableDebugLogging) {
                            $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][DEBUG][GUI] $fallbackMessage")
                        }
                    } catch { }
                }

                $scroll.Controls.Add((_Label 'Launch Arguments (fallback editor)' $fieldX $script:y 420 20 $fontBold))
                $script:y += 22
                Add-DynamicField -key 'LaunchArgs' -value $Profile.LaunchArgs

                $fallbackDetails = New-Object 'System.Collections.Generic.List[string]'
                if ($_.Exception -and -not [string]::IsNullOrWhiteSpace([string]$_.Exception.Message)) {
                    $fallbackDetails.Add("Reason: $($_.Exception.Message)") | Out-Null
                }
                if ($_.Exception -and -not [string]::IsNullOrWhiteSpace([string]$_.Exception.GetType().FullName)) {
                    $fallbackDetails.Add("Type: $($_.Exception.GetType().FullName)") | Out-Null
                }
                if ($_.InvocationInfo) {
                    if ($_.InvocationInfo.ScriptLineNumber) {
                        $fallbackDetails.Add("Line: $($_.InvocationInfo.ScriptLineNumber)") | Out-Null
                    }
                    $commandText = ''
                    try { $commandText = [string]$_.InvocationInfo.Line } catch { $commandText = '' }
                    if ([string]::IsNullOrWhiteSpace($commandText)) {
                        try { $commandText = [string]$_.InvocationInfo.PositionMessage } catch { $commandText = '' }
                    }
                    if (-not [string]::IsNullOrWhiteSpace($commandText)) {
                        $commandText = $commandText.Trim()
                        if ($commandText.Length -gt 220) {
                            $commandText = $commandText.Substring(0, 220) + '...'
                        }
                        $fallbackDetails.Add("Command: $commandText") | Out-Null
                    }
                }
                $fallbackDetailText = if ($fallbackDetails.Count -gt 0) { $fallbackDetails -join "`r`n" } else { '' }
                if ($launchArgRelayoutTrace -and $launchArgRelayoutTrace.Count -gt 0) {
                    $traceText = $launchArgRelayoutTrace -join "`r`n"
                    if (-not [string]::IsNullOrWhiteSpace($traceText)) {
                        if (-not [string]::IsNullOrWhiteSpace($fallbackDetailText)) {
                            $fallbackDetailText += "`r`n`r`nRelayout trace:`r`n$traceText"
                        } else {
                            $fallbackDetailText = "Relayout trace:`r`n$traceText"
                        }
                    }
                }
                $fallbackHintText = 'Generated LaunchArgs is currently mirroring the raw LaunchArgs field because the grouped editor is not available for this profile.'
                $fallbackHintWidth = [Math]::Max(320, $scroll.ClientSize.Width - ($fieldX + 24))
                $fallbackHintHeight = _MeasureProfileHelpHeight -text $fallbackHintText -width $fallbackHintWidth -font $fontLabel
                $fallbackHint = _Label $fallbackHintText $fieldX $script:y $fallbackHintWidth $fallbackHintHeight
                $fallbackHint.ForeColor = $clrYellow
                $fallbackHint.Anchor = 'Top,Left,Right'
                $fallbackHint.AutoSize = $false
                $scroll.Controls.Add($fallbackHint)
                $script:y += ($fallbackHintHeight + 4)

                if (-not [string]::IsNullOrWhiteSpace($fallbackDetailText)) {
                    $debugPanelHeight = 132
                    $debugPanel = _Panel $fieldX $script:y ([Math]::Max(320, $scroll.ClientSize.Width - ($fieldX + 24))) $debugPanelHeight $clrPanelSoft
                    $debugPanel.Anchor = 'Top,Left,Right'
                    $scroll.Controls.Add($debugPanel)

                    $debugTitle = _Label 'Grouped Editor Diagnostic' 10 8 ($debugPanel.Width - 20) 18 $fontBold
                    $debugTitle.ForeColor = $clrYellow
                    $debugTitle.Anchor = 'Top,Left,Right'
                    $debugPanel.Controls.Add($debugTitle)

                    $tbFallbackDebug = New-Object System.Windows.Forms.TextBox
                    $tbFallbackDebug.Location = [System.Drawing.Point]::new(10, 30)
                    $tbFallbackDebug.Size = [System.Drawing.Size]::new(($debugPanel.Width - 20), ($debugPanelHeight - 40))
                    $tbFallbackDebug.Anchor = 'Top,Left,Right'
                    $tbFallbackDebug.Multiline = $true
                    $tbFallbackDebug.ScrollBars = 'Vertical'
                    $tbFallbackDebug.WordWrap = $true
                    $tbFallbackDebug.ReadOnly = $true
                    $tbFallbackDebug.BackColor = [System.Drawing.Color]::FromArgb(21,25,37)
                    $tbFallbackDebug.ForeColor = $clrYellow
                    $tbFallbackDebug.BorderStyle = 'FixedSingle'
                    $tbFallbackDebug.Font = _ResolveUiFont -Font $fontLabel
                    $tbFallbackDebug.Text = $fallbackDetailText
                    $debugPanel.Controls.Add($tbFallbackDebug)

                    $script:y += ($debugPanelHeight + 8)
                }

                if ($script:_ProfileFields.ContainsKey('LaunchArgs')) {
                    $fallbackEntry = $script:_ProfileFields['LaunchArgs']
                    if ($fallbackEntry -and $fallbackEntry.Control) {
                        $fallbackControl = $fallbackEntry.Control
                        try {
                            $tbPreview.Text = $fallbackControl.Text
                        } catch { }
                        $fallbackControl.Add_TextChanged({
                            try {
                                $tbPreview.Text = $this.Text
                            } catch { }
                        })
                    }
                }
                return
            }

            $script:_ProfileFields['LaunchArgs'] = @{
                Kind         = 'launchargs'
                Defs         = $defs
                Controls     = $controlMap
                CustomBox    = $tbCustom
                BuildState   = ${function:_BuildLaunchArgStateFromControls}
                BuildPreview = ${function:_BuildLaunchArgsFromEditorState}
            }
        }

        $script:_ProfileFields = @{}
        if (-not $script:_ProfileFieldsByPrefix) { $script:_ProfileFieldsByPrefix = @{} }
        $script:_ProfileSeparators = @()
        $script:y = 14

        $profileHeader = _Panel $contentX $script:y ($scroll.ClientSize.Width - 52) 98 $clrPanelSoft
        $profileHeader.Anchor = 'Top,Left,Right'
        $profileHeader.BorderStyle = 'FixedSingle'
        $scroll.Controls.Add($profileHeader)

        $profileHeaderAccent = _Panel 0 0 4 98 $clrAccent
        $profileHeaderAccent.BorderStyle = 'None'
        $profileHeader.Controls.Add($profileHeaderAccent)

        $headerTitle = _Label "$($Profile.GameName)" 18 14 340 24 $fontTitle
        $headerTitle.ForeColor = $clrText
        $profileHeader.Controls.Add($headerTitle)
        $headerMeta = _Label "Prefix [$($script:_EditingPrefix)]   |   Process $($Profile.ProcessName)" 18 44 ($profileHeader.Width - 36) 18 $fontBold
        $headerMeta.ForeColor = $clrTextSoft
        $headerMeta.Anchor = 'Top,Left,Right'
        $profileHeader.Controls.Add($headerMeta)
        $headerHint = _Label 'Edit profile behavior, launch arguments, log paths, restart settings, and integrations.' 18 66 ($profileHeader.Width - 36) 18 $fontLabel
        $headerHint.ForeColor = $clrTextSoft
        $headerHint.Anchor = 'Top,Left,Right'
        $profileHeader.Controls.Add($headerHint)
        $script:y += 116

        function Add-SectionHeader([string]$title) {
            $lblSec = _Label $title $fieldX $script:y 320 24 $fontTitle
            $lblSec.ForeColor = $clrAccentAlt
            $scroll.Controls.Add($lblSec)
            $script:y += 28
            $sep = New-Object System.Windows.Forms.Panel
            $sep.Location = [System.Drawing.Point]::new($fieldX, $script:y)
            $sep.Size     = [System.Drawing.Size]::new([Math]::Max(100, $scroll.ClientSize.Width - 56), 2)
            $sep.BackColor = $clrBorder
            $sep.Anchor   = 'Top,Left,Right'
            $scroll.Controls.Add($sep)
            $script:_ProfileSeparators += $sep
            $script:y += 18
        }

        $sections = [ordered]@{
            'Basics' = @('GameName','Prefix','ProcessName','Executable','FolderPath')
            'Launch' = @('LaunchArgs','MaxRamGB','MinRamGB','AOTCache')
            'Logs'   = @('LogStrategy','ServerLogRoot','ServerLogSubDir','ServerLogFile','ServerLogPath','ServerLogNote')
            'REST'   = @('RestEnabled','RestHost','RestPort','RestPassword','RestProtocol','RestPollOnlyWhenRunning')
            'RCON'   = @('RconHost','RconPort','RconPassword')
            'Restart/Safety' = @('BlockStartIfRamPercentUsed','BlockStartIfFreeRamBelowGB','StartupTimeoutSeconds','ShutdownIfNoPlayersAfterStartupMinutes','ShutdownIfEmptyAfterLastPlayerLeavesMinutes','EnableAutoRestart','RestartDelaySeconds','MaxRestartsPerHour','SaveMethod','SaveWaitSeconds','StopMethod')
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

            if ($sec -eq 'Launch') {
                if ($Profile.Keys -contains 'LaunchArgs') {
                    Add-LaunchArgsEditor -Profile $Profile
                    $used['LaunchArgs'] = $true
                }
                if ($Profile.Keys -contains 'LaunchArgState') { $used['LaunchArgState'] = $true }
                foreach ($k in @('MaxRamGB','MinRamGB','AOTCache')) {
                    if ($Profile.Keys -contains $k) {
                        Add-DynamicField -key $k -value $Profile[$k]
                        $used[$k] = $true
                    }
                }
                continue
            }

            foreach ($k in $keys) {
                if (-not ($Profile.Keys -contains $k)) { continue }

                if ($k -eq 'SaveMethod') {
                    Add-SelectField -key $k -value $Profile[$k] -options @('none','stdin','http','rest','SatisfactoryApi')
                } elseif ($k -eq 'StopMethod') {
                    Add-SelectField -key $k -value $Profile[$k] -options @('processKill','processName','stdin','ctrlc','http','SatisfactoryApi')
                } elseif ($k -eq 'LogStrategy') {
                    Add-SelectField -key $k -value $Profile[$k] -options @('SingleFile','NewestFile','PZSessionFolder','ValheimUserFolder')
                } else {
                    Add-DynamicField -key $k -value $Profile[$k]
                }
                $used[$k] = $true
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

        # Capture the field map for this prefix so it doesn't get clobbered
        if ($script:_EditingPrefix) {
            $script:_ProfileFieldsByPrefix[$script:_EditingPrefix] = $script:_ProfileFields
        }

        $actionTop = $script:y + 16
        $actionCard = _Panel $contentX $actionTop ($scroll.ClientSize.Width - 52) 114 $clrPanelSoft
        $actionCard.Anchor = 'Top,Left,Right'
        $actionCard.BorderStyle = 'FixedSingle'
        $scroll.Controls.Add($actionCard)
        $actionTitle = _Label 'Profile Actions' 16 12 220 18 $fontBold
        $actionTitle.ForeColor = $clrAccentAlt
        $actionCard.Controls.Add($actionTitle)
        $actionHint = _Label 'Save profile changes or control the selected server directly from here.' 16 34 ($actionCard.Width - 32) 18 $fontLabel
        $actionHint.ForeColor = $clrTextSoft
        $actionHint.Anchor = 'Top,Left,Right'
        $actionCard.Controls.Add($actionHint)

        $actionButtonsRow = New-Object System.Windows.Forms.FlowLayoutPanel
        $actionButtonsRow.Location = [System.Drawing.Point]::new(16, 66)
        $actionButtonsRow.Size = [System.Drawing.Size]::new([Math]::Max(120, $actionCard.ClientSize.Width - 32), 40)
        $actionButtonsRow.Anchor = 'Top,Left,Right'
        $actionButtonsRow.WrapContents = $false
        $actionButtonsRow.AutoScroll = $true
        $actionButtonsRow.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
        $actionButtonsRow.BackColor = [System.Drawing.Color]::Transparent
        $actionCard.Controls.Add($actionButtonsRow)

        $btnSave = _Button 'Save Changes' 14 54 146 34 $clrGreen {
            # Use script-scope reference - the param $Profile is gone by click time
            $prof = $script:_EditingProfile
            if ($null -eq $prof -or -not ($prof -is [System.Collections.IDictionary])) {
                [System.Windows.Forms.MessageBox]::Show('No profile loaded for editing.','Error','OK','Error') | Out-Null
                return
            }

            $before = @{}
            foreach ($k in $prof.Keys) { $before[$k] = $prof[$k] }

            $fieldMap = $null
            if ($script:_ProfileFieldsByPrefix -and $script:_EditingPrefix -and
                $script:_ProfileFieldsByPrefix.ContainsKey($script:_EditingPrefix)) {
                $fieldMap = $script:_ProfileFieldsByPrefix[$script:_EditingPrefix]
            } else {
                $fieldMap = $script:_ProfileFields
            }

            foreach ($key in $fieldMap.Keys) {
                $entry = $fieldMap[$key]
                $ctrl  = $entry.Control
                $kind  = $entry.Kind

                if ($kind -eq 'bool') {
                    $prof[$key] = $ctrl.Checked
                    continue
                }

                if ($kind -eq 'select') {
                    if ($null -ne $ctrl.SelectedValue -and -not [string]::IsNullOrWhiteSpace([string]$ctrl.SelectedValue)) {
                        $prof[$key] = [string]$ctrl.SelectedValue
                    } elseif ($ctrl.SelectedItem -and $ctrl.SelectedItem.PSObject.Properties['Value']) {
                        $prof[$key] = [string]$ctrl.SelectedItem.Value
                    } elseif ($ctrl.SelectedItem) {
                        $prof[$key] = "$($ctrl.SelectedItem)"
                    } else {
                        $prof[$key] = "$($ctrl.Text)"
                    }
                    continue
                }

                if ($kind -eq 'launchargs') {
                    _EnsureProfileManagerLoaded
                    $defs = $entry.Defs
                    $controls = $entry.Controls
                    $state = $null
                    if ($entry.ContainsKey('BuildState') -and $entry.BuildState) {
                        $state = & $entry.BuildState
                    } else {
                        $state = @{ Args = @{}; CustomArgs = '' }
                        $seenKeys = @{}
                        foreach ($controlId in $controls.Keys) {
                            $c = $controls[$controlId]
                            if (-not $c -or -not $c.Def -or -not $c.Def.Key) { continue }
                            $argKey = "$($c.Def.Key)"
                            $argKeyLower = $argKey.ToLowerInvariant()
                            if ($seenKeys.ContainsKey($argKeyLower)) { continue }
                            $seenKeys[$argKeyLower] = $true
                            $enabled = $c.Check.Checked -eq $true
                            $val = if ($c.Text) { $c.Text.Text.Trim() } else { '' }
                            $state.Args[$argKey] = @{ Enabled = $enabled }
                            if ($c.Text -and $val -ne '') { $state.Args[$argKey]['Value'] = $val }
                        }
                        $state.CustomArgs = if ($entry.CustomBox) { $entry.CustomBox.Text.Trim() } else { '' }
                    }
                    $prof['LaunchArgState'] = $state
                    if ($entry.ContainsKey('BuildPreview') -and $entry.BuildPreview) {
                        $prof['LaunchArgs'] = (& $entry.BuildPreview $state)
                    } else {
                        $prof['LaunchArgs'] = Build-LaunchArgsFromState -GameName $prof.GameName -State $state
                    }
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
        $btnSave.Margin = [System.Windows.Forms.Padding]::new(0, 0, 10, 0)
        $actionButtonsRow.Controls.Add($btnSave)
        _SetMainControlToolTip -Control $btnSave -Text 'Save changes to the selected server profile.'

        $btnRestart = _Button 'Restart Server' 0 0 146 34 $clrAccent {
            # Use script-scope prefix - $Profile is gone by click time
            $pfxNow = $script:_EditingPrefix
            if (-not $pfxNow) {
                [System.Windows.Forms.MessageBox]::Show('No profile selected.','Error','OK','Error') | Out-Null
                return
            }
            _RunServerOpInBackground -Prefix $pfxNow -Operation 'Restart'
        }
        _SetMainControlToolTip -Control $btnRestart
        $btnRestart.Margin = [System.Windows.Forms.Padding]::new(0, 0, 10, 0)
        $actionButtonsRow.Controls.Add($btnRestart)

        $btnStop = _Button 'Stop Server' 0 0 132 34 $clrRed {
            # Use script-scope prefix - $Profile is gone by click time
            $pfxNow = $script:_EditingPrefix
            if (-not $pfxNow) {
                [System.Windows.Forms.MessageBox]::Show('No profile selected.','Error','OK','Error') | Out-Null
                return
            }
            _RunServerOpInBackground -Prefix $pfxNow -Operation 'Stop'
        }
        _SetMainControlToolTip -Control $btnStop
        $btnStop.Margin = [System.Windows.Forms.Padding]::new(0)
        $actionButtonsRow.Controls.Add($btnStop)
    }

    # =====================================================================
    # FIRST-RUN / ADD-GAME HELPERS
    # =====================================================================
    function _GetSteamLibraryCommonRoots {
        $roots = New-Object 'System.Collections.Generic.List[string]'
        $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

        $addRoot = {
            param([string]$BasePath)

            if ([string]::IsNullOrWhiteSpace($BasePath)) { return }

            $expanded = ''
            try { $expanded = [Environment]::ExpandEnvironmentVariables($BasePath) } catch { $expanded = $BasePath }
            if ([string]::IsNullOrWhiteSpace($expanded) -or -not (Test-Path -LiteralPath $expanded)) { return }

            $normalized = ''
            try { $normalized = [System.IO.Path]::GetFullPath($expanded) } catch { $normalized = $expanded }
            if ([string]::IsNullOrWhiteSpace($normalized)) { return }

            $commonPath = ''
            $trimmed = $normalized.TrimEnd('\')
            if ($trimmed -match '(?i)\\steamapps\\common$') {
                $commonPath = $trimmed
            } else {
                $commonPath = Join-Path $trimmed 'steamapps\common'
            }

            if (-not [string]::IsNullOrWhiteSpace($commonPath) -and (Test-Path -LiteralPath $commonPath) -and $seen.Add($commonPath)) {
                $roots.Add($commonPath) | Out-Null
            }
        }

        $addSteamLibrariesFromRoot = {
            param([string]$SteamRoot)

            if ([string]::IsNullOrWhiteSpace($SteamRoot)) { return }
            & $addRoot $SteamRoot

            $libraryVdfPath = Join-Path $SteamRoot 'steamapps\libraryfolders.vdf'
            if (-not (Test-Path -LiteralPath $libraryVdfPath)) { return }

            try {
                $vdfLines = Get-Content -LiteralPath $libraryVdfPath -ErrorAction Stop
                foreach ($line in $vdfLines) {
                    if ($line -match '"path"\s+"([^"]+)"') {
                        $libPath = $Matches[1] -replace '\\\\', '\'
                        & $addRoot $libPath
                    }
                }
            } catch { }
        }

        foreach ($steamRoot in @(
            (try { (Get-ItemProperty -Path 'HKCU:\Software\Valve\Steam' -Name 'SteamPath' -ErrorAction Stop).SteamPath } catch { $null }),
            (try { (Get-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\Valve\Steam' -Name 'InstallPath' -ErrorAction Stop).InstallPath } catch { $null }),
            (try { Join-Path ${env:ProgramFiles(x86)} 'Steam' } catch { $null }),
            (try { Join-Path $env:ProgramFiles 'Steam' } catch { $null })
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique) {
            & $addSteamLibrariesFromRoot $steamRoot
        }

        try {
            foreach ($drive in @(Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue)) {
                if (-not $drive.Root) { continue }
                foreach ($basePath in @(
                    (Join-Path $drive.Root 'SteamLibrary'),
                    (Join-Path $drive.Root 'Program Files (x86)\Steam'),
                    (Join-Path $drive.Root 'Program Files\Steam')
                )) {
                    & $addSteamLibrariesFromRoot $basePath
                }
            }
        } catch { }

        return @($roots)
    }

    function _GetAutoDetectedServerFolders {
        param([string]$KnownGameName)

        if ([string]::IsNullOrWhiteSpace($KnownGameName)) { return @() }

        $pmPath = Join-Path $script:ModuleRoot 'ProfileManager.psm1'
        Import-Module $pmPath -Force | Out-Null

        $commonRoots = @(_GetSteamLibraryCommonRoots)
        if ($commonRoots.Count -eq 0) { return @() }

        $targetIdentity = _NormalizeGameIdentity $KnownGameName
        $results = New-Object 'System.Collections.Generic.List[object]'
        $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

        foreach ($commonRoot in $commonRoots) {
            foreach ($dir in @(Get-ChildItem -LiteralPath $commonRoot -Directory -ErrorAction SilentlyContinue)) {
                $folderPath = $dir.FullName
                if (-not $seen.Add($folderPath)) { continue }

                $resolvedGame = $null
                try {
                    $resolvedGame = Resolve-KnownGameFromFolder -FolderPath $folderPath -PreferredGameName $KnownGameName -DisplayName $dir.Name
                } catch {
                    $resolvedGame = $null
                }

                if ((_NormalizeGameIdentity "$resolvedGame") -ne $targetIdentity) { continue }

                $exeHint = ''
                try {
                    $exeHint = Find-ServerExecutable -FolderPath $folderPath -Hints @()
                } catch { $exeHint = '' }

                $score = 0
                if ((_NormalizeGameIdentity $dir.Name) -eq $targetIdentity) { $score += 120 }
                if (-not [string]::IsNullOrWhiteSpace($exeHint)) { $score += 80 }
                $score += 10

                $results.Add([pscustomobject]@{
                    Path       = $folderPath
                    Root       = $commonRoot
                    Display    = $dir.Name
                    Executable = $exeHint
                    Score      = $score
                }) | Out-Null
            }
        }

        return @($results | Sort-Object @{ Expression = 'Score'; Descending = $true }, Display, Path)
    }

    function _FolderPickerNodeHasChildren {
        param([string]$Path)

        if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) { return $false }
        try {
            return ($null -ne (Get-ChildItem -LiteralPath $Path -Directory -ErrorAction SilentlyContinue | Select-Object -First 1))
        } catch {
            return $false
        }
    }

    function _NewFolderPickerNode {
        param(
            [string]$Path,
            [string]$Label = ''
        )

        if ([string]::IsNullOrWhiteSpace($Path)) { return $null }

        $node = New-Object System.Windows.Forms.TreeNode
        $nodePath = $Path
        try { $nodePath = [System.IO.Path]::GetFullPath($Path) } catch { $nodePath = $Path }
        $nodePath = $nodePath.TrimEnd('\')
        if ([string]::IsNullOrWhiteSpace($nodePath) -and $Path -match '^[A-Za-z]:\\?$') {
            $nodePath = $Path.TrimEnd('\') + '\'
        }

        if ([string]::IsNullOrWhiteSpace($Label)) {
            $leaf = ''
            try { $leaf = [System.IO.Path]::GetFileName($nodePath.TrimEnd('\')) } catch { $leaf = '' }
            if ([string]::IsNullOrWhiteSpace($leaf)) {
                $Label = $nodePath
            } else {
                $Label = $leaf
            }
        }

        $node.Text = $Label
        $node.Tag = $nodePath

        if (_FolderPickerNodeHasChildren -Path $nodePath) {
            $placeholder = New-Object System.Windows.Forms.TreeNode
            $placeholder.Text = '...'
            $placeholder.Tag = '__placeholder__'
            [void]$node.Nodes.Add($placeholder)
        }

        return $node
    }

    function _LoadFolderPickerChildren {
        param([System.Windows.Forms.TreeNode]$Node)

        if ($null -eq $Node) { return }
        $basePath = [string]$Node.Tag
        if ([string]::IsNullOrWhiteSpace($basePath) -or $basePath -eq '__placeholder__') { return }

        $hasPlaceholder = ($Node.Nodes.Count -eq 1 -and [string]$Node.Nodes[0].Tag -eq '__placeholder__')
        if (-not $hasPlaceholder) { return }

        $Node.Nodes.Clear()
        try {
            $children = @(Get-ChildItem -LiteralPath $basePath -Directory -ErrorAction SilentlyContinue | Sort-Object Name)
            foreach ($child in $children) {
                $childNode = _NewFolderPickerNode -Path $child.FullName -Label $child.Name
                if ($childNode) { [void]$Node.Nodes.Add($childNode) }
            }
        } catch { }
    }

    function _ShowEccFolderPicker {
        param(
            [string]$Title = 'Select Folder',
            [string]$Prompt = 'Choose a folder.',
            [string]$InitialPath = '',
            [System.Windows.Forms.IWin32Window]$Owner = $null
        )

        $newFolderPickerNodeLocal = ${function:_NewFolderPickerNode}
        $loadFolderPickerChildrenLocal = {
            param([System.Windows.Forms.TreeNode]$Node)

            if ($null -eq $Node) { return }
            $basePath = [string]$Node.Tag
            if ([string]::IsNullOrWhiteSpace($basePath) -or $basePath -eq '__placeholder__') { return }

            $hasPlaceholder = ($Node.Nodes.Count -eq 1 -and [string]$Node.Nodes[0].Tag -eq '__placeholder__')
            if (-not $hasPlaceholder) { return }

            $Node.Nodes.Clear()
            try {
                $children = @(Get-ChildItem -LiteralPath $basePath -Directory -ErrorAction SilentlyContinue | Sort-Object Name)
                foreach ($child in $children) {
                    $childNode = & $newFolderPickerNodeLocal -Path $child.FullName -Label $child.Name
                    if ($childNode) { [void]$Node.Nodes.Add($childNode) }
                }
            } catch { }
        }.GetNewClosure()

        $form = New-Object System.Windows.Forms.Form
        $form.Text = $Title
        $form.Size = [System.Drawing.Size]::new(760, 560)
        $form.MinimumSize = [System.Drawing.Size]::new(680, 500)
        $form.StartPosition = 'CenterParent'
        $form.FormBorderStyle = 'Sizable'
        $form.MaximizeBox = $false
        $form.MinimizeBox = $false
        $form.BackColor = $clrBg

        $headerCard = _Panel 10 10 724 86 $clrPanel
        $headerCard.Anchor = 'Top,Left,Right'
        $headerCard.BorderStyle = 'FixedSingle'
        $form.Controls.Add($headerCard)

        $headerAccent = _Panel 0 0 4 86 $clrAccentAlt
        $headerAccent.BorderStyle = 'None'
        $headerCard.Controls.Add($headerAccent)

        $lblTitle = _Label $Title 14 12 620 24 $fontTitle
        $lblTitle.Anchor = 'Top,Left,Right'
        $headerCard.Controls.Add($lblTitle)

        $lblPrompt = _Label $Prompt 14 42 690 34 $fontLabel
        $lblPrompt.ForeColor = $clrTextSoft
        $lblPrompt.Anchor = 'Top,Left,Right'
        $headerCard.Controls.Add($lblPrompt)

        $contentCard = _Panel 10 106 724 392 $clrPanel
        $contentCard.Anchor = 'Top,Left,Right,Bottom'
        $contentCard.BorderStyle = 'FixedSingle'
        $form.Controls.Add($contentCard)

        $lblCurrent = _Label 'Selected Folder' 12 10 180 20 $fontBold
        $contentCard.Controls.Add($lblCurrent)

        $tbCurrent = _TextBox 12 34 560 24 '' $false
        $tbCurrent.Anchor = 'Top,Left,Right'
        $tbCurrent.ReadOnly = $true
        $contentCard.Controls.Add($tbCurrent)

        $btnUp = _Button 'Up' 596 33 52 26 $clrPanelAlt $null
        $btnUp.Anchor = 'Top,Right'
        _SetMainControlToolTip -Control $btnUp -Text 'Move to the parent folder of the current selection.'
        $contentCard.Controls.Add($btnUp)

        $btnRefreshTree = _Button 'Refresh' 654 33 68 26 $clrPanelAlt $null
        $btnRefreshTree.Anchor = 'Top,Right'
        _SetMainControlToolTip -Control $btnRefreshTree -Text 'Reload the folder tree from disk.'
        $contentCard.Controls.Add($btnRefreshTree)

        $tree = New-Object System.Windows.Forms.TreeView
        $tree.Location = [System.Drawing.Point]::new(12, 70)
        $tree.Size = [System.Drawing.Size]::new(700, 308)
        $tree.Anchor = 'Top,Left,Right,Bottom'
        $tree.HideSelection = $false
        $tree.FullRowSelect = $true
        $tree.BackColor = [System.Drawing.Color]::FromArgb(30,30,40)
        $tree.ForeColor = $clrText
        $tree.BorderStyle = 'FixedSingle'
        $tree.Font = $fontLabel
        $contentCard.Controls.Add($tree)

        $footerLine = New-Object System.Windows.Forms.Panel
        $footerLine.Location = [System.Drawing.Point]::new(10, ($form.ClientSize.Height - 54))
        $footerLine.Size = [System.Drawing.Size]::new($form.ClientSize.Width - 20, 1)
        $footerLine.Anchor = 'Left,Right,Bottom'
        $footerLine.BackColor = $clrBorder
        $form.Controls.Add($footerLine)

        $footer = New-Object System.Windows.Forms.FlowLayoutPanel
        $footer.Location = [System.Drawing.Point]::new(10, ($form.ClientSize.Height - 42))
        $footer.Size = [System.Drawing.Size]::new($form.ClientSize.Width - 20, 32)
        $footer.Anchor = 'Right,Bottom'
        $footer.WrapContents = $false
        $footer.AutoScroll = $false
        $footer.AutoSize = $true
        $footer.AutoSizeMode = 'GrowAndShrink'
        $footer.FlowDirection = [System.Windows.Forms.FlowDirection]::RightToLeft
        $footer.BackColor = [System.Drawing.Color]::Transparent
        $form.Controls.Add($footer)

        $script:_FolderPickerResult = $null

        $btnCancel = _Button 'Cancel' 0 0 92 28 $clrMuted {
            $script:_FolderPickerResult = $null
            $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $form.Close()
        }
        $btnCancel.Margin = [System.Windows.Forms.Padding]::new(0)
        $footer.Controls.Add($btnCancel)

        $btnSelect = _Button 'Select Folder' 0 0 126 28 $clrGreen {
            $path = $tbCurrent.Text.Trim()
            if ([string]::IsNullOrWhiteSpace($path) -or -not (Test-Path -LiteralPath $path)) {
                [System.Windows.Forms.MessageBox]::Show(
                    'Choose a valid folder before continuing.',
                    $Title,'OK','Information') | Out-Null
                return
            }
            $script:_FolderPickerResult = $path
            $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $form.Close()
        }
        $btnSelect.Margin = [System.Windows.Forms.Padding]::new(10, 0, 0, 0)
        $footer.Controls.Add($btnSelect)

        $form.AcceptButton = $btnSelect
        $form.CancelButton = $btnCancel

        $selectTreePath = $null
        $selectTreePath = {
            param([string]$TargetPath)

            if ([string]::IsNullOrWhiteSpace($TargetPath)) { return $false }
            $resolved = $TargetPath
            try { $resolved = [System.IO.Path]::GetFullPath($TargetPath) } catch { $resolved = $TargetPath }
            if (-not (Test-Path -LiteralPath $resolved)) { return $false }

            $root = [System.IO.Path]::GetPathRoot($resolved)
            if ([string]::IsNullOrWhiteSpace($root)) { return $false }

            $rootNode = $null
            foreach ($candidateNode in @($tree.Nodes)) {
                if ([string]$candidateNode.Tag -ieq $root.TrimEnd('\')) {
                    $rootNode = $candidateNode
                    break
                }
            }
            if (-not $rootNode) { return $false }

            $currentNode = $rootNode
            & $loadFolderPickerChildrenLocal -Node $currentNode
            $relative = $resolved.Substring($root.Length).Trim('\')
            if (-not [string]::IsNullOrWhiteSpace($relative)) {
                $segments = @($relative -split '\\')
                $currentPath = $root.TrimEnd('\')
                foreach ($segment in $segments) {
                    if ([string]::IsNullOrWhiteSpace($segment)) { continue }
                    $currentPath = (Join-Path $currentPath $segment)
                    & $loadFolderPickerChildrenLocal -Node $currentNode
                    $nextNode = $null
                    foreach ($childNode in @($currentNode.Nodes)) {
                        if ([string]$childNode.Tag -ieq $currentPath) {
                            $nextNode = $childNode
                            break
                        }
                    }
                    if (-not $nextNode) { break }
                    $currentNode = $nextNode
                }
            }

            $tree.SelectedNode = $currentNode
            $currentNode.EnsureVisible()
            $currentNode.Expand()
            return $true
        }.GetNewClosure()

        $refreshRoots = {
            $tree.BeginUpdate()
            $tree.Nodes.Clear()
            foreach ($drive in @(Get-PSDrive -PSProvider FileSystem | Sort-Object Name)) {
                $drivePath = ''
                try { $drivePath = $drive.Root } catch { $drivePath = '' }
                if ([string]::IsNullOrWhiteSpace($drivePath) -or -not (Test-Path -LiteralPath $drivePath)) { continue }
                $driveNode = & $newFolderPickerNodeLocal -Path $drivePath -Label $drive.Name
                if ($driveNode) { [void]$tree.Nodes.Add($driveNode) }
            }
            $tree.EndUpdate()

            $pathToSelect = $InitialPath
            if ([string]::IsNullOrWhiteSpace($pathToSelect)) {
                try { $pathToSelect = [Environment]::GetFolderPath('Desktop') } catch { $pathToSelect = '' }
            }
            if (-not (& $selectTreePath $pathToSelect)) {
                if ($tree.Nodes.Count -gt 0) {
                    $tree.SelectedNode = $tree.Nodes[0]
                    $tree.Nodes[0].EnsureVisible()
                }
            }
        }.GetNewClosure()

        $tree.Add_BeforeExpand({
            param($sender, $e)
            & $loadFolderPickerChildrenLocal -Node $e.Node
        }.GetNewClosure())
        $tree.Add_AfterSelect({
            $selectedNode = $this.SelectedNode
            if ($selectedNode -and $selectedNode.Tag -and [string]$selectedNode.Tag -ne '__placeholder__') {
                $tbCurrent.Text = [string]$selectedNode.Tag
            }
        })
        $tree.Add_DoubleClick({
            if ($tree.SelectedNode -and $tree.SelectedNode.Tag -and [string]$tree.SelectedNode.Tag -ne '__placeholder__') {
                $path = [string]$tree.SelectedNode.Tag
                if (Test-Path -LiteralPath $path) {
                    $script:_FolderPickerResult = $path
                    $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
                    $form.Close()
                }
            }
        }.GetNewClosure())

        $btnRefreshTree.Add_Click({
            & $refreshRoots
        }.GetNewClosure())

        $btnUp.Add_Click({
            $current = $tbCurrent.Text.Trim()
            if ([string]::IsNullOrWhiteSpace($current) -or -not (Test-Path -LiteralPath $current)) { return }
            try {
                $parentInfo = [System.IO.Directory]::GetParent($current)
                if ($parentInfo -and $parentInfo.FullName) {
                    [void](& $selectTreePath $parentInfo.FullName)
                }
            } catch { }
        }.GetNewClosure())

        $layoutFolderPicker = {
            $margin = 10
            $headerCard.Size = [System.Drawing.Size]::new($form.ClientSize.Width - ($margin * 2), 86)
            $contentTop = 106
            $footerTop = $form.ClientSize.Height - 42
            $footerLine.SetBounds($margin, $footerTop - 12, $form.ClientSize.Width - ($margin * 2), 1)
            $footer.PerformLayout()
            $footer.Location = [System.Drawing.Point]::new([Math]::Max($margin, $form.ClientSize.Width - $margin - $footer.PreferredSize.Width), $footerTop)
            $contentCard.SetBounds($margin, $contentTop, $form.ClientSize.Width - ($margin * 2), [Math]::Max(240, $footerLine.Top - $contentTop - 10))
            $lblTitle.Width = [Math]::Max(240, $headerCard.ClientSize.Width - 28)
            $lblPrompt.Size = [System.Drawing.Size]::new([Math]::Max(240, $headerCard.ClientSize.Width - 28), 34)
            $tbCurrent.Width = [Math]::Max(220, $contentCard.ClientSize.Width - 148)
            $btnUp.Location = [System.Drawing.Point]::new($contentCard.ClientSize.Width - 124, 33)
            $btnRefreshTree.Location = [System.Drawing.Point]::new($contentCard.ClientSize.Width - 70, 33)
            $tree.Size = [System.Drawing.Size]::new([Math]::Max(240, $contentCard.ClientSize.Width - 24), [Math]::Max(150, $contentCard.ClientSize.Height - 82))
        }.GetNewClosure()
        $form.Add_Resize({ & $layoutFolderPicker }.GetNewClosure())

        & $refreshRoots
        & $layoutFolderPicker

        $dialogResult = if ($Owner) { $form.ShowDialog($Owner) } else { $form.ShowDialog() }
        $result = $script:_FolderPickerResult
        $script:_FolderPickerResult = $null
        if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
        if ([string]::IsNullOrWhiteSpace($result)) { return $null }
        return $result
    }

    function _SelectServerFolderForWizard {
        param([string]$KnownGameName = '')

        $detected = @()
        if (-not [string]::IsNullOrWhiteSpace($KnownGameName)) {
            try { $detected = @(_GetAutoDetectedServerFolders -KnownGameName $KnownGameName) } catch { $detected = @() }
        }

        if ($detected.Count -eq 1) {
            $candidate = $detected[0]
            $msg = "ECC found a likely $KnownGameName server folder:`n`n$($candidate.Path)`n`nUse this folder?"
            $choice = [System.Windows.Forms.MessageBox]::Show($msg, 'Detected Server Folder', 'YesNoCancel', 'Question')
            if ($choice -eq [System.Windows.Forms.DialogResult]::Yes) {
                return $candidate.Path
            }
            if ($choice -eq [System.Windows.Forms.DialogResult]::Cancel) {
                return $null
            }
        } elseif ($detected.Count -gt 1) {
            $pickForm = New-Object System.Windows.Forms.Form
            $pickForm.Text = "Select $KnownGameName Server Folder"
            $pickForm.Size = [System.Drawing.Size]::new(720, 430)
            $pickForm.MinimumSize = [System.Drawing.Size]::new(720, 430)
            $pickForm.BackColor = $clrBg
            $pickForm.StartPosition = 'CenterParent'
            $pickForm.FormBorderStyle = 'Sizable'
            $pickForm.MaximizeBox = $false
            $pickForm.MinimizeBox = $false

            $lbl = _Label "ECC found $($detected.Count) likely $KnownGameName folders. Pick one, or browse manually." 12 12 680 36
            $lbl.ForeColor = $clrTextSoft
            $lbl.Anchor = 'Top,Left,Right'
            $pickForm.Controls.Add($lbl)

            $list = New-Object System.Windows.Forms.ListBox
            $list.Location = [System.Drawing.Point]::new(12, 58)
            $list.Size = [System.Drawing.Size]::new(680, 280)
            $list.Anchor = 'Top,Left,Right,Bottom'
            $list.BackColor = [System.Drawing.Color]::FromArgb(26, 31, 44)
            $list.ForeColor = $clrText
            $list.Font = $fontMono
            $list.HorizontalScrollbar = $true
            $pickForm.Controls.Add($list)

            $footerPanel = New-Object System.Windows.Forms.Panel
            $footerPanel.Location = [System.Drawing.Point]::new(12, 344)
            $footerPanel.Size = [System.Drawing.Size]::new(680, 44)
            $footerPanel.Anchor = 'Left,Right,Bottom'
            $footerPanel.BackColor = [System.Drawing.Color]::Transparent
            $pickForm.Controls.Add($footerPanel)

            $buttonRow = New-Object System.Windows.Forms.FlowLayoutPanel
            $buttonRow.Location = [System.Drawing.Point]::new(0, 6)
            $buttonRow.Size = [System.Drawing.Size]::new($footerPanel.ClientSize.Width, 32)
            $buttonRow.Anchor = 'Top,Right'
            $buttonRow.WrapContents = $false
            $buttonRow.AutoScroll = $false
            $buttonRow.AutoSize = $true
            $buttonRow.AutoSizeMode = 'GrowAndShrink'
            $buttonRow.FlowDirection = [System.Windows.Forms.FlowDirection]::RightToLeft
            $buttonRow.BackColor = [System.Drawing.Color]::Transparent
            $footerPanel.Controls.Add($buttonRow)

            $pathLookup = @{}
            foreach ($candidate in $detected) {
                $display = if ([string]::IsNullOrWhiteSpace($candidate.Executable)) {
                    "$($candidate.Path)"
                } else {
                    "$($candidate.Path)   [$($candidate.Executable)]"
                }
                $pathLookup[$display] = $candidate.Path
                [void]$list.Items.Add($display)
            }
            if ($list.Items.Count -gt 0) { $list.SelectedIndex = 0 }

            $btnUse = _Button 'Use Selected' 12 350 132 30 $clrGreen {
                if ($list.SelectedItem) {
                    $script:_DetectedServerFolderSelection = $pathLookup[[string]$list.SelectedItem]
                    $pickForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
                    $pickForm.Close()
                }
            }
            $btnUse.Margin = [System.Windows.Forms.Padding]::new(10, 0, 0, 0)
            _SetMainControlToolTip -Control $btnUse -Text 'Use the highlighted detected server folder for this new profile.'

            $btnBrowseDetected = _Button 'Browse Manually' 154 350 144 30 $clrPanelAlt {
                $script:_DetectedServerFolderSelection = '__browse__'
                $pickForm.DialogResult = [System.Windows.Forms.DialogResult]::Retry
                $pickForm.Close()
            }
            $btnBrowseDetected.Margin = [System.Windows.Forms.Padding]::new(10, 0, 0, 0)
            _SetMainControlToolTip -Control $btnBrowseDetected -Text 'Skip the detected folders and browse to a server folder yourself.'

            $btnCancelDetected = _Button 'Cancel' 308 350 96 30 $clrMuted {
                $script:_DetectedServerFolderSelection = $null
                $pickForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                $pickForm.Close()
            }
            $btnCancelDetected.Margin = [System.Windows.Forms.Padding]::new(0)
            _SetMainControlToolTip -Control $btnCancelDetected -Text 'Close this picker without choosing a detected server folder.'
            $buttonRow.Controls.Add($btnCancelDetected)
            $buttonRow.Controls.Add($btnBrowseDetected)
            $buttonRow.Controls.Add($btnUse)

            $pickForm.AcceptButton = $btnUse
            $pickForm.CancelButton = $btnCancelDetected
            $layoutDetectedFolderPicker = {
                $margin = 12
                $footerHeight = 44
                $footerTop = $pickForm.ClientSize.Height - $footerHeight - $margin
                $footerPanel.SetBounds($margin, $footerTop, $pickForm.ClientSize.Width - ($margin * 2), $footerHeight)
                $buttonRow.PerformLayout()
                $buttonRow.Location = [System.Drawing.Point]::new([Math]::Max(0, $footerPanel.ClientSize.Width - $buttonRow.PreferredSize.Width), 6)
                $list.SetBounds($margin, 58, $pickForm.ClientSize.Width - ($margin * 2), [Math]::Max(180, $footerPanel.Top - 64))
                $lbl.Size = [System.Drawing.Size]::new($pickForm.ClientSize.Width - ($margin * 2), 36)
            }.GetNewClosure()
            $pickForm.Add_Resize({ & $layoutDetectedFolderPicker }.GetNewClosure())
            $list.Add_DoubleClick({
                if ($list.SelectedItem) {
                    $script:_DetectedServerFolderSelection = $pathLookup[[string]$list.SelectedItem]
                    $pickForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
                    $pickForm.Close()
                }
            }.GetNewClosure())

            & $layoutDetectedFolderPicker
            $dialogResult = $pickForm.ShowDialog()
            $selectedFolder = $script:_DetectedServerFolderSelection
            $script:_DetectedServerFolderSelection = $null

            if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK -and -not [string]::IsNullOrWhiteSpace($selectedFolder)) {
                return $selectedFolder
            }
            if ($dialogResult -eq [System.Windows.Forms.DialogResult]::Cancel) {
                return $null
            }
        }

        return (_ShowEccFolderPicker -Title 'Select Server Folder' -Prompt $(if ([string]::IsNullOrWhiteSpace($KnownGameName)) { 'Select your game server folder.' } else { "Select the $KnownGameName server folder." }))
    }

    function _ShowCreateProfileWizard {
        param(
            [string]$KnownGameName = ''
        )

        $folder = _SelectServerFolderForWizard -KnownGameName $KnownGameName
        if ([string]::IsNullOrWhiteSpace($folder)) { return $false }

        $finalName = $KnownGameName
        if ([string]::IsNullOrWhiteSpace($finalName)) {
            $gameName = [System.IO.Path]::GetFileName($folder)

            $nameForm                  = New-Object System.Windows.Forms.Form
            $nameForm.Text             = 'Game Name'
            $nameForm.Size             = [System.Drawing.Size]::new(380, 170)
            $nameForm.MinimumSize      = [System.Drawing.Size]::new(380, 170)
            $nameForm.BackColor        = $clrBg
            $nameForm.StartPosition    = 'CenterParent'
            $nameForm.FormBorderStyle  = 'Sizable'
            $nameForm.MaximizeBox      = $false
            $nameForm.MinimizeBox      = $false
            $lblNamePrompt = _Label 'Enter a display name:' 10 10 340 22
            $lblNamePrompt.Anchor = 'Top,Left,Right'
            $nameForm.Controls.Add($lblNamePrompt)
            $tbName = _TextBox 10 36 340 24 $gameName
            $tbName.Anchor = 'Top,Left,Right'
            $nameForm.Controls.Add($tbName)
            $nameFooter = New-Object System.Windows.Forms.FlowLayoutPanel
            $nameFooter.Location = [System.Drawing.Point]::new(10, 74)
            $nameFooter.Size = [System.Drawing.Size]::new($nameForm.ClientSize.Width - 20, 32)
            $nameFooter.Anchor = 'Top,Right'
            $nameFooter.WrapContents = $false
            $nameFooter.AutoScroll = $false
            $nameFooter.AutoSize = $true
            $nameFooter.AutoSizeMode = 'GrowAndShrink'
            $nameFooter.FlowDirection = [System.Windows.Forms.FlowDirection]::RightToLeft
            $nameFooter.BackColor = [System.Drawing.Color]::Transparent
            $nameForm.Controls.Add($nameFooter)
            $btnCreateName = (_Button 'Create' 10 74 140 30 $clrGreen {
                $script:_newGameName       = $tbName.Text.Trim()
                $nameForm.DialogResult     = [System.Windows.Forms.DialogResult]::OK
                $nameForm.Close()
            })
            $btnCreateName.Margin = [System.Windows.Forms.Padding]::new(10, 0, 0, 0)
            _SetMainControlToolTip -Control $btnCreateName -Text 'Create the new profile using this display name.'
            $btnCancelName = (_Button 'Cancel' 160 74 80 30 $clrMuted {
                $nameForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                $nameForm.Close()
            })
            $btnCancelName.Margin = [System.Windows.Forms.Padding]::new(0)
            _SetMainControlToolTip -Control $btnCancelName -Text 'Close this name prompt without creating the profile.'
            $nameFooter.Controls.Add($btnCancelName)
            $nameFooter.Controls.Add($btnCreateName)
            $layoutGameNameDialog = {
                $margin = 10
                $lblNamePrompt.Size = [System.Drawing.Size]::new($nameForm.ClientSize.Width - ($margin * 2), 22)
                $tbName.Size = [System.Drawing.Size]::new($nameForm.ClientSize.Width - ($margin * 2), 24)
                $nameFooter.PerformLayout()
                $nameFooter.Location = [System.Drawing.Point]::new([Math]::Max($margin, $nameForm.ClientSize.Width - $margin - $nameFooter.PreferredSize.Width), 74)
            }.GetNewClosure()
            $nameForm.Add_Resize({ & $layoutGameNameDialog }.GetNewClosure())
            $nameForm.AcceptButton = $btnCreateName
            $nameForm.CancelButton = $btnCancelName
            & $layoutGameNameDialog
            if ($nameForm.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return $false }

            $finalName = if ($script:_newGameName) { $script:_newGameName } else { $gameName }
            $script:_newGameName = $null
        }

        $configFolder = ''
        if ([string]::IsNullOrWhiteSpace($KnownGameName)) {
            $cfgPrompt = [System.Windows.Forms.MessageBox]::Show(
                'Would you like to select a config folder? (Optional)',
                'Config Folder','YesNo','Question')
            if ($cfgPrompt -eq [System.Windows.Forms.DialogResult]::Yes) {
                $configFolder = _ShowEccFolderPicker -Title 'Select Config Folder' -Prompt 'Select your server config folder. This step is optional.' -InitialPath $folder
            }
        }

        try {
            $pmPath = Join-Path $script:ModuleRoot 'ProfileManager.psm1'
            Import-Module $pmPath -Force
            $newProfile = New-GameProfile -FolderPath $folder -GameName $finalName -KnownGame $KnownGameName
            if (-not [string]::IsNullOrWhiteSpace($configFolder)) {
                $newProfile.ConfigRoot = $configFolder
            }

            $outPath = Save-GameProfile -Profile $newProfile -ProfilesDir $script:ProfilesDir
            $pfx = $newProfile.Prefix.ToUpper()

            $script:SharedState.Profiles[$pfx] = $newProfile
            $script:_SelectedProfilePrefix = $pfx

            _BuildProfilesList
            _BuildProfileEditor -Profile $newProfile
            _BuildServerDashboard

            [System.Windows.Forms.MessageBox]::Show(
                "Profile created:`nGame:   $($newProfile.GameName)`nPrefix: $pfx`nFile:   $outPath`n`nUse !${pfx}start in Discord.",
                'Created','OK','Information') | Out-Null
            return $true
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to create profile:`n$_",'Error','OK','Error') | Out-Null
            return $false
        }
    }

    function _GetProfileCreationChoices {
        return @(
            @{ Label = 'Project Zomboid'; GameName = 'Project Zomboid'; Color = $clrAccent },
            @{ Label = 'Minecraft'; GameName = 'Minecraft'; Color = $clrAccentAlt },
            @{ Label = 'Palworld'; GameName = 'Palworld'; Color = $clrAccentAlt },
            @{ Label = 'Satisfactory'; GameName = 'Satisfactory'; Color = $clrAccent },
            @{ Label = 'Valheim'; GameName = 'Valheim'; Color = $clrAccentAlt },
            @{ Label = '7 Days to Die'; GameName = '7 Days to Die'; Color = $clrAccent },
            @{ Label = 'Hytale'; GameName = 'Hytale'; Color = $clrAccentAlt },
            @{ Label = 'Custom / Other'; GameName = '__custom__'; Color = $clrGreen }
        )
    }

    function _ShowCreateProfileTypePicker {
        $choices = @(_GetProfileCreationChoices)
        $buttonW = 156
        $buttonH = 32
        $pickerHeight = 420

        $picker = New-Object System.Windows.Forms.Form
        $picker.Text = 'Add Game'
        $picker.Size = [System.Drawing.Size]::new(420, $pickerHeight)
        $picker.MinimumSize = [System.Drawing.Size]::new(420, $pickerHeight)
        $picker.StartPosition = 'CenterParent'
        $picker.FormBorderStyle = 'Sizable'
        $picker.MaximizeBox = $false
        $picker.MinimizeBox = $false
        $picker.BackColor = $clrBg

        $pickerTitle = _Label 'Choose a game type' 12 12 260 24 $fontTitle
        $pickerTitle.Anchor = 'Top,Left,Right'
        $picker.Controls.Add($pickerTitle)

        $pickerHint = _Label 'Choose a supported game to use a built-in profile template, or choose Custom / Other to build the profile manually.' 12 42 360 40 $fontLabel
        $pickerHint.ForeColor = $clrTextSoft
        $pickerHint.Anchor = 'Top,Left,Right'
        $picker.Controls.Add($pickerHint)

        $choicesPanel = New-Object System.Windows.Forms.FlowLayoutPanel
        $choicesPanel.Location = [System.Drawing.Point]::new(12, 92)
        $choicesPanel.Size = [System.Drawing.Size]::new($picker.ClientSize.Width - 24, 144)
        $choicesPanel.Anchor = 'Top,Left,Right,Bottom'
        $choicesPanel.WrapContents = $true
        $choicesPanel.AutoScroll = $true
        $choicesPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
        $choicesPanel.BackColor = [System.Drawing.Color]::Transparent
        $picker.Controls.Add($choicesPanel)

        foreach ($choice in $choices) {
            $btn = _Button $choice.Label 0 0 $buttonW $buttonH $choice.Color {
                $script:_SelectedCreateProfileGame = "$($this.Tag)"
                $picker.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $picker.Close()
            }
            $btn.Tag = $choice.GameName
            $btn.Font = $fontBold
            $btn.Margin = [System.Windows.Forms.Padding]::new(0, 0, 12, 10)
            $choicesPanel.Controls.Add($btn)
        }

        $syncPickerChoiceButtons = {
            $usableWidth = [Math]::Max(220, $choicesPanel.ClientSize.Width - [System.Windows.Forms.SystemInformation]::VerticalScrollBarWidth)
            $buttonWidth = [Math]::Max(140, [Math]::Floor(($usableWidth - 12) / 2))
            foreach ($ctrl in @($choicesPanel.Controls)) {
                if ($ctrl -is [System.Windows.Forms.Button]) {
                    $ctrl.Width = $buttonWidth
                }
            }
        }
        $choicesPanel.Add_SizeChanged({ & $syncPickerChoiceButtons }.GetNewClosure())
        & $syncPickerChoiceButtons

        $footerLine = New-Object System.Windows.Forms.Panel
        $footerLine.Location = [System.Drawing.Point]::new(12, ($picker.ClientSize.Height - 54))
        $footerLine.Size = [System.Drawing.Size]::new($picker.ClientSize.Width - 24, 1)
        $footerLine.Anchor = 'Left,Right,Bottom'
        $footerLine.BackColor = $clrBorder
        $picker.Controls.Add($footerLine)

        $pickerFooter = New-Object System.Windows.Forms.FlowLayoutPanel
        $pickerFooter.Location = [System.Drawing.Point]::new(12, ($picker.ClientSize.Height - 42))
        $pickerFooter.Size = [System.Drawing.Size]::new($picker.ClientSize.Width - 24, 30)
        $pickerFooter.Anchor = 'Right,Bottom'
        $pickerFooter.WrapContents = $false
        $pickerFooter.AutoScroll = $false
        $pickerFooter.AutoSize = $true
        $pickerFooter.AutoSizeMode = 'GrowAndShrink'
        $pickerFooter.FlowDirection = [System.Windows.Forms.FlowDirection]::RightToLeft
        $pickerFooter.BackColor = [System.Drawing.Color]::Transparent
        $picker.Controls.Add($pickerFooter)

        $btnCancelPicker = _Button 'Cancel' 0 0 92 28 $clrMuted {
            $picker.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $picker.Close()
        }
        $btnCancelPicker.Margin = [System.Windows.Forms.Padding]::new(0)
        _SetMainControlToolTip -Control $btnCancelPicker -Text 'Close the add-game picker without creating a profile.'
        $pickerFooter.Controls.Add($btnCancelPicker)
        $picker.CancelButton = $btnCancelPicker
        $layoutCreateProfilePicker = {
            $margin = 12
            $footerTop = $picker.ClientSize.Height - 42
            $footerLine.SetBounds($margin, $footerTop - 12, $picker.ClientSize.Width - ($margin * 2), 1)
            $pickerFooter.PerformLayout()
            $pickerFooter.Location = [System.Drawing.Point]::new([Math]::Max($margin, $picker.ClientSize.Width - $margin - $pickerFooter.PreferredSize.Width), $footerTop)
            $choicesPanel.SetBounds($margin, 92, $picker.ClientSize.Width - ($margin * 2), [Math]::Max(120, $footerLine.Top - 102))
            $pickerHint.Size = [System.Drawing.Size]::new($picker.ClientSize.Width - ($margin * 2), 40)
            & $syncPickerChoiceButtons
        }.GetNewClosure()
        $picker.Add_Resize({ & $layoutCreateProfilePicker }.GetNewClosure())
        & $layoutCreateProfilePicker

        $script:_SelectedCreateProfileGame = $null
        if ($picker.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
            $script:_SelectedCreateProfileGame = $null
            return $false
        }

        $selected = $script:_SelectedCreateProfileGame
        $script:_SelectedCreateProfileGame = $null
        if ([string]::IsNullOrWhiteSpace($selected)) { return $false }

        if ($selected -eq '__custom__') {
            return (_ShowCreateProfileWizard)
        }

        return (_ShowCreateProfileWizard -KnownGameName $selected)
    }

    function _BuildFirstRunProfilesEmptyState {
        param([System.Windows.Forms.Control]$Parent)

        if ($null -eq $Parent) { return }

        $title = _Label 'Set up your first server' 12 12 520 24 $fontTitle
        $title.Anchor = 'Top,Left,Right'
        $Parent.Controls.Add($title)

        $subtitle = _Label 'Choose a supported game to create your first ECC profile, or use Custom / Other if you want the manual path.' 12 42 520 38
        $subtitle.ForeColor = $clrTextSoft
        $subtitle.Anchor = 'Top,Left,Right'
        $Parent.Controls.Add($subtitle)

        $choices = @(_GetProfileCreationChoices)

        $choicesPanel = New-Object System.Windows.Forms.FlowLayoutPanel
        $choicesPanel.Location = [System.Drawing.Point]::new(12, 92)
        $choicesPanel.Size = [System.Drawing.Size]::new([Math]::Max(140, $Parent.ClientSize.Width - 24), [Math]::Max(160, $Parent.ClientSize.Height - 136))
        $choicesPanel.Anchor = 'Top,Left,Right,Bottom'
        $choicesPanel.WrapContents = $false
        $choicesPanel.AutoScroll = $true
        $choicesPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
        $choicesPanel.BackColor = [System.Drawing.Color]::Transparent
        $Parent.Controls.Add($choicesPanel)

        foreach ($choice in $choices) {
            $btnWidth = [Math]::Max(180, $choicesPanel.ClientSize.Width - 8)
            $btn = _Button $choice.Label 0 0 $btnWidth 32 $choice.Color {
                $selected = "$($this.Tag)"
                if ($selected -eq '__custom__') {
                    _ShowCreateProfileWizard | Out-Null
                } else {
                    _ShowCreateProfileWizard -KnownGameName $selected | Out-Null
                }
            }
            $btn.Tag = $choice.GameName
            $btn.Font = $fontBold
            $btn.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 10)
            $choicesPanel.Controls.Add($btn)
        }

        $syncFirstRunButtons = {
            $buttonWidth = [Math]::Max(180, $choicesPanel.ClientSize.Width - [System.Windows.Forms.SystemInformation]::VerticalScrollBarWidth - 8)
            foreach ($ctrl in @($choicesPanel.Controls)) {
                if ($ctrl -is [System.Windows.Forms.Button]) {
                    $ctrl.Width = $buttonWidth
                }
            }
        }
        $choicesPanel.Add_SizeChanged({ & $syncFirstRunButtons }.GetNewClosure())
        & $syncFirstRunButtons

        $note = _Label 'Choose a game above, or use Custom / Other if you want the manual path.' 12 ($Parent.ClientSize.Height - 28) 520 20
        $note.ForeColor = $clrTextSoft
        $note.Anchor = 'Left,Right,Bottom'
        $Parent.Controls.Add($note)
    }

    # =====================================================================
    # PROFILES LIST  (LEFT COLUMN)
    # =====================================================================
    function _GetProfilesPaneLayout {
        param(
            [int]$PanelWidth,
            [int]$PanelHeight
        )

        $profilesTopMargin = 12
        $profilesSideMargin = 12
        $profilesFooterMargin = 10
        $profilesFooterHeight = 48
        $profilesGapAboveFooter = 10
        $footerTop = [Math]::Max(
            ($profilesTopMargin + 80),
            ($PanelHeight - $profilesFooterHeight - $profilesFooterMargin)
        )
        $listHeight = [Math]::Max(80, ($footerTop - $profilesTopMargin - $profilesGapAboveFooter))

        return @{
            TopMargin       = $profilesTopMargin
            SideMargin      = $profilesSideMargin
            FooterMargin    = $profilesFooterMargin
            FooterHeight    = $profilesFooterHeight
            GapAboveFooter  = $profilesGapAboveFooter
            FooterTop       = $footerTop
            ListHeight      = $listHeight
            ListWidth       = [Math]::Max(120, ($PanelWidth - ($profilesSideMargin * 2)))
            FooterWidth     = [Math]::Max(120, ($PanelWidth - 16))
        }
    }

    function _BuildProfilesList {
        $panel = $script:_ProfilesPanel
        if ($null -eq $panel) { return }
        $savedListScroll = $null
        try {
            if ($script:_ProfilesListPanel) {
                $savedListScroll = _CaptureScrollPosition $script:_ProfilesListPanel
            }
        } catch { }
        $panel.Controls.Clear()

        $layout = _GetProfilesPaneLayout -PanelWidth $panel.Width -PanelHeight $panel.Height

        $listPanel = New-Object System.Windows.Forms.Panel
        $listPanel.Location   = [System.Drawing.Point]::new($layout.SideMargin, $layout.TopMargin)
        $listPanel.Size       = [System.Drawing.Size]::new($layout.ListWidth, $layout.ListHeight)
        $listPanel.Anchor     = 'Top,Left,Right,Bottom'
        $listPanel.BackColor  = $clrPanelSoft
        $listPanel.AutoScroll = $true
        $panel.Controls.Add($listPanel)
        $script:_ProfilesListPanel = $listPanel

        $ss = $script:SharedState
        $rowH = 32
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
                    [System.Drawing.Color]::FromArgb(58,66,96)
                } else {
                    $clrPanelSoft
                }
                $row.Padding = [System.Windows.Forms.Padding]::new(0)

                $rowAccent = _Panel 0 0 4 $rowH $clrAccent
                $rowAccent.BorderStyle = 'None'
                $rowAccent.Visible = ($script:_SelectedProfilePrefix -eq $pfx)
                $row.Controls.Add($rowAccent)

                $lbl = _Label "[$pfx] $gn" 12 7 ($row.Width - 20) 18
                $lbl.ForeColor = $clrText
                $lbl.Cursor = 'Hand'
                $lbl.Anchor = 'Left,Right,Top'
                if ($script:_SelectedProfilePrefix -eq $pfx) { $lbl.Font = $fontBold }
                $row.Controls.Add($lbl)

                $sep = New-Object System.Windows.Forms.Panel
                $sep.Location = [System.Drawing.Point]::new(0, $rowH - 1)
                $sep.Size     = [System.Drawing.Size]::new($row.Width, 1)
                $sep.Anchor   = 'Left,Right,Bottom'
                $sep.BackColor = $clrBorder
                $row.Controls.Add($sep)

                _BindClickHandler -Control $row -Handler {
                    $p = $this.Tag
                    $script:_SelectedProfilePrefix = $p
                    $prof = $script:SharedState.Profiles[$p]
                    if ($prof) { _BuildProfileEditor -Profile $prof }
                    _BuildProfilesList
                }
                _BindClickHandler -Control $lbl -Handler {
                    $p = $this.Parent.Tag
                    $script:_SelectedProfilePrefix = $p
                    $prof = $script:SharedState.Profiles[$p]
                    if ($prof) { _BuildProfileEditor -Profile $prof }
                    _BuildProfilesList
                }

                $listPanel.Controls.Add($row)
                $y += $rowH
            }
        } else {
            _BuildFirstRunProfilesEmptyState -Parent $listPanel
        }

        _RestoreScrollPosition -Control $listPanel -Position $savedListScroll

        $footerPanel = _Panel 8 $layout.FooterTop $layout.FooterWidth $layout.FooterHeight $clrPanelSoft
        $footerPanel.Anchor = 'Left,Right,Bottom'
        $footerPanel.BorderStyle = 'FixedSingle'
        $panel.Controls.Add($footerPanel)
        $script:_ProfilesFooterPanel = $footerPanel

        $btnAdd = _Button '+ Add Game' 10 7 108 32 $clrGreen { _ShowCreateProfileTypePicker | Out-Null }
        $btnAdd.Anchor = 'Left,Top'
        $btnAdd.Font   = $fontBold
        _SetMainControlToolTip -Control $btnAdd
        $footerPanel.Controls.Add($btnAdd)

        $btnRemove = _Button 'Remove' ($footerPanel.Width - 118) 7 108 32 $clrRed {
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
        $btnRemove.Anchor = 'Top,Right'
        $btnRemove.Font   = $fontBold
        _SetMainControlToolTip -Control $btnRemove
        $footerPanel.Controls.Add($btnRemove)
    }

    function _GetRuntimeStateEntry {
        param(
            [string]$Prefix,
            [hashtable]$SharedState
        )

        if ([string]::IsNullOrWhiteSpace($Prefix) -or -not $SharedState -or -not $SharedState.ContainsKey('ServerRuntimeState')) {
            return $null
        }

        $key = ''
        try { $key = $Prefix.ToUpperInvariant() } catch { $key = $Prefix }

        try {
            if ($SharedState.ServerRuntimeState.ContainsKey($key)) {
                return $SharedState.ServerRuntimeState[$key]
            }
        } catch { }

        return $null
    }

    function _SetObservedPlayersRuntimeState {
        param(
            [string]$Prefix,
            [string[]]$Names = @(),
            [int]$Count = -1,
            [hashtable]$SharedState
        )

        if ([string]::IsNullOrWhiteSpace($Prefix) -or -not $SharedState) { return }
        $safeNames = @(_ToStringArray -Value $Names)
        $resolvedCount = if ($Count -ge 0) { [Math]::Max([int]$Count, $safeNames.Count) } else { $safeNames.Count }
        try {
            Set-ObservedPlayersServerRuntimeState -Prefix $Prefix -Count $resolvedCount -SharedState $SharedState
        } catch { }
    }

    function _ResolveDashboardHealthVisual {
        param(
            [string]$HealthCode,
            [bool]$Running = $false
        )

        $normalized = if ([string]::IsNullOrWhiteSpace($HealthCode)) { 'healthy' } else { $HealthCode.Trim().ToLowerInvariant() }
        switch ($normalized) {
            'error' {
                return [ordered]@{
                    Code  = 'error'
                    Text  = 'ERROR'
                    Color = $clrRed
                }
            }
            'warning' {
                return [ordered]@{
                    Code  = 'warning'
                    Text  = 'WARNING'
                    Color = $clrYellow
                }
            }
            'waiting' {
                return [ordered]@{
                    Code  = 'waiting'
                    Text  = 'WAITING'
                    Color = $clrAccentAlt
                }
            }
            default {
                if (-not $Running) {
                    return [ordered]@{
                        Code  = 'healthy'
                        Text  = 'READY'
                        Color = $clrTextSoft
                    }
                }
                return [ordered]@{
                    Code  = 'healthy'
                    Text  = 'HEALTHY'
                    Color = $clrGreen
                }
            }
        }
    }

    function _BuildDashboardHealthToolTip {
        param(
            [hashtable]$HealthSnapshot
        )

        if (-not $HealthSnapshot) { return '' }

        $tooltipLines = New-Object System.Collections.Generic.List[string]

        $healthText = ''
        try { $healthText = [string]$HealthSnapshot.HealthText } catch { $healthText = '' }
        if (-not [string]::IsNullOrWhiteSpace($healthText)) {
            $tooltipLines.Add(('Health: {0}' -f $healthText)) | Out-Null
        }

        $summary = ''
        try { $summary = [string]$HealthSnapshot.Summary } catch { $summary = '' }
        if (-not [string]::IsNullOrWhiteSpace($summary)) {
            $tooltipLines.Add($summary) | Out-Null
        }

        $playerSource = ''
        try { $playerSource = [string]$HealthSnapshot.PlayerSource } catch { $playerSource = '' }
        if (-not [string]::IsNullOrWhiteSpace($playerSource)) {
            $tooltipLines.Add(('Player source: {0}' -f $playerSource)) | Out-Null
        }

        $lastEventText = ''
        try { $lastEventText = [string]$HealthSnapshot.LastEventText } catch { $lastEventText = '' }
        $lastEventAge = ''
        try { $lastEventAge = [string]$HealthSnapshot.LastEventAge } catch { $lastEventAge = '' }
        if (-not [string]::IsNullOrWhiteSpace($lastEventText)) {
            $eventLine = 'Last event: ' + $lastEventText
            if (-not [string]::IsNullOrWhiteSpace($lastEventAge) -and $lastEventAge -ne 'unknown') {
                $eventLine += (' ({0})' -f $lastEventAge)
            }
            $tooltipLines.Add($eventLine) | Out-Null
        }

        $lastPlayerSeen = ''
        try { $lastPlayerSeen = [string]$HealthSnapshot.LastPlayerSeen } catch { $lastPlayerSeen = '' }
        if (-not [string]::IsNullOrWhiteSpace($lastPlayerSeen)) {
            $tooltipLines.Add(('Last player seen: {0}' -f $lastPlayerSeen)) | Out-Null
        }

        $conservativeMode = $false
        try { $conservativeMode = [bool]$HealthSnapshot.ConservativeMode } catch { $conservativeMode = $false }
        $conservativeNote = ''
        try { $conservativeNote = [string]$HealthSnapshot.ConservativeNote } catch { $conservativeNote = '' }
        if ($conservativeMode -and -not [string]::IsNullOrWhiteSpace($conservativeNote)) {
            $tooltipLines.Add(('Conservative mode: {0}' -f $conservativeNote)) | Out-Null
        }

        return ($tooltipLines -join [Environment]::NewLine).Trim()
    }

    function _ApplyDashboardCardToolTips {
        param(
            [System.Windows.Forms.Control]$Card,
            [string]$ToolTipText
        )

        if ($null -eq $Card -or [string]::IsNullOrWhiteSpace($ToolTipText)) { return }

        foreach ($controlName in @('lblName','lblSubtitle','lblSource','lblStatus','lblHealth','lblUptime','lblTimers')) {
            $target = $null
            try { $target = ($Card.Controls.Find($controlName, $true) | Select-Object -First 1) } catch { $target = $null }
            if ($target) {
                _SetMainControlToolTip -Control $target -Text $ToolTipText
            }
        }

        _SetMainControlToolTip -Control $Card -Text $ToolTipText
    }

    function _GetDashboardStateInfo {
        param(
            [string]$Prefix,
            [object]$Profile,
            [bool]$Running,
            [object]$Entry,
            [hashtable]$SharedState
        )

        $prefixKey = if ([string]::IsNullOrWhiteSpace($Prefix)) { '' } else { $Prefix.ToUpperInvariant() }
        $state = _GetRuntimeStateEntry -Prefix $Prefix -SharedState $SharedState
        $code = ''
        $detail = ''
        $stateSince = $null
        $activity = $null
        $activityNote = ''
        $activityDetectionSupported = $null
        $healthSnapshot = $null
        if ($state) {
            try { $code = [string]$state.Code } catch { $code = '' }
            try { $detail = [string]$state.Detail } catch { $detail = '' }
            try { $stateSince = [datetime]$state.Since } catch { $stateSince = $null }
        }
        if ($SharedState -and $SharedState.ContainsKey('PlayerActivityState') -and $SharedState.PlayerActivityState -and $SharedState.PlayerActivityState.ContainsKey($prefixKey)) {
            try { $activity = $SharedState.PlayerActivityState[$prefixKey] } catch { $activity = $null }
            if ($activity) {
                try { $activityNote = [string]$activity.Note } catch { $activityNote = '' }
                try { $activityDetectionSupported = [bool]$activity.DetectionSupported } catch { $activityDetectionSupported = $null }
            }
        }
        try {
            $healthSnapshot = Get-ProfileHealthSnapshot -Prefix $Prefix -Profile $Profile -SharedState $SharedState -Running:$Running
        } catch {
            $healthSnapshot = $null
        }

        $healthVisual = _ResolveDashboardHealthVisual -HealthCode $(if ($healthSnapshot) { [string]$healthSnapshot.HealthCode } else { 'healthy' }) -Running:$Running
        $healthMeta = ''
        $healthToolTip = _BuildDashboardHealthToolTip -HealthSnapshot $healthSnapshot
        try {
            $playerSourceText = [string]$healthSnapshot.PlayerSource
            if (-not [string]::IsNullOrWhiteSpace($playerSourceText)) {
                $healthMeta = ('Source: {0}' -f $playerSourceText)
            }
        } catch { $healthMeta = '' }
        if ([string]::IsNullOrWhiteSpace($healthMeta)) {
            $healthMeta = if ($Running) { 'Source: Waiting for runtime health data' } else { 'Source: Profile standby' }
        }

        $badgeText = if ($Running) { 'ONLINE' } else { 'OFFLINE' }
        $statusColor = if ($Running) { $clrGreen } else { $clrRed }
        $statusBg = if ($Running) { [System.Drawing.Color]::FromArgb(36, 78, 60) } else { [System.Drawing.Color]::FromArgb(86, 40, 46) }
        $accentColor = if ($Running) { $clrGreen } else { $clrRed }

        $normalizedCode = if ([string]::IsNullOrWhiteSpace($code)) { '' } else { $code.ToLowerInvariant() }
        if (-not $Running -and $normalizedCode -in @('online','starting','stopping','restarting','waiting_first_player','idle_wait')) {
            $normalizedCode = ''
            $detail = ''
        }
        if (-not $Running -and $normalizedCode -eq 'stopped' -and $null -ne $stateSince) {
            try {
                if (((Get-Date) - $stateSince).TotalSeconds -ge 8) {
                    $normalizedCode = ''
                    $detail = ''
                }
            } catch { }
        }
        switch ($normalizedCode) {
            'online' {
                $badgeText = 'ONLINE'
                $statusColor = $clrGreen
                $statusBg = [System.Drawing.Color]::FromArgb(36, 78, 60)
                $accentColor = $clrGreen
            }
            'starting' {
                $badgeText = 'STARTING'
                $statusColor = $clrAccentAlt
                $statusBg = [System.Drawing.Color]::FromArgb(42, 59, 102)
                $accentColor = $clrAccentAlt
            }
            'restarting' {
                $badgeText = 'RESTART'
                $statusColor = $clrYellow
                $statusBg = [System.Drawing.Color]::FromArgb(92, 72, 28)
                $accentColor = $clrYellow
            }
            'stopping' {
                $badgeText = 'STOPPING'
                $statusColor = $clrYellow
                $statusBg = [System.Drawing.Color]::FromArgb(92, 72, 28)
                $accentColor = $clrYellow
            }
            'waiting_restart' {
                $badgeText = 'WAITING'
                $statusColor = $clrYellow
                $statusBg = [System.Drawing.Color]::FromArgb(92, 72, 28)
                $accentColor = $clrYellow
            }
            'waiting_first_player' {
                $badgeText = 'WAITING'
                $statusColor = $clrYellow
                $statusBg = [System.Drawing.Color]::FromArgb(92, 72, 28)
                $accentColor = $clrYellow
            }
            'blocked' {
                $badgeText = 'BLOCKED'
                $statusColor = $clrYellow
                $statusBg = [System.Drawing.Color]::FromArgb(92, 72, 28)
                $accentColor = $clrYellow
            }
            'failed' {
                $badgeText = 'FAILED'
                $statusColor = $clrRed
                $statusBg = [System.Drawing.Color]::FromArgb(86, 40, 46)
                $accentColor = $clrRed
            }
            'startup_failed' {
                $badgeText = 'FAILED'
                $statusColor = $clrRed
                $statusBg = [System.Drawing.Color]::FromArgb(86, 40, 46)
                $accentColor = $clrRed
            }
            'stopped' {
                $badgeText = 'STOPPED'
                $statusColor = $clrRed
                $statusBg = [System.Drawing.Color]::FromArgb(86, 40, 46)
                $accentColor = $clrAccent
            }
            'idle_wait' {
                $badgeText = 'IDLE'
                $statusColor = $clrYellow
                $statusBg = [System.Drawing.Color]::FromArgb(92, 72, 28)
                $accentColor = $clrYellow
            }
            'idle_shutdown' {
                $badgeText = 'IDLE'
                $statusColor = $clrYellow
                $statusBg = [System.Drawing.Color]::FromArgb(92, 72, 28)
                $accentColor = $clrYellow
            }
        }

        $subtitle = if (-not [string]::IsNullOrWhiteSpace($detail)) {
            $detail
        } elseif ($Running -and $activity -and $activityDetectionSupported -eq $false -and -not [string]::IsNullOrWhiteSpace($activityNote)) {
            $activityNote
        } elseif ($Running) {
            'Server is running and being monitored'
        } else {
            'This server profile is ready for control and monitoring.'
        }

            return [ordered]@{
                BadgeText   = $badgeText
                StatusColor = $statusColor
                StatusBg    = $statusBg
                AccentColor = $accentColor
                HealthText  = $healthVisual.Text
                HealthColor = $healthVisual.Color
                HealthCode  = $healthVisual.Code
                HealthMeta  = $healthMeta
                HealthToolTip = $healthToolTip
                Subtitle    = $subtitle
                StateCode   = $normalizedCode
            }
        }

    # =====================================================================
    # TIMER LINE BUILDER
    # Returns a single-line string showing next restart countdown and last
    # save elapsed time.  Called during initial card build and on every
    # _UpdateDashboardStatus tick so the labels stay live.
    # =====================================================================
    function _BuildTimerLine {
        param(
            [string]$Prefix,
            [object]$Profile,
            [object]$Entry,
            [hashtable]$SharedState
        )

        $parts = [System.Collections.Generic.List[string]]::new()
        $prefixKey = if ([string]::IsNullOrWhiteSpace($Prefix)) { '' } else { $Prefix.ToUpperInvariant() }
        $runtimeState = _GetRuntimeStateEntry -Prefix $Prefix -SharedState $SharedState

        if ($SharedState -and $SharedState.ContainsKey('PendingAutoRestarts') -and $SharedState.PendingAutoRestarts.ContainsKey($prefixKey)) {
            try {
                $pending = $SharedState.PendingAutoRestarts[$prefixKey]
                $dueAt = [datetime]$pending.DueAt
                $remainingSeconds = [Math]::Max(0, [int][Math]::Ceiling(($dueAt - (Get-Date)).TotalSeconds))
                if ($remainingSeconds -le 0) {
                    $parts.Add('Restart: imminent')
                } else {
                    $parts.Add("Restart in: ${remainingSeconds}s")
                }
            } catch { }
        }

        if ($SharedState -and $SharedState.ContainsKey('PendingScheduledRestarts') -and $SharedState.PendingScheduledRestarts.ContainsKey($prefixKey)) {
            try {
                $pendingScheduled = $SharedState.PendingScheduledRestarts[$prefixKey]
                $dueAt = [datetime]$pendingScheduled.DueAt
                $remainingSeconds = [Math]::Max(0, [int][Math]::Ceiling(($dueAt - (Get-Date)).TotalSeconds))
                if ($remainingSeconds -le 0) {
                    $parts.Add('Restart retry: imminent')
                } else {
                    $parts.Add("Restart retry in: ${remainingSeconds}s")
                }
            } catch { }
        }

        if ($runtimeState -and $null -ne $Entry -and $null -ne $Entry.StartTime) {
            $runtimeCode = ''
            try { $runtimeCode = [string]$runtimeState.Code } catch { $runtimeCode = '' }
            if ($runtimeCode.ToLowerInvariant() -eq 'starting') {
                $startupTimeoutSeconds = 180
                try {
                    $candidate = 0
                    if ($null -ne $Profile.StartupTimeoutSeconds -and [int]::TryParse("$($Profile.StartupTimeoutSeconds)", [ref]$candidate) -and $candidate -gt 0) {
                        $startupTimeoutSeconds = $candidate
                    }
                } catch { }

                $startupRemaining = [Math]::Max(0, [int][Math]::Ceiling($startupTimeoutSeconds - ((Get-Date) - [datetime]$Entry.StartTime).TotalSeconds))
                if ($startupRemaining -le 0) {
                    $parts.Add('Startup: due')
                } else {
                    $parts.Add("Startup in: ${startupRemaining}s")
                }
            }
        }

        if ($runtimeState -and $SharedState -and $SharedState.ContainsKey('PlayerActivityState') -and $SharedState.PlayerActivityState.ContainsKey($prefixKey)) {
            try {
                $activity = $SharedState.PlayerActivityState[$prefixKey]
                $runtimeCode = ''
                try { $runtimeCode = [string]$runtimeState.Code } catch { $runtimeCode = '' }
                if ($activity -and $activity.ShutdownDueAt -and ($runtimeCode -eq 'waiting_first_player' -or $runtimeCode -eq 'idle_wait')) {
                    $dueAt = [datetime]$activity.ShutdownDueAt
                    $remainingSeconds = [Math]::Max(0, [int][Math]::Ceiling(($dueAt - (Get-Date)).TotalSeconds))
                    $pendingRule = ''
                    try { $pendingRule = [string]$activity.PendingRule } catch { $pendingRule = '' }
                    if ($remainingSeconds -le 0) {
                        if ($pendingRule -eq 'signal_wait') {
                            $parts.Add('Player data: overdue')
                        } else {
                            $parts.Add('Idle: due')
                        }
                    } elseif ($pendingRule -eq 'signal_wait') {
                        $parts.Add("Player data in: ${remainingSeconds}s")
                    } elseif ($runtimeCode -eq 'waiting_first_player') {
                        $parts.Add("First player in: ${remainingSeconds}s")
                    } else {
                        $parts.Add("Idle in: ${remainingSeconds}s")
                    }
                }
            } catch { }
        }

        # ── Next scheduled restart countdown ──────────────────────────────────
        $schedEnabled = $true
        if ($SharedState.Settings -and $SharedState.Settings.ContainsKey('ScheduledRestartEnabled')) {
            $schedEnabled = [bool]$SharedState.Settings.ScheduledRestartEnabled
        }
        if ($null -ne $Profile.ScheduledRestartEnabled) {
            $schedEnabled = $schedEnabled -and [bool]$Profile.ScheduledRestartEnabled
        }

        if ($null -ne $Entry -and $null -ne $Entry.StartTime) {
            if ($schedEnabled) {
                $intervalHours = 6.0
                if ($SharedState.Settings -and $SharedState.Settings.ContainsKey('ScheduledRestartHours')) {
                    $v = 0.0
                    if ([double]::TryParse("$($SharedState.Settings.ScheduledRestartHours)", [ref]$v) -and $v -gt 0) {
                        $intervalHours = $v
                    }
                }
                $intervalMin   = $intervalHours * 60.0
                $uptimeMin     = ((Get-Date) - [datetime]$Entry.StartTime).TotalMinutes
                $remainingMin  = [Math]::Max(0, $intervalMin - $uptimeMin)

                if ($remainingMin -le 0) {
                    $parts.Add('Restart: imminent')
                } elseif ($remainingMin -lt 60) {
                    $parts.Add("Restart in: $([Math]::Ceiling($remainingMin))m")
                } else {
                    $h = [Math]::Floor($remainingMin / 60)
                    $m = [Math]::Floor($remainingMin % 60)
                    $parts.Add("Restart in: ${h}h ${m}m")
                }
            } else {
                $parts.Add('Restart: disabled')
            }
        }

        # ── Last auto-save elapsed ────────────────────────────────────────────
        $saveEnabled = $true
        if ($SharedState.Settings -and $SharedState.Settings.ContainsKey('AutoSaveEnabled')) {
            $saveEnabled = [bool]$SharedState.Settings.AutoSaveEnabled
        }
        if ($null -ne $Profile.AutoSaveEnabled) {
            $saveEnabled = $saveEnabled -and [bool]$Profile.AutoSaveEnabled
        }
        $hasSaveMethod = ($null -ne $Profile.SaveMethod -and
                          "$($Profile.SaveMethod)".Trim() -ne '' -and
                          "$($Profile.SaveMethod)".Trim() -ne 'none')

        if ($null -ne $Entry -and $null -ne $Entry.StartTime -and $saveEnabled -and $hasSaveMethod) {
            $intervalSaveMin = 30
            if ($SharedState.Settings -and $SharedState.Settings.ContainsKey('AutoSaveIntervalMinutes')) {
                $v = 0
                if ([int]::TryParse("$($SharedState.Settings.AutoSaveIntervalMinutes)", [ref]$v) -and $v -gt 0) {
                    $intervalSaveMin = $v
                }
            }
            if ($null -ne $Profile.AutoSaveIntervalMinutes) {
                $v = 0
                if ([int]::TryParse("$($Profile.AutoSaveIntervalMinutes)", [ref]$v) -and $v -gt 0) {
                    $intervalSaveMin = $v
                }
            }

            if ($SharedState.ContainsKey('LastAutoSave') -and
                $SharedState.LastAutoSave.ContainsKey($prefixKey)) {
                $lastSave    = [datetime]$SharedState.LastAutoSave[$prefixKey]
                $elapsedSave = ((Get-Date) - $lastSave).TotalMinutes
                $nextSaveMin = [Math]::Max(0, $intervalSaveMin - $elapsedSave)

                if ($nextSaveMin -le 1) {
                    $parts.Add("Save in: <1m")
                } else {
                    $parts.Add("Save in: $([Math]::Ceiling($nextSaveMin))m")
                }
            } else {
                # Server just started - no save has run yet
                $parts.Add("Save in: ${intervalSaveMin}m")
            }
        } elseif ($null -ne $Entry -and $null -ne $Entry.StartTime -and -not $hasSaveMethod) {
            $parts.Add('Save: N/A')
        } elseif ($null -ne $Entry -and $null -ne $Entry.StartTime -and -not $saveEnabled) {
            $parts.Add('Save: disabled')
        }

        return ($parts -join '   |   ')
    }

    # =====================================================================
    # SERVER DASHBOARD  (CENTER COLUMN)
    # =====================================================================
    function _BuildServerDashboard {
        $panel = $script:_ServerDashboardPanel
        if ($null -eq $panel) { return }
        $script:_DashboardCardCache = @{}
        $savedDashboardScroll = $null
        try {
            if ($script:_DashboardScrollPanel) {
                $savedDashboardScroll = _CaptureScrollPosition $script:_DashboardScrollPanel
            }
        } catch { }
        $panel.Controls.Clear()

        $ss = $script:SharedState
        if ($null -eq $ss -or -not $ss.Profiles -or $ss.Profiles.Count -eq 0) {
            $lblEmpty = _Label 'No server profiles yet.' 10 10 500 22 $fontTitle
            $panel.Controls.Add($lblEmpty)

            $lblHint = _Label 'Use the quick-start buttons in the left panel, or click + Add Game there to create your first server profile.' 10 38 620 38
            $lblHint.ForeColor = $clrTextSoft
            $panel.Controls.Add($lblHint)

            $script:_DashboardScrollPanel = $null
            return
        }

        # Scrollable inner panel - cards live here, not directly on $panel
        $scroll            = New-Object System.Windows.Forms.Panel
        $scroll.Dock       = 'Fill'
        $scroll.AutoScroll = $true
        $scroll.BackColor  = $clrBg
        $panel.Controls.Add($scroll)
        $script:_DashboardScrollPanel = $scroll

        $y = 4
        foreach ($pfx in ($ss.Profiles.Keys | Sort-Object)) {
            $profile = $ss.Profiles[$pfx]

            $running = $false
            $entry   = $null
            try {
                $status = Get-ServerStatus -Prefix $pfx
                if ($status -and $status.Running) {
                    $running = $true
                    $entry = @{
                        Pid       = $status.Pid
                        StartTime = $status.StartTime
                    }
                }
            } catch { }

            $cardW = $scroll.ClientSize.Width - 24
            if ($cardW -lt 460) { $cardW = 460 }
            $cardH = 176
            $card        = _Panel 8 $y $cardW $cardH $clrPanel
            $card.Anchor = 'Top,Left,Right'
            $card.Padding = [System.Windows.Forms.Padding]::new(16, 12, 16, 12)
            $card.Cursor  = 'Hand'
            $card.Tag     = $pfx
            $card.TabStop = $false

            $stateInfo = _GetDashboardStateInfo -Prefix $pfx -Profile $profile -Running:$running -Entry $entry -SharedState $ss
            $statusText  = $stateInfo.BadgeText
            $statusColor = $stateInfo.StatusColor
            $statusBg    = $stateInfo.StatusBg
            $accentColor = $stateInfo.AccentColor

            $cardAccent = _Panel 0 0 4 $cardH $accentColor
            $cardAccent.BorderStyle = 'None'
            $cardAccent.Name = 'pnlAccent'
            $card.Controls.Add($cardAccent)

            $cardLeft = 18
            $headerTop = 10
            $statusBadgeWidth = 116
            $statusBadgeHeight = 44
            $statusBadgeRightMargin = 16
            $cardHeader = _Panel $cardLeft $headerTop ($cardW - ($cardLeft * 2)) 44 $clrPanel
            $cardHeader.Anchor = 'Top,Left,Right'
            $cardHeader.BorderStyle = 'None'
            $cardHeader.Name = 'pnlCardHeader'
            $card.Controls.Add($cardHeader)
            try { $cardHeader.SendToBack() } catch { }

            $isHytaleProfile = ((_NormalizeGameIdentity (_GetProfileKnownGame -Profile $profile)) -eq 'hytale')
            $headerReserve = if ($isHytaleProfile) { 386 } else { 248 }

            $nameText = "$($profile.GameName) [$pfx]"
            $namePreferredWidth = [Math]::Max(80, [System.Windows.Forms.TextRenderer]::MeasureText($nameText, $fontTitle).Width + 6)
            $lblName = _Label $nameText 0 8 ([Math]::Min([Math]::Max(220, $cardHeader.Width - $headerReserve), $namePreferredWidth)) 24 $fontTitle
            $lblName.Name = 'lblName'
            $lblName.Tag = if ($isHytaleProfile) { 'HytaleHeader' } else { 'DefaultHeader' }
            $lblName.AutoEllipsis = $true
            $cardHeader.Controls.Add($lblName)

            $lblSubtitle = _Label $stateInfo.Subtitle $cardLeft 58 ([Math]::Max(220, $cardW - 164)) 18 $fontLabel
            $lblSubtitle.ForeColor = $clrTextSoft
            $lblSubtitle.Name = 'lblSubtitle'
            $lblSubtitle.Tag = if ($isHytaleProfile) { 'HytaleHeader' } else { 'DefaultHeader' }
            $card.Controls.Add($lblSubtitle)

            $lblSource = _Label $stateInfo.HealthMeta $cardLeft 78 ([Math]::Max(220, $cardW - 164)) 16
            $lblSource.ForeColor = $clrTextSoft
            $lblSource.Name = 'lblSource'
            $lblSource.Tag = if ($isHytaleProfile) { 'HytaleHeader' } else { 'DefaultHeader' }
            $card.Controls.Add($lblSource)

            $statusBadgeX = $card.ClientSize.Width - $statusBadgeWidth - $statusBadgeRightMargin
            $statusBadge = _Panel $statusBadgeX $headerTop $statusBadgeWidth $statusBadgeHeight $statusBg
            $statusBadge.Anchor = 'Top,Right'
            $statusBadge.Name   = 'pnlStatus'
            $card.Controls.Add($statusBadge)
            try { $card.Controls.SetChildIndex($statusBadge, 0) } catch { }

            $lblStatus   = _Label $statusText 0 3 $statusBadgeWidth 18 $fontBold
            $lblStatus.Anchor    = 'Top,Left,Right'
            $lblStatus.ForeColor = $statusColor
            $lblStatus.Name      = 'lblStatus'
            $lblStatus.TextAlign = 'MiddleCenter'
            $statusBadge.Controls.Add($lblStatus)

            $lblHealth = _Label $stateInfo.HealthText 0 23 $statusBadgeWidth 16 $fontBold
            $lblHealth.Anchor = 'Top,Left,Right'
            $lblHealth.ForeColor = $stateInfo.HealthColor
            $lblHealth.Name = 'lblHealth'
            $lblHealth.TextAlign = 'MiddleCenter'
            $statusBadge.Controls.Add($lblHealth)

            if ($running -and $entry) {
                $up = [Math]::Round(((Get-Date) - $entry.StartTime).TotalMinutes, 1)
                $lblUptime = _Label "PID $($entry.Pid) | Uptime: ${up} min" $cardLeft 98 ($cardW - 244) 18 $fontBold
            } else {
                $lblUptime = _Label 'Server is not running' $cardLeft 98 ($cardW - 244) 18 $fontBold
            }
            $lblUptime.Name = 'lblUptime'
            $card.Controls.Add($lblUptime)

            $timerText = _BuildTimerLine -Prefix $pfx -Profile $profile -Entry $entry -SharedState $ss
            $lblTimers            = _Label $timerText $cardLeft 120 ($cardW - 264) 14
            $lblTimers.ForeColor  = $clrYellow
            $lblTimers.Font       = New-Object System.Drawing.Font('Consolas', 7.5)
            $lblTimers.Name       = 'lblTimers'
            $card.Controls.Add($lblTimers)

            $chk           = [System.Windows.Forms.CheckBox]::new()
            $chk.Text      = 'Auto-Restart'
            $chk.Location  = [System.Drawing.Point]::new($cardLeft, 143)
            $chk.Size      = [System.Drawing.Size]::new(136, 18)
            $chk.ForeColor = $clrText
            $chk.BackColor = [System.Drawing.Color]::Transparent
            $chk.Font      = $fontLabel
            $chk.Checked   = ($profile.EnableAutoRestart -eq $true)
            $chk.Tag       = $pfx
            $chk.TabStop   = $false
            $chk.Add_CheckedChanged({
                $script:SharedState.Profiles[$this.Tag].EnableAutoRestart = $this.Checked
                _ClearDashboardFocus -FallbackControl $script:_DashboardScrollPanel
            })
            _SetMainControlToolTip -Control $chk
            $card.Controls.Add($chk)

            $restartButtonColor  = [System.Drawing.Color]::FromArgb(170, 122, 58)
            $commandsButtonColor = [System.Drawing.Color]::FromArgb(76, 118, 196)
            $configButtonColor   = [System.Drawing.Color]::FromArgb(62, 142, 154)
            $toolsButtonColor    = [System.Drawing.Color]::FromArgb(92, 126, 198)

            $buttonY = 136
            $btnStart     = _Button 'Start'   210 $buttonY 80 30 $clrGreen $null
            $btnStart.Name = 'btnStart'
            $btnStart.Tag = $pfx
            $btnStart.TabStop = $false
            $btnStart.Add_Click({
                $p = $this.Tag
                _RunServerOpInBackground -Prefix $p -Operation 'Start'
                _ClearDashboardFocus -FallbackControl $script:_DashboardScrollPanel
            })
            _SetMainControlToolTip -Control $btnStart -Text ("Start {0}." -f $profile.GameName)
            $card.Controls.Add($btnStart)

            $btnStop     = _Button 'Stop' 298 $buttonY 80 30 $clrRed $null
            $btnStop.Name = 'btnStop'
            $btnStop.Tag = $pfx
            $btnStop.TabStop = $false
            $btnStop.Add_Click({
                $p = $this.Tag
                _RunServerOpInBackground -Prefix $p -Operation 'Stop'
                _ClearDashboardFocus -FallbackControl $script:_DashboardScrollPanel
            })
            _SetMainControlToolTip -Control $btnStop -Text ("Stop {0} using its configured shutdown method." -f $profile.GameName)
            $card.Controls.Add($btnStop)

            $btnRestart     = _Button 'Restart' 386 $buttonY 84 30 $restartButtonColor $null
            $btnRestart.Name = 'btnRestart'
            $btnRestart.Tag = $pfx
            $btnRestart.TabStop = $false
            $btnRestart.Add_Click({
                $p = $this.Tag
                _RunServerOpInBackground -Prefix $p -Operation 'Restart'
                _ClearDashboardFocus -FallbackControl $script:_DashboardScrollPanel
            })
            _SetMainControlToolTip -Control $btnRestart -Text ("Restart {0} using its configured save and stop rules." -f $profile.GameName)
            $card.Controls.Add($btnRestart)

            $btnCommands     = _Button 'Commands' 478 $buttonY 88 30 $commandsButtonColor $null
            $btnCommands.Name = 'btnCommands'
            $btnCommands.Tag = $pfx
            $btnCommands.TabStop = $false
            $btnCommands.Add_Click({
                $p = $this.Tag
                $prof = $script:SharedState.Profiles[$p]
                if ($prof) { _OpenCommandsWindow -Profile $prof -Prefix $p -WindowSharedState $SharedState }
                _ClearDashboardFocus -FallbackControl $script:_DashboardScrollPanel
            })
            _SetMainControlToolTip -Control $btnCommands -Text ("Open the command tools for {0}." -f $profile.GameName)
            $card.Controls.Add($btnCommands)

            $configRoots = _GetConfigRootsForProfile -Profile $profile
            $hasConfig   = ($configRoots.Count -gt 0)
            $btnConfig   = _Button 'Config' 574 $buttonY 80 30 $configButtonColor $null
            $btnConfig.Name = 'btnConfig'
            $btnConfig.Tag = $pfx
            $btnConfig.Enabled = $hasConfig
            $btnConfig.TabStop = $false
            $btnConfig.Add_Click({
                $p = $this.Tag
                $prof = $script:SharedState.Profiles[$p]
                if ($prof) { _OpenConfigEditor -Profile $prof }
                _ClearDashboardFocus -FallbackControl $script:_DashboardScrollPanel
            })
            $configTip = if ($hasConfig) {
                "Open the config files ECC found for $($profile.GameName)."
            } else {
                "ECC has not found a config path for $($profile.GameName) yet."
            }
            _SetMainControlToolTip -Control $btnConfig -Text $configTip
            $card.Controls.Add($btnConfig)

            if ($isHytaleProfile) {
                $btnHytaleTools = _Button 'Manager' ($statusBadgeX - 108) 17 98 30 $toolsButtonColor $null
                $btnHytaleTools.Anchor = 'Top,Right'
                $btnHytaleTools.Name = 'btnHytaleTools'
                $btnHytaleTools.Tag = $pfx
                $btnHytaleTools.TabStop = $false
                $btnHytaleTools.Add_Click({
                    $p = $this.Tag
                    $prof = $script:SharedState.Profiles[$p]
                    if ($prof) { _OpenHytaleManagerWindow -Profile $prof -Prefix $p }
                    _ClearDashboardFocus -FallbackControl $script:_DashboardScrollPanel
                })
                _SetMainControlToolTip -Control $btnHytaleTools -Text 'Open the Hytale manager for updater, downloader, and mod tools.'
                $card.Controls.Add($btnHytaleTools)
                try { $card.Controls.SetChildIndex($btnHytaleTools, 0) } catch { }
            }

            _ApplyDashboardCardToolTips -Card $card -ToolTipText $stateInfo.HealthToolTip

            # Clicking the card body (not a button) opens the profile editor
            $card.Add_Click({
                $pfxCapture = [string]$this.Tag
                $script:_SelectedProfilePrefix = $pfxCapture
                $prof = $script:SharedState.Profiles[$pfxCapture]
                if ($prof) { _BuildProfileEditor -Profile $prof }
                _BuildProfilesList
                _ClearDashboardFocus -FallbackControl $script:_DashboardScrollPanel
            })

            $scroll.Controls.Add($card)
            $script:_DashboardCardCache[$pfx.ToUpperInvariant()] = @{
                Card        = $card
                StatusLabel = $lblStatus
                Subtitle    = $lblSubtitle
                Source      = $lblSource
                Health      = $lblHealth
                Uptime      = $lblUptime
                Timers      = $lblTimers
                StartButton = $btnStart
                StopButton  = $btnStop
                RestartButton = $btnRestart
                CommandsButton = $btnCommands
                ConfigButton = $btnConfig
                HytaleButton = $btnHytaleTools
                AutoRestartCheckBox = $chk
                StatusPanel = $statusBadge
                HeaderPanel = $cardHeader
                AccentPanel = $cardAccent
                NameLabel = $lblName
            }
            $y += ($cardH + 10)
        }

        _RestoreScrollPosition -Control $scroll -Position $savedDashboardScroll
    }

    # =====================================================================
    # LIGHTWEIGHT STATUS UPDATER  (called every timer tick - no full rebuild)
    # =====================================================================
    function _UpdateDashboardStatus {
        $dashboardLoopPerf = [System.Diagnostics.Stopwatch]::StartNew()
        $ss = $script:SharedState
        if (-not $ss -or -not $ss.Profiles) { return }

        $resolvedCardPlayerCounts = @{}
        $runningSnapshot = @{}
        $cardCache = if ($script:_DashboardCardCache -is [hashtable]) { $script:_DashboardCardCache } else { @{} }
        if ($cardCache.Count -eq 0 -and $script:_DashboardScrollPanel -and @($script:_DashboardScrollPanel.Controls).Count -gt 0) {
            try {
                _BuildServerDashboard
                $cardCache = if ($script:_DashboardCardCache -is [hashtable]) { $script:_DashboardCardCache } else { @{} }
            } catch { }
        }

        try {
            $syncCmd = Get-Command -Name 'Sync-RunningServersFromProcesses' -ErrorAction SilentlyContinue
            if ($syncCmd) {
                & $syncCmd -SharedState $ss | Out-Null
            }
        } catch {
            _GuiModuleLog -Message ("Dashboard running-server sync failed: {0}" -f $_.Exception.Message) -Level WARN
        }

        try {
            if ($ss.RunningServers) {
                foreach ($runningKey in @($ss.RunningServers.Keys)) {
                    if ([string]::IsNullOrWhiteSpace([string]$runningKey)) { continue }
                    $runningSnapshot[[string]$runningKey.ToUpperInvariant()] = $ss.RunningServers[$runningKey]
                }
            }
        } catch { }

        $dashboardProfileCount = @($ss.Profiles.Keys).Count
        $dashboardCardCount = 0
        try {
            if ($cardCache.Count -gt 0) {
                $dashboardCardCount = @($cardCache.Keys).Count
            } elseif ($script:_DashboardScrollPanel) {
                $dashboardCardCount = @($script:_DashboardScrollPanel.Controls).Count
            }
        } catch { $dashboardCardCount = 0 }
        $dashboardCardScans = 0
        $dashboardMatchedCards = 0
        foreach ($pfx in @($ss.Profiles.Keys)) {
            $entry   = $null
            $running = $false

            $dashboardCardScans++
            $cacheKey = $pfx.ToUpperInvariant()
            if (-not $cardCache.ContainsKey($cacheKey)) { continue }
            $dashboardMatchedCards++

            if ($runningSnapshot.ContainsKey($cacheKey)) {
                $running = $true
                $runningEntry = $runningSnapshot[$cacheKey]
                $entry = @{
                    Pid       = $runningEntry.Pid
                    StartTime = $runningEntry.StartTime
                }
            }

            $cardEntry = $cardCache[$cacheKey]
            $card = $cardEntry.Card
            $statusLabel = $cardEntry.StatusLabel
            $subtitleLabel = $cardEntry.Subtitle
            $sourceLabel = $cardEntry.Source
            $healthLabel = $cardEntry.Health
            $uptimeLabel = $cardEntry.Uptime
            $timerLabel = $cardEntry.Timers
            $startButton = $cardEntry.StartButton
            $stopButton = $cardEntry.StopButton
            $statusPanel = $cardEntry.StatusPanel
            $accentPanel = $cardEntry.AccentPanel
            if (-not $card -or -not $statusLabel -or -not $uptimeLabel -or -not $timerLabel) { continue }

            $profile = $ss.Profiles[$pfx]
            $stateInfo = _GetDashboardStateInfo -Prefix $pfx -Profile $profile -Running:$running -Entry $entry -SharedState $ss
            $statusLabel.Text = $stateInfo.BadgeText
            $statusLabel.ForeColor = $stateInfo.StatusColor
            if ($statusPanel) { $statusPanel.BackColor = $stateInfo.StatusBg }
            if ($accentPanel) { $accentPanel.BackColor = $stateInfo.AccentColor }
            if ($subtitleLabel) { $subtitleLabel.Text = $stateInfo.Subtitle }
            if ($sourceLabel) { $sourceLabel.Text = $stateInfo.HealthMeta }
            if ($healthLabel) {
                $healthLabel.Text = $stateInfo.HealthText
                $healthLabel.ForeColor = $stateInfo.HealthColor
            }
            _ApplyDashboardCardToolTips -Card $card -ToolTipText $stateInfo.HealthToolTip
            if ($stateInfo.Subtitle -match '(?i)\b(\d+)\s+player\(s\)\s+online\b') {
                try {
                    $resolvedCardPlayerCounts[$cacheKey] = [Math]::Max(0, [int]$Matches[1])
                } catch { }
            }

            $startDisabled = $running -or @('starting','restarting','stopping','waiting_restart','waiting_first_player','blocked','online') -contains ([string]$stateInfo.StateCode)
            _SetDashboardStartButtonState -Button $startButton -Enabled:(-not $startDisabled) -StateCode ([string]$stateInfo.StateCode)

            $stopEnabled = $running -or @('starting','restarting','stopping','waiting_restart','waiting_first_player','idle_wait','idle_shutdown','online') -contains ([string]$stateInfo.StateCode)
            _SetDashboardStopButtonState -Button $stopButton -Enabled:$stopEnabled -StateCode ([string]$stateInfo.StateCode)

            if ($running -and $entry) {
                $up = [Math]::Round(((Get-Date) - $entry.StartTime).TotalMinutes, 1)
                if ($up -ge 60) {
                    $uh = [Math]::Floor($up / 60)
                    $um = [Math]::Floor($up % 60)
                    $uptimeLabel.Text = "PID $($entry.Pid) | Uptime: ${uh}h ${um}m"
                } else {
                    $uptimeLabel.Text = "PID $($entry.Pid) | Uptime: ${up} min"
                }

                $timerLabel.Text = _BuildTimerLine -Prefix $pfx -Profile $profile -Entry $entry -SharedState $ss
            } else {
                $uptimeLabel.Text = 'Server is not running'
                $timerLabel.Text = _BuildTimerLine -Prefix $pfx -Profile $profile -Entry $entry -SharedState $ss
            }
        }

        $script:_DashboardCardPlayerCounts = $resolvedCardPlayerCounts
        $dashboardLoopPerf.Stop()
        _TraceGuiPerformanceSample -Area 'DashboardCardLoop' `
            -ElapsedMs $dashboardLoopPerf.Elapsed.TotalMilliseconds `
            -WarnAtMs 120 `
            -DebugAtMs 40 `
            -Detail ('profiles={0};cards={1};cardScans={2};matched={3}' -f `
                $dashboardProfileCount, $dashboardCardCount, $dashboardCardScans, $dashboardMatchedCards)
    }

    function _NormalizeDashboardCardTopRightControls {
        if (-not ($script:_DashboardCardCache -is [hashtable])) { return }

        foreach ($cacheKey in @($script:_DashboardCardCache.Keys)) {
            $entry = $null
            try { $entry = $script:_DashboardCardCache[$cacheKey] } catch { $entry = $null }
            if (-not $entry) { continue }

            $card = $entry.Card
            if ($card -isnot [System.Windows.Forms.Control]) { continue }

            $statusPanel = $entry.StatusPanel
            $hytaleBtn = $entry.HytaleButton

            if ($statusPanel -is [System.Windows.Forms.Control]) {
                try {
                    if ($statusPanel.Parent -ne $card) {
                        try { if ($statusPanel.Parent) { $statusPanel.Parent.Controls.Remove($statusPanel) } } catch { }
                        $card.Controls.Add($statusPanel)
                    }
                    $card.Controls.SetChildIndex($statusPanel, 0)
                    $statusPanel.Visible = $true
                    $statusPanel.BringToFront()
                } catch { }
            }

            if ($hytaleBtn -is [System.Windows.Forms.Control]) {
                try {
                    if ($hytaleBtn.Parent -ne $card) {
                        try { if ($hytaleBtn.Parent) { $hytaleBtn.Parent.Controls.Remove($hytaleBtn) } } catch { }
                        $card.Controls.Add($hytaleBtn)
                    }
                    $card.Controls.SetChildIndex($hytaleBtn, 0)
                    $hytaleBtn.Visible = $true
                    $hytaleBtn.BringToFront()
                } catch { }
            }
        }
    }

    $layoutDashboardCards = {
        $dashScroll = $script:_DashboardScrollPanel
        if (-not $dashScroll) { return }

        foreach ($card in @($dashScroll.Controls)) {
            if ($card -isnot [System.Windows.Forms.Panel]) { continue }

            try {
                $card.Width = $dashScroll.ClientSize.Width - 20
            } catch { }

            $cacheEntry = $null
            try {
                if ($script:_DashboardCardCache -is [hashtable] -and $card.Tag) {
                    $cacheKey = ([string]$card.Tag).ToUpperInvariant()
                    if ($script:_DashboardCardCache.ContainsKey($cacheKey)) {
                        $cacheEntry = $script:_DashboardCardCache[$cacheKey]
                    }
                }
            } catch { $cacheEntry = $null }

            $statusPanel = if ($cacheEntry) { $cacheEntry.StatusPanel } else { $null }
            $badgeLabel  = if ($cacheEntry) { $cacheEntry.StatusLabel } else { $null }
            $headerPanel = if ($cacheEntry) { $cacheEntry.HeaderPanel } else { $null }
            $nameLabel   = ($card.Controls.Find('lblName', $true) | Select-Object -First 1)
            $subtitleLbl = if ($cacheEntry) { $cacheEntry.Subtitle } else { $null }
            $sourceLbl   = if ($cacheEntry) { $cacheEntry.Source } else { $null }
            $healthLbl   = if ($cacheEntry) { $cacheEntry.Health } else { $null }
            $uptimeLabel = if ($cacheEntry) { $cacheEntry.Uptime } else { $null }
            $timerLabel  = if ($cacheEntry) { $cacheEntry.Timers } else { $null }
            $startBtn    = if ($cacheEntry) { $cacheEntry.StartButton } else { $null }
            $stopBtn     = if ($cacheEntry) { $cacheEntry.StopButton } else { $null }
            $restartBtn  = if ($cacheEntry) { $cacheEntry.RestartButton } else { $null }
            $commandsBtn = if ($cacheEntry) { $cacheEntry.CommandsButton } else { $null }
            $configBtn   = if ($cacheEntry) { $cacheEntry.ConfigButton } else { $null }
            $hytaleBtn   = if ($cacheEntry) { $cacheEntry.HytaleButton } else { $null }
            $autoRestartChk = if ($cacheEntry) { $cacheEntry.AutoRestartCheckBox } else { $null }
            $nameLabel   = if ($cacheEntry) { $cacheEntry.NameLabel } else { $nameLabel }

            $cardLeft = 18
            $headerTop = 10
            $statusPanelWidth = if ($statusPanel) { [Math]::Max(96, $statusPanel.Width) } else { 116 }
            $statusPanelHeight = 44
            $statusPanelMargin = 16
            $buttonGap = 8
            $buttonY = 136
            $checkboxY = 143

            if ($headerPanel) {
                $headerPanel.Location = [System.Drawing.Point]::new($cardLeft, $headerTop)
                $headerPanel.Size = [System.Drawing.Size]::new([Math]::Max(180, $card.ClientSize.Width - ($cardLeft * 2)), 44)
            }

            $statusPanelX = [Math]::Max(96, $card.ClientSize.Width - $statusPanelWidth - $statusPanelMargin)
            if ($statusPanel) {
                $statusPanel.Size = [System.Drawing.Size]::new($statusPanelWidth, $statusPanelHeight)
                $statusPanel.Location = [System.Drawing.Point]::new($statusPanelX, $headerTop)
                $statusPanel.Visible = $true
                try { $card.Controls.SetChildIndex($statusPanel, 0) } catch { }
                $statusPanel.BringToFront()
            }
            $isHytaleCard = ($null -ne $hytaleBtn)
            $nameRight = if ($nameLabel) { $nameLabel.Left + $nameLabel.Width } else { 0 }
            if ($hytaleBtn) {
                $hytaleBtnX = [Math]::Max($nameRight + 12, $statusPanelX - $hytaleBtn.Width - 10)
                $hytaleBtn.Location = [System.Drawing.Point]::new($hytaleBtnX, 17)
                $hytaleBtn.Visible = $true
                try { $card.Controls.SetChildIndex($hytaleBtn, 0) } catch { }
                $hytaleBtn.BringToFront()
            }

            $headerBlockLeft = if ($isHytaleCard -and $hytaleBtn) {
                $hytaleBtn.Left
            } else {
                $statusPanelX
            }
            $headerWidth = [Math]::Max(80, $headerBlockLeft - $headerPanel.Left - 12)
            if ($nameLabel) {
                try {
                    $namePreferredWidth = [Math]::Max(80, [System.Windows.Forms.TextRenderer]::MeasureText($nameLabel.Text, $nameLabel.Font).Width + 6)
                } catch {
                    $namePreferredWidth = $headerWidth
                }
                $nameLabel.Width = [Math]::Min($headerWidth, $namePreferredWidth)
                $nameLabel.BringToFront()
            }
            $textRight = [Math]::Max(180, $headerBlockLeft - $cardLeft - 12)
            if ($subtitleLbl) {
                $subtitleLbl.Location = [System.Drawing.Point]::new($cardLeft, 58)
                $subtitleLbl.Width = $textRight
            }
            if ($sourceLbl) {
                $sourceLbl.Location = [System.Drawing.Point]::new($cardLeft, 78)
                $sourceLbl.Width = $textRight
            }
            if ($healthLbl) {
                $healthLbl.Location = [System.Drawing.Point]::new(0, 23)
                $healthLbl.Size = [System.Drawing.Size]::new($statusPanelWidth, 16)
                $healthLbl.Visible = $true
                $healthLbl.BringToFront()
            }
            if ($badgeLabel) {
                $badgeLabel.Location = [System.Drawing.Point]::new(0, 3)
                $badgeLabel.Size = [System.Drawing.Size]::new($statusPanelWidth, 18)
                $badgeLabel.Visible = $true
                $badgeLabel.BringToFront()
            }
            if ($uptimeLabel) {
                $uptimeLabel.Location = [System.Drawing.Point]::new($cardLeft, 98)
                $uptimeLabel.Width = [Math]::Max(180, $textRight - 80)
            }
            if ($timerLabel) {
                $timerLabel.Location = [System.Drawing.Point]::new($cardLeft, 120)
                $timerLabel.Width = [Math]::Max(160, $textRight - 100)
            }

            $buttonRight = [Math]::Max(220, $card.ClientSize.Width - 16)

            if ($configBtn) {
                $buttonRight -= $configBtn.Width
                $configBtn.Location = [System.Drawing.Point]::new($buttonRight, $buttonY)
                $configBtn.BringToFront()
                $buttonRight -= $buttonGap
            }
            if ($commandsBtn) {
                $buttonRight -= $commandsBtn.Width
                $commandsBtn.Location = [System.Drawing.Point]::new($buttonRight, $buttonY)
                $commandsBtn.BringToFront()
                $buttonRight -= $buttonGap
            }
            if ($restartBtn) {
                $buttonRight -= $restartBtn.Width
                $restartBtn.Location = [System.Drawing.Point]::new($buttonRight, $buttonY)
                $restartBtn.BringToFront()
                $buttonRight -= $buttonGap
            }
            if ($stopBtn) {
                $buttonRight -= $stopBtn.Width
                $stopBtn.Location = [System.Drawing.Point]::new($buttonRight, $buttonY)
                $stopBtn.BringToFront()
                $buttonRight -= $buttonGap
            }
            if ($startBtn) {
                $buttonRight -= $startBtn.Width
                $startBtn.Location = [System.Drawing.Point]::new($buttonRight, $buttonY)
                $startBtn.BringToFront()
            }
            if ($autoRestartChk) {
                $autoRestartChk.Location = [System.Drawing.Point]::new($cardLeft, $checkboxY)
                $autoRestartChk.BringToFront()
            }
        }
        _NormalizeDashboardCardTopRightControls
    }.GetNewClosure()

    # =====================================================================
    # INITIAL BUILDS
    # =====================================================================
    _BuildProfilesList
    if ($script:_SelectedProfilePrefix -and $script:SharedState -and $script:SharedState.Profiles -and
        $script:SharedState.Profiles.ContainsKey($script:_SelectedProfilePrefix)) {
        _BuildProfileEditor -Profile $script:SharedState.Profiles[$script:_SelectedProfilePrefix]
    } else {
        _BuildProfileEditor $null
    }
    _BuildServerDashboard
    & $layoutDashboardCards
    _UpdateDashboardStatus

    # =====================================================================
    # LAYOUT REFLOW  (called on form resize)
    # =====================================================================
    $reflowLayoutCore = {
        $reflowPerf = [System.Diagnostics.Stopwatch]::StartNew()
        $cw       = $form.ClientSize.Width
        $ch       = $form.ClientSize.Height
        $sbHeight = $statusBar.Height
        $contentTop = $windowMargin + $topBarHeight + 10

        try {
            if ((_IsGuiDebugEnabled) -and $script:SharedState -and $script:SharedState.LogQueue) {
                $now = [DateTime]::UtcNow
                $key = 'RESIZE_TRACE::ReflowEntry'
                if (-not $script:SharedState.ContainsKey($key) -or (($now - [DateTime]$script:SharedState[$key]).TotalSeconds -ge 2)) {
                    $script:SharedState[$key] = $now
                    $script:SharedState.LogQueue.Enqueue(
                        "[{0}][INFO][GUI] RESIZE ReflowEntry :: client={1}x{2};statusHeight={3}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $cw, $ch, $sbHeight
                    )
                }
            }
        } catch { }

        if ($statusBar) {
            $statusBar.Location = [System.Drawing.Point]::new(0, [Math]::Max(0, $ch - $sbHeight))
            $statusBar.Width = $cw
        }

        & $normalizeResizeChrome

        if ($script:_WindowShellPanel) {
            $script:_WindowShellPanel.Location = [System.Drawing.Point]::new($shellInset, $shellInset)
            $script:_WindowShellPanel.Size = [System.Drawing.Size]::new(
                $cw - ($shellInset * 2),
                $ch - $sbHeight - ($shellInset * 2)
            )
        }
        if ($script:_WindowEdgeTop) {
            $script:_WindowEdgeTop.Location = [System.Drawing.Point]::new(0, 0)
            $script:_WindowEdgeTop.Size = [System.Drawing.Size]::new($cw, 2)
        }
        if ($script:_WindowEdgeLeft) {
            $script:_WindowEdgeLeft.Location = [System.Drawing.Point]::new(0, 0)
            $script:_WindowEdgeLeft.Size = [System.Drawing.Size]::new(2, $ch)
        }
        if ($script:_WindowEdgeRight) {
            $script:_WindowEdgeRight.Location = [System.Drawing.Point]::new([Math]::Max(0, $cw - 2), 0)
            $script:_WindowEdgeRight.Size = [System.Drawing.Size]::new(2, $ch)
        }
        & $layoutResizeChromeElements $cw $ch $sbHeight
        try {
            $resizeDetail = "client={0}x{1};statusTop={2};bottomGrip={3},{4},{5},{6};rightGrip={7},{8},{9},{10};bottomRight={11},{12},{13},{14}" -f `
                $cw, `
                $ch, `
                $(if ($statusBar) { $statusBar.Top } else { -1 }), `
                $(if ($script:_ResizeBottomGrip) { $script:_ResizeBottomGrip.Left } else { -1 }), `
                $(if ($script:_ResizeBottomGrip) { $script:_ResizeBottomGrip.Top } else { -1 }), `
                $(if ($script:_ResizeBottomGrip) { $script:_ResizeBottomGrip.Width } else { -1 }), `
                $(if ($script:_ResizeBottomGrip) { $script:_ResizeBottomGrip.Height } else { -1 }), `
                $(if ($script:_ResizeRightGrip) { $script:_ResizeRightGrip.Left } else { -1 }), `
                $(if ($script:_ResizeRightGrip) { $script:_ResizeRightGrip.Top } else { -1 }), `
                $(if ($script:_ResizeRightGrip) { $script:_ResizeRightGrip.Width } else { -1 }), `
                $(if ($script:_ResizeRightGrip) { $script:_ResizeRightGrip.Height } else { -1 }), `
                $(if ($script:_ResizeBottomRightGrip) { $script:_ResizeBottomRightGrip.Left } else { -1 }), `
                $(if ($script:_ResizeBottomRightGrip) { $script:_ResizeBottomRightGrip.Top } else { -1 }), `
                $(if ($script:_ResizeBottomRightGrip) { $script:_ResizeBottomRightGrip.Width } else { -1 }), `
                $(if ($script:_ResizeBottomRightGrip) { $script:_ResizeBottomRightGrip.Height } else { -1 })
            & $writeResizeTrace 'GripLayout' $resizeDetail
        } catch { }

        $topBar.Location = [System.Drawing.Point]::new($windowMargin, $windowMargin)
        $topBar.Width  = $cw - ($windowMargin * 2)
        $topBar.Height = $topBarHeight

        # Reposition top-bar right-side control groups on resize
        $availableActionWidth = [Math]::Max($actionPanelMinWidth, [Math]::Min($actionPanelDesiredWidth, $topBar.Width - $metricsPanel.Left - 440))
        $actionPanel.Location       = [System.Drawing.Point]::new($topBar.Width - $chromeClusterWidth - $actionPanelToChromeGap - $availableActionWidth, 14)
        $actionPanel.Width          = $availableActionWidth
        $metricsPanelRightGap = 16
        $metricsPanel.Width = [Math]::Max(420, $actionPanel.Left - $metricsPanel.Left - $metricsPanelRightGap)
        & $layoutTopMetrics
        & $layoutTopActions
        $windowChromeButtonYCurrent = [Math]::Max(0, [int][Math]::Floor((($topBar.Height - 1) - $windowChromeButtonHeight) / 2))
        $btnWinMin.Location         = [System.Drawing.Point]::new($topBar.Width - (($windowChromeButtonWidth * 3) + ($windowChromeGap * 2)), $windowChromeButtonYCurrent)
        $btnWinMax.Location         = [System.Drawing.Point]::new($topBar.Width - (($windowChromeButtonWidth * 2) + $windowChromeGap), $windowChromeButtonYCurrent)
        $btnWinClose.Location       = [System.Drawing.Point]::new($topBar.Width - $windowChromeButtonWidth,  $windowChromeButtonYCurrent)

        $leftW   = if ($script:_LeftCollapsed)  { $collapsedSize } else { $leftWidth }
        $rightW  = if ($script:_RightCollapsed) { $collapsedSize } else { $rightWidth }
        $bottomH = if ($script:_BottomCollapsed){ $bottomHeaderHeight } else { $bottomLogsHeight }

        $leftMinExpanded = 220
        $rightMinExpanded = 320
        $availableShellWidth = [Math]::Max(320, $cw - ($windowMargin * 2) - ($sideGap * 2))
        $centerMinExpanded = [Math]::Max(320, [Math]::Min(560, $availableShellWidth - $leftMinExpanded - $rightMinExpanded))

        if (-not $script:_LeftCollapsed -and -not $script:_RightCollapsed) {
            $requiredShellWidth = $leftW + $rightW + $centerMinExpanded
            if ($requiredShellWidth -gt $availableShellWidth) {
                $deficit = $requiredShellWidth - $availableShellWidth

                $rightShrinkCapacity = [Math]::Max(0, $rightW - $rightMinExpanded)
                if ($rightShrinkCapacity -gt 0) {
                    $rightShrink = [Math]::Min($deficit, $rightShrinkCapacity)
                    $rightW -= $rightShrink
                    $deficit -= $rightShrink
                }

                $leftShrinkCapacity = [Math]::Max(0, $leftW - $leftMinExpanded)
                if ($deficit -gt 0 -and $leftShrinkCapacity -gt 0) {
                    $leftShrink = [Math]::Min($deficit, $leftShrinkCapacity)
                    $leftW -= $leftShrink
                    $deficit -= $leftShrink
                }
            }
        } elseif (-not $script:_LeftCollapsed) {
            $leftW = [Math]::Min($leftW, [Math]::Max($leftMinExpanded, $availableShellWidth - $centerMinExpanded - $collapsedSize))
        } elseif (-not $script:_RightCollapsed) {
            $rightW = [Math]::Min($rightW, [Math]::Max($rightMinExpanded, $availableShellWidth - $centerMinExpanded - $collapsedSize))
        }

        $contentBottomInset = $windowMargin
        $contentRightInset = [Math]::Max($windowMargin, $rightResizeGutter)

        $bottomY = $ch - $sbHeight - $contentBottomInset - $bottomH
        if ($bottomY -lt ($contentTop + 200)) { $bottomY = $contentTop + 200 }
        try {
            if (_IsGuiDebugEnabled) {
                _QueueStatusMessage(("PANELREFLOW state left={0};right={1};bottom={2};targetLeftW={3};targetRightW={4};targetBottomH={5};client={6}x{7}" -f `
                    $(if ($script:_LeftCollapsed) { '1' } else { '0' }), `
                    $(if ($script:_RightCollapsed) { '1' } else { '0' }), `
                    $(if ($script:_BottomCollapsed) { '1' } else { '0' }), `
                    $leftW, $rightW, $bottomH, $cw, $ch))
            }
        } catch { }

        $bottomContainer.Location = [System.Drawing.Point]::new($windowMargin, $bottomY)
        $bottomContainer.Size     = [System.Drawing.Size]::new([Math]::Max(260, $cw - $windowMargin - $contentRightInset), $bottomH)

        $leftContainer.Location  = [System.Drawing.Point]::new($windowMargin, $contentTop)
        $leftContainer.Size      = [System.Drawing.Size]::new($leftW, $bottomY - $contentTop)

        $rightX = $cw - $contentRightInset - $rightW
        $rightContainer.Location = [System.Drawing.Point]::new($rightX, $contentTop)
        $rightContainer.Size     = [System.Drawing.Size]::new($rightW, $bottomY - $contentTop)

        $centerX = $windowMargin + $leftW + $sideGap
        $centerW = [Math]::Max(220, $rightX - $sideGap - $centerX)
        $centerCol.Location = [System.Drawing.Point]::new($centerX, $contentTop)
        $centerCol.Size     = [System.Drawing.Size]::new($centerW, $bottomY - $contentTop)

        # Left header/body sizing
        if ($script:_LeftCollapsed) {
            $leftHeader.Location = [System.Drawing.Point]::new(0, 0)
            $leftHeader.Size     = [System.Drawing.Size]::new($leftW, $leftContainer.Height)
            $leftBody.Visible    = $false
            if ($leftHeaderToggle) { $leftHeaderToggle.Text = '+' }
            $leftHeaderLabel.Text = _VerticalText 'Profiles'
            $leftHeaderLabel.Location = [System.Drawing.Point]::new(3, 6)
            $leftHeaderLabel.Size = [System.Drawing.Size]::new([Math]::Max(16, $leftW - 6), $leftHeader.Height - 6)
            $leftHeaderLabel.TextAlign = 'TopCenter'
        } else {
            $leftHeader.Location = [System.Drawing.Point]::new(0, 0)
            $leftHeader.Size     = [System.Drawing.Size]::new($leftW, $headerHeight)
            $leftBody.Location   = [System.Drawing.Point]::new(0, $headerHeight)
            $leftBody.Size       = [System.Drawing.Size]::new($leftW, $leftContainer.Height - $headerHeight)
            $leftBody.Visible    = $true
            if ($leftHeaderToggle) { $leftHeaderToggle.Text = '-' }
            $leftHeaderLabel.Text = 'Game Profiles'
            $leftHeaderLabel.Location = [System.Drawing.Point]::new(10, 8)
            $leftHeaderLabel.Size = [System.Drawing.Size]::new($leftW - 20, $headerHeight - 10)
            $leftHeaderLabel.TextAlign = 'MiddleLeft'
        }
        try {
            if ($leftHeaderLabel) { $leftHeaderLabel.BringToFront() }
            if ($leftHeaderToggle) { $leftHeaderToggle.BringToFront() }
        } catch { }

        # Right header/body sizing
        if ($script:_RightCollapsed) {
            $rightHeader.Location = [System.Drawing.Point]::new(0, 0)
            $rightHeader.Size     = [System.Drawing.Size]::new($rightW, $rightContainer.Height)
            $rightBody.Visible    = $false
            if ($rightHeaderToggle) { $rightHeaderToggle.Text = '+' }
            $rightHeaderLabel.Text = _VerticalText 'Editor'
            $rightHeaderLabel.Location = [System.Drawing.Point]::new(3, 6)
            $rightHeaderLabel.Size = [System.Drawing.Size]::new([Math]::Max(16, $rightW - 6), $rightHeader.Height - 6)
            $rightHeaderLabel.TextAlign = 'TopCenter'
        } else {
            $rightHeader.Location = [System.Drawing.Point]::new(0, 0)
            $rightHeader.Size     = [System.Drawing.Size]::new($rightW, $headerHeight)
            $rightBody.Location   = [System.Drawing.Point]::new(0, $headerHeight)
            $rightBody.Size       = [System.Drawing.Size]::new($rightW, $rightContainer.Height - $headerHeight)
            $rightBody.Visible    = $true
            if ($rightHeaderToggle) { $rightHeaderToggle.Text = '-' }
            $rightHeaderLabel.Text = 'Profile Editor'
            $rightHeaderLabel.Location = [System.Drawing.Point]::new(10, 8)
            $rightHeaderLabel.Size = [System.Drawing.Size]::new($rightW - 20, $headerHeight - 10)
            $rightHeaderLabel.TextAlign = 'MiddleLeft'
        }
        try {
            if ($rightHeaderLabel) { $rightHeaderLabel.BringToFront() }
            if ($rightHeaderToggle) { $rightHeaderToggle.BringToFront() }
        } catch { }

        # Bottom header/body - Dock handles sizing, just toggle visibility for collapse
        if ($script:_BottomCollapsed) {
            $bottomHeader.Height = $bottomContainer.Height   # expand header to fill when collapsed
            $bottomPanel.Visible = $false
            if ($bottomHeaderToggle) { $bottomHeaderToggle.Text = '+' }
        } else {
            $bottomHeader.Height = $bottomHeaderHeight
            $bottomPanel.Visible = $true
            if ($bottomHeaderToggle) { $bottomHeaderToggle.Text = '-' }
        }
        try {
            if ($bottomHeaderLabel) { $bottomHeaderLabel.BringToFront() }
            if ($bottomHeaderToggle) { $bottomHeaderToggle.BringToFront() }
        } catch { }

        # Bottom panel internals (only resize controls that actually exist)
        if ($script:_DiscordFooter -and $tbSend -and $btnSend -and $btnClearDisc) {
            $tbSend.Width = [Math]::Max(100, $script:_DiscordFooter.Width - 220)
            $btnSend.Location = [System.Drawing.Point]::new([Math]::Max(120, $script:_DiscordFooter.Width - 205), 3)
            $btnClearDisc.Location = [System.Drawing.Point]::new([Math]::Max(200, $script:_DiscordFooter.Width - 110), 3)
        }

        # Log TabControl is Dock=Fill inside bottomPanel - no manual sizing needed

        # Profiles list panel
        if ($script:_ProfilesListPanel) {
            $profilesLayout = _GetProfilesPaneLayout -PanelWidth $leftBody.Width -PanelHeight $leftBody.Height
            $script:_ProfilesListPanel.Location = [System.Drawing.Point]::new($profilesLayout.SideMargin, $profilesLayout.TopMargin)
            $script:_ProfilesListPanel.Size = [System.Drawing.Size]::new($profilesLayout.ListWidth, $profilesLayout.ListHeight)
        }
        if ($script:_ProfilesFooterPanel) {
            $profilesLayout = if ($profilesLayout) { $profilesLayout } else { _GetProfilesPaneLayout -PanelWidth $leftBody.Width -PanelHeight $leftBody.Height }
            $script:_ProfilesFooterPanel.Location = [System.Drawing.Point]::new(8, $profilesLayout.FooterTop)
            $script:_ProfilesFooterPanel.Size = [System.Drawing.Size]::new($profilesLayout.FooterWidth, $profilesLayout.FooterHeight)
        }

        # Profile editor scroll panel
        if ($script:_ProfileEditorPanel) {
            foreach ($ctrl in $script:_ProfileEditorPanel.Controls) {
                if ($ctrl -is [System.Windows.Forms.Panel]) {
                    $ctrl.Size = [System.Drawing.Size]::new($script:_ProfileEditorPanel.Width, $script:_ProfileEditorPanel.Height)
                    if ($script:_ProfileFields) {
                        $twNew = [Math]::Max(260, $ctrl.Width - (172 + 104))
                        foreach ($key in $script:_ProfileFields.Keys) {
                            $field = $script:_ProfileFields[$key]
                            if ($field -is [System.Windows.Forms.TextBox]) { $field.Width = $twNew }
                            if ($field -is [System.Collections.IDictionary] -and $field.Control -is [System.Windows.Forms.TextBox]) {
                                $field.Control.Width = $twNew
                            }
                            if ($field -is [System.Collections.IDictionary] -and $field.Kind -eq 'launchargs') {
                                if ($field.Controls) {
                                    foreach ($ck in $field.Controls.Keys) {
                                        $c = $field.Controls[$ck]
                                        if ($c -and $c.Text -is [System.Windows.Forms.TextBox]) {
                                            if ($c.Browse -is [System.Windows.Forms.Button]) {
                                                $btnW = $c.Browse.Width
                                                $c.Text.Width = [Math]::Max(120, $twNew - ($btnW + 6))
                                                $c.Browse.Location = [System.Drawing.Point]::new(
                                                    $c.Text.Location.X + $c.Text.Width + 6,
                                                    $c.Browse.Location.Y
                                                )
                                            } else {
                                                $c.Text.Width = $twNew
                                            }
                                        }
                                    }
                                }
                                if ($field.CustomBox -is [System.Windows.Forms.TextBox]) {
                                    $field.CustomBox.Width = $twNew
                                }
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

        & $layoutDashboardCards

        try {
            if ($script:_WindowShellPanel) { $script:_WindowShellPanel.SendToBack() }
            if ($topBar) { $topBar.BringToFront() }
            if ($leftContainer) { $leftContainer.BringToFront() }
            if ($centerCol) { $centerCol.BringToFront() }
            if ($rightContainer) { $rightContainer.BringToFront() }
            if ($bottomContainer) { $bottomContainer.BringToFront() }
            if ($statusBar) { $statusBar.BringToFront() }
        if ($script:_WindowEdgeTop) { $script:_WindowEdgeTop.BringToFront() }
        if ($script:_WindowEdgeLeft) { $script:_WindowEdgeLeft.BringToFront() }
        if ($script:_WindowEdgeRight) { $script:_WindowEdgeRight.BringToFront() }
        if ($script:_WindowEdgeBottom) { $script:_WindowEdgeBottom.BringToFront() }
        if ($script:_ResizeBottomGrip) { $script:_ResizeBottomGrip.BringToFront() }
        if ($script:_ResizeLeftGrip) { $script:_ResizeLeftGrip.BringToFront() }
        if ($script:_ResizeRightGrip) { $script:_ResizeRightGrip.BringToFront() }
        if ($script:_ResizeBottomLeftGrip) { $script:_ResizeBottomLeftGrip.BringToFront() }
        if ($script:_ResizeBottomRightGrip) { $script:_ResizeBottomRightGrip.BringToFront() }
        } catch { }

        try {
            $reflowPerf.Stop()
            $dashCardCount = 0
            try {
                if ($script:_DashboardScrollPanel) {
                    $dashCardCount = @($script:_DashboardScrollPanel.Controls).Count
                }
            } catch { $dashCardCount = 0 }
            _TraceGuiPerformanceSample -Area 'LayoutReflow' `
                -ElapsedMs $reflowPerf.Elapsed.TotalMilliseconds `
                -WarnAtMs 140 `
                -DebugAtMs 50 `
                -Detail ('client={0}x{1};cards={2};leftCollapsed={3};rightCollapsed={4};bottomCollapsed={5}' -f `
                    $cw, $ch, $dashCardCount, `
                    $(if ($script:_LeftCollapsed) { '1' } else { '0' }), `
                    $(if ($script:_RightCollapsed) { '1' } else { '0' }), `
                    $(if ($script:_BottomCollapsed) { '1' } else { '0' }))
        } catch { }
    }.GetNewClosure()

    $script:_ReflowLayoutHandler = $reflowLayoutCore

    $applyPaneCollapseVisuals = {
        $cw = $form.ClientSize.Width
        $ch = $form.ClientSize.Height
        $sbHeight = $statusBar.Height
        $contentTop = $windowMargin + $topBarHeight + 10

        $leftW   = if ($script:_LeftCollapsed)  { $collapsedSize } else { $leftWidth }
        $rightW  = if ($script:_RightCollapsed) { $collapsedSize } else { $rightWidth }
        $bottomH = if ($script:_BottomCollapsed){ $bottomHeaderHeight } else { $bottomLogsHeight }

        $contentBottomInset = $windowMargin
        $contentRightInset = [Math]::Max($windowMargin, $rightResizeGutter)

        $bottomY = $ch - $sbHeight - $contentBottomInset - $bottomH
        if ($bottomY -lt ($contentTop + 200)) { $bottomY = $contentTop + 200 }

        $rightX = $cw - $contentRightInset - $rightW
        $centerX = $windowMargin + $leftW + $sideGap
        $centerW = [Math]::Max(220, $rightX - $sideGap - $centerX)
        try {
            $form.SuspendLayout()
            $leftContainer.SuspendLayout()
            $rightContainer.SuspendLayout()
            $centerCol.SuspendLayout()
            $bottomContainer.SuspendLayout()

            $bottomContainer.SetBounds($windowMargin, $bottomY, [Math]::Max(260, $cw - $windowMargin - $contentRightInset), $bottomH)
            $leftContainer.SetBounds($windowMargin, $contentTop, $leftW, ($bottomY - $contentTop))
            $rightContainer.SetBounds($rightX, $contentTop, $rightW, ($bottomY - $contentTop))
            $centerCol.SetBounds($centerX, $contentTop, $centerW, ($bottomY - $contentTop))

            if ($script:_LeftCollapsed) {
                $leftHeader.SetBounds(0, 0, $leftContainer.Width, $leftContainer.Height)
                $leftBody.Visible = $false
                if ($leftHeaderToggle) { $leftHeaderToggle.Text = '+' }
            } else {
                $leftHeader.SetBounds(0, 0, $leftContainer.Width, $headerHeight)
                $leftBody.SetBounds(0, $headerHeight, $leftContainer.Width, ($leftContainer.Height - $headerHeight))
                $leftBody.Visible = $true
                if ($leftHeaderToggle) { $leftHeaderToggle.Text = '-' }
            }

            if ($script:_RightCollapsed) {
                $rightHeader.SetBounds(0, 0, $rightContainer.Width, $rightContainer.Height)
                $rightBody.Visible = $false
                if ($rightHeaderToggle) { $rightHeaderToggle.Text = '+' }
            } else {
                $rightHeader.SetBounds(0, 0, $rightContainer.Width, $headerHeight)
                $rightBody.SetBounds(0, $headerHeight, $rightContainer.Width, ($rightContainer.Height - $headerHeight))
                $rightBody.Visible = $true
                if ($rightHeaderToggle) { $rightHeaderToggle.Text = '-' }
            }

            if ($script:_BottomCollapsed) {
                $bottomHeader.Height = $bottomContainer.Height
                $bottomPanel.Visible = $false
                if ($bottomHeaderToggle) { $bottomHeaderToggle.Text = '+' }
            } else {
                $bottomHeader.Height = $bottomHeaderHeight
                $bottomPanel.Visible = $true
                if ($bottomHeaderToggle) { $bottomHeaderToggle.Text = '-' }
            }

            $bottomContainer.ResumeLayout($true)
            $centerCol.ResumeLayout($true)
            $rightContainer.ResumeLayout($true)
            $leftContainer.ResumeLayout($true)
            $form.ResumeLayout($true)
            $form.PerformLayout()
            $leftContainer.Refresh()
            $rightContainer.Refresh()
            $centerCol.Refresh()
            $bottomContainer.Refresh()
            $form.Refresh()
        } catch {
            try { $form.ResumeLayout($true) } catch { }
            try { _WriteGuiTrace ("Pane visual apply failed: {0}" -f $_.Exception.Message) 'ERROR' } catch { }
        }
    }.GetNewClosure()
    $script:_ApplyPaneCollapseVisuals = $applyPaneCollapseVisuals

    function _InvokePaneToggleDirect {
        param(
            [string]$PaneKey,
            [string]$ControlName
        )

        if ([string]::IsNullOrWhiteSpace($PaneKey)) { return }
        $nowTick = [DateTime]::UtcNow.Ticks
        try {
            if ($script:_PaneToggleDebounce.ContainsKey($PaneKey)) {
                $elapsedMs = [Math]::Abs(([TimeSpan]::FromTicks($nowTick - [long]$script:_PaneToggleDebounce[$PaneKey])).TotalMilliseconds)
                if ($elapsedMs -lt 250) {
                    try { _QueueStatusMessage ("PANELTOGGLE suppress pane={0};elapsedMs={1:n1}" -f $PaneKey, $elapsedMs) } catch { }
                    return
                }
            }
            $script:_PaneToggleDebounce[$PaneKey] = $nowTick
        } catch { }

        switch ($PaneKey) {
            'Left'   { $script:_LeftCollapsed = -not $script:_LeftCollapsed }
            'Right'  { $script:_RightCollapsed = -not $script:_RightCollapsed }
            'Bottom' { $script:_BottomCollapsed = -not $script:_BottomCollapsed }
            default  { return }
        }

        try {
            if (_IsGuiDebugEnabled) {
                _QueueStatusMessage ("PANELTOGGLE click pane={0};control={1};left={2};right={3};bottom={4}" -f `
                    $PaneKey, $ControlName, `
                    $(if ($script:_LeftCollapsed) { '1' } else { '0' }), `
                    $(if ($script:_RightCollapsed) { '1' } else { '0' }), `
                    $(if ($script:_BottomCollapsed) { '1' } else { '0' }))
            }
        } catch { }

        try {
            $cw = $form.ClientSize.Width
            $ch = $form.ClientSize.Height
            $sbHeight = $statusBar.Height
            $contentTop = $windowMargin + $topBarHeight + 10
            $leftW   = if ($script:_LeftCollapsed)  { $collapsedSize } else { $leftWidth }
            $rightW  = if ($script:_RightCollapsed) { $collapsedSize } else { $rightWidth }
            $bottomH = if ($script:_BottomCollapsed){ $bottomHeaderHeight } else { $bottomLogsHeight }
            $contentBottomInset = $windowMargin
            $contentRightInset = [Math]::Max($windowMargin, $rightResizeGutter)
            $bottomY = $ch - $sbHeight - $contentBottomInset - $bottomH
            if ($bottomY -lt ($contentTop + 200)) { $bottomY = $contentTop + 200 }
            $rightX = $cw - $contentRightInset - $rightW
            $centerX = $windowMargin + $leftW + $sideGap
            $centerW = [Math]::Max(220, $rightX - $sideGap - $centerX)

            $bottomContainer.SetBounds($windowMargin, $bottomY, [Math]::Max(260, $cw - $windowMargin - $contentRightInset), $bottomH)
            $leftContainer.SetBounds($windowMargin, $contentTop, $leftW, ($bottomY - $contentTop))
            $rightContainer.SetBounds($rightX, $contentTop, $rightW, ($bottomY - $contentTop))
            $centerCol.SetBounds($centerX, $contentTop, $centerW, ($bottomY - $contentTop))

            if ($script:_LeftCollapsed) {
                $leftHeader.SetBounds(0, 0, $leftContainer.Width, $leftContainer.Height)
                $leftBody.Visible = $false
                if ($leftHeaderToggle) { $leftHeaderToggle.Text = '+' }
                if ($leftHeaderLabel) {
                    $leftHeaderLabel.Text = _VerticalText 'Profiles'
                    $leftHeaderLabel.Location = [System.Drawing.Point]::new(3, 28)
                    $leftHeaderLabel.Size = [System.Drawing.Size]::new([Math]::Max(16, $leftContainer.Width - 6), [Math]::Max(40, $leftHeader.Height - 34))
                    $leftHeaderLabel.TextAlign = 'TopCenter'
                }
                if ($leftHeaderToggle) {
                    $leftHeaderToggle.Location = [System.Drawing.Point]::new(1, 4)
                    $leftHeaderToggle.Size = [System.Drawing.Size]::new([Math]::Max(24, $leftContainer.Width - 2), 22)
                }
            } else {
                $leftHeader.SetBounds(0, 0, $leftContainer.Width, $headerHeight)
                $leftBody.SetBounds(0, $headerHeight, $leftContainer.Width, ($leftContainer.Height - $headerHeight))
                $leftBody.Visible = $true
                if ($leftHeaderToggle) { $leftHeaderToggle.Text = '-' }
                if ($leftHeaderLabel) {
                    $leftHeaderLabel.Text = 'Game Profiles'
                    $leftHeaderLabel.Location = [System.Drawing.Point]::new(10, 8)
                    $leftHeaderLabel.Size = [System.Drawing.Size]::new([Math]::Max(80, $leftContainer.Width - 50), $headerHeight - 10)
                    $leftHeaderLabel.TextAlign = 'MiddleLeft'
                }
                if ($leftHeaderToggle) {
                    $leftHeaderToggle.Location = [System.Drawing.Point]::new([Math]::Max(6, $leftContainer.Width - 34), 6)
                    $leftHeaderToggle.Size = [System.Drawing.Size]::new(26, 22)
                }
            }

            if ($script:_RightCollapsed) {
                $rightHeader.SetBounds(0, 0, $rightContainer.Width, $rightContainer.Height)
                $rightBody.Visible = $false
                if ($rightHeaderToggle) { $rightHeaderToggle.Text = '+' }
                if ($rightHeaderLabel) {
                    $rightHeaderLabel.Text = _VerticalText 'Editor'
                    $rightHeaderLabel.Location = [System.Drawing.Point]::new(3, 28)
                    $rightHeaderLabel.Size = [System.Drawing.Size]::new([Math]::Max(16, $rightContainer.Width - 6), [Math]::Max(40, $rightHeader.Height - 34))
                    $rightHeaderLabel.TextAlign = 'TopCenter'
                }
                if ($rightHeaderToggle) {
                    $rightHeaderToggle.Location = [System.Drawing.Point]::new(1, 4)
                    $rightHeaderToggle.Size = [System.Drawing.Size]::new([Math]::Max(24, $rightContainer.Width - 2), 22)
                }
            } else {
                $rightHeader.SetBounds(0, 0, $rightContainer.Width, $headerHeight)
                $rightBody.SetBounds(0, $headerHeight, $rightContainer.Width, ($rightContainer.Height - $headerHeight))
                $rightBody.Visible = $true
                if ($rightHeaderToggle) { $rightHeaderToggle.Text = '-' }
                if ($rightHeaderLabel) {
                    $rightHeaderLabel.Text = 'Profile Editor'
                    $rightHeaderLabel.Location = [System.Drawing.Point]::new(10, 8)
                    $rightHeaderLabel.Size = [System.Drawing.Size]::new([Math]::Max(80, $rightContainer.Width - 50), $headerHeight - 10)
                    $rightHeaderLabel.TextAlign = 'MiddleLeft'
                }
                if ($rightHeaderToggle) {
                    $rightHeaderToggle.Location = [System.Drawing.Point]::new([Math]::Max(6, $rightContainer.Width - 34), 6)
                    $rightHeaderToggle.Size = [System.Drawing.Size]::new(26, 22)
                }
            }

            if ($script:_BottomCollapsed) {
                $bottomHeader.Height = $bottomContainer.Height
                $bottomPanel.Visible = $false
                if ($bottomHeaderToggle) { $bottomHeaderToggle.Text = '+' }
                if ($bottomHeaderLabel) {
                    $bottomHeaderLabel.Text = 'Logs'
                    $bottomHeaderLabel.Location = [System.Drawing.Point]::new(10, 7)
                    $bottomHeaderLabel.Size = [System.Drawing.Size]::new([Math]::Max(80, $bottomContainer.Width - 80), 16)
                    $bottomHeaderLabel.TextAlign = 'MiddleLeft'
                }
                if ($bottomHeaderToggle) {
                    $bottomHeaderToggle.Location = [System.Drawing.Point]::new([Math]::Max(6, $bottomContainer.Width - 34), 4)
                    $bottomHeaderToggle.Size = [System.Drawing.Size]::new(26, 22)
                }
            } else {
                $bottomHeader.Height = $bottomHeaderHeight
                $bottomPanel.Visible = $true
                if ($bottomHeaderToggle) { $bottomHeaderToggle.Text = '-' }
                if ($bottomHeaderLabel) {
                    $bottomHeaderLabel.Text = 'Logs'
                    $bottomHeaderLabel.Location = [System.Drawing.Point]::new(10, 7)
                    $bottomHeaderLabel.Size = [System.Drawing.Size]::new([Math]::Max(80, $bottomContainer.Width - 80), 16)
                    $bottomHeaderLabel.TextAlign = 'MiddleLeft'
                }
                if ($bottomHeaderToggle) {
                    $bottomHeaderToggle.Location = [System.Drawing.Point]::new([Math]::Max(6, $bottomContainer.Width - 34), 4)
                    $bottomHeaderToggle.Size = [System.Drawing.Size]::new(26, 22)
                }
            }

            try {
                if ($leftHeaderLabel) { $leftHeaderLabel.BringToFront() }
                if ($leftHeaderToggle) { $leftHeaderToggle.BringToFront() }
                if ($rightHeaderLabel) { $rightHeaderLabel.BringToFront() }
                if ($rightHeaderToggle) { $rightHeaderToggle.BringToFront() }
                if ($bottomHeaderLabel) { $bottomHeaderLabel.BringToFront() }
                if ($bottomHeaderToggle) { $bottomHeaderToggle.BringToFront() }
            } catch { }
            if (_IsGuiDebugEnabled) {
                try { _QueueStatusMessage ("PANELTOGGLE apply pane={0};mode=direct-layout;ok=1" -f $PaneKey) } catch { }
                _QueueStatusMessage ("PANELTOGGLE result pane={0};left={1}x{2};center={3}x{4};right={5}x{6};bottom={7}x{8};leftBody={9};rightBody={10};bottomPanel={11}" -f `
                    $PaneKey, `
                    $leftContainer.Width, $leftContainer.Height, `
                    $centerCol.Width, $centerCol.Height, `
                    $rightContainer.Width, $rightContainer.Height, `
                    $bottomContainer.Width, $bottomContainer.Height, `
                    $(if ($leftBody.Visible) { 'visible' } else { 'hidden' }), `
                    $(if ($rightBody.Visible) { 'visible' } else { 'hidden' }), `
                    $(if ($bottomPanel.Visible) { 'visible' } else { 'hidden' }))
            }
        } catch {
            if (_IsGuiDebugEnabled) {
                try { _QueueStatusMessage ("PANELTOGGLE error pane={0};control={1};msg={2}" -f $PaneKey, $ControlName, $_.Exception.Message) } catch { }
            }
        }
    }

    _BindClickHandler -Control $leftHeaderToggle -Handler { _InvokePaneToggleDirect 'Left' 'LeftHeaderToggle' }
    _BindClickHandler -Control $rightHeaderToggle -Handler { _InvokePaneToggleDirect 'Right' 'RightHeaderToggle' }
    _BindClickHandler -Control $bottomHeaderToggle -Handler { _InvokePaneToggleDirect 'Bottom' 'BottomHeaderToggle' }

    _BindClickHandler -Control $leftHeader -Handler { try { $leftHeaderToggle.PerformClick() } catch { } }
    _BindClickHandler -Control $leftHeaderLabel -Handler { try { $leftHeaderToggle.PerformClick() } catch { } }
    _BindClickHandler -Control $rightHeader -Handler { try { $rightHeaderToggle.PerformClick() } catch { } }
    _BindClickHandler -Control $rightHeaderLabel -Handler { try { $rightHeaderToggle.PerformClick() } catch { } }
    _BindClickHandler -Control $bottomHeader -Handler { try { $bottomHeaderToggle.PerformClick() } catch { } }
    _BindClickHandler -Control $bottomHeaderLabel -Handler { try { $bottomHeaderToggle.PerformClick() } catch { } }

    $invokeMainWindowReflow = {
        param([string]$reason = 'unknown')
        try {
            if (_IsGuiDebugEnabled) {
                _QueueStatusMessage ("RESIZE {0} :: client={1}x{2};state={3}" -f $reason, $form.ClientSize.Width, $form.ClientSize.Height, $form.WindowState)
            }
        } catch { }
        try { & $normalizeResizeChrome } catch { }
        if ($script:_ReflowLayoutHandler) { & $script:_ReflowLayoutHandler }
    }.GetNewClosure()

    $finalizeMainWindowResize = {
        param([string]$reason = 'unknown')
        try {
            $script:_LastMainWindowClientSignature = '{0}x{1}' -f $form.ClientSize.Width, $form.ClientSize.Height
        } catch { }
        try { if ($script:_ReflowLayoutHandler) { & $script:_ReflowLayoutHandler } } catch { }
        try { & $normalizeResizeChrome } catch { }
        try {
            if (_IsGuiDebugEnabled) {
                _QueueStatusMessage ("RESIZE Finalize{0} :: client={1}x{2};state={3}" -f $reason, $form.ClientSize.Width, $form.ClientSize.Height, $form.WindowState)
            }
        } catch { }
    }.GetNewClosure()

    $form.add_Resize({
        & $invokeMainWindowReflow 'FormResizeInvoke'
    }.GetNewClosure())
    $form.add_SizeChanged({
        & $invokeMainWindowReflow 'FormSizeChanged'
    }.GetNewClosure())
    $form.add_ClientSizeChanged({
        & $invokeMainWindowReflow 'FormClientSizeChanged'
    }.GetNewClosure())
    $form.add_ResizeEnd({
        & $finalizeMainWindowResize 'FormResizeEnd'
    }.GetNewClosure())
    $form.add_StyleChanged({
        & $invokeMainWindowReflow 'FormStyleChanged'
    }.GetNewClosure())
    & $script:_ReflowLayoutHandler
    try { & $normalizeResizeChrome } catch { }
    try {
        if (_IsGuiDebugEnabled) {
            _QueueStatusMessage ((
                "RESIZE InitialLayout :: client={0}x{1};statusTop={2};bottomGrip={3},{4},{5},{6};rightGrip={7},{8},{9},{10};bottomRight={11},{12},{13},{14}" -f
                $form.ClientSize.Width,
                $form.ClientSize.Height,
                $(if ($statusBar) { $statusBar.Top } else { -1 }),
                $(if ($script:_ResizeBottomGrip) { $script:_ResizeBottomGrip.Left } else { -1 }),
                $(if ($script:_ResizeBottomGrip) { $script:_ResizeBottomGrip.Top } else { -1 }),
                $(if ($script:_ResizeBottomGrip) { $script:_ResizeBottomGrip.Width } else { -1 }),
                $(if ($script:_ResizeBottomGrip) { $script:_ResizeBottomGrip.Height } else { -1 }),
                $(if ($script:_ResizeRightGrip) { $script:_ResizeRightGrip.Left } else { -1 }),
                $(if ($script:_ResizeRightGrip) { $script:_ResizeRightGrip.Top } else { -1 }),
                $(if ($script:_ResizeRightGrip) { $script:_ResizeRightGrip.Width } else { -1 }),
                $(if ($script:_ResizeRightGrip) { $script:_ResizeRightGrip.Height } else { -1 }),
                $(if ($script:_ResizeBottomRightGrip) { $script:_ResizeBottomRightGrip.Left } else { -1 }),
                $(if ($script:_ResizeBottomRightGrip) { $script:_ResizeBottomRightGrip.Top } else { -1 }),
                $(if ($script:_ResizeBottomRightGrip) { $script:_ResizeBottomRightGrip.Width } else { -1 }),
                $(if ($script:_ResizeBottomRightGrip) { $script:_ResizeBottomRightGrip.Height } else { -1 })
            ))
        }
    } catch { }

    $form.Add_Shown({
        try {
            $form.BeginInvoke([System.Windows.Forms.MethodInvoker]{
                try { & $normalizeResizeChrome } catch { }
                try {
                    if ($statusBar) { $statusBar.BringToFront() }
                    if ($script:_ResizeBottomGrip) { $script:_ResizeBottomGrip.BringToFront() }
                    if ($script:_ResizeLeftGrip) { $script:_ResizeLeftGrip.BringToFront() }
                    if ($script:_ResizeRightGrip) { $script:_ResizeRightGrip.BringToFront() }
                    if ($script:_ResizeBottomLeftGrip) { $script:_ResizeBottomLeftGrip.BringToFront() }
                    if ($script:_ResizeBottomRightGrip) { $script:_ResizeBottomRightGrip.BringToFront() }
                } catch { }
                try {
                    if (_IsGuiDebugEnabled) {
                        _QueueStatusMessage ((
                            "RESIZE FormShown :: client={0}x{1};statusTop={2};bottomGrip={3},{4},{5},{6};rightGrip={7},{8},{9},{10};bottomRight={11},{12},{13},{14}" -f
                            $form.ClientSize.Width,
                            $form.ClientSize.Height,
                            $(if ($statusBar) { $statusBar.Top } else { -1 }),
                            $(if ($script:_ResizeBottomGrip) { $script:_ResizeBottomGrip.Left } else { -1 }),
                            $(if ($script:_ResizeBottomGrip) { $script:_ResizeBottomGrip.Top } else { -1 }),
                            $(if ($script:_ResizeBottomGrip) { $script:_ResizeBottomGrip.Width } else { -1 }),
                            $(if ($script:_ResizeBottomGrip) { $script:_ResizeBottomGrip.Height } else { -1 }),
                            $(if ($script:_ResizeRightGrip) { $script:_ResizeRightGrip.Left } else { -1 }),
                            $(if ($script:_ResizeRightGrip) { $script:_ResizeRightGrip.Top } else { -1 }),
                            $(if ($script:_ResizeRightGrip) { $script:_ResizeRightGrip.Width } else { -1 }),
                            $(if ($script:_ResizeRightGrip) { $script:_ResizeRightGrip.Height } else { -1 }),
                            $(if ($script:_ResizeBottomRightGrip) { $script:_ResizeBottomRightGrip.Left } else { -1 }),
                            $(if ($script:_ResizeBottomRightGrip) { $script:_ResizeBottomRightGrip.Top } else { -1 }),
                            $(if ($script:_ResizeBottomRightGrip) { $script:_ResizeBottomRightGrip.Width } else { -1 }),
                            $(if ($script:_ResizeBottomRightGrip) { $script:_ResizeBottomRightGrip.Height } else { -1 })
                        ))
                    }
                } catch { }
            }) | Out-Null
        } catch { }
    }.GetNewClosure())


    $script:_AppShutdownState = @{
        InProgress         = $false
        AllowClose         = $false
        FinalCloseQueued   = $false
        FinalCloseAfter    = $null
        WorkerRunspace     = $null
        WorkerPS           = $null
        WorkerHandle       = $null
        WorkerCompleted    = $false
        WorkerEnded        = $false
        RequestedAt        = $null
        WaitStartedAt      = $null
        LastPendingSummary = ''
        LastRenderedPhase  = ''
    }

    function _QueueManagedShutdownLog {
        param(
            [string]$Message,
            [ValidateSet('INFO','WARN','ERROR','DEBUG')][string]$Level = 'INFO'
        )

        if ([string]::IsNullOrWhiteSpace($Message)) { return }
        $entry = "[{0}][{1}][GUI] {2}" -f (Get-Date -f 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
        try {
            if ($script:SharedState -and $script:SharedState.LogQueue) {
                $script:SharedState.LogQueue.Enqueue($entry)
            }
        } catch { }
        try {
            if ($script:_ProgramLogBox -and $script:_ProgramLogBox.IsHandleCreated -and -not $script:_ProgramLogBox.IsDisposed) {
                $null = $script:_ProgramLogBox.BeginInvoke([System.Windows.Forms.MethodInvoker]{
                    try { _WriteProgramLog $entry } catch { }
                })
            }
        } catch { }
        _GuiModuleLog -Message $Message -Level $Level
    }

    function _SetShutdownUiState {
        function Disable-InteractiveControls {
            param([System.Windows.Forms.Control]$Root)

            if ($null -eq $Root) { return }

            foreach ($child in @($Root.Controls)) {
                if ($null -eq $child) { continue }

                $shouldDisable = (
                    $child -is [System.Windows.Forms.ButtonBase] -or
                    $child -is [System.Windows.Forms.ComboBox] -or
                    $child -is [System.Windows.Forms.ListView] -or
                    $child -is [System.Windows.Forms.CheckBox] -or
                    $child -is [System.Windows.Forms.RadioButton] -or
                    $child -is [System.Windows.Forms.NumericUpDown] -or
                    $child -is [System.Windows.Forms.DateTimePicker]
                )

                if ($shouldDisable) {
                    try { $child.Enabled = $false } catch { }
                }

                Disable-InteractiveControls -Root $child
            }
        }

        try { $form.UseWaitCursor = $false } catch { }
        try { $form.Cursor = [System.Windows.Forms.Cursors]::Default } catch { }
        try { Disable-InteractiveControls -Root $form } catch { }
        try { $btnWinClose.Enabled = $false } catch { }
    }

    function _TestAsyncHandleCompleted {
        param($Handle)
        if ($null -eq $Handle) { return $true }
        try { return [bool]$Handle.IsCompleted } catch { return $false }
    }

    function _GetPendingShutdownDependencies {
        $pending = New-Object 'System.Collections.Generic.List[string]'

        if (-not (_TestAsyncHandleCompleted $script:MetricsHandle)) { $pending.Add('metrics worker') | Out-Null }
        if (-not (_TestAsyncHandleCompleted $script:LogTailHandle)) { $pending.Add('log tail worker') | Out-Null }

        if ($script:SharedState) {
            $listenerHandle = $null
            $monitorHandle = $null
            try { if ($script:SharedState.ContainsKey('ListenerHandle')) { $listenerHandle = $script:SharedState['ListenerHandle'] } } catch { $listenerHandle = $null }
            try { if ($script:SharedState.ContainsKey('MonitorHandle')) { $monitorHandle = $script:SharedState['MonitorHandle'] } } catch { $monitorHandle = $null }

            if (-not (_TestAsyncHandleCompleted $listenerHandle)) { $pending.Add('Discord listener') | Out-Null }
            if (-not (_TestAsyncHandleCompleted $monitorHandle)) { $pending.Add('server monitor') | Out-Null }
        }

        if ($script:_BulkStartTimer -and -not $script:_BulkStartTimer.IsDisposed) {
            try {
                if ($script:_BulkStartTimer.Enabled) {
                    $pending.Add('bulk start timer') | Out-Null
                }
            } catch { }
        }

        return @($pending.ToArray())
    }

    function _FinishManagedShutdownWorker {
        $state = $script:_AppShutdownState
        if (-not $state -or $state.WorkerEnded) { return }

        $state.WorkerEnded = $true
        try {
            if ($state.WorkerHandle -and $state.WorkerPS) {
                try { $state.WorkerPS.EndInvoke($state.WorkerHandle) | Out-Null } catch { }
            }
        } finally {
            try { if ($state.WorkerPS) { $state.WorkerPS.Dispose() } } catch { }
            try { if ($state.WorkerRunspace) { $state.WorkerRunspace.Close(); $state.WorkerRunspace.Dispose() } } catch { }
            $state.WorkerPS = $null
            $state.WorkerRunspace = $null
            $state.WorkerHandle = $null
        }
    }

    function _BeginManagedApplicationShutdown {
        if ($script:_UIReloadRequested -eq $true) { return }

        $state = $script:_AppShutdownState
        if ($state.InProgress) {
            try { $statusLabel.Text = 'Shutdown already in progress. Waiting for ECC to finish closing.' } catch { }
            return
        }

        $state.InProgress = $true
        $state.AllowClose = $false
        $state.FinalCloseQueued = $false
        $state.FinalCloseAfter = $null
        $state.WorkerCompleted = $false
        $state.WorkerEnded = $false
        $state.RequestedAt = Get-Date
        $state.WaitStartedAt = $null
        $state.LastPendingSummary = ''
        $state.LastRenderedPhase = ''

        try {
            $script:SharedState['ManagedShutdownComplete'] = $false
            $script:SharedState['AppShutdownPhase'] = 'Shutdown requested. Preparing managed shutdown sequence.'
            $script:SharedState['AppShutdownWorkerDone'] = $false
            $script:SharedState['ShutdownDisplayQueue'] = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
        } catch { }

        _SetShutdownUiState
        try { $statusLabel.Text = 'Shutdown requested. ECC is stopping managed servers before closing.' } catch { }
        _QueueManagedShutdownLog -Message 'Shutdown requested from the main window close action. The UI will stay open until managed servers and background workers are fully stopped.'

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

            function Write-ShutdownProgress {
                param(
                    [string]$Message,
                    [string]$Level = 'INFO'
                )

                if ([string]::IsNullOrWhiteSpace($Message)) { return }
                $entry = "[{0}][{1}][GUI] {2}" -f (Get-Date -f 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
                try { $SharedState['AppShutdownPhase'] = $Message } catch { }
                try {
                    if ($SharedState -and $SharedState.LogQueue) {
                        $SharedState.LogQueue.Enqueue($entry)
                    }
                } catch { }
                try {
                    if ($SharedState -and $SharedState.ContainsKey('ShutdownDisplayQueue') -and $SharedState['ShutdownDisplayQueue']) {
                        $SharedState['ShutdownDisplayQueue'].Enqueue($entry)
                    }
                } catch { }
            }

            function Wait-SharedHandle {
                param(
                    [string]$Name,
                    [string]$Key,
                    [int]$TimeoutSeconds = 20
                )

                $handle = $null
                try {
                    if ($SharedState.ContainsKey($Key)) {
                        $handle = $SharedState[$Key]
                    }
                } catch { $handle = $null }

                if ($null -eq $handle) {
                    Write-ShutdownProgress -Message ("Shutdown phase 3/4: {0} was already inactive." -f $Name)
                    return
                }

                Write-ShutdownProgress -Message ("Shutdown phase 3/4: waiting for {0} to exit." -f $Name)
                $deadline = (Get-Date).AddSeconds([Math]::Max(1, $TimeoutSeconds))
                while ((Get-Date) -lt $deadline) {
                    $done = $false
                    try { $done = [bool]$handle.IsCompleted } catch { $done = $false }
                    if ($done) {
                        Write-ShutdownProgress -Message ("Shutdown phase 3/4: {0} exited cleanly." -f $Name)
                        return
                    }
                    Start-Sleep -Milliseconds 200
                }

                Write-ShutdownProgress -Message ("Shutdown phase 3/4: {0} did not exit before timeout. Final close will continue after best-effort cleanup." -f $Name) -Level 'WARN'
            }

            try {
                Import-Module (Join-Path $ModulesDir 'Logging.psm1')        -Force
                Import-Module (Join-Path $ModulesDir 'ProfileManager.psm1') -Force
                Import-Module (Join-Path $ModulesDir 'ServerManager.psm1')  -Force
                Import-Module (Join-Path $ModulesDir 'DiscordListener.psm1') -Force

                $runningPrefixes = @()
                try { $runningPrefixes = @($SharedState.RunningServers.Keys | Sort-Object) } catch { $runningPrefixes = @() }
                $serverCount = @($runningPrefixes).Count
                Write-ShutdownProgress -Message ("Shutdown phase 1/4: stopping managed servers. {0} tracked server(s) queued for sequential shutdown." -f $serverCount)

                $index = 0
                foreach ($prefix in $runningPrefixes) {
                    $index++
                    $gameName = [string]$prefix
                    try {
                        if ($SharedState.Profiles -and $SharedState.Profiles.ContainsKey($prefix)) {
                            $profile = $SharedState.Profiles[$prefix]
                            if ($profile -and $profile.GameName) { $gameName = [string]$profile.GameName }
                        }
                    } catch { }

                    $stillRunning = $false
                    try { $stillRunning = ($SharedState.RunningServers -and $SharedState.RunningServers.ContainsKey($prefix)) } catch { $stillRunning = $false }
                    if (-not $stillRunning) {
                        Write-ShutdownProgress -Message ("Phase 1 detail: [{0}] {1} was already stopped before its shutdown turn." -f $prefix, $gameName)
                        continue
                    }

                    Write-ShutdownProgress -Message ("Phase 1 detail: stopping [{0}] {1} ({2}/{3}) using its configured save and stop methods." -f $prefix, $gameName, $index, $serverCount)
                    try {
                        Invoke-SafeShutdown -Prefix $prefix -Quiet | Out-Null
                        Write-ShutdownProgress -Message ("Phase 1 detail: [{0}] {1} shutdown sequence completed." -f $prefix, $gameName)
                    } catch {
                        Write-ShutdownProgress -Message ("Phase 1 detail: [{0}] {1} shutdown failed: {2}" -f $prefix, $gameName, $_.Exception.Message) -Level 'WARN'
                    }
                }

                try { $SharedState['ManagedShutdownComplete'] = $true } catch { }
                Write-ShutdownProgress -Message 'Shutdown phase 2/4: signaling background workers to stop.'
                try { $SharedState['StopMetricsWorker'] = $true } catch { }
                try { $SharedState['StopLogTailWorker'] = $true } catch { }
                try { $SharedState['StopListener'] = $true } catch { }
                try { $SharedState['StopMonitor'] = $true } catch { }
                Write-ShutdownProgress -Message 'Phase 2 detail: metrics worker stop requested.'
                Write-ShutdownProgress -Message 'Phase 2 detail: log tail worker stop requested.'
                Write-ShutdownProgress -Message 'Phase 2 detail: Discord listener stop requested.'
                Write-ShutdownProgress -Message 'Phase 2 detail: server monitor stop requested.'

                Write-ShutdownProgress -Message 'Shutdown phase 3/4: waiting for background workers to exit.'
                Wait-SharedHandle -Name 'Discord listener' -Key 'ListenerHandle' -TimeoutSeconds 20
                Wait-SharedHandle -Name 'server monitor' -Key 'MonitorHandle' -TimeoutSeconds 20

                Write-ShutdownProgress -Message 'Shutdown phase 4/4: managed shutdown orchestration is complete. Waiting for GUI workers and timers to finish.'
            } catch {
                Write-ShutdownProgress -Message ("Managed shutdown orchestration failed: {0}" -f $_.Exception.Message) -Level 'ERROR'
                try { $SharedState['StopMetricsWorker'] = $true } catch { }
                try { $SharedState['StopLogTailWorker'] = $true } catch { }
                try { $SharedState['StopListener'] = $true } catch { }
                try { $SharedState['StopMonitor'] = $true } catch { }
            } finally {
                try { $SharedState['AppShutdownWorkerDone'] = $true } catch { }
            }
        }) | Out-Null

        $state.WorkerRunspace = $rs
        $state.WorkerPS = $ps
        $state.WorkerHandle = $ps.BeginInvoke()
    }

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

    function _GetProgramLogCategory {
        param([string]$Line)

        if ([string]::IsNullOrWhiteSpace($Line)) { return 'empty' }

        $level = 'UNK'
        $source = 'General'
        if ($Line -match '^\[[^\]]+\]\[(?<level>[^\]]+)\]\[(?<source>[^\]]+)\]') {
            $level = [string]$matches['level']
            $source = [string]$matches['source']
        } elseif ($Line -match '^\[[^\]]+\]\[(?<level>[^\]]+)\]') {
            $level = [string]$matches['level']
        }

        return ('{0}:{1}' -f $level.ToUpperInvariant(), $source)
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

    try {
        _RegisterGuiActivityHooks -Root $form
    } catch { }

    $timer.add_Tick({
        try {
            $script:_GuiTickSequence = [int]$script:_GuiTickSequence + 1
            $uiTickId = [int]$script:_GuiTickSequence
            $uiTickPerf = [System.Diagnostics.Stopwatch]::StartNew()
            if ($script:_AppShutdownState -and $script:_AppShutdownState.InProgress) {
                try {
                    if ($script:SharedState -and $script:SharedState.ContainsKey('ShutdownDisplayQueue') -and $script:SharedState['ShutdownDisplayQueue']) {
                        for ($shutdownLogIndex = 0; $shutdownLogIndex -lt 12; $shutdownLogIndex++) {
                            $shutdownEntry = $null
                            if (-not $script:SharedState['ShutdownDisplayQueue'].TryDequeue([ref]$shutdownEntry)) { break }
                            if (-not [string]::IsNullOrWhiteSpace($shutdownEntry)) {
                                try { _WriteProgramLog $shutdownEntry } catch { }
                            }
                        }
                    }
                } catch { }

                $shutdownPhase = ''
                try {
                    if ($script:SharedState -and $script:SharedState.ContainsKey('AppShutdownPhase')) {
                        $shutdownPhase = [string]$script:SharedState['AppShutdownPhase']
                    }
                } catch { $shutdownPhase = '' }
                if (-not [string]::IsNullOrWhiteSpace($shutdownPhase)) {
                    try { $statusLabel.Text = $shutdownPhase } catch { }
                    if ($shutdownPhase -ne $script:_AppShutdownState.LastRenderedPhase) {
                        $script:_AppShutdownState.LastRenderedPhase = $shutdownPhase
                    }
                }

                if (-not $script:_AppShutdownState.WorkerCompleted) {
                    $script:_AppShutdownState.WorkerCompleted = (_TestAsyncHandleCompleted $script:_AppShutdownState.WorkerHandle)
                    if ($script:_AppShutdownState.WorkerCompleted) {
                        _FinishManagedShutdownWorker
                    }
                }

                if ($script:_AppShutdownState.WorkerCompleted) {
                    $pending = @(_GetPendingShutdownDependencies)
                    if ($pending.Count -gt 0) {
                        if ($null -eq $script:_AppShutdownState.WaitStartedAt) {
                            $script:_AppShutdownState.WaitStartedAt = Get-Date
                        }

                        $summary = $pending -join ', '
                        if ($summary -ne $script:_AppShutdownState.LastPendingSummary) {
                            $script:_AppShutdownState.LastPendingSummary = $summary
                            _QueueManagedShutdownLog -Message ("Shutdown phase 4/4: waiting for {0}." -f $summary)
                        }
                        try { $statusLabel.Text = "Finishing shutdown: waiting for $summary." } catch { }

                        $waitSeconds = 0
                        try { $waitSeconds = ((Get-Date) - [datetime]$script:_AppShutdownState.WaitStartedAt).TotalSeconds } catch { $waitSeconds = 0 }
                        if ($waitSeconds -lt 20) {
                            # Keep the window alive while workers wind down.
                        } else {
                            _QueueManagedShutdownLog -Message ("Shutdown phase 4/4: timed out waiting for {0}. Closing after best-effort cleanup." -f $summary) -Level WARN
                            $pending = @()
                        }
                    }

                    if ($pending.Count -eq 0 -and -not $script:_AppShutdownState.FinalCloseQueued) {
                        $script:_AppShutdownState.FinalCloseQueued = $true
                        $script:_AppShutdownState.FinalCloseAfter = (Get-Date).AddMilliseconds(900)
                        _QueueManagedShutdownLog -Message 'Shutdown complete. Closing the UI now.'
                        try { $statusLabel.Text = 'Shutdown complete. Closing UI...' } catch { }
                    }

                    if ($script:_AppShutdownState.FinalCloseQueued -and -not $script:_AppShutdownState.AllowClose) {
                        $closeReady = $false
                        try {
                            if ($script:_AppShutdownState.FinalCloseAfter -is [datetime] -and (Get-Date) -ge [datetime]$script:_AppShutdownState.FinalCloseAfter) {
                                $closeReady = $true
                            }
                        } catch { $closeReady = $true }

                        if ($closeReady) {
                            $script:_AppShutdownState.AllowClose = $true
                            try {
                                $form.BeginInvoke([System.Windows.Forms.MethodInvoker]{
                                    try { $form.Close() } catch { }
                                }) | Out-Null
                            } catch {
                                try { $form.Close() } catch { }
                            }
                        }
                    }
                }
            }
            try {
                $clientSignature = '{0}x{1}' -f $form.ClientSize.Width, $form.ClientSize.Height
                $lastClientSignature = ''
                try { $lastClientSignature = [string]$script:_LastMainWindowClientSignature } catch { $lastClientSignature = '' }
                if ($clientSignature -ne $lastClientSignature) {
                    $script:_LastMainWindowClientSignature = $clientSignature
                    try { & $normalizeResizeChrome } catch { }
                    try { if ($script:_ReflowLayoutHandler) { & $script:_ReflowLayoutHandler } } catch { }
                    try {
                        if (_IsGuiDebugEnabled) {
                            _QueueStatusMessage ("RESIZE TimerReflow :: client={0};tick={1}" -f $clientSignature, $uiTickId)
                        }
                    } catch { }
                }
            } catch { }
            $uiPerfLogTabsMs = 0.0
            $uiPerfProgramLogMs = 0.0
            $uiPerfProgramLogCount = 0
            $uiPerfGameLogMs = 0.0
            $uiPerfGameLogDequeued = 0
            $uiPerfGameLogPainted = 0
            $uiPerfListenerMs = 0.0
            $uiPerfDashboardMs = 0.0
            $uiPerfHeaderMs = 0.0
            $uiPerfParserMs = 0.0
            $uiPerfParserLines = 0
            $uiPerfParserPrefixes = @{}
            $editorScrollPanel = $null
            $savedEditorScroll = $null
            try {
                if ($script:_ProfileEditorPanel -and $script:_ProfileEditorPanel.Controls.Count -gt 0) {
                    $editorScrollPanel = ($script:_ProfileEditorPanel.Controls | Select-Object -First 1)
                    $savedEditorScroll = _CaptureScrollPosition $editorScrollPanel
                }
            } catch { }

            # --- Read latest metrics written by background metrics runspace ---
            # (unchanged)

            # --- NO LOG TAILING HERE ANYMORE ---
            # Background runspaces now handle all file reading.
            # This timer ONLY updates UI and flushes SharedState queues.

            # --- Manage per-game log tabs (UI only, no file I/O) ---
            $uiPhasePerf = [System.Diagnostics.Stopwatch]::StartNew()
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
                        _EnsureGameLogTab -Prefix $pfx -GameName $gameName -Profile $profile

                        $tabEntry = $script:_GameLogTabs[$pfx]
                        $rtb      = $tabEntry.RTB
                        $lblSrc   = $tabEntry.LblSrc

                        # Resolve which files SHOULD be tailed (background thread handles actual tailing)
                        $logFiles = _ResolveGameLogFiles -Profile $profile

                        # Detect PZ session folder change (UI only)
                        if ($profile.LogStrategy -eq 'PZSessionFolder' -and
                            $profile.ServerLogRoot -and (Test-Path $profile.ServerLogRoot)) {

                            $resolvedPz = _ResolveProjectZomboidSessionLogFiles -Profile $profile
                            $sessionKey = if ($resolvedPz -and $resolvedPz.SessionKey) { [string]$resolvedPz.SessionKey } else { '' }

                            if (-not [string]::IsNullOrWhiteSpace($sessionKey) -and $sessionKey -ne $tabEntry.LastSession) {
                                $rtb.Clear()
                                _ResetLogBoxState $rtb
                                $tabEntry.Files       = @{}
                                $tabEntry.LastSession = $sessionKey
                                $tabEntry.HistoryLoaded = $false
                                $tabEntry.PendingLines.Clear()
                                _AppendLog $rtb "[SESSION] New PZ session: $sessionKey" $clrYellow
                                # Reset start notification for new session
                                $script:_ServerStartNotified.Remove($pfx) | Out-Null
                                _LoadRecentGameLogHistory -Prefix $pfx -Profile $profile -Force
                            }
                        }

                        # If no log file found yet, show note once (use a flag, not Lines.Count)
                        if ($logFiles.Count -eq 0) {
                            if ($profile.DisableFileTail -eq $true) {
                                $lblSrc.Text = 'Activity feed (no file tail)'
                            }
                            $note = $profile.ServerLogNote
                            if (-not [string]::IsNullOrEmpty($note) -and -not $tabEntry.NoteShown) {
                                _AppendLog $rtb "[INFO] $note" $clrMuted
                                $tabEntry.NoteShown = $true
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
            $uiPhasePerf.Stop()
            $uiPerfLogTabsMs = $uiPhasePerf.Elapsed.TotalMilliseconds
            _TraceGuiPerformanceSample -Area 'GameLogTabSync' `
                -ElapsedMs $uiPerfLogTabsMs `
                -WarnAtMs 120 `
                -DebugAtMs 40 `
                -Detail ('tick={0};tabs={1};running={2}' -f $uiTickId, `
                    $(if ($script:_GameLogTabs) { @($script:_GameLogTabs.Keys).Count } else { 0 }), `
                    $(if ($script:SharedState -and $script:SharedState.RunningServers) { @($script:SharedState.RunningServers.Keys).Count } else { 0 }))

            # --- Drain SharedState LogQueue (UI only) ---
            $uiPhasePerf = [System.Diagnostics.Stopwatch]::StartNew()
            $programLogQueueBefore = 0
            $programLogCharsRendered = 0
            $programLogMaxLineLength = 0
            $programLogCategories = @{}
            try {
                if ($script:SharedState -and $script:SharedState.LogQueue) {
                    $programLogQueueBefore = [int]$script:SharedState.LogQueue.Count
                }
            } catch { $programLogQueueBefore = 0 }
            for ($i = 0; $i -lt 20; $i++) {
                $item = $null
                if (-not $script:SharedState.LogQueue.TryDequeue([ref]$item)) { break }
                $uiPerfProgramLogCount++
                $itemText = [string]$item
                $programLogCharsRendered += $itemText.Length
                if ($itemText.Length -gt $programLogMaxLineLength) { $programLogMaxLineLength = $itemText.Length }
                $category = _GetProgramLogCategory -Line $itemText
                if ($programLogCategories.ContainsKey($category)) {
                    $programLogCategories[$category] = [int]$programLogCategories[$category] + 1
                } else {
                    $programLogCategories[$category] = 1
                }
                if ($item -match '\[Discord\]') {
                    _WriteDiscordLog $item
                } else {
                    _WriteProgramLog $item
                }
            }
            $uiPhasePerf.Stop()
            $uiPerfProgramLogMs = $uiPhasePerf.Elapsed.TotalMilliseconds
            _TraceGuiPerformanceSample -Area 'ProgramLogDrain' `
                -ElapsedMs $uiPerfProgramLogMs `
                -WarnAtMs 90 `
                -DebugAtMs 30 `
                -Detail ('tick={0};drained={1}' -f $uiTickId, $uiPerfProgramLogCount)
            if (_IsGuiDebugEnabled) {
                try {
                    $topCategories = @($programLogCategories.GetEnumerator() |
                        Sort-Object Value -Descending |
                        Select-Object -First 3 |
                        ForEach-Object { '{0}={1}' -f $_.Key, $_.Value }) -join ','
                    if ([string]::IsNullOrWhiteSpace($topCategories)) { $topCategories = 'none' }
                    _GuiDirectLog -Level DEBUG -Message ('PROGRAMLOGDRAIN tick={0} queueBefore={1} drained={2} renderedChars={3} maxLineChars={4} categories={5}' -f `
                        $uiTickId, $programLogQueueBefore, $uiPerfProgramLogCount, $programLogCharsRendered, $programLogMaxLineLength, $topCategories)
                } catch { }
            }

            # --- Drain GameLogQueue (UI only) ---
            # We process up to 120 lines per tick from the queue.
            # RTB painting is budget-capped (20 painted lines per tick) to keep
            # the UI responsive, but detection logic (start-notify, players-capture)
            # runs on EVERY dequeued line regardless of paint budget or tab visibility.
            # This ensures Valheim "Game server connected" and 7DTD "INF StartGame done"
            # are never silently dropped just because another game's tab is filling up.
            $uiPhasePerf = [System.Diagnostics.Stopwatch]::StartNew()
            $gameLogPaintBudget = 20
            $gameLogPaintCount  = 0
            $activeTabPage = $null
            try { $activeTabPage = $script:_LogTabControl.SelectedTab } catch { }

            for ($i = 0; $i -lt 120; $i++) {
                $entry = $null
                if (-not $script:SharedState.GameLogQueue.TryDequeue([ref]$entry)) { break }
                $uiPerfGameLogDequeued++
                if ($null -eq $entry) { continue }
                $pfx  = $entry.Prefix
                $line = $entry.Line
                if (-not $pfx -or -not $line) { continue }
                if (-not $script:_GameLogTabs.ContainsKey($pfx)) { continue }
                $tabEntry = $script:_GameLogTabs[$pfx]

                # Only paint lines into the RTB if the tab is currently selected
                # AND we have not hit the per-tick paint budget.
                # Budget check NEVER breaks the loop — detection still runs below.
                $tabVisible = ($null -ne $activeTabPage -and $tabEntry.Tab -eq $activeTabPage)
                if ($tabVisible -and $gameLogPaintCount -lt $gameLogPaintBudget) {
                    _AppendLog $tabEntry.RTB $line (_LogColour $line)
                    $gameLogPaintCount++
                } elseif ($tabEntry.PendingLines) {
                    $tabEntry.PendingLines.Add("$line") | Out-Null
                    while ($tabEntry.PendingLines.Count -gt 20) {
                        $tabEntry.PendingLines.RemoveAt(0)
                    }
                }

                # Project Zomboid and 7DTD live player tracking from normal server logs.
                # This keeps idle rules in sync even without a manual players command.
                $parserPerf = [System.Diagnostics.Stopwatch]::StartNew()
                $uiPerfParserLines++
                try { $uiPerfParserPrefixes[[string]$pfx] = $true } catch { }
                $profile = $script:SharedState.Profiles[$pfx]
                $normalizedKnownGame = ''
                try { if ($profile) { $normalizedKnownGame = (_NormalizeGameIdentity (_GetProfileKnownGame -Profile $profile)) } } catch { $normalizedKnownGame = '' }
                $isPzProfile = ($profile -and $normalizedKnownGame -eq 'projectzomboid')
                $is7dProfile = ($profile -and $normalizedKnownGame -eq '7daystodie')
                $isMinecraftProfile = ($profile -and $normalizedKnownGame -eq 'minecraft')
                if ($isPzProfile) {
                    $debugPlayers = $false
                    try {
                        $debugPlayers = ($script:SharedState -and $script:SharedState.Settings -and [bool]$script:SharedState.Settings.EnableDebugLogging)
                    } catch { $debugPlayers = $false }

                    if ($line -match '^\[[^\]]+\]\s+(\d+)\s+"([^"]+)"\s+(attempting to join(?: used queue)?|allowed to join|fully connected)\b') {
                        $connectedId = [string]$Matches[1]
                        $connectedName = [string]$Matches[2]
                        $connectedStage = [string]$Matches[3]
                        _RememberSharedPzObservedPlayer -Prefix $pfx -PlayerId $connectedId -PlayerName $connectedName -SharedState $script:SharedState
                        if ($connectedStage -match '(?i)^fully connected') {
                            _AddSharedObservedPlayer -Prefix $pfx -PlayerName $connectedName -SharedState $script:SharedState
                            $currentNames = @(_GetSharedLatestPlayers -Prefix $pfx -SharedState $script:SharedState)
                            try { _SetObservedPlayersRuntimeState -Prefix $pfx -Names @($currentNames) -Count ([Math]::Max(1, $currentNames.Count)) -SharedState $script:SharedState } catch { }
                        }
                        if ($debugPlayers -and $script:SharedState.LogQueue) {
                            $namesText = (@(_GetSharedLatestPlayers -Prefix $pfx -SharedState $script:SharedState) -join ', ')
                            $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][DEBUG][PZLIVE] prefix=$pfx event=$connectedStage id=$connectedId player='$connectedName' players=$namesText")
                        }
                    }
                    elseif ($line -match '^\[[^\]]+\]\s+Connection\s+(?:disconnect|remove)\b.*?\bid=(\d+)\b') {
                        $disconnectedId = [string]$Matches[1]
                        $knownPlayersBefore = @(_GetSharedLatestPlayers -Prefix $pfx -SharedState $script:SharedState)
                        _RemoveSharedPzObservedPlayerById -Prefix $pfx -PlayerId $disconnectedId -SharedState $script:SharedState
                        $currentNames = @(_GetSharedLatestPlayers -Prefix $pfx -SharedState $script:SharedState)
                        try { _SetObservedPlayersRuntimeState -Prefix $pfx -Names @($currentNames) -SharedState $script:SharedState } catch { }
                        if ($debugPlayers -and $script:SharedState.LogQueue) {
                            $namesText = ($currentNames -join ', ')
                            $beforeText = if ($knownPlayersBefore.Count -gt 0) { $knownPlayersBefore -join ', ' } else { '<none>' }
                            $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][DEBUG][PZLIVE] prefix=$pfx event=disconnect id=$disconnectedId players-before=$beforeText players-after=$namesText")
                        }
                    }
                    elseif ($line -match '(?i)"([^"]+)"\s+(?:disconnected|lost connection|connection lost)\b') {
                        $disconnectedName = [string]$Matches[1]
                        _RemoveSharedObservedPlayer -Prefix $pfx -PlayerName $disconnectedName -SharedState $script:SharedState
                        $currentNames = @(_GetSharedLatestPlayers -Prefix $pfx -SharedState $script:SharedState)
                        try { _SetObservedPlayersRuntimeState -Prefix $pfx -Names @($currentNames) -SharedState $script:SharedState } catch { }
                        if ($debugPlayers -and $script:SharedState.LogQueue) {
                            $namesText = ($currentNames -join ', ')
                            $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][DEBUG][PZLIVE] prefix=$pfx event=leave-by-name player='$disconnectedName' players=$namesText")
                        }
                    }
                }
                elseif ($is7dProfile) {
                    $debugPlayers = $false
                    try {
                        $debugPlayers = ($script:SharedState -and $script:SharedState.Settings -and [bool]$script:SharedState.Settings.EnableDebugLogging)
                    } catch { $debugPlayers = $false }

                    $playerCountFromPulse = $null
                    if ($line -match '\bPly:\s*(\d+)\b') {
                        try { $playerCountFromPulse = [int]$Matches[1] } catch { $playerCountFromPulse = $null }
                    }

                    if ($line -match "GMSG:\s+Player\s+'([^']+)'\s+joined the game\b") {
                        $joinedName = [string]$Matches[1]
                        _AddSharedObservedPlayer -Prefix $pfx -PlayerName $joinedName -SharedState $script:SharedState
                        $currentNames = @(_GetSharedLatestPlayers -Prefix $pfx -SharedState $script:SharedState)
                        try { _SetObservedPlayersRuntimeState -Prefix $pfx -Names @($currentNames) -Count ([Math]::Max(1, $currentNames.Count)) -SharedState $script:SharedState } catch { }
                        if ($debugPlayers -and $script:SharedState.LogQueue) {
                            $namesText = ($currentNames -join ', ')
                            $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][DEBUG][7DTDLIVE] prefix=$pfx event=join player='$joinedName' players=$namesText")
                        }
                    }
                    elseif ($line -match "PlayerSpawnedInWorld\b.*?\bPlayerName='([^']+)'") {
                        $spawnedName = [string]$Matches[1]
                        _AddSharedObservedPlayer -Prefix $pfx -PlayerName $spawnedName -SharedState $script:SharedState
                        $currentNames = @(_GetSharedLatestPlayers -Prefix $pfx -SharedState $script:SharedState)
                        try { _SetObservedPlayersRuntimeState -Prefix $pfx -Names @($currentNames) -Count ([Math]::Max(1, $currentNames.Count)) -SharedState $script:SharedState } catch { }
                        if ($debugPlayers -and $script:SharedState.LogQueue) {
                            $namesText = ($currentNames -join ', ')
                            $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][DEBUG][7DTDLIVE] prefix=$pfx event=spawn player='$spawnedName' players=$namesText")
                        }
                    }
                    elseif ($line -match "Player\s+(.+?)\s+disconnected after\b") {
                        $disconnectedName = [string]$Matches[1]
                        _RemoveSharedObservedPlayer -Prefix $pfx -PlayerName $disconnectedName -SharedState $script:SharedState
                        $currentNames = @(_GetSharedLatestPlayers -Prefix $pfx -SharedState $script:SharedState)
                        try { _SetObservedPlayersRuntimeState -Prefix $pfx -Names @($currentNames) -SharedState $script:SharedState } catch { }
                        if ($debugPlayers -and $script:SharedState.LogQueue) {
                            $namesText = ($currentNames -join ', ')
                            $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][DEBUG][7DTDLIVE] prefix=$pfx event=disconnect player='$disconnectedName' players=$namesText")
                        }
                    }
                    elseif ($line -match "GMSG:\s+Player\s+'([^']+)'\s+left the game\b") {
                        $leftName = [string]$Matches[1]
                        _RemoveSharedObservedPlayer -Prefix $pfx -PlayerName $leftName -SharedState $script:SharedState
                        $currentNames = @(_GetSharedLatestPlayers -Prefix $pfx -SharedState $script:SharedState)
                        try { _SetObservedPlayersRuntimeState -Prefix $pfx -Names @($currentNames) -SharedState $script:SharedState } catch { }
                        if ($debugPlayers -and $script:SharedState.LogQueue) {
                            $namesText = ($currentNames -join ', ')
                            $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][DEBUG][7DTDLIVE] prefix=$pfx event=leave player='$leftName' players=$namesText")
                        }
                    }
                    elseif ($line -match "Player disconnected:\b.*?\bPlayerName='([^']+)'") {
                        $leftName = [string]$Matches[1]
                        _RemoveSharedObservedPlayer -Prefix $pfx -PlayerName $leftName -SharedState $script:SharedState
                        $currentNames = @(_GetSharedLatestPlayers -Prefix $pfx -SharedState $script:SharedState)
                        try { _SetObservedPlayersRuntimeState -Prefix $pfx -Names @($currentNames) -SharedState $script:SharedState } catch { }
                        if ($debugPlayers -and $script:SharedState.LogQueue) {
                            $namesText = ($currentNames -join ', ')
                            $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][DEBUG][7DTDLIVE] prefix=$pfx event=leave-detail player='$leftName' players=$namesText")
                        }
                    }

                    if ($null -ne $playerCountFromPulse) {
                        if ($playerCountFromPulse -le 0) {
                            Set-LatestPlayersSnapshot -Prefix $pfx -Names @() -Count 0 -SharedState $script:SharedState
                            try {
                                Set-JoinableServerRuntimeState -Prefix $pfx -SharedState $script:SharedState
                            } catch { }
                            if ($debugPlayers -and $script:SharedState.LogQueue) {
                                $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][DEBUG][7DTDLIVE] prefix=$pfx event=pulse count=0 players=<none>")
                            }
                        } else {
                            $knownPlayers = @(_GetSharedLatestPlayers -Prefix $pfx -SharedState $script:SharedState)
                            if ($knownPlayers.Count -gt 0) {
                                Set-LatestPlayersSnapshot -Prefix $pfx -Names @($knownPlayers) -Count $playerCountFromPulse -SharedState $script:SharedState
                                try { _SetObservedPlayersRuntimeState -Prefix $pfx -Names @($knownPlayers) -Count $playerCountFromPulse -SharedState $script:SharedState } catch { }
                            }
                            if ($debugPlayers -and $script:SharedState.LogQueue) {
                                $namesText = if ($knownPlayers.Count -gt 0) { $knownPlayers -join ', ' } else { '<unknown>' }
                                $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][DEBUG][7DTDLIVE] prefix=$pfx event=pulse count=$playerCountFromPulse players=$namesText")
                            }
                        }
                    }
                }
                elseif ($isMinecraftProfile) {
                    if ($line -match '(?i)\bthere are\s+(\d+)\s+of a max of\s+(\d+)\s+players?\s+online\b') {
                        $reportedCount = 0
                        try { $reportedCount = [int]$Matches[1] } catch { $reportedCount = 0 }

                        $minecraftStillRunning = $false
                        try {
                            $minecraftStillRunning = ($script:SharedState -and $script:SharedState.RunningServers -and $script:SharedState.RunningServers.ContainsKey($pfx))
                        } catch { $minecraftStillRunning = $false }
                        if (-not $minecraftStillRunning) { return }

                        try {
                            Set-LatestPlayersSnapshot -Prefix $pfx -Names @() -Count $reportedCount -SharedState $script:SharedState
                        } catch { }

                        try {
                            if ($reportedCount -gt 0) {
                                _SetObservedPlayersRuntimeState -Prefix $pfx -Names @() -Count $reportedCount -SharedState $script:SharedState
                            } else {
                                Set-JoinableServerRuntimeState -Prefix $pfx -Force -Detail 'Minecraft server is joinable and responded to a live player query.' -SharedState $script:SharedState
                            }
                        } catch { }
                    }
                }
                elseif ($profile -and $normalizedKnownGame -eq 'hytale') {
                    $debugPlayers = $false
                    try {
                        $debugPlayers = ($script:SharedState -and $script:SharedState.Settings -and [bool]$script:SharedState.Settings.EnableDebugLogging)
                    } catch { $debugPlayers = $false }
                    $startedHytaleWhoCapture = $false

                    if ($line -match "\[World\|[^\]]+\]\s+Player\s+'([^']+)'\s+joined world\b") {
                        $joinedName = [string]$Matches[1]
                        _AddSharedObservedPlayer -Prefix $pfx -PlayerName $joinedName -SharedState $script:SharedState
                        $currentNames = @(_GetSharedLatestPlayers -Prefix $pfx -SharedState $script:SharedState)
                        try { _SetObservedPlayersRuntimeState -Prefix $pfx -Names @($currentNames) -Count ([Math]::Max(1, $currentNames.Count)) -SharedState $script:SharedState } catch { }
                        if ($debugPlayers -and $script:SharedState.LogQueue) {
                            $namesText = if ($currentNames.Count -gt 0) { $currentNames -join ', ' } else { '<none>' }
                            $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][DEBUG][HYLIVE] prefix=$pfx event=joined-world player='$joinedName' players=$namesText")
                        }
                    }
                    elseif ($line -match "\[Hytale\].*?\s-\s(.+?)\s+at\s+.+\s+left with reason:" -or
                            $line -match "\[Universe\|P\]\s+Removing player\s+'([^']+)'") {
                            $leftName = [string]$Matches[1]
                        _RemoveSharedObservedPlayer -Prefix $pfx -PlayerName $leftName -SharedState $script:SharedState
                        $currentNames = @(_GetSharedLatestPlayers -Prefix $pfx -SharedState $script:SharedState)
                        try { _SetObservedPlayersRuntimeState -Prefix $pfx -Names @($currentNames) -SharedState $script:SharedState } catch { }
                        if ($debugPlayers -and $script:SharedState.LogQueue) {
                            $namesText = if ($currentNames.Count -gt 0) { $currentNames -join ', ' } else { '<none>' }
                            $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][DEBUG][HYLIVE] prefix=$pfx event=leave player='$leftName' players=$namesText")
                        }
                    }

                    if ($line -match '(?i)\[CommandManager\]\s+Console executed command:\s*who\b') {
                        $script:_HytaleWhoCapture[$pfx] = @{
                            Active  = $true
                            Count   = 0
                            Names   = New-Object System.Collections.Generic.List[string]
                            Seen    = @{}
                            Started = Get-Date
                        }
                        $startedHytaleWhoCapture = $true
                        if ($debugPlayers -and $script:SharedState.LogQueue) {
                            $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][DEBUG][HYLIVE] prefix=$pfx event=who-start")
                        }
                    }

                    $hyCap = $null
                    try { $hyCap = $script:_HytaleWhoCapture[$pfx] } catch { $hyCap = $null }
                    if ($hyCap -and $hyCap.Active -eq $true) {
                        if ($line -match '^[^()]+?\((\d+)\)\s*:\s*(.*)$') {
                            $reportedCount = 0
                            try { $reportedCount = [int]$Matches[1] } catch { $reportedCount = 0 }
                            $rosterText = [string]$Matches[2]
                            $rosterText = $rosterText.Trim()
                            $rosterText = $rosterText.TrimStart(':').Trim()

                            if (-not $hyCap.Seen) { $hyCap.Seen = @{} }
                            if (-not $hyCap.Names) { $hyCap.Names = New-Object System.Collections.Generic.List[string] }

                            if ($reportedCount -le 0 -or [string]::IsNullOrWhiteSpace($rosterText) -or $rosterText -match '^(?i)\(empty\)|empty$') {
                                $hyCap.Seen = @{}
                                $hyCap.Names = New-Object System.Collections.Generic.List[string]
                            } else {
                                foreach ($part in ($rosterText -split ',')) {
                                    $name = ([string]$part).Trim(" `t[](){}:")
                                    if ([string]::IsNullOrWhiteSpace($name)) { continue }
                                    $seenKey = $name.ToLowerInvariant()
                                    if (-not $hyCap.Seen.ContainsKey($seenKey)) {
                                        $hyCap.Seen[$seenKey] = $true
                                        $hyCap.Names.Add($name) | Out-Null
                                    }
                                }
                            }

                            $currentNames = @(_ToStringArray $hyCap.Names)
                            Set-LatestPlayersSnapshot -Prefix $pfx -Names @($currentNames) -Count $reportedCount -SharedState $script:SharedState
                            try { _SetObservedPlayersRuntimeState -Prefix $pfx -Names @($currentNames) -Count $reportedCount -SharedState $script:SharedState } catch { }

                            $script:_HytaleWhoCapture[$pfx] = $hyCap
                            if ($debugPlayers -and $script:SharedState.LogQueue) {
                                $namesText = if ($currentNames.Count -gt 0) { $currentNames -join ', ' } else { '<none>' }
                                $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][DEBUG][HYLIVE] prefix=$pfx event=who-line count=$reportedCount players=$namesText")
                            }
                        } elseif ((-not $startedHytaleWhoCapture) -and ($line -match '(?i)\[CommandManager\]\s+Console executed command:' -or (((Get-Date) - $hyCap.Started).TotalSeconds -gt 6))) {
                            $script:_HytaleWhoCapture.Remove($pfx) | Out-Null
                            if ($debugPlayers -and $script:SharedState.LogQueue) {
                                $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][DEBUG][HYLIVE] prefix=$pfx event=who-end")
                            }
                        }
                    }
                }
                elseif ($profile -and ((_NormalizeGameIdentity (_GetProfileKnownGame -Profile $profile)) -eq 'satisfactory')) {
                    $debugPlayers = $false
                    try {
                        $debugPlayers = ($script:SharedState -and $script:SharedState.Settings -and [bool]$script:SharedState.Settings.EnableDebugLogging)
                    } catch { $debugPlayers = $false }
                    $sfConn = $null
                    try { $sfConn = $script:_SatisfactoryConnectionCapture[$pfx] } catch { $sfConn = $null }
                    if ($null -eq $sfConn) {
                        $sfConn = @{
                            RecentRemoteAddr = ''
                            Connections      = @{}
                        }
                        $script:_SatisfactoryConnectionCapture[$pfx] = $sfConn
                    }

                    if ($line -match 'NotifyAcceptedConnection:.*RemoteAddr:\s+([^,]+),') {
                        $sfConn.RecentRemoteAddr = [string]$Matches[1]
                        $script:_SatisfactoryConnectionCapture[$pfx] = $sfConn
                    }

                    if ($line -match 'LogNet:\s+Join succeeded:\s+(.+)$') {
                        $joinedName = [string]$Matches[1]
                        if (-not [string]::IsNullOrWhiteSpace($sfConn.RecentRemoteAddr)) {
                            $sfConn.Connections[$sfConn.RecentRemoteAddr] = $joinedName
                            $script:_SatisfactoryConnectionCapture[$pfx] = $sfConn
                        }
                        _AddSharedObservedPlayer -Prefix $pfx -PlayerName $joinedName -SharedState $script:SharedState
                        $currentNames = @(_GetSharedLatestPlayers -Prefix $pfx -SharedState $script:SharedState)
                        try { Set-LatestPlayersSnapshot -Prefix $pfx -Names @($currentNames) -Count ([Math]::Max(1, $currentNames.Count)) -SharedState $script:SharedState } catch { }
                        try { _SetObservedPlayersRuntimeState -Prefix $pfx -Names @($currentNames) -Count ([Math]::Max(1, $currentNames.Count)) -SharedState $script:SharedState } catch { }
                        if ($debugPlayers -and $script:SharedState.LogQueue) {
                            $namesText = if ($currentNames.Count -gt 0) { $currentNames -join ', ' } else { '<none>' }
                            $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][DEBUG][SFLIVE] prefix=$pfx event=join player='$joinedName' players=$namesText")
                        }
                    }
                    elseif ($line -match 'UNetConnection::Close:\s+\[UNetConnection\]\s+RemoteAddr:\s+([^,]+),') {
                        $remoteAddr = [string]$Matches[1]
                        $leftName = ''
                        if ($sfConn.Connections.ContainsKey($remoteAddr)) {
                            $leftName = [string]$sfConn.Connections[$remoteAddr]
                            $sfConn.Connections.Remove($remoteAddr) | Out-Null
                            $script:_SatisfactoryConnectionCapture[$pfx] = $sfConn
                        } else {
                            $currentNames = @(_GetSharedLatestPlayers -Prefix $pfx -SharedState $script:SharedState)
                            if ($currentNames.Count -eq 1) {
                                $leftName = [string]$currentNames[0]
                            }
                        }

                        if (-not [string]::IsNullOrWhiteSpace($leftName)) {
                            _RemoveSharedObservedPlayer -Prefix $pfx -PlayerName $leftName -SharedState $script:SharedState
                            $currentNames = @(_GetSharedLatestPlayers -Prefix $pfx -SharedState $script:SharedState)
                            try { Set-LatestPlayersSnapshot -Prefix $pfx -Names @($currentNames) -Count $currentNames.Count -SharedState $script:SharedState } catch { }
                            try { _SetObservedPlayersRuntimeState -Prefix $pfx -Names @($currentNames) -SharedState $script:SharedState } catch { }
                            if ($debugPlayers -and $script:SharedState.LogQueue) {
                                $namesText = if ($currentNames.Count -gt 0) { $currentNames -join ', ' } else { '<none>' }
                                $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][DEBUG][SFLIVE] prefix=$pfx event=disconnect player='$leftName' remote='$remoteAddr' players=$namesText")
                            }
                        }
                    }
                }
                elseif ($profile -and ((_NormalizeGameIdentity (_GetProfileKnownGame -Profile $profile)) -eq 'valheim')) {
                    $debugPlayers = $false
                    try {
                        $debugPlayers = ($script:SharedState -and $script:SharedState.Settings -and [bool]$script:SharedState.Settings.EnableDebugLogging)
                    } catch { $debugPlayers = $false }

                    if ($null -eq $script:_ValheimPlayerCapture) {
                        $script:_ValheimPlayerCapture = $script:SharedState.ValheimPlayerCapture
                    }

                    $vhState = $null
                    try { $vhState = $script:_ValheimPlayerCapture[$pfx] } catch { $vhState = $null }
                    if ($null -eq $vhState) {
                        $vhState = @{
                            Zdos = @{}
                            LastReportedCount = -1
                        }
                        $script:_ValheimPlayerCapture[$pfx] = $vhState
                    }

                    $reportedCount = $null
                    if ($line -match '(?i)\bConnections\s+(\d+)\b') {
                        try { $reportedCount = [int]$Matches[1] } catch { $reportedCount = $null }
                    } elseif ($line -match '(?i)\bPlayer\s+(?:joined|connection lost)\s+server\b.*?\bnow\s+(\d+)\s+player\(s\)') {
                        try { $reportedCount = [int]$Matches[1] } catch { $reportedCount = $null }
                    }

                    if ($line -match '(?i)Got character ZDOID from\s+(.+?)\s*:\s*([-\d]+):[-\d]+') {
                        $joinedName = [string]$Matches[1]
                        $zdoId = [string]$Matches[2]
                        if (-not [string]::IsNullOrWhiteSpace($zdoId)) {
                            $vhState.Zdos[$zdoId] = $joinedName
                            $script:_ValheimPlayerCapture[$pfx] = $vhState
                        }
                        _AddSharedObservedPlayer -Prefix $pfx -PlayerName $joinedName -SharedState $script:SharedState
                        $currentNames = @(_GetSharedLatestPlayers -Prefix $pfx -SharedState $script:SharedState)
                        try { _SetObservedPlayersRuntimeState -Prefix $pfx -Names @($currentNames) -Count ([Math]::Max(1, $currentNames.Count)) -SharedState $script:SharedState } catch { }
                        if ($debugPlayers -and $script:SharedState.LogQueue) {
                            $namesText = if ($currentNames.Count -gt 0) { $currentNames -join ', ' } else { '<none>' }
                            $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][DEBUG][VHLIVE] prefix=$pfx event=join player='$joinedName' zdo='$zdoId' players=$namesText")
                        }
                    }
                    elseif ($line -match '(?i)Destroying abandoned non persistent zdo\s+([-\d]+):[-\d]+\b') {
                        $zdoId = [string]$Matches[1]
                        $leftName = ''
                        if (-not [string]::IsNullOrWhiteSpace($zdoId) -and $vhState.Zdos.ContainsKey($zdoId)) {
                            $leftName = [string]$vhState.Zdos[$zdoId]
                            $vhState.Zdos.Remove($zdoId) | Out-Null
                            $script:_ValheimPlayerCapture[$pfx] = $vhState
                        }

                        if (-not [string]::IsNullOrWhiteSpace($leftName)) {
                            _RemoveSharedObservedPlayer -Prefix $pfx -PlayerName $leftName -SharedState $script:SharedState
                            $currentNames = @(_GetSharedLatestPlayers -Prefix $pfx -SharedState $script:SharedState)
                            try { _SetObservedPlayersRuntimeState -Prefix $pfx -Names @($currentNames) -SharedState $script:SharedState } catch { }
                            if ($debugPlayers -and $script:SharedState.LogQueue) {
                                $namesText = if ($currentNames.Count -gt 0) { $currentNames -join ', ' } else { '<none>' }
                                $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][DEBUG][VHLIVE] prefix=$pfx event=leave player='$leftName' zdo='$zdoId' players=$namesText")
                            }
                        }
                    }

                    if ($null -ne $reportedCount) {
                        $vhState.LastReportedCount = $reportedCount
                        $script:_ValheimPlayerCapture[$pfx] = $vhState

                        if ($reportedCount -le 0) {
                            Set-LatestPlayersSnapshot -Prefix $pfx -Names @() -Count 0 -SharedState $script:SharedState
                            $vhState.Zdos = @{}
                            $script:_ValheimPlayerCapture[$pfx] = $vhState
                            try {
                                Set-JoinableServerRuntimeState -Prefix $pfx -SharedState $script:SharedState
                            } catch { }
                            if ($debugPlayers -and $script:SharedState.LogQueue) {
                                $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][DEBUG][VHLIVE] prefix=$pfx event=count count=0 players=<none>")
                            }
                        } else {
                            $knownPlayers = @(_GetSharedLatestPlayers -Prefix $pfx -SharedState $script:SharedState)
                            if ($knownPlayers.Count -gt 0 -and $knownPlayers.Count -le $reportedCount) {
                                Set-LatestPlayersSnapshot -Prefix $pfx -Names @($knownPlayers) -Count $reportedCount -SharedState $script:SharedState
                            } else {
                                # If the count dropped below our remembered roster, prefer the authoritative
                                # server count over stale names so the dashboard does not stick high.
                                Set-LatestPlayersSnapshot -Prefix $pfx -Names @() -Count $reportedCount -SharedState $script:SharedState
                            }
                            try {
                                $namesForState = if ($knownPlayers.Count -gt 0 -and $knownPlayers.Count -le $reportedCount) { @($knownPlayers) } else { @() }
                                _SetObservedPlayersRuntimeState -Prefix $pfx -Names $namesForState -Count $reportedCount -SharedState $script:SharedState
                            } catch { }
                            if ($debugPlayers -and $script:SharedState.LogQueue) {
                                $namesText = if ($knownPlayers.Count -gt 0 -and $knownPlayers.Count -le $reportedCount) { $knownPlayers -join ', ' } else { '<unknown>' }
                                $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][DEBUG][VHLIVE] prefix=$pfx event=count count=$reportedCount players=$namesText")
                            }
                        }
                    }
                }

                # Player list capture for Project Zomboid after !PZ players
                # (runs regardless of tab visibility)
                if ($line -match 'Players connected\s*\((\d+)\):') {
                    $reportedPzPlayers = 0
                    try { $reportedPzPlayers = [Math]::Max(0, [int]$Matches[1]) } catch { $reportedPzPlayers = 0 }

                    try {
                        Set-LatestPlayersSnapshot -Prefix $pfx -Names @() -Count $reportedPzPlayers -SharedState $script:SharedState
                    } catch { }

                    try {
                        _SyncPlayerActivityFromSnapshot -Prefix $pfx -Profile $profile -SharedState $script:SharedState -Count $reportedPzPlayers
                    } catch { }

                    try {
                        $pzRuntimeCode = ''
                        try {
                            $pzRuntime = Get-ServerRuntimeState -Prefix $pfx -SharedState $script:SharedState
                            if ($pzRuntime) { $pzRuntimeCode = [string]$pzRuntime.Code }
                        } catch { $pzRuntimeCode = '' }

                        $pzAlreadyNotified = $false
                        try {
                            if ($script:SharedState.ContainsKey('ServerStartNotified') -and $script:SharedState.ServerStartNotified) {
                                $pzAlreadyNotified = [bool]$script:SharedState.ServerStartNotified[$pfx]
                            }
                        } catch { $pzAlreadyNotified = $false }

                        if ($pzRuntimeCode.ToLowerInvariant() -ne 'online') {
                            Set-JoinableServerRuntimeState -Prefix $pfx -Force -Detail 'Project Zomboid server responded to a live players query and is joinable.' -SharedState $script:SharedState
                        }

                        if (-not $pzAlreadyNotified) {
                            try {
                                if ($script:SharedState.ContainsKey('ServerStartNotified') -and $script:SharedState.ServerStartNotified) {
                                    $script:SharedState.ServerStartNotified[$pfx] = $true
                                }
                            } catch { }

                            try {
                                Send-DiscordGameEvent -Profile $profile -Prefix $pfx -Event 'joinable' -Tag 'JOINABLE' -SharedState $script:SharedState | Out-Null
                            } catch {
                                $suppressJoinable = $false
                                try {
                                    $suppressJoinable = [bool](_ShouldSuppressDiscordLifecycleWebhook -Prefix $pfx -Tag 'JOINABLE' -SharedState $script:SharedState)
                                } catch { $suppressJoinable = $false }

                                if (-not $suppressJoinable) {
                                    try { _SendDiscordNotice (New-DiscordGameMessage -Profile $profile -Prefix $pfx -Event 'joinable') } catch { }
                                }
                            }
                        }
                    } catch { }

                    try { _MaybePromoteJoinableFromCommandResponse -Prefix $pfx -Signal 'pz-players-header' -SharedState $script:SharedState } catch { }
                }

                $reqs = $script:SharedState.PlayersRequests
                if ($reqs -and $reqs.ContainsKey($pfx)) {
                    $requestInfo = $reqs[$pfx]
                    $requestSource = 'Discord'
                    if ($requestInfo -is [hashtable] -and $requestInfo.ContainsKey('Source') -and -not [string]::IsNullOrWhiteSpace([string]$requestInfo.Source)) {
                        $requestSource = [string]$requestInfo.Source
                    }
                    $debugPlayers = $false
                    try {
                        $debugPlayers = ($script:SharedState -and $script:SharedState.Settings -and [bool]$script:SharedState.Settings.EnableDebugLogging)
                    } catch { $debugPlayers = $false }
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
                        if ($debugPlayers -and $script:SharedState.LogQueue) {
                            $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][DEBUG][PZCAP] prefix=$pfx source=$requestSource players-header expected=$expected")
                        }
                        if ($expected -eq 0) {
                            $gameName = $script:SharedState.Profiles[$pfx].GameName
                            Set-LatestPlayersSnapshot -Prefix $pfx -Names @() -Count 0 -SharedState $script:SharedState
                            if ($debugPlayers -and $script:SharedState.LogQueue) {
                                $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][DEBUG][PZCAP] prefix=$pfx finalized-empty")
                            }
                            if ($requestSource -eq 'Discord') {
                                try {
                                    Send-DiscordGameEvent -Profile $script:SharedState.Profiles[$pfx] -Prefix $pfx -Event 'players_none' -SharedState $script:SharedState | Out-Null
                                } catch {
                                    _SendDiscordNotice (New-DiscordGameMessage -Profile $script:SharedState.Profiles[$pfx] -Prefix $pfx -Event 'players_none')
                                }
                            }
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
                            if ($debugPlayers -and $script:SharedState.LogQueue) {
                                $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][DEBUG][PZCAP] prefix=$pfx captured='$name' count=$($cap.Names.Count)/$($cap.Expected)")
                            }
                        }

                        $elapsed = ((Get-Date) - $cap.Started).TotalSeconds
                        if ($cap.Names.Count -ge $cap.Expected -or $elapsed -gt 5) {
                            $gameName = $script:SharedState.Profiles[$pfx].GameName
                            $names = if ($cap.Names.Count -gt 0) { $cap.Names -join ', ' } else { 'none' }
                            Set-LatestPlayersSnapshot -Prefix $pfx -Names @($cap.Names) -Count $cap.Names.Count -SharedState $script:SharedState
                            try { _MaybePromoteJoinableFromCommandResponse -Prefix $pfx -Signal 'pz-players' -SharedState $script:SharedState } catch { }
                            if ($debugPlayers -and $script:SharedState.LogQueue) {
                                $script:SharedState.LogQueue.Enqueue("[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')][DEBUG][PZCAP] prefix=$pfx finalized names=$names elapsed=$([Math]::Round($elapsed,2))")
                            }
                            if ($requestSource -eq 'Discord') {
                                try {
                                    Send-DiscordGameEvent -Profile $script:SharedState.Profiles[$pfx] -Prefix $pfx -Event 'players_list' -Values @{ Names = $names } -SharedState $script:SharedState | Out-Null
                                } catch {
                                    _SendDiscordNotice (New-DiscordGameMessage -Profile $script:SharedState.Profiles[$pfx] -Prefix $pfx -Event 'players_list' -Values @{ Names = $names })
                                }
                            }
                            $reqs.Remove($pfx)
                            $script:_PlayersCapture.Remove($pfx)
                        } else {
                            $script:_PlayersCapture[$pfx] = $cap
                        }
                    }
                }

                # Detect server started markers and notify Discord once per session.
                # (runs regardless of tab visibility)
                #
                # Detection is keyed on PROFILE PREFIX (not GameName) so renaming
                # a game in the GUI never silently breaks the joinable notification.
                # Each check also accepts a GameName keyword as a fallback so profiles
                # that use a non-standard prefix still work.
                $profile = $script:SharedState.Profiles[$pfx]
                if ($profile -and -not $script:_ServerStartNotified.ContainsKey($pfx)) {
                    $gameName = if ($profile.GameName) { $profile.GameName } else { $pfx }

                    # Shared closure: fire the joinable webhook and mark prefix done
                    $notifyJoinable = {
                        $runtimeCodeNow = ''
                        $isRunningNow = $false

                        try {
                            $runtimeNow = _GetRuntimeStateEntry -Prefix $pfx -SharedState $script:SharedState
                            if ($runtimeNow) { $runtimeCodeNow = [string]$runtimeNow.Code }
                        } catch { $runtimeCodeNow = '' }

                        try {
                            $statusNow = Get-ServerStatus -Prefix $pfx
                            $isRunningNow = ($statusNow -and $statusNow.Running)
                        } catch { $isRunningNow = $false }

                        $runtimeCodeNow = $runtimeCodeNow.ToLowerInvariant()
                    $knownGameNow = ''
                    try { if ($profile) { $knownGameNow = (_NormalizeGameIdentity (_GetProfileKnownGame -Profile $profile)) } } catch { $knownGameNow = '' }
                    $minecraftRuntimeReadyHint = (($knownGameNow -eq 'minecraft') -and $runtimeCodeNow -in @('starting','waiting_first_player','idle_wait','online'))
                        if (((-not $isRunningNow) -and (-not $minecraftRuntimeReadyHint)) -or $runtimeCodeNow -in @('stopping','stopped','idle_shutdown','failed','blocked')) {
                            return
                        }

                        $script:_ServerStartNotified[$pfx] = $true
                        try {
                            Set-JoinableServerRuntimeState -Prefix $pfx -Force -SharedState $script:SharedState
                        } catch { }
                        try {
                            Send-DiscordGameEvent -Profile $profile -Prefix $pfx -Event 'joinable' -Tag 'JOINABLE' -SharedState $script:SharedState | Out-Null
                        } catch {
                            $suppressJoinable = $false
                            try {
                                $suppressJoinable = [bool](_ShouldSuppressDiscordLifecycleWebhook -Prefix $pfx -Tag 'JOINABLE' -SharedState $script:SharedState)
                            } catch { $suppressJoinable = $false }

                            if (-not $suppressJoinable) {
                                _SendDiscordNotice (New-DiscordGameMessage -Profile $profile -Prefix $pfx -Event 'joinable')
                            }
                        }
                    }

                    # --- Project Zomboid ---
                    # Log line: *** SERVER STARTED ***
                    if (($pfx -eq 'PZ' -or $profile.GameName -match 'Zomboid') -and $line -match 'SERVER STARTED') {
                        & $notifyJoinable
                    }
                    # --- Hytale ---
                    # Log line: [HytaleServer]          Hytale Server Booted!
                    # Match "Server Booted" broadly so the [HytaleServer] bracket prefix
                    # and variable whitespace never block the match.
                    elseif (($pfx -eq 'HY' -or $profile.GameName -match 'Hytale') -and
                            ($line -match 'Hytale Server Booted' -or $line -match 'Server Booted')) {
                        & $notifyJoinable
                    }
                    # --- Valheim ---
                    # Log line: Game server connected
                    elseif (($pfx -eq 'VH' -or $profile.GameName -match 'Valheim') -and $line -match 'Game server connected') {
                        & $notifyJoinable
                    }
                    # --- 7 Days to Die ---
                    # Log line: ...INF StartGame done...
                    elseif (($pfx -eq 'DZ' -or $profile.GameName -match '7 Days') -and $line -match 'INF StartGame done') {
                        & $notifyJoinable
                    }
                    # --- Palworld ---
                    # Joinable detection is handled by the server monitor via REST API polling.
                    # Palworld produces no log files so log-line detection never fires here.
                    # --- Minecraft ---
                    # Ready/joinable markers seen in live dedicated server logs:
                    #   Done (X.Xs)! For help, type "help"
                    #   RCON running on 0.0.0.0:25575
                    #   There are 0 of a max of 20 players online:
                    elseif (($pfx -eq 'MC' -or $profile.GameName -match 'Minecraft') -and (
                            $line -match 'For help, type' -or
                            $line -match 'RCON running on' -or
                            $line -match 'Thread RCON Listener started' -or
                            $line -match '(?i)\bthere are \d+ of a max of \d+ players online\b')) {
                        $minecraftStillRunning = $false
                        try {
                            $minecraftStillRunning = ($script:SharedState -and $script:SharedState.RunningServers -and $script:SharedState.RunningServers.ContainsKey($pfx))
                        } catch { $minecraftStillRunning = $false }
                        if ($minecraftStillRunning) {
                            & $notifyJoinable
                        }
                    }
                    # --- Satisfactory ---
                    # Log line: LogServer: Display: Server startup time elapsed and
                    #           saving/level loading is done, auto-pause is allowed...
                    # Confirmed from actual server log - this line fires exactly once
                    # when the world is fully loaded and the server is accepting players.
                    elseif (($pfx -eq 'SF' -or $profile.GameName -match 'Satisfactory') -and
                            $line -match 'Server startup time elapsed and saving/level loading is done' -and
                            $line -match 'WorldTimeSeconds\s*=\s*([0-9]+(?:\.[0-9]+)?)' -and
                            ([double]$matches[1]) -lt 45) {
                        & $notifyJoinable
                    }
                }

                $parserPerf.Stop()
                $uiPerfParserMs += $parserPerf.Elapsed.TotalMilliseconds
            }
            $uiPhasePerf.Stop()
            $uiPerfGameLogMs = $uiPhasePerf.Elapsed.TotalMilliseconds
            $uiPerfGameLogPainted = $gameLogPaintCount
            _TraceGuiPerformanceSample -Area 'GameLogQueueDrain' `
                -ElapsedMs $uiPerfGameLogMs `
                -WarnAtMs 180 `
                -DebugAtMs 60 `
                -Detail ('tick={0};dequeued={1};painted={2};budget={3}' -f $uiTickId, $uiPerfGameLogDequeued, $uiPerfGameLogPainted, $gameLogPaintBudget)
            _TraceGuiPerformanceSample -Area 'GameLogParserPass' `
                -ElapsedMs $uiPerfParserMs `
                -WarnAtMs 140 `
                -DebugAtMs 45 `
                -Detail ('tick={0};lines={1};prefixes={2}' -f $uiTickId, $uiPerfParserLines, @($uiPerfParserPrefixes.Keys).Count)

            # --- Listener control (Start/Stop/Restart) ---
            $uiPhasePerf = [System.Diagnostics.Stopwatch]::StartNew()
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
            $uiPhasePerf.Stop()
            $uiPerfListenerMs = $uiPhasePerf.Elapsed.TotalMilliseconds
            _TraceGuiPerformanceSample -Area 'ListenerUiSync' `
                -ElapsedMs $uiPerfListenerMs `
                -WarnAtMs 80 `
                -DebugAtMs 25 `
                -Detail ('tick={0};restartRequested={1};listenerRunning={2}' -f $uiTickId, `
                    $(if ($script:_ListenerRestartRequested) { '1' } else { '0' }), `
                    $(if (_ListenerIsRunning) { '1' } else { '0' }))

            # --- Bot status ---
            $isRunning = $script:SharedState.ContainsKey('ListenerRunning') -and $script:SharedState['ListenerRunning'] -eq $true
            if ($isRunning) {
                $lblBot.Text      = 'Bot: Online'
                $lblBot.ForeColor = $clrGreen
            } else {
                $lblBot.Text      = 'Bot: Offline'
                $lblBot.ForeColor = $clrRed
            }
            & $layoutTopMetrics

            # --- Lightweight dashboard refresh ---
            $dashboardDecision = _GetDashboardRefreshDecision
            if ($dashboardDecision.ShouldRun) {
                $uiPhasePerf = [System.Diagnostics.Stopwatch]::StartNew()
                _UpdateDashboardStatus
                $uiPhasePerf.Stop()
                $uiPerfDashboardMs = $uiPhasePerf.Elapsed.TotalMilliseconds
                try { $script:SharedState['LastDashboardRefreshAt'] = Get-Date } catch { }
                _TraceGuiPerformanceSample -Area 'DashboardStatusRefresh' `
                    -ElapsedMs $uiPerfDashboardMs `
                    -WarnAtMs 140 `
                    -DebugAtMs 45 `
                    -Detail ('tick={0};profiles={1};running={2};mode={3};reason={4};interval={5}' -f $uiTickId, `
                        $(if ($script:SharedState -and $script:SharedState.Profiles) { @($script:SharedState.Profiles.Keys).Count } else { 0 }), `
                        $(if ($script:SharedState -and $script:SharedState.RunningServers) { @($script:SharedState.RunningServers.Keys).Count } else { 0 }), `
                        [string]$dashboardDecision.Mode, `
                        [string]$dashboardDecision.Reason, `
                        [int]$dashboardDecision.IntervalSeconds)
            } else {
                $uiPerfDashboardMs = 0.0
                if (_IsGuiDebugEnabled) {
                    _TraceGuiPerformanceSample -Area 'DashboardStatusRefreshSkipped' `
                        -ElapsedMs ([double]$dashboardDecision.SecondsSinceActivity) `
                        -WarnAtMs 999999 `
                        -DebugAtMs 0 `
                        -Detail ('tick={0};mode={1};reason={2};interval={3}' -f $uiTickId, `
                            [string]$dashboardDecision.Mode, `
                            [string]$dashboardDecision.Reason, `
                            [int]$dashboardDecision.IntervalSeconds)
                }
            }

            # --- Status bar ---
            $rc = if ($script:SharedState.RunningServers) { $script:SharedState.RunningServers.Count } else { 0 }
            $tc = if ($script:SharedState.Profiles)       { $script:SharedState.Profiles.Count       } else { 0 }
            $eccUptime = (Get-Date) - $guiStartedAt
            $eccUptimeText = if ($eccUptime.TotalHours -ge 1) {
                '{0:D2}:{1:D2}:{2:D2}' -f [int][Math]::Floor($eccUptime.TotalHours), $eccUptime.Minutes, $eccUptime.Seconds
            } else {
                '{0:D2}:{1:D2}' -f $eccUptime.Minutes, $eccUptime.Seconds
            }
            $statusLabel.Text = "Profiles: $tc  |  Running: $rc  |  $(Get-Date -Format 'hh:mm:ss tt')  |  ECC Uptime: $eccUptimeText"

            # --- Apply latest metrics from background runspace ---
            $uiPhasePerf = [System.Diagnostics.Stopwatch]::StartNew()
            if ($script:SharedState.ContainsKey('_MetricCPU')) {
                $script:_cpuSmooth = _Smooth $script:_cpuSmooth ([double]$script:SharedState['_MetricCPU'])
                $script:_ramSmooth = _Smooth $script:_ramSmooth ([double]$script:SharedState['_MetricRAM'])
                $script:_netSmooth = _Smooth $script:_netSmooth ([double]$script:SharedState['_MetricNET'])

                $lblCPU.Text = 'CPU: {0:N0}%' -f $script:_cpuSmooth
                $lblRAM.Text = 'RAM: {0:N0}%' -f $script:_ramSmooth
                $netKbps = [Math]::Max(0.0, [double]$script:_netSmooth * 8.0)
                $lblNET.Text = if ($netKbps -ge 1000000.0) {
                    'NET: {0:N1} Gbps' -f ($netKbps / 1000000.0)
                } elseif ($netKbps -ge 1000.0) {
                    'NET: {0:N1} Mbps' -f ($netKbps / 1000.0)
                } else {
                    'NET: {0:N0} Kbps' -f $netKbps
                }
                _SetMetricColor $lblCPU $script:_cpuSmooth
                _SetMetricColor $lblRAM $script:_ramSmooth
                $lblNET.ForeColor = $clrText
            }

            $now = Get-Date
            try {
                $refreshDiskMetric = $true
                if ($script:_diskMetricSnapshotAt -is [datetime]) {
                    try { $refreshDiskMetric = (($now - $script:_diskMetricSnapshotAt).TotalSeconds -ge 15) } catch { $refreshDiskMetric = $true }
                }
                if ($refreshDiskMetric -or $null -eq $script:_diskMetricSnapshot) {
                    $script:_diskMetricSnapshot = _GetTrackedDiskMetricSnapshot -SharedState $script:SharedState
                    $script:_diskMetricSnapshotAt = $now
                }
                if ($script:_diskMetricSnapshot) {
                    $lblDisk.Text = [string]$script:_diskMetricSnapshot.Summary
                    $lblDisk.ForeColor = $script:_diskMetricSnapshot.Color
                    _SetMainControlToolTip -Control $lblDisk -Text ([string]$script:_diskMetricSnapshot.Tooltip)
                } else {
                    throw 'Disk metric snapshot was empty.'
                }
            } catch {
                $fallbackDiskRoot = ''
                try {
                    foreach ($profileEntry in @($script:SharedState.Profiles.Values)) {
                        $fallbackDiskRoot = _GetDriveRootFromPath -Path ([string]$profileEntry.FolderPath)
                        if (-not [string]::IsNullOrWhiteSpace($fallbackDiskRoot)) { break }
                    }
                } catch { $fallbackDiskRoot = '' }
                if ([string]::IsNullOrWhiteSpace($fallbackDiskRoot)) {
                    try { $fallbackDiskRoot = _GetDriveRootFromPath -Path (Get-Location).Path } catch { $fallbackDiskRoot = '' }
                }

                if (-not [string]::IsNullOrWhiteSpace($fallbackDiskRoot)) {
                    try {
                        $fallbackDrive = [System.IO.DriveInfo]::new($fallbackDiskRoot)
                        $fallbackLabel = $fallbackDiskRoot.TrimEnd('\')
                        if ($fallbackDrive.IsReady) {
                            $lblDisk.Text = 'DISK: {0} {1}' -f $fallbackLabel, (_FormatFreeSpaceText -Bytes ([double]$fallbackDrive.AvailableFreeSpace))
                            $fallbackPercentFree = if ($fallbackDrive.TotalSize -gt 0) { ([double]$fallbackDrive.AvailableFreeSpace / [double]$fallbackDrive.TotalSize) * 100.0 } else { -1 }
                            if ($fallbackPercentFree -lt 0) {
                                $lblDisk.ForeColor = $clrYellow
                            } elseif ($fallbackPercentFree -lt 10) {
                                $lblDisk.ForeColor = $clrRed
                            } elseif ($fallbackPercentFree -lt 20) {
                                $lblDisk.ForeColor = $clrYellow
                            } else {
                                $lblDisk.ForeColor = $clrGreen
                            }
                            _SetMainControlToolTip -Control $lblDisk -Text ('Fallback drive view: {0}' -f $fallbackLabel)
                        } else {
                            $lblDisk.Text = 'DISK: {0} unavailable' -f $fallbackLabel
                            $lblDisk.ForeColor = $clrYellow
                            _SetMainControlToolTip -Control $lblDisk -Text ('Fallback drive view: {0} is not ready.' -f $fallbackLabel)
                        }
                    } catch {
                        $lblDisk.Text = 'DISK: --'
                        $lblDisk.ForeColor = $clrYellow
                        _SetMainControlToolTip -Control $lblDisk -Text 'Tracked disk summary is temporarily unavailable.'
                    }
                } else {
                    $lblDisk.Text = 'DISK: --'
                    $lblDisk.ForeColor = $clrYellow
                    _SetMainControlToolTip -Control $lblDisk -Text 'Tracked disk summary is temporarily unavailable.'
                }
            }

            try {
                $playersMetric = $null
                $playersMetricIssue = ''
                $visibleCardTotal = 0
                $visibleTooltipEntries = @()
                try {
                    if ($script:_DashboardScrollPanel) {
                        foreach ($card in @($script:_DashboardScrollPanel.Controls)) {
                            if ($null -eq $card) { continue }

                            $cardPrefix = ''
                            try { $cardPrefix = [string]$card.Tag } catch { $cardPrefix = '' }
                            if ([string]::IsNullOrWhiteSpace($cardPrefix)) { continue }

                            $subtitleLabel = $null
                            try { $subtitleLabel = ($card.Controls.Find('lblSubtitle', $true) | Select-Object -First 1) } catch { $subtitleLabel = $null }
                            if (-not $subtitleLabel) { continue }

                            $subtitleText = ''
                            try { $subtitleText = [string]$subtitleLabel.Text } catch { $subtitleText = '' }
                            if ($subtitleText -notmatch '(?i)\b(\d+)\s+player\(s\)\s+online\b') { continue }

                            $count = 0
                            try { $count = [Math]::Max(0, [int]$Matches[1]) } catch { $count = 0 }
                            if ($count -le 0) { continue }

                            $visibleCardTotal += $count
                            $visibleTooltipEntries += ('{0}: {1}' -f $cardPrefix.ToUpperInvariant(), $count)
                        }
                    }
                } catch { }

                $runningServerCount = 0
                try {
                    if ($script:SharedState -and $script:SharedState.RunningServers) {
                        $runningServerCount = @($script:SharedState.RunningServers.Keys).Count
                    }
                } catch { $runningServerCount = 0 }

                if ($visibleCardTotal -le 0 -and $runningServerCount -le 0) {
                    $playersMetric = [pscustomobject]@{
                        Summary = 'PLAYERS: 0'
                        Tooltip = 'No trusted active players are detected right now.'
                        Color = $clrAccentAlt
                        TotalPlayers = 0
                        Breakdown = @()
                    }
                } elseif ($visibleCardTotal -gt 0) {
                    $visibleTooltip = 'Active players shown on the dashboard cards:'
                    if (@($visibleTooltipEntries).Count -gt 0) {
                        $visibleTooltip += [Environment]::NewLine + (@($visibleTooltipEntries) -join [Environment]::NewLine)
                    }
                    $visibleTooltip += [Environment]::NewLine + ('Total active players: {0}' -f $visibleCardTotal)

                    $playersMetric = [pscustomobject]@{
                        Summary = 'PLAYERS: {0}' -f $visibleCardTotal
                        Tooltip = [string]$visibleTooltip
                        Color = $clrGreen
                        TotalPlayers = $visibleCardTotal
                        Breakdown = @()
                    }
                }

                if (-not $playersMetric) {
                    try {
                        $playersMetric = _GetVisibleDashboardPlayersMetricSnapshot
                    } catch {
                        $playersMetricIssue = "stage=visible-snapshot msg=$($_.Exception.Message)"
                    }
                }
                if (-not $playersMetric) {
                    try {
                        $playersMetric = _GetActivePlayersMetricSnapshot -SharedState $script:SharedState
                    } catch {
                        if ([string]::IsNullOrWhiteSpace($playersMetricIssue)) {
                            $playersMetricIssue = "stage=active-snapshot msg=$($_.Exception.Message)"
                        }
                    }
                }
                if (-not $playersMetric) {
                    $playersMetric = [pscustomobject]@{
                        Summary = 'PLAYERS: 0'
                        Tooltip = 'No trusted active players are detected right now.'
                        Color = $clrAccentAlt
                        TotalPlayers = 0
                        Breakdown = @()
                    }
                }

                try {
                    $lblPlayers.Text = [string]$playersMetric.Summary
                } catch {
                    $lblPlayers.Text = 'PLAYERS: 0'
                    if ([string]::IsNullOrWhiteSpace($playersMetricIssue)) {
                        $playersMetricIssue = "stage=apply-summary msg=$($_.Exception.Message)"
                    }
                }

                try {
                    $metricColor = $playersMetric.Color
                    if ($metricColor -is [System.Drawing.Color]) {
                        $lblPlayers.ForeColor = $metricColor
                    } else {
                        $lblPlayers.ForeColor = $clrAccentAlt
                        if ([string]::IsNullOrWhiteSpace($playersMetricIssue)) {
                            $playersMetricIssue = 'stage=apply-color msg=Metric color was not a System.Drawing.Color.'
                        }
                    }
                } catch {
                    $lblPlayers.ForeColor = $clrAccentAlt
                    if ([string]::IsNullOrWhiteSpace($playersMetricIssue)) {
                        $playersMetricIssue = "stage=apply-color msg=$($_.Exception.Message)"
                    }
                }

                try {
                    _SetMainControlToolTip -Control $lblPlayers -Text ([string]$playersMetric.Tooltip)
                } catch {
                    if ([string]::IsNullOrWhiteSpace($playersMetricIssue)) {
                        $playersMetricIssue = "stage=apply-tooltip msg=$($_.Exception.Message)"
                    }
                }

                try {
                    _MaybeTracePlayersMetricSnapshot -SharedState $script:SharedState -MetricSnapshot $playersMetric
                } catch {
                    if ([string]::IsNullOrWhiteSpace($playersMetricIssue)) {
                        $playersMetricIssue = "stage=trace msg=$($_.Exception.Message)"
                    }
                }
                & $layoutTopMetrics

                if (-not [string]::IsNullOrWhiteSpace($playersMetricIssue)) {
                    try {
                        if ($script:SharedState -and $script:SharedState.LogQueue) {
                            $now = Get-Date
                            $lastAt = $null
                            $lastKey = ''
                            try { if ($script:SharedState.ContainsKey('LastPlayersMetricHeaderIssueAt')) { $lastAt = $script:SharedState['LastPlayersMetricHeaderIssueAt'] } } catch { $lastAt = $null }
                            try { if ($script:SharedState.ContainsKey('LastPlayersMetricHeaderIssueKey')) { $lastKey = [string]$script:SharedState['LastPlayersMetricHeaderIssueKey'] } } catch { $lastKey = '' }
                            if ($playersMetricIssue -ne $lastKey -or $null -eq $lastAt -or (($now - $lastAt).TotalSeconds -ge 15)) {
                                $script:SharedState['LastPlayersMetricHeaderIssueAt'] = $now
                                $script:SharedState['LastPlayersMetricHeaderIssueKey'] = $playersMetricIssue
                                $script:SharedState.LogQueue.Enqueue("[$($now.ToString('yyyy-MM-dd HH:mm:ss'))][WARN][GUI] Players metric degraded: $playersMetricIssue")
                            }
                        }
                    } catch { }
                }
            } catch {
                $lblPlayers.Text = 'PLAYERS: 0'
                $lblPlayers.ForeColor = $clrAccentAlt
                _SetMainControlToolTip -Control $lblPlayers -Text 'No trusted active players are detected right now.'
                try {
                    if ($script:SharedState -and $script:SharedState.LogQueue) {
                        $script:SharedState.LogQueue.Enqueue("[$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))][WARN][GUI] Players metric update failed: $($_.Exception.Message)")
                    }
                } catch { }
            }

            $uiPhasePerf.Stop()
            $uiPerfHeaderMs = $uiPhasePerf.Elapsed.TotalMilliseconds
            _TraceGuiPerformanceSample -Area 'HeaderMetricsRefresh' `
                -ElapsedMs $uiPerfHeaderMs `
                -WarnAtMs 160 `
                -DebugAtMs 50 `
                -Detail ('tick={0};players={1};diskRefresh={2}' -f $uiTickId, [string]$lblPlayers.Text, $(if ($refreshDiskMetric) { '1' } else { '0' }))

            try {
                if ($script:SharedState -and $script:SharedState.LogQueue) {
                    $alreadyLogged = $false
                    try { if ($script:SharedState.ContainsKey('ResizeTimerSnapshotLogged')) { $alreadyLogged = [bool]$script:SharedState['ResizeTimerSnapshotLogged'] } } catch { $alreadyLogged = $false }
                    if (-not $alreadyLogged) {
                        $cwLive = [Math]::Max(0, $form.ClientSize.Width)
                        $chLive = [Math]::Max(0, $form.ClientSize.Height)
                        $sbhLive = 0
                        try { if ($statusBar) { $sbhLive = [Math]::Max(0, $statusBar.Height) } } catch { $sbhLive = 0 }
                        try { & $layoutResizeChromeElements $cwLive $chLive $sbhLive } catch { }
                        try { if ($script:_WindowShellPanel) { $script:_WindowShellPanel.SendToBack() } } catch { }
                        try { if ($topBar) { $topBar.BringToFront() } } catch { }
                        try { if ($leftContainer) { $leftContainer.BringToFront() } } catch { }
                        try { if ($centerCol) { $centerCol.BringToFront() } } catch { }
                        try { if ($rightContainer) { $rightContainer.BringToFront() } } catch { }
                        try { if ($bottomContainer) { $bottomContainer.BringToFront() } } catch { }
                        try { if ($script:_WindowEdgeTop) { $script:_WindowEdgeTop.BringToFront() } } catch { }
                        try { if ($script:_WindowEdgeLeft) { $script:_WindowEdgeLeft.BringToFront() } } catch { }
                        try { if ($script:_WindowEdgeRight) { $script:_WindowEdgeRight.BringToFront() } } catch { }
                        try { if ($script:_WindowEdgeBottom) { $script:_WindowEdgeBottom.BringToFront() } } catch { }
                        try { if ($statusBar) { $statusBar.BringToFront() } } catch { }
                        try { if ($script:_ResizeBottomGrip) { $script:_ResizeBottomGrip.BringToFront() } } catch { }
                        try { if ($script:_ResizeLeftGrip) { $script:_ResizeLeftGrip.BringToFront() } } catch { }
                        try { if ($script:_ResizeRightGrip) { $script:_ResizeRightGrip.BringToFront() } } catch { }
                        try { if ($script:_ResizeBottomLeftGrip) { $script:_ResizeBottomLeftGrip.BringToFront() } } catch { }
                        try { if ($script:_ResizeBottomRightGrip) { $script:_ResizeBottomRightGrip.BringToFront() } } catch { }

                        $script:SharedState['ResizeTimerSnapshotLogged'] = $true
                        if (_IsGuiDebugEnabled) {
                            $script:SharedState.LogQueue.Enqueue((
                                "[{0}][INFO][GUI] RESIZE TimerSnapshot :: client={1}x{2};statusTop={3};bottomGrip={4},{5},{6},{7};rightGrip={8},{9},{10},{11};bottomRight={12},{13},{14},{15}" -f
                                (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'),
                                $form.ClientSize.Width,
                                $form.ClientSize.Height,
                                $(if ($statusBar) { $statusBar.Top } else { -1 }),
                                $(if ($script:_ResizeBottomGrip) { $script:_ResizeBottomGrip.Left } else { -1 }),
                                $(if ($script:_ResizeBottomGrip) { $script:_ResizeBottomGrip.Top } else { -1 }),
                                $(if ($script:_ResizeBottomGrip) { $script:_ResizeBottomGrip.Width } else { -1 }),
                                $(if ($script:_ResizeBottomGrip) { $script:_ResizeBottomGrip.Height } else { -1 }),
                                $(if ($script:_ResizeRightGrip) { $script:_ResizeRightGrip.Left } else { -1 }),
                                $(if ($script:_ResizeRightGrip) { $script:_ResizeRightGrip.Top } else { -1 }),
                                $(if ($script:_ResizeRightGrip) { $script:_ResizeRightGrip.Width } else { -1 }),
                                $(if ($script:_ResizeRightGrip) { $script:_ResizeRightGrip.Height } else { -1 }),
                                $(if ($script:_ResizeBottomRightGrip) { $script:_ResizeBottomRightGrip.Left } else { -1 }),
                                $(if ($script:_ResizeBottomRightGrip) { $script:_ResizeBottomRightGrip.Top } else { -1 }),
                                $(if ($script:_ResizeBottomRightGrip) { $script:_ResizeBottomRightGrip.Width } else { -1 }),
                                $(if ($script:_ResizeBottomRightGrip) { $script:_ResizeBottomRightGrip.Height } else { -1 })
                            ))
                        }
                    }
                }
            } catch { }

            _RestoreScrollPosition -Control $editorScrollPanel -Position $savedEditorScroll

            $uiTickPerf.Stop()
            _TraceGuiPerformanceSample -Area 'MainTimerTick' `
                -ElapsedMs $uiTickPerf.Elapsed.TotalMilliseconds `
                -WarnAtMs 350 `
                -DebugAtMs 120 `
                -Detail ('tick={0};tabs={1:N1};prog={2:N1}/{3};game={4:N1}/{5}/{6};parser={7:N1}/{8}/{9};listener={10:N1};dash={11:N1};header={12:N1}' -f `
                    $uiTickId, `
                    $uiPerfLogTabsMs, `
                    $uiPerfProgramLogMs, $uiPerfProgramLogCount, `
                    $uiPerfGameLogMs, $uiPerfGameLogDequeued, $uiPerfGameLogPainted, `
                    $uiPerfParserMs, $uiPerfParserLines, @($uiPerfParserPrefixes.Keys).Count, `
                    $uiPerfListenerMs, `
                    $uiPerfDashboardMs, `
                    $uiPerfHeaderMs)

        } catch {
            try {
                $now = Get-Date
                $shouldLog = $true
                if ($script:SharedState -and $script:SharedState.ContainsKey('LastGuiTimerErrorAt')) {
                    $lastGuiTimerErrorAt = $script:SharedState['LastGuiTimerErrorAt']
                    if ($lastGuiTimerErrorAt -is [datetime] -and (($now - $lastGuiTimerErrorAt).TotalSeconds -lt 5)) {
                        $shouldLog = $false
                    }
                }
                if ($shouldLog -and $script:SharedState -and $script:SharedState.LogQueue) {
                    $script:SharedState['LastGuiTimerErrorAt'] = $now
                    $script:SharedState.LogQueue.Enqueue("[$($now.ToString('yyyy-MM-dd HH:mm:ss'))][ERROR][GUI] Main timer tick failed: $($_.Exception.Message)")
                }
                if ($shouldLog) {
                    _GuiCrashBreadcrumb -Message ("Main timer tick failed: {0}" -f $_.Exception.ToString()) -Level ERROR
                }
            } catch { }
        }
    })


    $timer.Start()

    $form.add_FormClosing({
        param($sender, $e)

        $isUiReload = ($script:_UIReloadRequested -eq $true)
        if (-not $isUiReload -and -not ($script:_AppShutdownState -and $script:_AppShutdownState.AllowClose)) {
            try { $e.Cancel = $true } catch { }
            _BeginManagedApplicationShutdown
            return
        }

        _PersistWindowSettings
        $timer.Stop()
        if ($script:_ResizeHook)      { try { $script:_ResizeHook.ReleaseHandle() } catch {} }
        if ($script:SharedState) {
            $script:SharedState['StopMetricsWorker'] = $true
            $script:SharedState['StopLogTailWorker'] = $true
            if (-not $isUiReload) {
                try { $script:SharedState.Remove('ReloadWindowBounds') } catch { }
            }
        }

        if (-not $isUiReload) {
            if ($script:MetricsHandle) {
                try { $null = $script:MetricsHandle.AsyncWaitHandle.WaitOne(3000) } catch {}
                try { $script:MetricsPS.EndInvoke($script:MetricsHandle) | Out-Null } catch {}
            }
            if ($script:LogTailHandle) {
                try { $null = $script:LogTailHandle.AsyncWaitHandle.WaitOne(3000) } catch {}
                try { $script:LogTailPS.EndInvoke($script:LogTailHandle) | Out-Null } catch {}
            }
        }

        if ($script:MetricsPS)        { try { $script:MetricsPS.Dispose() } catch {} }
        if ($script:MetricsRunspace)  { try { $script:MetricsRunspace.Close(); $script:MetricsRunspace.Dispose() } catch {} }
        if ($script:LogTailPS)        { try { $script:LogTailPS.Dispose() } catch {} }
        if ($script:LogTailRunspace)  { try { $script:LogTailRunspace.Close(); $script:LogTailRunspace.Dispose() } catch {} }
        try {
            if ($script:SharedState -and $script:SharedState.ContainsKey('ShutdownDisplayQueue') -and $script:SharedState['ShutdownDisplayQueue']) {
                while ($true) {
                    $shutdownEntry = $null
                    if (-not $script:SharedState['ShutdownDisplayQueue'].TryDequeue([ref]$shutdownEntry)) { break }
                    if (-not [string]::IsNullOrWhiteSpace($shutdownEntry)) {
                        try { _WriteProgramLog $shutdownEntry } catch { }
                    }
                }
            }
        } catch { }
        _FinishManagedShutdownWorker
        $script:MetricsHandle = $null
        $script:LogTailHandle = $null
        if (-not $isUiReload) {
            $script:SharedState['StopListener'] = $true
            $script:SharedState['StopMonitor']  = $true
        }
    })

    [System.Windows.Forms.Application]::Run($form)
    $script:_UIReloadRequested = $false
}

Export-ModuleMember -Function Start-GUI, Get-ProjectZomboidSpawnerCatalogs, Get-ProjectZomboidSpawnerCatalogsFromCache, Get-MinecraftSpawnerCatalogs, Get-MinecraftSpawnerCatalogsFromCache, Get-HytaleSpawnerCatalogs, Get-HytaleSpawnerCatalogsFromCache
