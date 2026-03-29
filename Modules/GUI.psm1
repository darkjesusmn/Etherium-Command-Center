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
$script:AppVersion  = 'v0.65'
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

function _IsPerformanceDebugEnabled {
    try {
        if (-not $script:SharedState -or -not $script:SharedState.Settings) { return $false }
        if ([bool]$script:SharedState.Settings.EnablePerformanceDebugMode) { return $true }
        if ([bool]$script:SharedState.Settings.EnableDebugLogging) { return $true }
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

function _PZSharedLoadCatalogCache {
    param(
        [string]$CatalogName,
        [string]$CacheKey
    )

    $path = _PZSharedCatalogCachePath -CatalogName $CatalogName -CacheKey $CacheKey
    if (-not (Test-Path -LiteralPath $path)) { return $null }
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
    $Button.Font      = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    $Button.FlatStyle = 'Flat'
    $Button.FlatAppearance.BorderSize = 1
    $Button.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(74, 82, 110)
    $Button.BackColor = $resolvedBaseColor
    $Button.ForeColor = [System.Drawing.Color]::FromArgb(228, 234, 245)
    $Button.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $Button.Padding   = [System.Windows.Forms.Padding]::new(0, 0, 0, 2)
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
    $y = 20
    $settingsMargin = 20
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

    $lblSettingsTitle = _Label 'Bot Settings' 20 $y (& $getSettingsContentWidth) 28 $fontTitle
    $lblSettingsTitle.Anchor = 'Top,Left,Right'
    $Tab.Controls.Add($lblSettingsTitle)
    $y += 44

    $lblToken = _Label 'Bot Token (keep secret)' 20 $y (& $getSettingsContentWidth) 20
    $lblToken.Anchor = 'Top,Left,Right'
    $Tab.Controls.Add($lblToken)
    $y += 22
    $tbToken = _TextBox 20 $y 480 24 ($Settings.BotToken) $true
    $tbToken.Anchor = 'Top,Left,Right'
    $Tab.Controls.Add($tbToken)
    $y += 36

    $lblWebhook = _Label 'Webhook URL' 20 $y (& $getSettingsContentWidth) 20
    $lblWebhook.Anchor = 'Top,Left,Right'
    $Tab.Controls.Add($lblWebhook)
    $y += 22
    $tbWebhook = _TextBox 20 $y 480 24 ($Settings.WebhookUrl)
    $tbWebhook.Anchor = 'Top,Left,Right'
    $Tab.Controls.Add($tbWebhook)
    $y += 36

    $lblChannel = _Label 'Monitor Channel ID' 20 $y (& $getSettingsContentWidth) 20
    $lblChannel.Anchor = 'Top,Left,Right'
    $Tab.Controls.Add($lblChannel)
    $y += 22
    $tbChannel = _TextBox 20 $y 240 24 ($Settings.MonitorChannelId)
    $tbChannel.Anchor = 'Top,Left,Right'
    $Tab.Controls.Add($tbChannel)
    $y += 36

    $lblPrefix = _Label 'Command Prefix (default: !)' 20 $y (& $getSettingsContentWidth) 20
    $lblPrefix.Anchor = 'Top,Left,Right'
    $Tab.Controls.Add($lblPrefix)
    $y += 22
    $tbPrefix = _TextBox 20 $y 80 24 ($Settings.CommandPrefix)
    $Tab.Controls.Add($tbPrefix)
    $y += 36

    $lblPoll = _Label 'Poll Interval in seconds (default: 2)' 20 $y (& $getSettingsContentWidth) 20
    $lblPoll.Anchor = 'Top,Left,Right'
    $Tab.Controls.Add($lblPoll)
    $y += 22
    $tbPoll = _TextBox 20 $y 80 24 ([string]$Settings.PollIntervalSeconds)
    $Tab.Controls.Add($tbPoll)
    $y += 36

    $lblDebug = _Label 'Debug Logging (verbose)' 20 $y (& $getSettingsContentWidth) 20
    $lblDebug.Anchor = 'Top,Left,Right'
    $Tab.Controls.Add($lblDebug)
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
    $y += 36

    $lblPerf = _Label 'Performance Trace Mode' 20 $y (& $getSettingsContentWidth) 20
    $lblPerf.Anchor = 'Top,Left,Right'
    $Tab.Controls.Add($lblPerf)
    $y += 22
    $chkPerfTrace           = New-Object System.Windows.Forms.CheckBox
    $chkPerfTrace.Location  = [System.Drawing.Point]::new(20, $y)
    $chkPerfTrace.Size      = [System.Drawing.Size]::new(260, 20)
    $chkPerfTrace.Text      = 'Enable long-run perf tracing'
    $chkPerfTrace.ForeColor = $clrText
    $chkPerfTrace.BackColor = [System.Drawing.Color]::Transparent
    $chkPerfTrace.Font      = $fontLabel
    $chkPerfTrace.Checked   = ($Settings.EnablePerformanceDebugMode -eq $true)
    $Tab.Controls.Add($chkPerfTrace)
    $y += 24
    $perfHint = _Label 'Keeps normal logs quieter than full debug mode while enabling detailed UIPERF and LOGPERF traces for long-term lag hunts.' 20 $y 620 34
    $perfHint.ForeColor = $clrTextSoft
    $perfHint.AutoSize = $false
    $perfHint.Anchor = 'Top,Left,Right'
    $perfHint.Width = (& $getSettingsContentWidth)
    $perfHint.Height = & $measureSettingsLabelHeight $perfHint.Text $perfHint.Width $perfHint.Font 34
    $Tab.Controls.Add($perfHint)
    $y += ($perfHint.Height + 12)

    # ── Auto-Save section ─────────────────────────────────────────────────────
    $lblAutoSaveHeader = _Label 'Auto-Save Settings' 20 $y (& $getSettingsContentWidth) 22 $fontBold
    $lblAutoSaveHeader.Anchor = 'Top,Left,Right'
    $Tab.Controls.Add($lblAutoSaveHeader)
    $y += 28

    $chkAutoSave           = New-Object System.Windows.Forms.CheckBox
    $chkAutoSave.Location  = [System.Drawing.Point]::new(20, $y)
    $chkAutoSave.Size      = [System.Drawing.Size]::new(200, 20)
    $chkAutoSave.Text      = 'Enable auto-save for all games'
    $chkAutoSave.ForeColor = $clrText
    $chkAutoSave.BackColor = [System.Drawing.Color]::Transparent
    $chkAutoSave.Font      = $fontLabel
    $chkAutoSave.Checked   = ($Settings.AutoSaveEnabled -ne $false)
    $Tab.Controls.Add($chkAutoSave)
    $y += 28

    $lblAutoSave = _Label 'Auto-save interval (minutes, default 30)' 20 $y (& $getSettingsContentWidth) 20
    $lblAutoSave.Anchor = 'Top,Left,Right'
    $Tab.Controls.Add($lblAutoSave)
    $y += 22
    $tbAutoSave = _TextBox 20 $y 80 24 ([string]$(if ($Settings.AutoSaveIntervalMinutes) { $Settings.AutoSaveIntervalMinutes } else { '30' }))
    $Tab.Controls.Add($tbAutoSave)
    $y += 36

    # ── Scheduled Restart section ──────────────────────────────────────────────
    $lblSchedHeader = _Label 'Scheduled Restart Settings' 20 $y (& $getSettingsContentWidth) 22 $fontBold
    $lblSchedHeader.Anchor = 'Top,Left,Right'
    $Tab.Controls.Add($lblSchedHeader)
    $y += 28

    $chkSchedRestart           = New-Object System.Windows.Forms.CheckBox
    $chkSchedRestart.Location  = [System.Drawing.Point]::new(20, $y)
    $chkSchedRestart.Size      = [System.Drawing.Size]::new(240, 20)
    $chkSchedRestart.Text      = 'Enable scheduled restarts for all games'
    $chkSchedRestart.ForeColor = $clrText
    $chkSchedRestart.BackColor = [System.Drawing.Color]::Transparent
    $chkSchedRestart.Font      = $fontLabel
    $chkSchedRestart.Checked   = ($Settings.ScheduledRestartEnabled -ne $false)
    $Tab.Controls.Add($chkSchedRestart)
    $y += 28

    $lblSchedHours = _Label 'Restart interval (hours, default 6)' 20 $y (& $getSettingsContentWidth) 20
    $lblSchedHours.Anchor = 'Top,Left,Right'
    $Tab.Controls.Add($lblSchedHours)
    $y += 22
    $tbSchedHours = _TextBox 20 $y 80 24 ([string]$(if ($Settings.ScheduledRestartHours) { $Settings.ScheduledRestartHours } else { '6' }))
    $Tab.Controls.Add($tbSchedHours)
    $lblSchedWarn = _Label 'Warnings sent at 60, 30, 15, 10, 5, 2, 1 min before restart.' 110 ($y + 4) ([Math]::Max(220, (& $getSettingsContentWidth) - 90)) 20
    $lblSchedWarn.ForeColor = $clrTextSoft
    $lblSchedWarn.AutoSize = $false
    $lblSchedWarn.Anchor = 'Top,Left,Right'
    $lblSchedWarn.Height = & $measureSettingsLabelHeight $lblSchedWarn.Text $lblSchedWarn.Width $lblSchedWarn.Font 20
    $Tab.Controls.Add($lblSchedWarn)
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
    $saveActionsRow.Location = [System.Drawing.Point]::new(20, $y)
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
            $savedMessage = 'Settings saved. Performance trace mode applies immediately and the Discord listener will reconnect automatically.'
        }
        [System.Windows.Forms.MessageBox]::Show(
            $savedMessage,
            'Saved','OK','Information') | Out-Null
    }
    $saveBtn.Margin = [System.Windows.Forms.Padding]::new(0)
    $saveActionsRow.Controls.Add($saveBtn)

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
    $note.Width = (& $getSettingsContentWidth)
    $note.Height = & $measureSettingsLabelHeight $note.Text $note.Width $note.Font 120
    $note.Name = 'SettingsNote'
    $Tab.Controls.Add($note)

    # Keep the settings content sized and reflowed to the tab width so text doesn't clip
    $settingsTabLocal = $Tab
    $layoutSettingsTab = {
        $contentWidth = [Math]::Max(320, $settingsTabLocal.ClientSize.Width - 40)

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
            $perfHint.Location = [System.Drawing.Point]::new(20, $chkPerfTrace.Bottom + 4)
        }

        if ($lblAutoSaveHeader -is [System.Windows.Forms.Control]) {
            $lblAutoSaveHeader.Location = [System.Drawing.Point]::new(20, $perfHint.Bottom + 12)
        }
        if ($chkAutoSave -is [System.Windows.Forms.Control]) {
            $chkAutoSave.Location = [System.Drawing.Point]::new(20, $lblAutoSaveHeader.Bottom + 6)
        }
        if ($lblAutoSave -is [System.Windows.Forms.Control]) {
            $lblAutoSave.Location = [System.Drawing.Point]::new(20, $chkAutoSave.Bottom + 8)
        }
        if ($tbAutoSave -is [System.Windows.Forms.Control]) {
            $tbAutoSave.Location = [System.Drawing.Point]::new(20, $lblAutoSave.Bottom + 2)
        }

        if ($lblSchedHeader -is [System.Windows.Forms.Control]) {
            $lblSchedHeader.Location = [System.Drawing.Point]::new(20, $tbAutoSave.Bottom + 12)
        }
        if ($chkSchedRestart -is [System.Windows.Forms.Control]) {
            $chkSchedRestart.Location = [System.Drawing.Point]::new(20, $lblSchedHeader.Bottom + 6)
        }
        if ($lblSchedHours -is [System.Windows.Forms.Control]) {
            $lblSchedHours.Location = [System.Drawing.Point]::new(20, $chkSchedRestart.Bottom + 8)
        }
        if ($tbSchedHours -is [System.Windows.Forms.Control]) {
            $tbSchedHours.Location = [System.Drawing.Point]::new(20, $lblSchedHours.Bottom + 2)
        }
        if ($lblSchedWarn -is [System.Windows.Forms.Control]) {
            $lblSchedWarn.Location = [System.Drawing.Point]::new(110, $tbSchedHours.Top + 4)
            $lblSchedWarn.Width = [Math]::Max(220, $contentWidth - 90)
            $lblSchedWarn.Height = [System.Windows.Forms.TextRenderer]::MeasureText(
                [string]$lblSchedWarn.Text,
                $lblSchedWarn.Font,
                [System.Drawing.Size]::new([Math]::Max(120, $lblSchedWarn.Width), 0),
                [System.Windows.Forms.TextFormatFlags]::WordBreak
            ).Height + 6
        }

        if ($saveActionsRow -is [System.Windows.Forms.Control]) {
            $saveActionsRow.Location = [System.Drawing.Point]::new(20, [Math]::Max($tbSchedHours.Bottom + 20, $lblSchedWarn.Bottom + 16))
            $saveActionsRow.Width = $contentWidth
        }

        if ($note -is [System.Windows.Forms.Control]) {
            $note.Location = [System.Drawing.Point]::new(20, $saveActionsRow.Bottom + 16)
            $note.Width = $contentWidth
            $note.Height = [System.Windows.Forms.TextRenderer]::MeasureText(
                [string]$note.Text,
                $note.Font,
                [System.Drawing.Size]::new([Math]::Max(120, $contentWidth), 0),
                [System.Windows.Forms.TextFormatFlags]::WordBreak
            ).Height + 6
        }

        try {
            $settingsTabLocal.AutoScrollMinSize = [System.Drawing.Size]::new(0, $note.Bottom + 20)
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
        '^Remove$'          { return 'Remove the currently selected server profile from ECC.' }
        '^Reload UI$'       { return 'Rebuild the interface without resetting running servers or timers.' }
        '^Reload Bot$'      { return 'Reconnect and reload the Discord bot without touching running servers.' }
        '^Reload Commands$' { return 'Reload profiles and command catalog files from disk.' }
        '^Full Restart$'    { return 'Restart the full app. This will stop running servers.' }
        '^Settings$'        { return 'Open global ECC settings for bot, auto-save, and restart behavior.' }
        '^Send$'            { return 'Send the current message or command.' }
        '^Clear$'           { return 'Clear the current log or text panel.' }
        '^Copy$'            { return 'Copy the visible log text to the clipboard.' }
        '^Start$'           { return 'Start this server profile.' }
        '^Stop$'            { return 'Stop this server using its configured shutdown path.' }
        '^Restart$'         { return 'Restart this server using its configured save and stop rules.' }
        '^Commands$'        { return 'Open the command tools window for this server.' }
        '^Config$'          { return 'Open the detected config files for this server.' }
        '^Manager$'         { return 'Open the Hytale manager window for updater, downloader, and mod tools.' }
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

    if ($null -eq $SharedState -or -not ($SharedState -is [hashtable])) {
        throw "Start-GUI was called without a valid SharedState hashtable."
    }

    $script:SharedState = $SharedState
    $script:ProfilesDir = $ProfilesDir
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
if (-not $SharedState.ContainsKey('SatisfactoryConnectionCapture')) { $SharedState['SatisfactoryConnectionCapture'] = [hashtable]::Synchronized(@{}) }
if (-not $SharedState.ContainsKey('ValheimPlayerCapture')) { $SharedState['ValheimPlayerCapture'] = [hashtable]::Synchronized(@{}) }

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

        if ($w -lt $minWidth) { $w = $minWidth }
        if ($h -lt $minHeight) { $h = $minHeight }

        return @{
            Width  = $w
            Height = $h
            X      = $x
            Y      = $y
            HasPos = $hasPos
            State  = if ($settings.ContainsKey('WindowState')) { "$($settings.WindowState)" } else { 'Normal' }
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

        $lblHeader = _Label "Hytale Manager - $($Profile.GameName)" 12 10 520 24 $fontTitle
        $form.Controls.Add($lblHeader)

        $lblHint = _Label 'Use the updater tab for downloader/server tools and the mod tab for local jar management.' 12 38 980 18 $fontLabel
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

        $chkAutoRestart = New-Object System.Windows.Forms.CheckBox
        $chkAutoRestart.Text = 'Auto-restart after update'
        $chkAutoRestart.Size = [System.Drawing.Size]::new(190, 20)
        $chkAutoRestart.Checked = $true
        $chkAutoRestart.ForeColor = $clrText
        $chkAutoRestart.BackColor = [System.Drawing.Color]::Transparent
        $chkAutoRestart.Font = $fontLabel
        $chkAutoRestart.Margin = [System.Windows.Forms.Padding]::new(0, 4, 8, 0)
        $flowUpdatePrimaryActions.Controls.Add($chkAutoRestart)

        $lblWarn = _Label 'Checking current server state...' 12 116 ($groupUpdate.ClientSize.Width - 24) 36 $fontLabel
        $lblWarn.ForeColor = $clrYellow
        $lblWarn.Anchor = 'Top,Left,Right'
        $groupUpdate.Controls.Add($lblWarn)

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
        }

        $lblOverallStatus = _Label 'Checking file state...' 12 82 ($groupStatus.ClientSize.Width - 24) 18 $fontBold
        $lblOverallStatus.ForeColor = $clrTextSoft
        $lblOverallStatus.Anchor = 'Top,Left,Right'
        $groupStatus.Controls.Add($lblOverallStatus)

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
        foreach ($ctrl in @($btnRefreshMods, $btnToggleMod, $btnOpenModsFolder, $btnDeleteMod, $btnCheckConflicts, $btnOpenSelectedMod, $btnOpenConfigFolder, $btnLinkCurseForge, $btnCheckModUpdates, $btnUpdateSelectedMod, $btnOpenModPage, $btnGetMoreMods)) {
            $ctrl.Margin = [System.Windows.Forms.Padding]::new(0, 0, 8, 8)
            $flowModButtons.Controls.Add($ctrl)
        }

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

        $groupModNotes = New-Object System.Windows.Forms.GroupBox
        $groupModNotes.Text = 'Selected Mod Notes'
        $groupModNotes.Location = [System.Drawing.Point]::new(16, 468)
        $groupModNotes.Size = [System.Drawing.Size]::new($tabMain.ClientSize.Width - 48, [Math]::Max(150, $tabMain.ClientSize.Height - 526))
        $groupModNotes.Anchor = 'Top,Left,Right,Bottom'
        $groupModNotes.ForeColor = $clrText
        $tabMods.Controls.Add($groupModNotes)

        $lblSelectedMod = _Label 'No mod selected.' 12 22 ($groupModNotes.ClientSize.Width - 24) 34 $fontBold
        $lblSelectedMod.ForeColor = $clrTextSoft
        $lblSelectedMod.Anchor = 'Top,Left,Right'
        $lblSelectedMod.AutoEllipsis = $true
        $groupModNotes.Controls.Add($lblSelectedMod)

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
            try { return (ConvertFrom-Json (Get-Content -LiteralPath $Path -Raw) -AsHashtable) } catch { return @{} }
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

        $invokeCurseForgeApi = {
            param([string]$RelativePath)
            Invoke-RestMethod -Uri ('{0}{1}' -f $cfApiBase, $RelativePath) -Headers @{ 'x-api-key' = $cfApiKey; Accept = 'application/json' } -Method Get -ErrorAction Stop
        }.GetNewClosure()

        $getCurseForgeProject = {
            param([int]$ProjectId)
            $project = & $invokeCurseForgeApi ("/mods/{0}" -f $ProjectId)
            $files = & $invokeCurseForgeApi ("/mods/{0}/files?pageSize=20&sortDescending=true" -f $ProjectId)
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
            $response = & $invokeCurseForgeApi ("/mods/search?gameId={0}&classId={1}&searchFilter={2}&pageSize={3}" -f $cfGameId, $cfClassId, $encodedQuery, $PageSize)
            @($response.data)
        }.GetNewClosure()

        $getSelectedMod = { if ($lvMods.SelectedItems.Count -gt 0) { $lvMods.SelectedItems[0].Tag } else { $null } }.GetNewClosure()

        $refreshModList = {
            if (-not $modsFeatureReady) { $lvMods.Items.Clear(); return }
            if (-not (& $ensureModFolders)) { return }
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
            } finally {
                $lvMods.EndUpdate()
            }
        }.GetNewClosure()

        $updateModSelectionUi = {
            $selected = & $getSelectedMod
            if ($selected) {
                $lblSelectedMod.Text = '{0}  [{1}]' -f $selected.DisplayName, $selected.Status
                $txtModNotes.Text = if ($modNotes.ContainsKey($selected.FileName)) { [string]$modNotes[$selected.FileName] } else { '' }
            } else {
                $lblSelectedMod.Text = 'No mod selected.'
                $txtModNotes.Text = ''
            }
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

            $lblManual = & $newDialogLabel 'Or enter the project ID manually:' 0 6 240 18 $dialogFontLabel $dialogText
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
                    [System.Windows.Forms.MessageBox]::Show('Enter a mod name to search for.', 'Link Mod CF', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
                    return
                }

                $searchForm.UseWaitCursor = $true
                $btnSearchCurseForge.Enabled = $false
                $lvResults.BeginUpdate()
                try {
                    $lvResults.Items.Clear()
                    $results = @(& $searchCurseForgeMods -SearchQuery $query -PageSize 15)
                    foreach ($result in $results) {
                        if ($null -eq $result) { continue }
                        $author = 'Unknown'
                        if ($result.authors -and $result.authors.Count -gt 0) {
                            try { $author = [string]$result.authors[0].name } catch { $author = 'Unknown' }
                        }
                        $downloads = '0'
                        try {
                            if ($null -ne $result.downloadCount) { $downloads = ('{0:N0}' -f ([double]$result.downloadCount)) }
                        } catch { $downloads = '0' }

                        $item = New-Object System.Windows.Forms.ListViewItem([string]$result.name)
                        [void]$item.SubItems.Add($author)
                        [void]$item.SubItems.Add($downloads)
                        [void]$item.SubItems.Add([string]$result.id)
                        $item.Tag = $result
                        [void]$lvResults.Items.Add($item)
                    }

                    if ($lvResults.Items.Count -le 0) {
                        $empty = New-Object System.Windows.Forms.ListViewItem('No CurseForge results found')
                        $empty.ForeColor = $dialogSoftText
                        [void]$lvResults.Items.Add($empty)
                    }
                } catch {
                    [System.Windows.Forms.MessageBox]::Show(("CurseForge search failed.`r`n`r`n{0}" -f $_.Exception.Message), 'Link Mod CF', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
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
                        [System.Windows.Forms.MessageBox]::Show('Project ID must be a number.', 'Link Mod CF', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
                        return
                    }
                } elseif ($lvResults.SelectedItems.Count -gt 0 -and $null -ne $lvResults.SelectedItems[0].Tag) {
                    try { $projectId = [int]$lvResults.SelectedItems[0].Tag.id } catch { $projectId = 0 }
                }

                if ($projectId -le 0) {
                    [System.Windows.Forms.MessageBox]::Show('Pick a search result or enter a project ID.', 'Link Mod CF', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
                    return
                }

                $searchForm.UseWaitCursor = $true
                try {
                    $info = & $getCurseForgeProject $projectId
                    & $saveCurseForgeLink $selected $info
                    [System.Windows.Forms.MessageBox]::Show(
                        ("Linked '{0}' to CurseForge project:`r`n`r`n{1} ({2})" -f $selected.DisplayName, $info.ProjectName, $info.ProjectId),
                        'Link Mod CF',
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Information
                    ) | Out-Null
                    $searchForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
                    $searchForm.Close()
                } catch {
                    [System.Windows.Forms.MessageBox]::Show(("Failed to link CurseForge mod.`r`n`r`n{0}" -f $_.Exception.Message), 'Link Mod CF', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
                } finally {
                    $searchForm.UseWaitCursor = $false
                }
            }.GetNewClosure()

            $btnSearchCurseForge.Add_Click({ & $performSearch }.GetNewClosure())
            $txtSearch.Add_KeyDown({
                param($sender, $e)
                if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
                    & $performSearch
                    $e.SuppressKeyPress = $true
                }
            }.GetNewClosure())
            $txtProjectId.Add_TextChanged({ & $syncLinkButtonState }.GetNewClosure())
            $lvResults.Add_SelectedIndexChanged({ & $syncLinkButtonState }.GetNewClosure())
            $lvResults.Add_DoubleClick({ & $completeLink }.GetNewClosure())
            $btnCancelLink.Add_Click({ $searchForm.Close() }.GetNewClosure())
            $btnConfirmLink.Add_Click({ & $completeLink }.GetNewClosure())
            $searchForm.AcceptButton = $btnConfirmLink
            $searchForm.CancelButton = $btnCancelLink
            $searchForm.Add_Resize({ & $layoutCurseForgeLinkDialog }.GetNewClosure())
            $searchForm.Add_Shown({
                & $layoutCurseForgeLinkDialog
                & $performSearch
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
                    $info = & $getCurseForgeProject ([int]$meta.ProjectId)
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
                [System.Windows.Forms.MessageBox]::Show('Link this mod to CurseForge first.', 'Update Mod', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
                return
            }
            try {
                $meta = $cfMetadata[$selected.FileName]
                $info = & $getCurseForgeProject ([int]$meta.ProjectId)
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
                    & $showResultMessage 'Open Folder' 'The Hytale downloader folder could not be resolved.' $false
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
                [System.Windows.Forms.MessageBox]::Show(("Potential duplicates:`r`n`r`n{0}" -f ($conflicts -join "`r`n")), 'Check Conflicts', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
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
            [System.Windows.Forms.MessageBox]::Show('No obvious config folder was found for the selected mod.', 'Browse Configs', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
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
                $cfMetadata = & $loadJsonMap $cfMetadataPath
                $modNotes = & $loadJsonMap $modNotesPath
                & $refreshModList
                & $updateModSelectionUi
            } else {
                foreach ($ctrl in @($btnRefreshMods, $btnToggleMod, $btnOpenModsFolder, $btnDeleteMod, $btnCheckConflicts, $btnOpenSelectedMod, $btnOpenConfigFolder, $btnLinkCurseForge, $btnCheckModUpdates, $btnUpdateSelectedMod, $btnOpenModPage, $btnGetMoreMods, $btnSaveModNotes, $lvMods, $txtModNotes)) {
                    try { $ctrl.Enabled = $false } catch { }
                }
                & $appendLog '[WARN] Hytale mod-manager tools are disabled because the profile root could not be resolved.'
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
    $form.Size          = [System.Drawing.Size]::new($defaultWidth, $defaultHeight)
    $form.MinimumSize   = [System.Drawing.Size]::new($minWidth, $minHeight)
    $form.BackColor     = $clrBg
    $form.StartPosition = 'CenterScreen'
    $form.Icon          = [System.Drawing.SystemIcons]::Application
    # Remove the native title bar - we draw our own chrome in the top bar
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
    $form.Padding         = [System.Windows.Forms.Padding]::new($windowMargin)

    $savedWindow = _GetSavedWindowBounds
    if ($savedWindow) {
        $form.Size = [System.Drawing.Size]::new($savedWindow.Width, $savedWindow.Height)
        if ($savedWindow.HasPos) {
            $form.StartPosition = 'Manual'
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
    $windowEdgeTop.Anchor = 'Top,Left,Right'
    $windowEdgeTop.BorderStyle = 'None'
    $form.Controls.Add($windowEdgeTop)
    $script:_WindowEdgeTop = $windowEdgeTop

    $windowEdgeLeft = _Panel 0 0 2 $defaultHeight $clrEdge
    $windowEdgeLeft.Anchor = 'Top,Left,Bottom'
    $windowEdgeLeft.BorderStyle = 'None'
    $form.Controls.Add($windowEdgeLeft)
    $script:_WindowEdgeLeft = $windowEdgeLeft

    $windowEdgeRight = _Panel ($defaultWidth - 2) 0 2 $defaultHeight $clrEdge
    $windowEdgeRight.Anchor = 'Top,Right,Bottom'
    $windowEdgeRight.BorderStyle = 'None'
    $form.Controls.Add($windowEdgeRight)
    $script:_WindowEdgeRight = $windowEdgeRight

    $windowEdgeBottom = _Panel 0 ($defaultHeight - 2) $defaultWidth 2 $clrEdge
    $windowEdgeBottom.Anchor = 'Left,Right,Bottom'
    $windowEdgeBottom.BorderStyle = 'None'
    $form.Controls.Add($windowEdgeBottom)
    $script:_WindowEdgeBottom = $windowEdgeBottom

    $topBar        = _Panel $windowMargin $windowMargin ($defaultWidth - ($windowMargin * 2)) $topBarHeight $clrPanel
    $topBar.Anchor = 'Top,Left,Right'
    $form.Controls.Add($topBar)

    # -- Drag-to-move state --
    $script:_DragActive = $false
    $script:_DragOrigin = [System.Drawing.Point]::new(0, 0)

    $topBar.Add_MouseDown({
        param($s, $e)
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
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

    $lblAppTitle = _Label "Etherium Command Center $($script:AppVersion)" 18 10 340 20 $fontBold
    $lblAppTitle.ForeColor = $clrText
    $lblAppTitle.Cursor = 'Hand'
    $metricsPanel = _Panel 18 30 780 28 $clrPanelSoft
    $metricsPanel.BorderStyle = 'FixedSingle'
    $metricsPanel.Cursor = 'Hand'
    $lblCPU = _Label 'CPU: --%'          10  4 86 18 $fontBold
    $lblRAM = _Label 'RAM: --%'         120 4 88 18 $fontBold
    $lblNET = _Label 'NET: -- KB/s'     232 4 120 18 $fontBold
    $lblDisk = _Label 'DISK: --'        376 4 170 18 $fontBold
    $lblPlayers = _Label 'PLAYERS: --'  570 4 100 18 $fontBold
    $lblBot = _Label 'Bot: Unknown'     694 4 78 18 $fontBold
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
    _SetMainControlToolTip -Control $lblDisk -Text 'Shows the main tracked server drive. Hover for all tracked drives.'
    _SetMainControlToolTip -Control $lblPlayers -Text 'Shows the total trusted active player count across running servers.'

    $actionPanelWidth = 620
    $actionPanelRightOffset = 768
    $actionPanel = _Panel ($defaultWidth - $actionPanelRightOffset) 14 $actionPanelWidth 40 $clrPanelSoft
    $actionPanel.BorderStyle = 'FixedSingle'
    $topBar.Controls.Add($actionPanel)

    $topStartAllColor = [System.Drawing.Color]::FromArgb(58, 128, 94)
    $topStopAllColor = [System.Drawing.Color]::FromArgb(156, 84, 66)
    $topReloadUiColor = [System.Drawing.Color]::FromArgb(66, 112, 214)
    $topReloadBotColor = [System.Drawing.Color]::FromArgb(182, 140, 56)
    $topReloadCommandsColor = [System.Drawing.Color]::FromArgb(156, 78, 156)
    $topFullRestartColor = [System.Drawing.Color]::FromArgb(184, 72, 92)
    $topSettingsColor = [System.Drawing.Color]::FromArgb(84, 96, 128)

    $btnStartAll = _Button 'Start All' 10 4 70 30 $topStartAllColor {
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
    _SetMainControlToolTip -Control $btnStartAll -Text 'Starts every currently startable offline profile one by one with a short delay between launches.'
    $actionPanel.Controls.Add($btnStartAll)

    $btnStopAll = _Button 'Stop All' 84 4 70 30 $topStopAllColor {
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
    _SetMainControlToolTip -Control $btnStopAll -Text 'Queue a safe stop request for every profile that is currently running.'
    $actionPanel.Controls.Add($btnStopAll)

    $btnReloadUI = _Button 'Reload UI' 158 4 86 30 $topReloadUiColor {
        $script:_UIReloadRequested = $true
        $script:SharedState['ReloadUI'] = $true
        $statusLabel.Text = "Reloading UI only. Running servers and timers will be preserved."
        _QueueStatusMessage 'Reload UI requested. Preserving running servers, auto-save timers, and restart timers.'
        _SendDiscordNotice (New-DiscordSystemMessage -Event 'reload_ui')
        $form.Close()
    }
    _SetMainControlToolTip -Control $btnReloadUI
    $actionPanel.Controls.Add($btnReloadUI)

    $btnReloadBot = _Button 'Reload Bot' 248 4 86 30 $topReloadBotColor {
        $script:SharedState['RestartListener'] = $true
        $statusLabel.Text = "Reloading Discord bot. Running servers and timers are unchanged."
        _QueueStatusMessage 'Discord bot reload requested. Server timers remain intact.'
        _SendDiscordNotice (New-DiscordSystemMessage -Event 'reload_bot')
    }
    _SetMainControlToolTip -Control $btnReloadBot
    $actionPanel.Controls.Add($btnReloadBot)

    $btnReloadCommands = _Button 'Reload Commands' 338 4 100 30 $topReloadCommandsColor {
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

    $btnFullRestart = _Button 'Full Restart' 442 4 90 30 $topFullRestartColor {
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

    $btnSettings = _Button 'Settings' 536 4 74 30 $topSettingsColor {
        $settingsForm                  = New-Object System.Windows.Forms.Form
        $settingsForm.Text             = 'Settings'
        $settingsForm.Size             = [System.Drawing.Size]::new(700, 650)
        $settingsForm.MinimumSize      = [System.Drawing.Size]::new(650, 600)
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

    # -- Custom window chrome buttons (right edge of top bar) --
    # Sizing: 3 buttons x 46px wide with 2px gaps, sitting 1px from the top edge
    $chromeY = 0
    $chromeH = $topBarHeight - 1   # full height minus bottom border pixel

    $windowChromeButtonWidth = 42
    $windowChromeGap = 4
    $windowChromeMinGlyph = '_'
    $windowChromeMaxGlyph = '[]'
    $windowChromeCloseGlyph = 'X'
    $windowChromeMinBase = [System.Drawing.Color]::FromArgb(38, 50, 74)
    $windowChromeMaxBase = [System.Drawing.Color]::FromArgb(45, 58, 82)
    $windowChromeCloseBase = [System.Drawing.Color]::FromArgb(66, 42, 52)
    $windowChromeMinHover = [System.Drawing.Color]::FromArgb(86, 112, 168)
    $windowChromeMaxHover = [System.Drawing.Color]::FromArgb(88, 126, 102)
    $windowChromeCloseHover = [System.Drawing.Color]::FromArgb(186, 72, 90)
    $windowChromeButtonHeight = $chromeH
    $windowChromeButtonY = $chromeY

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

    # =====================================================================
    # THREE-COLUMN MIDDLE SECTION
    # =====================================================================
    $collapsedSize      = 28
    $headerHeight       = 34
    $bottomHeaderHeight = 30

    $script:_LeftCollapsed   = $false
    $script:_RightCollapsed  = $false
    $script:_BottomCollapsed = $false

    # Left container (Profiles)
    $leftContainer = _Panel $windowMargin ($topBarHeight + $windowMargin + 10) $leftWidth 600 $clrPanel
    $leftContainer.Anchor = 'Top,Left,Bottom'
    $form.Controls.Add($leftContainer)

    $leftHeader = _Panel 0 0 $leftWidth $headerHeight $clrPanelAlt
    $leftHeader.Anchor = 'Top,Left,Right'
    $leftContainer.Controls.Add($leftHeader)
    $script:_LeftHeader = $leftHeader

    $leftHeaderLabel = _Label 'Game Profiles' 10 8 200 18 $fontBold
    $leftHeaderLabel.ForeColor = $clrAccentAlt
    $leftHeader.Controls.Add($leftHeaderLabel)
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
    $rightContainer.Anchor = 'Top,Right,Bottom'
    $form.Controls.Add($rightContainer)

    $rightHeader = _Panel 0 0 $rightWidth $headerHeight $clrPanelAlt
    $rightHeader.Anchor = 'Top,Left,Right'
    $rightContainer.Controls.Add($rightHeader)
    $script:_RightHeader = $rightHeader

    $rightHeaderLabel = _Label 'Profile Editor' 10 8 200 18 $fontBold
    $rightHeaderLabel.ForeColor = $clrAccentAlt
    $rightHeader.Controls.Add($rightHeaderLabel)
    $rightHeaderAccent = _Panel 0 0 4 $headerHeight $clrAccent
    $rightHeaderAccent.BorderStyle = 'None'
    $rightHeader.Controls.Add($rightHeaderAccent)

    $rightBody = _Panel 0 $headerHeight $rightWidth (600 - $headerHeight) $clrPanel
    $rightBody.Anchor = 'Top,Left,Right,Bottom'
    $rightBody.BackColor = $clrPanel
    $rightContainer.Controls.Add($rightBody)
    $script:_ProfileEditorPanel = $rightBody

    # Center dashboard
    $centerCol = _Panel ($windowMargin + $leftWidth + $sideGap) ($topBarHeight + $windowMargin + 10) ($defaultWidth - $leftWidth - $rightWidth - ($sideGap * 2) - ($windowMargin * 2)) 600 $clrPanel
    $centerCol.Anchor = 'Top,Left,Right,Bottom'
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
    $bottomContainer.Anchor = 'Left,Right,Bottom'
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
    $bottomHeaderAccent = _Panel 0 0 4 $bottomHeaderHeight $clrAccent
    $bottomHeaderAccent.BorderStyle = 'None'
    $bottomHeader.Controls.Add($bottomHeaderAccent)

    # Collapse toggle handlers (click header or label)
    $leftHeader.Cursor = 'Hand'
    $leftHeaderLabel.Cursor = 'Hand'
    $toggleLeftShell = {
        $script:_LeftCollapsed = -not $script:_LeftCollapsed
        _ReflowLayout
    }.GetNewClosure()
    _BindClickHandler -Control $leftHeader -Handler $toggleLeftShell
    _BindClickHandler -Control $leftHeaderLabel -Handler $toggleLeftShell

    $rightHeader.Cursor = 'Hand'
    $rightHeaderLabel.Cursor = 'Hand'
    $toggleRightShell = {
        $script:_RightCollapsed = -not $script:_RightCollapsed
        _ReflowLayout
    }.GetNewClosure()
    _BindClickHandler -Control $rightHeader -Handler $toggleRightShell
    _BindClickHandler -Control $rightHeaderLabel -Handler $toggleRightShell

    $bottomHeader.Cursor = 'Hand'
    $bottomHeaderLabel.Cursor = 'Hand'
    $toggleBottomShell = {
        $script:_BottomCollapsed = -not $script:_BottomCollapsed
        _ReflowLayout
    }.GetNewClosure()
    _BindClickHandler -Control $bottomHeader -Handler $toggleBottomShell
    _BindClickHandler -Control $bottomHeaderLabel -Handler $toggleBottomShell

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
    $maxRtbLines = 500   # hard cap; trim fires every 50 lines over cap to amortize cost

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

        # Only trim when we have grown 50 lines past the cap, and only then do we
        # touch $rt.Lines (the expensive O(n) call).  This amortizes the cost to
        # once every 50 appends instead of once per append.
        $trimThreshold = $maxRtbLines + 50

        try {
            $rt.SelectionStart  = $rt.TextLength
            $rt.SelectionLength = 0
            $rt.SelectionColor  = $colour
            $rt.AppendText("$line`n")

            if ($currentCount -ge $trimThreshold) {
                # Now it is worth paying the Lines.Count cost
                $actualLines = $rt.Lines.Count
                if ($actualLines -gt $maxRtbLines) {
                    $removeCount = $actualLines - $maxRtbLines
                    $cutEnd = $rt.GetFirstCharIndexFromLine($removeCount)
                    if ($cutEnd -gt 0) {
                        $rt.SelectionStart  = 0
                        $rt.SelectionLength = $cutEnd
                        $rt.SelectedText    = ''
                    }
                }
                # Sync counter to reality after the trim
                $script:_RtbLineCounts[$rtKey] = $rt.Lines.Count
            }

            # Move caret to end and scroll (must come after any trim)
            $rt.SelectionStart  = $rt.TextLength
            $rt.SelectionLength = 0
            $rt.ScrollToCaret()

        } catch { }
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

    $btnSend = _Button 'Send' 0 0 90 28 $clrAccent {
        $msg = $tbSend.Text.Trim()
        if (-not $msg) { return }
        $msg = "[BOT] $msg"
        _SendDiscordNotice -Message $msg
        $tbSend.Text = ''
    }
    _SetMainControlToolTip -Control $btnSend -Text 'Send the text box message to the Discord bot output channel.'
    $btnClearDisc = _Button 'Clear' ([Math]::Max(206, $discordFooter.Width - 117)) 5 72 28 $clrMuted {
        if ($script:_DiscordLogBox) { $script:_DiscordLogBox.Clear() }
    }
    _SetMainControlToolTip -Control $btnClearDisc -Text 'Clear the visible Discord activity log in the UI.'
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
        if ($script:_ProgramLogBox) { $script:_ProgramLogBox.Clear() }
        if (-not [string]::IsNullOrEmpty($script:LogFilePath) -and (Test-Path $script:LogFilePath)) {
            $script:LogFilePos = (Get-Item $script:LogFilePath).Length
        }
    }
    _SetMainControlToolTip -Control $btnClearProg -Text 'Clear the visible program log panel without deleting log files.'
    $btnCopyProg = _Button 'Copy' 0 4 72 28 $clrPanel {
        $ok = _CopyLogText -Box $script:_ProgramLogBox
        if (-not $ok) {
            [System.Windows.Forms.MessageBox]::Show(
                'There is no Program Log text to copy yet.',
                'Copy Program Log','OK','Information') | Out-Null
        }
    }
    _SetMainControlToolTip -Control $btnCopyProg -Text 'Copy the visible program log text to the clipboard.'
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
                    "There is no log text to copy for $gameLogPrefixLocal yet.",
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
        _SetMainControlToolTip -Control $btnClearGame -Text "Clear the visible $gameLogPrefixLocal log panel."

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
        # 16 KB is enough for ~150-200 average log lines per tick, which is
        # already far more than the UI can paint (budget = 60 lines/tick).
        $maxReadBytes = 16384L   # 16 KB per file per tick

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
            if ($available -gt $maxReadBytes) {
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
                $lines  = @($notice) + @($lines | Where-Object { $_ -ne '' })
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

                        # Hard cap: never enqueue more than 50 lines per file per tick.
                        # _ReadNewLines already limits the read window, but the 16 KB
                        # window can still yield 200+ short lines on very chatty games.
                        # This final cap keeps the GameLogQueue shallow.
                        if ($lines.Count -gt 50) {
                            $dropped = $lines.Count - 50
                            $lines   = @("... [+$dropped lines not shown this tick] ...") + $lines[-50..-1]
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
            $primaryRoot = @(
                $profileDriveCounts.GetEnumerator() |
                Sort-Object -Property @{ Expression = { [int]$_.Value }; Descending = $true }, @{ Expression = { [string]$_.Key }; Descending = $false } |
                Select-Object -ExpandProperty Key -First 1
            )[0]
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

    function _ParseGeneratedConfig {
        param(
            [string]$Content,
            [string]$Extension
        )

        $supported = @('.ini','.cfg','.conf','.properties','.txt')
        $extLower = "$Extension".ToLowerInvariant()
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
                        return @{ Supported = $false; Reason = 'This Lua file contains nested tables or code values that still need the raw editor.'; Lines = $lines; Entries = @($entries) }
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

                return @{ Supported = $false; Reason = 'This Lua file has advanced syntax, so use the raw editor for now.'; Lines = $lines; Entries = @($entries) }
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

            return @{ Supported = $false; Reason = 'This file has complex lines that need the raw editor.'; Lines = $lines; Entries = @($entries) }
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

        $outLines = New-Object System.Collections.Generic.List[string]
        foreach ($entry in @($Entries)) {
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
            return (if ($Control.Checked) { 'true' } else { 'false' })
        }
        if ($Control -is [System.Windows.Forms.TextBox]) { return [string]$Control.Text }
        if ($Control -is [System.Windows.Forms.ComboBox]) {
            return if ($Control.SelectedItem) { "$($Control.SelectedItem)" } else { "$($Control.Text)" }
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
        $form.Size            = [System.Drawing.Size]::new(900, 620)
        $form.MinimumSize     = [System.Drawing.Size]::new(760, 520)
        $form.StartPosition   = 'CenterParent'
        $form.BackColor       = $clrBg
        $form.FormBorderStyle = 'Sizable'

        $header = _Label "Config Editor - $($Profile.GameName)" 10 10 500 22 $fontBold
        $header.Anchor = 'Top,Left,Right'
        $form.Controls.Add($header)

        $lblRoot = _Label "Root: $($roots[0])" 10 34 840 20 $fontLabel
        $lblRoot.Anchor = 'Top,Left,Right'
        $form.Controls.Add($lblRoot)

        $combo = $null
        if ($roots.Count -gt 1) {
            $combo = New-Object System.Windows.Forms.ComboBox
            $combo.Location = [System.Drawing.Point]::new(60, 32)
            $combo.Size     = [System.Drawing.Size]::new(680, 24)
            $combo.Anchor   = 'Top,Left,Right'
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

        $cfgTabAccent = $clrAccent
        $cfgTabPanel = $clrPanelSoft
        $cfgTabText = $clrText
        $cfgTabFont = $tabFont

        $editorTabs = New-Object System.Windows.Forms.TabControl
        $editorTabs.Location = [System.Drawing.Point]::new(270, 60)
        $editorTabs.Size = [System.Drawing.Size]::new(600, 470)
        $editorTabs.Anchor = 'Top,Left,Right,Bottom'
        $editorTabs.DrawMode = [System.Windows.Forms.TabDrawMode]::OwnerDrawFixed
        $editorTabs.SizeMode = [System.Windows.Forms.TabSizeMode]::Fixed
        $editorTabs.ItemSize = [System.Drawing.Size]::new(110, 24)
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
        $form.Controls.Add($editorTabs)

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
        $cfgFooter.Location = [System.Drawing.Point]::new(270, 536)
        $cfgFooter.Size = [System.Drawing.Size]::new(600, 36)
        $cfgFooter.Anchor = 'Left,Right,Bottom'
        $cfgFooter.BackColor = [System.Drawing.Color]::Transparent
        $form.Controls.Add($cfgFooter)

        $cfgFooterActions = New-Object System.Windows.Forms.FlowLayoutPanel
        $cfgFooterActions.Dock = 'Left'
        $cfgFooterActions.Size = [System.Drawing.Size]::new(220, 36)
        $cfgFooterActions.WrapContents = $false
        $cfgFooterActions.AutoScroll = $true
        $cfgFooterActions.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
        $cfgFooterActions.BackColor = [System.Drawing.Color]::Transparent
        $cfgFooter.Controls.Add($cfgFooterActions)

        $cfgFooterStatusHost = New-Object System.Windows.Forms.Panel
        $cfgFooterStatusHost.Dock = 'Fill'
        $cfgFooterStatusHost.BackColor = [System.Drawing.Color]::Transparent
        $cfgFooter.Controls.Add($cfgFooterStatusHost)

        $btnSave = _Button 'Save File' 0 0 100 30 $clrGreen $null
        $btnSave.Margin = [System.Windows.Forms.Padding]::new(0, 3, 10, 0)
        $cfgFooterActions.Controls.Add($btnSave)

        $btnRefresh = _Button 'Refresh' 0 0 100 30 $clrPanel $null
        $btnRefresh.Margin = [System.Windows.Forms.Padding]::new(0, 3, 0, 0)
        $cfgFooterActions.Controls.Add($btnRefresh)

        $lblStatus = _Label '' 0 8 ($cfgFooterStatusHost.ClientSize.Width - 6) 20 $fontLabel
        $lblStatus.Anchor = 'Left,Right,Top'
        $cfgFooterStatusHost.Controls.Add($lblStatus)

        $layoutConfigEditor = {
            $clientWidth = [Math]::Max(420, $form.ClientSize.Width)
            $clientHeight = [Math]::Max(320, $form.ClientSize.Height)
            $leftMargin = 10
            $rightMargin = 10
            $topContent = 60
            $bottomMargin = 12
            $gap = 10
            $footerHeight = 36
            $footerY = $clientHeight - $footerHeight - $bottomMargin
            $contentHeight = [Math]::Max(200, $footerY - $topContent - 6)
            $usableWidth = $clientWidth - $leftMargin - $rightMargin

            $listWidth = [Math]::Min(250, [Math]::Max(180, [Math]::Floor(($usableWidth - $gap) * 0.34)))
            $editorWidth = [Math]::Max(320, $usableWidth - $listWidth - $gap)
            if (($listWidth + $gap + $editorWidth) -gt $usableWidth) {
                $editorWidth = [Math]::Max(280, $usableWidth - $listWidth - $gap)
            }
            if (($listWidth + $gap + $editorWidth) -gt $usableWidth) {
                $listWidth = [Math]::Max(160, $usableWidth - $gap - $editorWidth)
            }

            $list.Location = [System.Drawing.Point]::new($leftMargin, $topContent)
            $list.Size = [System.Drawing.Size]::new($listWidth, $contentHeight)

            $editorTabs.Location = [System.Drawing.Point]::new($leftMargin + $listWidth + $gap, $topContent)
            $editorTabs.Size = [System.Drawing.Size]::new($editorWidth, $contentHeight)

            $cfgFooter.Location = [System.Drawing.Point]::new($leftMargin + $listWidth + $gap, $footerY)
            $cfgFooter.Size = [System.Drawing.Size]::new($editorWidth, $footerHeight)

            if ($combo) {
                $combo.Width = [Math]::Max(260, $clientWidth - 140)
            } else {
                $lblRoot.Width = [Math]::Max(220, $clientWidth - 20)
            }
        }.GetNewClosure()

        $allowedExt = @('.ini','.txt','.cfg','.json','.xml','.yml','.yaml','.properties','.conf','.lua')

        $state = [ordered]@{
            Root        = $roots[0]
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
            [System.Windows.Forms.MessageBox]::Show(
                "No commands found for '$($Profile.GameName)' in CommandCatalog.json.",
                'No Commands','OK','Information') | Out-Null
            return
        }

        $supportsPzItemSpawner = ((_NormalizeGameIdentity $knownGameName) -eq 'projectzomboid')

        $form                 = New-Object System.Windows.Forms.Form
        $form.Text            = "Commands - $($Profile.GameName)"
        $form.Size            = if ($supportsPzItemSpawner) { [System.Drawing.Size]::new(1380, 760) } else { [System.Drawing.Size]::new(900, 650) }
        $form.MinimumSize     = if ($supportsPzItemSpawner) { [System.Drawing.Size]::new(1260, 760) } else { [System.Drawing.Size]::new(780, 560) }
        $form.StartPosition   = 'CenterParent'
        $form.BackColor       = $clrBg
        $form.FormBorderStyle = 'Sizable'

        $lblHeader            = _Label "Commands - $($Profile.GameName)" 10 10 600 22 $fontBold
        $lblHeader.Anchor     = 'Top,Left,Right'
        $form.Controls.Add($lblHeader)

        $commandsPanelGap = if ($supportsPzItemSpawner) { 6 } else { 10 }
        $rightPanelWidth = if ($supportsPzItemSpawner) { 500 } else { 0 }
        $commandsFooterHeight = 212
        $commandsFooterBottomMargin = 10
        $commandsContentHeight = [Math]::Max(180, $form.ClientSize.Height - $commandsFooterHeight - $commandsFooterBottomMargin - 50)
        $listPanelWidth = if ($supportsPzItemSpawner) {
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

        if ($supportsPzItemSpawner) {
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

            $tabVehicles = New-Object System.Windows.Forms.TabPage
            $tabVehicles.Text = 'Vehicles'
            $tabVehicles.BackColor = $clrPanel
            $spawnTabControl.TabPages.Add($tabVehicles) | Out-Null

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
            $lblSpawnHint.Text = 'Loading local Build 42 item entries...'
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

            $btnBuildAddItem = _Button 'Build /additem' 0 0 150 28 $clrAccent $null
            $btnBuildAddItem.Margin = [System.Windows.Forms.Padding]::new(0)
            $itemActionsRow.Controls.Add($btnBuildAddItem)

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

            $spawnFooterPanel.BringToFront()
            $vehicleFooterPanel.BringToFront()
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

                    $result = Invoke-ServerCommandText -Prefix $TargetPrefix -Command 'players' -ForceRest:$false -ForceSatisfactoryApi:$false -VerboseDebug:$EnableVerboseDebug -SharedState $SharedState

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

        if ($supportsPzItemSpawner -and $spawnPanel) {
            $pzVehicleTextureManifest = $null
            try { $pzVehicleTextureManifest = _PZSharedLoadVehicleTextureMap } catch { $pzVehicleTextureManifest = $null }
            if ($null -eq $pzVehicleTextureManifest) {
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
                        $resolvedIconPath = _PZSharedResolveItemPreviewPath -ItemName ([string]$SelectedItem.ItemName) -IconName ([string]$SelectedItem.IconName) -AssetManifest (_PZSharedLoadImportedItemTextureMap)
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
                $cmd = Get-Command -Name 'Get-ProjectZomboidSpawnerCatalogsFromCache' -CommandType Function -ErrorAction Stop
                return & $cmd -Profile $CatalogProfile
            }.GetNewClosure()

            $loadPzCatalogsFull = {
                param([hashtable]$CatalogProfile)
                $cmd = Get-Command -Name 'Get-ProjectZomboidSpawnerCatalogs' -CommandType Function -ErrorAction Stop
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
                & $refreshPzVehiclesList $(if ($tbVehicleSearch) { $tbVehicleSearch.Text } else { '' })

                if (@($pzCatalogState.ItemSource).Count -eq 0) {
                    $lblSpawnHint.Text = 'Could not load local item data from the Project Zomboid install path.'
                }
                if (@($pzCatalogState.VehicleSource).Count -eq 0) {
                    $lblVehicleHint.Text = 'Could not load local vehicle data from the Project Zomboid install path.'
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
                if ($cachedItemCount -gt 0 -and $cachedVehicleCount -gt 0) {
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
                    $lblStatus.Text = 'Inserted additem command into the command box.'
                    $lblStatus.ForeColor = $cmdClrGreen
                }
            }.GetNewClosure())

            $lbItems.Add_KeyDown({
                if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
                    $cmdPreview = & $buildPzAddItemCommand
                    if ($cmdPreview) {
                        $tbCmd.Text = $cmdPreview
                        $lblStatus.Text = 'Inserted additem command into the command box.'
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
                        $lblStatus.Text = 'Inserted command with <user> placeholder. Pick a player or edit it manually below.'
                    } else {
                        $lblStatus.Text = 'Inserted additem command into the command box.'
                    }
                    $lblStatus.ForeColor = $cmdClrGreen
                } else {
                    $lblStatus.Text = 'Pick an item first.'
                    $lblStatus.ForeColor = $cmdClrRed
                }
            }.GetNewClosure())

            if ($btnInsertAddVehicle) {
                $btnInsertAddVehicle.Add_Click({
                    $cmdPreview = & $buildPzAddVehicleCommand
                    if ($cmdPreview) {
                        $tbCmd.Text = $cmdPreview
                        if ([string]::IsNullOrWhiteSpace($cmbSpawnPlayer.Text) -and -not ($cmbSpawnPlayer.SelectedItem -is [string] -and -not [string]::IsNullOrWhiteSpace($cmbSpawnPlayer.SelectedItem))) {
                            $lblStatus.Text = 'Inserted vehicle command with <user> placeholder. Pick a player or edit it manually below.'
                        } else {
                            $lblStatus.Text = 'Inserted addvehicle command into the command box.'
                        }
                        $lblStatus.ForeColor = $cmdClrGreen
                    } else {
                        $lblStatus.Text = 'Pick a vehicle first.'
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
                        $lblStatus.Text = 'Built additem command. You can review it below, then press Send.'
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
                        $lblStatus.Text = 'Pick a vehicle first.'
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
                    $cached = & (Get-Command -Name 'Get-ProjectZomboidSpawnerCatalogsFromCache' -CommandType Function -ErrorAction Stop) -Profile $PzProfile
                    $cachedItemCount = if ($cached -and $cached.PSObject.Properties.Name -contains 'Items') { @($cached.Items).Count } else { 0 }
                    $cachedVehicleCount = if ($cached -and $cached.PSObject.Properties.Name -contains 'Vehicles') { @($cached.Vehicles).Count } else { 0 }
                    if ($cachedItemCount -gt 0 -or $cachedVehicleCount -gt 0) {
                        return $cached
                    }
                    return & (Get-Command -Name 'Get-ProjectZomboidSpawnerCatalogs' -CommandType Function -ErrorAction Stop) -Profile $PzProfile
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
                        $lblSpawnHint.Text = 'Failed to load local item data.'
                        $lblVehicleHint.Text = 'Failed to load local vehicle data.'
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
            $tipText = ''
            if ($cmd.PSObject.Properties['Description'] -and $cmd.Description) {
                $tipText = [string]$cmd.Description
            } elseif ($cmd.Command) {
                $tipText = [string]$cmd.Command
            } else {
                $tipText = [string]$cmd.Label
            }
            if ($cmd.Category) { $tipText = "$tipText`nCategory: $($cmd.Category)" }
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
            $emptyCard = _Panel 14 14 ($panel.Width - 28) 86 $clrPanelSoft
            $emptyCard.Anchor = 'Top,Left,Right'
            $panel.Controls.Add($emptyCard)
            $emptyTitle = _Label 'Profile Editor' 14 14 260 22 $fontTitle
            $emptyTitle.ForeColor = $clrText
            $emptyCard.Controls.Add($emptyTitle)
            $emptySub = _Label 'Select a server to edit its profile, commands, paths, and runtime behavior.' 14 44 ($emptyCard.Width - 28) 18 $fontLabel
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

        $contentX = 16
        $fieldX = 18
        $lw  = 182
        $tw  = [Math]::Max(300, $scroll.ClientSize.Width - ($lw + 64))
        $th  = 24
        $gap = 36
        $helpGap = 18

        $editorToolTip = New-Object System.Windows.Forms.ToolTip
        $editorToolTip.AutoPopDelay = 12000
        $editorToolTip.InitialDelay = 350
        $editorToolTip.ReshowDelay  = 150
        $editorToolTip.ShowAlways   = $true

        function _FormatKeyLabel([string]$key) {
            return $key -replace '([a-z])([A-Z])', '$1 $2'
        }

        function _GetProfileFieldMeta([string]$key) {
            $map = @{
                GameName = @{
                    Label = 'Profile display name'
                    Help  = 'The friendly name ECC shows on the dashboard, in Discord, and in the profile list.'
                }
                Prefix = @{
                    Label = 'Command prefix'
                    Help  = 'Short ID used for bot commands and quick profile targeting, like PZ or VH.'
                }
                ProcessName = @{
                    Label = 'Process name to watch'
                    Help  = 'The running process ECC looks for when checking whether this server is online.'
                }
                Executable = @{
                    Label = 'Server executable'
                    Help  = 'The main launcher file ECC starts for this server. This can be an exe, bat, cmd, or jar path.'
                }
                FolderPath = @{
                    Label = 'Server folder'
                    Help  = 'Root folder ECC treats as the install location for this server.'
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
                    Help  = 'Optional ahead-of-time cache location used by games like Hytale when that launcher mode is enabled.'
                }
                LogStrategy = @{
                    Label = 'Log discovery mode'
                    Help  = 'Tells ECC how to find the active server log: one file, newest file, special PZ session logic, or Valheim user-folder logic.'
                    Options = @(
                        @{ Value = 'SingleFile';        Label = 'Single file (fixed path)' }
                        @{ Value = 'NewestFile';        Label = 'Newest file in folder' }
                        @{ Value = 'PZSessionFolder';   Label = 'Project Zomboid session folder' }
                        @{ Value = 'ValheimUserFolder'; Label = 'Valheim user log folder' }
                    )
                }
                ServerLogRoot = @{
                    Label = 'Log root folder'
                    Help  = 'Base folder ECC searches when it needs to resolve the active server log.'
                }
                ServerLogSubDir = @{
                    Label = 'Log subfolder'
                    Help  = 'Extra folder under the log root that ECC should check before looking for the live log file.'
                }
                ServerLogFile = @{
                    Label = 'Preferred log filename'
                    Help  = 'Exact log filename ECC should look for first when log discovery mode uses file matching.'
                }
                ServerLogPath = @{
                    Label = 'Resolved log path'
                    Help  = 'Direct path to the server log file when this profile uses one fixed log file instead of discovery rules.'
                }
                ServerLogNote = @{
                    Label = 'Log note'
                    Help  = 'Internal note for special log behavior. Usually this is informational and does not need regular editing.'
                }
                RestEnabled = @{
                    Label     = 'REST/API control enabled'
                    Help      = 'Turns on REST or HTTP-based control features for games that expose a server API.'
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
                    Help  = 'Authentication token, password, or API key ECC sends to the REST control endpoint.'
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
                    Help      = 'If the server crashes or exits unexpectedly, ECC can try to bring it back up for you.'
                    BoolLabel = 'Restart it automatically'
                }
                RestartDelaySeconds = @{
                    Label = 'Crash restart delay (seconds)'
                    Help  = 'How long ECC waits after a crash before trying to start the server again.'
                }
                MaxRestartsPerHour = @{
                    Label = 'Max crash restarts per hour'
                    Help  = 'Safety cap for repeated crash loops. Set this lower if you want ECC to stop hammering a broken server.'
                }
                BlockStartIfRamPercentUsed = @{
                    Label = 'Block start if RAM used is above (%)'
                    Help  = 'ECC will refuse to start this server if total system memory usage is already above this percent. Leave 0 to disable.'
                }
                BlockStartIfFreeRamBelowGB = @{
                    Label = 'Block start if free RAM is below (GB)'
                    Help  = 'ECC will refuse to start this server if the machine has less free memory than this amount. Leave 0 to disable.'
                }
                StartupTimeoutSeconds = @{
                    Label = 'Startup ready timeout (seconds)'
                    Help  = 'How long ECC waits for the server to become truly ready before marking startup as failed.'
                }
                ShutdownIfNoPlayersAfterStartupMinutes = @{
                    Label = 'Shut down if nobody joins within (minutes)'
                    Help  = 'After startup, ECC can shut the server down if nobody joins before this timer expires. Leave 0 to disable.'
                }
                ShutdownIfEmptyAfterLastPlayerLeavesMinutes = @{
                    Label = 'Shut down after last player leaves (minutes)'
                    Help  = 'Once a server becomes empty again, ECC can wait this many minutes and then shut it down. Leave 0 to disable.'
                }
                SaveMethod = @{
                    Label = 'Save method'
                    Help  = 'Which save path ECC uses before restart or shutdown.'
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
                    Help  = 'How long ECC waits after sending the save command before moving on to shutdown or restart.'
                }
                StopMethod = @{
                    Label = 'Stop method'
                    Help  = 'Which shutdown path ECC uses when you stop or restart the server.'
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
                    Help      = 'Keeps REST polling quiet when the server is offline.'
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
                    Help  = 'Main configuration directory ECC uses when it needs to read or write server config files.'
                }
                ConfigRoots = @{
                    Label = 'Additional config folders'
                    Help  = 'Extra config locations ECC should consider for this game. Leave as-is unless you know the server uses multiple config roots.'
                }
                Commands = @{
                    Label = 'Base command map'
                    Help  = 'The main command definitions ECC uses for actions like start, stop, restart, status, and players.'
                }
                ExtraCommands = @{
                    Label = 'Extra command map'
                    Help  = 'Additional custom bot or console commands for this server beyond the default control actions.'
                }
                StdinSaveCommand = @{
                    Label = 'STDIN save command'
                    Help  = 'Exact command ECC sends into the server console when save method is set to stdin.'
                }
                StdinStopCommand = @{
                    Label = 'STDIN stop command'
                    Help  = 'Exact command ECC sends into the server console when stop method is set to stdin.'
                }
                ExeHints = @{
                    Label = 'Executable hints'
                    Help  = 'Names ECC can use to recognize this server process or discover the correct launcher during detection.'
                }
                AssetFile = @{
                    Label = 'Asset package file'
                    Help  = 'Optional asset bundle or support file this profile expects to exist with the server.'
                }
                BackupDir = @{
                    Label = 'Backup folder'
                    Help  = 'Folder ECC or the game uses for backups, snapshots, or exported saves.'
                }
            }
            if ($map.ContainsKey($key)) { return $map[$key] }
            return $null
        }

        function Add-FieldHelp([string]$helpText, [int]$y, [System.Windows.Forms.Control[]]$targets) {
            if ([string]::IsNullOrWhiteSpace($helpText)) { return }

            $lblHelp = _Label $helpText ($lw + 34) $y $tw 30 $fontLabel
            $lblHelp.ForeColor = $clrTextSoft
            $lblHelp.Anchor = 'Top,Left,Right'
            $scroll.Controls.Add($lblHelp)

            foreach ($target in @($targets)) {
                if ($target) { $editorToolTip.SetToolTip($target, $helpText) }
            }
            $editorToolTip.SetToolTip($lblHelp, $helpText)
            $script:y += $helpGap
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
                $chk.Location  = [System.Drawing.Point]::new($lw + 34, $script:y)
                $chk.Size      = [System.Drawing.Size]::new(200, 20)
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
                $tb.Location    = [System.Drawing.Point]::new($lw + 34, $script:y)
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

            $tb        = _TextBox ($lw + 34) $script:y $tw $th ([string]$value) $false
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
            $cb.Location  = [System.Drawing.Point]::new($lw + 34, $script:y - 2)
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
            $previewLbl = _Label 'Generated LaunchArgs (preview)' $fieldX $script:y $lw 20 $fontBold
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
                $panelWidth = [Math]::Max(520, $scroll.ClientSize.Width - $fieldX - 26)
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
                                $fdlg = New-Object System.Windows.Forms.FolderBrowserDialog
                                $current = $targetTb.Text.Trim()
                                if ($current -and (Test-Path $current)) {
                                    $fdlg.SelectedPath = $current
                                }
                                if ($fdlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                                    $targetTb.Text = $fdlg.SelectedPath
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
                $tbCustom = _TextBox ($lw + 34) $customAnchorY $tw $th '' $false
                $tbCustom.Anchor = 'Top,Left,Right'
                if ($state.CustomArgs) { $tbCustom.Text = "$($state.CustomArgs)" }
                $scroll.Controls.Add($tbCustom)
                $script:y = $customAnchorY + $gap
                $tbCustom.Add_TextChanged({
                    if ($rebuildPreviewAction -is [scriptblock]) {
                        $null = $rebuildPreviewAction.Invoke()
                    }
                })

                $currentY = $customAnchorY
                foreach ($section in @($sectionPanels)) {
                    $sectionPanel = $null
                    if ($section -and $section.PSObject -and $section.PSObject.Properties.Name -contains 'Panel') {
                        $sectionPanel = $section.Panel
                    }
                    if (-not ($sectionPanel -is [System.Windows.Forms.Control])) { continue }
                    $sectionPanel.Location = [System.Drawing.Point]::new($fieldX, $currentY)
                    $sectionPanel.Width = $panelWidth
                    $currentY += $sectionPanel.Height + 8
                }
                if ($customLbl) {
                    $customLbl.Location = [System.Drawing.Point]::new($fieldX, $currentY)
                }
                if ($tbCustom) {
                    $tbCustom.Location = [System.Drawing.Point]::new(($lw + 34), $currentY)
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

                $fallbackHint = _Label 'Generated LaunchArgs is currently mirroring the raw LaunchArgs field because the grouped editor is not available for this profile.' $fieldX $script:y 760 18
                $fallbackHint.ForeColor = $clrYellow
                $scroll.Controls.Add($fallbackHint)
                $script:y += 20

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

        $profileHeader = _Panel $contentX $script:y ($scroll.ClientSize.Width - 36) 92 $clrPanelSoft
        $profileHeader.Anchor = 'Top,Left,Right'
        $profileHeader.BorderStyle = 'FixedSingle'
        $scroll.Controls.Add($profileHeader)

        $profileHeaderAccent = _Panel 0 0 4 92 $clrAccent
        $profileHeaderAccent.BorderStyle = 'None'
        $profileHeader.Controls.Add($profileHeaderAccent)

        $headerTitle = _Label "$($Profile.GameName)" 16 12 340 24 $fontTitle
        $headerTitle.ForeColor = $clrText
        $profileHeader.Controls.Add($headerTitle)
        $headerMeta = _Label "Prefix [$($script:_EditingPrefix)]   |   Process $($Profile.ProcessName)" 16 42 ($profileHeader.Width - 32) 18 $fontBold
        $headerMeta.ForeColor = $clrTextSoft
        $headerMeta.Anchor = 'Top,Left,Right'
        $profileHeader.Controls.Add($headerMeta)
        $headerHint = _Label 'Edit profile behavior, launch arguments, log paths, restart settings, and integrations.' 16 62 ($profileHeader.Width - 32) 16 $fontLabel
        $headerHint.ForeColor = $clrTextSoft
        $headerHint.Anchor = 'Top,Left,Right'
        $profileHeader.Controls.Add($headerHint)
        $script:y += 108

        function Add-SectionHeader([string]$title) {
            $lblSec = _Label $title $fieldX $script:y 320 22 $fontTitle
            $lblSec.ForeColor = $clrAccentAlt
            $scroll.Controls.Add($lblSec)
            $script:y += 24
            $sep = New-Object System.Windows.Forms.Panel
            $sep.Location = [System.Drawing.Point]::new($fieldX, $script:y)
            $sep.Size     = [System.Drawing.Size]::new([Math]::Max(100, $scroll.ClientSize.Width - 36), 2)
            $sep.BackColor = $clrBorder
            $sep.Anchor   = 'Top,Left,Right'
            $scroll.Controls.Add($sep)
            $script:_ProfileSeparators += $sep
            $script:y += 14
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

        $actionTop = $script:y + 12
        $actionCard = _Panel $contentX $actionTop ($scroll.ClientSize.Width - 36) 102 $clrPanelSoft
        $actionCard.Anchor = 'Top,Left,Right'
        $actionCard.BorderStyle = 'FixedSingle'
        $scroll.Controls.Add($actionCard)
        $actionTitle = _Label 'Profile Actions' 14 10 220 18 $fontBold
        $actionTitle.ForeColor = $clrAccentAlt
        $actionCard.Controls.Add($actionTitle)
        $actionHint = _Label 'Save profile changes or control the selected server directly from here.' 14 30 ($actionCard.Width - 28) 16 $fontLabel
        $actionHint.ForeColor = $clrTextSoft
        $actionHint.Anchor = 'Top,Left,Right'
        $actionCard.Controls.Add($actionHint)

        $actionButtonsRow = New-Object System.Windows.Forms.FlowLayoutPanel
        $actionButtonsRow.Location = [System.Drawing.Point]::new(14, 52)
        $actionButtonsRow.Size = [System.Drawing.Size]::new([Math]::Max(120, $actionCard.ClientSize.Width - 28), 40)
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

            $btnBrowseDetected = _Button 'Browse Manually' 154 350 144 30 $clrPanelAlt {
                $script:_DetectedServerFolderSelection = '__browse__'
                $pickForm.DialogResult = [System.Windows.Forms.DialogResult]::Retry
                $pickForm.Close()
            }
            $btnBrowseDetected.Margin = [System.Windows.Forms.Padding]::new(10, 0, 0, 0)

            $btnCancelDetected = _Button 'Cancel' 308 350 96 30 $clrMuted {
                $script:_DetectedServerFolderSelection = $null
                $pickForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                $pickForm.Close()
            }
            $btnCancelDetected.Margin = [System.Windows.Forms.Padding]::new(0)
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

        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = if ([string]::IsNullOrWhiteSpace($KnownGameName)) {
            'Select your game server folder'
        } else {
            "Select the $KnownGameName server folder"
        }
        if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
        if ([string]::IsNullOrWhiteSpace($dlg.SelectedPath)) { return $null }
        return $dlg.SelectedPath
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
            $btnCancelName = (_Button 'Cancel' 160 74 80 30 $clrMuted {
                $nameForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                $nameForm.Close()
            })
            $btnCancelName.Margin = [System.Windows.Forms.Padding]::new(0)
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
                $cfgDlg             = New-Object System.Windows.Forms.FolderBrowserDialog
                $cfgDlg.Description = 'Select your server config folder (optional)'
                if ($cfgDlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                    $configFolder = $cfgDlg.SelectedPath
                }
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
        $pickerHeight = 320

        $picker = New-Object System.Windows.Forms.Form
        $picker.Text = 'Add Game'
        $picker.Size = [System.Drawing.Size]::new(400, $pickerHeight)
        $picker.MinimumSize = [System.Drawing.Size]::new(400, $pickerHeight)
        $picker.StartPosition = 'CenterParent'
        $picker.FormBorderStyle = 'Sizable'
        $picker.MaximizeBox = $false
        $picker.MinimizeBox = $false
        $picker.BackColor = $clrBg

        $pickerTitle = _Label 'Choose a game type' 12 12 260 24 $fontTitle
        $pickerTitle.Anchor = 'Top,Left,Right'
        $picker.Controls.Add($pickerTitle)

        $pickerHint = _Label 'Pick a supported game to use a built-in profile template, or choose Custom / Other for the manual path.' 12 42 360 40 $fontLabel
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
            'Server profile ready for control and monitoring'
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
            $cardH = 146
            $card        = _Panel 8 $y $cardW $cardH $clrPanel
            $card.Anchor = 'Top,Left,Right'
            $card.Padding = [System.Windows.Forms.Padding]::new(14, 10, 14, 10)
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

            $isHytaleProfile = ((_NormalizeGameIdentity (_GetProfileKnownGame -Profile $profile)) -eq 'hytale')
            $headerReserve = if ($isHytaleProfile) { 370 } else { 240 }

            $lblName = _Label "$($profile.GameName) [$pfx]" 16 10 ([Math]::Max(220, $cardW - $headerReserve)) 24 $fontTitle
            $lblName.Name = 'lblName'
            $lblName.Tag = if ($isHytaleProfile) { 'HytaleHeader' } else { 'DefaultHeader' }
            $card.Controls.Add($lblName)

            $lblSubtitle = _Label $stateInfo.Subtitle 16 34 ([Math]::Max(220, $cardW - $headerReserve)) 18 $fontLabel
            $lblSubtitle.ForeColor = $clrTextSoft
            $lblSubtitle.Name = 'lblSubtitle'
            $lblSubtitle.Tag = if ($isHytaleProfile) { 'HytaleHeader' } else { 'DefaultHeader' }
            $card.Controls.Add($lblSubtitle)

            $lblSource = _Label $stateInfo.HealthMeta 16 52 ([Math]::Max(220, $cardW - $headerReserve)) 16
            $lblSource.ForeColor = $clrTextSoft
            $lblSource.Name = 'lblSource'
            $lblSource.Tag = if ($isHytaleProfile) { 'HytaleHeader' } else { 'DefaultHeader' }
            $card.Controls.Add($lblSource)

            $statusBadgeX = $cardW - 118
            $statusBadge = _Panel $statusBadgeX 10 96 24 $statusBg
            $statusBadge.Anchor = 'Top,Right'
            $statusBadge.Name   = 'pnlStatus'
            $card.Controls.Add($statusBadge)

            $lblStatus   = _Label $statusText 0 2 96 18 $fontBold
            $lblStatus.Anchor    = 'Top,Right'
            $lblStatus.ForeColor = $statusColor
            $lblStatus.Name      = 'lblStatus'
            $lblStatus.TextAlign = 'MiddleCenter'
            $statusBadge.Controls.Add($lblStatus)

            $lblHealth = _Label $stateInfo.HealthText $statusBadgeX 38 96 16 $fontBold
            $lblHealth.Anchor = 'Top,Right'
            $lblHealth.ForeColor = $stateInfo.HealthColor
            $lblHealth.Name = 'lblHealth'
            $lblHealth.TextAlign = 'MiddleCenter'
            $card.Controls.Add($lblHealth)

            if ($running -and $entry) {
                $up = [Math]::Round(((Get-Date) - $entry.StartTime).TotalMinutes, 1)
                $lblUptime = _Label "PID $($entry.Pid) | Uptime: ${up} min" 16 68 ($cardW - 240) 18 $fontBold
            } else {
                $lblUptime = _Label 'Server is not running' 16 68 ($cardW - 240) 18 $fontBold
            }
            $lblUptime.Name = 'lblUptime'
            $card.Controls.Add($lblUptime)

            $timerText = _BuildTimerLine -Prefix $pfx -Profile $profile -Entry $entry -SharedState $ss
            $lblTimers            = _Label $timerText 16 88 ($cardW - 260) 14
            $lblTimers.ForeColor  = $clrYellow
            $lblTimers.Font       = New-Object System.Drawing.Font('Consolas', 7.5)
            $lblTimers.Name       = 'lblTimers'
            $card.Controls.Add($lblTimers)

            $chk           = [System.Windows.Forms.CheckBox]::new()
            $chk.Text      = 'Auto-Restart'
            $chk.Location  = [System.Drawing.Point]::new(16, 114)
            $chk.Size      = [System.Drawing.Size]::new(132, 18)
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

            $restartButtonColor  = [System.Drawing.Color]::FromArgb(164, 118, 56)
            $commandsButtonColor = [System.Drawing.Color]::FromArgb(74, 118, 204)
            $configButtonColor   = [System.Drawing.Color]::FromArgb(54, 144, 148)
            $toolsButtonColor    = [System.Drawing.Color]::FromArgb(86, 124, 196)

            $btnStart     = _Button 'Start'   210 106 78 28 $clrGreen $null
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

            $btnStop     = _Button 'Stop' 296 106 78 28 $clrRed $null
            $btnStop.Name = 'btnStop'
            $btnStop.Tag = $pfx
            $btnStop.TabStop = $false
            $btnStop.Add_Click({
                $p = $this.Tag
                _RunServerOpInBackground -Prefix $p -Operation 'Stop'
                _ClearDashboardFocus -FallbackControl $script:_DashboardScrollPanel
            })
            _SetMainControlToolTip -Control $btnStop -Text ("Stop {0} using its configured shutdown path." -f $profile.GameName)
            $card.Controls.Add($btnStop)

            $btnRestart     = _Button 'Restart' 382 106 82 28 $restartButtonColor $null
            $btnRestart.Tag = $pfx
            $btnRestart.TabStop = $false
            $btnRestart.Add_Click({
                $p = $this.Tag
                _RunServerOpInBackground -Prefix $p -Operation 'Restart'
                _ClearDashboardFocus -FallbackControl $script:_DashboardScrollPanel
            })
            _SetMainControlToolTip -Control $btnRestart -Text ("Restart {0} using its configured save and stop rules." -f $profile.GameName)
            $card.Controls.Add($btnRestart)

            $btnCommands     = _Button 'Commands' 472 106 86 28 $commandsButtonColor $null
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
            $btnConfig   = _Button 'Config' 566 106 78 28 $configButtonColor $null
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
                "Open the config files detected for $($profile.GameName)."
            } else {
                "No config path is configured yet for $($profile.GameName)."
            }
            _SetMainControlToolTip -Control $btnConfig -Text $configTip
            $card.Controls.Add($btnConfig)

            if ($isHytaleProfile) {
                $btnHytaleTools = _Button 'Manager' ($statusBadgeX - 106) 10 96 24 $toolsButtonColor $null
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
        $dashboardProfileCount = @($ss.Profiles.Keys).Count
        $dashboardCardCount = 0
        try {
            if ($script:_DashboardScrollPanel) {
                $dashboardCardCount = @($script:_DashboardScrollPanel.Controls).Count
            }
        } catch { $dashboardCardCount = 0 }
        $dashboardCardScans = 0
        $dashboardMatchedCards = 0
        foreach ($pfx in @($ss.Profiles.Keys)) {
            $entry   = $null
            $running = $false

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

            if ($null -eq $script:_DashboardScrollPanel) { continue }
            foreach ($card in @($script:_DashboardScrollPanel.Controls)) {
                $dashboardCardScans++
                if ($card.Controls.Count -lt 4) { continue }
                if ([string]$card.Tag -ne $pfx) { continue }
                $dashboardMatchedCards++

                $statusLabel = ($card.Controls.Find('lblStatus', $true) | Select-Object -First 1)
                $subtitleLabel = ($card.Controls.Find('lblSubtitle', $true) | Select-Object -First 1)
                $sourceLabel = ($card.Controls.Find('lblSource', $true) | Select-Object -First 1)
                $healthLabel = ($card.Controls.Find('lblHealth', $true) | Select-Object -First 1)
                $uptimeLabel = ($card.Controls.Find('lblUptime', $true) | Select-Object -First 1)
                $timerLabel  = ($card.Controls.Find('lblTimers', $true) | Select-Object -First 1)
                $startButton = ($card.Controls.Find('btnStart', $true) | Select-Object -First 1)
                $stopButton  = ($card.Controls.Find('btnStop', $true) | Select-Object -First 1)
                $statusPanel = ($card.Controls.Find('pnlStatus', $true) | Select-Object -First 1)
                $accentPanel = ($card.Controls.Find('pnlAccent', $true) | Select-Object -First 1)
                if (-not $statusLabel -or -not $uptimeLabel -or -not $timerLabel) { continue }

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
                        $resolvedCardPlayerCounts[$pfx.ToUpperInvariant()] = [Math]::Max(0, [int]$Matches[1])
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

                    # Timer line: recompute every tick so countdown is live
                    $timerLabel.Text = _BuildTimerLine -Prefix $pfx -Profile $profile -Entry $entry -SharedState $ss
                } else {
                    $uptimeLabel.Text      = 'Server is not running'
                    $timerLabel.Text       = _BuildTimerLine -Prefix $pfx -Profile $profile -Entry $entry -SharedState $ss
                }
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

    # =====================================================================
    # LAYOUT REFLOW  (called on form resize)
    # =====================================================================
    function _ReflowLayout {
        $reflowPerf = [System.Diagnostics.Stopwatch]::StartNew()
        $cw       = $form.ClientSize.Width
        $ch       = $form.ClientSize.Height
        $sbHeight = $statusBar.Height
        $contentTop = $windowMargin + $topBarHeight + 10

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
        if ($script:_WindowEdgeBottom) {
            $script:_WindowEdgeBottom.Location = [System.Drawing.Point]::new(0, [Math]::Max(0, $ch - 2))
            $script:_WindowEdgeBottom.Size = [System.Drawing.Size]::new($cw, 2)
        }

        $topBar.Location = [System.Drawing.Point]::new($windowMargin, $windowMargin)
        $topBar.Width  = $cw - ($windowMargin * 2)
        $topBar.Height = $topBarHeight

        # Reposition top-bar right-side control groups on resize
        $actionPanel.Location       = [System.Drawing.Point]::new($topBar.Width - $actionPanelRightOffset, 14)
        $windowChromeButtonYCurrent = [Math]::Max(0, [int][Math]::Floor((($topBar.Height - 1) - $windowChromeButtonHeight) / 2))
        $btnWinMin.Location         = [System.Drawing.Point]::new($topBar.Width - (($windowChromeButtonWidth * 3) + ($windowChromeGap * 2)), $windowChromeButtonYCurrent)
        $btnWinMax.Location         = [System.Drawing.Point]::new($topBar.Width - (($windowChromeButtonWidth * 2) + $windowChromeGap), $windowChromeButtonYCurrent)
        $btnWinClose.Location       = [System.Drawing.Point]::new($topBar.Width - $windowChromeButtonWidth,  $windowChromeButtonYCurrent)

        $leftW   = if ($script:_LeftCollapsed)  { $collapsedSize } else { $leftWidth }
        $rightW  = if ($script:_RightCollapsed) { $collapsedSize } else { $rightWidth }
        $bottomH = if ($script:_BottomCollapsed){ $bottomHeaderHeight } else { $bottomLogsHeight }

        $availableShellWidth = [Math]::Max(320, $cw - ($windowMargin * 2) - ($sideGap * 2))
        $centerMinExpanded = 560
        $leftMinExpanded = 220
        $rightMinExpanded = 440

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

        $bottomY = $ch - $sbHeight - $windowMargin - $bottomH
        if ($bottomY -lt ($contentTop + 200)) { $bottomY = $contentTop + 200 }

        $bottomContainer.Location = [System.Drawing.Point]::new($windowMargin, $bottomY)
        $bottomContainer.Size     = [System.Drawing.Size]::new($cw - ($windowMargin * 2), $bottomH)

        $leftContainer.Location  = [System.Drawing.Point]::new($windowMargin, $contentTop)
        $leftContainer.Size      = [System.Drawing.Size]::new($leftW, $bottomY - $contentTop)

        $rightContainer.Location = [System.Drawing.Point]::new($cw - $rightW - $windowMargin, $contentTop)
        $rightContainer.Size     = [System.Drawing.Size]::new($rightW, $bottomY - $contentTop)

        $centerX = $windowMargin + $leftW + $sideGap
        $centerW = $cw - $leftW - $rightW - ($sideGap * 2) - ($windowMargin * 2)
        if ($centerW -lt $centerMinExpanded) { $centerW = $centerMinExpanded }
        $centerCol.Location = [System.Drawing.Point]::new($centerX, $contentTop)
        $centerCol.Size     = [System.Drawing.Size]::new($centerW, $bottomY - $contentTop)

        # Left header/body sizing
        if ($script:_LeftCollapsed) {
            $leftHeader.Location = [System.Drawing.Point]::new(0, 0)
            $leftHeader.Size     = [System.Drawing.Size]::new($leftW, $leftContainer.Height)
            $leftBody.Visible    = $false
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
            $leftHeaderLabel.Text = 'Game Profiles'
            $leftHeaderLabel.Location = [System.Drawing.Point]::new(10, 8)
            $leftHeaderLabel.Size = [System.Drawing.Size]::new($leftW - 20, $headerHeight - 10)
            $leftHeaderLabel.TextAlign = 'MiddleLeft'
        }

        # Right header/body sizing
        if ($script:_RightCollapsed) {
            $rightHeader.Location = [System.Drawing.Point]::new(0, 0)
            $rightHeader.Size     = [System.Drawing.Size]::new($rightW, $rightContainer.Height)
            $rightBody.Visible    = $false
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
            $rightHeaderLabel.Text = 'Profile Editor'
            $rightHeaderLabel.Location = [System.Drawing.Point]::new(10, 8)
            $rightHeaderLabel.Size = [System.Drawing.Size]::new($rightW - 20, $headerHeight - 10)
            $rightHeaderLabel.TextAlign = 'MiddleLeft'
        }

        # Bottom header/body - Dock handles sizing, just toggle visibility for collapse
        if ($script:_BottomCollapsed) {
            $bottomHeader.Height = $bottomContainer.Height   # expand header to fill when collapsed
            $bottomPanel.Visible = $false
        } else {
            $bottomHeader.Height = $bottomHeaderHeight
            $bottomPanel.Visible = $true
        }

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
                        $twNew = [Math]::Max(300, $ctrl.Width - (170 + 40))
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

        # Dashboard card widths - iterate the scroll panel's direct children
        $dashScroll = $script:_DashboardScrollPanel
        if ($dashScroll) {
            foreach ($card in @($dashScroll.Controls)) {
                if ($card -is [System.Windows.Forms.Panel]) {
                    $card.Width = $dashScroll.ClientSize.Width - 20
                    $statusPanel = ($card.Controls.Find('pnlStatus', $true) | Select-Object -First 1)
                    $nameLabel   = ($card.Controls.Find('lblName', $true) | Select-Object -First 1)
                    $subtitleLbl = ($card.Controls.Find('lblSubtitle', $true) | Select-Object -First 1)
                    $sourceLbl   = ($card.Controls.Find('lblSource', $true) | Select-Object -First 1)
                    $healthLbl   = ($card.Controls.Find('lblHealth', $true) | Select-Object -First 1)
                    $uptimeLabel = ($card.Controls.Find('lblUptime', $true) | Select-Object -First 1)
                    $timerLabel  = ($card.Controls.Find('lblTimers', $true) | Select-Object -First 1)
                    $hytaleBtn   = ($card.Controls.Find('btnHytaleTools', $true) | Select-Object -First 1)

                    $statusPanelX = $card.Width - 118
                    if ($statusPanel) {
                        $statusPanel.Location = [System.Drawing.Point]::new($statusPanelX, 10)
                    }
                    if ($healthLbl) {
                        $healthLbl.Location = [System.Drawing.Point]::new($statusPanelX, 38)
                    }

                    $isHytaleCard = $false
                    if ($hytaleBtn) {
                        $isHytaleCard = $true
                        $hytaleBtn.Location = [System.Drawing.Point]::new($statusPanelX - 128, 10)
                    }

                    $headerReserve = if ($isHytaleCard) { 370 } else { 240 }
                    $headerWidth = [Math]::Max(220, $card.Width - $headerReserve)
                    if ($nameLabel) { $nameLabel.Width = $headerWidth }
                    if ($subtitleLbl) { $subtitleLbl.Width = $headerWidth }
                    if ($sourceLbl) { $sourceLbl.Width = $headerWidth }
                    if ($uptimeLabel) { $uptimeLabel.Width = [Math]::Max(220, $card.Width - 240) }
                    if ($timerLabel) { $timerLabel.Width = [Math]::Max(200, $card.Width - 260) }
                }
            }
        }

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
            $script:_GuiTickSequence = [int]$script:_GuiTickSequence + 1
            $uiTickId = [int]$script:_GuiTickSequence
            $uiTickPerf = [System.Diagnostics.Stopwatch]::StartNew()
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
            for ($i = 0; $i -lt 20; $i++) {
                $item = $null
                if (-not $script:SharedState.LogQueue.TryDequeue([ref]$item)) { break }
                $uiPerfProgramLogCount++
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

            # --- Drain GameLogQueue (UI only) ---
            # We process up to 300 lines per tick from the queue.
            # RTB painting is budget-capped (60 painted lines per tick) to keep
            # the UI responsive, but detection logic (start-notify, players-capture)
            # runs on EVERY dequeued line regardless of paint budget or tab visibility.
            # This ensures Valheim "Game server connected" and 7DTD "INF StartGame done"
            # are never silently dropped just because another game's tab is filling up.
            $uiPhasePerf = [System.Diagnostics.Stopwatch]::StartNew()
            $gameLogPaintBudget = 60
            $gameLogPaintCount  = 0
            $activeTabPage = $null
            try { $activeTabPage = $script:_LogTabControl.SelectedTab } catch { }

            for ($i = 0; $i -lt 300; $i++) {
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
                    while ($tabEntry.PendingLines.Count -gt 80) {
                        $tabEntry.PendingLines.RemoveAt(0)
                    }
                }

                # Project Zomboid and 7DTD live player tracking from normal server logs.
                # This keeps idle rules in sync even without a manual players command.
                $parserPerf = [System.Diagnostics.Stopwatch]::StartNew()
                $uiPerfParserLines++
                try { $uiPerfParserPrefixes[[string]$pfx] = $true } catch { }
                $profile = $script:SharedState.Profiles[$pfx]
                $isPzProfile = ($profile -and ((_NormalizeGameIdentity (_GetProfileKnownGame -Profile $profile)) -eq 'projectzomboid'))
                $is7dProfile = ($profile -and ((_NormalizeGameIdentity (_GetProfileKnownGame -Profile $profile)) -eq '7daystodie'))
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
                elseif ($profile -and ((_NormalizeGameIdentity (_GetProfileKnownGame -Profile $profile)) -eq 'hytale')) {
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
                        if ((-not $isRunningNow) -or $runtimeCodeNow -in @('stopping','stopped','idle_shutdown','failed','blocked')) {
                            return
                        }

                        $script:_ServerStartNotified[$pfx] = $true
                        try {
                            Set-JoinableServerRuntimeState -Prefix $pfx -SharedState $script:SharedState
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
                    # Log line: Done (X.Xs)! For help, type "help"
                    elseif (($pfx -eq 'MC' -or $profile.GameName -match 'Minecraft') -and $line -match 'For help, type') {
                        & $notifyJoinable
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

            # --- Lightweight dashboard refresh ---
            $uiPhasePerf = [System.Diagnostics.Stopwatch]::StartNew()
            _UpdateDashboardStatus
            $uiPhasePerf.Stop()
            $uiPerfDashboardMs = $uiPhasePerf.Elapsed.TotalMilliseconds
            _TraceGuiPerformanceSample -Area 'DashboardStatusRefresh' `
                -ElapsedMs $uiPerfDashboardMs `
                -WarnAtMs 140 `
                -DebugAtMs 45 `
                -Detail ('tick={0};profiles={1};running={2}' -f $uiTickId, `
                    $(if ($script:SharedState -and $script:SharedState.Profiles) { @($script:SharedState.Profiles.Keys).Count } else { 0 }), `
                    $(if ($script:SharedState -and $script:SharedState.RunningServers) { @($script:SharedState.RunningServers.Keys).Count } else { 0 }))

            # --- Status bar ---
            $rc = if ($script:SharedState.RunningServers) { $script:SharedState.RunningServers.Count } else { 0 }
            $tc = if ($script:SharedState.Profiles)       { $script:SharedState.Profiles.Count       } else { 0 }
            $statusLabel.Text = "Profiles: $tc  |  Running: $rc  |  $(Get-Date -Format 'HH:mm:ss')"

            # --- Apply latest metrics from background runspace ---
            $uiPhasePerf = [System.Diagnostics.Stopwatch]::StartNew()
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
            } catch { }
        }
    })


    $timer.Start()

    $form.add_FormClosing({
        _PersistWindowSettings
        $timer.Stop()
        if ($script:_ResizeHook)      { try { $script:_ResizeHook.ReleaseHandle() } catch {} }
        if ($script:SharedState) {
            $script:SharedState['StopMetricsWorker'] = $true
            $script:SharedState['StopLogTailWorker'] = $true
        }

        $isUiReload = ($script:_UIReloadRequested -eq $true)
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

Export-ModuleMember -Function Start-GUI, Get-ProjectZomboidSpawnerCatalogs, Get-ProjectZomboidSpawnerCatalogsFromCache
