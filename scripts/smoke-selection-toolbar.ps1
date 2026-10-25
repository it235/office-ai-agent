$ErrorActionPreference = 'Stop'

# E4-01 selection toolbar smoke.
# Source-level invariants (single catalog, no second main chain, privacy) + behaviour tests
# for the pure geometry module, compiled in-place so no Word/COM instance is needed.

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')

function Read-Source([string]$relativePath) {
    $path = Join-Path $repoRoot $relativePath
    if (-not (Test-Path -LiteralPath $path)) {
        throw "Expected file is missing: $relativePath"
    }
    return Get-Content -LiteralPath $path -Raw -Encoding UTF8
}

$catalogPath = 'ShareRibbon\Agent\QuickActionCatalog.vb'
$catalog = Read-Source $catalogPath
$geometry = Read-Source 'WordAi\Selection\WordSelectionToolbarGeometry.vb'
$toolbar = Read-Source 'WordAi\Selection\WordSelectionToolbar.vb'
$controller = Read-Source 'WordAi\Selection\WordSelectionToolbarController.vb'
$thisAddIn = Read-Source 'WordAi\ThisAddIn.vb'
$chatControl = Read-Source 'ShareRibbon\Controls\BaseChatControl.vb'
$settings = Read-Source 'ShareRibbon\Config\ChatSettings.vb'
$onboarding = Read-Source 'ShareRibbon\Controls\Services\OnboardingService.vb'
$wordRegistry = Read-Source 'WordAi\Services\WordCapabilityRegistry.vb'
$shareProj = Read-Source 'ShareRibbon\ShareRibbon.vbproj'
$wordProj = Read-Source 'WordAi\WordAi.vbproj'

# --- 1. one catalog, one namespace that VB can actually resolve -------------

if ($catalog -notmatch '(?m)^\s*Namespace Agent\s*$') {
    throw 'QuickActionCatalog must live in Namespace Agent (ShareRibbon root).'
}
if ($catalog -match 'Namespace\s+Agent\.Entry') {
    throw 'Namespace Agent.Entry is forbidden: VB identifiers are case-insensitive, ' +
          'so the namespace hijacks the "entry" loop variable already used in ShareRibbon.Agent.*'
}
foreach ($compile in @('Agent\QuickActionCatalog.vb')) {
    if ($shareProj -notmatch [regex]::Escape($compile)) {
        throw "ShareRibbon.vbproj does not compile $compile"
    }
}
foreach ($compile in @('Selection\WordSelectionToolbarGeometry.vb',
                       'Selection\WordSelectionToolbar.vb',
                       'Selection\WordSelectionToolbarController.vb')) {
    if ($wordProj -notmatch [regex]::Escape($compile)) {
        throw "WordAi.vbproj does not compile $compile"
    }
}

# --- 2. action list: 3 word actions, every CapabilityId is a real capability --

# The catalog also carries the Excel/PowerPoint entry actions added by E4-02 and, since E6-01, the
# Ribbon button wording; this smoke only owns the Word selection list, so the id universe is read from
# the _wordActions block itself and not from "every id that starts with word.", which would let a ribbon
# button silently become a fourth toolbar row.
$wordBlockStart = $catalog.IndexOf('ReadOnly _wordActions ')
if ($wordBlockStart -lt 0) { throw 'QuickActionCatalog lost its _wordActions block' }
# Any access modifier ends the block: the next field is _emptyActions, then the E6-01 ribbon utterances.
$wordBlockEnd = $catalog.IndexOf('Shared ReadOnly', $wordBlockStart + 1)
if ($wordBlockEnd -lt 0) { $wordBlockEnd = $catalog.Length }
$wordActionsBlock = $catalog.Substring($wordBlockStart, $wordBlockEnd - $wordBlockStart)
$actionIds = [regex]::Matches($wordActionsBlock, '\.Id\s*=\s*"(?<id>word\.[^"]+)"') | ForEach-Object { $_.Groups['id'].Value }
if ($actionIds.Count -ne 3) {
    throw "E4-01 expects exactly 3 selection actions, found $($actionIds.Count): $($actionIds -join ', ')"
}
foreach ($id in $actionIds) {
    if ($id -notmatch '^word\.selection\.[a-z]+$') {
        throw "Selection action id must be namespaced word.selection.*: $id"
    }
}

$capabilityIds = [regex]::Matches($catalog, '\.CapabilityId\s*=\s*"(?<cap>word\.[^"]+)"') | ForEach-Object { $_.Groups['cap'].Value }
foreach ($capabilityId in $capabilityIds) {
    if ($wordRegistry -notmatch [regex]::Escape("""$capabilityId""")) {
        throw "Quick action declares capability '$capabilityId' but WordCapabilityRegistry has no such id."
    }
}
Write-Host "Catalog actions: $($actionIds -join ', ') (declared capabilities: $($capabilityIds -join ', '))"

# --- 3. single source of wording: onboarding examples come from the catalog ---

if ($onboarding -notmatch 'QuickActionCatalog\.SampleTasks') {
    throw 'OnboardingService must read example tasks from QuickActionCatalog, not keep its own list.'
}
$sampleMatch = [regex]::Match($catalog, '(?s)Private Shared ReadOnly _wordSamples.*?Case "excel"')
if (-not $sampleMatch.Success) {
    throw 'QuickActionCatalog lost the per-host sample task lists.'
}
foreach ($sample in [regex]::Matches($sampleMatch.Value, '"(?<s>[^"]+)"') | ForEach-Object { $_.Groups['s'].Value }) {
    $hostFiles = @($onboarding, $thisAddIn, $chatControl)
    if (($hostFiles -join "`n") -match [regex]::Escape($sample)) {
        throw "Sample wording is duplicated outside the catalog: $sample"
    }
}

# --- 4. timing + lifecycle invariants ---------------------------------------

$debounceMatch = [regex]::Match($controller, 'DebounceMs\s+As\s+Integer\s*=\s*(?<ms>\d+)')
if (-not $debounceMatch.Success) {
    throw 'Controller must keep a named DebounceMs constant.'
}
$debounceMs = [int]$debounceMatch.Groups['ms'].Value
if ($debounceMs -le 0 -or $debounceMs -gt 280) {
    throw "DebounceMs = $debounceMs cannot meet the 300ms acceptance window."
}

if ($controller -notmatch '(?m)^\s*If\s+_bound\s+Then\s+Return') {
    throw 'Controller Initialize must be idempotent (ghost text already binds WindowSelectionChange).'
}
if ($controller -notmatch 'AddHandler\s+_app\.WindowSelectionChange') {
    throw 'Controller must subscribe to Word WindowSelectionChange.'
}
if ($controller -notmatch 'RemoveHandler\s+_app\.WindowSelectionChange') {
    throw 'Controller Dispose must unbind WindowSelectionChange symmetrically (VSTO ComEvents).'
}
if ($thisAddIn -notmatch 'EnsureSelectionToolbar' -or $thisAddIn -notmatch 'selectionToolbarController\.Dispose') {
    throw 'ThisAddIn must lazily create the controller and dispose it on Shutdown.'
}
# the guard only inspects the Startup body itself, otherwise the later EnsureSelectionToolbar
# definition reads as a call
$startupBody = [regex]::Match($thisAddIn, 'Private Sub WordAi_Startup[\s\S]*?\r?\n    End Sub').Value
if ($startupBody -notmatch 'PhaseStartupManager' -or $startupBody -match 'EnsureSelectionToolbar') {
    throw 'Selection toolbar must stay off the VSTO startup critical path (create it on first pane open).'
}

# --- 5. host window validity: application-level add-in, document window only --

if ($controller -match 'Documents\.Count') {
    throw 'Do not gate on Documents.Count; WordAi is an application-level add-in, gate on ActiveWindow.'
}
if ($controller -notmatch '_app\.ActiveWindow') {
    throw 'Controller must check _app.ActiveWindow before showing the strip.'
}
if ($controller -notmatch 'wdWindowDocument') {
    throw 'Controller must reject non-document active windows (start page, mail, other WPS windows).'
}

# --- 6. one main chain: the host only hands a sentence to the panel ----------

if ($controller -notmatch 'StartAgentFromHostEntry\(') {
    throw 'Quick actions must go through StartAgentFromHostEntry.'
}
if ($chatControl -notmatch 'Public Sub StartAgentFromHostEntry[\s\S]{0,400}HandleStartAgent\(') {
    throw 'StartAgentFromHostEntry must forward to HandleStartAgent, the single Agent entry with its gates.'
}
foreach ($bypass in @('sendMessageToServer', 'createChatSection', 'ExecuteJavaScript')) {
    if ($controller -match $bypass -or $toolbar -match $bypass) {
        throw "Host entry must not bypass the panel with '$bypass'."
    }
}

# --- 7. privacy: only ids and counts may be logged --------------------------

foreach ($log in [regex]::Matches($controller, 'AppLogger\.\w+\([^\r\n]*')) {
    if ($log.Value -match '\.Text|RequestText') {
        throw "Selection text must never reach the log: $($log.Value.Trim())"
    }
}

# --- 8. non-activating strip so the Word selection survives the click --------

if ($toolbar -notmatch 'WS_EX_NOACTIVATE\s+As\s+Integer\s*=\s*&H8000') {
    throw 'Toolbar needs the WS_EX_NOACTIVATE extended style, otherwise clicking a button clears the selection.'
}
$exStyleLines = [regex]::Matches($toolbar, '[^\r\n]*ExStyle[^\r\n]*') | ForEach-Object { $_.Value }
if (-not ($exStyleLines | Where-Object { $_ -match 'WS_EX_NOACTIVATE' -and $_ -match 'WS_EX_TOOLWINDOW' })) {
    throw 'Toolbar never applies both the no-activate and tool-window styles in one CreateParams statement.'
}
if ($toolbar -notmatch 'Overrides ReadOnly Property ShowWithoutActivation') {
    throw 'Toolbar must override ShowWithoutActivation.'
}

# --- 9. default off + a real config switch ----------------------------------

if ($settings -notmatch 'EnableSelectionToolbar As Boolean = False') {
    throw 'The toolbar strips onto every selection: it must default to off until there is a UI switch.'
}
if ($settings -notmatch '"enableSelectionToolbar"') {
    throw 'chat-settings.json must carry the enableSelectionToolbar key.'
}
if ($settings -notmatch 'ContainsKey\("enableSelectionToolbar"\)') {
    throw 'LoadSettings must tolerate the key being absent (existing profiles have no toolbar setting).'
}
if ($chatControl -notmatch 'SaveSettings\([\s\S]{0,600}EnableSelectionToolbar\s*[,)]') {
    # E4-02 appends enableContextMenu after it, so the passthrough must be read as an argument list.
    throw 'A panel settings save must not silently reset enableSelectionToolbar.'
}

# --- 10. behaviour: geometry clamps into the work area ----------------------

$geometrySource = "Imports System`r`n" + $geometry
$tmpSource = Join-Path $env:TEMP 'smoke-selection-toolbar-geometry.vb'
Set-Content -LiteralPath $tmpSource -Value $geometrySource -Encoding UTF8
try {
    Add-Type -Path $tmpSource -ReferencedAssemblies 'System.Drawing', 'System.Windows.Forms' -ErrorAction Stop
} finally {
    Remove-Item -LiteralPath $tmpSource -Force -ErrorAction SilentlyContinue
}

$type = [Selection.WordSelectionToolbarGeometry]
# a maximised word window on a 1080p screen with a 40px taskbar: the owner rect is taller than the work area
$workArea = New-Object System.Drawing.Rectangle 0, 0, 1920, 1040
$ownerRect = New-Object System.Drawing.Rectangle 0, 0, 1920, 1080

$size = $type::ComputeStripSize(3)
if ($size.Width -le 0 -or $size.Height -le 0) {
    throw "ComputeStripSize(3) returned an empty strip: $size"
}
if ($type::ItemWidth -lt 60) {
    throw "ItemWidth = $($type::ItemWidth) cannot fit a four-character Chinese label."
}

# anchor near the right edge: the strip must be pulled back, never hang off screen
$nearEdge = New-Object System.Drawing.Point 1900, 300
$placed = $type::ResolveLocation($nearEdge, $ownerRect, $workArea, $size)
if ($placed.Top -ne ($nearEdge.Y + $type::OffsetBelowAnchor)) {
    throw "An anchor inside the window should place the strip just below it, got $($placed.Top)"
}
if ($placed.Right -gt $workArea.Right -or $placed.Left -lt $workArea.Left) {
    throw "Strip escapes the work area horizontally: $placed vs $workArea"
}

# anchor below the bottom edge: flip above the anchor, still inside the work area
$nearBottom = New-Object System.Drawing.Point 900, ($workArea.Bottom - 4)
$flipped = $type::ResolveLocation($nearBottom, $ownerRect, $workArea, $size)
if ($flipped.Bottom -gt $workArea.Bottom) {
    throw "Strip does not flip above the anchor near the bottom edge: $flipped"
}
if ($flipped.Top -ge $nearBottom.Y) {
    throw "Anchor near the bottom edge must flip the strip up instead of pushing it sideways: $flipped"
}

# untrusted anchor (keyboard selection, mouse already moved away): predictable top-centre position
$outside = New-Object System.Drawing.Point -32000, -32000
$fallback = $type::ResolveLocation($outside, $ownerRect, $workArea, $size)
if (-not $workArea.Contains($fallback.Location)) {
    throw "Invalid anchor must fall back inside the work area: $fallback"
}
if ($fallback.Top -ge $placed.Top) {
    throw "Top fallback must sit above the normal anchor position: fallback=$($fallback.Top) anchor=$($placed.Top)"
}
if ($fallback.Left -ne [int](($workArea.Width - $size.Width) / 2)) {
    throw "Fallback must be centred on the work area, got $($fallback.Left)"
}

Write-Host "Geometry: size=$size edge=$($placed.Location) flip=$($flipped.Location) fallback=$($fallback.Location)"
Write-Host 'Selection toolbar smoke passed.'
