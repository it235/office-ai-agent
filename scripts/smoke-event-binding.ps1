$ErrorActionPreference = 'Stop'

# E4-05 event binding hygiene smoke.
# Acceptance is "no duplicate binding; reopening a pane must not double-trigger". None of that is a
# compile error: the old code bound the same navigation handler twice (once declaratively, once with
# AddHandler) and removed only the AddHandler copy, so every navigation re-injected the settings and
# every panel reopen re-evaluated switches that had not been loaded from disk yet. These are the
# invariants that keep those from coming back.

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')

function Read-Source([string]$relativePath) {
    $path = Join-Path $repoRoot $relativePath
    if (-not (Test-Path -LiteralPath $path)) {
        throw "Expected file is missing: $relativePath"
    }
    return Get-Content -LiteralPath $path -Raw -Encoding UTF8
}

# "Must not appear" rules read code only: comments are where the forbidden mechanism gets explained.
function Get-CodeOnly([string]$source) {
    return (($source -split "`r?`n") | Where-Object { $_ -notmatch "^\s*'" }) -join "`n"
}

function Assert-Matches([string]$text, [string]$pattern, [string]$message, [string]$where) {
    if ($text -notmatch $pattern) {
        throw "$where : $message"
    }
}

function Assert-Absent([string]$text, [string]$pattern, [string]$message, [string]$where) {
    if ($text -match $pattern) {
        throw "$where : $message"
    }
}

function Count-Of([string]$text, [string]$pattern) {
    return ([regex]::Matches($text, $pattern)).Count
}

$chatControl = Read-Source 'ShareRibbon\Controls\BaseChatControl.vb'
$chatCode = Get-CodeOnly $chatControl
$settings = Get-CodeOnly (Read-Source 'ShareRibbon\Config\ChatSettings.vb')
$menu = Get-CodeOnly (Read-Source 'ShareRibbon\Agent\HostContextMenu.vb')
$watcher = Get-CodeOnly (Read-Source 'ShareRibbon\Controls\Services\HostHotkeyWatcher.vb')

$hosts = @(
    [pscustomobject]@{
        Name      = 'Word'
        AddIn     = 'WordAi\ThisAddIn.vb'
        Control   = 'WordAi\ChatControl.vb'
        AppInfo   = 'ApplicationInfo\("Word", OfficeApplicationType\.Word\)'
    },
    [pscustomobject]@{
        Name      = 'Excel'
        AddIn     = 'ExcelAi\ThisAddIn.vb'
        Control   = 'ExcelAi\ChatControl.vb'
        AppInfo   = 'ApplicationInfo\("Excel", OfficeApplicationType\.Excel\)'
    },
    [pscustomobject]@{
        Name      = 'PowerPoint'
        AddIn     = 'PowerPointAi\ThisAddIn.vb'
        Control   = 'PowerPointAi\ChatControl.vb'
        AppInfo   = 'ApplicationInfo\("PowerPoint", OfficeApplicationType\.PowerPoint\)'
    }
)

# --- 1. behaviour: the shared switch load really is load-once -----------------------------------

$dll = Join-Path $repoRoot 'ShareRibbon\bin\Debug\ShareRibbon.dll'
if (-not (Test-Path -LiteralPath $dll)) {
    throw "ShareRibbon assembly not found: $dll. Build the code projects first."
}
Add-Type -LiteralPath $dll

$appInfo = New-Object ShareRibbon.ApplicationInfo ('Word', [ShareRibbon.OfficeApplicationType]::Word)
$first = [ShareRibbon.ChatSettings]::EnsureLoaded($appInfo)
$second = [ShareRibbon.ChatSettings]::EnsureLoaded($appInfo)
if ($null -eq $first) {
    throw 'EnsureLoaded must return the loaded settings instance.'
}
if (-not [object]::ReferenceEquals($first, $second)) {
    throw 'EnsureLoaded must be idempotent: a second call must not build (and re-read) a new instance.'
}
Assert-Matches $settings 'If _loadedInstance IsNot Nothing Then Return' `
    'Load-once must be guarded in source: hosts call this on every pane visibility change.' 'ChatSettings'

# --- 2. the navigation bootstrap is bound once and runs once per document ------------------------

Assert-Absent $chatCode 'AddHandler ChatBrowser\.CoreWebView2\.NavigationCompleted' `
    'NavigationCompleted is already bound by the Handles clause; a second AddHandler runs the whole bootstrap twice per load.' 'BaseChatControl'
Assert-Absent $chatCode 'RemoveHandler ChatBrowser\.CoreWebView2\.NavigationCompleted' `
    'Removing only the imperative copy left the Handles-bound one live forever, so every later navigation re-injected the settings.' 'BaseChatControl'

$handlerCount = Count-Of $chatCode 'OnWebViewNavigationCompleted'
if ($handlerCount -ne 1) {
    throw "OnWebViewNavigationCompleted must be declared exactly once (declarative Handles only), found $handlerCount."
}
Assert-Matches $chatCode 'Handles ChatBrowser\.NavigationCompleted' `
    'The bootstrap must stay bound through the Handles clause, that is the single binding.' 'BaseChatControl'

# Once-per-document: set before the work, cleared only where a fresh document is actually navigated.
$resetIndex = $chatCode.IndexOf('_pageBootstrapDone = False')
$navigateIndex = $chatCode.IndexOf('ChatBrowser.CoreWebView2.NavigateToString(htmlContent)')
if ($resetIndex -lt 0 -or $navigateIndex -lt 0) {
    throw 'The bootstrap flag must be reset right before the page is navigated.'
}
if ($resetIndex -gt $navigateIndex) {
    throw 'The reset must happen before NavigateToString: a fresh document has to be allowed to bootstrap again.'
}
$bootstrap = [regex]::Match($chatCode, 'Private Async Sub OnWebViewNavigationCompleted[\s\S]*?\r?\n    End Sub')
if (-not $bootstrap.Success) {
    throw 'Cannot locate OnWebViewNavigationCompleted.'
}
Assert-Matches $bootstrap.Value 'If _pageBootstrapDone Then Return' `
    'The guard must bail out before any work: re-entrant navigation events are the bug this row removes.' 'BaseChatControl'
# Claim the slot before awaiting, otherwise two events inside the delay both pass the check.
$claimIndex = $bootstrap.Value.IndexOf('If _pageBootstrapDone Then Return')
$setIndex = $bootstrap.Value.IndexOf('_pageBootstrapDone = True')
$awaitIndex = $bootstrap.Value.IndexOf('Await ')
if ($setIndex -lt $claimIndex) {
    throw 'The flag must be checked before it is claimed.'
}
if ($setIndex -gt $awaitIndex) {
    throw 'The flag must be claimed before the first Await: the handler is Async Sub and two events can overlap.'
}
$flagTotal = Count-Of $chatCode '_pageBootstrapDone'
if ($flagTotal -ne 4) {
    throw "The bootstrap flag must appear exactly 4 times (declare, reset, check, claim), found $flagTotal."
}

# WebMessageReceived is the one binding that would double-handle every panel message.
$messageHandlers = Count-Of $chatCode 'AddHandler ChatBrowser\.WebMessageReceived'
if ($messageHandlers -ne 1) {
    throw "WebMessageReceived must be bound exactly once, found $messageHandlers AddHandler calls."
}
Assert-Absent $chatCode 'AddHandler Me\.Disposed, Sub\(sender As Object, args As EventArgs\) _commandPaletteService\.Shutdown\(\)[\s\S]{0,200}?AddHandler Me\.Disposed' `
    'The palette shutdown handler must not be attached twice.' 'BaseChatControl'
$paletteLazy = [regex]::Match($chatCode, 'If _commandPaletteService Is Nothing Then[\s\S]{0,400}?AddHandler Me\.Disposed')
if (-not $paletteLazy.Success) {
    throw 'The Disposed handler must live inside the lazy creation guard, otherwise every property read attaches another shutdown.'
}

# --- 3. hosts must not grow a second settings bootstrap ------------------------------------------

foreach ($host_ in $hosts) {
    $control = Get-CodeOnly (Read-Source $host_.Control)
    $load = [regex]::Match($control, 'Private Async Sub MainForm_Load\(sender As Object, e As EventArgs\) Handles MyBase\.Load[\s\S]*?\r?\n    End Sub')
    if (-not $load.Success) {
        throw "$($host_.Control) must bootstrap the panel from MainForm_Load (Handles MyBase.Load)."
    }
    Assert-Absent $load.Value 'InitializeSettings\(\)' `
        "$($host_.Name) injected the settings itself; the shared navigation bootstrap owns that exactly once."
    # A declarative Handles binding plus an AddHandler on the same Load event is the same double-fire.
    Assert-Absent $control 'AddHandler[^\r\n]*\.Load,' `
        "$($host_.Name) must not also AddHandler the Load event it already Handles."
    $explicitCalls = Count-Of $control 'InitializeSettings\(\)'
    if ($explicitCalls -ne 0) {
        throw "$($host_.Name) still calls InitializeSettings() $($explicitCalls) time(s) outside the shared bootstrap."
    }
}

# The declaration is "Protected Sub InitializeSettings()", so only count a bare invocation line.
$sharedBootstrapCalls = Count-Of $chatCode '(?m)^\s*InitializeSettings\(\)\s*$'
if ($sharedBootstrapCalls -ne 1) {
    throw "InitializeSettings() must be invoked from the single guarded bootstrap, found $sharedBootstrapCalls."
}

# --- 4. entry switches are read only after the settings file has been loaded ---------------------

foreach ($host_ in $hosts) {
    $addIn = Get-CodeOnly (Read-Source $host_.AddIn)
    $visible = [regex]::Match($addIn, 'Private Sub ChatTaskPane_VisibleChanged[\s\S]*?\r?\n    End Sub')
    if (-not $visible.Success) {
        throw "Cannot locate ChatTaskPane_VisibleChanged in $($host_.AddIn)."
    }
    Assert-Matches $visible.Value ('ChatSettings\.EnsureLoaded\(New ' + $host_.AppInfo + '\)') `
        "$($host_.Name) must load the settings file before reading the entry switches: the first pane open of a process used to see only the code defaults." `
        $host_.AddIn
    $loadIndex = $visible.Value.IndexOf('ChatSettings.EnsureLoaded(')
    $entryIndex = $visible.Value.IndexOf('EnsureContextMenu()')
    if ($loadIndex -lt 0 -or $entryIndex -lt 0 -or $loadIndex -gt $entryIndex) {
        throw "$($host_.AddIn) must call EnsureLoaded before the entry methods that read the switches."
    }

    # WPS width timers are created on demand and started on every visibility change.
    Assert-Matches $addIn 'If widthTimer Is Nothing Then[\s\S]{0,160}AddHandler widthTimer\.Tick' `
        "$($host_.Name) must attach the width timer Tick inside the creation guard." $host_.AddIn
    Assert-Matches $addIn 'If widthTimer1 Is Nothing Then[\s\S]{0,160}AddHandler widthTimer1\.Tick' `
        "$($host_.Name) must attach the deepseek width timer Tick inside the creation guard." $host_.AddIn
    $shutdown = [regex]::Match($addIn, 'Sub ThisAddIn_Shutdown[\s\S]*?\r?\n    End Sub')
    if (-not $shutdown.Success) {
        throw "Cannot locate ThisAddIn_Shutdown in $($host_.AddIn)."
    }
    Assert-Matches $shutdown.Value 'widthTimer\.Stop\(\)[\s\S]{0,120}widthTimer\.Dispose\(\)' `
        "$($host_.Name) must stop and dispose the timers it started, otherwise the tick keeps running after shutdown." $host_.AddIn
    Assert-Matches $shutdown.Value 'contextMenu\.Detach\(\)' `
        "$($host_.Name) must detach its context menu items on shutdown." $host_.AddIn
}

# The Word toolbar is the other consumer of a switch read from host code.
$toolbarController = Get-CodeOnly (Read-Source 'WordAi\Selection\WordSelectionToolbarController.vb')
Assert-Matches $toolbarController 'ChatSettings\.EnableSelectionToolbar' `
    'The toolbar must keep honouring its switch.' 'WordSelectionToolbarController'
Assert-Matches (Get-CodeOnly (Read-Source 'WordAi\ThisAddIn.vb')) `
    'EnsureLoaded[\s\S]{0,200}EnsureSelectionToolbar' `
    'Word must load the settings before the toolbar can read its switch.' 'WordAi\ThisAddIn.vb'

# --- 5. re-opening a pane must not stack menu items or handlers ----------------------------------

Assert-Matches $menu 'If _attached Then Detach\(\)' `
    'Attach must clear its previous items first: the menu is re-attached on every pane visibility change.' 'HostContextMenu'
Assert-Matches $menu 'AddHandler button\.Click, AddressOf OnButtonClick' `
    'Click must be hooked on the created button (the RCW has to stay referenced).' 'HostContextMenu'
Assert-Matches $menu 'RemoveHandler button\.Click, AddressOf OnButtonClick' `
    'Detach must unhook exactly the handlers Attach added.' 'HostContextMenu'
$detach = [regex]::Match($menu, 'Public Sub Detach\(\)[\s\S]*?\r?\n        End Sub')
if (-not $detach.Success) {
    throw 'Cannot locate HostContextMenu.Detach.'
}
Assert-Matches $detach.Value '_attached = False[\s\S]{0,300}_buttons\.Clear\(\)' `
    'Detach must clear both the flag and the retained buttons, or a re-attach sees a stale state.' 'HostContextMenu'

# The hotkey watcher keeps its own add/remove symmetry (E4-03), asserted here so this row owns the rule.
Assert-Matches $watcher 'AddHandler _timer\.Tick' 'The watcher must hook its timer tick.' 'HostHotkeyWatcher'
Assert-Matches $watcher 'RemoveHandler _timer\.Tick' 'StopWatching must remove the tick handler it added.' 'HostHotkeyWatcher'
$tickAdds = Count-Of $watcher 'AddHandler _timer\.Tick'
$tickRemoves = Count-Of $watcher 'RemoveHandler _timer\.Tick'
if ($tickAdds -ne $tickRemoves) {
    throw "Add/Remove must stay symmetric, got $tickAdds adds and $tickRemoves removes."
}

Write-Host 'Event binding hygiene smoke passed.'
