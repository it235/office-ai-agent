$ErrorActionPreference = 'Stop'

# E4-03 command palette + host hotkey smoke.
# The chord parser, the watcher registration rules and the palette queue are pure managed code, so
# they are asserted by calling the compiled assembly. What only shows up in the real host (a physical
# key press, WebView focus) is left to the machine checklist; everything checked here is an invariant
# that rots silently: the exclusive-resource rules, the single catalog source, the four-way front-end
# registration, and the settings round-trip that already swallowed two switches once.

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')

function Read-Source([string]$relativePath) {
    $path = Join-Path $repoRoot $relativePath
    if (-not (Test-Path -LiteralPath $path)) {
        throw "Expected file is missing: $relativePath"
    }
    return Get-Content -LiteralPath $path -Raw -Encoding UTF8
}

# Comments are where a forbidden mechanism gets explained, so a "must not appear" rule reads code only.
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

$watcher = Get-CodeOnly (Read-Source 'ShareRibbon\Controls\Services\HostHotkeyWatcher.vb')
$paletteService = Get-CodeOnly (Read-Source 'ShareRibbon\Controls\Services\CommandPaletteService.vb')
$settings = Read-Source 'ShareRibbon\Config\ChatSettings.vb'
$chatControl = Get-CodeOnly (Read-Source 'ShareRibbon\Controls\BaseChatControl.vb')
$shareProj = Read-Source 'ShareRibbon\ShareRibbon.vbproj'
$paletteJs = Read-Source 'ShareRibbon\Resources\js\command-palette.js'
$settingsJs = Read-Source 'ShareRibbon\Resources\js\settings-manager.js'
$template = Read-Source 'ShareRibbon\Resources\chat-template-refactored.html'
$styles = Read-Source 'ShareRibbon\Resources\css\styles.css'
$extractor = Get-CodeOnly (Read-Source 'ShareRibbon\Resources\ResourceExtractor.vb')
$resx = Read-Source 'ShareRibbon\My Project\Resources.resx'

# --- 1. behaviour: a chord is an exclusive resource, not a free-form string ---------------------

$dll = Join-Path $repoRoot 'ShareRibbon\bin\Debug\ShareRibbon.dll'
if (-not (Test-Path -LiteralPath $dll)) {
    throw "ShareRibbon assembly not found: $dll. Build the code projects first."
}
Add-Type -LiteralPath $dll

$chord = [ShareRibbon.HostHotkeyChord]

foreach ($case in @(
        @{ Text = 'Ctrl+Alt+K'; Want = 'Ctrl+Alt+K' },
        @{ Text = 'ctrl+k'; Want = 'Ctrl+K' },
        @{ Text = 'Control+Shift+F12'; Want = 'Ctrl+Shift+F12' },
        @{ Text = 'Alt+/'; Want = 'Alt+/' },
        @{ Text = 'Ctrl+.'; Want = 'Ctrl+.' })) {
    $parsed = $chord::Parse($case.Text)
    if ($null -eq $parsed) {
        throw "Parse('$($case.Text)') returned nothing; a valid chord must round-trip."
    }
    if (-not $parsed.IsSafe) {
        throw "Parse('$($case.Text)') must be registrable: it carries Ctrl or Alt."
    }
    if ($parsed.ToString() -ne $case.Want) {
        throw "Parse('$($case.Text)') normalizes to '$($parsed.ToString())', expected '$($case.Want)'."
    }
}

# A bare key would swallow normal typing, two main keys is a typo, a modifier alone has nothing to poll.
foreach ($bad in @('K', 'Space', 'Ctrl+', '', '   ', 'Ctrl+Shift+A+B')) {
    $parsed = $chord::Parse($bad)
    if ($null -ne $parsed -and $parsed.IsSafe) {
        throw "Parse('$bad') must not produce a registrable chord."
    }
}
# Shift+A has a real main key: it must be refused for being Shift-only, not accidentally unparseable.
$shiftOnly = $chord::Parse('Shift+A')
if ($null -eq $shiftOnly -or $shiftOnly.MainKey -eq 0 -or $shiftOnly.IsSafe) {
    throw 'Shift+A must parse to a main key and still be refused as unsafe (Ctrl or Alt is required).'
}

# --- 2. behaviour: one watcher, refuse instead of stacking --------------------------------------

$instance = New-Object ShareRibbon.HostHotkeyWatcher
if ($instance.IsWatching) {
    throw 'A fresh watcher must not already be watching.'
}
$noop = [Action] { }
if ($instance.Start('K', $noop)) {
    throw 'Start must refuse a chord without Ctrl/Alt: a bare key steals the keyboard.'
}
if (-not $instance.Start('Ctrl+Alt+K', $noop)) {
    throw 'Start must accept a safe chord.'
}
if (-not $instance.IsWatching) {
    throw 'IsWatching must report the live registration.'
}
if ($instance.Start('Ctrl+Alt+J', $noop)) {
    throw 'A second Start must fail: the hotkey is exclusive, two triggers means a double-open.'
}
$instance.StopWatching()
if ($instance.IsWatching) {
    throw 'StopWatching must clear the registration.'
}
$instance.Dispose()
$instance.Dispose()

# --- 3. behaviour: an open request made before the page is ready is not lost ---------------------

# The Func delegate is invoked from compiled VB, so the list has to live in the global scope.
$global:cpSent = New-Object 'System.Collections.Generic.List[string]'
$exec = [Func[string, System.Threading.Tasks.Task]] {
    param($js)
    $global:cpSent.Add($js) | Out-Null
    return [System.Threading.Tasks.Task]::CompletedTask
}
$appType = [Func[string]] { return 'Word' }
$global:paneCalls = 0
$ensureVisible = [Action] { $global:paneCalls += 1 }

$svc = New-Object ShareRibbon.CommandPaletteService ($exec, $appType, $ensureVisible)

$svc.Open()
if ($global:cpSent.Count -ne 0) {
    throw "Open() before the page is ready must queue, got: $($global:cpSent -join ' | ')"
}

$svc.PushState()
$synced = @($global:cpSent | Where-Object { $_ -like 'syncCommandPalette(*' })
$shown = @($global:cpSent | Where-Object { $_ -like 'showCommandPalette(*' })
if ($synced.Count -ne 1) { throw 'PushState must sync the palette state to the page exactly once.' }
if ($shown.Count -ne 1) { throw 'PushState must flush the queued open request; the keystroke must not be lost.' }

$payload = ($synced[0] -replace '^syncCommandPalette\(', '' -replace '\);$', '') | ConvertFrom-Json
if ($payload.appType -ne 'Word') { throw 'The palette state must carry the host app type.' }
if (@($payload.actions).Count -ne 3) {
    throw "The palette must receive the Word catalog (3 actions), got $(@($payload.actions).Count)."
}
if (-not ([string]$payload.actions[0].request).Trim()) {
    throw 'Every palette row needs a request text; a row without one cannot reach the Agent.'
}
if ($payload.PSObject.Properties.Name -notcontains 'shortcut') {
    throw 'The page needs the chord text to show it on the button and in the settings row.'
}

# --- 3b. E6-04: the same rows now arrive grouped by task package, still one flat list -----------

if (@($payload.packages).Count -lt 1) {
    throw "The Word palette state must carry at least the word.review-polish package header, got $(@($payload.packages).Count)."
}
$grouped = @($payload.actions | Where-Object { $_.PSObject.Properties.Name -contains 'packageId' })
if ($grouped.Count -ne 2) {
    throw "word.review-polish registers 2 entry actions, so exactly 2 rows must carry packageId; got $($grouped.Count)."
}
if ($grouped[0].packageId -ne 'word.review-polish') {
    throw "Grouped rows must be emitted first and in package order, got $($grouped[0].id)."
}
if (@($payload.actions | Where-Object { $_.id -eq 'word.selection.summarize' -and ($_.PSObject.Properties.Name -contains 'packageId') }).Count -ne 0) {
    throw 'An action that no package declares must still be sent, but ungrouped - not dropped or invented into a package.'
}
if (-not ([string]$payload.packages[0].title).Trim()) {
    throw 'A package header without a title gives the page nothing to group under.'
}
if ($global:paneCalls -ne 1) {
    throw "Open() must ask the host to expand the task pane first, calls=$($global:paneCalls)."
}

$global:cpSent.Clear()
$svc.Open()
if ($global:cpSent.Count -ne 1 -or $global:cpSent[0] -notlike 'showCommandPalette(*') {
    throw "Open() after the page is ready must push directly, got: $($global:cpSent -join ' | ')"
}

# --- 4. off by default, persisted, and read back the same way ------------------------------------

if ([ShareRibbon.ChatSettings]::EnableHotkey) {
    throw 'ChatSettings.EnableHotkey must default to False: a global chord must never arrive silently.'
}
if ([ShareRibbon.ChatSettings]::HotkeyChord -ne 'Ctrl+Alt+K') {
    throw "Unexpected default chord: $([ShareRibbon.ChatSettings]::HotkeyChord)"
}
Assert-Matches $settings 'EnableHotkey As Boolean = False' 'EnableHotkey must be declared with a False default.' 'ChatSettings'

foreach ($key in @('enableHotkey', 'hotkeyChord', 'enableSelectionToolbar', 'enableContextMenu')) {
    Assert-Matches $settings ('settings\("' + $key + '"\)') "SaveSettings must write $key." 'ChatSettings'
    Assert-Matches $settings ('ContainsKey\("' + $key + '"\)') "LoadSettings must read $key back." 'ChatSettings'
    Assert-Matches $settingsJs $key "settings-manager.js must submit $key, or saving the dialog drops it." 'settings-manager.js'
}

# SaveSettings used to rebuild the file from a fixed table and deleted enableMcpList with it.
Assert-Matches $settings 'JObject\.Parse\(File\.ReadAllText\(settingsPath\)\)' `
    'SaveSettings must merge into the existing JSON file instead of overwriting it.' 'ChatSettings'
Assert-Matches $settings 'HostHotkeyChord\.Parse' `
    'LoadSettings must re-validate the stored chord: a file holding a bare key must not register it.' 'ChatSettings'

# The three native-entry switches must come from the payload; writing the statics back is what made
# the E4-01/E4-02 switches read-only.
foreach ($key in @('enableHotkey', 'enableSelectionToolbar', 'enableContextMenu')) {
    Assert-Matches $chatControl ('ReadFlagSetting\(jsonDoc, "' + $key + '"') `
        "HandleSaveSettings must read $key from the payload." 'BaseChatControl'
}
Assert-Matches $chatControl 'jsonDoc\("hotkeyChord"\)' 'HandleSaveSettings must read hotkeyChord from the payload.' 'BaseChatControl'
Assert-Matches $chatControl 'enableHotkey AndAlso \(parsedChord Is Nothing OrElse Not parsedChord\.IsSafe\)' `
    'An unsafe chord must be refused at save time with the switch left off, not stored and silently dead.' 'BaseChatControl'

# --- 5. one catalog, one execution chain, four-way front-end registration ------------------------

Assert-Matches $paletteService 'QuickActionCatalog\.ActionsFor\(appType\)' `
    'CommandPaletteService must take its rows from QuickActionCatalog; no second action table.' 'CommandPaletteService'
Assert-Absent $paletteService 'startAgent|sendMessageToServer' `
    'The backend service must not start an Agent run itself; the page sends the single startAgent message.' 'CommandPaletteService'
Assert-Matches $paletteService 'GlobalStatusStrip\.ShowWarning' `
    'A refused registration has to tell the user; silent failure is the pattern this repo is removing.' 'CommandPaletteService'
Assert-Matches $paletteService '_watcher\.StopWatching\(\)' `
    'Changing the chord must stop the old poll first, or two chords stay live at once.' 'CommandPaletteService'

foreach ($entry in @(
        @{ Name = 'the single startAgent exit'; Pattern = "type:\s*'startAgent'" },
        @{ Name = 'catalog rows'; Pattern = 'commandPaletteActions' },
        @{ Name = 'natural language row'; Pattern = '__natural_language__' },
        @{ Name = 'close on Escape'; Pattern = "key === 'Escape'" },
        @{ Name = 'arrow navigation'; Pattern = 'ArrowDown' },
        @{ Name = 'cached state without a round trip'; Pattern = 'window\.commandPaletteState' })) {
    Assert-Matches $paletteJs $entry.Pattern "command-palette.js is missing $($entry.Name)." 'command-palette.js'
}
Assert-Absent $paletteJs 'fetch\(|XMLHttpRequest|invokeMethod' `
    'The palette must not reach for a model or the host directly; one message to the backend only.' 'command-palette.js'

# A JS file missing from resx + extractor + HTML is the classic "the panel keeps loading the old copy".
Assert-Matches $resx 'name="command_palette"' 'Resources.resx does not declare command_palette.' 'Resources.resx'
Assert-Matches $extractor '"command_palette", "command-palette\.js"' 'ResourceExtractor does not release command-palette.js.' 'ResourceExtractor'
Assert-Matches $template 'js/command-palette\.js' 'The panel HTML does not load command-palette.js.' 'chat-template.html'
Assert-Matches $shareProj ([regex]::Escape('Controls\Services\HostHotkeyWatcher.vb')) 'ShareRibbon.vbproj does not compile HostHotkeyWatcher.vb.' 'ShareRibbon.vbproj'
Assert-Matches $shareProj ([regex]::Escape('Controls\Services\CommandPaletteService.vb')) 'ShareRibbon.vbproj does not compile CommandPaletteService.vb.' 'ShareRibbon.vbproj'

foreach ($id in @('command-palette-button', 'settings-enable-hotkey', 'settings-hotkey-chord',
                  'settings-enable-selection-toolbar', 'settings-enable-context-menu')) {
    Assert-Matches $template ('id="' + $id + '"') "The panel HTML has no #$id element." 'chat-template.html'
}
Assert-Matches $styles '\.command-palette\s*\{' 'styles.css must carry the palette layout.' 'styles.css'
Assert-Matches $chatControl 'CommandPaletteSvc\.PushState\(\)' 'InitializeSettings must push the palette state once the page is ready.' 'BaseChatControl'
Assert-Matches $chatControl 'CommandPaletteSvc\.ApplyHotkey\(\)' 'The chord must be registered on load and after saving, without restarting the host.' 'BaseChatControl'
Assert-Matches $chatControl 'AddressOf EnsureHostTaskPaneVisible' 'BaseChatControl must hand the palette its pane-visible hook.' 'BaseChatControl'

foreach ($hostName in @('WordAi', 'ExcelAi', 'PowerPointAi')) {
    $control = Get-CodeOnly (Read-Source "$hostName\ChatControl.vb")
    Assert-Matches $control 'Overrides Sub EnsureHostTaskPaneVisible' `
        "$hostName must expand its task pane for the hotkey: a palette pushed into a collapsed pane is invisible." $hostName
}

# --- 6. watcher hygiene: no global hook, foreground-only, no reentrancy --------------------------

foreach ($forbidden in @('SetWindowsHookEx', 'RegisterHotKey', 'keybd_event', 'SendInput')) {
    Assert-Absent $watcher $forbidden "HostHotkeyWatcher must not use $forbidden." 'HostHotkeyWatcher'
}
Assert-Matches $watcher 'GetAsyncKeyState' 'HostHotkeyWatcher must poll GetAsyncKeyState.' 'HostHotkeyWatcher'
Assert-Matches $watcher 'GetForegroundWindow' 'Global key state must be gated on the host being foreground.' 'HostHotkeyWatcher'
Assert-Matches $watcher '_hostProcessId' 'The foreground check must compare the host process id.' 'HostHotkeyWatcher'
Assert-Matches $watcher 'MainKey <> 0 AndAlso \(Control OrElse Alt\)' 'IsSafe must require Ctrl or Alt in source.' 'HostHotkeyWatcher'
Assert-Matches $watcher 'IsKeyDown\(VK_SHIFT\) = chord\.Shift' 'Modifiers must match as an exact set; Ctrl+Alt+K plus Shift is a different chord.' 'HostHotkeyWatcher'
Assert-Matches $watcher '_onTrigger\(\)[\s\S]{0,300}?_timer\.Start\(\)' `
    'The timer must be restarted after the callback: it is stopped around it to keep one press to one open.' 'HostHotkeyWatcher'
Assert-Matches $watcher 'RemoveHandler _timer\.Tick' 'StopWatching must remove the handler it added (E4-05 hygiene rule).' 'HostHotkeyWatcher'

foreach ($file in @('ShareRibbon\Controls\Services\HostHotkeyWatcher.vb',
                    'ShareRibbon\Controls\Services\CommandPaletteService.vb')) {
    $lines = (Get-Content -LiteralPath (Join-Path $repoRoot $file) -Encoding UTF8).Count
    if ($lines -gt 500) {
        throw "$file is $lines lines; the project rule caps a class at 500."
    }
}

Write-Host 'Command palette smoke passed.'
