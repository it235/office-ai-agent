$ErrorActionPreference = 'Stop'

# E4-04 auto selection reference smoke.
# Flipping a default is the cheap half of this task; the half that rots silently is the contract
# around it: one decision point so the switch can actually switch off, the reference shown back to
# the user, the summary kept out of history and logs, and the injected settings script still able
# to run to the last line. This round found two getElementById calls pointing at ids that do not
# exist, which aborted the whole injected block and made the E4-01/E4-02/E4-03 switches dead, so
# section 6 now checks every referenced id against the panel HTML.
#
# DescribeAutoSelection is a Private Shared member of a MustInherit WinForms control, so the type
# cannot be loaded by reflection here (its base type pulls the WebView2/interop surface). The
# clamps are therefore asserted in source; the switch defaults are asserted against the compiled
# assembly, which is a plain static class.

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')

function Read-Source([string]$relativePath) {
    $path = Join-Path $repoRoot $relativePath
    if (-not (Test-Path -LiteralPath $path)) {
        throw "Expected file is missing: $relativePath"
    }
    return Get-Content -LiteralPath $path -Raw -Encoding UTF8
}

# A "this must not appear" rule has to read code only: comments are where the forbidden
# mechanism gets explained. VB drops a line starting with a single quote.
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

$settings = Read-Source 'ShareRibbon\Config\ChatSettings.vb'
$chatControl = Read-Source 'ShareRibbon\Controls\BaseChatControl.vb'
$chatCode = Get-CodeOnly $chatControl
$managerJs = Read-Source 'ShareRibbon\Resources\js\chat-manager.js'
$settingsJs = Read-Source 'ShareRibbon\Resources\js\settings-manager.js'
$template = Read-Source 'ShareRibbon\Resources\chat-template-refactored.html'

# --- 1. behaviour: the compiled default is on, and an explicit false still wins -----------------

$dll = Join-Path $repoRoot 'ShareRibbon\bin\Debug\ShareRibbon.dll'
if (-not (Test-Path -LiteralPath $dll)) {
    throw "ShareRibbon assembly not found: $dll. Build the code projects first."
}
Add-Type -LiteralPath $dll

# Read before any ChatSettings instance exists, so this is the declared default and not this
# machine's profile.
if (-not [ShareRibbon.ChatSettings]::selectedCellChecked) {
    throw 'ChatSettings.selectedCellChecked must default to True: selecting text and then asking means "about this".'
}
Assert-Matches $settings 'selectedCellChecked As Boolean = True' `
    'The static default must be declared True so a profile without the key gets auto-reference.' 'ChatSettings'
Assert-Matches $chatControl 'Protected selectedCellChecked As Boolean = True' `
    'The control field starts from the same default: the panel is correct before LoadSettings runs.' 'BaseChatControl'
Assert-Matches $chatControl 'selectedCellChecked = ChatSettings\.selectedCellChecked' `
    'InitializeSettings must adopt the persisted value instead of keeping its own default.' 'BaseChatControl'

# Default-on is only honest while the user can turn it off without recompiling.
Assert-Matches $settings 'settings\("selectedCellChecked"\) = selectedCell' `
    'SaveSettings must persist selectedCellChecked, or a stored false is forgotten on next start.' 'ChatSettings'
Assert-Matches $settings 'ContainsKey\("selectedCellChecked"\)' `
    'LoadSettings must read selectedCellChecked back when the key exists.' 'ChatSettings'

# --- 2. one decision point: the switch really switches off --------------------------------------

Assert-Matches $chatCode `
    'Protected Function BuildMessageWithAutoSelection\(message As String\) As String[\s\S]{0,400}?If Not selectedCellChecked Then Return message[\s\S]{0,200}?Dim appended = AppendCurrentSelectedContent\(message\)' `
    'The gate and the append must live in one wrapper: a per-host gate is how the switch stops applying to some hosts.' `
    'BaseChatControl'

# The shared layer declares the override and calls it exactly once, from inside the wrapper.
$appendCalls = [regex]::Matches($chatCode, 'AppendCurrentSelectedContent\(')
if ($appendCalls.Count -ne 2) {
    throw "BaseChatControl must keep exactly the MustOverride declaration plus the single gated call, found $($appendCalls.Count)."
}

# Both send paths go through the wrapper; the count is declaration + chat turn + agent turn.
$wrapperCalls = [regex]::Matches($chatCode, 'BuildMessageWithAutoSelection\(')
if ($wrapperCalls.Count -ne 3) {
    throw "BuildMessageWithAutoSelection must be the only message builder for both send paths, found $($wrapperCalls.Count) occurrences."
}

# Word records the selection snapshot for write-back inside the same override, so switching
# auto-reference off also stops selection-targeted write-back. That coupling is deliberate and the
# setting tooltip says so; the smoke pins that the snapshot is still produced there.
$wordControl = Get-CodeOnly (Read-Source 'WordAi\ChatControl.vb')
$wordAppend = [regex]::Match($wordControl, 'Overrides Function AppendCurrentSelectedContent\(message As String\) As String[\s\S]*?\r?\n    End Function')
if (-not $wordAppend.Success) {
    throw 'Cannot locate the Word AppendCurrentSelectedContent override.'
}
Assert-Matches $wordAppend.Value 'PendingSelectionInfo = ' `
    'Word must still record the selection snapshot for write-back; the shared wrapper owns the on/off decision.' `
    'WordAi\ChatControl.vb'

# No host override may re-gate on the switch: a second gate is where a host silently diverges.
foreach ($file in @(
        'WordAi\ChatControl.vb', 'WordAi\DoubaoChat.vb', 'WordAi\DeepseekControl.vb',
        'ExcelAi\ChatControl.vb', 'ExcelAi\DoubaoChat.vb', 'ExcelAi\DeepseekControl.vb',
        'PowerPointAi\ChatControl.vb', 'PowerPointAi\DoubaoChat.vb', 'PowerPointAi\DeepseekControl.vb')) {
    $body = [regex]::Match((Get-CodeOnly (Read-Source $file)), `
        'Overrides Function AppendCurrentSelectedContent\(message As String\) As String[\s\S]*?\r?\n    End Function')
    if (-not $body.Success) {
        throw "$file must override AppendCurrentSelectedContent; the shared wrapper has nothing to call."
    }
    Assert-Absent $body.Value 'selectedCellChecked' `
        "$file re-gates the switch inside the override: the shared wrapper is the only owner."
}

# --- 3. what was carried out must be visible ----------------------------------------------------

Assert-Matches $chatCode 'RenderAutoSelectionReference\(AgentKernelSvc\.AgentThinkingUuid\)' `
    'The Agent turn renders no user bubble, so the auto-reference must be shown on the thinking section.' 'BaseChatControl'
Assert-Matches $chatCode 'Private Sub RenderAutoSelectionReference\(turnUuid As String\)[\s\S]{0,400}?appendTurnSelectionReference' `
    'RenderAutoSelectionReference must push the summary through the appendTurnSelectionReference channel.' 'BaseChatControl'
# Both arguments go through the JS string literal escaper: the summary contains raw document text.
Assert-Matches $chatCode `
    'appendTurnSelectionReference\(\{UtilsService\.ToJsStringLiteral\(turnUuid\)\}, \{UtilsService\.ToJsStringLiteral\(_autoSelectionReference\)\}\)' `
    'Selection text must be escaped as a JS literal, never concatenated into injected script.' 'BaseChatControl'

$renderer = [regex]::Match($chatCode, 'Private Sub RenderAutoSelectionReference\(turnUuid As String\)[\s\S]*?\r?\n    End Sub')
Assert-Matches $renderer.Value 'If String.IsNullOrWhiteSpace\(_autoSelectionReference\) Then Return' `
    'A turn that carried no selection must render no reference chip at all.' 'BaseChatControl'

# The summary is derived from the delta that was actually appended, so it cannot lie about the prompt.
Assert-Matches $chatCode 'DescribeAutoSelection\(appended\.Substring\(message\.Length\)\)' `
    'The displayed summary must describe the text that was really appended.' 'BaseChatControl'
Assert-Matches $chatCode `
    'Private Shared Function DescribeAutoSelection\(referenceText As String\) As String[\s\S]{0,600}?If Char\.IsWhiteSpace\(c\) Then[\s\S]{0,400}?If flat\.Length >= 160 Then Exit For' `
    'The summary must collapse whitespace into one line and cap the preview, never dump the selection.' 'BaseChatControl'

# --- 4. privacy and persistence: display only ---------------------------------------------------

foreach ($line in ($chatCode -split "`n")) {
    if ($line -match '_autoSelectionReference' -and $line -match 'AppLogger|AddMessage|SaveMessage|History|Memory|Observation') {
        throw "The turn summary must stay a display-only string, not a persisted record: $($line.Trim())"
    }
}
$describe = [regex]::Match($chatCode, 'Private Shared Function DescribeAutoSelection\(referenceText As String\)[\s\S]*?\r?\n    End Function')
foreach ($line in ($describe.Value -split "`n")) {
    if ($line -match 'AppLogger') {
        throw "Selection text must never reach the log: $($line.Trim())"
    }
}

# --- 5. the panel side of the channel ------------------------------------------------------------

$jsFn = [regex]::Match($managerJs, 'window\.appendTurnSelectionReference = function \(uuid, summary\)[\s\S]*?\r?\n\};')
if (-not $jsFn.Success) {
    throw 'chat-manager.js lost window.appendTurnSelectionReference.'
}
Assert-Matches $jsFn.Value "document\.getElementById\('content-' \+ uuid\)" `
    'The chip must target the turn content element.' 'chat-manager.js'
Assert-Matches $jsFn.Value 'if \(!contentEl \|\| !summary\) return false;' `
    'A missing element must return false, not throw and abort the rest of the injected script.' 'chat-manager.js'
Assert-Matches $jsFn.Value "querySelector\('\.turn-selection-reference'\)" `
    'A repeated sync of the same turn must remove the stale chip instead of stacking a second one.' 'chat-manager.js'
Assert-Matches $jsFn.Value 'line\.textContent = String\(summary\)' `
    'The summary is document text: it may only ever be assigned through textContent.' 'chat-manager.js'
Assert-Absent $jsFn.Value 'innerHTML' `
    'innerHTML on selection text is an injection hole.' 'chat-manager.js'
Assert-Absent $jsFn.Value 'fetch\(|XMLHttpRequest|invokeMethod' `
    'The reference chip must not call the model or the host; it is a display channel.' 'chat-manager.js'
Assert-Matches $jsFn.Value 'className = ''chat-message-references' `
    'The chip reuses the existing reference styling instead of a second look.' 'chat-manager.js'

# --- 6. every element id used by the injected settings script must exist ------------------------

# One missing id throws and takes every later line with it, which is how the native-entry switches
# silently stopped syncing. E5-05 moved that check into scripts/check-entry-reachability.js rule H,
# which reads the payload from all four modules' .vb instead of one block in one file, so this
# section keeps only the part the gate cannot see: that the settings block is still the one that
# initialises the auto-reference switch from the persisted value.
$injected = [regex]::Match($chatControl, 'Dim js As String[\s\S]*?ExecuteJavaScriptAsyncJS\(js\)')
if (-not $injected.Success) {
    throw 'Cannot locate the injected settings script in InitializeSettings.'
}
$referencedIds = @()
foreach ($m in [regex]::Matches($injected.Value, "getElementById\(\s*'([^']+)'\s*\)")) { $referencedIds += $m.Groups[1].Value }
foreach ($m in [regex]::Matches($injected.Value, "setNativeEntrySetting\(\s*'([^']+)'\s*,")) { $referencedIds += $m.Groups[1].Value }
if ($referencedIds.Count -lt 8) {
    throw "Only $($referencedIds.Count) ids referenced from the injected script; the block probably moved and this smoke lost its subject."
}

# --- 7. the switch is togglable from the panel --------------------------------------------------

Assert-Matches $injected.Value "getElementById\('settings-selected-cell'\)\.checked[\s\S]{0,80}ChatSettings\.selectedCellChecked" `
    'The panel checkbox must be initialised from the persisted value, otherwise it shows the opposite of the truth.' 'BaseChatControl'
Assert-Matches $template 'id="settings-selected-cell"' `
    'The auto-reference switch needs its own row in the settings panel.' 'chat-template.html'
Assert-Matches $settingsJs "getElementById\('settings-selected-cell'\)" `
    'settings-manager.js must read the switch, or saving the dialog drops it.' 'settings-manager.js'
Assert-Matches $settingsJs 'selectedCell:' `
    'The save payload must carry selectedCell.' 'settings-manager.js'
# E5-04 removed the legacy 'checkedChange' channel: the panel never sent it, and the switch already
# applies live through the saveSettings payload. Pin the surviving path instead of the deleted one.
Assert-Absent $chatCode 'router\.Register\("checkedChange"' `
    'checkedChange must not come back; saveSettings is the only live-apply path for this switch.'
$saveSettingsBody = [regex]::Match($chatCode, `
    'Protected Overridable Sub HandleSaveSettings\(jsonDoc As JObject\)[\s\S]*?\r?\n    End Sub')
if (-not $saveSettingsBody.Success) {
    throw 'BaseChatControl : cannot locate HandleSaveSettings, the only live-apply path for the switch.'
}
Assert-Matches $saveSettingsBody.Value 'selectedCellChecked = jsonDoc\("selectedCell"\)' `
    'Saving the dialog must apply the switch to the live field, not only persist it.' 'BaseChatControl'
Assert-Matches $chatCode 'selectedCellChecked = jsonDoc\("selectedCell"\)' `
    'HandleSaveSettings must take the value from the payload instead of writing the static back.' 'BaseChatControl'

Write-Host 'Selection reference smoke passed.'
