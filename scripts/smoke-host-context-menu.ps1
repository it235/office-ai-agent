$ErrorActionPreference = 'Stop'

# E4-02 host context menu smoke.
# The catalog is checked by actually calling the compiled assembly, so "three actions per host"
# is a behaviour fact and not a grep. The rest are source invariants that do not show up as a
# compile error when they rot: menu ownership (Tag), reversibility, the single Agent entry,
# the default-off switch and the privacy rule.

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')

function Read-Source([string]$relativePath) {
    $path = Join-Path $repoRoot $relativePath
    if (-not (Test-Path -LiteralPath $path)) {
        throw "Expected file is missing: $relativePath"
    }
    return Get-Content -LiteralPath $path -Raw -Encoding UTF8
}

# Comments are where the forbidden mechanisms get explained, so a "this must not appear" rule has
# to read code only. VB drops a line starting with a single quote.
function Get-CodeOnly([string]$source) {
    return (($source -split "`r?`n") | Where-Object { $_ -notmatch "^\s*'" }) -join "`n"
}

$shared = Get-CodeOnly (Read-Source 'ShareRibbon\Agent\HostContextMenu.vb')
$catalog = Read-Source 'ShareRibbon\Agent\QuickActionCatalog.vb'
$settings = Read-Source 'ShareRibbon\Config\ChatSettings.vb'
$chatControl = Read-Source 'ShareRibbon\Controls\BaseChatControl.vb'
$shareProj = Read-Source 'ShareRibbon\ShareRibbon.vbproj'

$hosts = @(
    [pscustomobject]@{
        Name        = 'Word'
        Adapter     = 'WordAi\HostEntry\WordContextMenu.vb'
        AddIn       = 'WordAi\ThisAddIn.vb'
        Ribbon      = 'WordAi\Ribbon1.vb'
        Project     = 'WordAi\WordAi.vbproj'
        CompileItem = 'HostEntry\WordContextMenu.vb'
        StartupSub  = 'WordAi_Startup'
        AppType     = 'WordAppType'
        IdPrefix    = 'word\.'
    },
    [pscustomobject]@{
        Name        = 'Excel'
        Adapter     = 'ExcelAi\HostEntry\ExcelContextMenu.vb'
        AddIn       = 'ExcelAi\ThisAddIn.vb'
        Ribbon      = 'ExcelAi\Ribbon1.vb'
        Project     = 'ExcelAi\ExcelAi.vbproj'
        CompileItem = 'HostEntry\ExcelContextMenu.vb'
        StartupSub  = 'ExcelAi_Startup'
        AppType     = 'ExcelAppType'
        IdPrefix    = 'excel\.'
    },
    [pscustomobject]@{
        Name        = 'PowerPoint'
        Adapter     = 'PowerPointAi\HostEntry\PowerPointContextMenu.vb'
        AddIn       = 'PowerPointAi\ThisAddIn.vb'
        Ribbon      = 'PowerPointAi\Ribbon1.vb'
        Project     = 'PowerPointAi\PowerPointAi.vbproj'
        CompileItem = 'HostEntry\PowerPointContextMenu.vb'
        StartupSub  = 'PowerPointAi_Startup'
        AppType     = 'PowerPointAppType'
        IdPrefix    = 'ppt\.'
    }
)

# --- 1. behaviour: the compiled catalog really serves three honest actions per host ----------

$dll = Join-Path $repoRoot "ShareRibbon\bin\Debug\ShareRibbon.dll"
if (-not (Test-Path -LiteralPath $dll)) {
    throw "ShareRibbon assembly not found: $dll. Build the code projects first."
}
Add-Type -LiteralPath $dll

foreach ($host_ in $hosts) {
    $actions = @([ShareRibbon.Agent.QuickActionCatalog]::ActionsFor($host_.Name))
    if ($actions.Count -ne 3) {
        throw "$($host_.Name) must get exactly 3 context menu actions, found $($actions.Count)."
    }

    $ids = @()
    foreach ($action in $actions) {
        if ([string]::IsNullOrWhiteSpace($action.Id) -or
            [string]::IsNullOrWhiteSpace($action.Label) -or
            [string]::IsNullOrWhiteSpace($action.RequestText)) {
            throw "$($host_.Name) action $($action.Id) has an empty id, label or request text."
        }
        if ($action.Id -notmatch "^$($host_.IdPrefix)") {
            throw "$($host_.Name) action id must be namespaced '$($host_.IdPrefix)*': $($action.Id)"
        }
        if ($ids -contains $action.Id) {
            throw "$($host_.Name) declares the same action id twice: $($action.Id)"
        }
        $ids += $action.Id
        if ($host_.Name -ne 'Word' -and -not [string]::IsNullOrWhiteSpace($action.CapabilityId)) {
            # Excel/PowerPoint have no host capability registry yet (open decision #36), so an
            # entry there must not claim one.
            throw "$($host_.Name) action $($action.Id) declares CapabilityId '$($action.CapabilityId)' " +
                  "while that host has no capability registry to prove it."
        }
    }
    Write-Host "$($host_.Name) actions: $($ids -join ', ')"
}

# An unhosted app type must return nothing rather than borrow another host's wording.
if (@([ShareRibbon.Agent.QuickActionCatalog]::ActionsFor('Outlook')).Count -ne 0) {
    throw 'QuickActionCatalog must return an empty list for a host it does not serve.'
}

# --- 2. shared layer stays host-agnostic and owns only what it can undo ----------------------

if ($shared -notmatch '(?m)^\s*Namespace Agent\s*$') {
    throw 'HostContextMenu must live in Namespace Agent (ShareRibbon root).'
}
if ($shared -match 'Microsoft\.Office\.Interop\.(Word|Excel|PowerPoint)') {
    throw 'ShareRibbon must not reference host COM types; only Office.CommandBar level is allowed.'
}
if ($shared -notmatch 'TagPrefix As String = "aiq:"') {
    throw 'Menu items need a Tag prefix: it is the only proof of which controls this add-in owns.'
}
if ($shared -notmatch 'Temporary:=True') {
    throw 'Context menu controls must be added with Temporary:=True so they never persist in the user profile.'
}
if ($shared -notmatch 'For i = _bar\.Controls\.Count To 1 Step -1') {
    throw 'Removing owned controls must walk backwards; forward enumeration of a COM collection skips items.'
}
if ($shared -notmatch 'control\.Tag\.StartsWith\(TagPrefix') {
    throw 'Detach must delete by Tag prefix, never by caption text.'
}
if ($shared -notmatch '_buttons\.Add\(button\)') {
    throw 'Button RCWs must be retained, otherwise the click sink is collected and the menu goes dead.'
}
if ($shared -notmatch 'RemoveHandler button\.Click') {
    throw 'Detach must unhook every click handler it added.'
}

# The pinned Office 15 primary interop exposes no event surface on CommandBar at all.
# Reaching for BeforeRightClick therefore cannot compile, and any replacement must stay
# click-time so the menu never shows a stale grayed-out item.
if ($shared -match 'BeforeRightClick|CommandBarEvents') {
    throw 'BeforeRightClick is not available in the referenced Office PIA; availability is judged at click time.'
}
if ($shared -notmatch 'GlobalStatusStripAll\.ShowProgress') {
    throw 'A blocked action must answer with the reason on the status bar instead of launching the Agent.'
}
foreach ($bypass in @('sendMessageToServer', 'ExecuteJavaScript', 'createChatSection')) {
    if ($shared -match $bypass) {
        throw "The shared menu layer builds panel JS ($bypass); it must only invoke the injected request callback."
    }
}

# --- 3. each host: adapter, wiring and lifecycle --------------------------------------------

foreach ($host_ in $hosts) {
    $adapter = Get-CodeOnly (Read-Source $host_.Adapter)
    $addIn = Get-CodeOnly (Read-Source $host_.AddIn)
    $ribbon = Get-CodeOnly (Read-Source $host_.Ribbon)
    $proj = Read-Source $host_.Project

    if ($proj -notmatch [regex]::Escape($host_.CompileItem)) {
        throw "$($host_.Project) does not compile $($host_.CompileItem)"
    }
    # VB resolves a namespace over a same-named interop type, so a "Selection" namespace here
    # breaks unrelated PowerPoint files. The host entry layer must never use a type-like name.
    if ($adapter -match '(?m)^\s*Namespace Selection\s*$') {
        throw "$($host_.Adapter) must use Namespace HostEntry: 'Selection' collides with the interop type name."
    }
    if ($adapter -notmatch '(?m)^\s*Namespace HostEntry\s*$') {
        throw "$($host_.Adapter) must live in Namespace HostEntry."
    }
    if ($adapter -notmatch 'New HostContextMenu\(\)') {
        throw "$($host_.Adapter) must delegate registration to the shared HostContextMenu layer."
    }
    if ($adapter -notmatch "QuickActionCatalog\.ActionsFor\(QuickActionCatalog\.$($host_.AppType)\)") {
        throw "$($host_.Adapter) must take its actions from QuickActionCatalog for $($host_.AppType)."
    }
    if ($adapter -notmatch '_menu\.Attach\(bar,[\s\S]{0,200}AddressOf CurrentAvailability\)') {
        throw "$($host_.Adapter) must hand the shared layer its own availability probe."
    }
    if ($adapter -notmatch '\.Enabled = False[\s\S]{0,160}\.Reason = "([^"]+)"') {
        throw "$($host_.Adapter) must explain a degraded click with a reason, not just refuse."
    }
    # Host adapters only pick the bar and probe the selection; COM event plumbing stays in one place.
    if ($adapter -match 'AddHandler') {
        throw "$($host_.Adapter) must not wire COM events itself; the shared layer owns handler symmetry."
    }
    foreach ($bypass in @('sendMessageToServer', 'ExecuteJavaScript')) {
        if ($adapter -match $bypass) {
            throw "$($host_.Adapter) talks to the panel JS ($bypass) instead of the single Agent entry."
        }
    }

    if ($addIn -notmatch 'ChatSettings\.EnableContextMenu Then Return') {
        throw "$($host_.AddIn) must keep the context menu behind the ChatSettings.EnableContextMenu switch."
    }
    if ($addIn -notmatch 'contextMenu\.Attach\(AddressOf OnContextMenuAction\)') {
        throw "$($host_.AddIn) must attach the menu through one callback that returns to the panel."
    }
    if ($addIn -notmatch 'OnContextMenuAction\(action As ShareRibbon\.Agent\.QuickAction\)') {
        throw "$($host_.AddIn) menu callback must receive the catalog action, not a hand-written string."
    }
    if ($addIn -notmatch 'chatControl\.StartAgentFromHostEntry\(action\)') {
        throw "$($host_.AddIn) must run the request through StartAgentFromHostEntry, the single Agent entry."
    }
    if ($addIn -notmatch 'contextMenu\.Detach\(\)') {
        throw "$($host_.AddIn) must remove its menu items on shutdown."
    }

    $startupBody = [regex]::Match($addIn, "Private Sub $($host_.StartupSub)[\s\S]*?\r?\n    End Sub")
    if (-not $startupBody.Success) {
        throw "Cannot locate the $($host_.StartupSub) body in $($host_.AddIn)."
    }
    if ($startupBody.Value -match 'EnsureContextMenu') {
        throw 'The context menu must be attached when the panel opens, not in Startup: startup cost and ' +
              'a floating rewrite of the user menu for users who never open the panel.'
    }
    $visibleBody = [regex]::Match($addIn, 'Private Sub ChatTaskPane_VisibleChanged[\s\S]*?\r?\n    End Sub')
    if (-not $visibleBody.Success) {
        throw "Cannot locate ChatTaskPane_VisibleChanged in $($host_.AddIn)."
    }
    if ($visibleBody.Value -notmatch 'EnsureContextMenu\(\)') {
        throw "$($host_.AddIn) must attach the menu from the panel-visible path, so a click always has a listener."
    }

    # Ribbon buttons and the menu share one entry, so neither can grow a private execution path.
    $ribbonHelper = [regex]::Match($ribbon, 'Private Async Function StartAgentFromRibbonAsync[\s\S]*?\r?\n    End Function')
    if (-not $ribbonHelper.Success) {
        throw "Cannot locate StartAgentFromRibbonAsync in $($host_.Ribbon)."
    }
    if ($ribbonHelper.Value -match 'sendMessageToServer') {
        throw "$($host_.Ribbon) still injects panel JS to start the Agent; use StartAgentFromHostEntry."
    }
    if ($ribbonHelper.Value -notmatch 'StartAgentFromHostEntry\(') {
        throw "$($host_.Ribbon) must start the Agent through StartAgentFromHostEntry."
    }

    foreach ($file in @(@{ Path = $host_.Adapter; Text = $adapter }, @{ Path = $host_.AddIn; Text = $addIn })) {
        foreach ($line in ($file.Text -split "`n")) {
            if ($line -match 'AppLogger\.\w+\(' -and $line -notmatch 'Redact') {
                if ($line -match '\.Text\b|RequestText|Selection\.Text|\.Caption') {
                    throw "$($file.Path) logs document text: $($line.Trim())"
                }
            }
        }
    }
}

# --- 4. one switch, off by default, and actually persisted ----------------------------------

if ($settings -notmatch 'EnableContextMenu As Boolean = False') {
    throw 'ChatSettings.EnableContextMenu must default to False: a rewritten host menu must never arrive silently.'
}
if ($settings -notmatch '"enableContextMenu"') {
    throw 'The switch must be persisted as enableContextMenu in office_ai_chat_settings.json.'
}
if ($settings -notmatch 'ContainsKey\("enableContextMenu"\)') {
    throw 'ChatSettings must read enableContextMenu back from the settings file.'
}
if ($chatControl -notmatch 'ChatSettings\.EnableContextMenu,?\s*\)') {
    throw 'HandleSaveSettings must write ChatSettings.EnableContextMenu back, otherwise the switch is read-only.'
}
if ($shareProj -notmatch [regex]::Escape('Agent\HostContextMenu.vb')) {
    throw 'ShareRibbon.vbproj does not compile Agent\HostContextMenu.vb'
}
if ($catalog -notmatch 'ExcelAppType|PowerPointAppType') {
    throw 'QuickActionCatalog must keep the per-host app type constants used by every native entry.'
}

foreach ($file in @('ShareRibbon\Agent\HostContextMenu.vb', 'ShareRibbon\Agent\QuickActionCatalog.vb',
                    'WordAi\HostEntry\WordContextMenu.vb', 'ExcelAi\HostEntry\ExcelContextMenu.vb',
                    'PowerPointAi\HostEntry\PowerPointContextMenu.vb')) {
    $lines = (Get-Content -LiteralPath (Join-Path $repoRoot $file) -Encoding UTF8).Count
    if ($lines -gt 500) {
        throw "$file is $lines lines; the project rule caps a class at 500."
    }
}

Write-Host 'Host context menu smoke passed.'
