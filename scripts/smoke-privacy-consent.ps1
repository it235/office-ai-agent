param(
    [string]$Configuration = "Debug"
)

# E3-05: the privacy acknowledgement existed before this round, but nothing enforced it and the panel could
# only ever send accepted=true. This gate pins the three halves that were actually missing:
#   1) no cloud call without a recorded consent - proven by running the real AiGateway, UtilsService and
#      LLMUtil refusals, the real store write and the real pushed payload, not by grepping a comment;
#   2) the consent is viewable and revocable, and the data-boundary wording has exactly one owner;
#   3) every outbound request site in the repo is registered in one inventory with the reason it is
#      allowed to skip the guard - so "we gated the cloud calls" stays a checkable claim instead of a hope.
# The single-owner rule is data driven: the sentences are read out of PrivacyConsentGate.vb and then counted
# across the whole corpus, so rewording the copy never needs an edit here - while a second copy of the
# promise in VB, JS or HTML turns the gate red.
# The privacy flag is written for real, so secure.config is backed up byte for byte and restored in Finally.
# Keep this file ASCII only: PowerShell 5.1 reads a BOM-less script as ANSI.

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$bin = Join-Path $repoRoot "ShareRibbon\bin\$Configuration"
$dll = Join-Path $bin "ShareRibbon.dll"

if (-not (Test-Path -LiteralPath $dll)) {
    throw "ShareRibbon.dll not found. Run: .\scripts\build-code-projects.ps1 -Configuration $Configuration"
}

$securePath = Join-Path ([Environment]::GetFolderPath("ApplicationData")) "AiHelper\secure.config"

function Read-BytesIfExists($path) {
    if (Test-Path -LiteralPath $path) { return [System.IO.File]::ReadAllBytes($path) }
    return $null
}

function Restore-File($path, $bytes) {
    $dir = Split-Path -Parent $path
    if (-not (Test-Path -LiteralPath $dir)) { [void][System.IO.Directory]::CreateDirectory($dir) }
    if ($null -eq $bytes) {
        if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
    }
    else {
        [System.IO.File]::WriteAllBytes($path, $bytes)
    }
}

function Test-SameBytes($left, $right) {
    if ($null -eq $left -and $null -eq $right) { return $true }
    if ($null -eq $left -or $null -eq $right) { return $false }
    if ($left.Length -ne $right.Length) { return $false }
    for ($i = 0; $i -lt $left.Length; $i++) {
        if ($left[$i] -ne $right[$i]) { return $false }
    }
    return $true
}

# ---------- corpus for the static rules ----------
$sourceByPath = @{}
foreach ($module in @("ShareRibbon", "WordAi", "ExcelAi", "PowerPointAi")) {
    $dir = Join-Path $repoRoot $module
    if (-not (Test-Path -LiteralPath $dir)) { continue }
    foreach ($file in @(Get-ChildItem -Path $dir -Recurse -File -Include *.vb, *.js, *.html |
        Where-Object { $_.FullName -notmatch '\\(bin|obj|packages)\\' })) {
        $sourceByPath[$file.FullName] = Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8
    }
}

$gatePath = Join-Path $repoRoot "ShareRibbon\Controls\Services\PrivacyConsentGate.vb"
$baseChatPath = Join-Path $repoRoot "ShareRibbon\Controls\BaseChatControl.vb"
$onboardingPath = Join-Path $repoRoot "ShareRibbon\Controls\Services\OnboardingService.vb"
$aiGatewayPath = Join-Path $repoRoot "ShareRibbon\Services\Ai\AiGateway.vb"
$streamPath = Join-Path $repoRoot "ShareRibbon\Controls\Services\HttpStreamService.vb"
$storePath = Join-Path $repoRoot "ShareRibbon\Security\SecureConfigManager.vb"
$wizardJsPath = Join-Path $repoRoot "ShareRibbon\Resources\js\onboarding.js"
$ctaJsPath = Join-Path $repoRoot "ShareRibbon\Resources\js\chat-manager.js"
$settingsJsPath = Join-Path $repoRoot "ShareRibbon\Resources\js\settings-manager.js"
$panelHtmlPath = Join-Path $repoRoot "ShareRibbon\Resources\chat-template-refactored.html"
$vbprojPath = Join-Path $repoRoot "ShareRibbon\ShareRibbon.vbproj"
$translatePath = Join-Path $repoRoot "ShareRibbon\Translate\DocumentTranslateService.vb"

# Every place this product can put bytes on the wire, and who is supposed to stop them.
# coverage=guard  -> the file asks PrivacyConsentGate itself before building the request
# coverage=funnel -> reachable only from BaseChatControl.Send / HandleStartAgent, which already ask
# coverage=probe  -> the config window's connectivity test. Its body is the fixed word "Hi", so it
#                    ships no user content, and it is also the tool used to check an endpoint
#                    *before* consent exists - blocking it would make setup impossible.
# coverage=mcp    -> the endpoint is a tool server the user added, not a model interface; the W2 MCP
#                    authorization surface owns it, and it is listed here so nobody can quietly
#                    start routing document content through it without re-deciding this row.
$outboundInventory = @(
    [pscustomobject]@{ Path = 'ShareRibbon\Services\Ai\AiGateway.vb'; Sites = 1; Coverage = 'guard' },
    [pscustomobject]@{ Path = 'ShareRibbon\Controls\BaseChatControl.vb'; Sites = 1; Coverage = 'guard' },
    [pscustomobject]@{ Path = 'ShareRibbon\Controls\Services\UtilsService.vb'; Sites = 1; Coverage = 'guard' },
    [pscustomobject]@{ Path = 'ShareRibbon\Common\LLMUtil.vb'; Sites = 1; Coverage = 'guard' },
    [pscustomobject]@{ Path = 'ShareRibbon\Controls\Services\AutocompleteService.vb'; Sites = 1; Coverage = 'guard' },
    [pscustomobject]@{ Path = 'ShareRibbon\Controls\Services\OfficeCompletionService.vb'; Sites = 2; Coverage = 'guard' },
    [pscustomobject]@{ Path = 'ShareRibbon\Services\EmbeddingService.vb'; Sites = 1; Coverage = 'guard' },
    [pscustomobject]@{ Path = 'ShareRibbon\Translate\DocumentTranslateService.vb'; Sites = 1; Coverage = 'guard' },
    [pscustomobject]@{ Path = 'ExcelAi\ExcelDocumentTranslateService.vb'; Sites = 1; Coverage = 'guard' },
    [pscustomobject]@{ Path = 'ShareRibbon\Controls\Services\HttpStreamService.vb'; Sites = 3; Coverage = 'funnel' },
    [pscustomobject]@{ Path = 'ShareRibbon\Controls\Services\ModelConnectivityTester.vb'; Sites = 1; Coverage = 'probe' },
    [pscustomobject]@{ Path = 'ShareRibbon\Mcp\StreamJsonRpcMCPClient.vb'; Sites = 2; Coverage = 'mcp' }
)

# The project file is not part of the .vb/.js/.html sweep, but rule 1 reads it, so it must be in the corpus.
foreach ($path in @($gatePath, $baseChatPath, $onboardingPath, $aiGatewayPath, $streamPath, $storePath,
                    $wizardJsPath, $ctaJsPath, $settingsJsPath, $panelHtmlPath, $vbprojPath, $translatePath)) {
    if (-not (Test-Path -LiteralPath $path)) { throw "the consent contract references $path but it is not on disk." }
    if (-not $sourceByPath.ContainsKey($path)) {
        $sourceByPath[$path] = Get-Content -LiteralPath $path -Raw -Encoding UTF8
    }
}

# The boundary copy is read out of its owner, never retyped here: a gate that keeps its own copy of the
# wording becomes the second owner it is supposed to be looking for.
$gateBlock = [regex]::Match($sourceByPath[$gatePath], 'BoundaryItems As String\(\) = \{[\s\S]*?\}')
if (-not $gateBlock.Success) { throw "cannot read the BoundaryItems block out of PrivacyConsentGate.vb." }
$boundarySentences = @([regex]::Matches($gateBlock.Value, '"([^"]+)"') | ForEach-Object { $_.Groups[1].Value })
if ($boundarySentences.Count -lt 3) { throw "the data-boundary copy needs at least 3 statements, found $($boundarySentences.Count)." }

function Test-ConsentContract($state) {
    $problems = @()
    $gateSource = $state[$gatePath]
    $baseSource = $state[$baseChatPath]
    $wizardSource = $state[$onboardingPath]
    $gatewaySource = $state[$aiGatewayPath]
    $wizardJs = $state[$wizardJsPath]
    $ctaJs = $state[$ctaJsPath]
    $settingsJs = $state[$settingsJsPath]
    $html = $state[$panelHtmlPath]
    $project = $state[$vbprojPath]

    # 1. a gate class nobody compiles is decoration
    if ($project -notmatch [regex]::Escape('<Compile Include="Controls\Services\PrivacyConsentGate.vb" />')) {
        $problems += 'PrivacyConsentGate.vb is not compiled into ShareRibbon'
    }

    # 2. exactly one reader/writer of the flag: the store class plus this gate. Anyone else is a second opinion.
    foreach ($path in @($state.Keys)) {
        if ($path -eq $gatePath) { continue }
        if ($path -notlike '*.vb') { continue }
        if ($state[$path] -notmatch 'PrivacyPolicyAccepted') { continue }
        if ((Split-Path -Leaf $path) -ne 'SecureConfigManager.vb') {
            $problems += "touches the privacy flag outside its owner: $(Split-Path -Leaf $path)"
        }
    }
    if ($gateSource -notmatch 'SecureConfigManager\.PrivacySettings\.PrivacyPolicyAccepted = ') {
        $problems += 'no longer writes the acknowledgement into the local store'
    }
    if ($gateSource -notmatch 'Return SecureConfigManager\.PrivacySettings\.PrivacyPolicyAccepted') {
        $problems += 'no longer reads the acknowledgement from the local store'
    }

    # 3. the wizard stopped owning the flag and must not start owning the wording again
    if ($wizardSource -notmatch 'PrivacyConsentGate\.Apply\(accepted\)') {
        $problems += 'writes the privacy flag itself again instead of going through the gate'
    }
    if ($wizardSource -notmatch 'state\("privacyBoundary"\) = JArray\.FromObject\(PrivacyConsentGate\.BoundaryItems\)') {
        $problems += 'no longer pushes the boundary copy from its owner'
    }
    if ($wizardJs -notmatch 'state\.privacyBoundary') {
        $problems += 'onboarding.js renders the boundary list from its own copy again'
    }

    # 4. both funnels ask, and what they show is the one blocking card the panel already has
    if ($baseSource -notmatch 'Protected Function IsBlockedByPrivacy\(blockedAction As String\) As Boolean') {
        $problems += 'the shared privacy funnel gate is gone'
    }
    if ($baseSource -notmatch 'IsBlockedByPrivacy\(ModelConfigGate\.ActionChat\)') {
        $problems += 'the chat send funnel no longer asks for consent'
    }
    if ($baseSource -notmatch 'IsBlockedByPrivacy\(ModelConfigGate\.ActionAgent\)') {
        $problems += 'the Agent funnel no longer asks for consent'
    }
    $privacyBlock = [regex]::Match($baseSource, 'Protected Function IsBlockedByPrivacy[\s\S]{0,900}?\r?\n    End Function')
    if (-not $privacyBlock.Success) {
        $problems += 'cannot locate the IsBlockedByPrivacy body'
    }
    else {
        if ($privacyBlock.Value -notmatch 'PrivacyConsentGate\.Evaluate\(PrivacyConsentGate\.IsAccepted\(\), blockedAction\)') {
            $problems += 'the funnel gate decides consent on its own instead of asking the holder'
        }
        if ($privacyBlock.Value -notmatch 'ModelConfigGate\.BuildShowScript\(prompt\)') {
            $problems += 'the consent refusal stopped reusing the shared blocking card'
        }
    }

    # 5. the last-resort stop sits in front of the socket, not behind it
    $guardAt = $gatewaySource.IndexOf('If Not PrivacyConsentGate.IsAccepted() Then', [System.StringComparison]::Ordinal)
    $sendAt = $gatewaySource.IndexOf('_httpClient.SendAsync', [System.StringComparison]::Ordinal)
    if ($guardAt -lt 0) {
        $problems += 'AiGateway no longer refuses a request before any consent exists'
    }
    elseif ($sendAt -lt 0 -or $guardAt -gt $sendAt) {
        $problems += 'the consent refusal is not positioned before the gateway issues its request'
    }

    # 6. the panel branches on the action the host really sends, and can say no as well as yes
    if ($ctaJs -notmatch "data\.action === 'accept_privacy'") {
        $problems += 'chat-manager.js no longer branches the CTA button on the host action'
    }
    if ($ctaJs -notmatch 'setPrivacyConsent\(true\)') {
        $problems += 'the consent card lost the write-back that makes its button mean anything'
    }
    if ($wizardJs -notmatch "accepted: accepted \? 'true' : 'false'") {
        $problems += 'onboarding.js can only ever send consent, so withdrawing it has no path'
    }
    if ($wizardJs -notmatch "type: 'onboardingSetPrivacy'") {
        $problems += 'the consent write no longer goes through the message the host registers'
    }

    # 7. the settings row shows the host's answer, never a guess
    if ($html -notmatch 'id="settings-privacy-consent" disabled') {
        $problems += 'the settings consent row must start disabled until the host reports the real state'
    }
    if ($wizardJs -notmatch "getElementById\('settings-privacy-consent'\)") {
        $problems += 'nothing syncs the settings consent row from the host state'
    }
    if ($wizardJs -notmatch "getElementById\('settings-privacy-boundary'\)") {
        $problems += 'nothing renders the boundary copy into the settings row'
    }
    if ($settingsJs -notmatch "type: 'getOnboardingState'") {
        $problems += 'opening settings before the host push would show an unverified consent default'
    }

    # 8. the wording has exactly one owner across VB, JS and HTML
    foreach ($sentence in $boundarySentences) {
        $hits = 0
        foreach ($path in @($state.Keys)) {
            $hits += ([regex]::Matches($state[$path], [regex]::Escape($sentence))).Count
        }
        if ($hits -ne 1) {
            $problems += "the boundary copy exists $hits time(s) in the corpus while it must exist exactly once"
        }
    }

    # 9. every outbound request site has to declare who gates it, or the promise on the card is a lie.
    # This is the rule that turns "we gate the two panel funnels" into an auditable list: a new
    # HttpClient in any module either lands here with a coverage story, or the gate goes red.
    $observed = @{}
    foreach ($path in @($state.Keys)) {
        if ($path -notlike '*.vb') { continue }
        $count = ([regex]::Matches($state[$path], 'HttpRequestMessage\(HttpMethod\.Post|\.PostAsync\(')).Count
        if ($count -gt 0) { $observed[$path] = $count }
    }
    foreach ($entry in $outboundInventory) {
        $full = Join-Path $repoRoot $entry.Path
        if (-not $observed.ContainsKey($full)) {
            $problems += "the inventory still registers $($entry.Path) but it no longer sends anything out"
            continue
        }
        if ($observed[$full] -ne $entry.Sites) {
            $problems += "$($entry.Path) has $($observed[$full]) outbound request(s) while the inventory says $($entry.Sites)"
        }
        if ($entry.Coverage -eq 'guard' -and $state[$full] -notmatch 'PrivacyConsentGate') {
            $problems += "a guarded cloud exit stopped asking for consent: $($entry.Path)"
        }
        # an exemption is only true as long as the shape it rests on is still the shape on disk
        if ($entry.Coverage -eq 'funnel') {
            $sendEntries = ([regex]::Matches($state[$full], 'Public Async Function Send\w*Async')).Count
            if ($sendEntries -ne 1) {
                $problems += "$($entry.Path) is exempt because one gated send entry reaches it, but it now has $sendEntries"
            }
        }
        if ($entry.Coverage -eq 'probe' -and $state[$full] -notmatch '\{"content", "Hi"\}') {
            $problems += "$($entry.Path) is exempt only while it posts a fixed probe sentence, and that sentence is gone"
        }
        [void]$observed.Remove($full)
    }
    foreach ($extra in @($observed.Keys)) {
        $problems += "an outbound request site is not registered in the consent inventory: $(Split-Path -Leaf $extra) ($($observed[$extra]) site(s))"
    }
    return $problems
}

function Copy-State {
    $copy = @{}
    foreach ($key in @($sourceByPath.Keys)) { $copy[$key] = $sourceByPath[$key] }
    return $copy
}

function Invoke-Probe($name, $expect, $mutate) {
    $copy = Copy-State
    & $mutate $copy
    $found = @(Test-ConsentContract $copy)
    if ($found.Count -eq 0) {
        throw "consent probe '$name' slipped through - the rule it tests is vacuous"
    }
    $text = $found -join ' | '
    if ($text -notmatch [regex]::Escape($expect)) {
        throw "consent probe '$name' was caught by the wrong rule; expected '$expect', got: $text"
    }
    return 1
}

$nl = [Environment]::NewLine
$privacyBody = '        Dim prompt = PrivacyConsentGate.Evaluate(PrivacyConsentGate.IsAccepted(), blockedAction)' + $nl +
               '        If prompt Is Nothing Then Return False' + $nl +
               '        ExecuteJavaScriptAsyncJS(ModelConfigGate.BuildShowScript(prompt))'

$probes = @(
    [pscustomobject]@{ Name = 'gate-not-compiled'; Expect = 'is not compiled into ShareRibbon';
        Apply = { param($s) $s[$vbprojPath] = $s[$vbprojPath].Replace('<Compile Include="Controls\Services\PrivacyConsentGate.vb" />', '') } },
    [pscustomobject]@{ Name = 'gate-not-reads-store'; Expect = 'no longer reads the acknowledgement from the local store';
        Apply = { param($s) $s[$gatePath] = $s[$gatePath].Replace('Return SecureConfigManager.PrivacySettings.PrivacyPolicyAccepted', 'Return True') } },
    [pscustomobject]@{ Name = 'wizard-writes-flag'; Expect = 'writes the privacy flag itself again';
        Apply = { param($s) $s[$onboardingPath] = $s[$onboardingPath].Replace('PrivacyConsentGate.Apply(accepted)', 'SecureConfigManager.PrivacySettings.PrivacyPolicyAccepted = accepted') } },
    [pscustomobject]@{ Name = 'wizard-drops-boundary-push'; Expect = 'no longer pushes the boundary copy';
        Apply = { param($s) $s[$onboardingPath] = $s[$onboardingPath].Replace('state("privacyBoundary") = JArray.FromObject(PrivacyConsentGate.BoundaryItems)', '') } },
    [pscustomobject]@{ Name = 'js-keeps-own-wording'; Expect = 'renders the boundary list from its own copy';
        Apply = { param($s) $s[$wizardJsPath] = $s[$wizardJsPath].Replace('state.privacyBoundary', 'state.privacyCopy') } },
    [pscustomobject]@{ Name = 'chat-funnel-stops-asking'; Expect = 'chat send funnel no longer asks';
        Apply = { param($s) $s[$baseChatPath] = $s[$baseChatPath].Replace('IsBlockedByPrivacy(ModelConfigGate.ActionChat)', 'False') } },
    [pscustomobject]@{ Name = 'agent-funnel-stops-asking'; Expect = 'Agent funnel no longer asks';
        Apply = { param($s) $s[$baseChatPath] = $s[$baseChatPath].Replace('IsBlockedByPrivacy(ModelConfigGate.ActionAgent)', 'False') } },
    [pscustomobject]@{ Name = 'funnel-guesses-consent'; Expect = 'decides consent on its own';
        Apply = { param($s) $s[$baseChatPath] = $s[$baseChatPath].Replace('PrivacyConsentGate.Evaluate(PrivacyConsentGate.IsAccepted(), blockedAction)', 'PrivacyConsentGate.Evaluate(True, blockedAction)') } },
    [pscustomobject]@{ Name = 'consent-gets-its-own-card'; Expect = 'stopped reusing the shared blocking card';
        Apply = { param($s) $s[$baseChatPath] = $s[$baseChatPath].Replace($privacyBody, ($privacyBody.Replace('ModelConfigGate.BuildShowScript(prompt)', 'showConsentRequired(prompt.ToJson())'))) } },
    [pscustomobject]@{ Name = 'gateway-allows-anonymous'; Expect = 'no longer refuses a request before any consent';
        Apply = { param($s) $s[$aiGatewayPath] = $s[$aiGatewayPath].Replace('If Not PrivacyConsentGate.IsAccepted() Then', 'If Not False Then') } },
    [pscustomobject]@{ Name = 'refusal-moved-past-socket'; Expect = 'not positioned before the gateway issues its request';
        Apply = { param($s)
            $block = [regex]::Match($s[$aiGatewayPath], 'If Not PrivacyConsentGate\.IsAccepted\(\) Then[\s\S]{0,300}?End If').Value
            if ([string]::IsNullOrWhiteSpace($block)) { throw 'consent probe cannot find the AiGateway guard block to move' }
            $s[$aiGatewayPath] = $s[$aiGatewayPath].Replace($block, '') + $block } },
    [pscustomobject]@{ Name = 'cta-ignores-action'; Expect = 'branches the CTA button on the host action';
        Apply = { param($s) $s[$ctaJsPath] = $s[$ctaJsPath].Replace("data.action === 'accept_privacy'", "data.action === 'open_other'") } },
    [pscustomobject]@{ Name = 'consent-cannot-be-withdrawn'; Expect = 'can only ever send consent';
        Apply = { param($s) $s[$wizardJsPath] = $s[$wizardJsPath].Replace("accepted: accepted ? 'true' : 'false'", "accepted: 'true'") } },
    [pscustomobject]@{ Name = 'settings-row-unverified'; Expect = 'must start disabled until the host reports';
        Apply = { param($s) $s[$panelHtmlPath] = $s[$panelHtmlPath].Replace('id="settings-privacy-consent" disabled', 'id="settings-privacy-consent"') } },
    [pscustomobject]@{ Name = 'settings-opens-without-state'; Expect = 'unverified consent default';
        Apply = { param($s) $s[$settingsJsPath] = $s[$settingsJsPath].Replace("sendMessageToServer({ type: 'getOnboardingState' });", '') } },
    [pscustomobject]@{ Name = 'wording-copied-into-html'; Expect = 'it must exist exactly once';
        Apply = { param($s) $s[$panelHtmlPath] = $s[$panelHtmlPath] + ('<div>' + $boundarySentences[0] + '</div>') } },
    [pscustomobject]@{ Name = 'new-model-exit-unregistered'; Expect = 'not registered in the consent inventory';
        Apply = { param($s) $s[(Join-Path $repoRoot "ShareRibbon\Services\SomeNewSender.vb")] =
            'Using request As New HttpRequestMessage(HttpMethod.Post, apiUrl)' } },
    [pscustomobject]@{ Name = 'registered-exit-grew-a-socket'; Expect = 'while the inventory says 3';
        Apply = { param($s) $s[$streamPath] = $s[$streamPath] + ' New HttpRequestMessage(HttpMethod.Post, apiUrl)' } },
    [pscustomobject]@{ Name = 'guarded-exit-stops-asking'; Expect = 'a guarded cloud exit stopped asking for consent';
        Apply = { param($s) $s[$translatePath] = $s[$translatePath].Replace('PrivacyConsentGate', 'ConsentWatcher') } },
    [pscustomobject]@{ Name = 'funnel-exemption-grows-a-second-entry'; Expect = 'but it now has 2';
        Apply = { param($s) $s[$streamPath] = $s[$streamPath] + $nl + 'Public Async Function SendSomethingElseAsync()' } },
    [pscustomobject]@{ Name = 'probe-exemption-starts-shipping-text'; Expect = 'that sentence is gone';
        Apply = { param($s)
            $testerPath = Join-Path $repoRoot 'ShareRibbon\Controls\Services\ModelConnectivityTester.vb'
            $s[$testerPath] = $s[$testerPath].Replace('{"content", "Hi"}', '{"content", activeDocumentText}') } }
)

$clean = @(Test-ConsentContract (Copy-State))
if ($clean.Count -gt 0) {
    throw "the consent contract is already broken before any probe: $($clean -join ' | ')"
}

$caught = 0
foreach ($probe in $probes) {
    $caught += Invoke-Probe $probe.Name $probe.Expect $probe.Apply
}

# rule 2 has to catch a writer this gate has never heard of, so hand it one
$secondWriter = Copy-State
$secondWriter[(Join-Path $repoRoot "ShareRibbon\Controls\Services\SomeOtherWriter.vb")] =
    "Public Class SomeOtherWriter" + $nl + "    Public Sub Clear()" + $nl +
    "        SecureConfigManager.PrivacySettings.PrivacyPolicyAccepted = False" + $nl +
    "    End Sub" + $nl + "End Class"
$secondFound = @(Test-ConsentContract $secondWriter)
if ($secondFound.Count -eq 0 -or $secondFound[0] -notmatch 'touches the privacy flag outside its owner') {
    throw "a second writer of the privacy flag was not caught; got: $($secondFound -join ' | ')"
}
$caught++

Push-Location $bin
try {
    Add-Type -Path $dll

    $store = [ShareRibbon.Security.SecureConfigManager+PrivacySettings]
    $script:pushes = @()
    $execute = [System.Func[string, System.Threading.Tasks.Task]] {
        param($js)
        $script:pushes += $js
        return [System.Threading.Tasks.Task]::CompletedTask
    }
    $getAppType = [System.Func[string]] { param() return "word" }
    $onConfigChanged = [System.Action] { }

    # anchors assembled from codepoints - this file must stay ASCII for PowerShell 5.1
    $markNotSent = -join ([char]0x6D88, [char]0x606F)
    $markTask = -join ([char]0x4EFB, [char]0x52A1)

    # --- the card exists only while consent is really missing ---
    $store::PrivacyPolicyAccepted = $false
    if ([ShareRibbon.PrivacyConsentGate]::IsAccepted()) { throw "the gate must report the store, and the store says no." }
    if ($null -ne [ShareRibbon.PrivacyConsentGate]::Evaluate($true, "chat")) {
        throw "an accepted consent must produce no blocking card."
    }

    $chatPrompt = [ShareRibbon.PrivacyConsentGate]::Evaluate($false, "chat")
    $agentPrompt = [ShareRibbon.PrivacyConsentGate]::Evaluate($false, "agent")
    if ($chatPrompt.Reason -ne "privacy_not_accepted") { throw "the reason must be privacy_not_accepted, got '$($chatPrompt.Reason)'." }
    if ($chatPrompt.Action -ne "accept_privacy" -or $agentPrompt.Action -ne "accept_privacy") {
        throw "both funnels must hand the panel the same card action."
    }
    if ($chatPrompt.BlockedAction -ne "chat" -or $agentPrompt.BlockedAction -ne "agent") { throw "the card must know which entry it stopped." }
    foreach ($sentence in $boundarySentences) {
        if ($chatPrompt.Message -notlike "*$sentence*") { throw "the blocking card lost a boundary statement." }
    }
    if ($chatPrompt.Message -notlike "*$markNotSent*") { throw "the chat card must say the message was not sent." }
    if ($agentPrompt.Message -notlike "*$markTask*") { throw "the agent card must say the task did not start." }
    if ($chatPrompt.Message -eq $agentPrompt.Message) { throw "chat and agent refusals must not read identical." }

    # --- what the panel receives still has to carry the discriminator the button branches on ---
    $showScript = [ShareRibbon.ModelConfigGate]::BuildShowScript($chatPrompt)
    if ($showScript -notmatch "^showModelConfigRequired\(") { throw "the consent card must reuse the shared renderer, got '$showScript'." }
    $payloadJson = $showScript.Substring("showModelConfigRequired(".Length)
    $payloadJson = $payloadJson.Substring(0, $payloadJson.Length - 2)
    $payload = ConvertFrom-Json -InputObject $payloadJson
    if ($payload.action -ne "accept_privacy") { throw "the payload lost the action the panel branches on." }
    if ($payload.message -ne $chatPrompt.Message) { throw "the boundary copy did not survive the JSON round trip." }
    if (-not [string]::IsNullOrEmpty([ShareRibbon.ModelConfigGate]::BuildShowScript($null))) {
        throw "no prompt must mean no script."
    }

    $configGate = New-Object ShareRibbon.ModelConfigGate
    $configPrompt = $configGate.Evaluate("", "", "chat")
    if ($configPrompt.Action -ne "open_config") { throw "the config card must keep its own action, got '$($configPrompt.Action)'." }

    # --- consent really lands on disk, and so does the withdrawal ---
    $wizard = New-Object ShareRibbon.OnboardingService -ArgumentList @($execute, $getAppType, $onConfigChanged)
    $doc = New-Object Newtonsoft.Json.Linq.JObject
    [void]$doc.Add("accepted", [Newtonsoft.Json.Linq.JValue]::CreateString("true"))
    $wizard.HandleSetPrivacy($doc)
    if (-not [ShareRibbon.PrivacyConsentGate]::IsAccepted()) { throw "accepting through the wizard must reach the store." }
    if (([System.IO.File]::ReadAllText($securePath)) -notmatch "PrivacyPolicyAccepted=true") {
        throw "the acknowledgement did not land in secure.config."
    }

    $doc2 = New-Object Newtonsoft.Json.Linq.JObject
    [void]$doc2.Add("accepted", [Newtonsoft.Json.Linq.JValue]::CreateString("false"))
    $wizard.HandleSetPrivacy($doc2)
    if ([ShareRibbon.PrivacyConsentGate]::IsAccepted()) { throw "revoking through the settings row must reach the store." }
    if (([System.IO.File]::ReadAllText($securePath)) -notmatch "PrivacyPolicyAccepted=false") {
        throw "the withdrawal did not land in secure.config."
    }

    # --- without consent the shared exit refuses before touching the network ---
    $options = New-Object ShareRibbon.AiRequestOptions
    $options.ApiUrl = "https://127.0.0.1:1/chat/completions"
    $options.ApiKey = "smoke-key-not-real"
    $options.ModelName = "smoke-model"
    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    $refused = [ShareRibbon.AiGateway]::SendChatAsync($options).GetAwaiter().GetResult()
    $watch.Stop()
    if ($refused.Success) { throw "the gateway sent a request while no consent existed." }
    if ($refused.ErrorMessage -notmatch "privacy consent not accepted") { throw "the refusal is not the consent refusal: '$($refused.ErrorMessage)'." }
    if ($watch.ElapsedMilliseconds -gt 3000) { throw "the consent refusal must happen offline, took $($watch.ElapsedMilliseconds)ms." }

    # --- the two other shared senders refuse the same way, and refuse offline ---
    # Rule 9 proves the guard text is in these files; this proves it is the code that actually runs,
    # because a check nobody reaches also greps perfectly.
    $utilsWatch = [System.Diagnostics.Stopwatch]::StartNew()
    $utilsRefused = [ShareRibbon.UtilsService]::SendHttpRequestAsync(
        "https://127.0.0.1:1/chat/completions", "smoke-key-not-real", "{}").GetAwaiter().GetResult()
    $utilsWatch.Stop()
    if (-not [string]::IsNullOrWhiteSpace($utilsRefused)) {
        throw "UtilsService returned a payload while no consent existed: '$utilsRefused'."
    }
    if ($utilsWatch.ElapsedMilliseconds -gt 3000) { throw "the UtilsService refusal must happen offline, took $($utilsWatch.ElapsedMilliseconds)ms." }

    $llmWatch = [System.Diagnostics.Stopwatch]::StartNew()
    $llmRefused = [ShareRibbon.LLMUtil]::SendHttpRequest(
        "https://127.0.0.1:1/chat/completions", "smoke-key-not-real", "{}").GetAwaiter().GetResult()
    $llmWatch.Stop()
    if ($llmRefused -notmatch "privacy consent not accepted") {
        throw "LLMUtil did not refuse on consent, it said: '$llmRefused'."
    }
    if ($llmWatch.ElapsedMilliseconds -gt 3000) { throw "the LLMUtil refusal must happen offline, took $($llmWatch.ElapsedMilliseconds)ms." }

    # a conditional stop, not a blanket one: with consent the same call reaches the (dead) socket
    [ShareRibbon.PrivacyConsentGate]::Apply($true)
    $watch2 = [System.Diagnostics.Stopwatch]::StartNew()
    $allowed = [ShareRibbon.AiGateway]::SendChatAsync($options).GetAwaiter().GetResult()
    $watch2.Stop()
    if ($allowed.ErrorMessage -match "privacy consent") { throw "a consented request was still refused by the consent gate." }
    if ($watch2.ElapsedMilliseconds -gt 8000) { throw "the refused-socket case hung instead of failing fast." }

    # --- the pushed payload carries the flag and the owner's wording ---
    $count = $script:pushes.Count
    $wizard.HandleGetState()
    if ($script:pushes.Count -ne $count + 1) { throw "HandleGetState must push exactly one payload." }
    $push = $script:pushes[$count]
    if ($push -notmatch "^renderOnboarding\(") { throw "the wizard must push renderOnboarding(...), got '$push'." }
    $stateJson = $push.Substring("renderOnboarding(".Length)
    $stateJson = $stateJson.Substring(0, $stateJson.Length - 2)
    $state = ConvertFrom-Json -InputObject $stateJson
    if ($state.privacyAccepted -ne $true) { throw "the pushed state must report the real consent flag." }
    $pushedBoundary = @($state.privacyBoundary)
    if ($pushedBoundary.Count -ne $boundarySentences.Count) {
        throw "the wizard pushed $($pushedBoundary.Count) boundary statements but the owner has $($boundarySentences.Count)."
    }
    for ($i = 0; $i -lt $boundarySentences.Count; $i++) {
        if ($pushedBoundary[$i] -ne $boundarySentences[$i]) { throw "boundary statement #$i is not the owner's wording." }
    }

    Write-Host "OK: consent gates every registered cloud exit, is revocable, and the boundary copy has one owner."
    [pscustomobject]@{
        BoundarySentences = $boundarySentences.Count
        Rules             = 9
        ProbesRed         = $caught
        OutboundSites     = (@($outboundInventory | ForEach-Object { $_.Sites }) | Measure-Object -Sum).Sum
        GuardedFiles      = @($outboundInventory | Where-Object { $_.Coverage -eq 'guard' }).Count
        CardReason        = $chatPrompt.Reason
        CardAction        = $chatPrompt.Action
        RefusedMs         = $watch.ElapsedMilliseconds
        UtilsRefusedMs    = $utilsWatch.ElapsedMilliseconds
        LlmRefusedMs      = $llmWatch.ElapsedMilliseconds
        ConsentMs         = $watch2.ElapsedMilliseconds
    } | Format-List
}
finally {
    Restore-File $securePath $secureBackup
    Pop-Location
}

if (-not (Test-SameBytes $secureBackup (Read-BytesIfExists $securePath))) {
    throw "the consent smoke rewrote secure.config; the byte restore failed."
}
