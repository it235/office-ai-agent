param(
    [string]$Configuration = "Debug"
)

# E3-04: the panel showed a raw token count for streamed chat answers and nothing else - no per-run usage,
# no price, no budget, and the non-streaming Agent path did not read usage at all. This gate pins the four
# things that had to be true for that to stop being decoration:
#   1) one parser reads usage out of a provider response (OpenAI keys and Anthropic keys), and it is the only
#      place in the repo that spells those key names;
#   2) one ledger: every send - streamed or not - writes through ChatStateService.RecordUsage, so a run cost
#      and a session cost cannot come from different counts;
#   3) money only appears when the user configured a price: with no price the copy says so and shows no
#      amount, and no JS may compute or format a cost, otherwise two wordings drift apart;
#   4) the budget is a notice, not a gate - only consent blocks a request (E3-05).
# Runtime half loads the compiled DLL and runs the real parser, ledger, formatter and settings round trip.
# The chat settings file is written for real, so it is backed up byte for byte and restored in Finally.
# Keep this file ASCII only: PowerShell 5.1 reads a BOM-less script as ANSI, so a Chinese literal here would
# compare against mojibake instead of the strings the VB layer returns.

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$bin = Join-Path $repoRoot "ShareRibbon\bin\$Configuration"
$dll = Join-Path $bin "ShareRibbon.dll"

if (-not (Test-Path -LiteralPath $dll)) {
    throw "ShareRibbon.dll not found. Run: .\scripts\build-code-projects.ps1 -Configuration $Configuration"
}

$documents = [Environment]::GetFolderPath("MyDocuments")
$settingsPathOnDisk = Join-Path $documents "OfficeAiAppData\office_ai_chat_settings.json"

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
    foreach ($file in @(Get-ChildItem -Path $dir -Recurse -File -Include *.vb, *.js, *.html, *.vbproj |
        Where-Object { $_.FullName -notmatch '\\(bin|obj|packages)\\' })) {
        $sourceByPath[$file.FullName] = Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8
    }
}

$usagePath = Join-Path $repoRoot "ShareRibbon\Services\Ai\TokenUsage.vb"
$reporterPath = Join-Path $repoRoot "ShareRibbon\Controls\Services\UsageCostReporter.vb"
$chatStatePath = Join-Path $repoRoot "ShareRibbon\Controls\Services\ChatStateService.vb"
$settingsPath = Join-Path $repoRoot "ShareRibbon\Config\ChatSettings.vb"
$baseChatPath = Join-Path $repoRoot "ShareRibbon\Controls\BaseChatControl.vb"
$streamPath = Join-Path $repoRoot "ShareRibbon\Controls\Services\HttpStreamService.vb"
$kernelPath = Join-Path $repoRoot "ShareRibbon\Controls\Services\AgentKernelService.vb"
$aiGatewayPath = Join-Path $repoRoot "ShareRibbon\Services\Ai\AiGateway.vb"
$panelHtmlPath = Join-Path $repoRoot "ShareRibbon\Resources\chat-template-refactored.html"
$settingsJsPath = Join-Path $repoRoot "ShareRibbon\Resources\js\settings-manager.js"
$cardJsPath = Join-Path $repoRoot "ShareRibbon\Resources\js\agent-card.js"
$codeJsPath = Join-Path $repoRoot "ShareRibbon\Resources\js\code-handler.js"
$vbprojPath = Join-Path $repoRoot "ShareRibbon\ShareRibbon.vbproj"

foreach ($path in @($usagePath, $reporterPath, $chatStatePath, $settingsPath, $baseChatPath, $streamPath,
                    $kernelPath, $aiGatewayPath, $panelHtmlPath, $settingsJsPath, $cardJsPath, $codeJsPath, $vbprojPath)) {
    if (-not (Test-Path -LiteralPath $path)) { throw "the usage contract references $path but it is not on disk." }
}

# The four keys travel through five places each; list them once and let the rule walk the table so adding a
# fifth setting cannot mean forgetting one of the five links.
$keyInventory = @(
    [pscustomobject]@{ Name = 'pricePerMillionIn';  DomId = 'settings-price-in';      Reader = 'ReadNumberSetting' },
    [pscustomobject]@{ Name = 'pricePerMillionOut'; DomId = 'settings-price-out';     Reader = 'ReadNumberSetting' },
    [pscustomobject]@{ Name = 'priceCurrency';      DomId = 'settings-price-currency'; Reader = 'ReadString' },
    [pscustomobject]@{ Name = 'sessionTokenBudget'; DomId = 'settings-session-budget'; Reader = 'ReadIntegerSetting' }
)

# Wording is read out of the reporter, so rewording the copy never needs an edit here - while a second copy
# of the same promise in VB, JS or HTML turns the gate red.
function Get-OwnerWording($state) {
    $items = @()
    foreach ($m in [regex]::Matches($state[$reporterPath], 'Public Const (\w+) As String = "([^"]+)"')) {
        $items += [pscustomobject]@{ Name = $m.Groups[1].Value; Text = $m.Groups[2].Value }
    }
    return $items
}

function Test-UsageContract($state) {
    $problems = @()

    # 1. both new files must actually be compiled in - an orphan class proves nothing at runtime
    foreach ($compile in @('Services\Ai\TokenUsage.vb', 'Controls\Services\UsageCostReporter.vb')) {
        if ($state[$vbprojPath] -notmatch ('<Compile Include="' + [regex]::Escape($compile) + '" />')) {
            $problems += "$compile is not compiled into ShareRibbon"
        }
    }

    # 2. one parser owns the provider key names - a quoted index into a usage block is a re-parse
    foreach ($path in @($state.Keys)) {
        if ($path -eq $usagePath) { continue }
        if ($state[$path] -match '"(prompt_tokens|completion_tokens|input_tokens|output_tokens)"\)') {
            $problems += "$(Split-Path -Leaf $path) parses provider usage keys itself; TokenUsage.vb is the only owner"
        }
    }

    # 3. one ledger, one writer
    foreach ($path in @($state.Keys)) {
        if ($state[$path] -match '\bAddTokens\b') {
            $problems += "$(Split-Path -Leaf $path) still writes the session count through AddTokens; RecordUsage is the only entry"
        }
    }
    $callSites = 0
    foreach ($path in @($state.Keys)) {
        $callSites += ([regex]::Matches($state[$path], '\.RecordUsage\(')).Count
    }
    if ($callSites -ne 3) {
        $problems += "exactly 3 call sites may record usage (stream finalize, non-stream sender, multimodal), got $callSites"
    }
    if ($state[$chatStatePath] -notmatch 'Public Sub RecordUsage\(usage As TokenUsage\)') {
        $problems += "ChatStateService no longer defines RecordUsage - the ledger lost its owner"
    }
    if ($state[$chatStatePath] -notmatch 'CurrentSessionTotalTokens[\s\S]{0,160}Return _sessionUsage\.TotalTokens') {
        $problems += "the session token read-out must come from the same ledger the usage is recorded into"
    }
    # the streaming translator's last frame must be the parser's own type, not a second three-integer shape
    foreach ($path in @($state.Keys)) {
        if ($state[$path] -match '\bTokenInfo\b') {
            $problems += "$(Split-Path -Leaf $path) still carries TokenInfo; the ledger holds the same TokenUsage the parser returns"
        }
    }

    # 4. wording has one owner (data driven: the sentences are read out of the reporter)
    foreach ($item in @(Get-OwnerWording $state)) {
        $hits = 0
        foreach ($path in @($state.Keys)) {
            $hits += ([regex]::Matches($state[$path], [regex]::Escape($item.Text))).Count
        }
        if ($hits -ne 1) {
            $problems += "usage copy '$($item.Name)' occurs $hits time(s); it must exist exactly once"
        }
    }

    # 5. no money may be computed or worded in the panel
    $costPhrase = -join ([char]0x4F30, [char]0x7B97, [char]0x8D39, [char]0x7528)
    if ($state[$reporterPath] -notmatch [regex]::Escape($costPhrase)) {
        $problems += "the reporter no longer owns the cost sentence, so the rule below is vacuous"
    }
    foreach ($path in @($state.Keys)) {
        if ($path -notmatch '\.(js|html)$') { continue }
        if ($state[$path] -match [regex]::Escape($costPhrase)) {
            $problems += "$(Split-Path -Leaf $path) words the cost itself; the amount and its sentence come from UsageCostReporter"
        }
        if ($state[$path] -match '/\s*1000000') {
            $problems += "$(Split-Path -Leaf $path) divides by a million token price in the panel - that is a second calculator"
        }
    }

    # 6. every settings key is written, read back, sent by the panel, handled by the host, present in the DOM
    #    and seeded from the host, or the row is decoration the user cannot actually set
    foreach ($row in $keyInventory) {
        if ($state[$settingsPath] -notmatch ('settings\("' + $row.Name + '"\)')) {
            $problems += "ChatSettings.SaveSettings does not write $($row.Name), so saving the row loses it"
        }
        if ($state[$settingsPath] -notmatch ('settings, "' + $row.Name + '"')) {
            $problems += "ChatSettings.LoadSettings does not read $($row.Name) back"
        }
        if ($state[$settingsJsPath] -notmatch ('(' + $row.Name + ':|"' + $row.Name + '")')) {
            $problems += "the settings payload never carries $($row.Name) to the host"
        }
        if ($row.Reader -eq 'ReadString') {
            if ($state[$baseChatPath] -notmatch ('jsonDoc\("' + $row.Name + '"\)')) {
                $problems += "HandleSaveSettings does not read $($row.Name) from the payload - it would keep writing the static value back"
            }
        }
        elseif ($state[$baseChatPath] -notmatch ($row.Reader + '\(jsonDoc, "' + $row.Name + '"')) {
            $problems += "HandleSaveSettings does not read $($row.Name) through $($row.Reader)"
        }
        if ($state[$panelHtmlPath] -notmatch ('id="' + $row.DomId + '"')) {
            $problems += "the panel has no input #$($row.DomId) for $($row.Name)"
        }
        if ($state[$baseChatPath] -notmatch ("setNativeEntrySetting\('" + $row.DomId + "'")) {
            $problems += "the row #$($row.DomId) is never seeded from the stored value, so the panel would show an unset default"
        }
    }

    # 7. the terminal card must carry usage even on the paths that return early (no changes / not undoable)
    if ($state[$kernelPath] -notmatch '\{"usage", BuildUsageBlock\(\)\}') {
        $problems += "the Agent terminal descriptor no longer carries a usage block"
    }
    $usageAt = $state[$kernelPath].IndexOf('{"usage", BuildUsageBlock()}')
    $firstReturn = $state[$kernelPath].IndexOf('Return descriptor')
    if ($usageAt -lt 0 -or $firstReturn -lt 0 -or $usageAt -gt $firstReturn) {
        $problems += "the usage block must be attached before the first early return, or a read-only run shows no cost"
    }
    if ($state[$kernelPath] -notmatch 'UsageCostReporter\.SummaryText\(') {
        $problems += "BuildUsageBlock must take its wording from UsageCostReporter"
    }

    # 8. the streamed answer's footer note goes through the same reporter and the same escape helper
    if ($state[$streamPath] -notmatch 'UsageCostReporter\.SummaryText\(') {
        $problems += "the stream footer no longer takes its wording from UsageCostReporter"
    }
    $footerLine = @($state[$streamPath] -split "\r?\n") | Where-Object { $_ -match 'processStreamComplete\(' } | Select-Object -First 1
    if (-not $footerLine) {
        $problems += "HttpStreamService no longer completes a stream - the footer rule below is vacuous"
    }
    elseif ($footerLine -notmatch 'usageNote') {
        $problems += "the stream footer no longer passes a reporter-built note to the panel"
    }
    if ($state[$streamPath] -notmatch 'ToJsStringLiteral\(usageNote\)') {
        $problems += "the footer note reaches the script text unescaped"
    }
    if ($state[$codeJsPath] -notmatch 'function processStreamComplete\(uuid, totalTokens, usageNote\)') {
        $problems += "processStreamComplete must accept the usage note the host sends"
    }
    if ($state[$codeJsPath] -notmatch 'noteSpan\.textContent') {
        $problems += "the footer note must be written with textContent, never as markup"
    }

    # 9. the budget is a notice. Only consent blocks a request - if this file ever starts refusing sends,
    #    that is a second gate and needs its own decision, not a quiet insertion.
    if ($state[$reporterPath] -match 'IsBlockedByPrivacy|RefuseQuietly|Throw') {
        $problems += "UsageCostReporter is blocking or throwing; the budget may only warn"
    }
    # Count the asks, not the definition: every block site names the action it blocks.
    $blockSites = ([regex]::Matches($state[$baseChatPath], 'IsBlockedByPrivacy\(ModelConfigGate\.')).Count
    if ($blockSites -ne 3) {
        $problems += "BaseChatControl must keep exactly 3 consent block sites (chat send, agent start, non-stream sender), got $blockSites"
    }

    # 10. the run card renders usage through textContent only
    if ($state[$cardJsPath] -notmatch 'usageLine\.textContent') {
        $problems += "the terminal card must render the host's usage sentence with textContent"
    }
    if ($state[$cardJsPath] -match 'innerHTML[^;]*usage') {
        $problems += "the usage line is being injected as markup"
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
    $found = @(Test-UsageContract $copy)
    if ($found.Count -eq 0) {
        throw "usage probe '$name' slipped through - the rule it tests is vacuous"
    }
    $text = $found -join ' | '
    if ($text -notmatch [regex]::Escape($expect)) {
        throw "usage probe '$name' was caught by the wrong rule; expected '$expect', got: $text"
    }
    return 1
}

$probes = @(
    [pscustomobject]@{ Name = 'parser-not-compiled'; Expect = 'is not compiled into ShareRibbon';
        Apply = { param($s) $s[$vbprojPath] = $s[$vbprojPath].Replace('<Compile Include="Services\Ai\TokenUsage.vb" />', '') } },
    [pscustomobject]@{ Name = 'stream-reparses-usage'; Expect = 'parses provider usage keys itself';
        Apply = { param($s) $s[$streamPath] = $s[$streamPath] + "`r`n" + 'Dim p = CInt(jsonObj("usage")("prompt_tokens"))' } },
    [pscustomobject]@{ Name = 'second-ledger-writer'; Expect = 'RecordUsage is the only entry';
        Apply = { param($s) $s[$kernelPath] = $s[$kernelPath].Replace('_chatStateService?.BeginRunUsage()', '_chatStateService.AddTokens(10)') } },
    [pscustomobject]@{ Name = 'fourth-usage-caller'; Expect = 'exactly 3 call sites may record usage';
        Apply = { param($s) $s[$aiGatewayPath] = $s[$aiGatewayPath] + "`r`n" + 'state.RecordUsage(u)' } },
    [pscustomobject]@{ Name = 'session-total-detached'; Expect = 'the same ledger the usage is recorded into';
        Apply = { param($s) $s[$chatStatePath] = $s[$chatStatePath].Replace('Return _sessionUsage.TotalTokens', 'Return 0') } },
    [pscustomobject]@{ Name = 'second-usage-shape-returning'; Expect = 'the ledger holds the same TokenUsage the parser returns';
        Apply = { param($s) $s[$chatStatePath] = $s[$chatStatePath].Replace('Public Property LastUsage As TokenUsage', 'Public Property LastUsage As TokenInfo') } },
    [pscustomobject]@{ Name = 'wording-copied-into-js'; Expect = 'it must exist exactly once';
        Apply = { param($s) $s[$cardJsPath] = $s[$cardJsPath] + ('/* ' + (Get-OwnerWording $s)[0].Text + ' */') } },
    [pscustomobject]@{ Name = 'js-computes-cost'; Expect = 'divides by a million token price in the panel';
        Apply = { param($s) $s[$cardJsPath] = $s[$cardJsPath].Replace("box.appendChild(usageLine);", "const per = tokens/1000000; box.appendChild(usageLine);") } },
    [pscustomobject]@{ Name = 'price-key-not-persisted'; Expect = 'so saving the row loses it';
        Apply = { param($s) $s[$settingsPath] = $s[$settingsPath].Replace('settings("pricePerMillionIn")', 'settings("priceIn")') } },
    [pscustomobject]@{ Name = 'price-key-not-read-back'; Expect = 'does not read pricePerMillionOut back';
        Apply = { param($s) $s[$settingsPath] = $s[$settingsPath].Replace('ReadDouble(settings, "pricePerMillionOut"', 'ReadDouble(settings, "priceOut"') } },
    [pscustomobject]@{ Name = 'panel-never-sends-budget'; Expect = 'the settings payload never carries';
        Apply = { param($s) $s[$settingsJsPath] = $s[$settingsJsPath].Replace('sessionTokenBudget: document', 'budgetSentElsewhere: document') } },
    [pscustomobject]@{ Name = 'handler-writes-static-back'; Expect = 'would keep writing the static value back';
        Apply = { param($s) $s[$baseChatPath] = $s[$baseChatPath].Replace('If(jsonDoc("priceCurrency")?.Value(Of String)(), ChatSettings.PriceCurrency)', 'ChatSettings.PriceCurrency') } },
    [pscustomobject]@{ Name = 'price-row-unseeded'; Expect = 'never seeded from the stored value';
        Apply = { param($s) $s[$baseChatPath] = $s[$baseChatPath].Replace("setNativeEntrySetting('settings-price-in'", "setNativeEntrySetting('settings-price-x'") } },
    [pscustomobject]@{ Name = 'usage-missing-on-early-return'; Expect = 'attached before the first early return';
        Apply = { param($s) $s[$kernelPath] = $s[$kernelPath].Replace('            {"usage", BuildUsageBlock()}' + "`r`n", '') } },
    [pscustomobject]@{ Name = 'footer-note-dropped'; Expect = 'no longer passes a reporter-built note';
        Apply = { param($s) $s[$streamPath] = $s[$streamPath].Replace('{UtilsService.ToJsStringLiteral(usageNote)}', '') } },
    [pscustomobject]@{ Name = 'footer-note-unescaped'; Expect = 'reaches the script text unescaped';
        Apply = { param($s) $s[$streamPath] = $s[$streamPath].Replace('ToJsStringLiteral(usageNote)', 'SerializeForScript2(usageNote)') } },
    [pscustomobject]@{ Name = 'budget-starts-blocking'; Expect = 'the budget may only warn';
        Apply = { param($s) $s[$reporterPath] = $s[$reporterPath].Replace('Public Shared Function BudgetText', "Public Shared Function BudgetText`r`n        If True Then Throw New Exception('blocked')`r`n        End If") } },
    [pscustomobject]@{ Name = 'card-injects-markup'; Expect = 'written with textContent, never as markup';
        Apply = { param($s) $s[$codeJsPath] = $s[$codeJsPath].Replace('noteSpan.textContent = ', 'noteSpan.innerHTML = ') } }
)

$baseline = @(Test-UsageContract $sourceByPath)
if ($baseline.Count -ne 0) {
    throw "usage contract is red on the clean tree: " + ($baseline -join ' | ')
}

$caught = 0
foreach ($probe in $probes) {
    $caught += Invoke-Probe $probe.Name $probe.Expect $probe.Apply
}

# ---------- runtime: the real parser, ledger, formatter and settings round trip ----------
Push-Location $bin
$originalSettingsBytes = Read-BytesIfExists $settingsPathOnDisk
try {
    Add-Type -Path $dll

    $tokenUsage = [ShareRibbon.TokenUsage]
    $reporter = [ShareRibbon.UsageCostReporter]

    # --- parser: both provider shapes, and nothing invented when usage is absent ---
    $openAi = $tokenUsage::FromResponse('{"choices":[{"message":{"content":"x"}}],"usage":{"prompt_tokens":120,"completion_tokens":30,"total_tokens":150}}')
    if ($null -eq $openAi) { throw "an OpenAI usage block must parse." }
    if ($openAi.PromptTokens -ne 120 -or $openAi.CompletionTokens -ne 30 -or $openAi.TotalTokens -ne 150) {
        throw "OpenAI usage came back $($openAi.PromptTokens)/$($openAi.CompletionTokens)/$($openAi.TotalTokens)."
    }

    # reasoning models bill a total larger than in+out; the provider number must win.
    $withReasoning = $tokenUsage::FromResponse('{"usage":{"prompt_tokens":10,"completion_tokens":5,"total_tokens":40}}')
    if ($withReasoning.TotalTokens -ne 40) { throw "the provider total must win over in+out, got $($withReasoning.TotalTokens)." }

    $anthropic = $tokenUsage::FromResponse('{"content":[{"type":"text","text":"x"}],"usage":{"input_tokens":77,"output_tokens":23}}')
    if ($null -eq $anthropic) { throw "an Anthropic usage block must parse - this is the shape the streaming translator forwards." }
    if ($anthropic.PromptTokens -ne 77 -or $anthropic.CompletionTokens -ne 23 -or $anthropic.TotalTokens -ne 100) {
        throw "Anthropic aliases came back $($anthropic.PromptTokens)/$($anthropic.CompletionTokens)/$($anthropic.TotalTokens)."
    }

    if ($null -ne $tokenUsage::FromResponse('{"choices":[]}')) { throw "a response without usage must yield Nothing, not a zero charge." }
    if ($null -ne $tokenUsage::FromResponse('not json at all')) { throw "malformed response text must yield Nothing." }
    if ($null -ne $tokenUsage::FromResponse('{"usage":{"currency":"USD"}}')) { throw "a usage object without any count must yield Nothing." }

    $negative = $tokenUsage::FromResponse('{"usage":{"prompt_tokens":-7,"completion_tokens":4,"total_tokens":4}}')
    if ($null -eq $negative) { throw "a partially negative usage should still record the trustworthy side." }
    if ($negative.PromptTokens -ne 0 -or $negative.CompletionTokens -ne 4) {
        throw "negative prompt_tokens must clamp to 0, got $($negative.PromptTokens)/$($negative.CompletionTokens)."
    }

    $base = New-Object ShareRibbon.TokenUsage
    [void]$base.Add($openAi)
    [void]$base.Add($anthropic)
    if ($base.TotalTokens -ne 250 -or $base.PromptTokens -ne 197) {
        throw "Add must accumulate both responses, got $($base.PromptTokens)/$($base.TotalTokens)."
    }
    $snapshot = $base.Copy()
    [void]$base.Add($openAi)
    if ($snapshot.TotalTokens -ne 250) { throw "Copy handed out the live ledger - a snapshot must not keep counting." }

    Write-Host "usage parser: openai / anthropic / no-usage / negative / accumulation checked"

    # --- ledger: one count, per-run window, session span ---
    $state = New-Object ShareRibbon.ChatStateService
    $state.RecordUsage($null)
    if ($state.SessionUsage.TotalTokens -ne 0) { throw "recording Nothing must not change the ledger." }
    $state.RecordUsage((New-Object ShareRibbon.TokenUsage))
    if ($state.SessionUsage.TotalTokens -ne 0) { throw "a zero-usage response must not pad the ledger." }

    $state.BeginRunUsage()
    [void]$state.RecordUsage($openAi)
    if ($state.RunUsage.TotalTokens -ne 150) { throw "the run window must see this run's call, got $($state.RunUsage.TotalTokens)." }
    [void]$state.RecordUsage($anthropic)
    if ($state.RunUsage.TotalTokens -ne 250 -or $state.CurrentSessionTotalTokens -ne 250) {
        throw "session and run must both reach 250, got $($state.RunUsage.TotalTokens)/$($state.CurrentSessionTotalTokens)."
    }
    if ($state.SessionUsage.TotalTokens -ne $state.CurrentSessionTotalTokens) {
        throw "the session read-out and the ledger are two different numbers again."
    }
    $state.BeginRunUsage()
    if ($state.RunUsage.TotalTokens -ne 0) { throw "the previous run leaked into the next run's bill." }
    if ($state.CurrentSessionTotalTokens -ne 250) { throw "starting a run must not reset the session ledger." }
    [void]$state.RecordUsage($withReasoning)
    $state.ResetSessionTokens()
    if ($state.CurrentSessionTotalTokens -ne 0 -or $state.RunUsage.TotalTokens -ne 0) {
        throw "resetting the session must clear both the ledger and the run window."
    }

    Write-Host "ledger: single writer, run window, session span, reset checked"

    # --- formatter: unpriced shows no amount, priced shows exact arithmetic ---
    $originalIn = [ShareRibbon.ChatSettings]::PricePerMillionIn
    $originalOut = [ShareRibbon.ChatSettings]::PricePerMillionOut
    $originalCurrency = [ShareRibbon.ChatSettings]::PriceCurrency
    $originalBudget = [ShareRibbon.ChatSettings]::SessionTokenBudget

    [ShareRibbon.ChatSettings]::PricePerMillionIn = -1
    [ShareRibbon.ChatSettings]::PricePerMillionOut = -1
    [ShareRibbon.ChatSettings]::PriceCurrency = ''
    [ShareRibbon.ChatSettings]::SessionTokenBudget = 0

    $costPhrase = -join ([char]0x4F30, [char]0x7B97, [char]0x8D39, [char]0x7528)
    if ($reporter::IsPriced()) { throw "prices of -1 must not count as configured." }
    if ($reporter::EstimateCost(1000, 1000) -ne -1) { throw "an unpriced estimate must be -1, never 0 - 0 reads as free." }
    if ($reporter::CostText(1000, 1000) -ne '') { throw "no price means no cost sentence." }
    $unpriced = $reporter::SummaryText($anthropic, 250)
    if ($unpriced -match [regex]::Escape($costPhrase)) { throw "the card reports money while nothing is configured." }
    $tokensWord = '100 tokens'
    if ($unpriced -notmatch [regex]::Escape($tokensWord)) { throw "the unpriced card must still report the token count, got '$unpriced'." }
    if ($unpriced -notmatch [regex]::Escape($reporter::UnpricedText)) { throw "the unpriced card must say why there is no amount." }

    # half a price is not a price: with only one side set, an estimate would be quietly wrong.
    [ShareRibbon.ChatSettings]::PricePerMillionIn = 2
    [ShareRibbon.ChatSettings]::PricePerMillionOut = -1
    if ($reporter::IsPriced()) { throw "a one-sided price must not be treated as configured." }
    [ShareRibbon.ChatSettings]::PricePerMillionOut = 6
    [ShareRibbon.ChatSettings]::PricePerMillionIn = -1
    if ($reporter::IsPriced()) { throw "a one-sided price must not be treated as configured (other side)." }

    [ShareRibbon.ChatSettings]::PricePerMillionIn = 2
    [ShareRibbon.ChatSettings]::PriceCurrency = '$'
    # 1M in at 2 plus 1M out at 6 per million is exactly 8 - this is the arithmetic a user reads as a bill,
    # so the whole sentence is pinned instead of "it contains a dollar sign somewhere".
    if ($reporter::EstimateCost(1000000, 1000000) -ne 8) { throw "2 and 6 per million over 1M/1M must estimate 8, got $($reporter::EstimateCost(1000000, 1000000))." }
    $priced = $reporter::CostText(1000000, 1000000)
    $pricedExpected = $costPhrase + ' $8'
    if ($priced -ne $pricedExpected) { throw "the priced sentence must be '$pricedExpected', got '$priced'." }
    if ($reporter::CostText(500000, 250000) -ne ($costPhrase + ' $2.5')) { throw "a fractional estimate lost its decimals, got '$($reporter::CostText(500000, 250000))'." }

    $bigRun = $tokenUsage::FromResponse('{"usage":{"prompt_tokens":1000000,"completion_tokens":1000000,"total_tokens":2000000}}')
    # the card separator is a middle dot, built here so this file stays ASCII
    $sep = ' ' + [string][char]0xB7 + ' '
    if ($reporter::SummaryText($bigRun, 250) -ne ('2000000 tokens' + $sep + $pricedExpected)) {
        throw "the priced card must be the token count plus the amount, got '$($reporter::SummaryText($bigRun, 250))'."
    }
    if ($reporter::SummaryText($null, 250) -ne '') { throw "a run with no usage must render no money, not a guess." }

    Write-Host "formatter: unpriced and priced sentences checked end to end"

    # --- budget ---
    [ShareRibbon.ChatSettings]::SessionTokenBudget = 0
    if ($reporter::BudgetText(999999) -ne '') { throw "budget 0 must mean unlimited, not a warning on everything." }
    [ShareRibbon.ChatSettings]::SessionTokenBudget = 100
    if ($reporter::BudgetText(99) -ne '') { throw "under the budget must stay silent." }
    if ($reporter::BudgetText(100) -eq '') { throw "reaching the budget must warn." }
    if ($reporter::BudgetText(500) -notmatch '500') { throw "the warning must quote the number the user is over." }
    # the notice is additive: a priced run over budget shows the amount and the warning, never one instead of the other.
    $overBudget = $reporter::SummaryText($anthropic, 250)
    if ($overBudget -notmatch [regex]::Escape($costPhrase)) { throw "a priced card dropped the amount once the budget notice appeared." }
    if ($overBudget -notmatch [regex]::Escape($reporter::BudgetText(250))) { throw "the card must carry the budget wording the reporter owns, got '$overBudget'." }

    # --- settings round trip against the real file, restored afterwards ---
    # a key outside the SaveSettings table must survive, because this file used to be rebuilt from scratch
    $foreignKey = 'smokeForeignSetting'
    $seeded = '{"' + $foreignKey + '": "keepme"}'
    if (Test-Path -LiteralPath $settingsPathOnDisk) {
        $existing = Get-Content -LiteralPath $settingsPathOnDisk -Raw -Encoding UTF8
        $at = $existing.LastIndexOf('}')
        if ($at -lt 0) { throw "the chat settings file on disk is not a JSON object, so this gate cannot test the round trip." }
        $seeded = $existing.Insert($at, ',"' + $foreignKey + '": "keepme"')
    }
    [System.IO.File]::WriteAllText($settingsPathOnDisk, $seeded, [System.Text.Encoding]::UTF8)
    $appInfo = New-Object ShareRibbon.ApplicationInfo ('Word', [ShareRibbon.OfficeApplicationType]::Word)
    $writer = New-Object ShareRibbon.ChatSettings $appInfo
    $writer.SaveSettings(0.8, 5, $true, $true, $true, $false, 800, 'Ctrl+.', $false, $false, $false, 'Ctrl+Alt+K', 2.5, 7.5, '$', 54321)
    # SaveSettings also writes the statics, so clear them first - otherwise a load that read nothing still passes.
    [ShareRibbon.ChatSettings]::PricePerMillionIn = -1
    [ShareRibbon.ChatSettings]::PricePerMillionOut = -1
    [ShareRibbon.ChatSettings]::PriceCurrency = ''
    [ShareRibbon.ChatSettings]::SessionTokenBudget = 0
    $reader = New-Object ShareRibbon.ChatSettings $appInfo
    if ([ShareRibbon.ChatSettings]::PricePerMillionIn -ne 2.5 -or [ShareRibbon.ChatSettings]::PricePerMillionOut -ne 7.5) {
        throw "the prices did not survive a save/load round trip: $([ShareRibbon.ChatSettings]::PricePerMillionIn) / $([ShareRibbon.ChatSettings]::PricePerMillionOut)"
    }
    if ([ShareRibbon.ChatSettings]::PriceCurrency -ne '$') { throw "the currency symbol did not round trip." }
    if ([ShareRibbon.ChatSettings]::SessionTokenBudget -ne 54321) { throw "the budget did not round trip." }

    $stored = Get-Content -LiteralPath $settingsPathOnDisk -Raw -Encoding UTF8
    if ($stored -notmatch ('"' + $foreignKey + '"')) {
        throw "saving the new keys wiped a setting outside this table - SaveSettings must read-modify-write the same JObject."
    }

    # a cleared row means "stop estimating", so it must land as the unconfigured sentinel
    $writer.SaveSettings(0.8, 5, $true, $true, $true, $false, 800, 'Ctrl+.', $false, $false, $false, 'Ctrl+Alt+K', -1, -1, '', 0)
    [ShareRibbon.ChatSettings]::PricePerMillionIn = 2
    [ShareRibbon.ChatSettings]::SessionTokenBudget = 100
    [void](New-Object ShareRibbon.ChatSettings $appInfo)
    if ([ShareRibbon.ChatSettings]::PricePerMillionIn -ne -1 -or [ShareRibbon.ChatSettings]::SessionTokenBudget -ne 0) {
        throw "clearing the rows must return the sentinels, got $([ShareRibbon.ChatSettings]::PricePerMillionIn) / $([ShareRibbon.ChatSettings]::SessionTokenBudget)"
    }

    # garbage in the text box must not take the rest of LoadSettings down with it
    $writer.SaveSettings(0.8, 5, $true, $true, $true, $false, 800, 'Ctrl+.', $false, $false, $false, 'Ctrl+Alt+K', 3, 9, 'x', 100)
    $text = Get-Content -LiteralPath $settingsPathOnDisk -Raw -Encoding UTF8
    $broken = $text -replace '"pricePerMillionIn":\s*-?[\d.]+', '"pricePerMillionIn": "two dollars"'
    $broken = $broken -replace '"sessionTokenBudget":\s*-?[\d.]+', '"sessionTokenBudget": 777'
    if ($broken -eq $text) { throw "the malformed-price probe did not reach the file, so it proves nothing." }
    [System.IO.File]::WriteAllText($settingsPathOnDisk, $broken, [System.Text.Encoding]::UTF8)
    [ShareRibbon.ChatSettings]::PricePerMillionIn = -1
    [ShareRibbon.ChatSettings]::PricePerMillionOut = -1
    [ShareRibbon.ChatSettings]::PriceCurrency = ''
    [ShareRibbon.ChatSettings]::SessionTokenBudget = 0
    [void](New-Object ShareRibbon.ChatSettings $appInfo)
    if ([ShareRibbon.ChatSettings]::PricePerMillionIn -ne -1) {
        throw "a non-numeric stored price must land as unconfigured, never as a number made up from the text, got $([ShareRibbon.ChatSettings]::PricePerMillionIn)"
    }
    # the keys after the broken one must still load, or the garbage silently killed the rest of the settings
    if ([ShareRibbon.ChatSettings]::PricePerMillionOut -ne 9 -or [ShareRibbon.ChatSettings]::SessionTokenBudget -ne 777) {
        throw "a bad price stopped LoadSettings: out=$([ShareRibbon.ChatSettings]::PricePerMillionOut) budget=$([ShareRibbon.ChatSettings]::SessionTokenBudget)"
    }

    Write-Host "settings: round trip, clear-to-unconfigured, malformed value, untouched legacy keys checked"
}
finally {
    [ShareRibbon.ChatSettings]::PricePerMillionIn = $originalIn
    [ShareRibbon.ChatSettings]::PricePerMillionOut = $originalOut
    [ShareRibbon.ChatSettings]::PriceCurrency = $originalCurrency
    [ShareRibbon.ChatSettings]::SessionTokenBudget = $originalBudget
    Restore-File $settingsPathOnDisk $originalSettingsBytes
    Pop-Location
}

$restored = Read-BytesIfExists $settingsPathOnDisk
if (-not (Test-SameBytes $originalSettingsBytes $restored)) {
    throw "office_ai_chat_settings.json did not come back byte for byte - the gate wrote the user's config."
}

Write-Host "OK: usage has one parser, one ledger and one wording owner; money appears only when a price was set, and the budget warns without blocking."
[pscustomobject]@{
    ProbesRed     = $caught
    ContractRules = 10
    KeyRows       = $keyInventory.Count
    OpenAiTotal   = $openAi.TotalTokens
    AnthropicTotal = $anthropic.TotalTokens
    PricedCost    = $priced
    UnpricedCard  = $unpriced
} | Format-List
