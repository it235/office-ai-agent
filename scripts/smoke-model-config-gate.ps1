param(
    [string]$Configuration = "Debug"
)

# E3-03: an unconfigured model must be blocked before the request reaches the chat or agent chain,
# and the panel must receive a machine-readable CTA payload instead of a status-bar string.
# This runs against the compiled DLL, so it also proves the gate and the validator cannot drift apart.
# Keep this file ASCII only: PowerShell 5.1 reads a BOM-less script as ANSI, so any non-ASCII literal
# here would compare against mojibake instead of the strings the VB layer really returns.

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$bin = Join-Path $repoRoot "ShareRibbon\bin\$Configuration"
$dll = Join-Path $bin "ShareRibbon.dll"

if (-not (Test-Path -LiteralPath $dll)) {
    throw "ShareRibbon.dll not found. Run: .\scripts\build-code-projects.ps1 -Configuration $Configuration -Projects `"ShareRibbon\ShareRibbon.vbproj`""
}

Push-Location $bin
try {
    Add-Type -Path $dll

    $gate = New-Object ShareRibbon.ModelConfigGate
    $validator = New-Object ShareRibbon.ChatSendValidator

    $keyEnum = [ShareRibbon.ChatSendValidationFailure]::MissingApiKey
    $urlEnum = [ShareRibbon.ChatSendValidationFailure]::MissingApiUrl
    $questionEnum = [ShareRibbon.ChatSendValidationFailure]::MissingQuestion

    # 1. a complete config must not produce a CTA on either entry point
    foreach ($action in @("chat", "agent")) {
        $clean = $gate.Evaluate("https://api.example.test/v1/chat/completions", "smoke-key", $action)
        if ($clean -ne $null) { throw "a configured model must not be blocked ($action)." }
    }

    # 2. missing key wins over everything else and keeps the reason machine readable
    $keyPrompt = $gate.Evaluate("https://api.example.test/v1/chat/completions", "", "chat")
    if ($keyPrompt -eq $null) { throw "an empty api key must block the chat send." }
    if ($keyPrompt.Reason -ne "missing_api_key") { throw "reason must be missing_api_key, got '$($keyPrompt.Reason)'." }
    if ([string]::IsNullOrWhiteSpace($keyPrompt.Title)) { throw "the CTA needs a heading the user can read." }
    if ([string]::IsNullOrWhiteSpace($keyPrompt.Message)) { throw "the CTA needs a message the user can act on." }
    if ([string]::IsNullOrWhiteSpace($keyPrompt.ActionLabel)) { throw "the CTA needs a button label." }
    if ($keyPrompt.BlockedAction -ne "chat") { throw "the CTA must carry the blocked entry point." }

    # 3. whitespace counts as unconfigured: a pasted space must not look like a saved key
    $blankPrompt = $gate.Evaluate("https://api.example.test/v1/chat/completions", "   ", "agent")
    if ($blankPrompt -eq $null) { throw "a whitespace-only api key must block the agent start." }
    if ($blankPrompt.BlockedAction -ne "agent") { throw "the agent CTA must report the blocked action." }
    if ($blankPrompt.Message -eq $keyPrompt.Message) { throw "chat and agent share one copy, so the user cannot tell a stopped task from an unsent message." }

    # 4. key present and url missing must be a distinct reason, or the user is sent to the wrong field
    $urlPrompt = $gate.Evaluate("", "smoke-key", "chat")
    if ($urlPrompt -eq $null) { throw "an empty api url must block the send." }
    if ($urlPrompt.Reason -ne "missing_api_url") { throw "reason must be missing_api_url, got '$($urlPrompt.Reason)'." }

    # 5. the config gate and the send validator must be the same gate, not two copies
    $badConfig = $validator.Validate("", "", "")
    if ($badConfig.Failure -ne $keyEnum) { throw "an empty config must report the key first, got $($badConfig.Failure)." }
    $goodConfigNoQuestion = $validator.Validate("https://api.example.test/v1/chat/completions", "smoke-key", "")
    if ($goodConfigNoQuestion.Failure -ne $questionEnum) { throw "a configured model with no question must report the missing question, got $($goodConfigNoQuestion.Failure)." }
    $urlOnlyFailure = $validator.Validate("", "smoke-key", "hello")
    if ($urlOnlyFailure.Failure -ne $urlEnum) { throw "Validate must reuse ValidateConfig, got $($urlOnlyFailure.Failure)." }
    if (-not $validator.Validate("https://api.example.test/v1/chat/completions", "smoke-key", "hello").IsValid) { throw "a complete request must pass validation." }

    # 6. the payload is JSON, so the panel never has to trust hand-escaped script text
    $showPrefix = "showModelConfigRequired("
    $script = [ShareRibbon.ModelConfigGate]::BuildShowScript($keyPrompt)
    if ($script -notmatch "^showModelConfigRequired\(") { throw "show script must call showModelConfigRequired, got '$script'." }
    if ($script -notmatch "\);\s*$") { throw "show script must be a complete statement." }
    $json = $script.Substring($showPrefix.Length, $script.Length - $showPrefix.Length - 2)
    $parsed = ConvertFrom-Json -InputObject $json
    if ($parsed.reason -ne "missing_api_key") { throw "the CTA payload must parse as JSON with the reason intact." }
    if ([string]::IsNullOrWhiteSpace($parsed.message)) { throw "the CTA payload lost its message." }

    # 7. quotes and newlines in the copy must survive the round trip instead of breaking the script
    $hostile = New-Object ShareRibbon.ModelConfigPrompt
    $hostile.Reason = "missing_api_key"
    $hostile.Title = 'say "hello" <b>'
    $hostile.Message = "line1`nline2 with a \ backslash and 'quote'"
    $hostile.ActionLabel = "Configure API"
    $hostile.BlockedAction = "chat"
    $roundTrip = ConvertFrom-Json -InputObject $hostile.ToJson()
    if ($roundTrip.Title -ne 'say "hello" <b>') { throw "the prompt payload is not escaped, title came back '$($roundTrip.Title)'." }
    if ($roundTrip.Message -notmatch "line2") { throw "a newline in the prompt message breaks the injected script." }
    if ($hostile.ToJson() -match "`n") { throw "a raw newline reached the script text, WebView2 would get a broken statement." }

    if ([ShareRibbon.ModelConfigGate]::BuildShowScript($null) -ne "") { throw "no prompt must mean no script." }
    if ([ShareRibbon.ModelConfigGate]::BuildHideScript() -notmatch "^hideModelConfigRequired\(\);\s*$") { throw "the hide script must match the panel renderer." }

    Write-Host "OK: model config gate blocks both entries and emits a parseable CTA."
    [pscustomobject]@{
        KeyReason       = $keyPrompt.Reason
        UrlReason       = $urlPrompt.Reason
        ChatDiffersFromAgent = ($keyPrompt.Message -ne $blankPrompt.Message)
        ScriptLength    = $script.Length
    } | Format-List
}
finally {
    Pop-Location
}
