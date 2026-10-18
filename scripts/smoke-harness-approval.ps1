# Approval-lifecycle smoke: an unresolved approval must never lock the session forever.
# Covers both exits - timeout gives up, and a new turn supersedes the stale request.
# Either way the pending write is rejected, never silently executed.
param(
    [string]$Configuration = "Debug"
)

$ErrorActionPreference = "Stop"
$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$binDir = Join-Path $repoRoot "ShareRibbon\bin\$Configuration"
$dll = Join-Path $binDir "ShareRibbon.dll"

if (-not (Test-Path -LiteralPath $dll)) {
    throw "ShareRibbon assembly not found: $dll. Build the code projects first."
}

$newtonsoft = Join-Path $binDir "Newtonsoft.Json.dll"
if (Test-Path -LiteralPath $newtonsoft) {
    Add-Type -LiteralPath $newtonsoft -ErrorAction SilentlyContinue
}
Add-Type -LiteralPath $dll

$envelope = @{ toolId = "DeleteText"; parameters = @{ range = "selection" } } | ConvertTo-Json -Compress
# One canned reply satisfies the spec, plan and react parsers, so a single stub drives the whole chain.
$cannedResponse = @{
    goal          = "delete the selection"
    complexity    = "simple"
    understanding = "smoke"
    summary       = "smoke"
    steps         = @(@{ description = "delete the selection"; code = $envelope; language = "json" })
    action        = @{ tool = "DeleteText"; params = @{ range = "selection" } }
} | ConvertTo-Json -Compress -Depth 6

# Any host call recorded here means an unapproved write actually executed.
$script:hostCalls = @()

function New-SmokeKernel {
    $kernel = New-Object ShareRibbon.Agent.AgentKernel
    $kernel.SendAIRequest = [System.Func[string, string, System.Collections.Generic.List[ShareRibbon.HistoryMessage], System.Threading.Tasks.Task[string]]] {
        param($prompt, $system, $history)
        return [System.Threading.Tasks.Task]::FromResult($cannedResponse)
    }
    $kernel.ExecuteCodeWithToolResult = [System.Func[string, string, bool, ShareRibbon.Agent.ToolResult]] {
        param($code, $language, $preview)
        $script:hostCalls += $code
        return [ShareRibbon.Agent.ToolResult]::Succeed("smoke-host", "executed")
    }
    $kernel.Initialize()
    return $kernel
}

function New-SmokeTurn($sessionId) {
    $turn = New-Object ShareRibbon.Agent.Harness.UserTurn
    $turn.SessionId = $sessionId
    $turn.AppType = "Word"
    $turn.Text = "delete the selection"
    $turn.Mode = "agent"
    # Skill recall ranking decides which Word boundary applies, and that ranking depends on the skill
    # index and the resolved intent - it is not what this smoke tests. Without pinning, the same canned
    # DeleteText plan passes under the baseline skill and dies with TOOL_NOT_ALLOWED under a narrower
    # typesetting skill, so the approval lifecycle could not be trusted. Pin the baseline instead.
    $turn.SelectedSkills.Add((Get-BaselineWordSkill))
    return $turn
}

function Get-BaselineWordSkill {
    $svc = [ShareRibbon.Agent.AgentKernel].Assembly.GetTypes() |
        Where-Object { $_.Name -eq 'SkillsDirectoryService' } | Select-Object -First 1
    if (-not $svc) { throw "Smoke setup failed: SkillsDirectoryService not found in ShareRibbon.dll." }
    $slots = [object[]]::new(1)
    $slots[0] = $false
    $skills = $svc.GetMethod('GetAllSkills').Invoke($null, $slots)
    $baseline = @($skills) | Where-Object {
        $_.Application -eq 'Word' -and @($_.AllowedTools) -contains 'DeleteText'
    } | Select-Object -First 1
    if (-not $baseline) { throw "Smoke setup failed: no Word skill allows DeleteText, so the approval path cannot be exercised." }
    return $baseline
}

function Invoke-SmokeRun($harness, $turn) {
    $cts = New-Object System.Threading.CancellationTokenSource
    return $harness.RunAsync($turn, $cts.Token).GetAwaiter().GetResult()
}

function New-SmokeHarness($kernel) {
    return New-Object ShareRibbon.Agent.Harness.OfficeHarness($kernel, (New-Object ShareRibbon.Agent.Harness.NoopRunTraceStore))
}

# Case 1: the approval times out, and the session accepts the next turn again.
$timeoutHarness = New-SmokeHarness (New-SmokeKernel)
$timeoutHarness.ApprovalTimeout = [TimeSpan]::FromMilliseconds(400)

$first = Invoke-SmokeRun $timeoutHarness (New-SmokeTurn "smoke-approval-timeout-1")
if ($first.Status.ToString() -ne "AwaitingApproval") {
    throw "Smoke failed: expected AwaitingApproval, got $($first.Status) / $($first.DebugMessage)"
}
Start-Sleep -Milliseconds 2500

$afterTimeout = Invoke-SmokeRun $timeoutHarness (New-SmokeTurn "smoke-approval-timeout-2")
if ($afterTimeout.DebugMessage -eq "RUN_ALREADY_ACTIVE") {
    throw "Smoke failed: the timed-out approval still blocks the session."
}

# Case 2: the user ignores the buttons and sends a new turn; the stale approval gives way.
$supersedeHarness = New-SmokeHarness (New-SmokeKernel)
$supersedeHarness.ApprovalTimeout = [TimeSpan]::FromMinutes(10)

$pending = Invoke-SmokeRun $supersedeHarness (New-SmokeTurn "smoke-supersede-1")
if ($pending.Status.ToString() -ne "AwaitingApproval") {
    throw "Smoke failed: expected AwaitingApproval, got $($pending.Status) / $($pending.DebugMessage)"
}
$superseded = Invoke-SmokeRun $supersedeHarness (New-SmokeTurn "smoke-supersede-2")
if ($superseded.DebugMessage -eq "RUN_ALREADY_ACTIVE") {
    throw "Smoke failed: a new turn could not supersede the stale approval request."
}

$result = [pscustomobject]@{
    TimeoutConverged    = ($afterTimeout.Status.ToString() -eq "AwaitingApproval")
    SupersededConverged = ($superseded.Status.ToString() -eq "AwaitingApproval")
    HostCalls           = $script:hostCalls.Count
}
$result | Format-List

if ($result.TimeoutConverged -ne $true) {
    throw "Smoke failed: the run after an approval timeout did not reach a fresh approval prompt: $($afterTimeout.Status) / $($afterTimeout.UserMessage)"
}
if ($result.SupersededConverged -ne $true) {
    throw "Smoke failed: superseding an awaiting approval must start the new turn, not abandon it: $($superseded.Status) / $($superseded.UserMessage)"
}
if ($script:hostCalls.Count -gt 0) {
    throw "Smoke failed: an operation still awaiting approval reached the host executor."
}
