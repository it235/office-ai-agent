# Approval-lifecycle smoke: an unresolved approval must never lock the session forever.
# Covers both exits - timeout gives up, and a new turn supersedes the stale request.
# Either way the pending write is rejected, never silently executed. It also covers the three-state
# decision: allow once reaches the host, session approval is not consumed, abort stops the round.
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

# Case 3: an explicit "allow once" decision resumes the pending write and reaches the host.
$approveHarness = New-SmokeHarness (New-SmokeKernel)
$approveHarness.ApprovalTimeout = [TimeSpan]::FromMinutes(10)
$awaiting = Invoke-SmokeRun $approveHarness (New-SmokeTurn "smoke-approve-once-1")
if ($awaiting.Status.ToString() -ne "AwaitingApproval") {
    throw "Smoke failed: expected AwaitingApproval for allow-once, got $($awaiting.Status) / $($awaiting.DebugMessage)"
}
$callsBeforeDecision = $script:hostCalls.Count
$policy = [ShareRibbon.Agent.ApprovalDecision]
$decided = $approveHarness.DecideAsync($awaiting.RunId, $policy::Once, [System.Threading.CancellationToken]::None).GetAwaiter().GetResult()
if ($script:hostCalls.Count -le $callsBeforeDecision) {
    throw "Smoke failed: allow-once did not reach the host executor."
}

# The session decision must survive repeated checks and must never be consumed like a one-time token.
$context = New-Object ShareRibbon.Agent.ToolExecutionContext
$context.ApproveToolForSession("DeleteText")
for ($i = 0; $i -lt 3; $i++) {
    if (-not $context.IsToolApprovedForSession("deletetext")) {
        throw "Smoke failed: session approval must match case-insensitively and stay in effect."
    }
}
if ($context.ConsumeToolApproval("DeleteText")) {
    throw "Smoke failed: session approval was consumed by the one-time approval path."
}
$decideMethod = [ShareRibbon.Agent.Harness.IOfficeHarness].GetMethod("DecideAsync")
if ($null -eq $decideMethod) {
    throw "Smoke failed: IOfficeHarness no longer exposes the three-state approval decision."
}

# Case 4: Abort closes the whole round. Unlike Deny (skip this step), it must never reach the host and
# must not be reported as success, otherwise a walked-away user still gets writes applied.
$abortHarness = New-SmokeHarness (New-SmokeKernel)
$abortHarness.ApprovalTimeout = [TimeSpan]::FromMinutes(10)
$awaitingAbort = Invoke-SmokeRun $abortHarness (New-SmokeTurn "smoke-abort-1")
if ($awaitingAbort.Status.ToString() -ne "AwaitingApproval") {
    throw "Smoke failed: expected AwaitingApproval for abort, got $($awaitingAbort.Status) / $($awaitingAbort.DebugMessage)"
}
$callsBeforeAbort = $script:hostCalls.Count
$aborted = $abortHarness.DecideAsync($awaitingAbort.RunId, $policy::Abort, [System.Threading.CancellationToken]::None).GetAwaiter().GetResult()
if ($script:hostCalls.Count -ne $callsBeforeAbort) {
    throw "Smoke failed: an aborted approval still reached the host executor."
}
if ($aborted.Status.ToString() -eq "Succeeded") {
    throw "Smoke failed: an aborted round was reported as success."
}
$afterAbort = Invoke-SmokeRun $abortHarness (New-SmokeTurn "smoke-abort-2")
if ($afterAbort.DebugMessage -eq "RUN_ALREADY_ACTIVE") {
    throw "Smoke failed: an aborted round never released the session."
}

$result = [pscustomobject]@{
    TimeoutConverged    = ($afterTimeout.Status.ToString() -eq "AwaitingApproval")
    SupersededConverged = ($superseded.Status.ToString() -eq "AwaitingApproval")
    AllowOnceReachedHost = ($script:hostCalls.Count -gt $callsBeforeDecision)
    SessionDecisionKept  = ($context.IsToolApprovedForSession("DeleteText"))
    AbortSkippedHost     = ($script:hostCalls.Count -eq $callsBeforeAbort)
    AbortReleasedSession = ($afterAbort.DebugMessage -ne "RUN_ALREADY_ACTIVE")
    HostCalls           = $script:hostCalls.Count
}
$result | Format-List

if ($result.TimeoutConverged -ne $true) {
    throw "Smoke failed: the run after an approval timeout did not reach a fresh approval prompt: $($afterTimeout.Status) / $($afterTimeout.UserMessage)"
}
if ($result.SupersededConverged -ne $true) {
    throw "Smoke failed: superseding an awaiting approval must start the new turn, not abandon it: $($superseded.Status) / $($superseded.UserMessage)"
}
if ($callsBeforeDecision -gt 0) {
    throw "Smoke failed: an operation still awaiting approval reached the host executor."
}
if ($result.AbortSkippedHost -ne $true -or $result.AbortReleasedSession -ne $true) {
    throw "Smoke failed: an aborted round must neither write nor keep the session locked."
}
