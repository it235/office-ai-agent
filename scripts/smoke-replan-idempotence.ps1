# Replan idempotence smoke (E1-06).
# After three consecutive failures the loop may adopt a brand-new plan and restart at step 0.
# A write that already landed in an earlier plan version must never be executed a second time,
# and it must not add a second change-log row either.
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

function New-SmokeCall {
    param([string]$ToolId, [string]$ParamsJson)
    $call = New-Object ShareRibbon.Agent.ToolCall
    $call.ToolId = $ToolId
    $call.Parameters = [Newtonsoft.Json.Linq.JObject]::Parse($ParamsJson)
    return $call
}

function New-SmokeChange {
    param([bool]$Success, [bool]$Changed, [bool]$RolledBack, [string]$Summary)
    $record = New-Object ShareRibbon.Agent.AgentChangeRecord
    $record.Success = $Success
    $record.Changed = $Changed
    $record.RolledBack = $RolledBack
    $record.Summary = $Summary
    return $record
}

# ---- A. ledger contract: what counts as an applied write ----
$emptyLedger = New-Object ShareRibbon.Agent.AgentWriteLedger
if (-not [string]::IsNullOrWhiteSpace($emptyLedger.BuildAppliedWritesText())) {
    throw "Smoke failed: an empty ledger must not add noise to the prompt."
}
if ($null -ne $emptyLedger.FindAppliedInPriorPlan((New-SmokeCall "DeleteText" '{"range":"W1"}'))) {
    throw "Smoke failed: an empty ledger must not block any step."
}

$ledger = New-Object ShareRibbon.Agent.AgentWriteLedger
$failedChange = New-SmokeChange $false $true $false "failed write"
$rolledBackChange = New-SmokeChange $true $true $true "rolled back"
$committedChange = New-SmokeChange $true $true $false "wrote W1"

if ($null -ne $ledger.Record((New-SmokeCall "DeleteText" '{"range":"W1"}'), $null, "read step")) {
    throw "Smoke failed: a step with no change record is not an applied write."
}
if ($null -ne $ledger.Record((New-SmokeCall "DeleteText" '{"range":"W1"}'), $failedChange, "failed step")) {
    throw "Smoke failed: a failed write must stay re-executable after replan."
}
if ($null -ne $ledger.Record((New-SmokeCall "DeleteText" '{"range":"W1"}'), $rolledBackChange, "rolled step")) {
    throw "Smoke failed: a rolled-back write is not present in the document, it must not be skipped."
}
if ($null -eq $ledger.Record((New-SmokeCall "DeleteText" '{"range":"W1","text":"x"}'), $committedChange, "wrote W1")) {
    throw "Smoke failed: a committed write must enter the ledger."
}

# Same plan version: two identical steps are the repeat the user asked for, not a replan artifact.
if ($null -ne $ledger.FindAppliedInPriorPlan((New-SmokeCall "DeleteText" '{"text":"x","range":"W1"}'))) {
    throw "Smoke failed: steps inside the current plan version must never be treated as duplicates."
}

$ledger.BeginNextPlan()
$keyOrderTolerant = $ledger.FindAppliedInPriorPlan((New-SmokeCall "DeleteText" '{"text":"x","range":"W1"}'))
if ($null -eq $keyOrderTolerant) {
    throw "Smoke failed: replanned step repeating an applied write must be detected despite JSON key order."
}
if ($null -ne $ledger.FindAppliedInPriorPlan((New-SmokeCall "DeleteText" '{"range":"W2"}'))) {
    throw "Smoke failed: a different payload is a different write and must still execute."
}
if ($null -ne $ledger.FindAppliedInPriorPlan((New-SmokeCall "FormatRange" '{"range":"W1"}'))) {
    throw "Smoke failed: the fingerprint must include the tool id."
}

$writesText = $ledger.BuildAppliedWritesText()
if ($writesText -notmatch "APPLIED_WRITES") {
    throw "Smoke failed: the planner must be told which writes already landed."
}

# ---- B. loop end to end: injected failures trigger replan, applied writes are not replayed ----
function New-ActionJson {
    param([string]$ToolId, [string]$ParamsJson)
    return (@{ thought = "smoke"; action = @{ tool = $ToolId; params = ($ParamsJson | ConvertFrom-Json) } } |
        ConvertTo-Json -Compress -Depth 6)
}

function New-StepJson {
    param([string]$Description, [string]$ToolId, [string]$ParamsJson)
    return @{
        description = $Description
        code        = (@{ command = $ToolId; params = ($ParamsJson | ConvertFrom-Json) } | ConvertTo-Json -Compress -Depth 6)
        language    = "json"
    }
}

$specJson = '{"goal":"smoke replan","complexity":"simple"}'
$reflectionJson = '{"analysis":"host failure","strategy":"replan","reason":"smoke"}'

# Plan v1: two committed writes, then three injected host failures so the loop reflects.
$planJsonFirst = (@{
    understanding = "smoke"
    summary       = "smoke"
    capabilityGap = ""
    steps         = @(
        (New-StepJson "smoke write one" "InsertText" '{"content":"SMOKE_W1","position":"end"}'),
        (New-StepJson "smoke write two" "InsertText" '{"content":"SMOKE_W2","position":"end"}'),
        (New-StepJson "smoke fail one" "InsertText" '{"content":"SMOKE_FAIL_1","position":"end"}'),
        (New-StepJson "smoke fail two" "InsertText" '{"content":"SMOKE_FAIL_2","position":"end"}'),
        (New-StepJson "smoke fail three" "InsertText" '{"content":"SMOKE_FAIL_3","position":"end"}')
    )
} | ConvertTo-Json -Compress -Depth 8)

# Plan v2 the model produces after the failure: it re-plans the work it has already done.
$planJsonReplan = (@{
    understanding = "smoke replan"
    summary       = "smoke replan"
    capabilityGap = ""
    steps         = @(
        (New-StepJson "smoke write one" "InsertText" '{"content":"SMOKE_W1","position":"end"}'),
        (New-StepJson "smoke write three" "InsertText" '{"content":"SMOKE_W3","position":"end"}')
    )
} | ConvertTo-Json -Compress -Depth 8)

# The react step follows the plan step order; index 5 is the duplicated write inside plan v2.
$reactActions = @(
    (New-ActionJson "InsertText" '{"content":"SMOKE_W1","position":"end"}'),
    (New-ActionJson "InsertText" '{"content":"SMOKE_W2","position":"end"}'),
    (New-ActionJson "InsertText" '{"content":"SMOKE_FAIL_1","position":"end"}'),
    (New-ActionJson "InsertText" '{"content":"SMOKE_FAIL_2","position":"end"}'),
    (New-ActionJson "InsertText" '{"content":"SMOKE_FAIL_3","position":"end"}'),
    (New-ActionJson "InsertText" '{"content":"SMOKE_W1","position":"end"}'),
    (New-ActionJson "InsertText" '{"content":"SMOKE_W3","position":"end"}')
)

$script:prompts = @()
$script:hostCommands = @()
$script:reactCursor = 0
$script:sawReflection = $false

# Drive LoopEngine directly: it is the only component that owns the replan decision, and it keeps
# the run inside one session state machine instead of adding a second harness for the test.
$registry = [ShareRibbon.Agent.ToolRegistry]::new($null)
if ($registry.LoadFromRuntimeDirectories($binDir) -le 0) {
    throw "Smoke failed: no Word tool definitions were loaded from $binDir."
}
$engine = New-Object ShareRibbon.Agent.LoopEngine($registry,
    (New-Object ShareRibbon.Agent.AgentMemory),
    (New-Object ShareRibbon.Agent.PromptManager((Join-Path $binDir "Prompts"))))

$engine.SendAIRequest = [System.Func[string, string, System.Collections.Generic.List[ShareRibbon.HistoryMessage], System.Threading.Tasks.Task[string]]] {
    param($prompt, $system, $history)
    # The applied-write constraint is delivered as a system message, so the audit must see both halves.
    $script:prompts += ($prompt + [Environment]::NewLine + $system)

    if ($prompt -match '"complexity"') {
        return [System.Threading.Tasks.Task]::FromResult($specJson)
    }
    if ($prompt -match '"toolId":\s*"') {
        # Repair declines with plain text, so the loop exhausts the step and moves on.
        return [System.Threading.Tasks.Task]::FromResult("smoke repair declined")
    }
    if ($prompt -match '"strategy"') {
        $script:sawReflection = $true
        return [System.Threading.Tasks.Task]::FromResult($reflectionJson)
    }
    if ($prompt -match '"capabilityGap"') {
        if ($script:sawReflection) {
            return [System.Threading.Tasks.Task]::FromResult($planJsonReplan)
        }
        return [System.Threading.Tasks.Task]::FromResult($planJsonFirst)
    }

    $index = [Math]::Min($script:reactCursor, $reactActions.Count - 1)
    $script:reactCursor += 1
    return [System.Threading.Tasks.Task]::FromResult($reactActions[$index])
}

$engine.OnRequestApproval = [System.Func[ShareRibbon.Agent.ApprovalRequest, System.Threading.Tasks.Task[ShareRibbon.Agent.ApprovalDecision]]] {
    param($request)
    return [System.Threading.Tasks.Task]::FromResult([ShareRibbon.Agent.ApprovalDecision]::Session)
}

$registry.ExecuteCodeWithToolResult = [System.Func[string, string, bool, ShareRibbon.Agent.ToolResult]] {
    param($code, $language, $preview)
    $script:hostCommands += $code

    if ($code -match "SMOKE_FAIL") {
        return [ShareRibbon.Agent.ToolResult]::Failed(
            "smoke-host", "injected host failure", $null, "COM_ERROR", "injected host failure", "injected host failure", $true, $null)
    }

    $observation = @{
        kind       = "write"
        summary    = "applied marker"
        changed    = $true
        targetRefs = @("Office:ActiveDocument")
    }
    return [ShareRibbon.Agent.ToolResult]::Succeed("smoke-host", "applied", $null, $observation, "undo-1")
}

$session = New-Object ShareRibbon.Agent.AgentSession("smoke replan idempotence", "Word", "smoke document content")
$result = $engine.RunAsync($session, "smoke system prompt", $null).GetAwaiter().GetResult()
if ($null -eq $result) {
    throw "Smoke failed: LoopEngine.RunAsync returned no result after $($script:hostCommands.Count) host calls."
}

$w1Calls = @($script:hostCommands | Where-Object { $_ -match "SMOKE_W1" }).Count
$w2Calls = @($script:hostCommands | Where-Object { $_ -match "SMOKE_W2" }).Count
$w3Calls = @($script:hostCommands | Where-Object { $_ -match "SMOKE_W3" }).Count
$failCalls = @($script:hostCommands | Where-Object { $_ -match "SMOKE_FAIL" }).Count
$toldPrompts = @($script:prompts | Where-Object { $_ -match "APPLIED_WRITES" }).Count
# The ledger skip is only proven by the marker the executor leaves behind; W1 reaching the host once
# would also be true if the replanned simply dropped it.
$skipMarks = @($session.Iterations | Where-Object { $_.Explanation.ObservationJson -match "duplicate_write_skipped" }).Count

[PSCustomObject]@{
    Success     = $result.Success
    Cancelled   = $result.IsCancelled
    Iterations  = $result.IterationsCompleted
    HostCalls   = $script:hostCommands.Count
    W1          = $w1Calls
    W2          = $w2Calls
    W3          = $w3Calls
    Failures    = $failCalls
    LedgerSkips = $skipMarks
    PlanVersion = $session.Writes.PlanGeneration
    ChangeRows  = $session.ChangeLog.Entries.Count
    WrittenRows = $session.ChangeLog.WrittenCount
    ToldPrompts = $toldPrompts
} | Format-List

if ($script:sawReflection -eq $false) {
    throw "Smoke failed: the injected failures never reached the reflection/replan path."
}
if (-not $result.Success) {
    throw "Smoke failed: the run must finish after replan, got: $($result.Message)"
}
if ($w1Calls -ne 1) {
    throw "Smoke failed: write one reached the host $w1Calls times, replan must not replay an applied write."
}
if ($w2Calls -ne 1 -or $w3Calls -ne 1) {
    throw "Smoke failed: the other writes must land exactly once (W2=$w2Calls, W3=$w3Calls)."
}
if ($failCalls -ne 3) {
    throw "Smoke failed: expected three injected host failures, got $failCalls."
}
if ($session.Writes.PlanGeneration -ne 2) {
    throw "Smoke failed: replan must advance the plan version, got $($session.Writes.PlanGeneration)."
}
if ($skipMarks -ne 1) {
    throw "Smoke failed: the re-planned applied write must be skipped by the ledger exactly once, got $skipMarks."
}
if ($session.ChangeLog.Entries.Count -ne 3 -or $session.ChangeLog.WrittenCount -ne 3) {
    throw "Smoke failed: the change log carries duplicate rows after replan " +
         "(entries=$($session.ChangeLog.Entries.Count), written=$($session.ChangeLog.WrittenCount))."
}
if ($toldPrompts -lt 2) {
    throw "Smoke failed: neither the reflection prompt nor the replanned plan prompt was told about applied writes."
}
