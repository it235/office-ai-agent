# Step intervention smoke (E10-01 / E10-02).
# The plan card lets the user retry or skip a failed step, and rewrite a step's description and
# parameters. What must stay true: only the clicked step re-runs, no second terminal or plan owner
# appears, a click that arrives while the loop is mid-step is still honoured at the next boundary,
# and a rewritten step runs from its saved text without asking the model to rethink it.
# This drives the real LoopEngine with a scripted model and a scripted host executor, so the judgement
# is the same code path a Word/Excel/PPT host would take - no test-only loop was added.
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

function Assert-That {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "Smoke failed: $Message" }
}

$script:caseCount = 0

# ---- A. the inbox itself: one decision per step, nothing guessed ----
$inbox = New-Object ShareRibbon.Agent.AgentStepInterventions
$inbox.Queue(-1, [ShareRibbon.Agent.StepIntervention]::Retry)
$inbox.Queue(3, [ShareRibbon.Agent.StepIntervention]::None)
Assert-That ($inbox.PendingCount() -eq 0) "a negative step or an empty decision must never enter the inbox"

$inbox.Queue(2, [ShareRibbon.Agent.StepIntervention]::Retry)
$inbox.Queue(2, [ShareRibbon.Agent.StepIntervention]::Retry)
Assert-That ($inbox.PendingCount() -eq 1) "clicking the same step twice is one decision, got $($inbox.PendingCount())"

$inbox.Queue(0, [ShareRibbon.Agent.StepIntervention]::Skip)
Assert-That ($inbox.PendingIndices()[0] -eq 0) "pending steps come back in step order, the loop must rewind to the earliest one"
$inbox.Queue(2, [ShareRibbon.Agent.StepIntervention]::Skip)
$taken = $inbox.Take(2)
Assert-That ($null -ne $taken -and $taken.Action -eq [ShareRibbon.Agent.StepIntervention]::Skip) `
    "the later decision on a step must replace the earlier one, got $(if ($null -eq $taken) { 'nothing' } else { $taken.Action })"
Assert-That ($null -eq $inbox.Take(2)) "one decision may act on one failure only"

# A rewrite must carry its own text: taking the verb without the payload would apply a definition
# nobody wrote. Both fields travel with the one decision and are taken once.
$editedInboxCode = '{"command":"InsertText","params":{"content":"EDITED"}}'
$inbox.QueueEdit(1, "edited description", $editedInboxCode)
$editRequest = $inbox.Take(1)
Assert-That ($null -ne $editRequest -and $editRequest.Action -eq [ShareRibbon.Agent.StepIntervention]::Edit) `
    "a queued rewrite must come back as an edit decision"
Assert-That ($editRequest.Description -eq "edited description" -and $editRequest.Code -eq $editedInboxCode) `
    "the rewrite lost its text on the way out of the inbox"
Assert-That ($null -eq $inbox.Take(1)) "a rewrite may act on one step once"
# The verb-only door must not swallow an "edit" click that carries no new text.
$inbox.Queue(4, [ShareRibbon.Agent.StepIntervention]::Edit)
Assert-That ($null -eq $inbox.Take(4)) "an edit without text must never enter the verb-only door"

$inbox.Clear()
Assert-That ($inbox.PendingCount() -eq 0) "a replaced plan must not keep any decision addressed to the old step numbers"

# The verb list is closed: an unknown verb must not be read as "the user probably wanted a retry".
$verbs = @{
    retry    = [ShareRibbon.Agent.StepIntervention]::Retry
    SKIP     = [ShareRibbon.Agent.StepIntervention]::Skip
    abort    = [ShareRibbon.Agent.StepIntervention]::None
    delete   = [ShareRibbon.Agent.StepIntervention]::None
    ""       = [ShareRibbon.Agent.StepIntervention]::None
}
foreach ($verb in $verbs.Keys) {
    $parsed = [ShareRibbon.Agent.AgentStepInterventions]::Parse($verb)
    Assert-That ($parsed -eq $verbs[$verb]) "Parse($verb) returned $parsed, the panel and the loop would disagree about what the click meant"
}
$script:caseCount += 1
Write-Host "    ok the inbox holds one closed-vocabulary decision per step"

# ---- B. the single owner refuses what it cannot place ----
# AgentKernel.QueueStepAction is the only door from a click to the inbox. _session is private and is
# only ever assigned by a run, so the guard matrix is aimed at it through the field that owns it.
$sessionField = [ShareRibbon.Agent.AgentKernel].GetField(
    "_session", [System.Reflection.BindingFlags]::NonPublic -bor [System.Reflection.BindingFlags]::Instance)
Assert-That ($null -ne $sessionField) `
    "AgentKernel no longer holds the run in a _session field - the guards below would be judging nothing"

$kernel = New-Object ShareRibbon.Agent.AgentKernel
$live = New-Object ShareRibbon.Agent.AgentSession("smoke step intervention", "Word", "smoke document")
$live.Status = [ShareRibbon.Agent.AgentStatus]::Executing
$sessionField.SetValue($kernel, $live)

Assert-That (-not $kernel.QueueStepAction($null, 1, "retry")) "a click with no session id must be refused"
Assert-That (-not $kernel.QueueStepAction("other-run", 1, "retry")) `
    "a click aimed at another run must be refused, its step number means nothing here"
Assert-That (-not $kernel.QueueStepAction($live.Id, 1, "delete")) "an unknown verb must be refused, not guessed"
Assert-That ($live.Interventions.PendingCount() -eq 0) "a refused click must leave nothing queued"
Assert-That ($kernel.QueueStepAction($live.Id, 1, "retry")) "a live session must accept a decision for a step of its own plan"
Assert-That ($live.Interventions.PendingCount() -eq 1) "the accepted click must be queued for the loop to take at the boundary"

# A finished run has no step left to recover: the card must hear "this step is no longer in this run".
foreach ($terminal in @([ShareRibbon.Agent.AgentStatus]::Completed,
                        [ShareRibbon.Agent.AgentStatus]::Failed,
                        [ShareRibbon.Agent.AgentStatus]::Aborted)) {
    $live.Status = $terminal
    Assert-That (-not $kernel.QueueStepAction($live.Id, 1, "skip")) `
        "a terminal session ($terminal) must refuse a step decision instead of queueing it into a dead run"
}
# The rewrite door (E10-02) has more to judge than a verb: it must refuse a step it cannot place,
# a plan it cannot keep, and a rewrite that would swap the tool. The loop then executes whatever
# came through here without judging it a second time.
$live.Status = [ShareRibbon.Agent.AgentStatus]::Executing
$live.Interventions.Clear()
$plannedCode = '{"command":"InsertText","params":{"content":"SMOKE_A","position":"end"}}'
$plan = New-Object ShareRibbon.Agent.ExecutionPlan
$stepA = New-Object ShareRibbon.Agent.PlanStep
$stepA.StepNumber = 1
$stepA.Description = "smoke step A"
$stepA.Code = $plannedCode
$plan.Steps.Add($stepA)
$stepB = New-Object ShareRibbon.Agent.PlanStep
$stepB.StepNumber = 2
$stepB.Description = "smoke step B"
$stepB.Code = '{"command":"InsertText","params":{"content":"SMOKE_B","position":"end"}}'
$plan.Steps.Add($stepB)
$live.Plan = $plan

$rewriteCode = '{"command":"InsertText","params":{"content":"SMOKE_A_EDITED","position":"end"}}'
Assert-That ($kernel.QueueStepEdit("other-run", 0, "edited", $rewriteCode) -eq [ShareRibbon.Agent.StepEditResult]::NoLiveRun) `
    "a rewrite aimed at another run must be refused"
Assert-That ($kernel.QueueStepEdit($live.Id, 5, "edited", $rewriteCode) -eq [ShareRibbon.Agent.StepEditResult]::StepOutOfRange) `
    "a rewrite aimed at a step the plan does not have must be refused"
Assert-That ($kernel.QueueStepEdit($live.Id, 0, "   ", $rewriteCode) -eq [ShareRibbon.Agent.StepEditResult]::EmptyEdit) `
    "a rewrite without a description must be refused, the step would have no goal"
Assert-That ($kernel.QueueStepEdit($live.Id, 0, "edited", "not a tool call") -eq [ShareRibbon.Agent.StepEditResult]::CodeInvalid) `
    "a rewrite that is not a {command, params} object must be refused"
Assert-That ($kernel.QueueStepEdit($live.Id, 0, "edited", '{"command":"CreateChart","params":{}}') -eq [ShareRibbon.Agent.StepEditResult]::ToolChanged) `
    "a rewrite must not swap the tool behind the user's back"
$plan.Steps[0].Status = [ShareRibbon.Agent.StepStatus]::Running
Assert-That ($kernel.QueueStepEdit($live.Id, 0, "edited", $rewriteCode) -eq [ShareRibbon.Agent.StepEditResult]::StepNotEditable) `
    "a running step must not accept a rewrite that can no longer take effect"
$plan.Steps[0].Status = [ShareRibbon.Agent.StepStatus]::Pending
Assert-That ($live.Interventions.PendingCount() -eq 0) "a refused rewrite must leave nothing queued"
Assert-That (-not $kernel.QueueStepAction($live.Id, 0, "edit")) `
    "the verb-only door must refuse an edit that carries no text"
Assert-That ($kernel.QueueStepEdit($live.Id, 0, "  edited  ", $rewriteCode) -eq [ShareRibbon.Agent.StepEditResult]::Queued) `
    "a same-tool rewrite of a pending step must be queued"
$queuedEdit = $live.Interventions.Take(0)
Assert-That ($null -ne $queuedEdit) "the queued rewrite disappeared before the loop could take it"
Assert-That ($queuedEdit.Action -eq [ShareRibbon.Agent.StepIntervention]::Edit -and
             $queuedEdit.Code -eq $rewriteCode -and $queuedEdit.Description -eq "edited") `
    "the door must queue the trimmed description and the rewritten code together"

$sessionField.SetValue($kernel, $null)
Assert-That (-not $kernel.QueueStepAction($live.Id, 1, "retry")) "with no run in flight the door must stay shut"
Assert-That ($kernel.QueueStepEdit($live.Id, 0, "edited", $rewriteCode) -eq [ShareRibbon.Agent.StepEditResult]::NoLiveRun) `
    "with no run in flight the rewrite door must stay shut too"
$script:caseCount += 1
Write-Host "    ok AgentKernel is the only door, and it refuses id / verb / terminal / rewrite mismatches"

# ---- C. the loop consumes decisions at the step boundary ----
$Marks = @("SMOKE_A", "SMOKE_B", "SMOKE_C")

function New-StepJson {
    param([string]$Mark)
    return @{
        description = "smoke step $Mark"
        code        = (@{ command = "InsertText"; params = @{ content = $Mark; position = "end" } } |
            ConvertTo-Json -Compress -Depth 6)
        language    = "json"
    }
}

function New-ActionJson {
    param([string]$Mark)
    return (@{ thought = "smoke"; action = @{ tool = "InsertText"; params = @{ content = $Mark; position = "end" } } } |
        ConvertTo-Json -Compress -Depth 6)
}

function New-PlanJson {
    param([string[]]$StepMarks)
    return (@{
        understanding = "smoke"
        summary       = "smoke"
        capabilityGap = ""
        steps         = @($StepMarks | ForEach-Object { New-StepJson $_ })
    } | ConvertTo-Json -Compress -Depth 8)
}

function Invoke-InterventionRun {
    param(
        # How many host calls a marker loses before it starts winning. -1 never wins.
        [hashtable]$FailFirst,
        # Markers whose failures the executor declares non-recoverable, i.e. not fit for AI repair.
        [string[]]$NonRecoverable = @(),
        # Clicks replayed by the panel: OnStep sees a failure, Target gets the decision queued.
        [hashtable[]]$Clicks = @()
    )

    $script:planMarks = @($FailFirst.Keys | ForEach-Object { $_ } | Sort-Object { [Array]::IndexOf($Marks, $_) })
    $script:failFirst = $FailFirst
    $script:nonRecoverable = @($NonRecoverable)
    $script:clickPlan = @{}
    foreach ($click in $Clicks) { $script:clickPlan[[string]$click.OnStep] = $click }
    $script:hostCalls = @{}
    foreach ($mark in $Marks) { $script:hostCalls[$mark] = 0 }
    $script:lastHostCode = @{}
    $script:reactCalls = @{}
    foreach ($mark in $Marks) { $script:reactCalls[$mark] = 0 }
    $script:statusEvents = @()
    $script:reflections = 0
    $script:clicksQueued = 0

    $registry = [ShareRibbon.Agent.ToolRegistry]::new($null)
    if ($registry.LoadFromRuntimeDirectories($binDir) -le 0) {
        throw "no tool definitions loaded from $binDir - the run would normalize nothing"
    }
    $engine = New-Object ShareRibbon.Agent.LoopEngine($registry,
        (New-Object ShareRibbon.Agent.AgentMemory),
        (New-Object ShareRibbon.Agent.PromptManager((Join-Path $binDir "Prompts"))))

    $engine.SendAIRequest = [System.Func[string, string, System.Collections.Generic.List[ShareRibbon.HistoryMessage], System.Threading.Tasks.Task[string]]] {
        param($prompt, $system, $history)
        $combined = $prompt + [Environment]::NewLine + $system

        if ($combined -match '"complexity"') {
            return [System.Threading.Tasks.Task]::FromResult('{"goal":"smoke step intervention","complexity":"simple"}')
        }
        if ($combined -match '"toolId":\s*"') {
            # Step-level repair declines, so one failed step costs exactly one host call.
            return [System.Threading.Tasks.Task]::FromResult("smoke repair declined")
        }
        if ($combined -match '"strategy"') {
            $script:reflections += 1
            return [System.Threading.Tasks.Task]::FromResult('{"analysis":"nothing to add","strategy":"none","reason":"smoke"}')
        }
        if ($combined -match '"capabilityGap"') {
            return [System.Threading.Tasks.Task]::FromResult((New-PlanJson $script:planMarks))
        }

        # The react answer follows the step the loop is thinking about, so a retry re-issues the same call.
        # The counter is how the rewrite case proves the loop never asked the model about that step.
        foreach ($mark in $script:planMarks) {
            if ($combined -match $mark) {
                $script:reactCalls[$mark] += 1
                return [System.Threading.Tasks.Task]::FromResult((New-ActionJson $mark))
            }
        }
        return [System.Threading.Tasks.Task]::FromResult("smoke has no action to offer")
    }

    $engine.OnRequestApproval = [System.Func[ShareRibbon.Agent.ApprovalRequest, System.Threading.Tasks.Task[ShareRibbon.Agent.ApprovalDecision]]] {
        param($request)
        return [System.Threading.Tasks.Task]::FromResult([ShareRibbon.Agent.ApprovalDecision]::Session)
    }

    $registry.ExecuteCodeWithToolResult = [System.Func[string, string, bool, ShareRibbon.Agent.ToolResult]] {
        param($code, $language, $preview)
        $mark = $null
        foreach ($candidate in $script:planMarks) {
            if ($code -match $candidate) { $mark = $candidate; break }
        }
        if ($null -eq $mark) { throw "Smoke failed: host saw a call for no known step: $code" }
        $script:hostCalls[$mark] += 1
        $script:lastHostCode[$mark] = $code

        $failUntil = $script:failFirst[$mark]
        if ($failUntil -lt 0 -or $script:hostCalls[$mark] -le $failUntil) {
            $recoverable = -not ($script:nonRecoverable -contains $mark)
            return [ShareRibbon.Agent.ToolResult]::Failed(
                "smoke-host", "injected host failure", $null, "COM_ERROR", "injected host failure",
                "injected host failure", $recoverable, $null)
        }

        $observation = @{
            kind       = "write"
            summary    = "applied marker"
            changed    = $true
            targetRefs = @("Office:ActiveDocument")
        }
        return [ShareRibbon.Agent.ToolResult]::Succeed("smoke-host", "applied", $null, $observation, "undo-1")
    }

    # The panel's side of the click: it sees a step go to failed and queues a decision for the loop.
    # AgentKernel.QueueStepAction is proven above; from here the loop only ever reads the same inbox.
    $engine.OnStepStatusChanged = [Action[int, string, string]] {
        param($idx, $status, $msg)
        $script:statusEvents += "$idx|$status"
        if ($status -ne [ShareRibbon.Agent.LoopEngine]::StepStatusFailed) { return }
        $click = $script:clickPlan[[string]$idx]
        if ($null -eq $click -or $click.Remaining -eq 0) { return }
        if ($click.Remaining -gt 0) { $click.Remaining -= 1 }
        # The panel has two doors, exactly like the real one: a verb, or a rewrite with its text.
        if ($click.Action -eq "edit") {
            $script:session.Interventions.QueueEdit($click.Target, $click.Description, $click.Code)
        } else {
            $script:session.Interventions.Queue($click.Target,
                [ShareRibbon.Agent.AgentStepInterventions]::Parse($click.Action))
        }
        $script:clicksQueued += 1
    }

    $script:session = New-Object ShareRibbon.Agent.AgentSession("smoke step intervention", "Word", "smoke document")
    $result = $engine.RunAsync($script:session, "smoke system prompt", $null).GetAwaiter().GetResult()

    return [pscustomobject]@{
        Result       = $result
        Session      = $script:session
        Calls        = [hashtable]$script:hostCalls.Clone()
        LastCodes    = [hashtable]$script:lastHostCode.Clone()
        ReactCalls   = [hashtable]$script:reactCalls.Clone()
        Events       = @($script:statusEvents)
        Reflections  = $script:reflections
        ClicksQueued = $script:clicksQueued
    }
}

function Get-Click {
    param([int]$OnStep,
          [int]$Target = -1,
          [string]$Action = "retry",
          [int]$Remaining = 1,
          [string]$Description = "",
          [string]$Code = "")
    if ($Target -lt 0) { $Target = $OnStep }
    return @{
        OnStep      = $OnStep
        Target      = $Target
        Action      = $Action
        Remaining   = $Remaining
        Description = $Description
        Code        = $Code
    }
}

function Get-CallCount {
    param([hashtable]$Calls, [string]$Mark)
    return $Calls[$Mark]
}

# --- 1. nobody clicked: the boundary must not touch anyone -------------------------------
$case = Invoke-InterventionRun -FailFirst @{ SMOKE_A = 0; SMOKE_B = 0; SMOKE_C = 0 }
Assert-That ($case.Result.Success) "a clean three step run must still succeed: $($case.Result.Message)"
Assert-That ((Get-CallCount $case.Calls "SMOKE_A") -eq 1 -and
             (Get-CallCount $case.Calls "SMOKE_B") -eq 1 -and
             (Get-CallCount $case.Calls "SMOKE_C") -eq 1) `
    "every step must reach the host exactly once with nobody clicking, got $($case.Calls)"
Assert-That ($case.Session.Writes.PlanGeneration -eq 1) "the plan version has one owner, got $($case.Session.Writes.PlanGeneration)"
Assert-That ($case.Reflections -eq 0) "a run that never failed must not spend reflection"
$script:caseCount += 1
Write-Host "    ok no click, no change: the settled-step guard leaves a clean run alone"

# --- 2. retry one failed step: only that step re-runs ------------------------------------
$case = Invoke-InterventionRun -FailFirst @{ SMOKE_A = 0; SMOKE_B = 1; SMOKE_C = 0 } `
    -Clicks @(Get-Click -OnStep 1 -Action "retry")
Assert-That ($case.ClicksQueued -eq 1) "the panel never saw a failed step, so the click path was not exercised"
Assert-That ($case.Result.Success) "a step the user asked to retry and that then worked must finish the job: $($case.Result.Message)"
Assert-That ((Get-CallCount $case.Calls "SMOKE_B") -eq 2) "the retried step must reach the host twice, got $(Get-CallCount $case.Calls "SMOKE_B")"
Assert-That ((Get-CallCount $case.Calls "SMOKE_A") -eq 1) "an already successful write must not be replayed by a retry elsewhere"
Assert-That ((Get-CallCount $case.Calls "SMOKE_C") -eq 1) "the step after the retry must run exactly once, got $(Get-CallCount $case.Calls "SMOKE_C")"
Assert-That ($case.Session.Writes.PlanGeneration -eq 1) "a user retry is not a replan, the plan version must stay 1"
Assert-That ($case.Reflections -eq 0) "a user retry must not spend the AI repair budget"
Assert-That ($case.Session.Interventions.PendingCount() -eq 0) "the boundary must drain the inbox, $($case.Session.Interventions.PendingCount()) left over"
$retryPath = @($case.Events | Where-Object { $_ -like "1|*" })
Assert-That (($retryPath -join ",") -match "1\|failed,1\|pending") `
    "the card must be told the step is queued again, got: $($retryPath -join ',')"
Assert-That ($retryPath[-1] -eq "1|completed") "the retried step must report its real outcome, got $($retryPath[-1])"
$script:caseCount += 1
Write-Host "    ok retry re-runs only the clicked step, spends no replan and does not move the plan version"

# --- 3. skip a step that will never work --------------------------------------------------
$case = Invoke-InterventionRun -FailFirst @{ SMOKE_A = 0; SMOKE_B = -1; SMOKE_C = 0 } `
    -Clicks @(Get-Click -OnStep 1 -Action "skip")
Assert-That ((Get-CallCount $case.Calls "SMOKE_B") -eq 1) "a step the user skipped must not be attempted again, got $(Get-CallCount $case.Calls "SMOKE_B")"
Assert-That ((Get-CallCount $case.Calls "SMOKE_C") -eq 1) "skipping one step must not stop the steps behind it"
Assert-That ($case.Session.Plan.Steps[1].Status -eq [ShareRibbon.Agent.StepStatus]::Skipped) `
    "the skipped step must carry the skip state the loop decided, got $($case.Session.Plan.Steps[1].Status)"
Assert-That ($case.Result.Success) "a run where the user chose to skip is not a failure: $($case.Result.Message)"
Assert-That ($case.Session.Writes.PlanGeneration -eq 1) "a user skip is not a replan either"
$script:caseCount += 1
Write-Host "    ok skip settles the step once and lets the rest of the plan finish"

# --- 4. a non-recoverable failure waits for the queued decision --------------------------
# Without the deferral this is the shape the user would get: the executor says "not recoverable",
# the loop returns a terminal failure in the same step body, and the retry click they just made
# becomes a message about a run that already ended.
$case = Invoke-InterventionRun -FailFirst @{ SMOKE_A = 0; SMOKE_B = 1; SMOKE_C = 0 } `
    -NonRecoverable @("SMOKE_B")
Assert-That (-not $case.Result.Success) "an un-recoverable step with no click must still report failure"
Assert-That ((Get-CallCount $case.Calls "SMOKE_C") -eq 0) "the run must stop at the unrecoverable step, SMOKE_C reached the host"

$case = Invoke-InterventionRun -FailFirst @{ SMOKE_A = 0; SMOKE_B = 1; SMOKE_C = 0 } `
    -NonRecoverable @("SMOKE_B") -Clicks @(Get-Click -OnStep 1 -Action "retry")
Assert-That ($case.ClicksQueued -eq 1) "the unrecoverable failure never reached the panel handler, so this case proved nothing"
Assert-That ($case.Result.Success) "a retry the user queued must outrank the executor's own non-recoverable verdict: $($case.Result.Message)"
Assert-That ((Get-CallCount $case.Calls "SMOKE_B") -eq 2) "the unrecoverable step must run once more and then win"
Assert-That ((Get-CallCount $case.Calls "SMOKE_C") -eq 1) "the steps behind it must continue after the recovery"
$script:caseCount += 1
Write-Host "    ok a non-recoverable failure defers to a queued decision instead of ending the run"

# --- 5. rewinding goes to the earliest settled step the user named -----------------------
# B never recovers, so the loop is about to move past it; the click points back at A, which already
# succeeded. The boundary must land on A alone and then walk forward over every settled step - B is not
# what the user asked for, and replaying it would replay a write for a decision that named a different step.
$case = Invoke-InterventionRun -FailFirst @{ SMOKE_A = 0; SMOKE_B = -1; SMOKE_C = 0 } `
    -Clicks @(Get-Click -OnStep 1 -Target 0 -Action "retry")
Assert-That ((Get-CallCount $case.Calls "SMOKE_A") -eq 2) "the step the user named must re-run even though it had already succeeded, got $(Get-CallCount $case.Calls "SMOKE_A")"
Assert-That ((Get-CallCount $case.Calls "SMOKE_B") -eq 1) "the walk forward must not replay a step nobody pointed at, got $(Get-CallCount $case.Calls "SMOKE_B")"
Assert-That ((Get-CallCount $case.Calls "SMOKE_C") -eq 1) "the walk must stop at the first step still waiting, got $(Get-CallCount $case.Calls "SMOKE_C")"
$script:caseCount += 1
Write-Host "    ok a rewind to an earlier step re-runs that step only"

# --- 6. a click addressed to a step that is no longer in the plan is dropped -------------
$case = Invoke-InterventionRun -FailFirst @{ SMOKE_A = 0; SMOKE_B = -1; SMOKE_C = 0 } `
    -Clicks @(Get-Click -OnStep 1 -Target 42 -Action "retry")
Assert-That ($case.Session.Interventions.PendingCount() -eq 0) "a decision for step 42 of a three step plan must be consumed and dropped"
Assert-That ((Get-CallCount $case.Calls "SMOKE_A") -eq 1 -and (Get-CallCount $case.Calls "SMOKE_C") -eq 1) `
    "a dropped decision must not disturb any real step, got $($case.Calls)"
$script:caseCount += 1
Write-Host "    ok a decision for a step number the plan does not have is dropped, not applied to a neighbour"

# --- 7. a user who clicks retry forever is still bounded --------------------------------
# A user retry resets the no-progress counter, because the user giving new direction is progress.
# That removes the reflect guard as the thing that stops this, so MaxIterations must be.
$case = Invoke-InterventionRun -FailFirst @{ SMOKE_A = 0; SMOKE_B = -1; SMOKE_C = 0 } `
    -Clicks @(Get-Click -OnStep 1 -Action "retry" -Remaining -1)
Assert-That ($null -ne $case.Result) "the run must return even while the user keeps asking for the same retry"
Assert-That (-not $case.Result.Success) "a step that never works must never be reported as done"
Assert-That ($case.ClicksQueued -gt 1) "the panel must have clicked more than once, otherwise this case judged nothing"
Assert-That ($case.Session.CurrentIteration -le 15) `
    "MaxIterations still bounds the loop after the third entry clause, got $($case.Session.CurrentIteration)"
Assert-That ((Get-CallCount $case.Calls "SMOKE_A") -eq 1) "the bounded spin must stay on the step the user pointed at, SMOKE_A reached the host $(Get-CallCount $case.Calls "SMOKE_A") times"
Assert-That ($case.Session.Writes.PlanGeneration -eq 1) "user retries must not silently become replans, got version $($case.Session.Writes.PlanGeneration)"
$script:caseCount += 1
Write-Host "    ok endless retries stay bounded by MaxIterations and never invent a plan version"

# --- 8. rewriting a failed step: the saved parameters are what runs, and only that step --------
# The acceptance is not "the text was stored" but "the host saw the rewritten parameters without
# the model being asked to rethink the step": a prompt-only edit would let the model keep its own
# plan and the card would be describing an execution that never happened.
$editedB = '{ "command": "InsertText", "params": { "content": "SMOKE_B_EDITED", "position": "end" } }'
$case = Invoke-InterventionRun -FailFirst @{ SMOKE_A = 0; SMOKE_B = 1; SMOKE_C = 0 } `
    -Clicks @(Get-Click -OnStep 1 -Action "edit" -Description "smoke edited B" -Code $editedB)
Assert-That ($case.ClicksQueued -eq 1) "the rewrite never reached the inbox, so this case proved nothing"
Assert-That ($case.Result.Success) "a rewritten step that then works must finish the job: $($case.Result.Message)"
Assert-That ((Get-CallCount $case.Calls "SMOKE_B") -eq 2) "the rewritten step must run again, got $(Get-CallCount $case.Calls "SMOKE_B")"
Assert-That ((Get-CallCount $case.Calls "SMOKE_A") -eq 1 -and (Get-CallCount $case.Calls "SMOKE_C") -eq 1) `
    "a rewrite must not replay the steps around it, got $($case.Calls)"
$sent = $case.LastCodes["SMOKE_B"] | ConvertFrom-Json
Assert-That ($sent.command -eq "InsertText" -and $sent.params.content -eq "SMOKE_B_EDITED" -and
             $sent.params.position -eq "end") `
    "the host must receive the saved parameters, got: $($case.LastCodes['SMOKE_B'])"
Assert-That ($case.ReactCalls["SMOKE_B"] -eq 1) `
    "the rewritten step was planned again by the model ($($case.ReactCalls['SMOKE_B']) rounds), so the saved parameters were only a suggestion"
Assert-That ($case.Session.Plan.Steps[1].PinnedByUser) `
    "the loop must mark the step as user-defined, otherwise the next run drops the saved text"
Assert-That ($case.Session.Plan.Steps[1].Description -eq "smoke edited B") `
    "the plan and the card disagree about this step's description"
Assert-That ($case.Session.Writes.PlanGeneration -eq 1) `
    "a rewrite is not a replan: the card must stay at plan version 1, got $($case.Session.Writes.PlanGeneration)"
Assert-That ($case.Session.Interventions.PendingCount() -eq 0) "the boundary must drain the rewritten decision"
$editPath = @($case.Events | Where-Object { $_ -like "1|*" })
Assert-That (($editPath -join ",") -match "1\|failed,1\|pending") `
    "the card must be told the rewritten step is queued again, got: $($editPath -join ',')"
$script:caseCount += 1
Write-Host "    ok a rewritten step runs from its saved parameters with no model round and no plan version bump"

Write-Host "PASS step intervention: $script:caseCount cases (2 inbox/door contracts + 7 injected-failure runs + 1 rewrite run over the real LoopEngine)"
