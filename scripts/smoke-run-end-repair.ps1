param(
    [string]$Configuration = "Debug"
)

$ErrorActionPreference = "Stop"

# E6-05b behaviour smoke: the run-end contract gate.
# check-task-package-contract.ps1 proves the assertions are written legal; smoke-task-package-verifier.ps1
# proves the verifier judges one ledger correctly. Neither proves that a run whose package is unmet actually
# gets a repair round, stops once the replan budget is spent, or is left alone when it has no entry identity.
# This one drives the real LoopEngine with a scripted model and a scripted host executor.

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

# The gate resolves a package from the entry action, exactly like a real run, so the fixtures below
# would judge the wrong contract in silence if those bindings drifted. The declared onFail is checked
# per package because the route must follow the package, not a hardcoded strategy.
@(
    @{ Name = "excel.clean-and-fix"; Action = "excel.selection.formula"; OnFail = "repair" },
    @{ Name = "excel.insight"; Action = "excel.selection.explain"; OnFail = "report" }
) | ForEach-Object {
    $resolved = [ShareRibbon.Agent.TaskPackageCatalog]::FindByAction($_.Action)
    if ($null -eq $resolved) {
        throw "$($_.Name) no longer resolves from $($_.Action) - the run-end gate would never see this package"
    }
    if ($resolved.Id -ne $_.Name) {
        throw "$($_.Action) now belongs to $($resolved.Id), not $($_.Name) - the fixtures would judge the wrong contract"
    }
    $declared = [ShareRibbon.Agent.TaskPackageOutcomeVerifier]::OnFail($resolved)
    if ($declared -ne $_.OnFail) {
        throw "$($_.Name) declares onFail=$declared, not $($_.OnFail) - the scenarios below assume the other route"
    }
}

$WriteTools = @("CleanData", "WriteData", "ApplyFormula", "FormatRange")

# ToolRegistry validates required parameters, so each scripted tool needs its own payload shape.
$ToolParams = @{
    DataAnalysis = @{ sourceRange = "A1:A10"; type = "summary" }
    CleanData    = @{ range = "A1:A10"; operation = "trim" }
}

function Get-ToolParams {
    param([string]$ToolId)
    if ($ToolParams.ContainsKey($ToolId)) { return $ToolParams[$ToolId] }
    throw "no scripted payload for tool $ToolId - add it to ToolParams"
}

function New-StepJson {
    param([string]$ToolId)
    return @{
        description = "smoke $ToolId"
        code        = (@{ command = $ToolId; params = (Get-ToolParams $ToolId) } | ConvertTo-Json -Compress -Depth 6)
        language    = "json"
    }
}

function New-PlanJson {
    param([string[]]$ToolIds)
    return (@{
        understanding = "smoke"
        summary       = "smoke"
        capabilityGap = ""
        steps         = @($ToolIds | ForEach-Object { New-StepJson $_ })
    } | ConvertTo-Json -Compress -Depth 8)
}

function New-ActionJson {
    param([string]$ToolId)
    return (@{ thought = "smoke"; action = @{ tool = $ToolId; params = (Get-ToolParams $ToolId) } } |
        ConvertTo-Json -Compress -Depth 6)
}

function Set-PlanOffered {
    # Whatever plan the scripted model hands out, the react answers must follow its steps.
    param([string[]]$ToolIds)
    $script:reactActions = @($ToolIds | ForEach-Object { New-ActionJson $_ })
    $script:reactCursor = 0
}

function Invoke-SmokeRun {
    param(
        [string]$EntryActionId,
        [string[]]$FirstPlan,
        [string[]]$RepairPlan,
        # How many replan offers the model is willing to make. 0 means it cooperates with nothing.
        [int]$RepairOffers = 0
    )

    $script:reactActions = @()
    $script:reactCursor = 0
    $script:hostCommands = @()
    $script:prompts = @()
    $script:reflections = 0
    $script:sawRepairAsk = $false
    $script:plansIssued = 0
    $script:planQueue = @((New-PlanJson $FirstPlan))
    for ($i = 0; $i -lt $RepairOffers; $i++) {
        # Past the offers the model returns something that is not a plan at all - a real failure mode.
        $script:planQueue += $(if ($RepairPlan.Count -gt 0) { New-PlanJson $RepairPlan } else { '{"steps":[]}' })
    }
    $script:firstPlan = $FirstPlan
    $script:repairPlan = $RepairPlan
    $script:repairOffers = $RepairOffers

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
        $script:prompts += $combined

        if ($combined -match '"complexity"') {
            return [System.Threading.Tasks.Task]::FromResult('{"goal":"smoke run-end repair","complexity":"simple"}')
        }
        if ($combined -match '"toolId":\s*"') {
            # Step-level repair: decline it, the loop must then move on with the plan.
            return [System.Threading.Tasks.Task]::FromResult("smoke repair declined")
        }
        if ($combined -match '"strategy"') {
            $script:reflections += 1
            if ($combined -match "RUN_END_CONTRACT_UNMET") { $script:sawRepairAsk = $true }
            if ($script:repairOffers -le 0) {
                return [System.Threading.Tasks.Task]::FromResult('{"analysis":"nothing to add","strategy":"none","reason":"smoke"}')
            }
            $script:repairOffers -= 1
            return [System.Threading.Tasks.Task]::FromResult('{"analysis":"the declared write never landed","strategy":"replan","reason":"smoke"}')
        }
        if ($combined -match '"capabilityGap"') {
            $slot = [Math]::Min($script:plansIssued, $script:planQueue.Count - 1)
            $script:plansIssued += 1
            Set-PlanOffered $(if ($slot -eq 0) { $script:firstPlan } else { $script:repairPlan })
            return [System.Threading.Tasks.Task]::FromResult($script:planQueue[$slot])
        }
        if ($script:reactActions.Count -eq 0) {
            return [System.Threading.Tasks.Task]::FromResult("smoke has no action to offer")
        }
        $cursor = [Math]::Min($script:reactCursor, $script:reactActions.Count - 1)
        $script:reactCursor += 1
        return [System.Threading.Tasks.Task]::FromResult($script:reactActions[$cursor])
    }

    $engine.OnRequestApproval = [System.Func[ShareRibbon.Agent.ApprovalRequest, System.Threading.Tasks.Task[ShareRibbon.Agent.ApprovalDecision]]] {
        param($request)
        return [System.Threading.Tasks.Task]::FromResult([ShareRibbon.Agent.ApprovalDecision]::Session)
    }

    $registry.ExecuteCodeWithToolResult = [System.Func[string, string, bool, ShareRibbon.Agent.ToolResult]] {
        param($code, $language, $preview)
        $script:hostCommands += $code
        foreach ($toolId in $WriteTools) {
            if ($code -match $toolId) {
                $writeObservation = @{ kind = "write"; summary = "applied"; changed = $true }
                return [ShareRibbon.Agent.ToolResult]::Succeed("smoke-host", "applied", $null, $writeObservation, "undo-1")
            }
        }
        $readObservation = @{ kind = "read"; summary = "inspected"; changed = $false }
        return [ShareRibbon.Agent.ToolResult]::Succeed("smoke-host", "inspected", $null, $readObservation)
    }

    $session = New-Object ShareRibbon.Agent.AgentSession("smoke run-end repair", "Excel", "smoke workbook")
    $session.EntryActionId = $EntryActionId
    $result = $engine.RunAsync($session, "smoke system prompt", $null).GetAwaiter().GetResult()

    return [pscustomobject]@{
        Result       = $result
        Session      = $session
        Reflections  = $script:reflections
        SawRepairAsk = $script:sawRepairAsk
        HostCalls    = $script:hostCommands.Count
        Prompts      = @($script:prompts)
    }
}

function Assert-That {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "Smoke failed: $Message" }
}

$script:caseCount = 0

# --- 1. an unmet contract gets a repair round, and the repaired run is allowed to succeed --------
$case = Invoke-SmokeRun -EntryActionId "excel.selection.formula" `
    -FirstPlan @("DataAnalysis") -RepairPlan @("CleanData") -RepairOffers 1
Assert-That ($case.Reflections -eq 1) "unmet package must ask for repair exactly once, got $($case.Reflections)"
Assert-That $case.SawRepairAsk "the repair request must carry the unmet success criteria as its observation"
$namedFacts = @($case.Prompts | Where-Object { $_ -match "writtenCount" -and $_ -match "changedCount" }).Count
Assert-That ($namedFacts -ge 1) "the observation must name the unmet facts, not just say the run was unsatisfying"
Assert-That ($case.Session.Writes.PlanGeneration -eq 2) "repair must advance the plan version, got $($case.Session.Writes.PlanGeneration)"
Assert-That ($case.Session.ChangeLog.WrittenCount -eq 1) "the repaired run must land the missing write once, got $($case.Session.ChangeLog.WrittenCount)"
Assert-That ($case.Result.Success) "a run that met its contract after repair must not still report unmet: $($case.Result.Message)"
$script:caseCount += 1
Write-Host "    ok unmet excel.clean-and-fix gets one repair round and then succeeds"

# --- 2. the budget is real: two repair rounds, then report - never an endless repair loop --------
$case = Invoke-SmokeRun -EntryActionId "excel.selection.formula" `
    -FirstPlan @("DataAnalysis") -RepairPlan @("DataAnalysis") -RepairOffers 4
Assert-That ($case.Reflections -eq 2) "run-end repair must stop at MaxReplanAttempts (2), got $($case.Reflections)"
Assert-That (-not $case.Result.Success) "a package that stayed unmet must never report success"
Assert-That ($case.Result.Message -match "excel\.clean-and-fix") "the failure must name the package, got: $($case.Result.Message)"
Assert-That ($case.Session.ChangeLog.WrittenCount -eq 0) "no write landed; the ledger must still say so, got $($case.Session.ChangeLog.WrittenCount)"
$script:caseCount += 1
Write-Host "    ok repeated unmet repairs stop at the replan budget and still report unmet"

# --- 3. no entry identity: the gate must not bind a package by guessing --------------------------
$case = Invoke-SmokeRun -EntryActionId "" -FirstPlan @("DataAnalysis") -RepairPlan @("CleanData") -RepairOffers 1
Assert-That ($case.Reflections -eq 0) "a free-text run has no contract, so the gate must not ask for repair, got $($case.Reflections)"
Assert-That ($case.Session.Writes.PlanGeneration -eq 1) "the plan version must stay 1 without a repair, got $($case.Session.Writes.PlanGeneration)"
Assert-That ($case.Result.Success) "the same read-only run without a package is simply a read-only run: $($case.Result.Message)"
$script:caseCount += 1
Write-Host "    ok free text is not bound to any package at run end"

# --- 4. a satisfied contract does not trigger a repair ------------------------------------------
$case = Invoke-SmokeRun -EntryActionId "excel.selection.explain" -FirstPlan @("DataAnalysis") `
    -RepairPlan @("CleanData") -RepairOffers 1
Assert-That ($case.Reflections -eq 0) "excel.insight was met, so no repair may run, got $($case.Reflections)"
Assert-That ($case.Result.Success) "a met read-only contract must finish as success: $($case.Result.Message)"
Assert-That ($case.Session.Writes.PlanGeneration -eq 1) "a met contract must not consume a plan version, got $($case.Session.Writes.PlanGeneration)"
$script:caseCount += 1
Write-Host "    ok a package whose contract is met is left alone"

# --- 5. a package that declares report must not be sent to repair -------------------------------
$case = Invoke-SmokeRun -EntryActionId "excel.selection.explain" -FirstPlan @("CleanData") `
    -RepairPlan @("DataAnalysis") -RepairOffers 1
Assert-That ($case.Reflections -eq 0) "excel.insight declares onFail=report, so the route may not ask for a repair plan, got $($case.Reflections)"
Assert-That (-not $case.Result.Success) "a read-only package that wrote into the sheet must never report success"
Assert-That ($case.Result.Message -match "excel\.insight") "the failure must name excel.insight, got: $($case.Result.Message)"
Assert-That ($case.Session.ChangeLog.WrittenCount -eq 1) "the write really landed and the ledger must carry it, got $($case.Session.ChangeLog.WrittenCount)"
$script:caseCount += 1
Write-Host "    ok onFail=report is read from the package, not hardcoded to repair"

# --- 6. a model that offers no usable plan must not be rewarded with a repaired success ---------
$case = Invoke-SmokeRun -EntryActionId "excel.selection.formula" `
    -FirstPlan @("DataAnalysis") -RepairPlan @() -RepairOffers 0
Assert-That ($case.Reflections -eq 1) "the unmet contract must still ask once, got $($case.Reflections)"
Assert-That ($case.Session.Writes.PlanGeneration -eq 1) "an unusable repair answer must not advance the plan version, got $($case.Session.Writes.PlanGeneration)"
Assert-That (-not $case.Result.Success) "no repair plan means the contract is still unmet, so success would be a lie"
Assert-That ($case.Result.Message -match "excel\.clean-and-fix") "the failure must name the package, got: $($case.Result.Message)"
$script:caseCount += 1
Write-Host "    ok a repair round that yields no plan reports unmet instead of succeeding"

Write-Host "PASS run-end repair: $script:caseCount contract cases over the real LoopEngine"
