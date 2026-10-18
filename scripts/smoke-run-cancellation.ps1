# Stop-convergence smoke: pressing stop must converge the run to Cancelled without
# turning an already-received model reply into a host write operation.
# Case A cancels while the task is still being analysed; case B cancels right as the
# ReAct thought comes back, which is the window the old step-boundary check missed.
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

function Invoke-SmokeStopRun {
    param([int]$CancelOnCall)

    $script:llmCalls = 0
    $script:hostCalls = @()
    $script:cancelOnCall = $CancelOnCall
    $script:runCts = New-Object System.Threading.CancellationTokenSource

    $kernel = New-Object ShareRibbon.Agent.AgentKernel
    $kernel.SendAIRequest = [System.Func[string, string, System.Collections.Generic.List[ShareRibbon.HistoryMessage], System.Threading.Tasks.Task[string]]] {
        param($prompt, $system, $history)
        $script:llmCalls += 1
        # The stop click lands while this request is in flight; the reply must never reach the host.
        if ($script:llmCalls -eq $script:cancelOnCall) { $script:runCts.Cancel() }
        return [System.Threading.Tasks.Task]::FromResult($cannedResponse)
    }
    $kernel.ExecuteCodeWithToolResult = [System.Func[string, string, bool, ShareRibbon.Agent.ToolResult]] {
        param($code, $language, $preview)
        $script:hostCalls += $code
        return [ShareRibbon.Agent.ToolResult]::Succeed("smoke-host", "executed")
    }
    $kernel.Initialize()

    $harness = New-Object ShareRibbon.Agent.Harness.OfficeHarness($kernel, (New-Object ShareRibbon.Agent.Harness.NoopRunTraceStore))
    $turn = New-Object ShareRibbon.Agent.Harness.UserTurn
    $turn.SessionId = "smoke-stop-$CancelOnCall"
    $turn.AppType = "Word"
    $turn.Text = "delete the selection"
    $turn.Mode = "agent"

    $result = $harness.RunAsync($turn, $script:runCts.Token).GetAwaiter().GetResult()
    return [pscustomobject]@{
        Status    = $result.Status.ToString()
        HostCalls = $script:hostCalls.Count
        LlmCalls  = $script:llmCalls
    }
}

foreach ($case in @(1, 3)) {
    $outcome = Invoke-SmokeStopRun -CancelOnCall $case
    $outcome | Format-List

    if ($outcome.Status -ne "Cancelled") {
        throw "Smoke failed: cancelling on LLM call $case must converge to Cancelled, got $($outcome.Status)."
    }
    if ($outcome.HostCalls -gt 0) {
        throw "Smoke failed: a cancelled run still pushed a write to the host executor (call $case)."
    }
    if ($outcome.LlmCalls -gt $case) {
        throw "Smoke failed: the run kept asking the model after the user stopped it (call $case)."
    }
}
