param(
    [string]$Configuration = "Debug"
)

$ErrorActionPreference = "Stop"
$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$assemblyPath = Join-Path $repoRoot "ShareRibbon\bin\$Configuration\ShareRibbon.dll"
$catalogPath = Join-Path $repoRoot "tests\golden\l0\catalog.json"
$toolsPath = Join-Path $repoRoot "ShareRibbon\Tools"

if (-not (Test-Path -LiteralPath $assemblyPath)) {
    throw "ShareRibbon assembly not found: $assemblyPath. Build the code projects first."
}
if (-not (Test-Path -LiteralPath $catalogPath)) {
    throw "Golden L0 catalog not found: $catalogPath"
}

Add-Type -Path $assemblyPath
$catalog = Get-Content -LiteralPath $catalogPath -Raw | ConvertFrom-Json
$results = @()

foreach ($case in $catalog.cases) {
    if ($case.kind -eq "contract") {
        if ($case.contractType -eq "context_pack") {
            $officeContext = New-Object ShareRibbon.Agent.Context.OfficeContext
            $officeContext.AppType = "Word"
            $officeContext.Selection = New-Object ShareRibbon.Agent.Context.SelectionInfo
            $officeContext.Selection.Address = "Selection"
            $officeContext.Selection.ItemCount = 1
            $officeContext.Selection.DataType = "text"
            $officeContext.Selection.Preview = ("x" * 7000)
            $pack = [ShareRibbon.Agent.Context.ContextPack]::FromOfficeContext($officeContext, ("y" * 9000), 4000)
            if ($pack.SchemaVersion -ne $case.expectedSchemaVersion -or
                $pack.Budget.Strategy -ne $case.expectedBudgetStrategy -or
                -not $pack.Budget.Truncated -or
                $pack.Budget.UsedChars -gt $pack.Budget.MaxChars -or
                [string]::IsNullOrWhiteSpace($pack.ToJson())) {
                throw "Golden case failed: $($case.id)"
            }
        }
        elseif ($case.contractType -eq "approval_api") {
            $interfaceType = [ShareRibbon.Agent.Harness.IOfficeHarness]
            $approve = $interfaceType.GetMethod("ApproveAsync")
            $cancel = $interfaceType.GetMethod("CancelAsync")
            $resume = $interfaceType.GetMethod("ResumeAsync")
            $awaiting = [System.Enum]::Parse([ShareRibbon.Agent.Harness.HarnessRunStatus], "AwaitingApproval")
            if ($null -eq $approve -or $null -eq $cancel -or $null -eq $resume -or $null -eq $awaiting) {
                throw "Golden case failed: $($case.id)"
            }
        }
        elseif ($case.contractType -eq "tool_result_only") {
            $serviceType = [ShareRibbon.CodeExecutionService]
            $legacyMethods = @("ExecuteCodeWithResult", "ExecuteJsonCommand")
            foreach ($methodName in $legacyMethods) {
                if ($null -ne $serviceType.GetMethod($methodName)) {
                    throw "Golden case $($case.id): legacy method still exists: $methodName"
                }
            }
            if ($null -ne $serviceType.GetProperty("JsonCommandExecutor")) {
                throw "Golden case $($case.id): legacy JsonCommandExecutor still exists"
            }

            $sourceFiles = @(
                "ShareRibbon\Controls\BaseChatControl.vb",
                "WordAi\ChatControl.vb",
                "ExcelAi\ChatControl.vb",
                "PowerPointAi\ChatControl.vb"
            )
            foreach ($sourceFile in $sourceFiles) {
                $source = Get-Content -LiteralPath (Join-Path $repoRoot $sourceFile) -Raw
                if ($source -match 'Protected\s+(Overridable|Overrides)\s+Function\s+ExecuteJsonCommand\s*\(') {
                    throw "Golden case $($case.id): Boolean JSON override found in $sourceFile"
                }
            }
        }
        elseif ($case.contractType -eq "powerpoint_create_fallback") {
            $intentService = [ShareRibbon.IntentRecognitionService]::new("PowerPoint")
            $intent = $intentService.IdentifyIntent("ppt")
            if ($intent.OfficeIntent.ToString() -ne "SLIDE_CREATE" -or $intent.Confidence -lt 0.5) {
                throw "Golden case $($case.id): explicit PPT creation request fell back to $($intent.OfficeIntent)"
            }

            $parseMethod = [ShareRibbon.IntentRecognitionService].GetMethod(
                "ParseLLMIntentResponse",
                [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic)
            $malformedResult = $parseMethod.Invoke($intentService, @('{"intentType":', 'ppt'))
            if ($null -ne $malformedResult) {
                throw "Golden case $($case.id): malformed LLM JSON returned a default intent and would overwrite the fallback"
            }

            $routeSource = Get-Content -LiteralPath (Join-Path $repoRoot "ShareRibbon\Controls\Services\ChatRoutingOrchestrator.vb") -Raw
            if ($routeSource -notmatch 'taskSpecRequiresExecution' -or
                $routeSource -notmatch 'Not\s+taskSpecRequiresExecution') {
                throw "Golden case $($case.id): TaskSpec execution fallback is missing"
            }
        }
        elseif ($case.contractType -eq "powerpoint_com_compat") {
            $classified = [ShareRibbon.ExceptionClassifier]::Classify(
                [InvalidCastException]::new("QueryInterface E_NOINTERFACE 不支持此接口"))
            if ($classified.ErrorCode -ne "COM_ERROR" -or $classified.Recoverable) {
                throw "Golden case $($case.id): E_NOINTERFACE must be terminal"
            }

            $rendererSource = Get-Content -LiteralPath (Join-Path $repoRoot "PowerPointAi\Design\PowerPointSceneRenderer.vb") -Raw
            if ($rendererSource -match 'As\s+(FillFormat|LineFormat|ShadowFormat|ColorFormat|Adjustments|PictureFormat|TextFrame2|TextRange2|Font2|ParagraphFormat2)') {
                throw "Golden case $($case.id): renderer reintroduced a strongly typed Office style interface"
            }

            $pptChatSource = Get-Content -LiteralPath (Join-Path $repoRoot "PowerPointAi\ChatControl.vb") -Raw
            if ($pptChatSource -notmatch 'ProfessionalDeckExecutor\.ExecuteAsToolResult\(\s*GetPowerPointEnvelopeParams\(envelope\),\s*preview\)') {
                throw "Golden case $($case.id): CreateSlides preview flag is not propagated"
            }
        }
        elseif ($case.contractType -eq "excel_observer_compat") {
            $excelChatSource = Get-Content -LiteralPath (Join-Path $repoRoot "ExcelAi\ChatControl.vb") -Raw
            if ($excelChatSource -match 'BuildExcelRangePreview\(target\s+As\s+Microsoft\.Office\.Interop\.Excel\.Range' -or
                $excelChatSource -match 'CountExcelFormulaErrors\(target\s+As\s+Microsoft\.Office\.Interop\.Excel\.Range' -or
                $excelChatSource -match 'TryCast\(app\.Selection,\s*Microsoft\.Office\.Interop\.Excel\.Range\)') {
                throw "Golden case $($case.id): Excel observer reintroduced a strongly typed Range cast"
            }
            if ($excelChatSource -notmatch 'CodeObservationFailed' -or
                $excelChatSource -notmatch 'recoverable:=False') {
                throw "Golden case $($case.id): observer failures can still repeat an uncertain write"
            }
        }
        else {
            $observation = [Newtonsoft.Json.Linq.JObject]::Parse('{"kind":"write","summary":"changed","changed":true}')
            $result = [ShareRibbon.Agent.ToolResult]::Succeed("contract.test", "ok", $null, $observation)
            if (-not $result.Success -or $null -eq $result.Observation -or [string]::IsNullOrWhiteSpace($result.ToObserveSummary())) {
                throw "Golden case failed: $($case.id)"
            }
        }
        $results += [pscustomobject]@{ Id = $case.id; Status = "PASS"; ErrorCode = ""; HostCalls = 0 }
        continue
    }

    $registry = [ShareRibbon.Agent.ToolRegistry]::new($null)
    $registry.LoadFromDirectory($toolsPath)
    $hostCalls = [System.Collections.Generic.List[string]]::new()
    $registry.ExecuteCodeWithToolResult = [System.Func[string,string,bool,ShareRibbon.Agent.ToolResult]] {
        param($code, $language, $preview)
        $hostCalls.Add($code)
        return [ShareRibbon.Agent.ToolResult]::Succeed("fake-host", "executed")
    }

    $context = New-Object ShareRibbon.Agent.ToolExecutionContext
    $context.AppType = $case.appType
    $context.RunId = $case.id
    $context.CorrelationId = $case.id
    $context.PrimarySkillName = "golden-l0"
    $context.AllowedTools = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($tool in $case.allowedTools) {
        [void]$context.AllowedTools.Add([string]$tool)
    }
    $context.EnforceAllowedTools = $true
    if ($null -ne $case.approvedTools) {
        foreach ($tool in $case.approvedTools) {
            $context.ApproveTool([string]$tool)
        }
    }

    $params = New-Object Newtonsoft.Json.Linq.JObject
    if ($null -ne $case.params) {
        # JObject.FromObject() on a PSCustomObject serializes PowerShell's wrapper and yields a
        # single "CliXml" property, so fixture params never reached the tool. Flatten first.
        $paramBag = [System.Collections.Generic.Dictionary[string, object]]::new()
        foreach ($prop in $case.params.PSObject.Properties) {
            $paramBag[$prop.Name] = $prop.Value
        }
        $params = [Newtonsoft.Json.Linq.JObject]::FromObject($paramBag)
    }

    $task = $registry.ExecuteToolAsync($context, [string]$case.toolId, $params)
    $result = $task.GetAwaiter().GetResult()
    $expectedSuccess = ($case.expectedSuccess -eq $true)
    if ($result.Success -ne $expectedSuccess) {
        throw "Golden case $($case.id) expected success=$expectedSuccess, got $($result.Success) code=$($result.ErrorCode) message=$($result.Message)"
    }
    if (-not $expectedSuccess -and $result.ErrorCode -ne $case.expectedErrorCode) {
        throw "Golden case $($case.id) expected $($case.expectedErrorCode), got $($result.ErrorCode)"
    }
    if ($hostCalls.Count -ne [int]$case.expectedHostCalls) {
        throw "Golden case $($case.id) expected host calls $($case.expectedHostCalls), got $($hostCalls.Count)"
    }
    if ($case.id -eq "S-approval-resume") {
        $second = $registry.ExecuteToolAsync($context, [string]$case.toolId, $params).GetAwaiter().GetResult()
        if ($second.Success -or $second.ErrorCode -ne "SAFETY_NEEDS_APPROVAL" -or $hostCalls.Count -ne 1) {
            throw "Golden case $($case.id) approval token was not single-use"
        }
    }

    $results += [pscustomobject]@{
        Id = $case.id
        Status = "PASS"
        ErrorCode = $result.ErrorCode
        HostCalls = $hostCalls.Count
    }
}

$results | Format-Table -AutoSize
Write-Host "Golden L0 passed: $($results.Count) cases"
