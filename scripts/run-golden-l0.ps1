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
$catalog = Get-Content -LiteralPath $catalogPath -Raw -Encoding UTF8 | ConvertFrom-Json
$results = @()

# E6-06: build a real change log out of the case's step observations, exactly the way the host does it
# during a run. The fixture stays data (catalog.json) so a package's contract can be re-read without
# touching this script, and a step the log must NOT record is itself an assertion.
function New-TaskPackageChangeLog {
    param($Case)
    $log = New-Object ShareRibbon.Agent.AgentRunChangeLog
    if ($null -eq $Case.canUndo) {
        throw "Golden case $($Case.id) must state canUndo - the host undo-stack flag decides nativeUndoable"
    }
    $canUndo = [bool]$Case.canUndo
    $index = 0
    foreach ($step in @($Case.steps)) {
        $index += 1
        # ConvertFrom-Json hands back a PSCustomObject; JObject.FromObject on that yields a single
        # "CliXml" property, so the observation has to round-trip through its own JSON text.
        $observation = [Newtonsoft.Json.Linq.JObject]::Parse(($step.observation | ConvertTo-Json -Depth 8))
        $toolId = [string]$step.toolId
        $result = [ShareRibbon.Agent.ToolResult]::Succeed($toolId, "golden step $index", $null, $observation)
        $record = $log.AddFromStep($toolId, $index, "golden step $index", $result, "golden-undo-$index", $canUndo)
        $expectedRecorded = ($null -eq $step.recorded) -or [bool]$step.recorded
        if ($expectedRecorded -and $null -eq $record) {
            throw "Golden case $($Case.id) step ${index}: the change log rejected the fixture - the write-observation contract drifted"
        }
        if ((-not $expectedRecorded) -and $null -ne $record) {
            throw "Golden case $($Case.id) step ${index}: a non-write observation entered the change log, so a read-only run would claim it modified the document"
        }
    }
    return $log
}

foreach ($case in $catalog.cases) {
    if ($case.kind -eq "task_package_outcome") {
        $entryActionId = [string]$case.entryActionId
        $packageId = [string]$case.packageId
        $package = [ShareRibbon.Agent.TaskPackageCatalog]::FindByAction($entryActionId)
        if ($null -eq $package) {
            throw "Golden case $($case.id): entry action $entryActionId no longer resolves to a package, so E6-05 would never verify this task"
        }
        if ($package.Id -ne $packageId) {
            throw "Golden case $($case.id): $entryActionId now belongs to package $($package.Id), not $packageId"
        }

        $changeLog = New-TaskPackageChangeLog $case
        $reasons = @([ShareRibbon.Agent.TaskPackageOutcomeVerifier]::UnmetReasons($package, $changeLog))
        $joined = $reasons -join " | "
        if ($reasons.Count -ne [int]$case.expectedUnmetCount) {
            throw "Golden case $($case.id) expected $($case.expectedUnmetCount) unmet assertion(s), got $($reasons.Count): $joined"
        }
        if ($case.expectedOutcome -eq "satisfied") {
            if ($reasons.Count -gt 0) {
                throw "Golden case $($case.id) ($packageId) should meet its declared success criteria, got: $joined"
            }
        }
        elseif ($case.expectedOutcome -eq "unmet") {
            if ($reasons.Count -eq 0) {
                throw "Golden case $($case.id) ($packageId) must be able to fail - the verifier answered satisfied for a run the package does not accept"
            }
            foreach ($fact in @($case.expectedUnmetFacts)) {
                if (-not $joined.Contains([string]$fact)) {
                    throw "Golden case $($case.id) must name '$fact' in the failure text, got: $joined"
                }
            }
        }
        else {
            throw "Golden case $($case.id) has expectedOutcome '$($case.expectedOutcome)'; only satisfied/unmet are meaningful"
        }

        $results += [pscustomobject]@{
            Id = $case.id
            Status = "PASS"
            ErrorCode = "$packageId/$($case.expectedOutcome)"
            HostCalls = $changeLog.Entries.Count
        }
        continue
    }

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
                $source = Get-Content -LiteralPath (Join-Path $repoRoot $sourceFile) -Raw -Encoding UTF8
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

            $routeSource = Get-Content -LiteralPath (Join-Path $repoRoot "ShareRibbon\Controls\Services\ChatRoutingOrchestrator.vb") -Raw -Encoding UTF8
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

            $rendererSource = Get-Content -LiteralPath (Join-Path $repoRoot "PowerPointAi\Design\PowerPointSceneRenderer.vb") -Raw -Encoding UTF8
            if ($rendererSource -match 'As\s+(FillFormat|LineFormat|ShadowFormat|ColorFormat|Adjustments|PictureFormat|TextFrame2|TextRange2|Font2|ParagraphFormat2)') {
                throw "Golden case $($case.id): renderer reintroduced a strongly typed Office style interface"
            }

            $pptChatSource = Get-Content -LiteralPath (Join-Path $repoRoot "PowerPointAi\ChatControl.vb") -Raw -Encoding UTF8
            if ($pptChatSource -notmatch 'ProfessionalDeckExecutor\.ExecuteAsToolResult\(\s*GetPowerPointEnvelopeParams\(envelope\),\s*preview\)') {
                throw "Golden case $($case.id): CreateSlides preview flag is not propagated"
            }
        }
        elseif ($case.contractType -eq "excel_observer_compat") {
            $excelChatSource = Get-Content -LiteralPath (Join-Path $repoRoot "ExcelAi\ChatControl.vb") -Raw -Encoding UTF8
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

# Coverage is a golden assertion of its own: a catalog whose cases stopped covering the packages on disk
# would still print "passed", and a package rejected at load time disappears from coverage silently.
$loadedCount = [ShareRibbon.Agent.TaskPackageCatalog]::LoadFromRuntimeDirectories()
$loadErrors = @([ShareRibbon.Agent.TaskPackageCatalog]::ValidationErrors)
if ($loadErrors.Count -gt 0) {
    throw "task package(s) rejected while loading: $($loadErrors -join ' | ') - a rejected package cannot be covered by a golden case"
}
$loadedPackages = @([ShareRibbon.Agent.TaskPackageCatalog]::Packages)
if ($loadedCount -lt 1 -or $loadedPackages.Count -lt 1) {
    throw "no task package loaded - the task_package_outcome cases below would prove nothing"
}
foreach ($package in $loadedPackages) {
    $ownCases = @($catalog.cases | Where-Object { $_.kind -eq "task_package_outcome" -and $_.packageId -eq $package.Id })
    $passCases = @($ownCases | Where-Object { $_.expectedOutcome -eq "satisfied" })
    $failCases = @($ownCases | Where-Object { $_.expectedOutcome -eq "unmet" })
    if ($passCases.Count -lt 1 -or $failCases.Count -lt 1) {
        throw "task package $($package.Id) needs >=1 pass and >=1 fail golden case, got pass=$($passCases.Count) fail=$($failCases.Count)"
    }
    if (@($package.GoldenCases).Count -ne $ownCases.Count) {
        throw "task package $($package.Id) declares $(@($package.GoldenCases).Count) golden case(s) but the catalog holds $($ownCases.Count) for it"
    }
}

$results | Format-Table -AutoSize
Write-Host "Golden L0 passed: $($results.Count) cases"
