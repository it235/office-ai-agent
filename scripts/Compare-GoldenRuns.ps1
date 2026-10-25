param(
    [string]$Baseline,
    [string]$Candidate,
    [string]$OutFile
)

# E11-03: compare two golden-set run records.
# With -Baseline and -Candidate it validates both records, refuses incomparable
# runs, and prints a markdown report. With no arguments it runs the synthetic
# self-check that is wired into scripts\run-code-checks.ps1.

$ErrorActionPreference = "Stop"

$script:StatusValues = @("completed", "needs_approval", "blocked", "needs_clarification", "failed", "not-run")
$script:ResultValues = @("pass", "fail", "not-run")

function Add-RecordError {
    param([System.Collections.ArrayList]$List, [string]$Message)
    [void]$List.Add($Message)
}

function Read-GoldenRunRecord {
    param([Parameter(Mandatory = $true)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        throw "run record not found: $Path"
    }
    $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    return ($raw | ConvertFrom-Json)
}

function Get-JsonArray {
    param($Value)
    if ($null -eq $Value) { return @() }
    return @($Value)
}

function Get-RequiredString {
    param($Object, [string]$Name)
    if ($null -eq $Object) { return "" }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) { return "" }
    return [string]$property.Value
}

function Get-RequiredInt {
    param($Object, [string]$Name)
    $value = $null
    if ($null -ne $Object) {
        $property = $Object.PSObject.Properties[$Name]
        if ($null -ne $property) { $value = $property.Value }
    }
    if ($null -eq $value) { return -1 }
    try { return [int]$value } catch { return -1 }
}

function Test-GoldenRunRecord {
    param($Record)
    $errors = New-Object System.Collections.ArrayList
    if ($null -eq $Record) {
        Add-RecordError $errors "record is empty"
        return ,$errors
    }

    if ((Get-RequiredInt $Record "schemaVersion") -ne 1) {
        Add-RecordError $errors "schemaVersion must be 1"
    }
    if ([string]::IsNullOrWhiteSpace((Get-RequiredString $Record "label"))) {
        Add-RecordError $errors "label is required"
    }
    if ([string]::IsNullOrWhiteSpace((Get-RequiredString $Record "generatedAt"))) {
        Add-RecordError $errors "generatedAt is required"
    }
    $commit = Get-RequiredString $Record "commit"
    if ($commit -notmatch "^[0-9a-fA-F]{7,64}$") {
        Add-RecordError $errors "commit must be a real build commit (7-64 hex chars)"
    }

    $fixtures = $null
    if ($null -ne $Record.PSObject.Properties["fixtures"]) { $fixtures = $Record.fixtures }
    if ($null -eq $fixtures) {
        Add-RecordError $errors "fixtures map is required"
    }
    else {
        $fixtureCount = 0
        foreach ($property in $fixtures.PSObject.Properties) {
            $fixtureCount++
            if ([string]$property.Value -notmatch "^[0-9a-fA-F]{12,64}$") {
                Add-RecordError $errors "fixture '$($property.Name)' has no usable content hash"
            }
        }
        if ($fixtureCount -eq 0) {
            Add-RecordError $errors "fixtures map is empty"
        }
    }

    $cases = @(Get-JsonArray $Record.cases)
    if ($cases.Count -eq 0) {
        Add-RecordError $errors "cases list is empty"
    }
    $seen = @{}
    foreach ($case in $cases) {
        if ($null -eq $case) {
            Add-RecordError $errors "cases contains a null entry"
            continue
        }
        $caseId = Get-RequiredString $case "id"
        if ([string]::IsNullOrWhiteSpace($caseId)) {
            Add-RecordError $errors "case without id"
            continue
        }
        if ($seen.ContainsKey($caseId)) {
            Add-RecordError $errors "duplicate case id '$caseId'"
        }
        $seen[$caseId] = $true

        $result = Get-RequiredString $case "result"
        if ($script:ResultValues -notcontains $result) {
            Add-RecordError $errors "case '$caseId' has unknown result '$result'"
            continue
        }
        $status = Get-RequiredString $case "status"
        if ($script:StatusValues -notcontains $status) {
            Add-RecordError $errors "case '$caseId' has unknown status '$status'"
        }
        $steps = Get-RequiredInt $case "steps"
        $userRework = Get-RequiredInt $case "userRework"
        $repairs = Get-RequiredInt $case "repairs"
        $durationMs = Get-RequiredInt $case "durationMs"
        if ($steps -lt 0) { Add-RecordError $errors "case '$caseId' steps must be >= 0" }
        if ($userRework -lt 0) { Add-RecordError $errors "case '$caseId' userRework must be >= 0" }
        if ($repairs -lt 0) { Add-RecordError $errors "case '$caseId' repairs must be >= 0" }
        if ($durationMs -lt 0) { Add-RecordError $errors "case '$caseId' durationMs must be >= 0" }

        $assertions = @(Get-JsonArray $case.assertions)
        foreach ($assertion in $assertions) {
            if ($null -eq $assertion -or [string]::IsNullOrWhiteSpace((Get-RequiredString $assertion "type"))) {
                Add-RecordError $errors "case '$caseId' has an assertion without type"
                continue
            }
            $passedProperty = $assertion.PSObject.Properties["passed"]
            if ($null -eq $passedProperty -or $passedProperty.Value -isnot [bool]) {
                Add-RecordError $errors "case '$caseId' assertion '$($assertion.type)' needs boolean passed"
            }
        }

        if ($result -eq "not-run") {
            if ($status -ne "not-run") {
                Add-RecordError $errors "case '$caseId' is not-run but its status is '$status'"
            }
            if ($steps -ne 0 -or $durationMs -ne 0 -or $userRework -ne 0 -or $repairs -ne 0) {
                Add-RecordError $errors "case '$caseId' is not-run but carries execution numbers"
            }
            if ($assertions.Count -gt 0) {
                Add-RecordError $errors "case '$caseId' is not-run but has assertion results"
            }
        }
        else {
            if ($status -eq "not-run") {
                Add-RecordError $errors "case '$caseId' has result '$result' but status not-run"
            }
            if ([string]::IsNullOrWhiteSpace((Get-RequiredString $case "runId"))) {
                Add-RecordError $errors "case '$caseId' is executed but has no runId"
            }
        }

        if ($null -ne $case.PSObject.Properties["attempts"]) {
            $attempts = @(Get-JsonArray $case.attempts)
            foreach ($attempt in $attempts) {
                if ($script:ResultValues -notcontains [string]$attempt) {
                    Add-RecordError $errors "case '$caseId' has unknown attempt result '$attempt'"
                }
            }
        }
    }

    return ,$errors
}

function Get-CaseMap {
    param($Record)
    $map = @{}
    foreach ($case in @(Get-JsonArray $Record.cases)) {
        $map[[string]$case.id] = $case
    }
    return $map
}

function Get-RunSummary {
    param($Record)
    $executed = 0
    $passed = 0
    $failed = 0
    $notRun = 0
    $reworkFree = 0
    $repairs = 0
    $durationMs = 0
    foreach ($case in @(Get-JsonArray $Record.cases)) {
        if ([string]$case.result -eq "not-run") {
            $notRun++
            continue
        }
        $executed++
        $durationMs += (Get-RequiredInt $case "durationMs")
        $repairs += (Get-RequiredInt $case "repairs")
        if ([string]$case.result -eq "pass") {
            $passed++
            if ((Get-RequiredInt $case "userRework") -eq 0) { $reworkFree++ }
        }
        else {
            $failed++
        }
    }
    return [pscustomobject]@{
        Executed       = $executed
        Passed         = $passed
        Failed         = $failed
        NotRun         = $notRun
        ReworkFree     = $reworkFree
        Repairs        = $repairs
        PassRate       = $(if ($executed -gt 0) { $passed / $executed } else { $null })
        ReworkFreeRate = $(if ($executed -gt 0) { $reworkFree / $executed } else { $null })
        DurationMs     = $durationMs
    }
}

function Format-Rate {
    param($Rate)
    if ($null -eq $Rate) { return "n/a (not-run)" }
    return ([string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:0.0}%", ($Rate * 100)))
}

function Format-PointDelta {
    param($Before, $After)
    if ($null -eq $Before -or $null -eq $After) { return "n/a" }
    $delta = ($After - $Before) * 100
    if ($delta -ge 0) { return ("+" + [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:0.0}pp", $delta)) }
    return [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:0.0}pp", $delta)
}

function Test-CaseFlaky {
    param($Case)
    if ($null -eq $Case.PSObject.Properties["attempts"]) { return $false }
    $attempts = @(Get-JsonArray $Case.attempts)
    if ($attempts.Count -lt 2) { return $false }
    return (($attempts -contains "pass") -and ($attempts -contains "fail"))
}

function Compare-GoldenRunRecords {
    param($BaselineRecord, $CandidateRecord)
    $reasons = New-Object System.Collections.ArrayList
    $result = [pscustomobject]@{
        Refused          = $false
        RefusalReasons   = $reasons
        BaselineSummary  = $null
        CandidateSummary = $null
        NewFailures      = @()
        Fixed            = @()
        Disappeared      = @()
        Added            = @()
        Flaky            = @()
    }

    if ([int]$BaselineRecord.schemaVersion -ne [int]$CandidateRecord.schemaVersion) {
        Add-RecordError $reasons "schemaVersion differs; records come from different contracts"
    }
    $baselineFixtures = $BaselineRecord.fixtures
    $candidateFixtures = $CandidateRecord.fixtures
    $baselineIds = @($baselineFixtures.PSObject.Properties.Name)
    $candidateIds = @($candidateFixtures.PSObject.Properties.Name)
    foreach ($id in $baselineIds) {
        if ($candidateIds -notcontains $id) {
            Add-RecordError $reasons "fixture '$id' exists in baseline only; the corpus changed"
        }
    }
    foreach ($id in $candidateIds) {
        if ($baselineIds -notcontains $id) {
            Add-RecordError $reasons "fixture '$id' exists in candidate only; the corpus changed"
        }
    }
    foreach ($id in $baselineIds) {
        if ($candidateIds -contains $id) {
            $before = [string]$baselineFixtures.PSObject.Properties[$id].Value
            $after = [string]$candidateFixtures.PSObject.Properties[$id].Value
            if ($before -ne $after) {
                Add-RecordError $reasons "fixture '$id' hash changed ($before -> $after); refusing to compare"
            }
        }
    }

    $baselineSummary = Get-RunSummary $BaselineRecord
    $candidateSummary = Get-RunSummary $CandidateRecord
    $result.BaselineSummary = $baselineSummary
    $result.CandidateSummary = $candidateSummary
    if ($baselineSummary.Executed -eq 0) {
        Add-RecordError $reasons "baseline has no executed case; not-run cannot become a score"
    }
    if ($candidateSummary.Executed -eq 0) {
        Add-RecordError $reasons "candidate has no executed case; not-run cannot become a score"
    }

    if ($reasons.Count -gt 0) {
        $result.Refused = $true
        return $result
    }

    $result.BaselineSummary = $baselineSummary
    $result.CandidateSummary = $candidateSummary
    $baselineMap = Get-CaseMap $BaselineRecord
    $candidateMap = Get-CaseMap $CandidateRecord

    $newFailures = New-Object System.Collections.ArrayList
    $fixed = New-Object System.Collections.ArrayList
    $disappeared = New-Object System.Collections.ArrayList
    $added = New-Object System.Collections.ArrayList
    $flaky = New-Object System.Collections.ArrayList

    foreach ($id in @($baselineMap.Keys | Sort-Object)) {
        if (-not $candidateMap.ContainsKey($id)) {
            [void]$disappeared.Add($id)
            continue
        }
        $before = [string]$baselineMap[$id].result
        $after = [string]$candidateMap[$id].result
        if ($before -eq "pass" -and $after -eq "fail") { [void]$newFailures.Add($id) }
        if ($before -eq "fail" -and $after -eq "pass") { [void]$fixed.Add($id) }
        if ((Test-CaseFlaky $baselineMap[$id]) -or (Test-CaseFlaky $candidateMap[$id])) {
            if ($flaky -notcontains $id) { [void]$flaky.Add($id) }
        }
    }
    foreach ($id in @($candidateMap.Keys | Sort-Object)) {
        if (-not $baselineMap.ContainsKey($id)) {
            [void]$added.Add($id)
            if (Test-CaseFlaky $candidateMap[$id]) { [void]$flaky.Add($id) }
        }
    }

    $result.NewFailures = @($newFailures)
    $result.Fixed = @($fixed)
    $result.Disappeared = @($disappeared)
    $result.Added = @($added)
    $result.Flaky = @($flaky)
    return $result
}

function Format-CaseList {
    param([string[]]$Cases)
    if ($null -eq $Cases -or $Cases.Count -eq 0) { return "- none" }
    $lines = New-Object System.Collections.ArrayList
    foreach ($case in $Cases) { [void]$lines.Add("- " + $case) }
    return ($lines -join "`n")
}

function New-GoldenRunReport {
    param($BaselineRecord, $CandidateRecord, $Comparison)
    $lines = New-Object System.Collections.ArrayList
    [void]$lines.Add("# Golden run comparison")
    [void]$lines.Add("")
    [void]$lines.Add("| field | baseline | candidate |")
    [void]$lines.Add("|---|---|---|")
    [void]$lines.Add("| label | $($BaselineRecord.label) | $($CandidateRecord.label) |")
    [void]$lines.Add("| commit | $($BaselineRecord.commit) | $($CandidateRecord.commit) |")
    [void]$lines.Add("| generatedAt | $($BaselineRecord.generatedAt) | $($CandidateRecord.generatedAt) |")
    [void]$lines.Add("")
    [void]$lines.Add("## Metrics")
    [void]$lines.Add("")
    [void]$lines.Add("| metric | baseline | candidate | delta |")
    [void]$lines.Add("|---|---|---|---|")
    $before = $Comparison.BaselineSummary
    $after = $Comparison.CandidateSummary
    [void]$lines.Add("| executed | $($before.Executed) | $($after.Executed) | |")
    [void]$lines.Add("| pass rate | $(Format-Rate $before.PassRate) | $(Format-Rate $after.PassRate) | $(Format-PointDelta $before.PassRate $after.PassRate) |")
    [void]$lines.Add("| rework-free rate | $(Format-Rate $before.ReworkFreeRate) | $(Format-Rate $after.ReworkFreeRate) | $(Format-PointDelta $before.ReworkFreeRate $after.ReworkFreeRate) |")
    [void]$lines.Add("| internal repairs | $($before.Repairs) | $($after.Repairs) | |")
    [void]$lines.Add("| not-run | $($before.NotRun) | $($after.NotRun) | |")
    [void]$lines.Add("| total duration ms | $($before.DurationMs) | $($after.DurationMs) | |")
    [void]$lines.Add("")
    [void]$lines.Add("## New failures (pass -> fail)")
    [void]$lines.Add("")
    [void]$lines.Add((Format-CaseList $Comparison.NewFailures))
    [void]$lines.Add("")
    [void]$lines.Add("## Fixed (fail -> pass)")
    [void]$lines.Add("")
    [void]$lines.Add((Format-CaseList $Comparison.Fixed))
    [void]$lines.Add("")
    [void]$lines.Add("## Disappeared cases (baseline only)")
    [void]$lines.Add("")
    [void]$lines.Add((Format-CaseList $Comparison.Disappeared))
    [void]$lines.Add("")
    [void]$lines.Add("## Added cases (candidate only)")
    [void]$lines.Add("")
    [void]$lines.Add((Format-CaseList $Comparison.Added))
    [void]$lines.Add("")
    [void]$lines.Add("## Flaky (attempts contain pass and fail)")
    [void]$lines.Add("")
    [void]$lines.Add((Format-CaseList $Comparison.Flaky))
    [void]$lines.Add("")
    [void]$lines.Add("Rework-free means result=pass and userRework=0. Internal replan/repair is reported separately and never counted as user rework.")
    return ($lines -join "`n")
}

function New-SyntheticRun {
    param(
        [string]$Label,
        [string]$Commit,
        [hashtable]$FixtureHashes,
        [object[]]$Cases
    )
    return [pscustomobject]@{
        schemaVersion = 1
        label         = $Label
        commit        = $Commit
        generatedAt   = "2026-10-25T00:00:00Z"
        fixtures      = [pscustomobject]$FixtureHashes
        cases         = @($Cases)
    }
}

function New-SyntheticCase {
    param(
        [string]$Id,
        [string]$Result,
        [int]$UserRework = 0,
        [string[]]$Attempts = @()
    )
    $case = [pscustomobject]@{
        id          = $Id
        result      = $Result
        status      = $(if ($Result -eq "pass") { "completed" } elseif ($Result -eq "fail") { "failed" } else { "not-run" })
        steps       = $(if ($Result -eq "pass") { 4 } else { 3 })
        userRework  = $UserRework
        repairs     = $(if ($Result -eq "pass") { 1 } else { 2 })
        durationMs  = $(if ($Result -eq "pass") { 12000 } else { 18000 })
        runId       = $(if ($Result -eq "not-run") { "" } else { "run_" + $Id })
        assertions  = @()
    }
    if ($Result -eq "not-run") {
        $case.steps = 0
        $case.repairs = 0
        $case.durationMs = 0
    }
    else {
        $case.assertions = @([pscustomobject]@{ type = "run_status"; passed = ($Result -eq "pass") })
    }
    if ($Attempts.Count -gt 0) {
        $case | Add-Member -NotePropertyName "attempts" -NotePropertyValue @($Attempts) -Force
    }
    return $case
}

function Assert-SelfTestEqual {
    param($Actual, $Expected, [string]$Message)
    if ($Actual -ne $Expected) {
        throw "self-test failed: $Message (expected '$Expected', got '$Actual')"
    }
}

function Invoke-SelfTest {
    $hashes = @{ "word/report-zh" = "96b4db150237aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"; "excel/sales-zh" = "f22997502b7fbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb" }
    $baseline = New-SyntheticRun -Label "baseline" -Commit "abc1234" -FixtureHashes $hashes -Cases @(
        (New-SyntheticCase "A-pass" "pass" 0),
        (New-SyntheticCase "B-rework" "pass" 1),
        (New-SyntheticCase "C-fail" "fail" 0),
        (New-SyntheticCase "D-disappear" "pass" 0),
        (New-SyntheticCase "E-flaky" "pass" 0 -Attempts @("pass", "fail")),
        (New-SyntheticCase "F-fixed" "fail" 0)
    )
    $candidate = New-SyntheticRun -Label "candidate" -Commit "def5678" -FixtureHashes $hashes -Cases @(
        (New-SyntheticCase "A-pass" "fail" 0),
        (New-SyntheticCase "B-rework" "pass" 0),
        (New-SyntheticCase "C-fail" "fail" 0),
        (New-SyntheticCase "E-flaky" "pass" 0 -Attempts @("pass", "fail")),
        (New-SyntheticCase "F-fixed" "pass" 0),
        (New-SyntheticCase "G-added" "pass" 0)
    )

    foreach ($record in @($baseline, $candidate)) {
        $errors = Test-GoldenRunRecord $record
        if ($errors.Count -gt 0) { throw "self-test failed: synthetic record rejected: $($errors[0])" }
    }

    $comparison = Compare-GoldenRunRecords -BaselineRecord $baseline -CandidateRecord $candidate
    Assert-SelfTestEqual $comparison.Refused $false "comparable records must not be refused"
    Assert-SelfTestEqual $comparison.BaselineSummary.Passed 4 "baseline pass count"
    Assert-SelfTestEqual $comparison.CandidateSummary.Passed 4 "candidate pass count"
    Assert-SelfTestEqual (Format-Rate $comparison.BaselineSummary.ReworkFreeRate) "50.0%" "baseline rework-free rate"
    Assert-SelfTestEqual (Format-Rate $comparison.CandidateSummary.ReworkFreeRate) "66.7%" "candidate rework-free rate"
    Assert-SelfTestEqual ($comparison.NewFailures -join ",") "A-pass" "new failure"
    Assert-SelfTestEqual ($comparison.Fixed -join ",") "F-fixed" "fixed case"
    Assert-SelfTestEqual ($comparison.Disappeared -join ",") "D-disappear" "disappeared case"
    Assert-SelfTestEqual ($comparison.Added -join ",") "G-added" "added case"
    Assert-SelfTestEqual ($comparison.Flaky -join ",") "E-flaky" "flaky case"

    $report = New-GoldenRunReport -BaselineRecord $baseline -CandidateRecord $candidate -Comparison $comparison
    foreach ($needle in @("## Metrics", "## New failures", "## Disappeared cases", "## Flaky", "A-pass", "E-flaky")) {
        if (-not $report.Contains($needle)) { throw "self-test failed: report is missing '$needle'" }
    }

    $badHash = New-SyntheticRun -Label "bad-hash" -Commit "def5678" -FixtureHashes @{
        "word/report-zh" = "ffffffffffffcccccccccccccccccccccccccccccccccccccccccccccccc"
        "excel/sales-zh" = $hashes["excel/sales-zh"]
    } -Cases @((New-SyntheticCase "A-pass" "pass" 0))
    $refused = Compare-GoldenRunRecords -BaselineRecord $baseline -CandidateRecord $badHash
    Assert-SelfTestEqual $refused.Refused $true "fixture hash mismatch must be refused"
    if (-not (($refused.RefusalReasons -join "`n").Contains("hash changed"))) {
        throw "self-test failed: hash refusal does not say why"
    }

    $notRun = New-SyntheticRun -Label "not-run" -Commit "abc1234" -FixtureHashes $hashes -Cases @(
        (New-SyntheticCase "A-pass" "not-run" 0),
        (New-SyntheticCase "B-rework" "not-run" 0)
    )
    $notRunErrors = Test-GoldenRunRecord $notRun
    Assert-SelfTestEqual $notRunErrors.Count 0 "an honest not-run record is valid"
    $notRunComparison = Compare-GoldenRunRecords -BaselineRecord $baseline -CandidateRecord $notRun
    Assert-SelfTestEqual $notRunComparison.Refused $true "a not-run candidate must not become a score"
    Assert-SelfTestEqual (Format-Rate $notRunComparison.CandidateSummary.ReworkFreeRate) "n/a (not-run)" "not-run rate wording"

    $broken = New-SyntheticRun -Label "broken" -Commit "abc1234" -FixtureHashes $hashes -Cases @(
        (New-SyntheticCase "A-pass" "maybe" 0)
    )
    $brokenErrors = Test-GoldenRunRecord $broken
    if ($brokenErrors.Count -eq 0) { throw "self-test failed: unknown result was accepted" }

    $noRunId = New-SyntheticRun -Label "no-run-id" -Commit "abc1234" -FixtureHashes $hashes -Cases @(
        ([pscustomobject]@{ id = "A"; result = "pass"; status = "completed"; steps = 1; userRework = 0; repairs = 0; durationMs = 1; runId = ""; assertions = @() })
    )
    if ((Test-GoldenRunRecord $noRunId).Count -eq 0) { throw "self-test failed: executed case without runId was accepted" }

    Write-Host "PASS golden run comparison: synthetic records classify new failures, fixes, disappeared, added, flaky; hash mismatch and not-run are refused"
}

if ([string]::IsNullOrWhiteSpace($Baseline) -and [string]::IsNullOrWhiteSpace($Candidate)) {
    Invoke-SelfTest
    exit 0
}
if ([string]::IsNullOrWhiteSpace($Baseline) -or [string]::IsNullOrWhiteSpace($Candidate)) {
    throw "usage: Compare-GoldenRuns.ps1 -Baseline <run.json> -Candidate <run.json> [-OutFile <report.md>]"
}

$baselineRecord = Read-GoldenRunRecord -Path $Baseline
$candidateRecord = Read-GoldenRunRecord -Path $Candidate
$validationErrors = New-Object System.Collections.ArrayList
foreach ($recordError in (Test-GoldenRunRecord $baselineRecord)) {
    Add-RecordError $validationErrors "baseline: $recordError"
}
foreach ($recordError in (Test-GoldenRunRecord $candidateRecord)) {
    Add-RecordError $validationErrors "candidate: $recordError"
}
if ($validationErrors.Count -gt 0) {
    foreach ($recordError in $validationErrors) { Write-Host "FAIL: $recordError" }
    throw "run record validation failed"
}

$comparison = Compare-GoldenRunRecords -BaselineRecord $baselineRecord -CandidateRecord $candidateRecord
if ($comparison.Refused) {
    foreach ($reason in $comparison.RefusalReasons) { Write-Host "REFUSED: $reason" }
    throw "records are not comparable"
}

$reportText = New-GoldenRunReport -BaselineRecord $baselineRecord -CandidateRecord $candidateRecord -Comparison $comparison
if (-not [string]::IsNullOrWhiteSpace($OutFile)) {
    $parent = Split-Path -Parent $OutFile
    if (-not [string]::IsNullOrWhiteSpace($parent) -and -not (Test-Path -LiteralPath $parent)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }
    Set-Content -LiteralPath $OutFile -Value $reportText -Encoding UTF8
    Write-Host "report written: $OutFile"
}
Write-Output $reportText
