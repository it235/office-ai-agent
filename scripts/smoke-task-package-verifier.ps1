param(
    [string]$Configuration = "Debug"
)

$ErrorActionPreference = "Stop"
$script:caseCount = 0

# E6-05 behaviour smoke for the task package verifier.
# check-task-package-contract.ps1 only proves the JSON and the vocabularies agree statically.
# This one loads the compiled ShareRibbon.dll and feeds real AgentRunChangeLog instances to the real
# packages, so a green gate cannot hide a verifier that always answers "satisfied" (or always "unmet").

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$dllDir = Join-Path $repoRoot "ShareRibbon\bin\$Configuration"
$shareRibbonDll = Join-Path $dllDir "ShareRibbon.dll"
if (-not (Test-Path -LiteralPath $shareRibbonDll)) {
    throw "missing $shareRibbonDll - run scripts\build-code-projects.ps1 -Configuration $Configuration first"
}

@(
    (Join-Path $repoRoot "packages\Newtonsoft.Json.13.0.3\lib\net45\Newtonsoft.Json.dll"),
    $shareRibbonDll
) | ForEach-Object {
    [Reflection.Assembly]::LoadFrom($_) | Out-Null
}

# One write the host actually committed, with a native undo unit and both document fingerprints.
$wordWriteJson = '{"kind":"write","changed":true,"summary":"smoke write","undoUnits":1,"before":{"documentTextHash":"hash-before"},"after":{"documentTextHash":"hash-after"}}'
$plainWriteJson = '{"kind":"write","changed":true,"summary":"smoke write","before":{"documentTextHash":"hash-before"},"after":{"documentTextHash":"hash-after"}}'
$noChangeJson = '{"kind":"write","changed":false,"summary":"smoke read","before":{"documentTextHash":"hash-before"},"after":{"documentTextHash":"hash-before"}}'
$rolledBackJson = '{"kind":"write","changed":true,"summary":"smoke write","rollback":{"succeeded":true},"before":{"documentTextHash":"hash-before"},"after":{"documentTextHash":"hash-before"}}'

function New-ChangeLog {
    param(
        [string[]]$JsonList,
        [bool]$CanUndo = $true
    )
    $log = New-Object ShareRibbon.Agent.AgentRunChangeLog
    $index = 0
    foreach ($json in $JsonList) {
        $index += 1
        # A JObject is enumerable: it must be parsed here, not returned from a helper,
        # or PowerShell hands back its child properties and the fixture silently loses "kind".
        $observation = [Newtonsoft.Json.Linq.JObject]::Parse($json)
        $result = [ShareRibbon.Agent.ToolResult]::Succeed("SmokeWrite", "smoke step $index", $null, $observation)
        $record = $log.AddFromStep("SmokeWrite", $index, "smoke step $index", $result, "undo-$index", [bool]$CanUndo)
        if ($null -eq $record) {
            throw "AddFromStep rejected the fixture at index $index - the change log contract drifted"
        }
    }
    return $log
}

function New-EmptyChangeLog {
    return (New-ChangeLog -JsonList @($noChangeJson))
}

function New-SyntheticPackage {
    param(
        [string]$Fact,
        [string]$Op,
        [int]$Value = 0
    )
    $package = New-Object ShareRibbon.Agent.TaskPackage
    $package.Id = "smoke.synthetic"
    $package.Verifier = New-Object ShareRibbon.Agent.TaskPackageVerifier
    $package.Verifier.OnFail = "report"
    $assertion = New-Object ShareRibbon.Agent.TaskPackageAssertion
    $assertion.Fact = $Fact
    $assertion.Op = $Op
    $assertion.Value = $Value
    $assertion.Why = "synthetic case"
    $package.Verifier.Assertions.Add($assertion) | Out-Null
    return $package
}

function Assert-Case {
    param(
        [string]$Name,
        $Package,
        $Log,
        [int]$ExpectedCount,
        [string[]]$MustContain = @()
    )
    $reasons = @([ShareRibbon.Agent.TaskPackageOutcomeVerifier]::UnmetReasons($Package, $Log))
    $joined = $reasons -join " | "
    if ($reasons.Count -ne $ExpectedCount) {
        throw "$Name expected $ExpectedCount unmet assertion(s), got $($reasons.Count): $joined"
    }
    foreach ($needle in $MustContain) {
        if (-not $joined.Contains($needle)) {
            throw "$Name must name '$needle' in the failure text, got: $joined"
        }
    }
    $script:caseCount += 1
    Write-Host "    ok $Name"
}

# --- 1. the packages on disk are the ones being verified ---------------------

$cleanAndFix = [ShareRibbon.Agent.TaskPackageCatalog]::FindByAction("excel.selection.formula")
$insight = [ShareRibbon.Agent.TaskPackageCatalog]::FindByAction("excel.selection.explain")
$reviewPolish = [ShareRibbon.Agent.TaskPackageCatalog]::FindByAction("word.selection.rewrite")
$resolved = @(
    @{ Name = "excel.clean-and-fix"; Package = $cleanAndFix },
    @{ Name = "excel.insight"; Package = $insight },
    @{ Name = "word.review-polish"; Package = $reviewPolish }
)
foreach ($pair in $resolved) {
    if ($null -eq $pair.Package) {
        throw "entry action no longer resolves to package $($pair.Name) - the verifier would never run for it"
    }
    if (@($pair.Package.Verifier.Assertions).Count -lt 2) {
        throw "$($pair.Name) must carry at least 2 assertions, got $(@($pair.Package.Verifier.Assertions).Count)"
    }
}
Write-Host "    ok three packages resolve from their entry actions"

# --- 2. satisfied vs unmet, on the same fixtures ----------------------------

Assert-Case "clean-and-fix with a committed write" $cleanAndFix (New-ChangeLog @($plainWriteJson, $plainWriteJson)) 0
Assert-Case "clean-and-fix with nothing written" $cleanAndFix (New-EmptyChangeLog) 2 @("writtenCount", "changedCount")
Assert-Case "clean-and-fix when everything rolled back" $cleanAndFix (New-ChangeLog @($rolledBackJson)) 1 @("writtenCount")
Assert-Case "insight stays read-only" $insight (New-EmptyChangeLog) 0
Assert-Case "insight that wrote into the sheet" $insight (New-ChangeLog @($plainWriteJson)) 2 @("writtenCount", "hasChanges")
Assert-Case "review-polish with native undo on every write" $reviewPolish (New-ChangeLog @($wordWriteJson, $wordWriteJson)) 0
Assert-Case "review-polish without undo evidence" $reviewPolish (New-ChangeLog @($plainWriteJson, $plainWriteJson)) 1 @("nativeUndoable")
Assert-Case "review-polish with nothing written" $reviewPolish (New-EmptyChangeLog) 2 @("writtenCount", "nativeUndoable")
Assert-Case "review-polish without a host undo stack" $reviewPolish (New-ChangeLog -JsonList @($wordWriteJson) -CanUndo $false) 1 @("nativeUndoable")

# --- 3. fail closed: a broken assertion never reads as satisfied -------------

$satisfiedLog = New-ChangeLog @($wordWriteJson)
Assert-Case "unknown fact name" (New-SyntheticPackage "notAFact" "gte" 1) $satisfiedLog 1 @("notAFact")
Assert-Case "boolean op on a number" (New-SyntheticPackage "writtenCount" "isTrue") $satisfiedLog 1 @("writtenCount")
Assert-Case "numeric op on a boolean" (New-SyntheticPackage "hasChanges" "gte" 1) $satisfiedLog 1 @("hasChanges")
Assert-Case "unknown operator" (New-SyntheticPackage "writtenCount" "contains" 1) $satisfiedLog 1 @("contains")
Assert-Case "empty fact name" (New-SyntheticPackage "" "gte" 1) $satisfiedLog 1
Assert-Case "no change log at all" $cleanAndFix $null 2 @("writtenCount", "changedCount")

if ([ShareRibbon.Agent.TaskPackageOutcomeVerifier]::OnFail($cleanAndFix) -ne "repair") {
    throw "verifier must read the declared onFail from the package, got '$([ShareRibbon.Agent.TaskPackageOutcomeVerifier]::OnFail($cleanAndFix))'"
}
if ([ShareRibbon.Agent.TaskPackageOutcomeVerifier]::OnFail($null) -ne "") {
    throw "OnFail of a missing package must be empty, not an invented strategy"
}
Write-Host "    ok onFail still comes from the package"

# --- 4. the runtime reader and the gate vocabulary are one list --------------

$facts = @([ShareRibbon.Agent.TaskPackageCatalog]::VerifiableFacts)
foreach ($fact in $facts) {
    $factValue = $null
    if (-not $satisfiedLog.TryFactValue($fact, [ref]$factValue)) {
        throw "fact '$fact' is in the vocabulary but the runtime reader rejects it"
    }
}
$wrongCaseValue = $null
if ($satisfiedLog.TryFactValue("WrittenCount", [ref]$wrongCaseValue)) {
    throw "TryFactValue must speak the camelCase fact name, not the PascalCase property name"
}
$emptyLog = New-Object ShareRibbon.Agent.AgentRunChangeLog
$writtenValue = $null
if (-not $emptyLog.TryFactValue("writtenCount", [ref]$writtenValue)) {
    throw "TryFactValue could not read writtenCount"
}
if ([int]$writtenValue -ne 0) {
    throw "a fresh change log must report writtenCount 0, got $writtenValue"
}
Write-Host "    ok all $($facts.Count) facts are readable at runtime, and only under the fact name"

Write-Host "PASS task package verifier: 3 real packages, $script:caseCount assertion cases, $($facts.Count) facts readable at runtime"
