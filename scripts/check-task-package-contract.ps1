$ErrorActionPreference = 'Stop'

# E6-03 task package contract gate.
#
# A task package may own only the fields nobody else owns (preview policy, verifier, success
# criteria, deliverable, explanation, recovery, golden-case references, existing implementation).
# Everything else is a reference to an id that already lives elsewhere, and this gate is what makes
# those references carry weight: package -> entry action -> capability -> skill -> golden case ->
# reuse path must all resolve on disk. A broken reference fails the build instead of letting the
# catalog become a fifth drifting list (plan section 9, items 20 and 22).
#
# The schema vocabulary is not duplicated here. It is read back out of TaskPackageCatalog.vb and the
# assertion facts out of AgentRunChangeLog's own properties, so the code stays the single owner and
# this script only proves the JSON on disk agrees with it.
#
# Non-vacuity is part of the contract: the script mutates copies of real packages in memory and
# requires the validator to catch every mutation, and requires the clean packages to pass first.
# A gate that fails everything is as useless as a gate that passes everything.

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')

function Read-Source([string]$relativePath) {
    $path = Join-Path $repoRoot $relativePath
    if (-not (Test-Path -LiteralPath $path)) {
        throw "Expected file is missing: $relativePath"
    }
    return Get-Content -LiteralPath $path -Raw -Encoding UTF8
}

function Get-Vocabulary([string]$loaderSource, [string]$name) {
    $pattern = "Allowed${name}\s+As\s+IReadOnlyCollection\(Of String\)\s*=\s*ToSet\(([^)]*)\)"
    $match = [regex]::Match($loaderSource, $pattern)
    if (-not $match.Success) {
        throw "Loader vocabulary '$name' is no longer a ToSet(...) list in TaskPackageCatalog.vb"
    }
    $values = @([regex]::Matches($match.Groups[1].Value, '"([^"]+)"') | ForEach-Object { $_.Groups[1].Value })
    if ($values.Count -eq 0) {
        throw "Loader vocabulary '$name' is empty - the gate would pass anything"
    }
    return $values
}

function Get-Camel([string]$name) {
    if ([string]::IsNullOrEmpty($name)) { return $name }
    return ([string]$name[0]).ToLowerInvariant() + $name.Substring(1)
}

# --- vocabulary and the referenced inventories, each read from its current owner ---

$loaderSource = Read-Source 'ShareRibbon\Agent\TaskPackageCatalog.vb'
$parserSource = Read-Source 'ShareRibbon\Agent\TaskPackageCatalog.Parsing.vb'
$modelsSource = Read-Source 'ShareRibbon\Agent\TaskPackageModels.vb'
$changeLogSource = Read-Source 'ShareRibbon\Agent\Execution\AgentChangeLog.vb'
$catalogSource = Read-Source 'ShareRibbon\Agent\QuickActionCatalog.vb'
$shareProj = Read-Source 'ShareRibbon\ShareRibbon.vbproj'

$ctx = [pscustomobject]@{
    TopLevelKeys       = Get-Vocabulary $loaderSource 'TopLevelKeys'
    PreviewKeys        = Get-Vocabulary $loaderSource 'PreviewKeys'
    VerifierKeys       = Get-Vocabulary $loaderSource 'VerifierKeys'
    AssertionKeys      = Get-Vocabulary $loaderSource 'AssertionKeys'
    ExplanationKeys    = Get-Vocabulary $loaderSource 'ExplanationKeys'
    RecoveryKeys       = Get-Vocabulary $loaderSource 'RecoveryKeys'
    ImplementationKeys = Get-Vocabulary $loaderSource 'ImplementationKeys'
    PreviewModes       = Get-Vocabulary $loaderSource 'PreviewModes'
    OnFailStrategies   = Get-Vocabulary $loaderSource 'OnFailStrategies'
    RecoveryStrategies = Get-Vocabulary $loaderSource 'RecoveryStrategies'
    Ops                = Get-Vocabulary $loaderSource 'Ops'
    Applications       = Get-Vocabulary $loaderSource 'Applications'
    SchemaVersion      = 0
    NumericFacts       = @()
    BooleanFacts       = @()
    Actions            = @{}
    Capabilities       = @{}
    Skills             = @()
    GoldenCases        = @()
    GoldenOutcomeCases = @()
}

$versionMatch = [regex]::Match($loaderSource, 'AllowedSchemaVersion\s+As\s+Integer\s*=\s*(\d+)')
if (-not $versionMatch.Success) { throw 'AllowedSchemaVersion constant not found in TaskPackageCatalog.vb' }
$ctx.SchemaVersion = [int]$versionMatch.Groups[1].Value

# Facts come from AgentRunChangeLog only; AgentRunUndoScope is a different contract and must not
# widen the vocabulary behind the loader's back.
$logStart = $changeLogSource.IndexOf('Public Class AgentRunChangeLog')
$logEnd = $changeLogSource.IndexOf('Public Class AgentRunUndoScope')
if ($logStart -lt 0 -or $logEnd -le $logStart) {
    throw 'Cannot locate the AgentRunChangeLog class block - the fact vocabulary has no owner to read'
}
$logBlock = $changeLogSource.Substring($logStart, $logEnd - $logStart)
foreach ($m in [regex]::Matches($logBlock, 'Public ReadOnly Property (\w+) As (Integer|Long|Boolean)')) {
    $name = Get-Camel $m.Groups[1].Value
    if ($m.Groups[2].Value -eq 'Boolean') { $ctx.BooleanFacts += $name } else { $ctx.NumericFacts += $name }
}
if ($ctx.NumericFacts.Count -lt 3 -or $ctx.BooleanFacts.Count -lt 2) {
    throw "AgentRunChangeLog exposes too few facts (numeric=$($ctx.NumericFacts.Count), boolean=$($ctx.BooleanFacts.Count))"
}

# Entry actions are read per host block, so a Word action cannot vouch for an Excel package.
foreach ($hostBlock in @(
        [pscustomobject]@{ App = 'Word'; Field = '_wordActions' },
        [pscustomobject]@{ App = 'Excel'; Field = '_excelActions' },
        [pscustomobject]@{ App = 'PowerPoint'; Field = '_powerPointActions' })) {
    $start = $catalogSource.IndexOf("ReadOnly $($hostBlock.Field) ")
    if ($start -lt 0) { throw "QuickActionCatalog lost its $($hostBlock.Field) block" }
    # Any access modifier ends the block: the E6-01 ribbon utterances are Public Shared ReadOnly and
    # must not be absorbed into the PowerPoint palette universe, or a package could vouch for a
    # ribbon button that the palette never shows.
    $end = $catalogSource.IndexOf('Shared ReadOnly', $start + 1)
    if ($end -lt 0) { $end = $catalogSource.Length }
    $block = $catalogSource.Substring($start, $end - $start)
    $ctx.Actions[$hostBlock.App] = @([regex]::Matches($block, '\.Id\s*=\s*"([^"]+)"') | ForEach-Object { $_.Groups[1].Value })
    if ($ctx.Actions[$hostBlock.App].Count -eq 0) {
        throw "QuickActionCatalog $($hostBlock.Field) declares no action ids"
    }
}

# Capability registries are per host. Where a host has none (open decision #36) the package must
# declare no capabilityIds, so an empty registry is the honest answer rather than a grep miss.
$capabilityOwners = @{ Word = 'WordAi\Services\WordCapabilityRegistry.vb' }
foreach ($app in $capabilityOwners.Keys) {
    $source = Read-Source $capabilityOwners[$app]
    $ctx.Capabilities[$app] = @([regex]::Matches($source, '\.Id\s*=\s*"([a-z][\w.-]+)"') | ForEach-Object { $_.Groups[1].Value })
    if ($ctx.Capabilities[$app].Count -eq 0) { throw "$($capabilityOwners[$app]) declares no capability ids" }
}

$skillRoot = Join-Path $repoRoot 'ShareRibbon\Skills'
$ctx.Skills = @(Get-ChildItem -LiteralPath $skillRoot -Directory |
    Where-Object { Test-Path -LiteralPath (Join-Path $_.FullName 'SKILL.md') } |
    ForEach-Object { $_.Name })
if ($ctx.Skills.Count -lt 3) { throw "only $($ctx.Skills.Count) directory skills found under ShareRibbon\Skills" }

$goldenCatalog = Join-Path $repoRoot 'tests\golden\l0\catalog.json'
if (Test-Path -LiteralPath $goldenCatalog) {
    $golden = Get-Content -LiteralPath $goldenCatalog -Raw -Encoding UTF8 | ConvertFrom-Json
    $ctx.GoldenCases = @($golden.cases | ForEach-Object { $_.id })
    # E6-06: the package-outcome cases are the ones that carry a contract verdict, so the gate reads them
    # as data (which package, which polarity, which fact) instead of trusting the runner to notice.
    $ctx.GoldenOutcomeCases = @($golden.cases | Where-Object { $_.kind -eq 'task_package_outcome' })
}

# --- the validator: one function, so the probes are judged by exactly what judges the files ---

function Add-Violation($results, [string]$message) {
    $results.Add($message) | Out-Null
}

function Test-KeyWhitelist($node, [string[]]$allowed, [string]$section, $results) {
    if ($node -eq $null) { return }
    foreach ($key in @($node.PSObject.Properties.Name)) {
        if ($allowed -notcontains $key) {
            Add-Violation $results "$section has field '$key' that TaskPackageCatalog does not accept"
        }
    }
}

function Test-Contract($pkg, [string]$fileName, $ctx) {
    # A List, not an array: an empty array returned from a PS function collapses to $null and then
    # reads back as one empty violation, which is exactly how "everything is dirty" masquerades.
    $results = New-Object 'System.Collections.Generic.List[string]'
    $allFacts = @($ctx.NumericFacts + $ctx.BooleanFacts)

    if ($pkg.schemaVersion -ne $ctx.SchemaVersion) {
        Add-Violation $results "schemaVersion=$($pkg.schemaVersion), loader only accepts $($ctx.SchemaVersion)"
    }
    Test-KeyWhitelist $pkg $ctx.TopLevelKeys 'package' $results

    if ($pkg.id -ne ($fileName -replace '\.json$', '')) {
        Add-Violation $results "file $fileName does not match package id '$($pkg.id)'"
    }
    if ($ctx.Applications -notcontains $pkg.application) {
        Add-Violation $results "application '$($pkg.application)' is outside the vocabulary"
    }
    $prefix = (([string]$pkg.id) -split '\.')[0]
    $expectedPrefix = switch ($pkg.application) {
        'Word' { 'word' }
        'Excel' { 'excel' }
        'PowerPoint' { 'ppt' }
        default { '' }
    }
    if ($expectedPrefix -and $prefix -ne $expectedPrefix -and $prefix -ne 'powerpoint') {
        Add-Violation $results "id prefix '$prefix' contradicts application '$($pkg.application)'"
    }

    $entryIds = @($pkg.entryActions)
    if ($entryIds.Count -lt 1) {
        Add-Violation $results 'entryActions is empty - a package with no entry must not be declared at all (E6-07)'
    }
    $hostActions = @()
    if ($ctx.Actions.ContainsKey($pkg.application)) { $hostActions = $ctx.Actions[$pkg.application] }
    foreach ($action in $entryIds) {
        if ($hostActions -notcontains $action) {
            Add-Violation $results "entry action '$action' is not in QuickActionCatalog's $($pkg.application) list"
        }
    }

    $caps = @($pkg.capabilityIds)
    if ($ctx.Capabilities.ContainsKey($pkg.application)) {
        foreach ($cap in $caps) {
            if ($ctx.Capabilities[$pkg.application] -notcontains $cap) {
                Add-Violation $results "capability '$cap' is not registered for $($pkg.application)"
            }
        }
    } elseif ($caps.Count -gt 0) {
        Add-Violation $results "$($pkg.application) has no capability registry yet (open decision #36), so the package must declare no capabilityIds"
    }

    if ($ctx.Skills -notcontains $pkg.skill) {
        Add-Violation $results "skill '$($pkg.skill)' is not a directory skill with a SKILL.md"
    } else {
        $skillText = Get-Content -LiteralPath (Join-Path $skillRoot ($pkg.skill + '\SKILL.md')) -Raw -Encoding UTF8
        if ($skillText -notmatch ("(?m)^name:\s*" + [regex]::Escape($pkg.skill) + "\s*$")) {
            Add-Violation $results "skill '$($pkg.skill)' front matter name does not match the directory"
        }
        if ($skillText -notmatch '(?m)^allowed-tools:\s*\S') {
            Add-Violation $results "skill '$($pkg.skill)' declares no allowed-tools, so the package has no tool boundary to inherit"
        }
    }

    $policy = $pkg.previewPolicy
    if ($policy -eq $null) {
        Add-Violation $results 'previewPolicy block is missing'
    } else {
        Test-KeyWhitelist $policy $ctx.PreviewKeys 'previewPolicy' $results
        if ($ctx.PreviewModes -notcontains $policy.mode) {
            Add-Violation $results "previewPolicy.mode '$($policy.mode)' is outside the vocabulary"
        }
        if ($policy.mode -eq 'none' -and [bool]$policy.requireBeforeWrite) {
            Add-Violation $results 'previewPolicy.mode=none cannot require a before-write confirmation'
        }
    }

    $verifier = $pkg.verifier
    $assertions = @()
    if ($verifier -eq $null) {
        Add-Violation $results 'verifier block is missing'
    } else {
        Test-KeyWhitelist $verifier $ctx.VerifierKeys 'verifier' $results
        if ($ctx.OnFailStrategies -notcontains $verifier.onFail) {
            Add-Violation $results "verifier.onFail '$($verifier.onFail)' is outside the vocabulary"
        }
        $assertions = @($verifier.assertions)
        if ($assertions.Count -lt 1) {
            Add-Violation $results 'verifier.assertions must hold at least one assertion'
        }
    }
    foreach ($item in $assertions) {
        if ($item -eq $null) { continue }
        Test-KeyWhitelist $item $ctx.AssertionKeys 'assertion' $results
        if ($allFacts -notcontains $item.fact) {
            Add-Violation $results "assertion fact '$($item.fact)' is not a property of AgentRunChangeLog"
            continue
        }
        if ($ctx.Ops -notcontains $item.op) {
            Add-Violation $results "assertion op '$($item.op)' is outside the vocabulary"
        }
        $isBoolean = $ctx.BooleanFacts -contains $item.fact
        if ($isBoolean -and @('isTrue', 'isFalse') -notcontains $item.op) {
            Add-Violation $results "boolean fact '$($item.fact)' only supports isTrue/isFalse"
        }
        if (-not $isBoolean -and @('isTrue', 'isFalse') -contains $item.op) {
            Add-Violation $results "numeric fact '$($item.fact)' cannot use $($item.op)"
        }
        if (-not $isBoolean -and @('eq', 'gte', 'lte') -contains $item.op) {
            if (@($item.PSObject.Properties.Name) -notcontains 'value') {
                Add-Violation $results "assertion on '$($item.fact)' has no integer value"
            } elseif ($item.value -isnot [int] -and $item.value -isnot [long]) {
                Add-Violation $results "assertion value on '$($item.fact)' must be an integer"
            }
        }
    }

    # A read-only package that still demands a write confirmation is the same contradiction the
    # loader rejects, stated from the verifier side instead of the preview side.
    $declaresReadOnly = @($assertions | Where-Object {
        ($_.fact -eq 'writtenCount' -and $_.op -eq 'eq' -and [int]$_.value -eq 0) -or
        ($_.fact -eq 'hasChanges' -and $_.op -eq 'isFalse')
    }).Count -gt 0
    if ($declaresReadOnly -and $policy -and ($policy.mode -ne 'none' -or [bool]$policy.requireBeforeWrite)) {
        Add-Violation $results 'a read-only package (writtenCount eq 0 / hasChanges isFalse) must use previewPolicy.mode=none'
    }

    $explanation = $pkg.explanation
    if ($explanation -eq $null) {
        Add-Violation $results 'explanation block is missing'
    } else {
        Test-KeyWhitelist $explanation $ctx.ExplanationKeys 'explanation' $results
        foreach ($prop in @($explanation.PSObject.Properties.Name)) {
            foreach ($m in [regex]::Matches([string]$explanation.$prop, '\{([A-Za-z][A-Za-z0-9_]*)\}')) {
                if ($allFacts -notcontains $m.Groups[1].Value) {
                    Add-Violation $results "explanation.$prop references unknown fact placeholder '$($m.Groups[1].Value)'"
                }
            }
        }
    }

    $recovery = $pkg.recovery
    if ($recovery -eq $null) {
        Add-Violation $results 'recovery block is missing'
    } else {
        Test-KeyWhitelist $recovery $ctx.RecoveryKeys 'recovery' $results
        if ($ctx.RecoveryStrategies -notcontains $recovery.strategy) {
            Add-Violation $results "recovery.strategy '$($recovery.strategy)' is outside the vocabulary"
        }
    }

    if (@($pkg.successCriteria).Count -lt 1) {
        Add-Violation $results 'successCriteria needs at least one observable condition'
    }
    foreach ($criterion in @($pkg.successCriteria)) {
        if ([string]::IsNullOrWhiteSpace([string]$criterion)) {
            Add-Violation $results 'successCriteria holds an empty line'
        }
    }

    foreach ($case in @($pkg.goldenCases)) {
        if ($ctx.GoldenCases.Count -eq 0) {
            Add-Violation $results 'goldenCases declared but tests/golden/l0/catalog.json has no cases to point at'
            break
        }
        if ($ctx.GoldenCases -notcontains $case) {
            Add-Violation $results "golden case '$case' is not in tests/golden/l0/catalog.json"
            continue
        }
        $owned = @($ctx.GoldenOutcomeCases | Where-Object { $_.id -eq $case })
        if ($owned.Count -ne 1) {
            Add-Violation $results "golden case '$case' is referenced but not exactly one task_package_outcome case (found $($owned.Count))"
            continue
        }
        if ($owned[0].packageId -ne $pkg.id) {
            Add-Violation $results "golden case '$case' belongs to package '$($owned[0].packageId)', not '$($pkg.id)'"
        }
    }

    # E6-06: a contract is only proven when the same assertions have been seen both met and unmet.
    # One polarity alone is as useless as a verifier that always answers the same way.
    $ownCases = @($ctx.GoldenOutcomeCases | Where-Object { $_.packageId -eq $pkg.id })
    if (@($ownCases | Where-Object { $_.expectedOutcome -eq 'satisfied' }).Count -lt 1) {
        Add-Violation $results "'$($pkg.id)' has no satisfied task_package_outcome case - nothing proves the contract can ever be met"
    }
    if (@($ownCases | Where-Object { $_.expectedOutcome -eq 'unmet' }).Count -lt 1) {
        Add-Violation $results "'$($pkg.id)' has no unmet task_package_outcome case - nothing proves its assertions can turn red"
    }
    if ($ownCases.Count -ne @($pkg.goldenCases).Count) {
        Add-Violation $results "'$($pkg.id)' declares $(@($pkg.goldenCases).Count) golden case(s) while the catalog holds $($ownCases.Count) for it"
    }

    $impl = $pkg.existingImplementation
    if ($impl -eq $null) {
        Add-Violation $results 'existingImplementation block is missing'
    } else {
        Test-KeyWhitelist $impl $ctx.ImplementationKeys 'existingImplementation' $results
        $reuse = @($impl.reuse)
        if ($reuse.Count -lt 1) {
            Add-Violation $results 'existingImplementation.reuse must name at least one existing file'
        }
        foreach ($path in $reuse) {
            $diskPath = Join-Path $repoRoot (([string]$path) -replace '/', [IO.Path]::DirectorySeparatorChar)
            if (-not (Test-Path -LiteralPath $diskPath)) {
                Add-Violation $results "existingImplementation.reuse points at a file that is not on disk: $path"
            }
        }
        if ([string]::IsNullOrWhiteSpace([string]$impl.gap)) {
            Add-Violation $results 'existingImplementation.gap must say what this package actually adds'
        }
    }

    foreach ($field in @('title', 'deliverable')) {
        if ([string]::IsNullOrWhiteSpace([string]$pkg.$field)) {
            Add-Violation $results "$field must be non-empty"
        }
    }

    return $results
}

# --- run the real files ---

$packageDir = Join-Path $repoRoot 'ShareRibbon\TaskPackages'
$packages = @(Get-ChildItem -LiteralPath $packageDir -Filter *.json -File)
if ($packages.Count -lt 3) {
    throw "only $($packages.Count) task packages under ShareRibbon\TaskPackages - E6-02 declared 3, so either a file went missing or this floor is stale"
}

function Get-Package([string]$id) {
    foreach ($file in $packages) {
        if ($file.Name -eq ($id + '.json')) {
            return [pscustomobject]@{
                Model = (Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8 | ConvertFrom-Json)
                Name  = $file.Name
            }
        }
    }
    throw "package '$id' is not on disk"
}

$violations = 0
$totalGolden = 0
$totalAssertions = 0
$packageIndex = @{}
foreach ($file in $packages) {
    $pkg = Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
    $problems = @(Test-Contract $pkg $file.Name $ctx)
    $packageIndex[[string]$pkg.id] = $pkg
    $totalGolden += @($pkg.goldenCases).Count
    $totalAssertions += @($pkg.verifier.assertions).Count
    if ($problems.Count -gt 0) {
        $violations += $problems.Count
        Write-Host "  FAIL $($file.Name)"
        foreach ($problem in $problems) { Write-Host "    - $problem" }
    } else {
        Write-Host "  ok   $($file.Name) ($(@($pkg.entryActions).Count) entries, $(@($pkg.verifier.assertions).Count) assertions)"
    }
}
if ($violations -gt 0) {
    throw "$violations task package contract violation(s) across $($packages.Count) packages"
}

# --- E6-06: the golden cases are contract data, so they get judged like the packages themselves -----
# run-golden-l0.ps1 executes them; this half proves they could not have been written to pass anything:
# the case must point at a declared package, start from one of its real entry actions, state which
# polarity it is, name real facts, and describe steps the change log would actually record.

function Test-GoldenCases($cases, $index, $ctx) {
    $results = New-Object 'System.Collections.Generic.List[string]'
    $allFacts = @($ctx.NumericFacts + $ctx.BooleanFacts)
    if (@($cases).Count -lt 1) {
        Add-Violation $results 'tests/golden/l0/catalog.json holds no task_package_outcome case, so no package contract is ever executed'
        return $results
    }
    $ids = @($cases | ForEach-Object { [string]$_.id })
    foreach ($dup in @($ids | Group-Object | Where-Object { $_.Count -gt 1 })) {
        Add-Violation $results "golden case id '$($dup.Name)' is used by $($dup.Count) cases, so a package cannot point at one of them"
    }

    foreach ($case in @($cases)) {
        $label = "golden case '$($case.id)'"
        $pkg = $index[[string]$case.packageId]
        if ($null -eq $pkg) {
            Add-Violation $results "$label names package '$($case.packageId)' that is not declared on disk"
            continue
        }
        if (@($pkg.entryActions) -notcontains [string]$case.entryActionId) {
            Add-Violation $results "$label starts from '$($case.entryActionId)', not one of $($pkg.id)'s entry actions - run end would resolve a different package"
        }

        $outcome = [string]$case.expectedOutcome
        $count = 0
        if (-not [int]::TryParse([string]$case.expectedUnmetCount, [ref]$count)) {
            Add-Violation $results "$label does not state expectedUnmetCount as a number"
            $count = -1
        }
        $facts = @($case.expectedUnmetFacts)
        if ($outcome -eq 'satisfied') {
            if ($count -ne 0) { Add-Violation $results "$label expects a satisfied contract but counts $count unmet assertion(s)" }
            if ($facts.Count -gt 0) { Add-Violation $results "$label is satisfied yet still names unmet facts: $($facts -join ', ')" }
        }
        elseif ($outcome -eq 'unmet') {
            if ($count -lt 1) { Add-Violation $results "$label expects an unmet contract but counts $count unmet assertion(s)" }
            if ($facts.Count -lt 1) { Add-Violation $results "$label is unmet but names no fact, so it cannot say which criterion failed" }
        }
        else {
            Add-Violation $results "$label has expectedOutcome '$outcome'; only satisfied/unmet are meaningful"
        }
        foreach ($fact in $facts) {
            if ($allFacts -notcontains [string]$fact) {
                Add-Violation $results "$label names '$fact' outside the fact vocabulary the change log can report"
            }
        }

        if ($null -eq $case.canUndo) {
            Add-Violation $results "$label must state canUndo - the host undo-stack flag decides nativeUndoable"
        }
        $steps = @($case.steps)
        if ($steps.Count -lt 1) {
            Add-Violation $results "$label has no step fixture, so its ledger is invented rather than observed"
        }
        $stepIndex = 0
        foreach ($step in $steps) {
            $stepIndex += 1
            $stepLabel = "$label step $stepIndex"
            if ([string]::IsNullOrWhiteSpace([string]$step.toolId)) {
                Add-Violation $results "$stepLabel has no toolId"
            }
            $kind = [string]$step.observation.kind
            if ([string]::IsNullOrWhiteSpace($kind)) {
                Add-Violation $results "$stepLabel has an observation without kind - the change log records nothing it cannot classify"
                continue
            }
            $isWrite = ($kind -eq 'write') -or ($kind -eq 'code_execution')
            $recorded = ($null -eq $step.recorded) -or [bool]$step.recorded
            if ($recorded -and -not $isWrite) {
                Add-Violation $results "$stepLabel is marked recorded but kind '$kind' is not a write observation - the change log would drop it and the case would assert on an empty ledger"
            }
            if ((-not $recorded) -and $isWrite) {
                Add-Violation $results "$stepLabel says the ledger must drop it, but kind '$kind' is a write observation the log records"
            }
            if ($recorded -and $null -eq $step.observation.changed) {
                Add-Violation $results "$stepLabel needs a boolean 'changed' - without it the log refuses the record and the fixture silently produces no facts"
            }
        }
    }
    return $results
}

function Copy-GoldenCases($cases) {
    return @($cases | ForEach-Object { $_ | ConvertTo-Json -Depth 12 | ConvertFrom-Json })
}

$goldenViolations = @(Test-GoldenCases $ctx.GoldenOutcomeCases $packageIndex $ctx)
foreach ($v in $goldenViolations) { Write-Host "  FAIL $v" }
if ($goldenViolations.Count -gt 0) {
    throw "$($goldenViolations.Count) golden case violation(s) across $($ctx.GoldenOutcomeCases.Count) case(s)"
}

$goldenProbes = @(
    [pscustomobject]@{ Name = 'golden-ghost-package'; Index = 0; Expect = 'that is not declared on disk';
        Apply = { param($c) $c[0].packageId = 'excel.not-a-package' } },
    [pscustomobject]@{ Name = 'golden-wrong-entry'; Index = 0; Expect = "not one of $($ctx.GoldenOutcomeCases[0].packageId)'s entry actions";
        Apply = { param($c) $c[0].entryActionId = 'excel.selection.explain' } },
    [pscustomobject]@{ Name = 'golden-outcome-outside-vocab'; Index = 0; Expect = 'only satisfied/unmet are meaningful';
        Apply = { param($c) $c[0].expectedOutcome = 'probably' } },
    [pscustomobject]@{ Name = 'golden-satisfied-still-counts'; Index = 0; Expect = 'expects a satisfied contract but counts';
        Apply = { param($c) $c[0].expectedUnmetCount = 2 } },
    [pscustomobject]@{ Name = 'golden-invented-fact'; Index = 1; Expect = 'outside the fact vocabulary';
        Apply = { param($c) $c[1].expectedUnmetFacts = @('magicSpoon') } },
    [pscustomobject]@{ Name = 'golden-unmet-without-facts'; Index = 1; Expect = 'names no fact';
        Apply = { param($c) $c[1].expectedUnmetFacts = @() } },
    [pscustomobject]@{ Name = 'golden-hides-can-undo'; Index = 1; Expect = 'must state canUndo';
        Apply = { param($c) $c[1].PSObject.Properties.Remove('canUndo') } },
    [pscustomobject]@{ Name = 'golden-empty-steps'; Index = 2; Expect = 'has no step fixture';
        Apply = { param($c) $c[2].steps = @() } },
    [pscustomobject]@{ Name = 'golden-read-counted-as-write'; Index = 5; Expect = 'the change log would drop it';
        Apply = { param($c) $c[5].steps[0].recorded = $true } },
    [pscustomobject]@{ Name = 'golden-write-hidden-from-ledger'; Index = 0; Expect = 'the log records';
        Apply = { param($c) $c[0].steps[0] = [pscustomobject]@{ toolId = 'ReplaceText'; recorded = $false;
            observation = $c[0].steps[0].observation } } },
    [pscustomobject]@{ Name = 'golden-duplicate-id'; Index = 0; Expect = 'is used by';
        Apply = { param($c) $c[0].id = $c[1].id } }
)
$goldenCaught = 0
foreach ($probe in $goldenProbes) {
    $copy = Copy-GoldenCases $ctx.GoldenOutcomeCases
    & $probe.Apply $copy
    $found = @(Test-GoldenCases $copy $packageIndex $ctx)
    $text = $found -join ' | '
    if ($found.Count -eq 0) {
        throw "golden probe '$($probe.Name)' slipped through - the golden case rule is vacuous"
    }
    if ($text -notmatch [regex]::Escape($probe.Expect)) {
        throw "golden probe '$($probe.Name)' was caught by the wrong rule; expected '$($probe.Expect)', got: $text"
    }
    $goldenCaught += 1
}

# The golden cases only prove the contract the runtime actually loads. The candidate-directory algorithm
# hits both the source tree and the build output, so a duplicate id ignored without comparing content lets
# an unread stale copy answer for the package - which is how a golden case passes against a dead contract.
if ($loaderSource -notmatch '_sourceTexts\.TryGetValue\(pkg\.Id, previous\)') {
    throw 'TaskPackageCatalog no longer compares a same-id package from another directory - a stale build output can shadow the contract these cases were written against'
}
if ($loaderSource -notmatch '_sourceTexts\(pkg\.Id\) = canonical') {
    throw 'TaskPackageCatalog does not record the contract text it loaded first, so the cross-directory comparison cannot work'
}

# --- registration invariants: a contract nobody loads or copies is decoration ---

foreach ($required in @('Agent\TaskPackageCatalog.vb', 'Agent\TaskPackageCatalog.Parsing.vb', 'Agent\TaskPackageModels.vb')) {
    if ($shareProj -notmatch [regex]::Escape("<Compile Include=""$required""")) {
        throw "ShareRibbon.vbproj does not compile $required - the loader would silently ship without it"
    }
}
if ($shareProj -notmatch ('TaskPackages' + [regex]::Escape('\**\*.json'))) {
    throw 'ShareRibbon.vbproj does not copy TaskPackages\**\*.json to the output directory'
}
if (($loaderSource + "`n" + $parserSource) -notmatch 'QuickActionCatalog\.ActionsFor') {
    throw 'TaskPackageCatalog no longer cross-checks entry actions against QuickActionCatalog at load time'
}
if ($parserSource -notmatch 'Partial Class TaskPackageCatalog') {
    throw 'TaskPackageCatalog.Parsing.vb is no longer part of the same class - the schema would have two owners'
}
if ($modelsSource -notmatch 'TaskPackageCatalog') {
    throw 'TaskPackageModels.vb lost the pointer to where its schema vocabulary lives'
}

# --- E6-01: the three Ribbons must not hold their own Agent wording ------------------------------
# The compile-time guarantee is the helper signature: StartAgentFromRibbonAsync takes a QuickAction,
# so passing a literal request string is no longer valid VB. This section proves the closure around
# that one guarantee: the referenced member exists, belongs to the host that references it and is
# used by exactly one button; an utterance nobody references is deleted instead of kept as
# decoration; ribbon ids keep the <host>.ribbon.<button> shape and never leak into the palette
# universe, because E6-01 moves text, it does not change which rows the command palette shows.

function Get-RibbonNames([string]$catalog) {
    return @([regex]::Matches($catalog, 'Public Shared ReadOnly (\w*Ribbon\w*) As QuickAction') |
        ForEach-Object { $_.Groups[1].Value })
}

function Copy-RibbonState($state) {
    $copy = [ordered]@{ catalog = $state.catalog; sources = [ordered]@{} }
    foreach ($file in $state.sources.Keys) { $copy.sources[$file] = $state.sources[$file] }
    return $copy
}

function Test-RibbonUtterances([string]$catalog, $sources, $prefixes) {
    $results = New-Object 'System.Collections.Generic.List[string]'
    $declared = @(Get-RibbonNames $catalog)
    if ($declared.Count -eq 0) {
        Add-Violation $results 'QuickActionCatalog declares no ribbon utterance - the Ribbons carry their own wording again'
        return $results
    }

    $regionStart = [regex]::Match($catalog, 'Public Shared ReadOnly \w*Ribbon\w* As QuickAction').Index
    $regionEnd = $catalog.IndexOf('Private Shared ReadOnly _wordSamples')
    if ($regionEnd -lt 0) { $regionEnd = $catalog.Length }
    if ($regionEnd -le $regionStart) {
        Add-Violation $results 'the ribbon block in QuickActionCatalog is not between the palette rows and the sample tasks, so the two universes cannot be told apart'
        return $results
    }
    $ribbonBlock = $catalog.Substring($regionStart, $regionEnd - $regionStart)
    $paletteIds = @([regex]::Matches($catalog.Substring(0, $regionStart), '\.Id\s*=\s*"([^"]+)"') | ForEach-Object { $_.Groups[1].Value })
    if ($paletteIds.Count -lt 3) {
        Add-Violation $results 'no palette action id found before the ribbon block - the palette universe was read from the wrong slice'
    }

    $ribbonIds = @([regex]::Matches($ribbonBlock, '\.Id\s*=\s*"([^"]+)"') | ForEach-Object { $_.Groups[1].Value })
    if ($ribbonIds.Count -ne $declared.Count) {
        Add-Violation $results "ribbon block declares $($declared.Count) utterances but carries $($ribbonIds.Count) ids - each needs its own <host>.ribbon.<button> id"
    }
    foreach ($id in $ribbonIds) {
        if ($id -notmatch '^[a-z]+\.ribbon\.[a-z-]+$') {
            Add-Violation $results "ribbon id $id is not of the <host>.ribbon.<button> shape"
        }
        if ($paletteIds -contains $id) {
            Add-Violation $results "ribbon id $id also appears in ActionsFor - E6-01 must not change what the palette shows"
        }
    }

    $usedBy = @{}
    foreach ($file in ($sources.Keys | Sort-Object)) {
        $src = [string]$sources[$file]
        $prefix = [string]$prefixes[$file]
        $sites = 0
        if ($src -notmatch 'StartAgentFromRibbonAsync\(action As Agent\.QuickAction\)') {
            Add-Violation $results "$file no longer takes action As Agent.QuickAction - a literal Agent request would compile again"
        }
        foreach ($m in [regex]::Matches($src, 'StartAgentFromRibbonAsync\s*\(\s*"')) {
            Add-Violation $results "$file passes a literal Agent request string into StartAgentFromRibbonAsync"
        }
        foreach ($m in [regex]::Matches($src, 'StartAgentFromRibbonAsync\(Agent\.QuickActionCatalog\.(\w+)\)')) {
            $name = $m.Groups[1].Value
            $sites++
            if (-not $usedBy.ContainsKey($name)) { $usedBy[$name] = 0 }
            $usedBy[$name] = $usedBy[$name] + 1
            if ($declared -notcontains $name) {
                Add-Violation $results "$file references QuickActionCatalog.$name which is not a declared ribbon utterance"
            }
            if ($name -notmatch ('^' + $prefix + 'Ribbon')) {
                Add-Violation $results "$file references $name which is not that host's ribbon utterance"
            }
        }
        if ($sites -eq 0) {
            Add-Violation $results "$file references no ribbon utterance - its Agent button would keep its own wording"
        }
    }

    foreach ($name in $declared) {
        if (-not $usedBy.ContainsKey($name)) {
            Add-Violation $results "ribbon utterance $name is declared but no Ribbon entry uses it"
        } elseif ($usedBy[$name] -gt 1) {
            Add-Violation $results "ribbon utterance $name is referenced $($usedBy[$name]) times - one button means one utterance"
        }
    }
    return $results
}

$ribbonHosts = [ordered]@{
    'WordAi\Ribbon1.vb'       = 'Word'
    'ExcelAi\Ribbon1.vb'      = 'Excel'
    'PowerPointAi\Ribbon1.vb' = 'Ppt'
}
$ribbonSources = [ordered]@{}
foreach ($file in $ribbonHosts.Keys) {
    $text = Read-Source $file
    if (-not $text) { throw "$file is missing - the ribbon utterance check has nothing to read" }
    $ribbonSources[$file] = $text
}
$ribbonState = [ordered]@{ catalog = $catalogSource; sources = $ribbonSources }

$ribbonViolations = @(Test-RibbonUtterances $ribbonState.catalog $ribbonState.sources $ribbonHosts)
foreach ($v in $ribbonViolations) { Write-Host "  FAIL $v" }
if ($ribbonViolations.Count -gt 0) {
    throw "$($ribbonViolations.Count) ribbon utterance ownership violation(s)"
}

$ribbonNames = @(Get-RibbonNames $catalogSource)
$ribbonSites = 0
foreach ($file in $ribbonSources.Keys) {
    $ribbonSites += [regex]::Matches($ribbonSources[$file], 'StartAgentFromRibbonAsync\(Agent\.QuickActionCatalog\.').Count
}

# Probes: judged by the same function, so a rule that only looks enforced because nobody tried to
# break it cannot pass. Each mutation is applied to a copy; the clean catalog is what was checked above.
foreach ($ribbonProbe in @(
        [pscustomobject]@{ Name = 'ribbon-literal-returns'; Expect = 'passes a literal Agent request string';
            Apply = { param($s) $s.sources['WordAi\Ribbon1.vb'] = $s.sources['WordAi\Ribbon1.vb'].Replace(
                'StartAgentFromRibbonAsync(Agent.QuickActionCatalog.WordRibbonAnalyze)',
                'StartAgentFromRibbonAsync("wording kept inside the ribbon again")') } },
        [pscustomobject]@{ Name = 'ribbon-signature-softened'; Expect = 'no longer takes action As Agent.QuickAction';
            Apply = { param($s) $s.sources['ExcelAi\Ribbon1.vb'] = $s.sources['ExcelAi\Ribbon1.vb'].Replace(
                'StartAgentFromRibbonAsync(action As Agent.QuickAction)',
                'StartAgentFromRibbonAsync(request As String)') } },
        [pscustomobject]@{ Name = 'ribbon-member-undeclared'; Expect = 'not a declared ribbon utterance';
            Apply = { param($s) $s.sources['PowerPointAi\Ribbon1.vb'] = $s.sources['PowerPointAi\Ribbon1.vb'].Replace(
                'Agent.QuickActionCatalog.PptRibbonProofread', 'Agent.QuickActionCatalog.PptRibbonGhost') } },
        [pscustomobject]@{ Name = 'ribbon-utterance-reused'; Expect = 'one button means one utterance';
            Apply = { param($s) $s.sources['PowerPointAi\Ribbon1.vb'] = $s.sources['PowerPointAi\Ribbon1.vb'].Replace(
                'StartAgentFromRibbonAsync(Agent.QuickActionCatalog.PptRibbonProofread)',
                'StartAgentFromRibbonAsync(Agent.QuickActionCatalog.PptRibbonAnalyze)') } },
        [pscustomobject]@{ Name = 'ribbon-utterance-of-another-host'; Expect = "not that host's ribbon utterance";
            Apply = { param($s) $s.sources['WordAi\Ribbon1.vb'] = $s.sources['WordAi\Ribbon1.vb'].Replace(
                'Agent.QuickActionCatalog.WordRibbonAnalyze', 'Agent.QuickActionCatalog.ExcelRibbonAnalyze') } },
        [pscustomobject]@{ Name = 'ribbon-id-shape-drift'; Expect = 'not of the <host>.ribbon.<button> shape';
            Apply = { param($s) $s.catalog = $s.catalog.Replace('.Id = "word.ribbon.analyze"', '.Id = "word.analyze"') } },
        [pscustomobject]@{ Name = 'ribbon-id-in-palette'; Expect = 'also appears in ActionsFor';
            Apply = { param($s) $s.catalog = $s.catalog.Replace('.Id = "word.selection.rewrite"', '.Id = "word.ribbon.analyze"') } } )) {
    $state = Copy-RibbonState $ribbonState
    & $ribbonProbe.Apply $state
    $found = @(Test-RibbonUtterances $state.catalog $state.sources $ribbonHosts)
    if ($found.Count -eq 0) {
        throw "ribbon probe '$($ribbonProbe.Name)' slipped through - the ownership rule is vacuous"
    }
    if (($found -join ' | ') -notmatch [regex]::Escape($ribbonProbe.Expect)) {
        throw "ribbon probe '$($ribbonProbe.Name)' was caught by the wrong rule; expected '$($ribbonProbe.Expect)', got: $($found -join ' | ')"
    }
}

if ($ribbonSites -ne $ribbonNames.Count) {
    throw "ribbon wording is not one-to-one with buttons (utterances=$($ribbonNames.Count), references=$ribbonSites)"
}

# --- E6-05: the entry identity must reach the session, and facts stay owned by the change log ----
# --- E6-05b: the run-end repair route must stay a route, not grow into a second planner -----------
# A package can only be verified if the run knows which entry action started it. The chain is short and
# every link is greppable: panel/host entry -> payload -> StartAgentAsync -> UserTurn -> AgentKernel ->
# AgentSession -> the run-end gate. Dropping any link does not fail a build, it silently disables the
# verifier, so this section is the closure: the identity has a home, both entry paths state it (free text
# states "none" explicitly), and the only reader of change facts is AgentRunChangeLog itself.
# The run-end gate is judged the same way: it re-enters the existing loop, reuses the existing replan
# budget and route, keeps the E1-06 plan-version handoff, and still ends in exactly one terminal.

$identityFiles = [ordered]@{
    'ShareRibbon\Agent\AgentSession.vb'                      = $null
    'ShareRibbon\Agent\Harness\HarnessModels.vb'             = $null
    'ShareRibbon\Agent\Harness\OfficeHarness.vb'             = $null
    'ShareRibbon\Agent\AgentKernel.vb'                       = $null
    'ShareRibbon\Agent\LoopEngine.vb'                        = $null
    'ShareRibbon\Agent\LoopEngine.Coverage.vb'               = $null
    'ShareRibbon\Agent\TaskPackageOutcomeVerifier.vb'        = $null
    'ShareRibbon\Agent\TaskPackageCatalog.vb'                = $loaderSource
    'ShareRibbon\Agent\Execution\AgentChangeLog.vb'          = $changeLogSource
    'ShareRibbon\Controls\Services\AgentKernelService.vb'    = $null
    'ShareRibbon\Controls\BaseChatControl.vb'                = $null
    'ShareRibbon\Resources\js\command-palette.js'            = $null
    'ShareRibbon\ShareRibbon.vbproj'                         = $shareProj
}
$identityState = [ordered]@{}
foreach ($file in $identityFiles.Keys) {
    $text = $identityFiles[$file]
    if ($null -eq $text) { $text = Read-Source $file }
    if (-not $text) { throw "$file is missing - the entry identity check has nothing to read" }
    $identityState[$file] = $text
}

function Copy-IdentityState($state) {
    $copy = [ordered]@{}
    foreach ($file in $state.Keys) { $copy[$file] = $state[$file] }
    return $copy
}

function Test-EntryIdentity($state) {
    $results = New-Object 'System.Collections.Generic.List[string]'

    if ($state['ShareRibbon\Agent\AgentSession.vb'] -notmatch 'Public Property EntryActionId As String') {
        Add-Violation $results 'AgentSession no longer carries the entry identity, so no package can be bound to a run'
    }
    if ($state['ShareRibbon\Agent\Harness\HarnessModels.vb'] -notmatch 'Public Property EntryActionId As String') {
        Add-Violation $results 'UserTurn drops the entry identity between the panel and the kernel'
    }
    if ($state['ShareRibbon\Agent\Harness\OfficeHarness.vb'] -notmatch 'turn\.EntryActionId') {
        Add-Violation $results 'OfficeHarness does not hand the entry identity to AgentKernel'
    }
    $kernel = $state['ShareRibbon\Agent\AgentKernel.vb']
    if ($kernel -notmatch 'Optional entryActionId As String') {
        Add-Violation $results 'AgentKernel.ExecuteAsync cannot receive an entry identity at all'
    }
    if ($kernel -notmatch '_session\.EntryActionId = If\(entryActionId') {
        Add-Violation $results 'AgentKernel does not bind the entry identity onto the session'
    }
    if ($kernel -match 'EntryActionId = .*(userRequest|QuickActionCatalog)') {
        Add-Violation $results 'AgentKernel infers the entry identity from the wording - a guessed package applies the wrong contract'
    }

    $service = $state['ShareRibbon\Controls\Services\AgentKernelService.vb']
    if ($service -notmatch 'AgentEntryActionId = If\(entryActionId') {
        Add-Violation $results 'AgentKernelService no longer registers the entry identity for this turn'
    }
    if ($service -notmatch '\.EntryActionId = AgentEntryActionId') {
        Add-Violation $results 'the UserTurn is built without an entry identity'
    }

    $panel = $state['ShareRibbon\Controls\BaseChatControl.vb']
    if ($panel -notmatch 'Public Sub StartAgentFromHostEntry\(action As Agent\.QuickAction\)') {
        Add-Violation $results 'StartAgentFromHostEntry accepts a bare request again, so the entry id becomes optional'
    }
    if ($panel -notmatch 'payload\("entryActionId"\) = If\(action\.Id') {
        Add-Violation $results 'the host entry path does not pass the id of the action it was given'
    }
    if ($panel -notmatch 'jsonDoc\("entryActionId"\)') {
        Add-Violation $results 'HandleStartAgent no longer reads the entry identity out of the payload'
    }
    if ($panel -notmatch 'HandleStartAgentCore\(question As String[\s\S]{0,160}entryActionId As String\)') {
        Add-Violation $results 'HandleStartAgentCore can be called without stating the entry identity'
    }
    if ($panel -notmatch 'StartAgentAsync\(finalMessageToLLM[\s\S]{0,160}entryActionId\)') {
        Add-Violation $results 'the Agent launch path forgets to pass the entry identity into the service'
    }
    if ($panel -notmatch 'analysis\?\.SelectedSkills,\s*\r?\n\s*""\)') {
        Add-Violation $results 'the chat-input Agent path must state an empty entry id explicitly, not inherit the previous turn'
    }
    if ($panel -notmatch 'AgentKernelSvc\.AgentEntryActionId\)') {
        Add-Violation $results 'refining a plan loses the entry identity, so the package contract stops applying to the same task'
    }

    $palette = $state['ShareRibbon\Resources\js\command-palette.js']
    if ($palette -notmatch "item\.id === '__natural_language__' \? '' :") {
        Add-Violation $results 'the free-text palette row would send its placeholder id as an entry action'
    }
    if ($palette -notmatch 'entryActionId: entryActionId') {
        Add-Violation $results 'the command panel no longer tells the panel which action ran'
    }

    $coverage = $state['ShareRibbon\Agent\LoopEngine.Coverage.vb']
    if ($coverage -notmatch 'ValidateTaskPackageOutcome\(session\)') {
        Add-Violation $results 'the run-end gate no longer asks the task package whether its success criteria were met'
    }
    if ($coverage -notmatch 'FindByAction\(session\.EntryActionId\)') {
        Add-Violation $results 'the package is not resolved from the entry identity, so assertions could bind to any package'
    }
    if ($coverage -notmatch 'TaskPackageOutcomeVerifier\.UnmetReasons\(pkg, session\.ChangeLog\)') {
        Add-Violation $results 'the run-end gate does not read the real change log, so it cannot say unmet'
    }
    if ($coverage -notmatch 'String\.IsNullOrWhiteSpace\(session\.EntryActionId\)') {
        Add-Violation $results 'a free-text run can be bound to a task package again, so "unmet" may report a contract the user never asked for'
    }
    # E6-05b: the run-end repair route must stay a *route*, not grow into a second planner. Every literal here
    # is the seam where the gate hands the unmet assertions to the existing reflect/replan machinery.
    if ($coverage -notmatch 'ReflectAndReplanAsync\(session, observation, systemPrompt\)') {
        Add-Violation $results 'run-end repair no longer reuses the existing reflect/replan route'
    }
    if ($coverage -match 'SendAIRequest\(') {
        Add-Violation $results 'LoopEngine.Coverage.vb opens its own model call, so run-end repair became a second AI route'
    }
    if ($coverage -notmatch 'If decision\?\.Strategy <> "replan" Then') {
        Add-Violation $results 'run-end repair no longer rejects retry/skip, and at run end there is no failed step to retry or skip'
    }
    if ($coverage -notmatch 'ValidateReplannedCoverage\(session, decision\.Plan\)') {
        Add-Violation $results 'a repair plan that loses required coverage is accepted at run end'
    }
    if ($coverage -notmatch 'PackageWantsRunEndRepair\(session As AgentSession\)') {
        Add-Violation $results 'nothing asks the package which failure strategy applies, so onFail stopped being read'
    }
    if ($coverage -notmatch 'declared onFail=rollback but no run-end restore route exists') {
        Add-Violation $results 'the unwired rollback strategy is no longer recorded as unwired, so the gap became invisible instead of reported'
    }

    $loopSource = $state['ShareRibbon\Agent\LoopEngine.vb']
    if ($loopSource -notmatch 'OrElse Not runEndGated') {
        Add-Violation $results 'the loop stops at the last plan step again, so the run-end package gate can never ask for a repair round'
    }
    $successExits = [regex]::Matches($loopSource, 'AgentResult\.SuccessResult').Count
    if ($successExits -ne 1) {
        Add-Violation $results "LoopEngine now has $successExits success terminal(s); the package verdict is only the last gate when there is exactly one"
    }
    $gateBlock = [regex]::Match($loopSource, 'RUN-END CONTRACT GATE[\s\S]{0,1600}?Continue While')
    if (-not $gateBlock.Success) {
        Add-Violation $results 'the run-end contract gate block cannot be located, so its guarantees are unreadable'
    } else {
        $gateRules = @(
            @('IncompletePlanSteps\(session\)\.Count = 0',
              'run-end repair fires while a plan step is still failing, so one problem gets two repair loops'),
            @('PackageWantsRunEndRepair\(session\)',
              'the run-end route no longer reads onFail from the package, so every unmet run repairs itself'),
            @('replanAttempts < MaxReplanAttempts',
              'run-end repair no longer shares the existing replan budget'),
            @('session\.Writes\.BeginNextPlan\(\)',
              'an accepted repair plan no longer advances the plan version, so an applied write can execute twice (E1-06)'),
            @('runEndGated = False',
              'the gate is never re-armed after a repair round, so a second repair cannot happen'),
            @('If repairPlan Is Nothing Then Exit While',
              'a repair round that yields no plan no longer falls back to the unmet report')
        )
        foreach ($gateRule in $gateRules) {
            if ($gateBlock.Value -notmatch $gateRule[0]) {
                Add-Violation $results $gateRule[1]
            }
        }
    }

    $verifier = $state['ShareRibbon\Agent\TaskPackageOutcomeVerifier.vb']
    if ($verifier -notmatch 'TryFactValue\(fact, actual\)') {
        Add-Violation $results 'the verifier no longer reads facts through the change log, so it invented its own'
    }
    foreach ($factProperty in @('WrittenCount', 'ChangedCount', 'HasChanges', 'UndoableUnits', 'RestoreCount', 'NativeUndoable', 'FullyUndoable')) {
        if ($verifier -match [regex]::Escape('.' + $factProperty)) {
            Add-Violation $results "the verifier touches AgentRunChangeLog.$factProperty directly - facts need one owner"
        }
    }
    # The verifier has two satisfied branches (numeric, boolean). A single "Return """ grep would survive
    # one branch being broken off, so pin both: an assertion that really passed must be able to say so.
    foreach ($branch in @('If passed Then Return ""',
                          'If DirectCast(actual, Boolean) = expected Then Return ""')) {
        if ($verifier -notmatch [regex]::Escape($branch)) {
            Add-Violation $results "the verifier lost a satisfied branch ($branch), so a run that met its contract still reports unmet"
        }
    }

    if ($state['ShareRibbon\Agent\TaskPackageCatalog.vb'] -notmatch 'AgentRunChangeLog\.ToFactName\(p\.Name\)') {
        Add-Violation $results 'TaskPackageCatalog keeps its own fact-name rule, so the vocabulary and the runtime reader can drift'
    }
    $logSource = $state['ShareRibbon\Agent\Execution\AgentChangeLog.vb']
    if ($logSource -notmatch 'Public Function TryFactValue') {
        Add-Violation $results 'AgentRunChangeLog no longer exposes the fact reader the verifier depends on'
    }
    $readerBlock = [regex]::Match($logSource, 'Public Function TryFactValue[\s\S]{0,1200}?\r?\n        End Function')
    if (-not $readerBlock.Success) {
        Add-Violation $results 'cannot locate the TryFactValue body - the fact reader shape is unreadable'
    }
    foreach ($guard in @('GetType\(Integer\)', 'GetType\(Long\)', 'GetType\(Boolean\)', 'GetIndexParameters', 'p\.CanRead')) {
        if ($readerBlock.Value -notmatch $guard) {
            Add-Violation $results "TryFactValue no longer restricts facts with $guard - it can read a property the vocabulary does not list"
        }
    }

    if ($state['ShareRibbon\ShareRibbon.vbproj'] -notmatch [regex]::Escape('<Compile Include="Agent\TaskPackageOutcomeVerifier.vb" />')) {
        Add-Violation $results 'TaskPackageOutcomeVerifier.vb is not compiled into ShareRibbon'
    }
    return $results
}

$identityViolations = @(Test-EntryIdentity $identityState)
foreach ($v in $identityViolations) { Write-Host "  FAIL $v" }
if ($identityViolations.Count -gt 0) {
    throw "$($identityViolations.Count) entry identity violation(s)"
}

$identityProbes = @(
    [pscustomobject]@{ Name = 'session-loses-identity'; Expect = 'AgentSession no longer carries the entry identity';
        Apply = { param($s) $s['ShareRibbon\Agent\AgentSession.vb'] = $s['ShareRibbon\Agent\AgentSession.vb'].Replace(
            'Public Property EntryActionId As String = ""', 'Public Property EntryNote As String = ""') } },
    [pscustomobject]@{ Name = 'harness-drops-identity'; Expect = 'OfficeHarness does not hand the entry identity';
        Apply = { param($s) $s['ShareRibbon\Agent\Harness\OfficeHarness.vb'] = $s['ShareRibbon\Agent\Harness\OfficeHarness.vb'].Replace(
            'turn.EntryActionId,', '') } },
    [pscustomobject]@{ Name = 'kernel-guesses-identity'; Expect = 'infers the entry identity from the wording';
        Apply = { param($s) $s['ShareRibbon\Agent\AgentKernel.vb'] = $s['ShareRibbon\Agent\AgentKernel.vb'].Replace(
            '_session.EntryActionId = If(entryActionId, "")', '_session.EntryActionId = QuickActionCatalog.Guess(userRequest)') } },
    [pscustomobject]@{ Name = 'host-entry-back-to-string'; Expect = 'accepts a bare request again';
        Apply = { param($s) $s['ShareRibbon\Controls\BaseChatControl.vb'] = $s['ShareRibbon\Controls\BaseChatControl.vb'].Replace(
            'Public Sub StartAgentFromHostEntry(action As Agent.QuickAction)', 'Public Sub StartAgentFromHostEntry(request As String)') } },
    [pscustomobject]@{ Name = 'chat-path-inherits-identity'; Expect = 'must state an empty entry id explicitly';
        Apply = { param($s) $s['ShareRibbon\Controls\BaseChatControl.vb'] = $s['ShareRibbon\Controls\BaseChatControl.vb'].Replace(
            "analysis?.SelectedSkills,`r`n                             `"`")", 'analysis?.SelectedSkills)') } },
    [pscustomobject]@{ Name = 'palette-drops-entry-id'; Expect = 'no longer tells the panel which action ran';
        Apply = { param($s) $s['ShareRibbon\Resources\js\command-palette.js'] = $s['ShareRibbon\Resources\js\command-palette.js'].Replace(
            'entryActionId: entryActionId', 'shownAs: entryActionId') } },
    [pscustomobject]@{ Name = 'palette-sends-placeholder-id'; Expect = 'free-text palette row would send its placeholder id';
        Apply = { param($s) $s['ShareRibbon\Resources\js\command-palette.js'] = $s['ShareRibbon\Resources\js\command-palette.js'].Replace(
            "item.id === '__natural_language__' ? '' : String(item.id || '')", 'String(item.id || "")') } },
    [pscustomobject]@{ Name = 'run-end-skips-package'; Expect = 'no longer asks the task package';
        Apply = { param($s) $s['ShareRibbon\Agent\LoopEngine.Coverage.vb'] = $s['ShareRibbon\Agent\LoopEngine.Coverage.vb'].Replace(
            'ValidateTaskPackageOutcome(session)', '""') } },
    [pscustomobject]@{ Name = 'verifier-counts-facts-itself'; Expect = 'touches AgentRunChangeLog.WrittenCount directly';
        Apply = { param($s) $s['ShareRibbon\Agent\TaskPackageOutcomeVerifier.vb'] = $s['ShareRibbon\Agent\TaskPackageOutcomeVerifier.vb'].Replace(
            'Return reasons', 'If log IsNot Nothing AndAlso log.WrittenCount > 0 Then reasons.Add("counted by hand")' + [Environment]::NewLine + '            Return reasons') } },
    [pscustomobject]@{ Name = 'verifier-never-satisfied'; Expect = 'lost a satisfied branch';
        Apply = { param($s) $s['ShareRibbon\Agent\TaskPackageOutcomeVerifier.vb'] = $s['ShareRibbon\Agent\TaskPackageOutcomeVerifier.vb'].Replace(
            'If passed Then Return ""', 'If passed Then Return "still unmet"') } },
    [pscustomobject]@{ Name = 'verifier-boolean-never-satisfied'; Expect = 'lost a satisfied branch';
        Apply = { param($s) $s['ShareRibbon\Agent\TaskPackageOutcomeVerifier.vb'] = $s['ShareRibbon\Agent\TaskPackageOutcomeVerifier.vb'].Replace(
            'If DirectCast(actual, Boolean) = expected Then Return ""', 'If False Then Return ""') } },
    [pscustomobject]@{ Name = 'fact-name-rule-split'; Expect = 'keeps its own fact-name rule';
        Apply = { param($s) $s['ShareRibbon\Agent\TaskPackageCatalog.vb'] = $s['ShareRibbon\Agent\TaskPackageCatalog.vb'].Replace(
            'AgentRunChangeLog.ToFactName(p.Name)', 'p.Name.Substring(1)') } },
    [pscustomobject]@{ Name = 'fact-reader-opens-up'; Expect = 'no longer restricts facts with';
        Apply = { param($s) $s['ShareRibbon\Agent\Execution\AgentChangeLog.vb'] = $s['ShareRibbon\Agent\Execution\AgentChangeLog.vb'].Replace(
            'p.PropertyType Is GetType(Long) OrElse', 'p.PropertyType Is GetType(String) OrElse') } },
    [pscustomobject]@{ Name = 'verifier-not-compiled'; Expect = 'is not compiled into ShareRibbon';
        Apply = { param($s) $s['ShareRibbon\ShareRibbon.vbproj'] = $s['ShareRibbon\ShareRibbon.vbproj'].Replace(
            '<Compile Include="Agent\TaskPackageOutcomeVerifier.vb" />', '') } },

    # E6-05b: the run-end repair route. Each probe breaks exactly one guarantee of that route.
    [pscustomobject]@{ Name = 'run-end-gate-unreachable'; Expect = 'can never ask for a repair round';
        Apply = { param($s) $s['ShareRibbon\Agent\LoopEngine.vb'] = $s['ShareRibbon\Agent\LoopEngine.vb'].Replace(
            'OrElse Not runEndGated', '') } },
    [pscustomobject]@{ Name = 'repair-fires-with-failing-steps'; Expect = 'while a plan step is still failing';
        Apply = { param($s) $s['ShareRibbon\Agent\LoopEngine.vb'] = $s['ShareRibbon\Agent\LoopEngine.vb'].Replace(
            'IncompletePlanSteps(session).Count = 0 AndAlso', '') } },
    [pscustomobject]@{ Name = 'repair-ignores-package-strategy'; Expect = 'so every unmet run repairs itself';
        Apply = { param($s) $s['ShareRibbon\Agent\LoopEngine.vb'] = $s['ShareRibbon\Agent\LoopEngine.vb'].Replace(
            'PackageWantsRunEndRepair(session) AndAlso', 'True AndAlso') } },
    [pscustomobject]@{ Name = 'repair-owns-a-second-budget'; Expect = 'shares the existing replan budget';
        Apply = { param($s) $s['ShareRibbon\Agent\LoopEngine.vb'] = $s['ShareRibbon\Agent\LoopEngine.vb'].Replace(
            'replanAttempts < MaxReplanAttempts Then', 'repairRound < 99 Then') } },
    [pscustomobject]@{ Name = 'repair-loses-plan-version'; Expect = 'an applied write can execute twice';
        Apply = { param($s) $s['ShareRibbon\Agent\LoopEngine.vb'] = $s['ShareRibbon\Agent\LoopEngine.vb'].Replace(
            'session.Writes.BeginNextPlan()', '') } },
    [pscustomobject]@{ Name = 'gate-never-re-armed'; Expect = 'never re-armed after a repair round';
        Apply = { param($s) $s['ShareRibbon\Agent\LoopEngine.vb'] = $s['ShareRibbon\Agent\LoopEngine.vb'].Replace(
            'runEndGated = False', '') } },
    [pscustomobject]@{ Name = 'repair-fallback-reuses-old-plan'; Expect = 'no longer falls back to the unmet report';
        Apply = { param($s) $s['ShareRibbon\Agent\LoopEngine.vb'] = $s['ShareRibbon\Agent\LoopEngine.vb'].Replace(
            'If repairPlan Is Nothing Then Exit While', 'If repairPlan Is Nothing Then repairPlan = session.Plan') } },
    [pscustomobject]@{ Name = 'second-success-terminal'; Expect = 'success terminal(s)';
        Apply = { param($s) $s['ShareRibbon\Agent\LoopEngine.vb'] = $s['ShareRibbon\Agent\LoopEngine.vb'].Replace(
            'Return AgentResult.SuccessResult(session.Id, finalMsg)',
            'Return AgentResult.SuccessResult(session.Id, finalMsg) + " repaired"' + [Environment]::NewLine + '                Return AgentResult.SuccessResult(session.Id, finalMsg)') } },
    [pscustomobject]@{ Name = 'repair-builds-its-own-route'; Expect = 'no longer reuses the existing reflect/replan route';
        Apply = { param($s) $s['ShareRibbon\Agent\LoopEngine.Coverage.vb'] = $s['ShareRibbon\Agent\LoopEngine.Coverage.vb'].Replace(
            'Await ReflectAndReplanAsync(session, observation, systemPrompt)', 'Await RepairFromModelAsync(observation)') } },
    [pscustomobject]@{ Name = 'coverage-opens-model-call'; Expect = 'second AI route';
        Apply = { param($s) $s['ShareRibbon\Agent\LoopEngine.Coverage.vb'] = $s['ShareRibbon\Agent\LoopEngine.Coverage.vb'].Replace(
            'Return decision.Plan',
            'Await SendAIRequest(observation, systemPrompt, Nothing)' + [Environment]::NewLine + '            Return decision.Plan') } },
    [pscustomobject]@{ Name = 'repair-accepts-retry-and-skip'; Expect = 'no failed step to retry or skip';
        Apply = { param($s) $s['ShareRibbon\Agent\LoopEngine.Coverage.vb'] = $s['ShareRibbon\Agent\LoopEngine.Coverage.vb'].Replace(
            'If decision?.Strategy <> "replan" Then', 'If decision Is Nothing Then') } },
    [pscustomobject]@{ Name = 'repair-plan-coverage-unchecked'; Expect = 'loses required coverage';
        Apply = { param($s) $s['ShareRibbon\Agent\LoopEngine.Coverage.vb'] = $s['ShareRibbon\Agent\LoopEngine.Coverage.vb'].Replace(
            'ValidateReplannedCoverage(session, decision.Plan)', '""') } },
    [pscustomobject]@{ Name = 'free-text-binds-a-package'; Expect = 'bound to a task package again';
        Apply = { param($s) $s['ShareRibbon\Agent\LoopEngine.Coverage.vb'] = $s['ShareRibbon\Agent\LoopEngine.Coverage.vb'].Replace(
            'String.IsNullOrWhiteSpace(session.EntryActionId)', 'False') } },
    [pscustomobject]@{ Name = 'rollback-gap-goes-silent'; Expect = 'recorded as unwired';
        Apply = { param($s) $s['ShareRibbon\Agent\LoopEngine.Coverage.vb'] = $s['ShareRibbon\Agent\LoopEngine.Coverage.vb'].Replace(
            'declared onFail=rollback but no run-end restore route exists', 'declared onFail=rollback and is handled') } },
    [pscustomobject]@{ Name = 'strategy-helper-deleted'; Expect = 'onFail stopped being read';
        Apply = { param($s) $s['ShareRibbon\Agent\LoopEngine.Coverage.vb'] = $s['ShareRibbon\Agent\LoopEngine.Coverage.vb'].Replace(
            'Private Shared Function PackageWantsRunEndRepair(session As AgentSession) As Boolean',
            'Private Shared Function PackageWantsSomething(session As AgentSession) As Boolean') } }
)
$identityCaught = 0
foreach ($probe in $identityProbes) {
    $copy = Copy-IdentityState $identityState
    & $probe.Apply $copy
    $found = @(Test-EntryIdentity $copy)
    $text = $found -join ' | '
    if ($found.Count -eq 0) {
        throw "identity probe '$($probe.Name)' slipped through - the entry identity rule is vacuous"
    }
    if ($text -notmatch [regex]::Escape($probe.Expect)) {
        throw "identity probe '$($probe.Name)' was caught by the wrong rule; expected '$($probe.Expect)', got: $text"
    }
    $identityCaught++
}

$identitySites = [regex]::Matches($identityState['ShareRibbon\Controls\BaseChatControl.vb'], 'entryActionId').Count

# A verifier that cannot be reached from a run is decoration: the payload key, the session field and the
# run-end call must all exist, and the free-text path must be able to say "no contract".
if ($identitySites -lt 6) {
    throw "the panel threads the entry identity through only $identitySites place(s) - a link in the chain went missing"
}

# --- non-vacuity probes: mutate a copy, the same validator must catch it ---

$writeBaseId = 'excel.clean-and-fix'
$readBaseId = 'excel.insight'
foreach ($baselineId in @($writeBaseId, $readBaseId)) {
    $b = Get-Package $baselineId
    if (@(Test-Contract $b.Model $b.Name $ctx).Count -gt 0) {
        throw "the probe baseline package $baselineId is already dirty, so the probes below prove nothing"
    }
}

function Copy-Package($source) {
    return ($source | ConvertTo-Json -Depth 12 | ConvertFrom-Json)
}

$probes = @(
    [pscustomobject]@{ Name = 'unknown-field'; Base = $writeBaseId; Expect = 'does not accept';
        Apply = { param($p) $p | Add-Member -NotePropertyName 'allowedTools' -NotePropertyValue @('WriteData') } },
    [pscustomobject]@{ Name = 'ghost-entry-action'; Base = $writeBaseId; Expect = 'not in QuickActionCatalog';
        Apply = { param($p) $p.entryActions = @('excel.selection.transcribe') } },
    [pscustomobject]@{ Name = 'invented-fact'; Base = $writeBaseId; Expect = 'not a property of AgentRunChangeLog';
        Apply = { param($p) $p.verifier.assertions[0].fact = 'undoEverything' } },
    [pscustomobject]@{ Name = 'numeric-op-on-boolean'; Base = $writeBaseId; Expect = 'only supports isTrue/isFalse';
        Apply = { param($p) $p.verifier.assertions[0].fact = 'hasChanges'; $p.verifier.assertions[0].op = 'gte'; $p.verifier.assertions[0].value = 1 } },
    [pscustomobject]@{ Name = 'dead-reuse-path'; Base = $writeBaseId; Expect = 'not on disk';
        Apply = { param($p) $p.existingImplementation.reuse = @('ShareRibbon/Agent/TaskPackageCatalogX.vb') } },
    [pscustomobject]@{ Name = 'readonly-demands-preview'; Base = $readBaseId; Expect = 'must use previewPolicy.mode=none';
        Apply = { param($p) $p.previewPolicy.mode = 'summary'; $p.previewPolicy.requireBeforeWrite = $true } },
    [pscustomobject]@{ Name = 'skill-not-on-disk'; Base = $writeBaseId; Expect = 'not a directory skill';
        Apply = { param($p) $p.skill = 'excel-table-agents' } },
    [pscustomobject]@{ Name = 'entry-less-package'; Base = $writeBaseId; Expect = 'no entry must not be declared';
        Apply = { param($p) $p.entryActions = @() } },
    [pscustomobject]@{ Name = 'ungrouped-capability'; Base = $writeBaseId; Expect = 'not registered for Word';
        Apply = { param($p) $p.application = 'Word'; $p.entryActions = @('word.selection.proofread'); $p.capabilityIds = @('word.not-there') } },
    [pscustomobject]@{ Name = 'unknown-golden-case'; Base = $writeBaseId; Expect = 'golden case';
        Apply = { param($p) $p.goldenCases = @('L0-does-not-exist') } },
    [pscustomobject]@{ Name = 'invented-placeholder'; Base = $writeBaseId; Expect = 'unknown fact placeholder';
        Apply = { param($p) $p.explanation.whatChanged = 'wrote {magicSpoon} cells' } },
    [pscustomobject]@{ Name = 'id-filename-mismatch'; Base = $writeBaseId; Expect = 'does not match package id';
        Apply = { param($p) $p.id = 'excel.insightly' } },
    [pscustomobject]@{ Name = 'schema-drift'; Base = $writeBaseId; Expect = 'loader only accepts';
        Apply = { param($p) $p.schemaVersion = $p.schemaVersion + 1 } },
    [pscustomobject]@{ Name = 'mode-outside-vocab'; Base = $writeBaseId; Expect = 'outside the vocabulary';
        Apply = { param($p) $p.previewPolicy.mode = 'telepathy' } }
)

$caught = 0
foreach ($probe in $probes) {
    $b = Get-Package $probe.Base
    $mutated = Copy-Package $b.Model
    & $probe.Apply $mutated
    $problems = @(Test-Contract $mutated $b.Name $ctx)
    $text = $problems -join ' | '
    if ($problems.Count -eq 0) {
        throw "probe '$($probe.Name)' slipped through the validator - the rule is vacuous"
    }
    if ($text -notmatch [regex]::Escape($probe.Expect)) {
        throw "probe '$($probe.Name)' was caught by the wrong rule; expected '$($probe.Expect)', got: $text"
    }
    $caught++
}

# A validator that only ever says "no" also looks green, so pin the sizes the files rely on.
if ($ctx.NumericFacts.Count -lt 3 -or $ctx.Ops.Count -ne 5 -or $ctx.PreviewModes.Count -ne 4) {
    throw "assertion vocabulary drifted (numeric facts=$($ctx.NumericFacts.Count), ops=$($ctx.Ops.Count), preview modes=$($ctx.PreviewModes.Count))"
}

Write-Host ""
Write-Host "PASS task-package contract: $($packages.Count) packages, $totalAssertions assertions, $(@($ctx.TopLevelKeys).Count) schema keys, $((@($ctx.NumericFacts + $ctx.BooleanFacts)).Count) facts, $caught probes turned red on 2 clean baselines, golden references=$totalGolden ($($ctx.GoldenOutcomeCases.Count) outcome cases, $goldenCaught golden probes turned red), ribbon utterances=$($ribbonNames.Count) at $ribbonSites button(s), $identityCaught entry-identity probes turned red at $identitySites panel site(s)"
