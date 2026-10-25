param()

$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$tasksDir = Join-Path $repoRoot "tests\golden\tasks"
$fixturesDir = Join-Path $tasksDir "fixtures"

. (Join-Path $PSScriptRoot "golden\GoldenFixtureBuilder.ps1")

function Get-Optional {
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $null }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }
    return $property.Value
}

function Get-OptionalArray {
    param($Object, [string]$Name)
    $value = Get-Optional $Object $Name
    if ($null -eq $value) { return ,@() }
    return ,@($value)
}

function Add-CorpusError {
    param([System.Collections.ArrayList]$List, [string]$Message)
    [void]$List.Add($Message)
}

function Copy-JsonObject {
    param($Value)
    if ($null -eq $Value) { return $null }
    $json = $Value | ConvertTo-Json -Depth 20
    return (ConvertFrom-Json -InputObject $json)
}

function Find-CaseIndex {
    param($Cases, [string]$Id)
    for ($i = 0; $i -lt @($Cases).Count; $i++) {
        if ([string](Get-Optional $Cases[$i] "id") -eq $Id) { return $i }
    }
    return -1
}

function Test-NoCommittedBinaries {
    param([string[]]$Paths)
    $errors = New-Object System.Collections.ArrayList
    foreach ($path in @($Paths)) {
        if ([string]::IsNullOrWhiteSpace($path)) { continue }
        if ([string]$path -match "\.(docx|xlsx|pptx)$") {
            Add-CorpusError $errors "office binary must not be committed: $path"
        }
    }
    return ,$errors
}

function Find-ForbiddenCaseKeys {
    param($Node)
    $found = New-Object System.Collections.ArrayList
    if ($null -eq $Node) { return ,$found }
    if ($Node -is [System.Management.Automation.PSCustomObject]) {
        foreach ($property in $Node.PSObject.Properties) {
            if ($property.Name -match "^(snapshot|expectedFile|outputFile|goldenFile|expectedDocument)$") {
                [void]$found.Add($property.Name)
            }
            foreach ($nested in (Find-ForbiddenCaseKeys -Node $property.Value)) { [void]$found.Add($nested) }
        }
    }
    elseif ($Node -is [System.Collections.IEnumerable] -and -not ($Node -is [string])) {
        foreach ($item in @($Node)) {
            foreach ($nested in (Find-ForbiddenCaseKeys -Node $item)) { [void]$found.Add($nested) }
        }
    }
    return ,$found
}

function Convert-GoldenColumnIndex {
    param([string]$Column)
    $value = 0
    foreach ($ch in $Column.ToUpperInvariant().ToCharArray()) {
        $value = $value * 26 + ([int][char]$ch - 64)
    }
    return $value
}

function Test-GoldenRangeReference {
    param([string]$Reference, [int]$MaxRow, [int]$MaxColumn)
    $errors = New-Object System.Collections.ArrayList
    $match = [regex]::Match([string]$Reference, "^\`$?([A-Za-z]+)\`$?(\d+)(?::\`$?([A-Za-z]+)\`$?(\d+))?$")
    if (-not $match.Success) {
        Add-CorpusError $errors "range '$Reference' is not A1 or A1:B2"
        return ,$errors
    }
    $column1 = Convert-GoldenColumnIndex $match.Groups[1].Value
    $row1 = [int]$match.Groups[2].Value
    if ($match.Groups[3].Success) {
        $column2 = Convert-GoldenColumnIndex $match.Groups[3].Value
        $row2 = [int]$match.Groups[4].Value
    }
    else {
        $column2 = $column1
        $row2 = $row1
    }
    if ($column1 -gt $column2 -or $row1 -gt $row2) {
        Add-CorpusError $errors "range '$Reference' is inverted"
    }
    if ($column2 -gt $MaxColumn -or $row2 -gt $MaxRow) {
        Add-CorpusError $errors "range '$Reference' exceeds the used area ${MaxRow}x${MaxColumn}"
    }
    return ,$errors
}

function Test-GoldenSelection {
    param($Case, $Model)
    $errors = New-Object System.Collections.ArrayList
    $app = [string](Get-Optional $Case "appType")
    $setup = Get-Optional $Case "setup"
    $selection = Get-Optional $setup "selection"
    if ($null -eq $selection) {
        Add-CorpusError $errors "case has no setup.selection"
        return ,$errors
    }
    $kind = [string](Get-Optional $selection "kind")

    if ($app -eq "Word") {
        $paragraphs = @(Get-Optional $Model "Paragraphs")
        $tables = @(Get-Optional $Model "Tables")
        switch ($kind) {
            "whole-document" { }
            "paragraph_range" {
                $from = [int](Get-Optional $selection "from")
                $to = [int](Get-Optional $selection "to")
                if ($from -lt 1 -or $to -lt $from -or $to -gt $paragraphs.Count) {
                    Add-CorpusError $errors "paragraph_range $from..$to is outside the fixture ($($paragraphs.Count) paragraph(s))"
                }
            }
            "table" {
                $index = [int](Get-Optional $selection "index")
                if ($index -lt 0 -or $index -ge $tables.Count) {
                    Add-CorpusError $errors "table index $index is outside the fixture ($($tables.Count) table(s))"
                }
            }
            default { Add-CorpusError $errors "unknown Word selection kind '$kind'" }
        }
    }
    elseif ($app -eq "Excel") {
        $sheetMap = @{}
        foreach ($sheet in @(Get-Optional $Model "Sheets")) { $sheetMap[[string]$sheet.Name] = $sheet }
        switch ($kind) {
            "whole-workbook" { }
            "range" {
                $sheetName = [string](Get-Optional $selection "sheet")
                $rangeRef = [string](Get-Optional $selection "range")
                if (-not $sheetMap.ContainsKey($sheetName)) {
                    Add-CorpusError $errors "range selection names unknown sheet '$sheetName'"
                }
                elseif ([string]::IsNullOrWhiteSpace($rangeRef)) {
                    Add-CorpusError $errors "range selection has no range"
                }
                else {
                    $maxRow = 0
                    $maxColumn = 0
                    foreach ($cell in @($sheetMap[$sheetName].Cells)) {
                        $cellMatch = [regex]::Match([string]$cell.Ref, "^([A-Za-z]+)(\d+)$")
                        if ($cellMatch.Success) {
                            $row = [int]$cellMatch.Groups[2].Value
                            $column = Convert-GoldenColumnIndex $cellMatch.Groups[1].Value
                            if ($row -gt $maxRow) { $maxRow = $row }
                            if ($column -gt $maxColumn) { $maxColumn = $column }
                        }
                    }
                    foreach ($rangeError in (Test-GoldenRangeReference -Reference $rangeRef -MaxRow $maxRow -MaxColumn $maxColumn)) {
                        Add-CorpusError $errors $rangeError
                    }
                }
            }
            default { Add-CorpusError $errors "unknown Excel selection kind '$kind'" }
        }
    }
    elseif ($app -eq "PowerPoint") {
        $slides = @(Get-Optional $Model "Slides")
        switch ($kind) {
            "whole-deck" { }
            "slide" {
                $index = [int](Get-Optional $selection "index")
                if ($index -lt 1 -or $index -gt $slides.Count) {
                    Add-CorpusError $errors "slide index $index is outside the fixture ($($slides.Count) slide(s))"
                }
            }
            "shape" {
                $slideIndex = [int](Get-Optional $selection "slide")
                $shapeName = [string](Get-Optional $selection "name")
                if ($slideIndex -lt 1 -or $slideIndex -gt $slides.Count) {
                    Add-CorpusError $errors "shape selection slide $slideIndex is outside the fixture ($($slides.Count) slide(s))"
                }
                else {
                    $names = @($slides[$slideIndex - 1].Shapes | ForEach-Object { [string]$_.Name })
                    if ($names -notcontains $shapeName) {
                        Add-CorpusError $errors "shape selection names unknown shape '$shapeName' on slide $slideIndex"
                    }
                }
            }
            default { Add-CorpusError $errors "unknown PowerPoint selection kind '$kind'" }
        }
    }
    else {
        Add-CorpusError $errors "cannot validate selection for appType '$app'"
    }

    return ,$errors
}

function Get-GoldenFixtureText {
    param($Model)
    $parts = New-Object System.Collections.ArrayList
    $kind = [string]$Model.Kind
    if ($kind -eq "word") {
        foreach ($paragraph in @($Model.Paragraphs)) { [void]$parts.Add([string]$paragraph.Text) }
        foreach ($table in @($Model.Tables)) {
            foreach ($row in @($table)) {
                foreach ($cell in @($row)) { [void]$parts.Add([string]$cell) }
            }
        }
    }
    elseif ($kind -eq "excel") {
        foreach ($sheet in @($Model.Sheets)) {
            [void]$parts.Add([string]$sheet.Name)
            foreach ($cell in @($sheet.Cells)) {
                [void]$parts.Add([string]$cell.Value)
                if (-not [string]::IsNullOrWhiteSpace([string]$cell.Formula)) {
                    [void]$parts.Add("=" + [string]$cell.Formula)
                }
            }
        }
    }
    else {
        foreach ($slide in @($Model.Slides)) {
            foreach ($shape in @($slide.Shapes)) {
                foreach ($line in @($shape.Paragraphs)) { [void]$parts.Add([string]$line) }
            }
        }
    }
    return ($parts -join "`n")
}

function Test-GoldenCorpus {
    param(
        [object[]]$Cases,
        [object[]]$Specs,
        [hashtable]$Context
    )

    $errors = New-Object System.Collections.ArrayList
    $vocabulary = $Context.Vocabulary
    $caseList = @($Cases)
    $specList = @($Specs)

    $seenFixtureIds = @{}
    foreach ($spec in $specList) {
        $fixtureId = [string](Get-Optional $spec "id")
        if ([string]::IsNullOrWhiteSpace($fixtureId)) {
            Add-CorpusError $errors "fixture spec has no id"
            continue
        }
        if ($seenFixtureIds.ContainsKey($fixtureId)) {
            Add-CorpusError $errors "duplicate fixture id '$fixtureId'"
        }
        else {
            $seenFixtureIds[$fixtureId] = $true
        }
        if ([string](Get-Optional $spec "dataPolicy") -ne "fictional") {
            Add-CorpusError $errors "fixture '$fixtureId' must declare dataPolicy=fictional"
        }
        if ([string]::IsNullOrWhiteSpace([string](Get-Optional $spec "owner"))) {
            Add-CorpusError $errors "fixture '$fixtureId' has no owner"
        }
        $kind = [string](Get-Optional $spec "kind")
        if (@("word", "excel", "ppt") -notcontains $kind) {
            Add-CorpusError $errors "fixture '$fixtureId' has unknown kind '$kind'"
        }
        if (-not $Context.FixtureText.ContainsKey($fixtureId)) {
            Add-CorpusError $errors "fixture '$fixtureId' was not generated and read back"
            continue
        }
        foreach ($needle in (Get-OptionalArray $spec "expectContains")) {
            $text = [string]$needle
            if ($text.Length -gt 0 -and -not $Context.FixtureText[$fixtureId].Contains($text)) {
                Add-CorpusError $errors "fixture '$fixtureId' read-back is missing expected text '$text'"
            }
        }
        $expectedSlideCount = Get-Optional $spec "expectSlideCount"
        if ($null -ne $expectedSlideCount) {
            $actual = [int]$Context.FixtureSlideCounts[$fixtureId]
            if ($actual -ne [int]$expectedSlideCount) {
                Add-CorpusError $errors "fixture '$fixtureId' has $actual slide(s), spec expects $expectedSlideCount"
            }
        }
    }

    if ($caseList.Count -lt 20) {
        Add-CorpusError $errors "corpus has $($caseList.Count) case(s), at least 20 are required"
    }

    $readyCount = 0
    $gapCount = 0
    $negativeCount = 0
    $nightlyCount = 0
    foreach ($case in $caseList) {
        $readiness = [string](Get-Optional $case "readiness")
        if ($readiness -eq "ready") { $readyCount++ }
        if ($readiness -eq "gap") { $gapCount++ }
        if ((Get-OptionalArray $case "tags") -contains "negative") { $negativeCount++ }
        if ([string](Get-Optional $case "ciTier") -eq "nightly") { $nightlyCount++ }
    }

    foreach ($app in @("Word", "Excel", "PowerPoint")) {
        $count = @($caseList | Where-Object { [string](Get-Optional $_ "appType") -eq $app }).Count
        if ($count -lt 4) {
            Add-CorpusError $errors "appType $app has $count case(s), at least 4 are required"
        }
    }
    if ($negativeCount -lt 5) {
        Add-CorpusError $errors "corpus has $negativeCount negative case(s), at least 5 are required"
    }
    if ($readyCount -lt 15) {
        Add-CorpusError $errors "corpus has $readyCount ready case(s), at least 15 are required"
    }
    if ($gapCount -lt 1) {
        Add-CorpusError $errors "corpus has no gap case; missing owners must stay visible"
    }
    if ($nightlyCount -lt 1) {
        Add-CorpusError $errors "no case is tagged ciTier=nightly; D13 needs a nightly L1 subset"
    }

    $allTags = New-Object System.Collections.ArrayList
    foreach ($case in $caseList) {
        foreach ($tag in (Get-OptionalArray $case "tags")) { [void]$allTags.Add([string]$tag) }
    }
    foreach ($category in @("S-skill-gate", "S-vba-disabled", "destructive", "large-range", "vague")) {
        if ($allTags -notcontains $category) {
            Add-CorpusError $errors "negative coverage is missing the '$category' category"
        }
    }

    $seenCaseIds = @{}
    foreach ($case in $caseList) {
        $caseId = [string](Get-Optional $case "id")
        if ([string]::IsNullOrWhiteSpace($caseId)) {
            Add-CorpusError $errors "case has no id"
            continue
        }
        if ($caseId -notmatch "^[A-Z][A-Za-z0-9-]+$") {
            Add-CorpusError $errors "case id '$caseId' must start uppercase and stay alphanumeric/dash"
        }
        if ($seenCaseIds.ContainsKey($caseId)) {
            Add-CorpusError $errors "duplicate case id '$caseId'"
        }
        else {
            $seenCaseIds[$caseId] = $true
        }

        $app = [string](Get-Optional $case "appType")
        if (@("Word", "Excel", "PowerPoint") -notcontains $app) {
            Add-CorpusError $errors "case '$caseId' has unknown appType '$app'"
            continue
        }

        $layer = (Get-OptionalArray $case "layer")
        if ($layer.Count -eq 0) {
            Add-CorpusError $errors "case '$caseId' has no layer"
        }
        foreach ($entryLayer in $layer) {
            if (@($vocabulary.layers) -notcontains $entryLayer) {
                Add-CorpusError $errors "case '$caseId' uses unknown layer '$entryLayer'"
            }
        }
        if ($layer -notcontains "L1") {
            Add-CorpusError $errors "case '$caseId' must include L1; the corpus is measured on the real host"
        }

        $ciTier = [string](Get-Optional $case "ciTier")
        if (@($vocabulary.ciTiers) -notcontains $ciTier) {
            Add-CorpusError $errors "case '$caseId' uses unknown ciTier '$ciTier'"
        }

        $readiness = [string](Get-Optional $case "readiness")
        if (@($vocabulary.readiness) -notcontains $readiness) {
            Add-CorpusError $errors "case '$caseId' uses unknown readiness '$readiness'"
        }

        if ([string]::IsNullOrWhiteSpace([string](Get-Optional $case "title"))) {
            Add-CorpusError $errors "case '$caseId' has no title"
        }
        if ([string]::IsNullOrWhiteSpace([string](Get-Optional $case "userText"))) {
            Add-CorpusError $errors "case '$caseId' has no userText"
        }
        if ([string]::IsNullOrWhiteSpace([string](Get-Optional $case "owner"))) {
            Add-CorpusError $errors "case '$caseId' has no owner"
        }

        $fixtureId = [string](Get-Optional $case "fixtureId")
        if (-not $Context.FixtureIds.ContainsKey($fixtureId)) {
            Add-CorpusError $errors "case '$caseId' points at unknown fixture '$fixtureId'"
        }
        else {
            if ([string]$Context.FixtureAppTypes[$fixtureId] -ne $app) {
                Add-CorpusError $errors "case '$caseId' appType $app does not match fixture '$fixtureId' appType $($Context.FixtureAppTypes[$fixtureId])"
            }
            foreach ($selectionError in (Test-GoldenSelection -Case $case -Model $Context.FixtureModels[$fixtureId])) {
                Add-CorpusError $errors "case '$caseId': $selectionError"
            }
        }

        $entry = [string](Get-Optional $case "entry")
        $entryActionId = [string](Get-Optional $case "entryActionId")
        $ribbonId = [string](Get-Optional $case "ribbonId")
        if (@($vocabulary.entries) -notcontains $entry) {
            Add-CorpusError $errors "case '$caseId' uses unknown entry '$entry'"
        }
        elseif ($entry -eq "quick-action") {
            if ([string]::IsNullOrWhiteSpace($entryActionId)) {
                Add-CorpusError $errors "case '$caseId' uses a quick-action entry without entryActionId"
            }
            elseif (@($Context.QuickActionIds) -notcontains $entryActionId) {
                Add-CorpusError $errors "case '$caseId' names ghost quick-action '$entryActionId'"
            }
            if (-not [string]::IsNullOrWhiteSpace($ribbonId)) {
                Add-CorpusError $errors "case '$caseId' mixes a quick-action entry with ribbonId"
            }
        }
        elseif ($entry -eq "ribbon") {
            if ([string]::IsNullOrWhiteSpace($ribbonId)) {
                Add-CorpusError $errors "case '$caseId' uses a ribbon entry without ribbonId"
            }
            elseif (@($Context.RibbonIds) -notcontains $ribbonId) {
                Add-CorpusError $errors "case '$caseId' names ghost ribbon '$ribbonId'"
            }
            if (-not [string]::IsNullOrWhiteSpace($entryActionId)) {
                Add-CorpusError $errors "case '$caseId' mixes a ribbon entry with entryActionId"
            }
        }
        elseif ($entry -eq "chat") {
            if (-not [string]::IsNullOrWhiteSpace($entryActionId) -or -not [string]::IsNullOrWhiteSpace($ribbonId)) {
                Add-CorpusError $errors "case '$caseId' declares a chat entry plus a native entry id"
            }
        }

        $expect = Get-Optional $case "expect"
        if ($null -eq $expect) {
            Add-CorpusError $errors "case '$caseId' has no expect block"
            continue
        }
        foreach ($expectKey in @($expect.PSObject.Properties.Name)) {
            if (@($vocabulary.expectKeys) -notcontains $expectKey) {
                Add-CorpusError $errors "case '$caseId' has unknown expect key '$expectKey'"
            }
        }

        $status = [string](Get-Optional $expect "status")
        if (@($vocabulary.statuses) -notcontains $status) {
            Add-CorpusError $errors "case '$caseId' has unknown expect.status '$status'"
        }

        $maxStepsValue = Get-Optional $expect "maxSteps"
        if ($null -eq $maxStepsValue) {
            Add-CorpusError $errors "case '$caseId' has no expect.maxSteps budget"
        }
        else {
            $maxSteps = [int]$maxStepsValue
            if ($maxSteps -lt 1 -or $maxSteps -gt 20) {
                Add-CorpusError $errors "case '$caseId' maxSteps $maxSteps is outside 1..20"
            }
        }
        $maxRepairsValue = Get-Optional $expect "maxRepairs"
        if ($null -eq $maxRepairsValue) {
            Add-CorpusError $errors "case '$caseId' has no expect.maxRepairs budget"
        }
        else {
            $maxRepairs = [int]$maxRepairsValue
            if ($maxRepairs -lt 0 -or $maxRepairs -gt 4) {
                Add-CorpusError $errors "case '$caseId' maxRepairs $maxRepairs is outside 0..4"
            }
        }

        $requiredCodes = (Get-OptionalArray $expect "requiredErrorCodes")
        foreach ($code in $requiredCodes) {
            if (@($Context.ErrorCodes) -notcontains [string]$code) {
                Add-CorpusError $errors "case '$caseId' requires unknown error code '$code'"
            }
        }
        if ($requiredCodes.Count -gt 0 -and $status -eq "completed") {
            Add-CorpusError $errors "case '$caseId' requires error codes but expects a completed run"
        }
        foreach ($code in (Get-OptionalArray $expect "forbiddenErrorCodes")) {
            if (@($Context.ErrorCodes) -notcontains [string]$code) {
                Add-CorpusError $errors "case '$caseId' forbids unknown error code '$code'"
            }
        }

        $document = Get-Optional $expect "document"
        if ($null -eq $document) {
            Add-CorpusError $errors "case '$caseId' has no document assertions"
        }
        else {
            foreach ($documentKey in @($document.PSObject.Properties.Name)) {
                if ($documentKey -ne "assertions") {
                    Add-CorpusError $errors "case '$caseId' has unknown document key '$documentKey'"
                }
            }
            $assertions = (Get-OptionalArray $document "assertions")
            if ($assertions.Count -eq 0) {
                Add-CorpusError $errors "case '$caseId' has no document assertions"
            }
            foreach ($assertion in $assertions) {
                $assertionType = [string](Get-Optional $assertion "type")
                if ([string]::IsNullOrWhiteSpace($assertionType)) {
                    Add-CorpusError $errors "case '$caseId' has an assertion without type"
                    continue
                }
                $definition = $vocabulary.assertions.PSObject.Properties[$assertionType]
                if ($null -eq $definition -or $null -eq $definition.Value) {
                    Add-CorpusError $errors "case '$caseId' uses unknown assertion type '$assertionType'"
                    continue
                }
                $definitionValue = $definition.Value
                if (@($definitionValue.apps) -notcontains $app) {
                    Add-CorpusError $errors "case '$caseId' uses assertion '$assertionType' outside appType $app"
                }
                $allowedParams = @("type") + @($definitionValue.params)
                foreach ($paramName in @($assertion.PSObject.Properties.Name)) {
                    if ($allowedParams -notcontains $paramName) {
                        Add-CorpusError $errors "case '$caseId' assertion '$assertionType' has unknown param '$paramName'"
                    }
                }
                foreach ($requiredParam in @($definitionValue.required)) {
                    if ($null -eq (Get-Optional $assertion $requiredParam)) {
                        Add-CorpusError $errors "case '$caseId' assertion '$assertionType' is missing '$requiredParam'"
                    }
                }
                if ($assertionType -eq "run_status" -and @($vocabulary.statuses) -notcontains [string](Get-Optional $assertion "status")) {
                    Add-CorpusError $errors "case '$caseId' run_status assertion uses an unknown status"
                }
            }
        }

        $trace = Get-Optional $expect "trace"
        if ($null -ne $trace) {
            foreach ($traceProperty in $trace.PSObject.Properties) {
                $traceKey = [string]$traceProperty.Name
                if (-not $Context.TraceKeys.ContainsKey($traceKey)) {
                    Add-CorpusError $errors "case '$caseId' has unknown trace key '$traceKey'"
                    continue
                }
                $traceType = [string]$Context.TraceKeys[$traceKey]
                $traceValue = Get-Optional $trace $traceKey
                if ($traceType -eq "int") {
                    try { [void][int]$traceValue }
                    catch { Add-CorpusError $errors "case '$caseId' trace.$traceKey must be an integer" }
                }
                elseif ($traceType -eq "bool") {
                    if ($traceValue -isnot [bool]) {
                        Add-CorpusError $errors "case '$caseId' trace.$traceKey must be a boolean"
                    }
                }
                elseif ($traceType -eq "string-array") {
                    if (@($traceValue).Count -eq 0) {
                        Add-CorpusError $errors "case '$caseId' trace.$traceKey must be a non-empty array"
                    }
                }
            }
            $minWrite = Get-Optional $trace "minWriteSteps"
            $maxWrite = Get-Optional $trace "maxWriteSteps"
            if ($null -ne $minWrite -and $null -ne $maxWrite -and [int]$minWrite -gt [int]$maxWrite) {
                Add-CorpusError $errors "case '$caseId' trace.minWriteSteps exceeds maxWriteSteps"
            }
            if ($status -ne "completed" -and ($null -eq $maxWrite -or [int]$maxWrite -ne 0)) {
                Add-CorpusError $errors "case '$caseId' expects '$status' but does not cap trace.maxWriteSteps at 0"
            }
        }
        elseif ($status -ne "completed") {
            Add-CorpusError $errors "case '$caseId' expects '$status' without a trace block"
        }

        foreach ($forbiddenKey in (Find-ForbiddenCaseKeys -Node $case)) {
            Add-CorpusError $errors "case '$caseId' carries whole-document snapshot key '$forbiddenKey'"
        }

        $traceability = Get-Optional $case "traceability"
        $objective = [string](Get-Optional $traceability "objective")
        if (@($vocabulary.objectives) -notcontains $objective) {
            Add-CorpusError $errors "case '$caseId' has unknown objective '$objective'"
        }
        $planRow = [string](Get-Optional $traceability "planRow")
        if ([string]::IsNullOrWhiteSpace($planRow) -or -not $Context.PlanText.Contains($planRow)) {
            Add-CorpusError $errors "case '$caseId' traceability points at missing plan row '$planRow'"
        }

        $options = Get-Optional $case "harnessOptions"
        $skillName = [string](Get-Optional $case "skill")
        if ($null -eq $options) {
            Add-CorpusError $errors "case '$caseId' has no harnessOptions"
        }
        else {
            $temperature = Get-Optional $options "temperature"
            if ($null -eq $temperature -or [double]$temperature -ne 0) {
                Add-CorpusError $errors "case '$caseId' must run at temperature 0"
            }
            if ((Get-Optional $options "allowVba") -ne $false) {
                Add-CorpusError $errors "case '$caseId' must set allowVba=false"
            }
            $hint = [string](Get-Optional $options "skillHint")
            if (-not [string]::IsNullOrWhiteSpace($hint) -and -not [string]::IsNullOrWhiteSpace($skillName) -and $hint -ne $skillName) {
                Add-CorpusError $errors "case '$caseId' skillHint '$hint' differs from skill '$skillName'"
            }
        }

        $packageId = [string](Get-Optional $case "packageId")
        if (-not [string]::IsNullOrWhiteSpace($packageId)) {
            if (-not $Context.Packages.ContainsKey($packageId)) {
                Add-CorpusError $errors "case '$caseId' names ghost package '$packageId'"
            }
            else {
                $package = $Context.Packages[$packageId]
                if ($entry -eq "quick-action" -and (@($package.entryActions) -notcontains $entryActionId)) {
                    Add-CorpusError $errors "case '$caseId' entry '$entryActionId' is not declared by package '$packageId'"
                }
                if (-not [string]::IsNullOrWhiteSpace($skillName) -and [string]$package.skill -ne $skillName) {
                    Add-CorpusError $errors "case '$caseId' skill '$skillName' differs from package '$packageId' skill '$($package.skill)'"
                }
            }
        }

        if ($readiness -eq "ready") {
            if ([string]::IsNullOrWhiteSpace($skillName)) {
                Add-CorpusError $errors "ready case '$caseId' has no skill"
            }
            elseif (-not $Context.Skills.ContainsKey($skillName)) {
                Add-CorpusError $errors "case '$caseId' names ghost skill '$skillName'"
            }
            else {
                if ([string]$Context.Skills[$skillName].Name -ne $skillName) {
                    Add-CorpusError $errors "skill '$skillName' front matter name does not match its directory"
                }
                if (@($Context.Skills[$skillName].AllowedTools).Count -eq 0) {
                    Add-CorpusError $errors "skill '$skillName' declares no allowed-tools"
                }
            }
            if ($null -ne (Get-Optional $case "gap")) {
                Add-CorpusError $errors "ready case '$caseId' still carries a gap block"
            }
        }
        elseif ($readiness -eq "gap") {
            if ($entry -ne "chat") {
                Add-CorpusError $errors "gap case '$caseId' must stay on the chat entry; it has no native entry yet"
            }
            if (-not [string]::IsNullOrWhiteSpace($packageId)) {
                Add-CorpusError $errors "gap case '$caseId' must not claim package '$packageId'"
            }
            if (-not [string]::IsNullOrWhiteSpace($skillName)) {
                Add-CorpusError $errors "gap case '$caseId' must not claim an owned skill"
            }
            $gap = Get-Optional $case "gap"
            if ($null -eq $gap) {
                Add-CorpusError $errors "gap case '$caseId' has no gap block"
            }
            else {
                $gapOwner = [string](Get-Optional $gap "owner")
                if ([string]::IsNullOrWhiteSpace($gapOwner) -or -not (Test-Path -LiteralPath (Join-Path $Context.RepoRoot $gapOwner))) {
                    Add-CorpusError $errors "gap case '$caseId' names a missing owner path '$gapOwner'"
                }
                $gapRow = [string](Get-Optional $gap "planRow")
                if ([string]::IsNullOrWhiteSpace($gapRow) -or -not $Context.PlanText.Contains($gapRow)) {
                    Add-CorpusError $errors "gap case '$caseId' names a missing plan row '$gapRow'"
                }
                if ([string]::IsNullOrWhiteSpace([string](Get-Optional $gap "missing"))) {
                    Add-CorpusError $errors "gap case '$caseId' does not say what is missing"
                }
                $probeEntry = [string](Get-Optional $gap "probeEntryId")
                $probePackage = [string](Get-Optional $gap "probePackageId")
                if ([string]::IsNullOrWhiteSpace($probeEntry) -and [string]::IsNullOrWhiteSpace($probePackage)) {
                    Add-CorpusError $errors "gap case '$caseId' has no probe ids; promotion would be invisible"
                }
                if (-not [string]::IsNullOrWhiteSpace($probeEntry) -and
                    (@($Context.QuickActionIds) -contains $probeEntry -or @($Context.RibbonIds) -contains $probeEntry)) {
                    Add-CorpusError $errors "gap case '$caseId' probe entry '$probeEntry' now resolves; promote the case to ready"
                }
                if (-not [string]::IsNullOrWhiteSpace($probePackage) -and $Context.Packages.ContainsKey($probePackage)) {
                    Add-CorpusError $errors "gap case '$caseId' probe package '$probePackage' now resolves; promote the case to ready"
                }
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($skillName) -and $Context.Skills.ContainsKey($skillName)) {
            foreach ($toolGroup in (Get-OptionalArray $expect "mustUseToolsAnyOf")) {
                foreach ($tool in @($toolGroup)) {
                    if (@($Context.Skills[$skillName].AllowedTools) -notcontains [string]$tool) {
                        Add-CorpusError $errors "case '$caseId' expects tool '$tool' outside skill '$skillName' allowed-tools"
                    }
                }
            }
        }
    }

    $snapshotDir = Join-Path $Context.RepoRoot "tests\golden\tasks\expect"
    if (Test-Path -LiteralPath $snapshotDir) {
        Add-CorpusError $errors "tests/golden/tasks/expect exists; whole-document snapshots are forbidden"
    }

    return ,$errors
}

function Get-QuickActionIdsFromBlock {
    param([string]$Source, [string]$BlockName)
    $pattern = "(?s)" + [regex]::Escape($BlockName) + "\s+As IReadOnlyList\(Of QuickAction\)\s*=\s*New QuickAction\(\)\s*\{(.*?)\r?\n        \}"
    $match = [regex]::Match($Source, $pattern)
    if (-not $match.Success) {
        throw "QuickActionCatalog.vb no longer has the $BlockName block; the corpus gate cannot resolve entries"
    }
    $ids = @([regex]::Matches($match.Groups[1].Value, '\.Id = "([^"]+)"') | ForEach-Object { $_.Groups[1].Value })
    if ($ids.Count -eq 0) {
        throw "QuickActionCatalog.vb $BlockName block has no ids; the corpus gate would be blind"
    }
    return ,$ids
}

function Get-RibbonIds {
    param([string]$Source)
    $ids = New-Object System.Collections.ArrayList
    foreach ($match in [regex]::Matches($Source, "(?s)Public Shared ReadOnly \w+ As QuickAction = New QuickAction With \{(.*?)\r?\n        \}")) {
        $idMatch = [regex]::Match($match.Groups[1].Value, '\.Id = "([^"]+)"')
        if ($idMatch.Success) { [void]$ids.Add($idMatch.Groups[1].Value) }
    }
    if ($ids.Count -eq 0) {
        throw "QuickActionCatalog.vb has no ribbon QuickAction declarations; the corpus gate would be blind"
    }
    return ,$ids
}

function Get-PackageCatalog {
    param([string]$PackageDir)
    $packages = @{}
    foreach ($file in @(Get-ChildItem -LiteralPath $PackageDir -Filter *.json -File)) {
        $package = Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
        $id = [string]$package.id
        if (-not [string]::IsNullOrWhiteSpace($id)) { $packages[$id] = $package }
    }
    if ($packages.Count -eq 0) {
        throw "no task packages found under $PackageDir; entry/package resolution would be blind"
    }
    return $packages
}

function Get-SkillCatalog {
    param([string]$SkillsRoot)
    $skills = @{}
    foreach ($directory in @(Get-ChildItem -LiteralPath $SkillsRoot -Directory)) {
        $skillFile = Join-Path $directory.FullName "SKILL.md"
        if (-not (Test-Path -LiteralPath $skillFile)) { continue }
        $name = ""
        $allowed = @()
        $inFrontMatter = $false
        $frontMatterDone = $false
        foreach ($line in @(Get-Content -LiteralPath $skillFile -Encoding UTF8)) {
            if (-not $frontMatterDone -and $line.Trim() -eq "---") {
                if ($inFrontMatter) { $frontMatterDone = $true } else { $inFrontMatter = $true }
                continue
            }
            if (-not $inFrontMatter -or $frontMatterDone) { continue }
            $nameMatch = [regex]::Match($line, "^name:\s*(.+)$")
            if ($nameMatch.Success) { $name = $nameMatch.Groups[1].Value.Trim(); continue }
            $toolsMatch = [regex]::Match($line, "^allowed-tools:\s*(.+)$")
            if ($toolsMatch.Success) {
                $allowed = @($toolsMatch.Groups[1].Value.Split(",") | ForEach-Object { $_.Trim() } | Where-Object { $_ })
            }
        }
        $skills[$directory.Name] = [pscustomobject]@{
            Name = $name
            AllowedTools = @($allowed)
            Path = $skillFile
        }
    }
    if ($skills.Count -eq 0) {
        throw "no skills found under $SkillsRoot; skill resolution would be blind"
    }
    return $skills
}

function Get-ExceptionCodes {
    param([string]$Source)
    $codes = @([regex]::Matches($Source, 'Public Const Code\w+ As String = "([A-Z_]+)"') | ForEach-Object { $_.Groups[1].Value })
    if ($codes.Count -eq 0) {
        throw "ExceptionClassifier.vb exposes no error code constants; the corpus gate would be blind"
    }
    return ,$codes
}

function Test-GoldenFixtureDigestStability {
    param($ManifestA, $ManifestB)
    $errors = New-Object System.Collections.ArrayList
    $mapA = @{}
    foreach ($fixture in @($ManifestA.fixtures)) { $mapA[[string]$fixture.id] = [string]$fixture.contentSha256 }
    $mapB = @{}
    foreach ($fixture in @($ManifestB.fixtures)) { $mapB[[string]$fixture.id] = [string]$fixture.contentSha256 }
    if ($mapA.Count -ne $mapB.Count) {
        Add-CorpusError $errors "fixture generation produced $($mapA.Count) and $($mapB.Count) fixtures on two runs"
    }
    foreach ($id in @($mapA.Keys)) {
        if (-not $mapB.ContainsKey($id)) {
            Add-CorpusError $errors "fixture '$id' was generated on the first run only"
        }
        elseif ($mapA[$id] -ne $mapB[$id]) {
            Add-CorpusError $errors "fixture '$id' canonical digest is not stable across two runs"
        }
        if (@($mapA[$id]).Count -eq 0 -or $mapA[$id].Length -ne 64) {
            Add-CorpusError $errors "fixture '$id' canonical digest is not a SHA256 hex string"
        }
    }
    return ,$errors
}

function Test-ProbeCaught {
    param($ProbeCases, $ProbeSpecs, [hashtable]$Context)
    $probeErrors = Test-GoldenCorpus -Cases $ProbeCases -Specs $ProbeSpecs -Context $Context
    return ($probeErrors.Count -gt 0)
}

Initialize-GoldenOpenXml -RepoRoot $repoRoot

$vocabulary = Get-Content -LiteralPath (Join-Path $tasksDir "assertions.v1.json") -Raw -Encoding UTF8 | ConvertFrom-Json
$catalog = Get-Content -LiteralPath (Join-Path $tasksDir "catalog.json") -Raw -Encoding UTF8 | ConvertFrom-Json
$specs = @(Get-ChildItem -LiteralPath $fixturesDir -Filter *.json -File -Recurse | Sort-Object FullName | ForEach-Object {
    Get-Content -LiteralPath $_.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
})

$catalogSource = Get-Content -LiteralPath (Join-Path $repoRoot "ShareRibbon\Agent\QuickActionCatalog.vb") -Raw -Encoding UTF8
$quickActionIds = @(
    (Get-QuickActionIdsFromBlock -Source $catalogSource -BlockName "_wordActions") +
    (Get-QuickActionIdsFromBlock -Source $catalogSource -BlockName "_excelActions") +
    (Get-QuickActionIdsFromBlock -Source $catalogSource -BlockName "_powerPointActions")
)
$ribbonIds = Get-RibbonIds -Source $catalogSource
foreach ($prefix in @("word.ribbon.", "excel.ribbon.", "ppt.ribbon.")) {
    if (@($ribbonIds | Where-Object { $_ -like "$prefix*" }).Count -eq 0) {
        throw "QuickActionCatalog.vb has no $prefix declarations; the corpus gate would be blind"
    }
}

$packages = Get-PackageCatalog -PackageDir (Join-Path $repoRoot "ShareRibbon\TaskPackages")
$skills = Get-SkillCatalog -SkillsRoot (Join-Path $repoRoot "ShareRibbon\Skills")
$errorCodes = Get-ExceptionCodes -Source (Get-Content -LiteralPath (Join-Path $repoRoot "ShareRibbon\Common\ExceptionClassifier.vb") -Raw -Encoding UTF8)
foreach ($requiredCode in @("SAFETY_NEEDS_APPROVAL", "VBA_DISABLED", "TOOL_NOT_ALLOWED")) {
    if ($errorCodes -notcontains $requiredCode) {
        throw "ExceptionClassifier.vb no longer exposes $requiredCode; the corpus gate cannot ground negative cases"
    }
}
$planText = Get-Content -LiteralPath (Join-Path $repoRoot "openspec\changes\product-ai-native-strategy-execution-plan.md") -Raw -Encoding UTF8

$tempA = Join-Path ([System.IO.Path]::GetTempPath()) ("golden-corpus-a-" + [Guid]::NewGuid().ToString("N"))
$tempB = Join-Path ([System.IO.Path]::GetTempPath()) ("golden-corpus-b-" + [Guid]::NewGuid().ToString("N"))
$preErrors = New-Object System.Collections.ArrayList
$fixtureText = @{}
$fixtureSlideCounts = @{}
$fixtureAppTypes = @{}
$fixtureModels = @{}
$fixtureIds = @{}

try {
    $manifestA = New-GoldenFixtures -SpecDir $fixturesDir -OutputDir $tempA
    $manifestB = New-GoldenFixtures -SpecDir $fixturesDir -OutputDir $tempB
    foreach ($digestError in (Test-GoldenFixtureDigestStability -ManifestA $manifestA -ManifestB $manifestB)) {
        Add-CorpusError $preErrors $digestError
    }

    foreach ($spec in $specs) {
        $fixtureId = [string]$spec.id
        $fixtureIds[$fixtureId] = $true
        $fixtureAppTypes[$fixtureId] = [string]$spec.appType
        $fileName = $fixtureId.Replace("/", "-") + (Get-GoldenExtension -Kind ([string]$spec.kind))
        $path = Join-Path $tempA $fileName
        if (-not (Test-Path -LiteralPath $path)) {
            Add-CorpusError $preErrors "fixture '$fixtureId' was not written to $fileName"
            continue
        }
        $model = Get-GoldenFixtureModel -Path $path
        $fixtureModels[$fixtureId] = $model
        $fixtureText[$fixtureId] = Get-GoldenFixtureText -Model $model
        $fixtureSlideCounts[$fixtureId] = @($model.Slides).Count
    }

    $gitFiles = @(& git -C $repoRoot ls-files -- tests/golden 2>$null)
    if ($LASTEXITCODE -ne 0) {
        throw "git ls-files failed; the corpus gate cannot prove no office binary is committed"
    }
    foreach ($binaryError in (Test-NoCommittedBinaries -Paths $gitFiles)) {
        Add-CorpusError $preErrors $binaryError
    }
    $diskBinaries = @(Get-ChildItem -LiteralPath (Join-Path $repoRoot "tests\golden") -Recurse -File |
        Where-Object { $_.FullName -notmatch "\\\.generated\\" -and $_.Extension -in @(".docx", ".xlsx", ".pptx") })
    foreach ($binary in $diskBinaries) {
        Add-CorpusError $preErrors "generated office fixture is sitting in the repo tree: $($binary.FullName)"
    }

    $traceKeys = @{}
    foreach ($property in $vocabulary.traceKeys.PSObject.Properties) { $traceKeys[$property.Name] = [string]$property.Value }

    $context = @{
        Vocabulary = $vocabulary
        QuickActionIds = $quickActionIds
        RibbonIds = $ribbonIds
        Packages = $packages
        Skills = $skills
        ErrorCodes = $errorCodes
        PlanText = $planText
        FixtureText = $fixtureText
        FixtureSlideCounts = $fixtureSlideCounts
        FixtureAppTypes = $fixtureAppTypes
        FixtureModels = $fixtureModels
        FixtureIds = $fixtureIds
        TraceKeys = $traceKeys
        RepoRoot = $repoRoot
    }

    $baseCases = @($catalog.cases)
    $baseSpecs = @($specs)
    $caseErrors = Test-GoldenCorpus -Cases $baseCases -Specs $baseSpecs -Context $context

    foreach ($errorMessage in @($preErrors) + @($caseErrors)) {
        Write-Host "FAIL: $errorMessage"
    }
    if ($preErrors.Count -gt 0 -or $caseErrors.Count -gt 0) {
        throw "golden corpus check failed with $($preErrors.Count + $caseErrors.Count) error(s)"
    }

    $probeResults = New-Object System.Collections.ArrayList

    $cases = @($baseCases | Where-Object { [string](Get-Optional $_ "readiness") -ne "gap" })
    [void]$probeResults.Add([pscustomobject]@{ Name = "drop-gap-cases"; Caught = (Test-ProbeCaught -ProbeCases $cases -ProbeSpecs $baseSpecs -Context $context) })

    $cases = @($baseCases | Select-Object -Skip 8)
    [void]$probeResults.Add([pscustomobject]@{ Name = "drop-cases-below-threshold"; Caught = (Test-ProbeCaught -ProbeCases $cases -ProbeSpecs $baseSpecs -Context $context) })

    $cases = @(Copy-JsonObject $baseCases)
    $cases[1].id = $cases[0].id
    [void]$probeResults.Add([pscustomobject]@{ Name = "duplicate-case-id"; Caught = (Test-ProbeCaught -ProbeCases $cases -ProbeSpecs $baseSpecs -Context $context) })

    $cases = @(Copy-JsonObject $baseCases)
    $index = Find-CaseIndex -Cases $cases -Id "W-proofread-selection"
    $cases[$index].expect.document.assertions[0].type = "not_a_real_assertion"
    [void]$probeResults.Add([pscustomobject]@{ Name = "unknown-assertion"; Caught = (Test-ProbeCaught -ProbeCases $cases -ProbeSpecs $baseSpecs -Context $context) })

    $cases = @(Copy-JsonObject $baseCases)
    $index = Find-CaseIndex -Cases $cases -Id "W-proofread-selection"
    $cases[$index].expect.status = "probably_done"
    [void]$probeResults.Add([pscustomobject]@{ Name = "unknown-status"; Caught = (Test-ProbeCaught -ProbeCases $cases -ProbeSpecs $baseSpecs -Context $context) })

    $cases = @(Copy-JsonObject $baseCases)
    $index = Find-CaseIndex -Cases $cases -Id "W-proofread-selection"
    $cases[$index].entryActionId = "word.selection.does-not-exist"
    [void]$probeResults.Add([pscustomobject]@{ Name = "ghost-quick-action"; Caught = (Test-ProbeCaught -ProbeCases $cases -ProbeSpecs $baseSpecs -Context $context) })

    $cases = @(Copy-JsonObject $baseCases)
    $index = Find-CaseIndex -Cases $cases -Id "W-proofread-selection"
    $cases[$index].entryActionId = "word.ribbon.analyze"
    [void]$probeResults.Add([pscustomobject]@{ Name = "ribbon-id-used-as-quick-action"; Caught = (Test-ProbeCaught -ProbeCases $cases -ProbeSpecs $baseSpecs -Context $context) })

    $cases = @(Copy-JsonObject $baseCases)
    $index = Find-CaseIndex -Cases $cases -Id "W-numbering-check"
    $cases[$index].gap.probeEntryId = "word.selection.proofread"
    [void]$probeResults.Add([pscustomobject]@{ Name = "gap-entry-probe-resolves"; Caught = (Test-ProbeCaught -ProbeCases $cases -ProbeSpecs $baseSpecs -Context $context) })

    $cases = @(Copy-JsonObject $baseCases)
    $index = Find-CaseIndex -Cases $cases -Id "P-brand-template-gap"
    $cases[$index].gap.probePackageId = "excel.insight"
    [void]$probeResults.Add([pscustomobject]@{ Name = "gap-package-probe-resolves"; Caught = (Test-ProbeCaught -ProbeCases $cases -ProbeSpecs $baseSpecs -Context $context) })

    $cases = @(Copy-JsonObject $baseCases)
    $index = Find-CaseIndex -Cases $cases -Id "W-numbering-check"
    $cases[$index] | Add-Member -NotePropertyName "packageId" -NotePropertyValue "word.review-polish" -Force
    [void]$probeResults.Add([pscustomobject]@{ Name = "gap-claims-package"; Caught = (Test-ProbeCaught -ProbeCases $cases -ProbeSpecs $baseSpecs -Context $context) })

    $cases = @(Copy-JsonObject $baseCases)
    $index = Find-CaseIndex -Cases $cases -Id "W-proofread-selection"
    $cases[$index].packageId = "excel.insight"
    [void]$probeResults.Add([pscustomobject]@{ Name = "package-entry-mismatch"; Caught = (Test-ProbeCaught -ProbeCases $cases -ProbeSpecs $baseSpecs -Context $context) })

    $cases = @(Copy-JsonObject $baseCases)
    $index = Find-CaseIndex -Cases $cases -Id "W-proofread-selection"
    $cases[$index].fixtureId = "word/does-not-exist"
    [void]$probeResults.Add([pscustomobject]@{ Name = "missing-fixture"; Caught = (Test-ProbeCaught -ProbeCases $cases -ProbeSpecs $baseSpecs -Context $context) })

    $cases = @(Copy-JsonObject $baseCases)
    $index = Find-CaseIndex -Cases $cases -Id "W-proofread-selection"
    $cases[$index].fixtureId = "excel/sales-zh"
    [void]$probeResults.Add([pscustomobject]@{ Name = "app-fixture-mismatch"; Caught = (Test-ProbeCaught -ProbeCases $cases -ProbeSpecs $baseSpecs -Context $context) })

    $cases = @(Copy-JsonObject $baseCases)
    $index = Find-CaseIndex -Cases $cases -Id "W-proofread-selection"
    $cases[$index].setup.selection.to = 99
    [void]$probeResults.Add([pscustomobject]@{ Name = "selection-out-of-range"; Caught = (Test-ProbeCaught -ProbeCases $cases -ProbeSpecs $baseSpecs -Context $context) })

    $cases = @(Copy-JsonObject $baseCases)
    $index = Find-CaseIndex -Cases $cases -Id "W-clear-document-negative"
    $cases[$index].expect.requiredErrorCodes = @("NOT_A_REAL_CODE")
    [void]$probeResults.Add([pscustomobject]@{ Name = "unknown-error-code"; Caught = (Test-ProbeCaught -ProbeCases $cases -ProbeSpecs $baseSpecs -Context $context) })

    $cases = @(Copy-JsonObject $baseCases)
    $index = Find-CaseIndex -Cases $cases -Id "W-clear-document-negative"
    $cases[$index].expect.status = "completed"
    [void]$probeResults.Add([pscustomobject]@{ Name = "negative-marked-completed"; Caught = (Test-ProbeCaught -ProbeCases $cases -ProbeSpecs $baseSpecs -Context $context) })

    $cases = @(Copy-JsonObject $baseCases)
    $index = Find-CaseIndex -Cases $cases -Id "W-clear-document-negative"
    $cases[$index].expect.trace.PSObject.Properties.Remove("maxWriteSteps")
    [void]$probeResults.Add([pscustomobject]@{ Name = "blocked-without-write-cap"; Caught = (Test-ProbeCaught -ProbeCases $cases -ProbeSpecs $baseSpecs -Context $context) })

    $cases = @(Copy-JsonObject $baseCases)
    $index = Find-CaseIndex -Cases $cases -Id "W-proofread-selection"
    $cases[$index].layer = @("L2")
    [void]$probeResults.Add([pscustomobject]@{ Name = "layer-without-l1"; Caught = (Test-ProbeCaught -ProbeCases $cases -ProbeSpecs $baseSpecs -Context $context) })

    $cases = @(Copy-JsonObject $baseCases)
    $index = Find-CaseIndex -Cases $cases -Id "W-proofread-selection"
    $cases[$index].harnessOptions.allowVba = $true
    [void]$probeResults.Add([pscustomobject]@{ Name = "vba-enabled"; Caught = (Test-ProbeCaught -ProbeCases $cases -ProbeSpecs $baseSpecs -Context $context) })

    $cases = @(Copy-JsonObject $baseCases)
    $index = Find-CaseIndex -Cases $cases -Id "W-proofread-selection"
    $cases[$index].expect.maxSteps = 0
    [void]$probeResults.Add([pscustomobject]@{ Name = "zero-step-budget"; Caught = (Test-ProbeCaught -ProbeCases $cases -ProbeSpecs $baseSpecs -Context $context) })

    $cases = @(Copy-JsonObject $baseCases)
    $index = Find-CaseIndex -Cases $cases -Id "W-proofread-selection"
    $cases[$index].ciTier = "pull-request"
    [void]$probeResults.Add([pscustomobject]@{ Name = "unknown-ci-tier"; Caught = (Test-ProbeCaught -ProbeCases $cases -ProbeSpecs $baseSpecs -Context $context) })

    $cases = @(Copy-JsonObject $baseCases)
    $index = Find-CaseIndex -Cases $cases -Id "W-proofread-selection"
    $cases[$index].expect.trace | Add-Member -NotePropertyName "snapshot" -NotePropertyValue $true -Force
    [void]$probeResults.Add([pscustomobject]@{ Name = "unknown-trace-key"; Caught = (Test-ProbeCaught -ProbeCases $cases -ProbeSpecs $baseSpecs -Context $context) })

    $cases = @(Copy-JsonObject $baseCases)
    $index = Find-CaseIndex -Cases $cases -Id "W-proofread-selection"
    $cases[$index].expect.document | Add-Member -NotePropertyName "snapshot" -NotePropertyValue "whole-document" -Force
    [void]$probeResults.Add([pscustomobject]@{ Name = "document-snapshot"; Caught = (Test-ProbeCaught -ProbeCases $cases -ProbeSpecs $baseSpecs -Context $context) })

    $cases = @(Copy-JsonObject $baseCases)
    $index = Find-CaseIndex -Cases $cases -Id "W-proofread-selection"
    $cases[$index].expect.document.assertions[0].type = "slide_count_equals"
    $cases[$index].expect.document.assertions[0] | Add-Member -NotePropertyName "count" -NotePropertyValue 1 -Force
    [void]$probeResults.Add([pscustomobject]@{ Name = "assertion-app-mismatch"; Caught = (Test-ProbeCaught -ProbeCases $cases -ProbeSpecs $baseSpecs -Context $context) })

    $cases = @(Copy-JsonObject $baseCases)
    $index = Find-CaseIndex -Cases $cases -Id "W-proofread-selection"
    $cases[$index].skill = "not-a-real-skill"
    [void]$probeResults.Add([pscustomobject]@{ Name = "ghost-skill"; Caught = (Test-ProbeCaught -ProbeCases $cases -ProbeSpecs $baseSpecs -Context $context) })

    $probeSpecs = @(Copy-JsonObject $baseSpecs)
    $specIndex = 0
    for ($i = 0; $i -lt $probeSpecs.Count; $i++) {
        if ([string]$probeSpecs[$i].id -eq "word/report-zh") { $specIndex = $i; break }
    }
    $probeSpecs[$specIndex].expectContains = @($probeSpecs[$specIndex].expectContains) + @("definitely-not-present")
    [void]$probeResults.Add([pscustomobject]@{ Name = "fixture-readback-text-missing"; Caught = (Test-ProbeCaught -ProbeCases $baseCases -ProbeSpecs $probeSpecs -Context $context) })

    $probeSpecs = @(Copy-JsonObject $baseSpecs)
    $probeSpecs[1].id = $probeSpecs[0].id
    [void]$probeResults.Add([pscustomobject]@{ Name = "duplicate-fixture-id"; Caught = (Test-ProbeCaught -ProbeCases $baseCases -ProbeSpecs $probeSpecs -Context $context) })

    $binaryProbe = Test-NoCommittedBinaries -Paths @("tests/golden/tasks/fixtures/word/real-customer.docx")
    [void]$probeResults.Add([pscustomobject]@{ Name = "committed-office-binary"; Caught = ($binaryProbe.Count -gt 0) })

    $missed = @($probeResults | Where-Object { -not $_.Caught })
    if ($missed.Count -gt 0) {
        foreach ($missedProbe in $missed) { Write-Host "FAIL: mutation probe '$($missedProbe.Name)' did not turn the gate red" }
        throw "$($missed.Count) mutation probe(s) did not turn the gate red"
    }

    $readyCount = @($baseCases | Where-Object { [string](Get-Optional $_ "readiness") -eq "ready" }).Count
    $gapCount = @($baseCases | Where-Object { [string](Get-Optional $_ "readiness") -eq "gap" }).Count
    $negativeCount = @($baseCases | Where-Object { (Get-OptionalArray $_ "tags") -contains "negative" }).Count

    Write-Host ("PASS golden corpus: {0} cases (Word {1}, Excel {2}, PowerPoint {3}), {4} ready, {5} gap, {6} negative, {7} fixtures, {8} mutation probes turned red" -f
        $baseCases.Count,
        @($baseCases | Where-Object { $_.appType -eq "Word" }).Count,
        @($baseCases | Where-Object { $_.appType -eq "Excel" }).Count,
        @($baseCases | Where-Object { $_.appType -eq "PowerPoint" }).Count,
        $readyCount,
        $gapCount,
        $negativeCount,
        $specs.Count,
        $probeResults.Count)
}
finally {
    if (Test-Path -LiteralPath $tempA) { Remove-Item -LiteralPath $tempA -Recurse -Force -ErrorAction SilentlyContinue }
    if (Test-Path -LiteralPath $tempB) { Remove-Item -LiteralPath $tempB -Recurse -Force -ErrorAction SilentlyContinue }
}
