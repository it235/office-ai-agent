param(
    [string]$Configuration = "Debug",
    [switch]$SkipBuild,
    [switch]$SkipNodeCheck
)

$ErrorActionPreference = "Stop"

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")

function Invoke-Step {
    param(
        [string]$Name,
        [scriptblock]$Action
    )

    Write-Host ""
    Write-Host "==> $Name"
    $global:LASTEXITCODE = 0
    & $Action
    if ($LASTEXITCODE -ne 0) {
        throw "$Name failed with exit code $LASTEXITCODE"
    }
    Write-Host "PASS: $Name"
}

Push-Location $repoRoot
try {
    if (-not $SkipBuild) {
        Invoke-Step "Build code projects ($Configuration)" {
            powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $repoRoot "scripts\build-code-projects.ps1") -Configuration $Configuration
        }
    }

    # E5-05: this list used to be hand-maintained, and it was maintained twice - here and in
    # build\RunReleaseChecks.ps1 - which is how the two drifted apart (22 gates here, 5 there, so the
    # capability-list guards P7/P8 were never release gates). Discover the gates from disk instead.
    # The inventory is inverted on purpose: every scripts\*.ps1 is a gate unless it names a reason to
    # be here. A name-prefix rule (smoke-*.ps1) looked general but was not - run-professional-design-
    # smoke.ps1 threw on real failures and never got discovered, so it never ran. Now a new gate is
    # picked up the day it lands, and opting out stays an explicit, readable decision.
    $notGates = @{
        "run-code-checks.ps1"            = "this runner; recursing into itself would loop"
        "build-code-projects.ps1"        = "build entry, not a check"
        "build-installer-prep.ps1"       = "installer build entry, not a check; its audits live in build\"
        "ensure-vsto-temp-keys.ps1"      = "developer-certificate setup, not a check"
        "audit-installer-duplicates.ps1" = "informational duplicate-SourcePath report, it never fails; the installer gates live in build\"
    }
    $gateNames = @(
        Get-ChildItem -LiteralPath (Join-Path $repoRoot "scripts") -Filter *.ps1 -File |
            Where-Object { -not ($notGates.Keys -contains $_.Name) } |
            Sort-Object Name |
            ForEach-Object { $_.Name }
    )
    if ($gateNames.Count -lt 20) {
        throw "only $($gateNames.Count) gate scripts discovered under scripts\ - the discovery itself is broken"
    }
    foreach ($notGate in $notGates.Keys) {
        if (-not (Test-Path -LiteralPath (Join-Path $repoRoot "scripts\$notGate"))) {
            throw "gate opt-out names a script that is not on disk anymore: $notGate"
        }
    }
    Write-Host ""
    Write-Host "    discovered $($gateNames.Count) gates under scripts\; opted out: $(($notGates.Keys | Sort-Object) -join ', ')"

    foreach ($name in $gateNames) {
        $fullPath = Join-Path $repoRoot "scripts\$name"
        Invoke-Step "scripts\$name" {
            if ($name -eq "run-golden-l0.ps1") {
                powershell -NoProfile -ExecutionPolicy Bypass -File $fullPath -Configuration $Configuration
            }
            else {
                powershell -NoProfile -ExecutionPolicy Bypass -File $fullPath
            }
        }
    }

    if (-not $SkipNodeCheck) {
        # E5-05: skipping silently was the bug - node is not on every machine, and when it is missing
        # the panel-contract gates simply did not run while the suite still reported "all passed".
        $node = Get-Command node -ErrorAction SilentlyContinue
        if (-not $node) {
            throw "node not found; the panel-contract gates (JS syntax, agent card, history grouping, entry reachability) cannot run. Install node or pass -SkipNodeCheck to say so on purpose."
        }

        # E5-04: the JS checklist here used to be hand-maintained, and it rotted in two directions:
        # ralph-loop.js had been deleted yet was skipped silently, while new panel JS was never
        # syntax-checked at all. Enumerate the real folder instead of trusting a checklist.
        $jsDir = Join-Path $repoRoot "ShareRibbon\Resources\js"
        $jsFiles = Get-ChildItem -LiteralPath $jsDir -Filter *.js -File |
            ForEach-Object { $_.FullName }
        if ($jsFiles.Count -eq 0) {
            throw "no JS assets found under $jsDir"
        }

        foreach ($fullPath in $jsFiles) {
            Invoke-Step ("node --check " + $fullPath.Substring($repoRoot.Path.Length + 1)) {
                node --check $fullPath
            }
        }

        # E1-08: the agent plan card contract (trust chips / step status tokens / terminal change
        # summary) is exercised against a stub DOM, so a VB-to-JS payload rename fails the gate.
        Invoke-Step "scripts\check-agent-card-ui.js" {
            node (Join-Path $repoRoot "scripts\check-agent-card-ui.js")
        }

        # E2-05: history sidebar grouping (document buckets / time buckets / guided empty state)
        # is asserted against the same payload field names the VB service serialises.
        Invoke-Step "scripts\check-history-grouping-ui.js" {
            node (Join-Path $repoRoot "scripts\check-history-grouping-ui.js")
        }

        # E5-04/E5-05: message types, inline events, DOM ids, vsto bridge methods, visible ribbon
        # buttons and both directions of the VB -> JS injection contract must all close.
        # ChatCommandRouter only Debug.Writes an unknown type, so an unreachable entry is silent.
        Invoke-Step "scripts\check-entry-reachability.js" {
            node (Join-Path $repoRoot "scripts\check-entry-reachability.js")
        }
    }

    Write-Host ""
    Write-Host "All code checks passed."
}
finally {
    Pop-Location
}
