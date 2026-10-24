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

    $smokeScripts = @(
        "scripts\audit-p0-guardrails.ps1",
        "scripts\smoke-db-schema-drift.ps1",
        "scripts\smoke-empty-db-initialization.ps1",
        "scripts\smoke-change-log-and-session-scope.ps1",
        "scripts\smoke-script-safety.ps1",
        "scripts\smoke-harness-approval.ps1",
        "scripts\smoke-run-cancellation.ps1",
        "scripts\smoke-memory-pipeline.ps1",
        "scripts\smoke-skills-registry.ps1",
        "scripts\smoke-skills-usage-migration.ps1",
        "scripts\smoke-word-capability-registry.ps1",
        "scripts\smoke-ai-gateway-provider.ps1",
        "scripts\smoke-model-config-gate.ps1",
        "scripts\smoke-config-snapshot.ps1",
        "scripts\smoke-onboarding-wizard.ps1",
        "scripts\run-golden-l0.ps1"
    )

    foreach ($smoke in $smokeScripts) {
        $fullPath = Join-Path $repoRoot $smoke
        if (-not (Test-Path -LiteralPath $fullPath)) {
            throw "Smoke script not found: $smoke"
        }

        Invoke-Step $smoke {
            if ($smoke -eq "scripts\run-golden-l0.ps1") {
                powershell -NoProfile -ExecutionPolicy Bypass -File $fullPath -Configuration $Configuration
            }
            else {
                powershell -NoProfile -ExecutionPolicy Bypass -File $fullPath
            }
        }
    }

    if (-not $SkipNodeCheck) {
        $node = Get-Command node -ErrorAction SilentlyContinue
        if ($node) {
            $jsFiles = @(
                "ShareRibbon\Resources\js\message-sender.js",
                "ShareRibbon\Resources\js\agent-card.js",
                "ShareRibbon\Resources\js\office-ai-bridge.js",
                "ShareRibbon\Resources\js\chat-manager.js",
                "ShareRibbon\Resources\js\onboarding.js",
                "ShareRibbon\Resources\js\history-manager.js",
                "ShareRibbon\Resources\js\ralph-loop.js"
            )

            foreach ($js in $jsFiles) {
                $fullPath = Join-Path $repoRoot $js
                if (Test-Path -LiteralPath $fullPath) {
                    Invoke-Step "node --check $js" {
                        node --check $fullPath
                    }
                }
            }
        }
        else {
            Write-Warning "node not found; skipping JS syntax checks. Use -SkipNodeCheck to make this explicit."
        }
    }

    Write-Host ""
    Write-Host "All code checks passed."
}
finally {
    Pop-Location
}
