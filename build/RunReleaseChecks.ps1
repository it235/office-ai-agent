param(
    [switch]$SkipBuild,
    [switch]$SkipSmoke,
    [switch]$VerifySignatures
)

$ErrorActionPreference = 'Stop'
$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')

function Invoke-Step {
    param(
        [string]$Name,
        [scriptblock]$Action
    )

    Write-Host ""
    Write-Host "==> $Name"
    $global:LASTEXITCODE = 0
    & $Action
    # E5-05: a child script that ends with `exit 1` does not stop the caller and does not throw, so
    # every step below that audits with an exit code used to print PASS even when it failed. Measured
    # before the fix: a child `exit 1` produced "PASS" plus LASTEXITCODE=1 and the run kept going.
    if ($LASTEXITCODE -ne 0) {
        throw "$Name failed with exit code $LASTEXITCODE"
    }
    Write-Host "PASS: $Name"
}

function Get-MSBuildPath {
    $candidates = @(
        'C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe',
        'C:\Program Files\Microsoft Visual Studio\18\Professional\MSBuild\Current\Bin\MSBuild.exe',
        'C:\Program Files\Microsoft Visual Studio\18\Enterprise\MSBuild\Current\Bin\MSBuild.exe',
        'C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe',
        'C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe',
        'C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe',
        'C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe',
        'C:\Program Files (x86)\Microsoft Visual Studio\2019\Professional\MSBuild\Current\Bin\MSBuild.exe',
        'C:\Program Files (x86)\Microsoft Visual Studio\2019\Enterprise\MSBuild\Current\Bin\MSBuild.exe'
    )

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) {
            return $candidate
        }
    }

    $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
    if (Test-Path -LiteralPath $vswhere) {
        $installPath = & $vswhere -latest -products * -requires Microsoft.Component.MSBuild -property installationPath
        if (-not [string]::IsNullOrWhiteSpace($installPath)) {
            $msbuild = Join-Path $installPath 'MSBuild\Current\Bin\MSBuild.exe'
            if (Test-Path -LiteralPath $msbuild) {
                return $msbuild
            }
        }
    }

    $cmd = Get-Command MSBuild.exe -ErrorAction SilentlyContinue
    if ($cmd) {
        return $cmd.Source
    }

    throw 'MSBuild.exe not found. Install Visual Studio Build Tools or add MSBuild.exe to PATH.'
}

Set-Location $repoRoot

Invoke-Step 'Version audit' {
    & (Join-Path $repoRoot 'build\AuditVersion.ps1')
}

Invoke-Step 'Installer Release audit' {
    & (Join-Path $repoRoot 'build\AuditInstallerRelease.ps1')
}

if (-not $SkipBuild) {
    $msbuild = Get-MSBuildPath
    $projects = @(
        'ShareRibbon\ShareRibbon.vbproj',
        'WordAi\WordAi.vbproj',
        'ExcelAi\ExcelAi.vbproj',
        'PowerPointAi\PowerPointAi.vbproj'
    )

    foreach ($project in $projects) {
        Invoke-Step "Build Release $project" {
            & $msbuild (Join-Path $repoRoot $project) /p:Configuration=Release /p:Platform=AnyCPU
            if ($LASTEXITCODE -ne 0) {
                throw "MSBuild failed for $project"
            }
        }
    }
}

Invoke-Step 'Installer input audit' {
    & (Join-Path $repoRoot 'build\AuditInstallerInputs.ps1')
}

if (-not $SkipSmoke) {
    # E5-05: this was a second, private copy of the gate list and it had drifted to 5 of 22 entries,
    # so the capability-consistency guards (P7/P8) and every panel-contract gate were never part of a
    # release check even though the plan row says a drifting list must fail the build. The list now
    # has one owner (scripts\run-code-checks.ps1). It is called without -SkipBuild on purpose: the
    # reflection-based gates read bin\Debug, not the Release output built above.
    Invoke-Step 'Code gates (scripts\run-code-checks.ps1)' {
        & (Join-Path $repoRoot 'scripts\run-code-checks.ps1')
    }
}

Invoke-Step 'Git diff whitespace check' {
    $git = Get-Command git.exe -ErrorAction SilentlyContinue
    if ($null -eq $git) {
        $git = Get-Command git -ErrorAction SilentlyContinue
    }
    if ($null -eq $git) {
        throw 'git not found.'
    }

    & $git.Source -C $repoRoot diff --check
    if ($LASTEXITCODE -ne 0) {
        throw 'git diff --check failed.'
    }
}

if ($VerifySignatures) {
    Invoke-Step 'Signature verification' {
        & (Join-Path $repoRoot 'build\VerifySignatures.ps1')
    }
}

Write-Host ""
Write-Host 'Release checks complete.'
