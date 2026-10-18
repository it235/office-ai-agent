param(
    [switch]$Build,
    [string]$Configuration = "Debug",
    [string]$Platform = "AnyCPU"
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$msbuild = "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"

if ($Build) {
    if (-not (Test-Path $msbuild)) {
        throw "MSBuild not found: $msbuild"
    }
    & $msbuild (Join-Path $repoRoot "WordAi\WordAi.vbproj") /p:Configuration=$Configuration /p:Platform=$Platform
    if ($LASTEXITCODE -ne 0) {
        throw "MSBuild failed, exit code: $LASTEXITCODE"
    }
}

$bin = Join-Path $repoRoot "ShareRibbon\bin\$Configuration"
$dll = Join-Path $bin "ShareRibbon.dll"
if (-not (Test-Path $dll)) {
    throw "ShareRibbon.dll not found. Run: .\scripts\smoke-script-safety.ps1 -Build"
}

Push-Location $bin
try {
    Add-Type -Path $dll

    $denyAction = [ShareRibbon.Agent.Execution.SafetyAction]::Deny
    $approvalAction = [ShareRibbon.Agent.Execution.SafetyAction]::RequireApproval
    $allowAction = [ShareRibbon.Agent.Execution.SafetyAction]::Allow

    function Get-Verdict([string]$code, [bool]$alwaysConfirm) {
        return [ShareRibbon.Agent.Execution.SafetyGate]::EvaluateScript($code, "risky", $alwaysConfirm)
    }

    function Assert-Verdict([string]$code, [bool]$alwaysConfirm, $expected, [string]$label) {
        $verdict = Get-Verdict $code $alwaysConfirm
        if ($verdict.Action -ne $expected) {
            throw "$label : expected $expected but got $($verdict.Action) for [$code] ($($verdict.Reason))"
        }
        Write-Host "  ok  $label -> $($verdict.Action)"
    }

    Write-Host "==> Dangerous script surfaces are denied even when obfuscated with spacing or case"
    Assert-Verdict 'Shell ( "calc.exe" )' $false $denyAction "shell with space"
    Assert-Verdict 'KILL("a.txt")' $false $denyAction "kill uppercase"
    Assert-Verdict 'CreateObject ( "WScript.Shell" )' $false $denyAction "wscript shell via CreateObject"
    Assert-Verdict 'new ActiveXObject("MSScriptControl.ScriptControl")' $false $denyAction "script control instantiation"
    Assert-Verdict 'Application . Quit()' $false $denyAction "application quit with spaces"
    Assert-Verdict 'fso. DeleteFile "x"' $false $denyAction "filesystem delete file"
    Assert-Verdict 'Documents . Close' $false $denyAction "documents close"

    Write-Host "==> Destructive but legitimate operations require approval"
    Assert-Verdict 'Selection. Delete()' $false $approvalAction "selection delete"
    Assert-Verdict 'ActiveDocument.SaveCopyAs "copy.docx"' $false $approvalAction "save copy as"

    Write-Host "==> Benign code is allowed on the UI path and still confirmed on the agent path"
    Assert-Verdict 'ActiveSheet.Name = "Report"' $false $allowAction "benign assignment (ui path)"
    Assert-Verdict 'ActiveSheet.Name = "Report"' $true $approvalAction "benign assignment (agent vba path)"

    Write-Host "Script safety gate smoke passed."
}
finally {
    Pop-Location
}
