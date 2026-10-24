param(
    [string]$Configuration = "Debug"
)

# E3-01: the in-panel wizard only owns what the config window never had - the privacy acknowledgement and
# the example tasks. Provider choice, api key entry and the connectivity probe belong to ConfigApiForm.
# This smoke pins that boundary: the pushed payload must not carry a provider list, the wizard step must
# come only from the host, and exercising the whole wizard must not write the config file at all.
# The tail keeps the shared connectivity probe checks, because ConfigApiForm is now its only caller and
# still has to refuse an unusable target before any socket is opened.
# Keep this file ASCII only: PowerShell 5.1 reads a BOM-less script as ANSI.

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$bin = Join-Path $repoRoot "ShareRibbon\bin\$Configuration"
$dll = Join-Path $bin "ShareRibbon.dll"

if (-not (Test-Path -LiteralPath $dll)) {
    throw "ShareRibbon.dll not found. Run: .\scripts\build-code-projects.ps1 -Configuration $Configuration -Projects `"ShareRibbon\ShareRibbon.vbproj`""
}

$configPath = Join-Path ([Environment]::GetFolderPath("MyDocuments")) "OfficeAiAppData\office_ai_config.json"
$securePath = Join-Path ([Environment]::GetFolderPath("ApplicationData")) "AiHelper\secure.config"

function Read-BytesIfExists($path) {
    if (Test-Path -LiteralPath $path) { return [System.IO.File]::ReadAllBytes($path) }
    return $null
}

function Restore-File($path, $bytes) {
    $dir = Split-Path -Parent $path
    if (-not (Test-Path -LiteralPath $dir)) { [void][System.IO.Directory]::CreateDirectory($dir) }
    if ($null -eq $bytes) {
        if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
    }
    else {
        [System.IO.File]::WriteAllBytes($path, $bytes)
    }
}

function Test-SameBytes($left, $right) {
    if ($null -eq $left -and $null -eq $right) { return $true }
    if ($null -eq $left -or $null -eq $right) { return $false }
    if ($left.Length -ne $right.Length) { return $false }
    for ($i = 0; $i -lt $left.Length; $i++) {
        if ($left[$i] -ne $right[$i]) { return $false }
    }
    return $true
}

$configBackup = Read-BytesIfExists $configPath
$secureBackup = Read-BytesIfExists $securePath

Push-Location $bin
try {
    Add-Type -Path $dll

    $script:pushes = @()
    $script:appType = "excel"
    $script:configRefreshCalls = 0

    $execute = [System.Func[string, System.Threading.Tasks.Task]] {
        param($js)
        $script:pushes += $js
        return [System.Threading.Tasks.Task]::CompletedTask
    }
    $getAppType = [System.Func[string]] { param() return $script:appType }
    $onConfigChanged = [System.Action] { $script:configRefreshCalls += 1 }

    function Get-PushedState {
        param([int]$index)
        $script = $script:pushes[$index]
        if ($script -notmatch "^renderOnboarding\(") {
            throw "the wizard must push renderOnboarding(...);, got '$script'."
        }
        $json = $script.Substring("renderOnboarding(".Length)
        $json = $json.Substring(0, $json.Length - 2)
        return ConvertFrom-Json -InputObject $json
    }

    function New-Doc($map) {
        $doc = New-Object Newtonsoft.Json.Linq.JObject
        foreach ($key in $map.Keys) {
            [void]$doc.Add($key, [Newtonsoft.Json.Linq.JValue]::CreateString([string]$map[$key]))
        }
        return $doc
    }

    function New-Wizard {
        return New-Object ShareRibbon.OnboardingService -ArgumentList @($execute, $getAppType, $onConfigChanged)
    }

    # 1. a first-run user with nothing configured: the wizard opens on privacy and reports the gap
    $wiz = New-Wizard
    [ShareRibbon.ConfigSettings]::ApiUrl = ""
    [ShareRibbon.ConfigSettings]::ApiKey = ""
    [ShareRibbon.Security.SecureConfigManager+PrivacySettings]::PrivacyPolicyAccepted = $false
    $count = $script:pushes.Count
    $wiz.HandleGetState()
    if ($script:pushes.Count -ne $count + 1) { throw "HandleGetState must push exactly one payload." }
    $state = Get-PushedState $count

    if ($state.needsSetup -ne $true) { throw "an empty url and key must report needsSetup." }
    if ($state.privacyAccepted -ne $false) { throw "the payload must report the real privacy flag." }
    if ($state.step -ne "privacy") { throw "an unaccepted privacy policy must open the privacy step, got '$($state.step)'." }
    if ($state.appType -ne "excel") { throw "the payload must carry the host type so examples match the app." }
    # The provider list was the duplication: it must not come back through the payload.
    if ($null -ne $state.providers) { throw "the wizard must not carry a provider list; ConfigApiForm owns providers." }

    # 2. needsSetup follows the same validator the send gate uses, not a wizard-local opinion
    [ShareRibbon.ConfigSettings]::ApiUrl = "https://example.invalid/v1/chat/completions"
    [ShareRibbon.ConfigSettings]::ApiKey = "sk-smoke-not-a-real-key"
    $count = $script:pushes.Count
    $wiz.HandleGetState()
    if ((Get-PushedState $count).needsSetup -ne $false) { throw "a url plus a key must count as configured." }
    if ((Get-PushedState $count).step -ne "privacy") { throw "privacy still comes first, got '$((Get-PushedState $count).step)'." }

    # 3. accepting privacy is the only state the wizard itself persists, and it must reach the store
    $wiz.HandleSetPrivacy((New-Doc @{ accepted = "true" }))
    if (-not [ShareRibbon.Security.SecureConfigManager+PrivacySettings]::PrivacyPolicyAccepted) {
        throw "accepting the privacy policy must persist."
    }
    $afterPrivacy = Get-PushedState ($script:pushes.Count - 1)
    if ($afterPrivacy.privacyAccepted -ne $true) { throw "the pushed payload must show the acknowledgement." }
    if ($afterPrivacy.step -ne "example") { throw "with privacy accepted and a model configured the wizard ends on the example step, got '$($afterPrivacy.step)'." }

    # 4. withdrawing the acknowledgement moves the wizard back, so the step never lives in the panel
    $wiz.HandleSetPrivacy((New-Doc @{ accepted = "false" }))
    if ([ShareRibbon.Security.SecureConfigManager+PrivacySettings]::PrivacyPolicyAccepted) {
        throw "revoking the privacy acknowledgement must persist."
    }
    if ((Get-PushedState ($script:pushes.Count - 1)).step -ne "privacy") { throw "the step must follow the host state on every push." }
    $wiz.HandleSetPrivacy((New-Doc @{ }))
    if ([ShareRibbon.Security.SecureConfigManager+PrivacySettings]::PrivacyPolicyAccepted) {
        throw "a payload without an accepted flag must not be read as consent."
    }
    $wiz.HandleSetPrivacy((New-Doc @{ accepted = "TRUE" }))

    # 5. examples follow the host, because a Word deck tip in Excel is noise
    $excelExamples = @(Get-PushedState ($script:pushes.Count - 1)).examples
    if ($excelExamples.Count -lt 2) { throw "the wizard needs example tasks to end on, got $($excelExamples.Count)." }
    $script:appType = "powerpoint"
    $wiz.HandleGetState()
    $pptExamples = @(Get-PushedState ($script:pushes.Count - 1)).examples
    if (-not (Compare-Object $excelExamples $pptExamples)) {
        throw "example tasks must differ per host; excel and powerpoint returned the same list."
    }
    $script:appType = "word"
    $wiz.HandleGetState()
    $wordExamples = @(Get-PushedState ($script:pushes.Count - 1)).examples
    if (-not (Compare-Object $excelExamples $wordExamples)) {
        throw "the fallback example list must still be word-shaped, not the excel list."
    }
    $script:appType = "excel"

    # 6. finishing only asks the host to re-read the model display - the wizard never touches config itself
    $refreshes = $script:configRefreshCalls
    $wiz.HandleFinish()
    if ($script:configRefreshCalls -ne $refreshes + 1) { throw "HandleFinish must refresh the model display exactly once." }
    if (-not (Test-SameBytes $configBackup (Read-BytesIfExists $configPath))) {
        throw "exercising the wizard rewrote the config file; provider and key writes belong to ConfigApiForm."
    }

    # 7. the shared connectivity probe still refuses an unusable target before any socket is opened
    $tester = New-Object ShareRibbon.ModelConnectivityTester
    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    $noUrl = $tester.TestAsync("", "any-key", "some-model", "default", "smoke").GetAwaiter().GetResult()
    $noModel = $tester.TestAsync("https://127.0.0.1:1/chat/completions", "any-key", "", "default", "smoke").GetAwaiter().GetResult()
    $watch.Stop()
    if ($noUrl.Success -or $noUrl.Reason -ne "invalid_config") { throw "an empty api url must fail as invalid_config, got '$($noUrl.Reason)'." }
    if ($noModel.Success -or $noModel.Reason -ne "invalid_config") { throw "an empty model name must fail as invalid_config, got '$($noModel.Reason)'." }
    if ([string]::IsNullOrWhiteSpace($noUrl.Detail)) { throw "the tester must explain why nothing was attempted." }
    if ($watch.ElapsedMilliseconds -gt 3000) { throw "invalid config must be rejected offline, took $($watch.ElapsedMilliseconds)ms." }
    $json = ConvertFrom-Json -InputObject $noUrl.ToJson()
    if ($json.reason -ne "invalid_config") { throw "the test result must stay machine readable for the panel." }

    Write-Host "OK: the wizard stays a two-step shell over the existing config window and never writes config itself."
    [pscustomobject]@{
        Steps           = "privacy,example"
        ExcelExamples   = $excelExamples.Count
        PptExamples     = $pptExamples.Count
        ConfigPushes    = $script:pushes.Count
        OfflineReason   = $noUrl.Reason
    } | Format-List
}
finally {
    Restore-File $configPath $configBackup
    Restore-File $securePath $secureBackup
    Pop-Location
}
