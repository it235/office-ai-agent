param(
    [string]$Configuration = "Debug"
)

# E3-03 root cause guard: ConfigSettings is the process-wide snapshot every send reads for its
# url/key. It used to be written only inside an "if this provider is selected" loop at load time,
# so clearing a url or a key in the config window left the previous values live. The panel then sent
# to the old endpoint and the unconfigured gate was satisfied by the stale key.
# Keep this file ASCII only: PowerShell 5.1 reads a BOM-less script as ANSI.

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$bin = Join-Path $repoRoot "ShareRibbon\bin\$Configuration"
$dll = Join-Path $bin "ShareRibbon.dll"

if (-not (Test-Path -LiteralPath $dll)) {
    throw "ShareRibbon.dll not found. Run: .\scripts\build-code-projects.ps1 -Configuration $Configuration -Projects `"ShareRibbon\ShareRibbon.vbproj`""
}

Push-Location $bin
try {
    Add-Type -Path $dll

    $itemType = [Type]"ShareRibbon.ConfigManager+ConfigItem"
    $modelType = [Type]"ShareRibbon.ConfigManager+ConfigItemModel"
    $chatEnum = [ShareRibbon.ConfigManager+ModelType]::Chat
    $embedEnum = [ShareRibbon.ConfigManager+ModelType]::Embedding

    # PS 5.1 cannot build a generic List type from a [Type] variable, so use the name form.
    function New-ModelList {
        return New-Object "System.Collections.Generic.List[ShareRibbon.ConfigManager+ConfigItemModel]"
    }
    function New-ProviderList {
        return New-Object "System.Collections.Generic.List[ShareRibbon.ConfigManager+ConfigItem]"
    }

    function New-Model($name, [bool]$selected, [object]$type, [bool]$mcpable, [bool]$fim, [string]$fimUrl, [string]$reasoning) {
        $m = New-Object $modelType
        $m.modelName = $name
        $m.selected = $selected
        if ($null -ne $type) { $m.modelType = $type }
        $m.mcpable = $mcpable
        $m.fimSupported = $fim
        $m.fimUrl = $fimUrl
        $m.reasoningMode = $reasoning
        return $m
    }

    function New-Provider($platform, $url, $key, [bool]$selected, $models) {
        $p = New-Object $itemType
        $p.platform = $platform
        $p.url = $url
        $p.key = $key
        $p.selected = $selected
        $p.model = New-ModelList
        foreach ($m in $models) { $p.model.Add($m) }
        return $p
    }

    function Set-ConfigData($providers) {
        $list = New-ProviderList
        foreach ($p in $providers) { $list.Add($p) }
        [ShareRibbon.ConfigManager]::ConfigData = $list
        [ShareRibbon.ConfigManager]::ApplyConfigSnapshot()
    }

    function Snapshot() {
        return @{
            platform = [ShareRibbon.ConfigSettings]::platform
            url      = [ShareRibbon.ConfigSettings]::ApiUrl
            key      = [ShareRibbon.ConfigSettings]::ApiKey
            model    = [ShareRibbon.ConfigSettings]::ModelName
            mcpable  = [ShareRibbon.ConfigSettings]::mcpable
            fim      = [ShareRibbon.ConfigSettings]::fimSupported
            fimUrl   = [ShareRibbon.ConfigSettings]::fimUrl
            reasoning = [ShareRibbon.ConfigSettings]::ReasoningMode
        }
    }

    # 1. a selected provider plus a selected chat model must land in the snapshot
    $chat = New-Model "smoke-chat" $true $chatEnum $true $true "" "auto"
    $embedding = New-Model "smoke-embed" $true $embedEnum $false $false "" "default"
    $other = New-Model "other-chat" $true $chatEnum $false $false "" "default"
    $providers = @(
        (New-Provider "Smoke Other" "https://other.test/v1/chat/completions" "other-key" $false @($other)),
        (New-Provider "Smoke Main" "https://main.test/v1/chat/completions" "main-key" $true @($chat, $embedding))
    )
    Set-ConfigData $providers
    $s = Snapshot
    if ($s.platform -ne "Smoke Main") { throw "snapshot must follow the selected provider, got '$($s.platform)'." }
    if ($s.url -ne "https://main.test/v1/chat/completions") { throw "snapshot url wrong, got '$($s.url)'." }
    if ($s.key -ne "main-key") { throw "snapshot key wrong, got '$($s.key)'." }
    if ($s.model -ne "smoke-chat") { throw "snapshot must take the selected Chat model, got '$($s.model)'." }
    if (-not $s.mcpable) { throw "mcpable must come from the selected chat model." }
    if ($s.reasoning -ne "auto") { throw "reasoning mode must come from the selected chat model." }
    if ($s.fimUrl -ne "https://main.test/v1/chat/completions") { throw "an empty fimUrl must fall back to the provider url, got '$($s.fimUrl)'." }

    # 2. the reported defect: clearing url and key must not leave the previous endpoint live
    $providers[1].url = ""
    $providers[1].key = ""
    Set-ConfigData $providers
    $s = Snapshot
    if ($s.url -ne "") { throw "a cleared url must empty the snapshot, got '$($s.url)'." }
    if ($s.key -ne "") { throw "a cleared key must empty the snapshot, got '$($s.key)'." }

    # 3. nothing selected at all must also empty it, not keep the last provider
    foreach ($p in $providers) { $p.selected = $false }
    Set-ConfigData $providers
    $s = Snapshot
    if ($s.url -ne "" -or $s.key -ne "" -or $s.model -ne "") { throw "with no selected provider the snapshot must be empty." }

    # 4. the gate must then agree that this configuration is unusable
    $gate = New-Object ShareRibbon.ModelConfigGate
    $prompt = $gate.Evaluate([ShareRibbon.ConfigSettings]::ApiUrl, [ShareRibbon.ConfigSettings]::ApiKey, "chat")
    if ($null -eq $prompt) { throw "an emptied snapshot must block the chat send." }
    if ($prompt.Reason -ne "missing_api_key") { throw "emptied config must read as missing_api_key, got '$($prompt.Reason)'." }

    # 5. a selected provider whose chat model was deselected keeps the endpoint but loses the model,
    #    which is what routes the E3-01 wizard back to the key/model step
    $providers[1].url = "https://main.test/v1/chat/completions"
    $providers[1].key = "main-key"
    $providers[1].selected = $true
    $providers[1].model[0].selected = $false
    Set-ConfigData $providers
    $s = Snapshot
    if ($s.key -ne "main-key") { throw "the endpoint key must stay live while its chat model is deselected." }
    if ($s.model -ne "") { throw "a deselected chat model must clear ModelName, got '$($s.model)'." }
    if ($s.fim) { throw "fim flags must reset with the model." }

    # 6. a null ConfigData (window opened before any load) must not throw and must read as unconfigured
    [ShareRibbon.ConfigManager]::ConfigData = $null
    [ShareRibbon.ConfigManager]::ApplyConfigSnapshot()
    if ([ShareRibbon.ConfigSettings]::ApiUrl -ne "") { throw "null config data must clear the snapshot." }
    if ([ShareRibbon.ConfigSettings]::ApiKey -ne "") { throw "null config data must clear the key." }

    Write-Host "CONFIG SNAPSHOT SMOKE PASSED"
}
finally {
    Pop-Location
}
