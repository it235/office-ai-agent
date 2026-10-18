param(
    [string]$VersionFile = (Join-Path $PSScriptRoot '..\Version.txt')
)

$ErrorActionPreference = 'Stop'
$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
$issues = New-Object System.Collections.Generic.List[string]

if (-not (Test-Path -LiteralPath $VersionFile)) {
    throw "Version file not found: $VersionFile"
}

$version = (Get-Content -LiteralPath $VersionFile -Raw).Trim()
if ($version -notmatch '^\d+\.\d+\.\d+$') {
    $issues.Add("Version.txt must use major.minor.patch format, got: $version")
}

$expectedAssemblyVersion = '2.0.0.0'
$expectedFileVersion = "$version.0"
$expectedInformationalVersion = $version

function Test-FileValue {
    param(
        [string]$RelativePath,
        [string]$Pattern,
        [string]$Expected,
        [string]$Label
    )

    $path = Join-Path $repoRoot $RelativePath
    if (-not (Test-Path -LiteralPath $path)) {
        $script:issues.Add("Missing file for ${Label}: $RelativePath")
        return
    }

    $content = Get-Content -LiteralPath $path -Raw -Encoding Default
    $match = [regex]::Match($content, $Pattern)
    if (-not $match.Success) {
        $script:issues.Add("Missing ${Label} in $RelativePath")
        return
    }

    $actual = $match.Groups[1].Value
    if ($actual -ne $Expected) {
        $script:issues.Add("${Label} mismatch in ${RelativePath}: expected $Expected, got $actual")
    }
}

$assemblyInfoFiles = @(
    'ShareRibbon\My Project\AssemblyInfo.vb',
    'ExcelAi\My Project\AssemblyInfo.vb',
    'WordAi\My Project\AssemblyInfo.vb',
    'PowerPointAi\My Project\AssemblyInfo.vb'
)

foreach ($relativePath in $assemblyInfoFiles) {
    Test-FileValue $relativePath '<Assembly:\s*AssemblyVersion\("([^"]+)"\)>' $expectedAssemblyVersion 'AssemblyVersion'
    Test-FileValue $relativePath '<Assembly:\s*AssemblyFileVersion\("([^"]+)"\)>' $expectedFileVersion 'AssemblyFileVersion'
    Test-FileValue $relativePath '<Assembly:\s*AssemblyInformationalVersion\("([^"]+)"\)>' $expectedInformationalVersion 'AssemblyInformationalVersion'
}

$vstoProjects = @(
    'ExcelAi\ExcelAi.vbproj',
    'WordAi\WordAi.vbproj',
    'PowerPointAi\PowerPointAi.vbproj'
)

foreach ($relativePath in $vstoProjects) {
    Test-FileValue $relativePath '<ApplicationVersion>([^<]+)</ApplicationVersion>' $expectedFileVersion 'ApplicationVersion'
}

Test-FileValue 'OfficeAgent\OfficeAgent.vdproj' '"ProductVersion"\s*=\s*"8:([^"]+)"' $version 'ProductVersion'

if ($issues.Count -gt 0) {
    foreach ($issue in $issues) {
        Write-Warning $issue
    }
    exit 1
}

Write-Host "Version audit passed: Version=$version, AssemblyVersion=$expectedAssemblyVersion, FileVersion=$expectedFileVersion"
