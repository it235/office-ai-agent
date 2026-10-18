param(
    [string]$VersionFile = (Join-Path $PSScriptRoot '..\Version.txt'),
    [string]$Version
)

$ErrorActionPreference = 'Stop'
$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')

if ([string]::IsNullOrWhiteSpace($Version)) {
    $Version = (Get-Content -LiteralPath $VersionFile -Raw).Trim()
}

if ($Version -notmatch '^\d+\.\d+\.\d+$') {
    throw "Version must use major.minor.patch format, got: $Version"
}

$assemblyVersion = '2.0.0.0'
$fileVersion = "$Version.0"
$informationalVersion = $Version

$assemblyInfoFiles = @(
    'ShareRibbon\My Project\AssemblyInfo.vb',
    'ExcelAi\My Project\AssemblyInfo.vb',
    'WordAi\My Project\AssemblyInfo.vb',
    'PowerPointAi\My Project\AssemblyInfo.vb'
)

foreach ($relativePath in $assemblyInfoFiles) {
    $path = Join-Path $repoRoot $relativePath
    $content = Get-Content -LiteralPath $path -Raw -Encoding Default
    $content = [regex]::Replace($content, 'AssemblyVersion\(".*?"\)', "AssemblyVersion(""$assemblyVersion"")")
    $content = [regex]::Replace($content, 'AssemblyFileVersion\(".*?"\)', "AssemblyFileVersion(""$fileVersion"")")

    if ($content -match 'AssemblyInformationalVersion\(".*?"\)') {
        $content = [regex]::Replace($content, 'AssemblyInformationalVersion\(".*?"\)', "AssemblyInformationalVersion(""$informationalVersion"")")
    } else {
        $content = [regex]::Replace(
            $content,
            '(<Assembly:\s*AssemblyFileVersion\(".*?"\)>\s*)',
            "`$1`r`n<Assembly: AssemblyInformationalVersion(""$informationalVersion"")>`r`n")
    }

    Set-Content -LiteralPath $path -Value $content -Encoding Default -NoNewline
    Write-Host "Updated $relativePath"
}

$vstoProjects = @(
    'ExcelAi\ExcelAi.vbproj',
    'WordAi\WordAi.vbproj',
    'PowerPointAi\PowerPointAi.vbproj'
)

foreach ($relativePath in $vstoProjects) {
    $path = Join-Path $repoRoot $relativePath
    $content = Get-Content -LiteralPath $path -Raw -Encoding Default
    if ($content -match '<ApplicationVersion>.*?</ApplicationVersion>') {
        $content = [regex]::Replace($content, '<ApplicationVersion>.*?</ApplicationVersion>', "<ApplicationVersion>$fileVersion</ApplicationVersion>")
    } else {
        $content = [regex]::Replace($content, '(<GenerateManifests>true</GenerateManifests>\s*)', "`$1    <ApplicationVersion>$fileVersion</ApplicationVersion>`r`n")
    }
    Set-Content -LiteralPath $path -Value $content -Encoding Default -NoNewline
    Write-Host "Updated $relativePath"
}

$vdproj = Join-Path $repoRoot 'OfficeAgent\OfficeAgent.vdproj'
if (Test-Path -LiteralPath $vdproj) {
    $content = Get-Content -LiteralPath $vdproj -Raw -Encoding Default
    $content = [regex]::Replace($content, '"ProductVersion"\s*=\s*"8:[^"]+"', """ProductVersion"" = ""8:$Version""")
    Set-Content -LiteralPath $vdproj -Value $content -Encoding Default -NoNewline
    Write-Host 'Updated OfficeAgent\OfficeAgent.vdproj'
}

Write-Host "Version update complete: AssemblyVersion=$assemblyVersion, FileVersion=$fileVersion, ProductVersion=$Version"
