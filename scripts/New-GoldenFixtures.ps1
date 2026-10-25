param(
    [string]$OutputDir = ""
)

$ErrorActionPreference = "Stop"

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
. (Join-Path $PSScriptRoot "golden\GoldenFixtureBuilder.ps1")

if ([string]::IsNullOrWhiteSpace($OutputDir)) {
    $OutputDir = Join-Path $repoRoot "tests\golden\.generated"
}

Initialize-GoldenOpenXml -RepoRoot $repoRoot
$manifest = New-GoldenFixtures -SpecDir (Join-Path $repoRoot "tests\golden\tasks\fixtures") -OutputDir $OutputDir

$fixtures = @($manifest.fixtures)
Write-Host ""
Write-Host ("Golden fixtures generated: {0} -> {1}" -f $fixtures.Count, $OutputDir)
foreach ($fixture in $fixtures) {
    Write-Host ("  {0}  {1} ({2} bytes)" -f $fixture.contentSha256.Substring(0, 12), $fixture.file, $fixture.bytes)
}
Write-Host ("Manifest: {0}" -f (Join-Path $OutputDir "manifest.json"))
