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
$sqliteDll = Join-Path $bin "System.Data.SQLite.dll"
$migrationsDir = Join-Path $repoRoot "ShareRibbon\Storage\Migrations"
$snapshotPath = Join-Path $repoRoot "ShareRibbon\Storage\OfficeAiDbSchema.current.sql"

if (-not (Test-Path $sqliteDll)) {
    throw "System.Data.SQLite.dll not found. Run: .\scripts\smoke-db-schema-drift.ps1 -Build"
}
if (-not (Test-Path $migrationsDir)) {
    throw "Migrations dir not found: $migrationsDir"
}
if (-not (Test-Path $snapshotPath)) {
    throw "Schema snapshot not found: $snapshotPath"
}

function Invoke-SqlFile {
    param(
        [System.Data.SQLite.SQLiteConnection]$Connection,
        [string]$Path
    )

    # 与生产端 OfficeAiDatabase 的 File.ReadAllText 保持一致，按 UTF-8 读取，
    # 否则 PowerShell 5.1 会用系统 ANSI 码页解码，注释里的中文会破坏语句切分。
    $sql = [System.IO.File]::ReadAllText($Path)
    if ([string]::IsNullOrWhiteSpace($sql)) {
        return
    }

    $cmd = $Connection.CreateCommand()
    $cmd.CommandText = $sql
    [void]$cmd.ExecuteNonQuery()
}

function Invoke-SqlText {
    param(
        [System.Data.SQLite.SQLiteConnection]$Connection,
        [string]$Sql
    )

    $cmd = $Connection.CreateCommand()
    $cmd.CommandText = $Sql
    [void]$cmd.ExecuteNonQuery()
}

function New-SqliteConnection {
    param([string]$Path)

    $conn = New-Object System.Data.SQLite.SQLiteConnection("Data Source=$Path;Version=3;")
    $conn.Open()
    return $conn
}

function Get-SchemaShape {
    param([System.Data.SQLite.SQLiteConnection]$Connection)

    $tables = New-Object System.Collections.Generic.List[string]
    $tableCmd = $Connection.CreateCommand()
    $tableCmd.CommandText = "SELECT name FROM sqlite_master WHERE type = 'table' AND name NOT LIKE 'sqlite_%' ORDER BY name"
    $reader = $tableCmd.ExecuteReader()
    try {
        while ($reader.Read()) {
            $tables.Add($reader.GetString(0))
        }
    }
    finally {
        $reader.Dispose()
    }

    $columns = New-Object System.Collections.Generic.List[string]
    foreach ($table in $tables) {
        $colCmd = $Connection.CreateCommand()
        $colCmd.CommandText = "PRAGMA table_info([" + $table.Replace("]", "]]") + "])"
        $colReader = $colCmd.ExecuteReader()
        try {
            while ($colReader.Read()) {
                $name = [string]$colReader["name"]
                $type = ([string]$colReader["type"]).ToUpperInvariant()
                $notNull = [int]$colReader["notnull"]
                $pk = [int]$colReader["pk"]
                $columns.Add("$table|$name|$type|$notNull|$pk")
            }
        }
        finally {
            $colReader.Dispose()
        }
    }

    $indexes = New-Object System.Collections.Generic.List[string]
    $indexCmd = $Connection.CreateCommand()
    $indexCmd.CommandText = "SELECT tbl_name, name FROM sqlite_master WHERE type = 'index' AND name NOT LIKE 'sqlite_autoindex%' ORDER BY tbl_name, name"
    $indexReader = $indexCmd.ExecuteReader()
    try {
        while ($indexReader.Read()) {
            $indexes.Add($indexReader.GetString(0) + "|" + $indexReader.GetString(1))
        }
    }
    finally {
        $indexReader.Dispose()
    }

    return [pscustomobject]@{
        Tables = @($tables)
        Columns = @($columns | Sort-Object)
        Indexes = @($indexes | Sort-Object)
    }
}

function Compare-Set {
    param(
        [string]$Name,
        [string[]]$Expected,
        [string[]]$Actual
    )

    $missing = @($Expected | Where-Object { $Actual -notcontains $_ })
    $extra = @($Actual | Where-Object { $Expected -notcontains $_ })
    if ($missing.Count -gt 0 -or $extra.Count -gt 0) {
        throw "$Name drift detected. Missing=[$($missing -join '; ')], Extra=[$($extra -join '; ')]"
    }
}

$workDir = Join-Path ([System.IO.Path]::GetTempPath()) ("office-ai-schema-drift-" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $workDir | Out-Null
$migrationDb = Join-Path $workDir "migration.db"
$snapshotDb = Join-Path $workDir "snapshot.db"

Push-Location $bin
try {
    Add-Type -Path $sqliteDll

    $migrationConn = New-SqliteConnection $migrationDb
    try {
        Get-ChildItem -Path $migrationsDir -Filter "*.sql" | Sort-Object Name | ForEach-Object {
            Invoke-SqlFile -Connection $migrationConn -Path $_.FullName
        }
        Invoke-SqlText -Connection $migrationConn -Sql "CREATE TABLE IF NOT EXISTS data_migration_marker (migration_key TEXT PRIMARY KEY, applied_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')));"
        $migrationShape = Get-SchemaShape $migrationConn
    }
    finally {
        $migrationConn.Dispose()
    }

    $snapshotConn = New-SqliteConnection $snapshotDb
    try {
        Invoke-SqlFile -Connection $snapshotConn -Path $snapshotPath
        $snapshotShape = Get-SchemaShape $snapshotConn
    }
    finally {
        $snapshotConn.Dispose()
    }

    Compare-Set -Name "Tables" -Expected $snapshotShape.Tables -Actual $migrationShape.Tables
    Compare-Set -Name "Columns" -Expected $snapshotShape.Columns -Actual $migrationShape.Columns
    Compare-Set -Name "Indexes" -Expected $snapshotShape.Indexes -Actual $migrationShape.Indexes

    [pscustomobject]@{
        WorkDir = $workDir
        Tables = $snapshotShape.Tables.Count
        Columns = $snapshotShape.Columns.Count
        Indexes = $snapshotShape.Indexes.Count
        Drift = $false
    } | Format-List
}
finally {
    try {
        [System.Data.SQLite.SQLiteConnection]::ClearAllPools()
    }
    catch {
    }
    Pop-Location
}
