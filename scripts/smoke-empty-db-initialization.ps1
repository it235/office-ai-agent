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
    throw "ShareRibbon.dll not found. Run: .\scripts\smoke-empty-db-initialization.ps1 -Build"
}

$documents = [Environment]::GetFolderPath([Environment+SpecialFolder]::MyDocuments)
$dbDir = Join-Path $documents "OfficeAiAppData-Debug"
$dbPath = Join-Path $dbDir "office_ai.db"
if (-not (Test-Path $dbDir)) {
    New-Item -ItemType Directory -Path $dbDir | Out-Null
}

$backupDir = Join-Path $dbDir ("empty-db-smoke-backup-" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $backupDir | Out-Null

$dbFiles = @($dbPath, "$dbPath-wal", "$dbPath-shm")
$movedOriginals = @()

try {
    foreach ($file in $dbFiles) {
        if (Test-Path $file) {
            $target = Join-Path $backupDir (Split-Path -Leaf $file)
            Move-Item -LiteralPath $file -Destination $target
            $movedOriginals += [pscustomobject]@{ Source = $file; Backup = $target }
        }
    }

    $childScript = Join-Path $backupDir "verify-empty-db.ps1"
    $childContent = @'
param(
    [string]$Bin,
    [string]$Dll
)

$ErrorActionPreference = "Stop"

Push-Location $Bin
try {
    Add-Type -Path $Dll
    [ShareRibbon.OfficeAiDatabase]::EnsureInitialized()

    $conn = New-Object System.Data.SQLite.SQLiteConnection([ShareRibbon.OfficeAiDatabase]::GetConnectionString())
    $conn.Open()

    function Get-ScalarValue([string]$sql) {
        $c = $conn.CreateCommand()
        $c.CommandText = $sql
        $value = $c.ExecuteScalar()
        $c.Dispose()
        return $value
    }

    function Invoke-ExecSql([string]$sql) {
        $c = $conn.CreateCommand()
        $c.CommandText = $sql
        [void]$c.ExecuteNonQuery()
        $c.Dispose()
    }

    try {
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = "SELECT version FROM schema_version LIMIT 1"
        $version = [int]$cmd.ExecuteScalar()
        if ($version -ne 17) {
            throw "Expected schema version 17, got $version"
        }

        $requiredTables = @(
            "atomic_memory",
            "user_profile",
            "conversation",
            "conversation_event",
            "session_summary",
            "memory_item",
            "memory_embedding",
            "memory_job",
            "skills_registry",
            "agent_run",
            "agent_run_step",
            "agent_run_change",
            "data_migration_marker"
        )

        foreach ($table in $requiredTables) {
            $check = $conn.CreateCommand()
            $check.CommandText = "SELECT COUNT(1) FROM sqlite_master WHERE type = 'table' AND name = @name"
            [void]$check.Parameters.AddWithValue("@name", $table)
            if ([int]$check.ExecuteScalar() -ne 1) {
                throw "Expected table not found: $table"
            }
        }

        $changeColumns = @()
        $colCmd = $conn.CreateCommand()
        $colCmd.CommandText = "PRAGMA table_info(agent_run_change)"
        $colReader = $colCmd.ExecuteReader()
        while ($colReader.Read()) { $changeColumns += [string]$colReader["name"] }
        $colReader.Dispose()
        $colCmd.Dispose()
        foreach ($column in @("rolled_back", "undo_units", "doc_hash_before", "doc_hash_after")) {
            if ($changeColumns -notcontains $column) {
                throw "agent_run_change must have column $column after migration 016"
            }
        }

        # ---- migration 017 gives the session profile its document binding columns ----
        $profileColumns = @()
        $colCmd = $conn.CreateCommand()
        $colCmd.CommandText = "PRAGMA table_info(session_summary)"
        $colReader = $colCmd.ExecuteReader()
        while ($colReader.Read()) { $profileColumns += [string]$colReader["name"] }
        $colReader.Dispose()
        $colCmd.Dispose()
        foreach ($column in @("app_type", "document_id", "document_name", "custom_title", "updated_at")) {
            if ($profileColumns -notcontains $column) {
                throw "session_summary must have column $column after migration 017"
            }
        }
        $profileUnique = [int](Get-ScalarValue "SELECT COUNT(1) FROM sqlite_master WHERE type = 'index' AND name = 'idx_session_summary_session_unique'")
        if ($profileUnique -ne 1) {
            throw "session_summary must keep one profile row per session via idx_session_summary_session_unique"
        }

        # ---- chat history is single-sourced from conversation_event ----
        $sessionA = "smoke-a-" + [Guid]::NewGuid().ToString("N")
        $sessionB = "smoke-b-" + [Guid]::NewGuid().ToString("N")

        [ShareRibbon.ConversationRepository]::InsertMessage($sessionA, "user", "first question of session A")
        [ShareRibbon.ConversationRepository]::InsertMessage($sessionA, "assistant", "answer of session A")
        [ShareRibbon.ConversationRepository]::InsertMessage($sessionB, "user", "only question of session B")

        $history = [ShareRibbon.ConversationRepository]::GetMessagesBySession($sessionA)
        if ($history.Count -ne 2) {
            throw "chat history must read back both turns from conversation_event, got $($history.Count)"
        }
        if ($history[0].Role -ne "user" -or $history[1].Role -ne "assistant") {
            throw "chat history role order broken"
        }

        # two turns in one session must not grow a second profile row
        if ([int](Get-ScalarValue "SELECT COUNT(1) FROM session_summary WHERE session_id = '$sessionA'") -ne 1) {
            throw "the session profile must stay one row per session"
        }

        $entryA = [ShareRibbon.ConversationRepository]::GetRecentSessions(50) |
            Where-Object { $_.SessionId -eq $sessionA }
        if (-not $entryA) {
            throw "session list must aggregate from conversation_event: session A missing"
        }
        if ($entryA.Title -ne "first question of session A") {
            throw "session title must come from the first user message, got '$($entryA.Title)'"
        }

        # ---- cascade delete only owns session-scoped memories ----
        $eventA = [string](Get-ScalarValue "SELECT event_id FROM conversation_event WHERE session_id = '$sessionA' AND role = 'user' LIMIT 1")
        $eventB = [string](Get-ScalarValue "SELECT event_id FROM conversation_event WHERE session_id = '$sessionB' LIMIT 1")

        $owned = New-Object ShareRibbon.MemoryItemRecord
        $owned.SourceEventId = $eventA
        $owned.Scope = "session"
        $owned.AppType = "Word"
        $owned.MemoryType = "fact"
        $owned.Content = "session-owned memory " + [Guid]::NewGuid().ToString("N")
        $ownedId = [ShareRibbon.AgentMemoryRepository]::UpsertMemoryItem($owned)

        $shared = New-Object ShareRibbon.MemoryItemRecord
        $shared.SourceEventId = $eventA
        $shared.Scope = "user"
        $shared.AppType = "Word"
        $shared.MemoryType = "fact"
        $shared.Content = "cross-session memory " + [Guid]::NewGuid().ToString("N")
        $sharedId = [ShareRibbon.AgentMemoryRepository]::UpsertMemoryItem($shared)

        $other = New-Object ShareRibbon.MemoryItemRecord
        $other.SourceEventId = $eventB
        $other.Scope = "session"
        $other.AppType = "Word"
        $other.MemoryType = "fact"
        $other.Content = "other session memory " + [Guid]::NewGuid().ToString("N")
        $otherId = [ShareRibbon.AgentMemoryRepository]::UpsertMemoryItem($other)

        Invoke-ExecSql "INSERT INTO memory_embedding (embedding_id, memory_id, embedding_model, embedding_json) VALUES ('emb-owned', '$ownedId', 'smoke-model', '[]')"
        Invoke-ExecSql "INSERT INTO memory_embedding (embedding_id, memory_id, embedding_model, embedding_json) VALUES ('emb-shared', '$sharedId', 'smoke-model', '[]')"

        # a later session restating the same fact must not steal the evidence anchor
        $restated = New-Object ShareRibbon.MemoryItemRecord
        $restated.SourceEventId = $eventB
        $restated.Scope = "user"
        $restated.AppType = "Word"
        $restated.MemoryType = "fact"
        $restated.Content = $shared.Content
        [void][ShareRibbon.AgentMemoryRepository]::UpsertMemoryItemMerged($restated)
        $anchorAfterMerge = [string](Get-ScalarValue "SELECT IFNULL(source_event_id,'') FROM memory_item WHERE memory_id = '$sharedId'")
        if ($anchorAfterMerge -ne $eventA) {
            throw "memory evidence anchor must keep the first writer, got '$anchorAfterMerge'"
        }

        [void][ShareRibbon.ConversationRepository]::DeleteSessionCascade($sessionA)

        if ([int](Get-ScalarValue "SELECT COUNT(1) FROM conversation_event WHERE session_id = '$sessionA'") -ne 0) {
            throw "deleted session left conversation_event rows behind"
        }
        if ([int](Get-ScalarValue "SELECT COUNT(1) FROM memory_item WHERE memory_id = '$ownedId'") -ne 0) {
            throw "session-scoped memory of the deleted session must be removed"
        }
        if ([int](Get-ScalarValue "SELECT COUNT(1) FROM memory_embedding WHERE memory_id = '$ownedId'") -ne 0) {
            throw "embedding of a removed session memory must be removed"
        }
        if ([int](Get-ScalarValue "SELECT COUNT(1) FROM memory_item WHERE memory_id = '$sharedId'") -ne 1) {
            throw "cross-session memory must survive deleting one session"
        }
        if ([string](Get-ScalarValue "SELECT IFNULL(source_event_id,'') FROM memory_item WHERE memory_id = '$sharedId'") -ne "") {
            throw "shared memory anchor must be detached from the deleted session"
        }
        if ([int](Get-ScalarValue "SELECT COUNT(1) FROM memory_embedding WHERE memory_id = '$sharedId'") -ne 1) {
            throw "shared memory embedding must not be deleted by a session cascade"
        }
        if ([int](Get-ScalarValue "SELECT COUNT(1) FROM memory_item WHERE memory_id = '$otherId'") -ne 1) {
            throw "another session's memory must stay untouched"
        }
        [void][ShareRibbon.ConversationRepository]::DeleteSessionCascade($sessionB)

        # ---- migration 012 backfills legacy conversation rows ----
        $legacySession = "smoke-legacy-" + [Guid]::NewGuid().ToString("N")
        Invoke-ExecSql "INSERT INTO conversation (session_id, role, content, create_time) VALUES ('$legacySession', 'user', 'legacy question', '2026-01-01 10:00:00')"
        Invoke-ExecSql "INSERT INTO conversation (session_id, role, content, create_time) VALUES ('$legacySession', 'assistant', 'legacy answer', '2026-01-01 10:00:05')"
        Invoke-ExecSql "UPDATE schema_version SET version = 11"

        $runMigrations = [ShareRibbon.OfficeAiDatabase].GetMethod(
            "RunVersionedMigrations",
            [System.Reflection.BindingFlags]::NonPublic -bor [System.Reflection.BindingFlags]::Static)
        if (-not $runMigrations) {
            throw "RunVersionedMigrations not reflectable; migration harness shape changed"
        }
        [void]$runMigrations.Invoke($null, @($conn.PSObject.BaseObject))

        if ([int](Get-ScalarValue "SELECT COUNT(1) FROM conversation_event WHERE session_id = '$legacySession' AND event_type = 'chat_message'") -ne 2) {
            throw "migration 012 must backfill legacy conversation rows into conversation_event"
        }
        if ([int](Get-ScalarValue "SELECT version FROM schema_version LIMIT 1") -ne 17) {
            throw "schema version must be 17 after migrations 012 through 017"
        }
        $legacyEntry = [ShareRibbon.ConversationRepository]::GetRecentSessions(200) |
            Where-Object { $_.SessionId -eq $legacySession }
        if (-not $legacyEntry) {
            throw "upgraded legacy session must appear in the aggregated session list"
        }
        if ($legacyEntry.CreatedAt -ne "2026-01-01 10:00:00") {
            throw "legacy session timestamp must come from the earliest backfilled message, got '$($legacyEntry.CreatedAt)'"
        }
        if ($legacyEntry.DocumentId -ne "") {
            throw "an upgraded legacy session must start unbound instead of inventing a document"
        }
        [void][ShareRibbon.ConversationRepository]::DeleteSessionCascade($legacySession)
    }
    finally {
        if ($conn) {
            $conn.Dispose()
        }
        [System.Data.SQLite.SQLiteConnection]::ClearAllPools()
    }

    [pscustomobject]@{
        Database = [ShareRibbon.OfficeAiDatabase]::GetDatabasePath()
        SchemaVersion = 17
        RequiredTables = $requiredTables.Count
        HistorySingleSource = $true
        SessionProfileSchema = $true
        CascadeDeleteScoped = $true
        LegacyBackfill = $true
        EmptyInitialization = $true
    } | Format-List
}
finally {
    Pop-Location
}
'@
    Set-Content -LiteralPath $childScript -Value $childContent -Encoding UTF8

    & powershell -NoProfile -ExecutionPolicy Bypass -File $childScript -Bin $bin -Dll $dll
    if ($LASTEXITCODE -ne 0) {
        throw "Empty DB initialization child process failed, exit code: $LASTEXITCODE"
    }
}
finally {
    $generatedDir = Join-Path $backupDir "generated"
    New-Item -ItemType Directory -Path $generatedDir -Force | Out-Null

    foreach ($file in $dbFiles) {
        if (Test-Path $file) {
            Move-Item -LiteralPath $file -Destination (Join-Path $generatedDir (Split-Path -Leaf $file))
        }
    }

    foreach ($item in $movedOriginals) {
        if (Test-Path $item.Backup) {
            Move-Item -LiteralPath $item.Backup -Destination $item.Source
        }
    }

    Remove-Item -LiteralPath $backupDir -Recurse -Force
}
