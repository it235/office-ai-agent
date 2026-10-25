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
        if ($version -ne 19) {
            throw "Expected schema version 19, got $version"
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
        # migration 018 stores the replayable before-state snapshot for hosts without native undo
        if ($changeColumns -notcontains "restore_json") {
            throw "agent_run_change must have column restore_json after migration 018"
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

        # ---- E2-04 fixtures: legacy ownership shapes migration 019 has to resolve ----
        # (a) the "Unknown" sentinel an early Excel build used to store. It equals no host and is
        #     not the empty-string exemption, so the session was invisible in all three hosts.
        #     The session profile knows the document, so the extension is the ownership evidence.
        $sentinelSession = "smoke-sentinel-" + [Guid]::NewGuid().ToString("N")
        Invoke-ExecSql "INSERT INTO conversation_event (event_id, session_id, app_type, event_type, role, content, created_at) VALUES ('ev-a1', '$sentinelSession', 'Unknown', 'chat_message', 'user', 'sentinel question', '2026-01-02 09:00:00')"
        Invoke-ExecSql "INSERT INTO conversation_event (event_id, session_id, app_type, event_type, role, content, created_at) VALUES ('ev-a2', '$sentinelSession', 'unknown ', 'chat_message', 'assistant', 'sentinel answer', '2026-01-02 09:00:04')"
        Invoke-ExecSql "INSERT INTO session_summary (session_id, app_type, document_id, document_name, created_at, updated_at) VALUES ('$sentinelSession', '', '', 'C:\reports\budget.XLSX', '2026-01-02 09:00:00', '2026-01-02 09:00:04')"

        # (b) the newest trace carries the sentinel while an older trace knows the host: after
        #     normalization the older real host must win, and the profile must follow the events.
        $maskedSession = "smoke-masked-" + [Guid]::NewGuid().ToString("N")
        Invoke-ExecSql "INSERT INTO conversation_event (event_id, session_id, app_type, event_type, role, content, created_at) VALUES ('ev-b1', '$maskedSession', 'Unknown', 'chat_message', 'user', 'masked question', '2026-01-03 09:00:00')"
        Invoke-ExecSql "INSERT INTO agent_run (run_id, session_id, app_type, status, started_at) VALUES ('run-b-old', '$maskedSession', 'Excel', 'completed', '2026-01-03 09:00:01')"
        Invoke-ExecSql "INSERT INTO agent_run (run_id, session_id, app_type, status, started_at) VALUES ('run-b-new', '$maskedSession', 'Unknown', 'completed', '2026-01-03 10:00:00')"
        Invoke-ExecSql "INSERT INTO session_summary (session_id, app_type, document_id, document_name, created_at, updated_at) VALUES ('$maskedSession', '', '', '', '2026-01-03 09:00:00', '2026-01-03 10:00:00')"

        # (c) no trace and no document at all: must fall back to the empty-string exemption
        #     (generic, visible in every host) instead of being guessed or dropped.
        $ownerlessSession = "smoke-ownerless-" + [Guid]::NewGuid().ToString("N")
        Invoke-ExecSql "INSERT INTO conversation_event (event_id, session_id, app_type, event_type, role, content, created_at) VALUES ('ev-c1', '$ownerlessSession', 'Unknown', 'chat_message', 'user', 'ownerless question', '2026-01-04 09:00:00')"
        Invoke-ExecSql "INSERT INTO conversation_event (event_id, session_id, app_type, event_type, role, content, created_at) VALUES ('ev-c2', '$ownerlessSession', '', 'chat_message', 'assistant', 'ownerless answer', '2026-01-04 09:00:05')"

        # (d) a properly owned session must not be re-decided by the extension or a later document.
        $wordSession = "smoke-word-" + [Guid]::NewGuid().ToString("N")
        Invoke-ExecSql "INSERT INTO conversation_event (event_id, session_id, app_type, event_type, role, content, created_at) VALUES ('ev-d1', '$wordSession', 'Word', 'chat_message', 'user', 'word question', '2026-01-05 09:00:00')"
        Invoke-ExecSql "INSERT INTO session_summary (session_id, app_type, document_id, document_name, created_at, updated_at) VALUES ('$wordSession', 'Word', '', 'C:\notes\draft.xlsx', '2026-01-05 09:00:00', '2026-01-05 09:00:02')"

        # (e) the memory tables read app_type with the same exemption rule, so the same sentinel
        #     made old memories unreachable in every host. Measured on the installed database:
        #     123 of 231 atomic_memory rows carried "Unknown". Migration 019 must clear them.
        Invoke-ExecSql "INSERT INTO atomic_memory (timestamp, content, session_id, app_type, memory_type) VALUES (1767312000, 'sentinel memory', '$sentinelSession', 'Unknown', 'long_term')"
        Invoke-ExecSql "INSERT INTO atomic_memory (timestamp, content, session_id, app_type, memory_type) VALUES (1767312001, 'case variant memory', '$wordSession', 'word ', 'long_term')"
        Invoke-ExecSql "INSERT INTO memory_item (memory_id, scope, app_type, content, status) VALUES ('mem-a', 'user', 'Unknown', 'sentinel memory item', 'active')"
        Invoke-ExecSql "INSERT INTO memory_item (memory_id, scope, app_type, content, status) VALUES ('mem-b', 'user', 'PowerPoint', 'owned memory item', 'active')"

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
        if ([int](Get-ScalarValue "SELECT version FROM schema_version LIMIT 1") -ne 19) {
            throw "schema version must be 19 after migrations 012 through 019"
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

        function Get-SessionOwner([string]$sessionId) {
            $distinct = [int](Get-ScalarValue "SELECT COUNT(DISTINCT IFNULL(app_type, '')) FROM conversation_event WHERE session_id = '$sessionId'")
            if ($distinct -ne 1) {
                throw "session $sessionId must resolve to exactly one owner, found $distinct distinct app_type values"
            }
            return [string](Get-ScalarValue "SELECT app_type FROM conversation_event WHERE session_id = '$sessionId' LIMIT 1")
        }

        function Get-ProfileOwner([string]$sessionId) {
            return [string](Get-ScalarValue "SELECT app_type FROM session_summary WHERE session_id = '$sessionId'")
        }

        # ---- E2-04: migration 019 makes ownership decidable again ----
        foreach ($tableName in @('conversation_event', 'agent_run', 'agent_run_change', 'session_summary', 'atomic_memory', 'memory_item')) {
            $leftover = [int](Get-ScalarValue "SELECT COUNT(1) FROM $tableName WHERE IFNULL(app_type, '') <> '' AND app_type NOT IN ('Word', 'Excel', 'PowerPoint')")
            if ($leftover -ne 0) {
                throw "migration 019 must leave only the three hosts or the empty exemption in $tableName, found $leftover row(s)"
            }
        }

        # The memory tables share the same read rule, so the sentinel had to go there too: the row
        # must be recallable in every host through the real read path, not merely look normalised.
        $sentinelMemoryApp = [string](Get-ScalarValue "SELECT app_type FROM atomic_memory WHERE content = 'sentinel memory'")
        if ($sentinelMemoryApp -ne '') {
            throw "an Unknown atomic_memory row must fall back to the generic exemption, got '$sentinelMemoryApp'"
        }
        $caseVariantApp = [string](Get-ScalarValue "SELECT app_type FROM atomic_memory WHERE content = 'case variant memory'")
        if ($caseVariantApp -ne 'Word') {
            throw "a case and whitespace variant of a real host must be canonicalised, got '$caseVariantApp'"
        }
        if ([string](Get-ScalarValue "SELECT app_type FROM memory_item WHERE memory_id = 'mem-a'") -ne '') {
            throw "an Unknown memory_item row must fall back to the generic exemption"
        }
        $ownedMemoryItem = [string](Get-ScalarValue "SELECT app_type FROM memory_item WHERE memory_id = 'mem-b'")
        if ($ownedMemoryItem -ne 'PowerPoint') {
            throw "a properly owned memory_item row must stay owned, got '$ownedMemoryItem'"
        }
        foreach ($hostName in @('Word', 'Excel', 'PowerPoint')) {
            $recalled = [ShareRibbon.MemoryRepository]::ListAtomicMemories(200, 0, $hostName) |
                Where-Object { $_.Content -eq 'sentinel memory' }
            if (-not $recalled) {
                throw "a normalised atomic_memory row must be recallable in $hostName through ListAtomicMemories"
            }
        }

        if ((Get-SessionOwner $sentinelSession) -ne 'Excel') {
            throw "the Unknown-era session must be resolved to Excel from its document extension, got '$(Get-SessionOwner $sentinelSession)'"
        }
        if ((Get-ProfileOwner $sentinelSession) -ne 'Excel') {
            throw "the session profile must carry the resolved owner, otherwise host grouping cannot use it"
        }
        if ((Get-SessionOwner $maskedSession) -ne 'Excel') {
            throw "an older real trace must win over a newer Unknown one after normalization"
        }
        if ((Get-ProfileOwner $maskedSession) -ne 'Excel') {
            throw "the profile row must follow the owner resolved from its events"
        }
        if ((Get-SessionOwner $ownerlessSession) -ne '') {
            throw "a session with no evidence must stay the generic exemption instead of an invented host"
        }
        if ((Get-SessionOwner $wordSession) -ne 'Word') {
            throw "an already owned session must not be re-decided by the document extension"
        }

        # Visibility is the acceptance criterion: no legacy session may be invisible in every host,
        # and a resolved owner must not stay visible in the wrong host either.
        foreach ($hostName in @('Word', 'Excel', 'PowerPoint')) {
            $ownerlessVisible = [ShareRibbon.ConversationRepository]::GetRecentSessions(200, $hostName) |
                Where-Object { $_.SessionId -eq $ownerlessSession }
            if (-not $ownerlessVisible) {
                throw "an ownerless legacy session must stay visible in $hostName after the upgrade"
            }
        }
        $excelList = [ShareRibbon.ConversationRepository]::GetRecentSessions(200, 'Excel')
        $wordList = [ShareRibbon.ConversationRepository]::GetRecentSessions(200, 'Word')
        foreach ($ownedSession in @($sentinelSession, $maskedSession)) {
            if (-not ($excelList | Where-Object { $_.SessionId -eq $ownedSession })) {
                throw "$ownedSession resolved to Excel and must appear in the Excel list"
            }
            if ($wordList | Where-Object { $_.SessionId -eq $ownedSession }) {
                throw "$ownedSession resolved to Excel and must not also appear in Word"
            }
        }
        if (-not ($wordList | Where-Object { $_.SessionId -eq $wordSession })) {
            throw "the Word-owned session must stay in the Word list"
        }

        # A half-failed upgrade must converge, so 019 has to be safe to re-run. This also proves the
        # inline fallback in OfficeAiDatabase and the shipped file agree.
        $migration019 = Join-Path $Bin 'Storage\Migrations\019_session_ownership_normalization.sql'
        if (-not (Test-Path $migration019)) {
            throw "migration 019 is not deployed next to the DLL: $migration019"
        }
        Invoke-ExecSql ([System.IO.File]::ReadAllText($migration019))
        foreach ($expected in @(@($sentinelSession, 'Excel'), @($maskedSession, 'Excel'), @($ownerlessSession, ''), @($wordSession, 'Word'))) {
            if ((Get-SessionOwner $expected[0]) -ne $expected[1]) {
                throw "migration 019 must be idempotent; re-running it changed the owner of $($expected[0])"
            }
        }

        foreach ($fixtureSession in @($sentinelSession, $maskedSession, $ownerlessSession, $wordSession)) {
            Invoke-ExecSql "DELETE FROM agent_run WHERE session_id = '$fixtureSession'"
            [void][ShareRibbon.ConversationRepository]::DeleteSessionCascade($fixtureSession)
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
        SchemaVersion = 19
        RequiredTables = $requiredTables.Count
        HistorySingleSource = $true
        SessionProfileSchema = $true
        CascadeDeleteScoped = $true
        LegacyBackfill = $true
        SessionOwnershipNormalized = $true
        MemoryOwnershipNormalized = $true
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
