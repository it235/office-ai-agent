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
if (-not (Test-Path (Join-Path $bin "ShareRibbon.dll"))) {
    throw "ShareRibbon.dll not found. Run: .\scripts\smoke-memory-pipeline.ps1 -Build"
}

Push-Location $bin
try {
    Add-Type -Path (Join-Path $bin "ShareRibbon.dll")

    [ShareRibbon.OfficeAiDatabase]::EnsureInitialized()
    $dbPath = [ShareRibbon.OfficeAiDatabase]::GetDatabasePath()

    $session = "smoke-" + [guid]::NewGuid().ToString("N")
    # A per-run marker: retrieving by it proves this run's memory reached the store, instead of
    # passing because some older memory happens to match a generic keyword.
    $marker = "smokemem" + [guid]::NewGuid().ToString("N")

    # The pipeline default is the LLM extractor, which needs a live key and a network. This gate
    # checks the queue -> extractor -> memory_item -> retrieval chain, so pin the deterministic
    # rule extractor and keep the result independent of model availability.
    [ShareRibbon.AgentMemoryPipelineService]::Extractor = New-Object ShareRibbon.RuleBasedMemoryExtractor

    # event_type must be 'chat_message': both the history read path and
    # GetUnprocessedConversationEventsBySession filter on it, so any other value is invisible.
    $userEvent = New-Object ShareRibbon.ConversationEventRecord
    $userEvent.SessionId = $session
    $userEvent.AppType = "Word"
    $userEvent.EventType = "chat_message"
    $userEvent.Role = "user"
    $userEvent.Content = "Remember my default official document body style ($marker): FangSong No.3, first-line indent 2 chars, justified alignment."
    $userId = [ShareRibbon.AgentMemoryRepository]::AppendConversationEvent($userEvent)

    $assistantEvent = New-Object ShareRibbon.ConversationEventRecord
    $assistantEvent.SessionId = $session
    $assistantEvent.AppType = "Word"
    $assistantEvent.EventType = "chat_message"
    $assistantEvent.Role = "assistant"
    $assistantEvent.Content = "I have saved your official document body formatting preference."
    $assistantId = [ShareRibbon.AgentMemoryRepository]::AppendConversationEvent($assistantEvent)

    $payload = New-Object Newtonsoft.Json.Linq.JObject
    $payload.Add("user_event_id", [Newtonsoft.Json.Linq.JValue]::CreateString($userId))
    $payload.Add("assistant_event_id", [Newtonsoft.Json.Linq.JValue]::CreateString($assistantId))

    $job = New-Object ShareRibbon.MemoryJobRecord
    $job.JobType = "extract_memory"
    # TargetId is the session id, matching MemoryTurnRecorder; the pipeline resolves the events
    # through GetUnprocessedConversationEventsBySession.
    $job.TargetId = $session
    $job.PayloadJson = $payload.ToString([Newtonsoft.Json.Formatting]::None)
    $job.Status = "pending"
    [ShareRibbon.AgentMemoryRepository]::EnqueueJob($job) | Out-Null

    [ShareRibbon.AgentMemoryPipelineService]::ProcessPendingJobsAsync(50).GetAwaiter().GetResult()

    $memories = [ShareRibbon.AgentMemoryRepository]::RetrieveMemoryItems($marker, "Word", "", "", 10)
    $skills = [ShareRibbon.AgentMemoryRepository]::RetrieveSkillRegistry("", "", "Word", 5)

    # ---- retrieval must respect TTL: an expired memory is invisible even when it matches ----
    $ttlMarker = $marker + "-expired"
    $expired = New-Object ShareRibbon.MemoryItemRecord
    $expired.Scope = "user"
    $expired.AppType = "Word"
    $expired.MemoryType = "fact"
    $expired.Content = "expired fixture memory $ttlMarker"
    $expired.Status = "active"
    $expired.ExpiresAt = (Get-Date).AddMinutes(-5).ToString("yyyy-MM-dd HH:mm:ss")
    [void][ShareRibbon.AgentMemoryRepository]::UpsertMemoryItem($expired)
    if (([ShareRibbon.AgentMemoryRepository]::RetrieveMemoryItems($ttlMarker, "Word", "", "", 10)).Count -ne 0) {
        throw "an expired memory item must not be retrieved"
    }

    # ---- queue semantics: idempotent enqueue, backoff retry, dead-letter ----
    $queueConn = New-Object System.Data.SQLite.SQLiteConnection([ShareRibbon.OfficeAiDatabase]::GetConnectionString())
    $queueConn.Open()
    try {
        function Get-QueueScalar([string]$sql) {
            $c = $script:queueConn.CreateCommand()
            $c.CommandText = $sql
            $v = $c.ExecuteScalar()
            $c.Dispose()
            return $v
        }

        function Set-QueueSql([string]$sql) {
            $c = $script:queueConn.CreateCommand()
            $c.CommandText = $sql
            [void]$c.ExecuteNonQuery()
            $c.Dispose()
        }

        $target = "smoke-queue-" + [Guid]::NewGuid().ToString("N")
        function New-QueueJob([string]$forTarget) {
            $j = New-Object ShareRibbon.MemoryJobRecord
            $j.JobType = "extract_memory"
            $j.TargetId = $forTarget
            $j.PayloadJson = "{}"
            $j.Status = "pending"
            return $j
        }

        $firstId = [ShareRibbon.AgentMemoryRepository]::EnqueueJob((New-QueueJob $target))
        $secondId = [ShareRibbon.AgentMemoryRepository]::EnqueueJob((New-QueueJob $target))
        if ($firstId -ne $secondId) {
            throw "duplicate enqueue must reuse the still-active job, got $firstId and $secondId"
        }
        if ([int](Get-QueueScalar "SELECT COUNT(1) FROM memory_job WHERE job_id = '$firstId' AND status IN ('pending','processing')") -ne 1) {
            throw "exactly one active job is expected per (job_type, target_id)"
        }

        $dupRejected = $false
        try {
            Set-QueueSql "INSERT INTO memory_job (job_id, job_type, target_id, status, payload_json) VALUES ('smoke-dup-row', 'extract_memory', '$target', 'pending', '{}')"
        } catch {
            $dupRejected = $true
        }
        if (-not $dupRejected) {
            throw "the partial unique index must reject a second active job for the same target"
        }

        Set-QueueSql "UPDATE memory_job SET status = 'processing', attempt_count = 1, next_run_at = NULL, updated_at = datetime('now','localtime') WHERE job_id = '$firstId'"
        if ([ShareRibbon.AgentMemoryRepository]::MarkJobFailed($firstId, "transient failure") -ne "pending") {
            throw "a job that still has retry budget must go back to pending"
        }
        if ([int](Get-QueueScalar "SELECT (next_run_at > datetime('now','localtime')) FROM memory_job WHERE job_id = '$firstId'") -ne 1) {
            throw "a retried job must be rescheduled with backoff"
        }

        Set-QueueSql "UPDATE memory_job SET status = 'processing', attempt_count = 3, updated_at = datetime('now','localtime') WHERE job_id = '$firstId'"
        if ([ShareRibbon.AgentMemoryRepository]::MarkJobFailed($firstId, "still failing") -ne "failed") {
            throw "a job out of retry budget must be dead-lettered"
        }

        $staleId = [ShareRibbon.AgentMemoryRepository]::EnqueueJob((New-QueueJob $target))
        Set-QueueSql "UPDATE memory_job SET status = 'processing', attempt_count = 3, updated_at = datetime('now','localtime','-40 minutes') WHERE job_id = '$staleId'"
        $requeuedId = [ShareRibbon.AgentMemoryRepository]::EnqueueJob((New-QueueJob ($target + "-requeue")))
        Set-QueueSql "UPDATE memory_job SET status = 'processing', attempt_count = 1, updated_at = datetime('now','localtime','-40 minutes') WHERE job_id = '$requeuedId'"
        [void][ShareRibbon.AgentMemoryRepository]::ReclaimStaleProcessingJobs(30)
        if ([string](Get-QueueScalar "SELECT status FROM memory_job WHERE job_id = '$staleId'") -ne "failed") {
            throw "a stale job out of retry budget must be dead-lettered, not left in processing"
        }
        if ([string](Get-QueueScalar "SELECT status FROM memory_job WHERE job_id = '$requeuedId'") -ne "pending") {
            throw "a stale job with retry budget left must return to pending"
        }

        Set-QueueSql "DELETE FROM memory_job WHERE job_id IN ('$firstId', '$staleId', '$requeuedId')"
    }
    finally {
        if ($queueConn) {
            $queueConn.Dispose()
        }
    }

    $result = [pscustomobject]@{
        Database = $dbPath
        Session = $session
        Marker = $marker
        UserEventId = $userId
        AssistantEventId = $assistantId
        MemoryCount = $memories.Count
        FirstMemory = $(if ($memories.Count -gt 0) { $memories[0].Content } else { "" })
        SkillRegistryRows = $skills.Count
        QueueIdempotent = $true
        QueueBackoffRetry = $true
        QueueDeadLetter = $true
        RetrievalHonorsTtl = $true
    }

    $result | Format-List

    if ($memories.Count -lt 1) {
        throw "Smoke failed: the inserted structured memory was not retrieved."
    }
    if (-not ($memories | Where-Object { $_.Content -like "*$marker*" })) {
        throw "the retrieved memory must be the one this run extracted, got '$($memories[0].Content)'"
    }
}
finally {
    # This gate writes into the developer's real memory store, so the fixture must leave no trace:
    # the session cascade removes the extracted session-scoped memory, and the marker delete catches
    # the expired item, which has scope=user and no session to cascade from.
    try {
        if ($session) {
            [void][ShareRibbon.ConversationRepository]::DeleteSessionCascade($session)
        }
        $clean = New-Object System.Data.SQLite.SQLiteConnection([ShareRibbon.OfficeAiDatabase]::GetConnectionString())
        $clean.Open()
        try {
            $cmd = $clean.CreateCommand()
            $cmd.CommandText = "DELETE FROM memory_embedding WHERE memory_id IN (SELECT memory_id FROM memory_item WHERE content LIKE '%smokemem%')"
            [void]$cmd.ExecuteNonQuery()
            $cmd.Dispose()
            $cmd = $clean.CreateCommand()
            $cmd.CommandText = "DELETE FROM memory_item WHERE content LIKE '%smokemem%'"
            [void]$cmd.ExecuteNonQuery()
            $cmd.Dispose()
        }
        finally {
            $clean.Dispose()
            [System.Data.SQLite.SQLiteConnection]::ClearAllPools()
        }
    }
    catch {
        Write-Warning "memory smoke fixture cleanup failed: $($_.Exception.Message)"
    }
    Pop-Location
}
