param(
    [string]$Configuration = "Debug"
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$bin = Join-Path $repoRoot "ShareRibbon\bin\$Configuration"
if (-not (Test-Path (Join-Path $bin "ShareRibbon.dll"))) {
    throw "ShareRibbon.dll not found. Build first: scripts\build-code-projects.ps1"
}

Push-Location $bin
try {
    # E1-07 assertions build Newtonsoft JObject payloads by hand, so load the dependency explicitly
    Add-Type -Path (Join-Path $bin "Newtonsoft.Json.dll")
    Add-Type -Path (Join-Path $bin "ShareRibbon.dll")

    [ShareRibbon.OfficeAiDatabase]::EnsureInitialized()

    # ---- change log aggregation: only observed write/code_execution results ----
    $log = New-Object ShareRibbon.Agent.AgentRunChangeLog

    $writeObservation = @{
        kind = "write"
        summary = "wrote A1"
        changed = $true
        targetRefs = @("Sheet1!A1")
    }
    $writeResult = [ShareRibbon.Agent.ToolResult]::Succeed("ApplyFormula", "wrote A1", $null, $writeObservation, "undo-1")
    $record1 = $log.AddFromStep("ApplyFormula", 1, "write formula", $writeResult, "undo-1", $true)
    if ($null -eq $record1) { throw "observed write result must be recorded" }

    $readObservation = @{ kind = "read"; summary = "read A1"; changed = $false }
    $readResult = [ShareRibbon.Agent.ToolResult]::Succeed("ReadData", "read A1", $null, $readObservation)
    if ($null -ne $log.AddFromStep("ReadData", 2, "read data", $readResult, "", $false)) {
        throw "read observations must not enter the change log"
    }

    $noChangedObservation = @{ kind = "write"; summary = "unknown" }
    $noChangedResult = [ShareRibbon.Agent.ToolResult]::Succeed("InsertText", "unknown", $null, $noChangedObservation)
    if ($null -ne $log.AddFromStep("InsertText", 3, "no evidence", $noChangedResult, "", $false)) {
        throw "write observations without a boolean changed flag must not enter the change log"
    }

    $partialObservation = @{
        kind = "write"
        summary = "half applied"
        changed = $true
        targetRefs = @("Slide 2")
        rollback = @{ attempted = $false; succeeded = $false; partialApply = $true; error = "no rollback" }
    }
    $partialResult = [ShareRibbon.Agent.ToolResult]::Failed("FormatSlide", "half applied", $null, "PARTIAL_APPLY", "half applied", "", $false, $partialObservation)
    $record2 = $log.AddFromStep("FormatSlide", 4, "format slide", $partialResult, "", $false)
    if ($null -eq $record2) { throw "failed but observed write must be recorded" }
    if (-not $record2.PartialApply) { throw "failed write with changed=true must be marked as partial apply" }
    if ($record2.RollbackError -ne "no rollback") { throw "rollback error must be carried into the change record" }

    $scriptObservation = @{ kind = "code_execution"; summary = "ran script"; changed = $true; targetRefs = @("Office:ActiveDocument") }
    $scriptResult = [ShareRibbon.Agent.ToolResult]::Succeed("ExecuteJavaScript", "ran script", $null, $scriptObservation)
    $record3 = $log.AddFromStep("ExecuteJavaScript", 5, "run script", $scriptResult, "", $false)
    if ($null -eq $record3) { throw "code_execution observations must be recorded" }

    if ($log.Entries.Count -ne 3) { throw "expected 3 change records, got $($log.Entries.Count)" }
    if ($log.ChangedCount -ne 3) { throw "expected 3 changed records, got $($log.ChangedCount)" }

    # ---- persistence: run-scoped read back ----
    $runId = "smoke-run-" + [Guid]::NewGuid().ToString("N")
    $sessionId = "smoke-session-" + [Guid]::NewGuid().ToString("N")
    [ShareRibbon.AgentChangeLogRepository]::SaveRun($runId, $sessionId, "Excel", $log)

    $loaded = [ShareRibbon.AgentChangeLogRepository]::GetByRun($runId)
    if ($loaded.Count -ne 3) { throw "expected 3 persisted change records, got $($loaded.Count)" }
    if ($loaded[0].ToolId -ne "ApplyFormula" -or $loaded[2].ToolId -ne "ExecuteJavaScript") {
        throw "persisted change records must keep step order"
    }
    if (-not $loaded[0].Changed -or $loaded[0].UndoAvailable -or $loaded[0].UndoPointId -ne "undo-1") {
        throw "Excel writes must keep their undo point but must not claim a host undo unit"
    }
    if (-not $loaded[1].PartialApply) { throw "partial apply flag must survive persistence" }
    if ($loaded[0].TargetRefs.Count -ne 1 -or $loaded[0].TargetRefs[0].ToString() -ne "Sheet1!A1") {
        throw "target refs must survive persistence"
    }

    # ---- run undo scope: only residual writes with host units and end-to-end hashes count ----
    $undoLog = New-Object ShareRibbon.Agent.AgentRunChangeLog
    $undoObservation = @{
        kind = "write"
        summary = "inserted two paragraphs"
        changed = $true
        undoUnits = 2
        before = @{ documentTextHash = "hash-before" }
        after = @{ documentTextHash = "hash-after" }
        items = @(
            @{ changed = $true; undoUnits = 1; before = @{ documentTextHash = "hash-before" }; after = @{ documentTextHash = "hash-mid" } },
            @{ changed = $true; undoUnits = 1; before = @{ documentTextHash = "hash-mid" }; after = @{ documentTextHash = "hash-after" } }
        )
    }
    $undoResult = [ShareRibbon.Agent.ToolResult]::Succeed("WordCommands", "inserted", $null, $undoObservation, "undo-batch")
    $undoRecord = $undoLog.AddFromStep("WordCommands", 1, "insert twice", $undoResult, "undo-batch", $true)
    # the aggregate already reports the batch total; summing items again would undo one action too many
    if ($undoRecord.UndoUnits -ne 2) { throw "batch undo units must keep the aggregate total, got $($undoRecord.UndoUnits)" }
    if ($undoRecord.DocHashBefore -ne "hash-before" -or $undoRecord.DocHashAfter -ne "hash-after") {
        throw "batch document hashes must span the whole batch"
    }
    if (-not $undoLog.FullyUndoable) { throw "a fully observed Word batch must be undoable" }

    $fallbackLog = New-Object ShareRibbon.Agent.AgentRunChangeLog
    $fallbackObservation = @{
        kind = "write"
        summary = "items only"
        changed = $true
        before = @{ documentTextHash = "hash-before" }
        after = @{ documentTextHash = "hash-after" }
        items = @(
            @{ changed = $true; undoUnits = 1 },
            @{ changed = $true; undoUnits = 1 },
            @{ changed = $false; undoUnits = 5 }
        )
    }
    $fallbackResult = [ShareRibbon.Agent.ToolResult]::Succeed("WordCommands", "items only", $null, $fallbackObservation, "undo-items")
    $fallbackRecord = $fallbackLog.AddFromStep("WordCommands", 2, "items only", $fallbackResult, "undo-items", $true)
    if ($fallbackRecord.UndoUnits -ne 2) { throw "batch undo units must fall back to committed items, got $($fallbackRecord.UndoUnits)" }

    $rolledLog = New-Object ShareRibbon.Agent.AgentRunChangeLog
    $rolledObservation = @{
        kind = "write"
        summary = "rolled back"
        changed = $true
        undoUnits = 1
        rollback = @{ attempted = $true; succeeded = $true; error = "" }
        before = @{ documentTextHash = "hash-before" }
        after = @{ documentTextHash = "hash-after" }
    }
    $rolledResult = [ShareRibbon.Agent.ToolResult]::Failed("InsertText", "rolled back", $null, "UNKNOWN", "rolled back", "", $false, $rolledObservation)
    $rolledRecord = $rolledLog.AddFromStep("InsertText", 1, "failed then rolled back", $rolledResult, "", $false)
    if (-not $rolledRecord.RolledBack) { throw "a verified rollback must be marked as rolled back" }
    if ($rolledRecord.PartialApply) { throw "a fully rolled-back failure must not count as partial apply" }
    if ($rolledLog.WrittenCount -ne 0 -or $rolledLog.FullyUndoable) {
        throw "rolled-back writes must not remain as residual changes"
    }

    $undoRunId = "smoke-undo-" + [Guid]::NewGuid().ToString("N")
    [ShareRibbon.AgentChangeLogRepository]::SaveRun($undoRunId, $sessionId, "Word", $undoLog)
    $scope = [ShareRibbon.AgentChangeLogRepository]::GetUndoScopeByRun($undoRunId)
    if (-not $scope.FullyUndoable -or $scope.UndoUnits -ne 2 -or $scope.WrittenCount -ne 1) {
        throw "persisted Word run must expose a 2-unit undo scope"
    }
    if ($scope.BaselineHash -ne "hash-before" -or $scope.ExpectedHash -ne "hash-after") {
        throw "undo scope must keep the run start and end document hashes"
    }

    [ShareRibbon.AgentChangeLogRepository]::MarkRunRolledBack($undoRunId)
    $afterUndo = [ShareRibbon.AgentChangeLogRepository]::GetUndoScopeByRun($undoRunId)
    if ($afterUndo.WrittenCount -ne 0 -or $afterUndo.FullyUndoable) {
        throw "a rolled-back run must no longer offer undo"
    }

    # ---- E1-07 restore snapshots: hosts without a whole-stack undo replay a captured region instead ----
    # The snapshot travels on ToolResult.RestoreSnapshot, never inside Observation: observation text is
    # fed to the planner prompt and persisted in the run trace, so document content must not ride along.
    $restoreRunId = "smoke-restore-" + [Guid]::NewGuid().ToString("N")
    $restoreLog = New-Object ShareRibbon.Agent.AgentRunChangeLog
    $regionObservation = @{
        kind = "write"
        summary = "wrote a region"
        changed = $true
        targetRefs = @("Sheet1!A1:C3")
    }
    $restoreNode = [ShareRibbon.Agent.AgentRunChangeLog]::BuildRestoreNode(
        "excel_range_state",
        ([Newtonsoft.Json.Linq.JObject]::Parse('{"address":"Sheet1!A1:C3","secretPayload":"cell-content-must-not-leak"}')))
    if ($null -eq $restoreNode) { throw "a small reversible restore node must be accepted" }
    if ($restoreNode["reversible"].ToString() -ne "True") { throw "restore nodes must declare themselves reversible" }

    $regionResult = [ShareRibbon.Agent.ToolResult]::Succeed("ApplyFormula", "wrote a region", $null, $regionObservation)
    $regionResult.RestoreSnapshot = $restoreNode
    $regionRecord = $restoreLog.AddFromStep("ApplyFormula", 1, "write region", $regionResult, "", $false)
    if ($null -eq $regionRecord -or -not $regionRecord.HasRestore) {
        throw "a write carrying a restore snapshot must record it"
    }
    if ($regionRecord.NativeUndoUnit) { throw "an Excel-style write has no host undo unit" }
    if ($restoreLog.NativeUndoable) { throw "restore-only writes must not be reported as native undo" }
    if ($restoreLog.RestoreCount -ne 1 -or -not $restoreLog.FullyUndoable) {
        throw "a run whose every residual write has a snapshot must be undoable by replay"
    }
    if ($regionRecord.ToJson().ToString().Contains("cell-content-must-not-leak")) {
        throw "change records must only report that a snapshot exists, never the snapshot body"
    }

    # all-or-nothing: one replayable step plus one unreplayable step is not a safe one-click undo
    $mixedRunId = "smoke-restore-mixed-" + [Guid]::NewGuid().ToString("N")
    $halfNode = [Newtonsoft.Json.Linq.JObject]::Parse('{"kind":"excel_range_state","state":{"address":"A1"}}')
    $halfResult = [ShareRibbon.Agent.ToolResult]::Succeed("FormatRange", "formatted", $null, @{
        kind = "write"; summary = "formatted a range"; changed = $true; targetRefs = @("Sheet1!A1") })
    $halfResult.RestoreSnapshot = $halfNode
    [void]$restoreLog.AddFromStep("FormatRange", 2, "format without reversibility", $halfResult, "", $false)
    if ($restoreLog.RestoreCount -ne 1 -or $restoreLog.WrittenCount -ne 2 -or $restoreLog.FullyUndoable) {
        throw "a run with one unrestorable write must not claim a complete undo"
    }
    [ShareRibbon.AgentChangeLogRepository]::SaveRun($mixedRunId, $sessionId, "Excel", $restoreLog)
    $mixedScope = [ShareRibbon.AgentChangeLogRepository]::GetUndoScopeByRun($mixedRunId)
    if ($mixedScope.FullyUndoable -or $mixedScope.NativeUndoable) {
        throw "persisted mixed run must stay undo-refused"
    }
    if (-not $mixedScope.Reason) { throw "a refused undo must carry a reason for the user" }

    # persistence: the snapshot body survives only in agent_run_change.restore_json
    $persistLog = New-Object ShareRibbon.Agent.AgentRunChangeLog
    [void]$persistLog.AddFromStep("ApplyFormula", 1, "write region", $regionResult, "", $false)
    [ShareRibbon.AgentChangeLogRepository]::SaveRun($restoreRunId, $sessionId, "Excel", $persistLog)
    $restored = [ShareRibbon.AgentChangeLogRepository]::GetByRun($restoreRunId)
    if ($restored.Count -ne 1 -or -not $restored[0].HasRestore) {
        throw "restore snapshots must survive persistence"
    }
    if ($restored[0].RestoreJson -notlike "*cell-content-must-not-leak*") {
        throw "the persisted snapshot must keep the region payload for replay"
    }
    $snapshots = [ShareRibbon.AgentChangeLogRepository]::GetRestoreSnapshotsByRun($restoreRunId)
    if ($snapshots.Count -ne 1 -or $snapshots[0]["kind"].ToString() -ne "excel_range_state") {
        throw "undo replay must read back the restore node by kind"
    }
    $restoreScope = [ShareRibbon.AgentChangeLogRepository]::GetUndoScopeByRun($restoreRunId)
    if (-not $restoreScope.FullyUndoable -or $restoreScope.NativeUndoable -or $restoreScope.RestoreCount -ne 1) {
        throw "a persisted Excel run must expose a restore-only undo scope"
    }

    # size gate: a snapshot bigger than the contract limit is dropped, not silently truncated
    $oversize = [Newtonsoft.Json.Linq.JObject]::Parse('{"text":"' + ("x" * 61000) + '"}')
    if ($null -ne [ShareRibbon.Agent.AgentRunChangeLog]::BuildRestoreNode("excel_range_state", $oversize)) {
        throw "restore snapshots over the size limit must be refused"
    }

    # ---- session scope: history is filtered by host, legacy rows stay visible ----
    $excelSession = "smoke-excel-" + [Guid]::NewGuid().ToString("N")
    $wordSession = "smoke-word-" + [Guid]::NewGuid().ToString("N")
    $legacySession = "smoke-legacy-" + [Guid]::NewGuid().ToString("N")

    [ShareRibbon.ConversationRepository]::InsertMessage($excelSession, "user", "excel question", $false, "Excel")
    [ShareRibbon.ConversationRepository]::InsertMessage($wordSession, "user", "word question", $false, "Word")
    [ShareRibbon.ConversationRepository]::InsertMessage($legacySession, "user", "legacy question", $false, "")

    $excelSessions = [ShareRibbon.ConversationRepository]::GetRecentSessions(200, "Excel")
    $wordSessions = [ShareRibbon.ConversationRepository]::GetRecentSessions(200, "Word")

    if (-not ($excelSessions | Where-Object { $_.SessionId -eq $excelSession })) {
        throw "Excel session must appear in Excel history"
    }
    if ($excelSessions | Where-Object { $_.SessionId -eq $wordSession }) {
        throw "Word session must not leak into Excel history"
    }
    if (-not ($wordSessions | Where-Object { $_.SessionId -eq $wordSession })) {
        throw "Word session must appear in Word history"
    }
    if (-not ($excelSessions | Where-Object { $_.SessionId -eq $legacySession })) {
        throw "legacy rows without app_type must stay visible to every host"
    }

    # ---- E2-04 write boundary: an app_type that is neither a host nor the empty exemption
    # makes the session invisible in every host, so InsertMessage must normalise it to '' ----
    $sentinelSession = "smoke-sentinel-" + [Guid]::NewGuid().ToString("N")
    [ShareRibbon.ConversationRepository]::InsertMessage($sentinelSession, "user", "sentinel question", $false, "Unknown ")
    $excelAfter = [ShareRibbon.ConversationRepository]::GetRecentSessions(200, "Excel")
    $wordAfter = [ShareRibbon.ConversationRepository]::GetRecentSessions(200, "Word")
    $pptAfter = [ShareRibbon.ConversationRepository]::GetRecentSessions(200, "PowerPoint")
    foreach ($hostSessions in @($excelAfter, $wordAfter, $pptAfter)) {
        if (-not ($hostSessions | Where-Object { $_.SessionId -eq $sentinelSession })) {
            throw "a rejected app_type must fall back to the generic exemption and stay visible in every host"
        }
    }

    $sentinelProbe = New-Object System.Data.SQLite.SQLiteConnection([ShareRibbon.OfficeAiDatabase]::GetConnectionString())
    $sentinelProbe.Open()
    try {
        $cmd = $sentinelProbe.CreateCommand()
        $cmd.CommandText = "SELECT IFNULL(app_type,'?') FROM conversation_event WHERE session_id = @sid"
        [void]$cmd.Parameters.AddWithValue("@sid", $sentinelSession)
        $storedEventApp = [string]$cmd.ExecuteScalar()
        $cmd.Dispose()

        $cmd = $sentinelProbe.CreateCommand()
        $cmd.CommandText = "SELECT IFNULL(app_type,'?') FROM session_summary WHERE session_id = @sid"
        [void]$cmd.Parameters.AddWithValue("@sid", $sentinelSession)
        $storedProfileApp = [string]$cmd.ExecuteScalar()
        $cmd.Dispose()
    }
    finally {
        $sentinelProbe.Dispose()
        [System.Data.SQLite.SQLiteConnection]::ClearAllPools()
    }
    if ($storedEventApp -ne "") { throw "InsertMessage must store the normalised host, got '$storedEventApp'" }
    if ($storedProfileApp -ne "") { throw "the session profile must carry the same normalised host, got '$storedProfileApp'" }

    # ---- E2-02 session/document binding: group by document, first write wins ----
    $docA = "smoke-doc-a-" + [Guid]::NewGuid().ToString("N")
    $docB = "smoke-doc-b-" + [Guid]::NewGuid().ToString("N")
    $sessionOnA = "smoke-bind-a-" + [Guid]::NewGuid().ToString("N")
    $sessionOnB = "smoke-bind-b-" + [Guid]::NewGuid().ToString("N")
    $sessionGeneral = "smoke-bind-general-" + [Guid]::NewGuid().ToString("N")

    [ShareRibbon.ConversationRepository]::InsertMessage($sessionOnA, "user", "summarize the headings", $false, "Word", $docA, "report.docx")
    [ShareRibbon.ConversationRepository]::InsertMessage($sessionOnB, "user", "check the formulas", $false, "Word", $docB, "budget.xlsx")
    [ShareRibbon.ConversationRepository]::InsertMessage($sessionGeneral, "user", "general question", $false, "Word", "", "")
    # a later turn must not silently move the session to whichever document is active now
    [ShareRibbon.ConversationRepository]::InsertMessage($sessionOnA, "assistant", "here are the headings", $false, "Word", $docB, "other.docx")

    $onlyA = [ShareRibbon.ConversationRepository]::GetRecentSessions(200, "Word", $docA, "document")
    if (-not ($onlyA | Where-Object { $_.SessionId -eq $sessionOnA })) {
        throw "document scope must return the session bound to that document"
    }
    if ($onlyA | Where-Object { $_.SessionId -eq $sessionOnB -or $_.SessionId -eq $sessionGeneral }) {
        throw "document scope must not leak other documents or general sessions"
    }

    $allForA = [ShareRibbon.ConversationRepository]::GetRecentSessions(200, "Word", $docA, "all")
    $bindItems = $allForA | Where-Object { $_.SessionId -eq $sessionOnA -or $_.SessionId -eq $sessionOnB -or $_.SessionId -eq $sessionGeneral }
    if (@($bindItems).Count -ne 3) { throw "all scope must list every bound session, got $(@($bindItems).Count)" }
    if ($bindItems[0].SessionId -ne $sessionOnA) { throw "current document must rank first in all scope" }
    if ($bindItems[0].DocumentName -ne "report.docx") { throw "the session profile must keep the document display name" }
    if ($bindItems[1].SessionId -ne $sessionOnB) { throw "other documents must stay visible behind the current one" }
    $generalItem = $bindItems | Where-Object { $_.SessionId -eq $sessionGeneral }
    if ($generalItem.DocumentId -ne "") { throw "an unbound session must stay unbound instead of claiming a document" }

    # ---- E2-03 auto title and rename ----
    if ($generalItem.IsRenamed) { throw "a session without a custom title must not be reported as renamed" }
    if ($generalItem.Title -ne "general question") { throw "auto title must come from the first user turn, got '$($generalItem.Title)'" }

    # string multiplication, not "New-Object string[] 40 | ForEach-Object": the latter
    # silently collapses to one word and then the long-title assertions test nothing.
    $longText = "detail " * 40
    $longSession = "smoke-bind-long-" + [Guid]::NewGuid().ToString("N")
    [ShareRibbon.ConversationRepository]::InsertMessage($longSession, "user", "rewrite the summary. " + $longText, $false, "Word", $docA, "report.docx")
    $longItem = [ShareRibbon.ConversationRepository]::GetRecentSessions(200, "Word", $docA, "document") |
        Where-Object { $_.SessionId -eq $longSession }
    if ($longItem.Title -ne "rewrite the summary.") {
        throw "auto title must break at punctuation instead of hard-cutting, got '$($longItem.Title)'"
    }

    # word boundary and the decimal-point guard: "3.5" must not end the title mid-number
    $wordSession = "smoke-bind-word-" + [Guid]::NewGuid().ToString("N")
    [ShareRibbon.ConversationRepository]::InsertMessage($wordSession, "user", "rescale the chart to 3.5 times and keep the legend readable", $false, "Word", $docA, "report.docx")
    $wordItem = [ShareRibbon.ConversationRepository]::GetRecentSessions(200, "Word", $docA, "document") |
        Where-Object { $_.SessionId -eq $wordSession }
    if ($wordItem.Title -notmatch "^rescale the chart to 3\.5") {
        throw "auto title must not break on a decimal point, got '$($wordItem.Title)'"
    }
    if ($wordItem.Title.Length -gt 31) { throw "auto title must stay short, got $($wordItem.Title.Length) chars" }
    if ($wordItem.Title -notmatch [char]::Ellipsis + "$") { throw "a truncated auto title must say so, got '$($wordItem.Title)'" }

    [void][ShareRibbon.ConversationRepository]::DeleteSessionCascade($longSession)
    [void][ShareRibbon.ConversationRepository]::DeleteSessionCascade($wordSession)

    if (-not [ShareRibbon.ConversationRepository]::RenameSession($sessionOnB, "Budget review")) {
        throw "rename must succeed for an existing session"
    }
    $renamed = [ShareRibbon.ConversationRepository]::GetRecentSessions(200, "Word", $docA, "all") |
        Where-Object { $_.SessionId -eq $sessionOnB }
    if (-not $renamed.IsRenamed -or $renamed.Title -ne "Budget review") {
        throw "custom title must win over the auto title, got '$($renamed.Title)'"
    }

    if ([ShareRibbon.ConversationRepository]::RenameSession("smoke-missing-session-" + [Guid]::NewGuid().ToString("N"), "nope")) {
        throw "rename must fail for a session that has no messages"
    }

    # a session created before migration 017 has no profile row; renaming must create one
    $conn = New-Object System.Data.SQLite.SQLiteConnection([ShareRibbon.OfficeAiDatabase]::GetConnectionString())
    $conn.Open()
    try {
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = "DELETE FROM session_summary WHERE session_id = @sid"
        [void]$cmd.Parameters.AddWithValue("@sid", $sessionGeneral)
        [void]$cmd.ExecuteNonQuery()
        $cmd.Dispose()
    }
    finally {
        $conn.Dispose()
        [System.Data.SQLite.SQLiteConnection]::ClearAllPools()
    }
    if (-not [ShareRibbon.ConversationRepository]::RenameSession($sessionGeneral, "Old session")) {
        throw "rename must add a profile row for a legacy session without one"
    }
    $legacyRenamed = [ShareRibbon.ConversationRepository]::GetRecentSessions(200, "Word", $docA, "all") |
        Where-Object { $_.SessionId -eq $sessionGeneral }
    if ($legacyRenamed.Title -ne "Old session") { throw "legacy rename must reach the list, got '$($legacyRenamed.Title)'" }
    if ($legacyRenamed.DocumentId -ne "") { throw "backfilling a profile row must not invent a document binding" }

    # ---- E2-05 grouping keys the sidebar buckets by ----
    # The group header buckets by last activity, so the list must carry both the first and the newest
    # event time: with CreatedAt only, a session opened long ago and still running today lands in "older".
    $backdate = New-Object System.Data.SQLite.SQLiteConnection([ShareRibbon.OfficeAiDatabase]::GetConnectionString())
    $backdate.Open()
    try {
        $cmd = $backdate.CreateCommand()
        $cmd.CommandText = "UPDATE conversation_event SET created_at = '2020-01-01 00:00:00' WHERE session_id = @sid AND role = 'user'"
        [void]$cmd.Parameters.AddWithValue("@sid", $sessionOnA)
        [void]$cmd.ExecuteNonQuery()
        $cmd.Dispose()
    }
    finally {
        $backdate.Dispose()
        [System.Data.SQLite.SQLiteConnection]::ClearAllPools()
    }

    $groupKeys = [ShareRibbon.ConversationRepository]::GetRecentSessions(200, "Word", $docA, "all") |
        Where-Object { $_.SessionId -eq $sessionOnA }
    if ($groupKeys.CreatedAt -notlike "2020-01-01*") {
        throw "CreatedAt must stay the first message time, got '$($groupKeys.CreatedAt)'"
    }
    if ([string]::IsNullOrWhiteSpace($groupKeys.LastActiveAt) -or $groupKeys.LastActiveAt -eq $groupKeys.CreatedAt) {
        throw "LastActiveAt must follow the newest event instead of duplicating MIN(created_at)"
    }
    if ($groupKeys.DocumentId -ne $docA) {
        throw "the list must carry the document hash, otherwise two files with the same name merge into one group"
    }

    foreach ($sid in @($excelSession, $wordSession, $legacySession, $sentinelSession, $sessionOnA, $sessionOnB, $sessionGeneral)) {
        [void][ShareRibbon.ConversationRepository]::DeleteSessionCascade($sid)
    }

    $probe = New-Object System.Data.SQLite.SQLiteConnection([ShareRibbon.OfficeAiDatabase]::GetConnectionString())
    $probe.Open()
    try {
        $cmd = $probe.CreateCommand()
        $cmd.CommandText = "SELECT COUNT(*) FROM session_summary WHERE session_id = @sid"
        [void]$cmd.Parameters.AddWithValue("@sid", $sessionGeneral)
        $leftProfile = [int]$cmd.ExecuteScalar()
        $cmd.Dispose()
    }
    finally {
        $probe.Dispose()
        [System.Data.SQLite.SQLiteConnection]::ClearAllPools()
    }
    if ($leftProfile -ne 0) { throw "deleting a session must also delete its profile row" }

    $cleanup = New-Object System.Data.SQLite.SQLiteConnection([ShareRibbon.OfficeAiDatabase]::GetConnectionString())
    $cleanup.Open()
    try {
        $cmd = $cleanup.CreateCommand()
        $cmd.CommandText = "DELETE FROM agent_run_change WHERE run_id = @run_id"
        [void]$cmd.Parameters.AddWithValue("@run_id", $runId)
        [void]$cmd.ExecuteNonQuery()
        $cmd.Dispose()

        $cmd = $cleanup.CreateCommand()
        $cmd.CommandText = "DELETE FROM agent_run_change WHERE run_id = @run_id"
        [void]$cmd.Parameters.AddWithValue("@run_id", $undoRunId)
        [void]$cmd.ExecuteNonQuery()
        $cmd.Dispose()
    }
    finally {
        $cleanup.Dispose()
        [System.Data.SQLite.SQLiteConnection]::ClearAllPools()
    }

    [pscustomobject]@{
        Database = [ShareRibbon.OfficeAiDatabase]::GetDatabasePath()
        ChangeRecords = $loaded.Count
        ChangedRecords = $log.ChangedCount
        PartialApplyCarried = $true
        UndoPointCarried = $true
        UndoScopeVerified = $true
        RestoreReplayUndoable = $true
        MixedRunUndoRefused = $true
        RollbackExcluded = $true
        HostScopeIsolated = $true
        LegacyRowsVisible = $true
        UnknownAppTypeNormalizedAtWrite = $true
        DocumentScopeIsolated = $true
        CurrentDocumentPreferred = $true
        GroupingKeysReturned = $true
        FirstDocumentBindingWins = $true
        AutoTitleReadable = $true
        RenameWinsOverAutoTitle = $true
        ProfileRowDeletedWithSession = $true
    } | Format-List
}
finally {
    # A failed assertion skips the inline deletes, and this script writes into the real
    # user database. Sweep our own "smoke-" fixtures so a red run does not leave test
    # sessions in the history sidebar. Use the cascade helper: raw deletes trip the FKs on
    # memory/event tables. Best effort, so the sweep never masks the real error.
    try {
        $sweep = New-Object System.Data.SQLite.SQLiteConnection([ShareRibbon.OfficeAiDatabase]::GetConnectionString())
        $sweep.Open()
        $ids = @()
        try {
            $cmd = $sweep.CreateCommand()
            $cmd.CommandText = "SELECT DISTINCT session_id FROM conversation_event WHERE session_id LIKE 'smoke-%' UNION SELECT DISTINCT session_id FROM session_summary WHERE session_id LIKE 'smoke-%'"
            $rdr = $cmd.ExecuteReader()
            while ($rdr.Read()) { $ids += $rdr.GetString(0) }
            $rdr.Close()
            $cmd.Dispose()

            $cmd = $sweep.CreateCommand()
            $cmd.CommandText = "DELETE FROM agent_run_change WHERE run_id LIKE 'smoke-%'"
            [void]$cmd.ExecuteNonQuery()
            $cmd.Dispose()
        }
        finally {
            $sweep.Dispose()
            [System.Data.SQLite.SQLiteConnection]::ClearAllPools()
        }
        foreach ($sid in $ids) {
            [void][ShareRibbon.ConversationRepository]::DeleteSessionCascade($sid)
        }
    }
    catch {
        Write-Warning "smoke fixture sweep failed: $($_.Exception.Message)"
    }
    Pop-Location
}
