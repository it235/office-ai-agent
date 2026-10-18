param(
    [int]$LargeFileThreshold = 800
)

$ErrorActionPreference = "Stop"

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$sourceRoots = @("ShareRibbon", "WordAi", "ExcelAi", "PowerPointAi")

$allowedLargeFiles = @(
    "WordAi\ChatControl.vb",
    "ShareRibbon\Controls\BaseChatControl.vb",
    "PowerPointAi\ChatControl.vb",
    "ShareRibbon\Config\ConfigApiForm.vb",
    "ShareRibbon\Controls\BaseDataCapturePane.vb",
    "ExcelAi\ExcelDirectOperationService.vb",
    "ShareRibbon\Controls\Services\IntentRecognitionService.vb",
    "ShareRibbon\Mcp\MCPConfigForm.vb",
    "ShareRibbon\Services\Reformat\SmartFormattingOrchestrator.vb",
    "ShareRibbon\Controls\Services\HttpStreamService.vb",
    "WordAi\WordDocumentTranslateService.vb",
    "ExcelAi\ChatControl.vb",
    "ShareRibbon\Services\Reformat\DocumentAnalyzer.vb",
    "ShareRibbon\Controls\BaseDeepseekChat.vb",
    "ShareRibbon\Controls\ReformatTemplateEditorControl.vb",
    "ShareRibbon\Controls\Services\ReformatService.vb",
    "ShareRibbon\Storage\MemoryRepository.vb",
    "ExcelAi\ExcelJsonCommandSchema.vb",
    "ShareRibbon\Config\ConfigPromptForm.vb",
    "ShareRibbon\Agent\ToolRegistry.vb",
    "ShareRibbon\Controls\Services\ChatFormatterAgent.vb",
    "ShareRibbon\Config\PromptManager.vb",
    "ShareRibbon\Config\ReformatTemplateManager.vb",
    "ShareRibbon\Services\SkillsService.vb",
    "ShareRibbon\Config\ReformatTemplateEditorForm.vb",
    "ShareRibbon\Controls\BaseDoubaoChat.vb",
    "ShareRibbon\Storage\OfficeAiDatabase.vb",
    "ShareRibbon\Storage\AgentMemoryRepository.vb"
)

function Get-RelativePath {
    param([string]$Path)
    return $Path.Substring($repoRoot.Path.Length + 1)
}

function Get-SourceFiles {
    Get-ChildItem -Path ($sourceRoots | ForEach-Object { Join-Path $repoRoot $_ }) -Recurse -Filter *.vb -File |
        Where-Object {
            $_.FullName -notmatch "\\(bin|obj)\\" -and
            $_.Name -notlike "*.Designer.vb"
        }
}

function Test-CodeLineMatch {
    param(
        [string]$Path,
        [string]$Pattern
    )

    foreach ($line in Get-Content -LiteralPath $Path -Encoding UTF8) {
        $trimmed = $line.TrimStart()
        if ($trimmed.StartsWith("'")) {
            continue
        }
        if ($line -match $Pattern) {
            return $true
        }
    }
    return $false
}

$failures = New-Object System.Collections.Generic.List[string]

# Reads a file for pattern guards, dropping comment-only lines so that an explanatory
# note can never satisfy (or trip) a guard that is supposed to talk about real code.
function Get-CodeSource {
    param([string]$RelativePath)

    $fullPath = Join-Path $repoRoot $RelativePath
    if (-not (Test-Path -LiteralPath $fullPath)) {
        $failures.Add("Guard target missing: $RelativePath")
        return ""
    }

    $kept = New-Object System.Collections.Generic.List[string]
    # -Encoding UTF8 is mandatory here: Windows PowerShell 5.1 would otherwise read these BOM-less
    # UTF-8 files as ANSI, and a Chinese comment byte can swallow the following newline.
    foreach ($line in Get-Content -LiteralPath $fullPath -Encoding UTF8) {
        if ($line.TrimStart().StartsWith("'")) {
            continue
        }
        $kept.Add($line)
    }
    return ($kept -join "`n")
}

# P0-1 guardrail: no new >800-line VB files outside the baseline list.
$largeFiles = Get-SourceFiles | ForEach-Object {
    $lineCount = (Get-Content -LiteralPath $_.FullName -Encoding UTF8).Count
    if ($lineCount -gt $LargeFileThreshold) {
        [pscustomobject]@{
            Path = Get-RelativePath $_.FullName
            Lines = $lineCount
        }
    }
}

$unexpectedLarge = $largeFiles | Where-Object { $allowedLargeFiles -notcontains $_.Path }
foreach ($item in $unexpectedLarge) {
    $failures.Add("New oversized VB file detected: $($item.Path) ($($item.Lines) lines). Move responsibilities into a service instead of growing a god class.")
}

# P0-2 guardrail: Ralph UI must not be loaded as a product entrypoint.
$htmlFiles = Get-ChildItem -Path (Join-Path $repoRoot "ShareRibbon\Resources") -Recurse -Include *.html -File
foreach ($html in $htmlFiles) {
    $content = Get-Content -LiteralPath $html.FullName -Raw
    if ($content -match '<script[^>]+ralph-loop\.js') {
        $failures.Add("Ralph loop script is still loaded by HTML: $(Get-RelativePath $html.FullName)")
    }
}

# P0-3 guardrail: Task.Run in host projects that also import Office interop must be explicitly reviewed.
$hostProjectPaths = @(
    (Join-Path $repoRoot "WordAi"),
    (Join-Path $repoRoot "ExcelAi"),
    (Join-Path $repoRoot "PowerPointAi")
)

$taskRunInteropFiles = Get-ChildItem -Path $hostProjectPaths -Recurse -Filter *.vb -File |
    Where-Object { $_.FullName -notmatch "\\(bin|obj)\\" -and $_.Name -notlike "*.Designer.vb" } |
    Where-Object {
        (Test-CodeLineMatch -Path $_.FullName -Pattern "Task\.Run") -and
        (Test-CodeLineMatch -Path $_.FullName -Pattern "Microsoft\.Office\.Interop")
    } |
    ForEach-Object { Get-RelativePath $_.FullName }

$allowedTaskRunInteropFiles = @(
    "ExcelAi\ThisAddIn.vb",
    "ExcelAi\ChatControl.vb",
    "PowerPointAi\ChatControl.vb",
    "WordAi\ChatControl.vb"
)

foreach ($path in $taskRunInteropFiles) {
    if ($allowedTaskRunInteropFiles -notcontains $path) {
        $failures.Add("Task.Run + Office Interop file needs review: $path")
    }
}

# P2-A guardrail: STA-bound host work must stay on the main thread.
# Excel's VBA preview used to block the UI thread while its COM calls waited on that same thread;
# Word's proofread correction pushed document writes onto the thread pool and swallowed the errors.
$excelPreviewSource = Get-CodeSource "ExcelAi\EnhancedPreviewAndConfirm.vb"
if ($excelPreviewSource -match "Task\.Run|SyncOverAsync") {
    $failures.Add("Excel VBA preview left the UI/STA thread: ExcelAi\EnhancedPreviewAndConfirm.vb")
}

$wordChatSource = Get-CodeSource "WordAi\ChatControl.vb"
if ($wordChatSource -match "Sub HandleProofreadFocusMode[\s\S]{0,900}?Task\.Run") {
    $failures.Add("Proofread corrections are dispatched off the UI thread: WordAi\ChatControl.vb")
}
if ($wordChatSource -match "Function ApplyCorrectionToWord[\s\S]{0,200}?Task\.Run") {
    $failures.Add("Word document replace runs on a thread-pool thread: WordAi\ChatControl.vb")
}

# Word routes are launched from background routing threads, so every object-model capture phase has
# to be marshalled back to UI/STA. Deterministic COM writes (direct formatting) marshal whole-call.
if ($wordChatSource -notmatch "RunUiActionSync\([\s\S]{0,1600}?For Each p In selRange\.Paragraphs") {
    $failures.Add("Proofread paragraph walk left the UI/STA thread: WordAi\ChatControl.vb ExecuteProofreadAsync")
}
if ($wordChatSource -notmatch "RunUiActionSync\([\s\S]{0,700}?CreateReformatJobFromRange\(targetRange, scopeKind\)") {
    $failures.Add("Reformat paragraph capture left the UI/STA thread: WordAi\ChatControl.vb TriggerSmartReformat")
}
if ($wordChatSource -notmatch "RunUiActionSync\([\s\S]{0,400}?RunDirectFormatting\(userMessage\)") {
    $failures.Add("Direct formatting COM writes run off the UI/STA thread: WordAi\ChatControl.vb")
}
if ($wordChatSource -notmatch "RunUiFuncSync\([\s\S]{0,400}?HasUsableProofreadSelection\(") {
    $failures.Add("Proofread selection probe reads the object model off UI/STA: WordAi\ChatControl.vb")
}

# P2-A guardrail: host context capture is reached from background routing/agent threads, so the
# shared host adapter must marshal every object-model read back to UI/STA. RAG enrichment is the
# opposite case: it blocks up to 5s and must stay off the UI thread.
$baseChatSource = Get-CodeSource "ShareRibbon\Controls\BaseChatControl.vb"
$stCapturePatterns = @(
    "RunUiFuncSync\(Function\(\) GetContextSnapshot\(\)\)",
    "RunUiFuncSync\(Function\(\) CaptureOfficeContext\(appType\)\)",
    "RunUiFuncSync\(Function\(\) GetApplicationType\(\)\)",
    "RunUiFuncSync\(Function\(\) CaptureCurrentSelectionInfo\("
)
foreach ($pattern in $stCapturePatterns) {
    if ($baseChatSource -notmatch $pattern) {
        $failures.Add("Host context capture lost its UI/STA marshal: $pattern")
    }
}
$officeContentCaptures = [regex]::Matches($baseChatSource, "RunUiFuncSync\(Function\(\) GetCurrentOfficeContent\(\)\)").Count
if ($officeContentCaptures -lt 2) {
    $failures.Add("Both Agent start paths must read office content on UI/STA (found $officeContentCaptures).")
}
if ($baseChatSource -match "IChatRoutingHost_EnrichContextForIntent[\s\S]{0,400}?RunUi(Func|Action)Sync") {
    $failures.Add("RAG enrichment was pushed onto the UI thread: ShareRibbon\Controls\BaseChatControl.vb")
}

# P2-C guardrail: a failed Word write must not leave the document half-applied.
# One write command == one custom undo record; Undo(1) is only legal inside the rollback helper,
# and only when this round actually opened that record (otherwise it reverts the user's own edits).
$wordUndoHelper = [regex]::Match($wordChatSource, "(?m)^    Private Function TryRollbackWordWrite[\s\S]*?^    End Function").Value
if ([string]::IsNullOrWhiteSpace($wordUndoHelper)) {
    $failures.Add("Word write rollback helper is missing: WordAi\ChatControl.vb TryRollbackWordWrite")
}
# The command label is Chinese in source; spell it with \u escapes so this .ps1 stays pure ASCII
# (Windows PowerShell 5.1 reads BOM-less UTF-8 as ANSI and a Chinese byte can eat the next char).
if ($wordChatSource -notmatch "Function ExecuteWordCommandWithToolResult[\s\S]{0,1800}?StartCustomRecord\(`"AI\u547D\u4EE4:`" & normalizedToolId\)") {
    $failures.Add("Word write command no longer opens a custom undo record: WordAi\ChatControl.vb")
}
if ($wordChatSource -notmatch "If undoStarted Then[\s\S]{0,200}?EndCustomRecord\(\)") {
    $failures.Add("Word undo record is not closed under a guard: WordAi\ChatControl.vb")
}
$rollbackCalls = [regex]::Matches($wordChatSource, "= TryRollbackWordWrite\(").Count
if ($rollbackCalls -lt 2) {
    $failures.Add("Both the failure and exception paths must attempt rollback (found $rollbackCalls).")
}
if ($wordUndoHelper -notmatch "If Not undoStarted Then[\s\S]{0,300}?Return rollback") {
    $failures.Add("Word rollback can Undo without this round's custom record: WordAi\ChatControl.vb")
}
# Automatic Undo lives only in the helper; HandleUndoReformat is a separate user-initiated undo.
$wordCommandFn = [regex]::Match($wordChatSource, "(?m)^    Private Function ExecuteWordCommandWithToolResult[\s\S]*?^    End Function").Value
if ([string]::IsNullOrWhiteSpace($wordCommandFn)) {
    $failures.Add("Word tool-result command executor is missing: WordAi\ChatControl.vb")
}
if ($wordCommandFn -match "\.Undo\(1\)") {
    $failures.Add("Word executor undoes directly instead of via the rollback helper: WordAi\ChatControl.vb")
}
if ([regex]::Matches($wordUndoHelper, "\.Undo\(1\)").Count -ne 1) {
    $failures.Add("The Word rollback helper must hold exactly one Undo(1) call.")
}if ($wordChatSource -notmatch "observation\(`"rollback`"\) = rollback") {
    $failures.Add("Word rollback result is not reported in the observation: WordAi\ChatControl.vb")
}
if ($wordChatSource -notmatch "ElseIf rollbackAttempted Then[\s\S]{0,240}?ExceptionClassifier\.CodePartialApply" -or
    $wordChatSource -notmatch "errorCode:=hostCode") {
    $failures.Add("Word partial-apply is not surfaced as CodePartialApply on the returned errorCode: WordAi\ChatControl.vb")
}
if ($wordChatSource -notmatch "observation\(`"hostReason`"\) = _lastWordWriteFailure") {
    $failures.Add("Word host executor no longer reports a named failure reason into the observation: WordAi\ChatControl.vb")
}
if ($wordChatSource -notmatch "_lastWordWriteFailure = Nothing") {
    $failures.Add("The Word failure-reason channel is not cleared per command, so a stale reason can leak into the next tool result: WordAi\ChatControl.vb")
}

# P2-C guardrail: Excel has no per-command undo record, so it must stay honest instead of faking a
# rollback. Application.Undo in Excel would also revert the user's earlier edits, so the host reports
# attempted=False + partial apply and stops automatic replay (a replay would re-apply on changed data).
$excelChatSource = Get-CodeSource "ExcelAi\ChatControl.vb"
if ($excelChatSource -notmatch "Function BuildExcelRollbackNote[\s\S]{0,700}?attempted`", False") {
    $failures.Add("Excel rollback note no longer states that no rollback was attempted: ExcelAi\ChatControl.vb")
}
$rollbackNotes = [regex]::Matches($excelChatSource, "observation\(`"rollback`"\) = BuildExcelRollbackNote\(").Count
if ($rollbackNotes -lt 3) {
    $failures.Add("Every Excel write outcome must carry a rollback note (found $rollbackNotes, need >= 3).")
}
if ($excelChatSource -notmatch "If\(partialApply,\s*ExceptionClassifier\.CodePartialApply") {
    $failures.Add("Excel partial apply is not reported as CodePartialApply: ExcelAi\ChatControl.vb")
}
if ($excelChatSource -notmatch "recoverable:=Not partialApply") {
    $failures.Add("Excel partial apply still allows automatic replay: ExcelAi\ChatControl.vb")
}
if ($excelChatSource -notmatch "crashResult\.Observation = crashObservation") {
    $failures.Add("Excel exception path lost its observation: ExcelAi\ChatControl.vb")
}
$excelUndoFiles = Get-SourceFiles | Where-Object {
    (Get-RelativePath $_.FullName) -like "ExcelAi\*" -and (Test-CodeLineMatch -Path $_.FullName -Pattern "Application\.Undo|\\.Undo\(")
} | ForEach-Object { Get-RelativePath $_.FullName }
foreach ($file in $excelUndoFiles) {
    $failures.Add("Excel calls Undo automatically; it cannot be scoped to one AI command: $file")
}

# P3-B guardrail: PowerPoint generic commands must reach the same partial-apply contract as Word/Excel.
# Only the CreateSlides chain can delete its own pages; a plain command envelope that already moved the
# presentation must stop automatic replay instead of pretending the deck is untouched.
$pptChatSource = Get-CodeSource "PowerPointAi\ChatControl.vb"
if ($pptChatSource -notmatch "Function ExecuteJsonCommandWithToolResult[\s\S]{0,4000}?ExceptionClassifier\.CodeObservationFailed") {
    $failures.Add("PowerPoint treats an unverified success as final: PowerPointAi\ChatControl.vb")
}
$pptRollbackNotes = [regex]::Matches($pptChatSource, "observation\(`"rollback`"\) = BuildPowerPointRollbackNote\(").Count
if ($pptRollbackNotes -lt 2) {
    $failures.Add("Every PowerPoint write outcome must carry a rollback note (found $pptRollbackNotes, need >= 2).")
}
if ($pptChatSource -notmatch "If\(partialApply,\s*ExceptionClassifier\.CodePartialApply") {
    $failures.Add("PowerPoint partial apply is not reported as CodePartialApply: PowerPointAi\ChatControl.vb")
}
if ($pptChatSource -notmatch "recoverable:=Not partialApply") {
    $failures.Add("PowerPoint partial apply still allows automatic replay: PowerPointAi\ChatControl.vb")
}

# Declarative operation batches: a step that already landed cannot be un-applied as a group, so the
# executor must expose rollback + non-replay instead of only an English warning string.
$pptBatchSource = Get-CodeSource "PowerPointAi\Runtime\PowerPointOperationExecutor.vb"
if ($pptBatchSource -match "Atomic batch partially applied") {
    $failures.Add("Partial-apply warning is still English-only user text: PowerPointOperationExecutor.vb")
}
$pptBatchRollbacks = [regex]::Matches($pptBatchSource, "failureObservation\(`"rollback`"\) = BuildOperationBatchRollbackNote\(").Count
if ($pptBatchRollbacks -lt 2) {
    $failures.Add("Both batch failure catches must publish a rollback note (found $pptBatchRollbacks, need >= 2).")
}
if ($pptBatchSource -notmatch "recoverable:=If\(partialApplied, False, ex\.Recoverable\)") {
    $failures.Add("A partially applied batch stays auto-replayable: PowerPointOperationExecutor.vb")
}
if ($pptBatchSource -notmatch "recoverable:=If\(crashedPartially, False, classified\.Recoverable\)") {
    $failures.Add("A batch that crashed after partial apply stays auto-replayable: PowerPointOperationExecutor.vb")
}

# P2-D guardrail: source text must never carry replacement characters. A tool once re-saved these
# files UTF-8 from GBK and destroyed every Chinese literal, so users saw garbage in dialogs.
# Checked on raw bytes because Get-CodeSource drops comment lines and the damage lives in comments too.
$mangledPattern = "[\ufffd]|\u951f\u65a4\u62f7"
foreach ($file in (Get-SourceFiles)) {
    $relative = Get-RelativePath $file.FullName
    # Legacy GBK sources are still readable; only flag text that is damaged in its own encoding.
    $bytes = [System.IO.File]::ReadAllBytes($file.FullName)
    $utf8Strict = New-Object System.Text.UTF8Encoding($false, $true)
    try {
        $text = $utf8Strict.GetString($bytes)
    } catch [System.Text.DecoderFallbackException] {
        $text = [System.Text.Encoding]::GetEncoding("GB18030").GetString($bytes)
    }
    if ($text -match $mangledPattern) {
        $failures.Add("Source text contains mojibake / replacement characters: $relative")
    }
}

# P0-4 guardrail: every script-injection surface must route through the shared SafetyGate helper.
$scriptSurfaceFiles = Get-SourceFiles | Where-Object {
    (Test-CodeLineMatch -Path $_.FullName -Pattern "AddFromString\(|MSScriptControl")
}
foreach ($file in $scriptSurfaceFiles) {
    $relative = Get-RelativePath $file.FullName
    if ($relative -eq "ShareRibbon\Controls\Services\CodeExecutionService.vb") {
        continue
    }
    if (-not (Test-CodeLineMatch -Path $file.FullName -Pattern "ConfirmScriptExecution|ScriptSafetyBlocked")) {
        $failures.Add("Ungated script/VBA injection surface: $relative")
    }
}

# P1-F guardrail: an approval request must own a timeout exit.
# A pending approval that nobody answered used to hold the session forever, and the next turn
# then hit the RUN_ALREADY_ACTIVE guard with no way back.
$harnessFiles = @(Get-SourceFiles | Where-Object {
    (Get-RelativePath $_.FullName) -like "ShareRibbon\Agent\Harness\OfficeHarness*.vb"
})
$harnessSource = ($harnessFiles | ForEach-Object { Get-Content -LiteralPath $_.FullName -Raw -Encoding UTF8 }) -join "`n"
if ($harnessSource -match "HarnessRunStatus\.AwaitingApproval" -or $harnessSource -match "ApprovalCallback") {
    if ($harnessSource -notmatch "(?m)^\s*ArmApprovalTimeout") {
        $failures.Add("Awaiting-approval run has no timeout exit: OfficeHarness partials")
    }
    if ($harnessSource -notmatch "AwaitingApproval\s*=\s*True") {
        $failures.Add("Guard cannot tell an approval wait apart from live host work: OfficeHarness partials")
    }
    if ($harnessSource -notmatch "FirstOrDefault\(Function\(item\)[^\r\n]*IsCompleted") {
        $failures.Add("Run guard no longer distinguishes an approval wait from live host work: OfficeHarness partials")
    }
}

# P1-G guardrail: "stop" must reach the in-flight model request, not only the next step boundary.
# Losing the ambient scope silently reverts the turn to paying full price for a reply it will throw away.
$turnSource = Get-CodeSource "ShareRibbon\Controls\Services\AgentKernelService.vb"
if ($turnSource -notmatch "AgentCancellationScope\.Enter\(runCts\.Token\)") {
    $failures.Add("Agent turn does not publish its cancellation token to the ambient scope")
}

$chatSource = Get-CodeSource "ShareRibbon\Controls\BaseChatControl.vb"
if ($chatSource -notmatch "CreateLinkedTokenSource\(ambientCancellation\)") {
    $failures.Add("Non-streaming AI request no longer observes the turn cancellation scope")
}

$loopSource = Get-CodeSource "ShareRibbon\Agent\LoopEngine.vb"
if ($loopSource -notmatch "Catch ex As OperationCanceledException When cancellationToken\.IsCancellationRequested") {
    $failures.Add("An aborted model request converges as a failure instead of a user stop: LoopEngine")
}
if (([regex]::Matches($loopSource, "If cancellationToken\.IsCancellationRequested Then Return FinishCancelled")).Count -lt 2) {
    $failures.Add("Cancelled turn can still reach the host executor after a model reply: LoopEngine")
}

# P3-A guardrail: refusing the preview is a cancellation, not a successful write.
# The Boolean host contract deliberately returns True on cancel (so the UI only re-enables buttons),
# therefore the ToolResult layer must read an explicit marker or the Loop records a change that never happened.
$cancelHostFiles = @(
    "WordAi\ChatControl.vb",
    "ExcelAi\ChatControl.vb",
    "PowerPointAi\ChatControl.vb"
)
foreach ($relative in $cancelHostFiles) {
    $hostLines = ((Get-CodeSource $relative) -split "`n")
    for ($i = 0; $i -lt $hostLines.Count; $i++) {
        if ($hostLines[$i] -notmatch "handleExecutionCancelled") {
            continue
        }
        $window = ($hostLines[$i..([Math]::Min($i + 3, $hostLines.Count - 1))] -join "`n")
        if ($window -notmatch "NotifyPreviewCancelled\(\)|ToolResult\.PreviewCancelled\(") {
            $failures.Add("Preview cancel point does not register the cancel marker: $relative")
        }
    }
    $hostSource = Get-CodeSource $relative
    if ($hostSource -notmatch "Function ExecuteJsonCommandWithToolResult[\s\S]{0,4000}?ConsumePreviewCancelled\(\)") {
        $failures.Add("ToolResult entry does not consume the cancel marker: $relative")
    }
    if ($hostSource -match "ToolResult\.Succeed\(`"`"") {
        $failures.Add("A success result carries an empty toolId, so nothing is attributable: $relative")
    }
}

# Word keeps a legacy Boolean executor as its last resort. That path has no snapshot, so it can never
# claim a verified write; it must report the observation gap and stop automatic replay.
if ($wordChatSource -notmatch "If fallbackOk Then[\s\S]{0,700}?ExceptionClassifier\.CodeObservationFailed") {
    $failures.Add("Word fallback reports an unverified write as success: WordAi\ChatControl.vb")
}

$toolRegistrySource = Get-CodeSource "ShareRibbon\Agent\ToolRegistry.vb"
if ($toolRegistrySource -notmatch "Function PreviewCancelled[\s\S]{0,900}?ExceptionClassifier\.CodeCancelled") {
    $failures.Add("PreviewCancelled no longer reports the CANCELLED error code: ShareRibbon\Agent\ToolRegistry.vb")
}
if ($toolRegistrySource -notmatch "Function PreviewCancelled[\s\S]{0,900}?recoverable:=False") {
    $failures.Add("A cancelled preview stays auto-replayable: ShareRibbon\Agent\ToolRegistry.vb")
}

if ($chatSource -notmatch "If Not approved Then NotifyPreviewCancelled\(\)") {
    $failures.Add("VBA preview rejection is not marked as a cancel: ShareRibbon\Controls\BaseChatControl.vb")
}
if ($chatSource -notmatch "PreviewCancelledReader = AddressOf ConsumePreviewCancelled") {
    $failures.Add("CodeExecutionService cannot see the host cancel marker: ShareRibbon\Controls\BaseChatControl.vb")
}

$codeExecSource = Get-CodeSource "ShareRibbon\Controls\Services\CodeExecutionService.vb"
if ($codeExecSource -notmatch "Function ExecuteCodeWithToolResult[\s\S]{0,3000}?If CancelledByUser\(\) Then[\s\S]{0,140}?ToolResult\.PreviewCancelled") {
    $failures.Add("Script path can still report a cancelled preview as success: CodeExecutionService.vb")
}
$scriptExecFn = [regex]::Match($codeExecSource, "(?m)^        Public Function ExecuteCodeWithToolResult[\s\S]*?^        End Function").Value
if ($scriptExecFn -notmatch "recoverable:=False") {
    $failures.Add("A failed script run is auto-replayed even though it may have half-applied: CodeExecutionService.vb")
}

# P4-A guardrail: model-produced markdown may only reach the DOM through the sanitizer.
# marked passes raw HTML through, so an unsanitized innerHTML assignment turns a model
# response into executable script inside the add-in pane, and that script can post
# executeCode messages back into the host command entry point.
function Get-JsSource {
    param([string]$RelativePath)

    $fullPath = Join-Path $repoRoot $RelativePath
    if (-not (Test-Path -LiteralPath $fullPath)) {
        $failures.Add("JS guard target missing: $RelativePath")
        return ""
    }

    $kept = New-Object System.Collections.Generic.List[string]
    foreach ($line in Get-Content -LiteralPath $fullPath -Encoding UTF8) {
        if ($line.TrimStart().StartsWith("//")) {
            continue
        }
        $kept.Add($line)
    }
    return ($kept -join "`n")
}

function Get-JsSetLiteral {
    param(
        [string]$Source,
        [string]$SetName
    )

    return [regex]::Match($Source, [regex]::Escape($SetName) + " = (?:new Set\()?\[[\s\S]*?\]").Value
}

$jsResourceDir = "ShareRibbon\Resources\js"
$jsFiles = Get-ChildItem -Path (Join-Path $repoRoot $jsResourceDir) -Filter *.js -File |
    Where-Object { $_.Name -notlike "*.min.js" }

foreach ($js in $jsFiles) {
    $jsRelative = Join-Path $jsResourceDir $js.Name
    $jsSource = Get-JsSource $jsRelative
    foreach ($line in ($jsSource -split "`n")) {
        if ($line -match "marked\.parse" -and $line -notmatch "sanitizeHtml\(|renderMarkdown\(") {
            $failures.Add("Markdown reaches the DOM without sanitizing: $jsRelative")
        }
    }
    if ($js.Name -ne "utils.js" -and $jsSource -match "function (sanitizeHtml|renderMarkdown)\(") {
        $failures.Add("A second sanitizer escaped into $jsRelative; utils.js is the single choke point")
    }
}

$utilsJsSource = Get-JsSource (Join-Path $jsResourceDir "utils.js")
if ($utilsJsSource -notmatch "new DOMParser\(\)") {
    $failures.Add("sanitizeHtml must parse through an inert DOMParser document, never by regex substitution")
}
if ($utilsJsSource -notmatch "function isSafeUrl\(") {
    $failures.Add("sanitizeHtml lost its URL scheme check: $jsResourceDir\utils.js")
}
$forbiddenTags = Get-JsSetLiteral $utilsJsSource "HTML_FORBIDDEN_TAGS"
foreach ($tag in @("script", "iframe", "form", "input", "object", "embed", "style", "base")) {
    if ($forbiddenTags -notmatch [regex]::Escape("'$tag'")) {
        $failures.Add("HTML sanitizer no longer drops '$tag' outright: $jsResourceDir\utils.js")
    }
}
$allowedTags = Get-JsSetLiteral $utilsJsSource "HTML_ALLOWED_TAGS"
foreach ($tag in @("p", "ul", "ol", "li", "table", "pre", "code", "button", "svg", "span")) {
    if ($allowedTags -notmatch [regex]::Escape("'$tag'")) {
        $failures.Add("HTML sanitizer would silently destroy legitimate markdown tag '$tag': $jsResourceDir\utils.js")
    }
}
if ($allowedTags -match "'(script|style|iframe|form)'") {
    $failures.Add("HTML sanitizer allowlist carries a forbidden tag: $jsResourceDir\utils.js")
}
$allowedAttrs = Get-JsSetLiteral $utilsJsSource "HTML_ALLOWED_ATTRS"
if ($allowedAttrs -match "'(on\w+|style|srcdoc|formaction|xlink:href)'") {
    $failures.Add("HTML sanitizer attribute allowlist carries an executable attribute: $jsResourceDir\utils.js")
}
$unsafeSchemes = Get-JsSetLiteral $utilsJsSource "HTML_SAFE_URL_SCHEMES"
if ($unsafeSchemes -match "javascript|data:|vbscript") {
    $failures.Add("HTML sanitizer accepts an executable URL scheme: $jsResourceDir\utils.js")
}
if ($utilsJsSource -notmatch "renderMarkdown[\s\S]{0,600}?element\.innerHTML = sanitizeHtml\(marked\.parse") {
    $failures.Add("renderMarkdown no longer routes marked output through sanitizeHtml: $jsResourceDir\utils.js")
}

$coreJsSource = Get-JsSource (Join-Path $jsResourceDir "core.js")
if ($coreJsSource -match "onclick=`"(copy|execute|edit)Code\(") {
    $failures.Add("core.js still emits inline handlers on code-block buttons")
}
foreach ($action in @("copy", "edit", "execute")) {
    if ($coreJsSource -notmatch [regex]::Escape("data-code-action=`"$action`"")) {
        $failures.Add("core.js lost the data-code-action marker for '$action': code buttons become inert")
    }
}

$codeHandlerSource = Get-JsSource (Join-Path $jsResourceDir "code-handler.js")
$codeActionTable = [regex]::Match($codeHandlerSource, "const CODE_ACTION_HANDLERS = \{[\s\S]*?\}").Value
foreach ($pair in @("copy: copyCode", "execute: executeCode", "edit: editCode")) {
    if ($codeActionTable -notmatch [regex]::Escape($pair)) {
        $failures.Add("code-handler.js action table lost '$pair': that code-block button becomes inert")
    }
}
if ($codeHandlerSource -notmatch "document\.addEventListener\('click'" -or $codeHandlerSource -notmatch "closest\('\[data-code-action\]'\)") {
    $failures.Add("code-handler.js no longer delegates clicks for sanitized code-block buttons")
}

$markdownRendererSource = Get-JsSource (Join-Path $jsResourceDir "markdown-renderer.js")
if ($markdownRendererSource -match "escapedJson") {
    $failures.Add("markdown-renderer.js re-introduced the hand-escaped inline JSON payload")
}
if ($markdownRendererSource -match "on(click|load|error|mouse\w+|focus)\s*=") {
    $failures.Add("markdown-renderer.js emits an inline event handler; card actions must be delegated: $jsResourceDir\markdown-renderer.js")
}
if ($markdownRendererSource -notmatch "sanitizeHtml\(cardHtml\)") {
    $failures.Add("Template card HTML is injected without sanitizing: markdown-renderer.js")
}
$cardBuilder = [regex]::Match($markdownRendererSource, "_createTemplateCard\(template, cardIndex\)[\s\S]*?\n    /\*\*").Value
if ($cardBuilder -notmatch "data-template-action=`"save`"" -or $cardBuilder -notmatch "data-template-action=`"preview`"") {
    $failures.Add("Template card builder lost its delegated action markers: $jsResourceDir\markdown-renderer.js")
}
if ($markdownRendererSource -notmatch "dataset\.templateJson = payload") {
    $failures.Add("Template card JSON payload is no longer mounted through dataset: $jsResourceDir\markdown-renderer.js")
}
if ($markdownRendererSource -match "replace\(/&quot;/g") {
    $failures.Add("TemplateCardActions still decodes HTML entities by hand instead of using dataset")
}

$utilsVbSource = Get-CodeSource "ShareRibbon\Controls\Services\UtilsService.vb"
if ($utilsVbSource -notmatch "data-code-action") {
    $failures.Add("Exported history HTML cannot restore code-block actions: UtilsService.vb")
}

$chatTemplatePath = Join-Path $repoRoot "ShareRibbon\Resources\chat-template-refactored.html"
function Get-ScriptLineIndex {
    param(
        [string]$FileName
    )

    $index = 0
    foreach ($line in Get-Content -LiteralPath $chatTemplatePath -Encoding UTF8) {
        $index++
        if ($line -match [regex]::Escape("js/$FileName")) {
            return $index
        }
    }
    return -1
}

$utilsLine = Get-ScriptLineIndex "utils.js"
foreach ($dependent in @("markdown-renderer.js", "chat-manager.js", "message-sender.js", "reformat-template.js")) {
    $dependentLine = Get-ScriptLineIndex $dependent
    if ($utilsLine -lt 0 -or $dependentLine -lt 0 -or $utilsLine -gt $dependentLine) {
        $failures.Add("chat-template-refactored.html loads $dependent before utils.js; sanitize helpers would be undefined")
    }
}

# P4-B guardrail: the VB -> JS bridge must hand values over as literals or JSON, never splice
# raw text into JS source, and never write raw HTML into the WebView from VB.
$bridgeFiles = @(
    "ShareRibbon\Controls\BaseChatControl.vb",
    "ShareRibbon\Controls\Services\AgentKernelService.vb",
    "ShareRibbon\Controls\Services\AgentKernelService.RunControl.vb",
    "ShareRibbon\Controls\Services\ReformatService.vb",
    "ShareRibbon\Controls\Services\HttpStreamService.vb",
    "ShareRibbon\Controls\Services\ChatFormatterAgent.vb",
    "WordAi\ChatControl.vb",
    "PowerPointAi\ChatControl.vb",
    "WordAi\Ribbon1.vb",
    "ExcelAi\Ribbon1.vb",
    "PowerPointAi\Ribbon1.vb"
)

foreach ($bridgeFile in $bridgeFiles) {
    $bridgeSource = Get-CodeSource $bridgeFile
    if ($bridgeSource -match "innerHTML\s*=\s*'<") {
        $failures.Add("VB bridge writes raw HTML into the WebView: $bridgeFile")
    }
    if ($bridgeSource -match "getElementById\('content-\{") {
        $failures.Add("VB bridge reaches into the chat DOM instead of calling a front-end helper: $bridgeFile")
    }
    if ($bridgeSource -match "'\{(?:_escapeJs|EscapeJavaScriptString)\(") {
        $failures.Add("VB bridge still hand-escapes a value into a JS single-quoted literal: $bridgeFile")
    }
}

$utilsServiceSource = Get-CodeSource "ShareRibbon\Controls\Services\UtilsService.vb"
if ($utilsServiceSource -notmatch "ChrW\(&H2028\)") {
    $failures.Add("EscapeJavaScriptString no longer escapes U+2028, which terminates a JS string literal: UtilsService.vb")
}
if ($utilsServiceSource -notmatch "ChrW\(&H2029\)") {
    $failures.Add("EscapeJavaScriptString no longer escapes U+2029, which terminates a JS string literal: UtilsService.vb")
}
$toJsLiteralBody = ""
if ($utilsServiceSource -match "Function ToJsStringLiteral[\s\S]*?End Function") {
    $toJsLiteralBody = $Matches[0]
}
if ($toJsLiteralBody -notmatch "JsonConvert\.SerializeObject") {
    $failures.Add("ToJsStringLiteral is no longer backed by JsonConvert, so callers lose quote handling: UtilsService.vb")
}
$serializeForScriptBody = ""
if ($utilsServiceSource -match "Function SerializeForScript[\s\S]*?End Function") {
    $serializeForScriptBody = $Matches[0]
}
if ($serializeForScriptBody -notmatch "EscapeJsonForScriptEmbedding") {
    $failures.Add("SerializeForScript no longer routes JSON through the script-embedding escape: UtilsService.vb")
}

$agentKernelSource = Get-CodeSource "ShareRibbon\Controls\Services\AgentKernelService.vb"
if ($agentKernelSource -notmatch "showAgentPlanCard\(\{planJson\}\)") {
    $failures.Add("Agent plan card bridge changed shape, front-end contract may break: AgentKernelService.vb")
}
if ($agentKernelSource -match '""language"":""') {
    $failures.Add("Agent plan steps are hand-spliced into JSON again, model text can break out: AgentKernelService.vb")
}
if ($agentKernelSource -notmatch "SerializeForScript\(New JObject From") {
    $failures.Add("Agent plan card is no longer serialized by Newtonsoft: AgentKernelService.vb")
}
if ($agentKernelSource -notmatch "setThinkingStatusText\(") {
    $failures.Add("Agent status line no longer goes through the plain-text helper: AgentKernelService.vb")
}

$chatManagerJs = Get-JsSource (Join-Path $jsResourceDir "chat-manager.js")
foreach ($helper in @("setSectionPlainText", "setThinkingStatusText", "showThinkingIndicator")) {
    if ($chatManagerJs -notmatch [regex]::Escape("window.$helper = function")) {
        $failures.Add("chat-manager.js lost the $helper bridge helper that VB now depends on")
    }
}
if ($chatManagerJs -notmatch "setSectionPlainText[\s\S]{0,400}?\.textContent") {
    $failures.Add("setSectionPlainText no longer writes via textContent, so VB-supplied text becomes markup: chat-manager.js")
}
if ($chatManagerJs -match "setSectionPlainText[\s\S]{0,400}?\.innerHTML") {
    $failures.Add("setSectionPlainText writes innerHTML again: chat-manager.js")
}

$reformatServiceSource = Get-CodeSource "ShareRibbon\Controls\Services\ReformatService.vb"
if ($reformatServiceSource -match "enterAiTemplateEditor") {
    $failures.Add("ReformatService pushes a whole serialized template into the WebView again: ReformatService.vb")
}
if ($reformatServiceSource -match "ThumbnailBase64|SourceFileContent") {
    $failures.Add("Template list bridge would carry thumbnail Base64 or source document text into the UI: ReformatService.vb")
}
if ($reformatServiceSource -notmatch "SerializeForScript\(allItems\)") {
    $failures.Add("loadReformatTemplateList is no longer serialized for script embedding: ReformatService.vb")
}

$formatterAgentSource = Get-CodeSource "ShareRibbon\Controls\Services\ChatFormatterAgent.vb"
if ($formatterAgentSource -match '<strong>\{(?:GetDocumentTypeLabel\(plan\)|plan\.StandardName)\}') {
    $failures.Add("Formatting card writes model-produced names into HTML without escaping: ChatFormatterAgent.vb")
}
if ($formatterAgentSource -notmatch '<strong>\{Esc\(GetDocumentTypeLabel\(plan\)\)\}') {
    $failures.Add("Formatting card document type label lost its Esc() wrapper: ChatFormatterAgent.vb")
}
if ($formatterAgentSource -notmatch "HtmlEncode\(If\(value") {
    $failures.Add("ChatFormatterAgent lost its HTML escaping helper for card text: ChatFormatterAgent.vb")
}

$wordChatSource = Get-CodeSource "WordAi\ChatControl.vb"
if ($wordChatSource -match "EscapeHtmlForInline") {
    $failures.Add("WordAi still builds inline HTML for the chat bridge instead of setSectionPlainText: WordAi\ChatControl.vb")
}
if ($wordChatSource -notmatch "setSectionPlainText\(") {
    $failures.Add("WordAi user echo no longer uses the plain-text bridge helper: WordAi\ChatControl.vb")
}

# P4-C guardrail: every JS function the VB bridge calls must exist in a WebView resource.
# A missing name fails silently inside the pane ("x is not defined" in its console only), so
# the user just sees a button that does nothing; showIdentifyingStatus and addThinkingMessage
# survived that way until this scan was added.
$bridgeCallSkip = @(
    "document", "window", "console", "localStorage", "JSON", "Object", "Array", "String",
    "Number", "setTimeout", "setInterval", "alert", "this", "if", "for", "while", "switch",
    "return", "function", "var", "let", "const", "requestAnimationFrame", "getComputedStyle",
    "Event", "CustomEvent", "Date", "Math", "Promise"
)

$jsGlobals = New-Object 'System.Collections.Generic.HashSet[string]'
$jsScanFiles = @(Get-ChildItem -LiteralPath (Join-Path $repoRoot $jsResourceDir) -Filter *.js -Recurse)
$jsScanFiles += @(Get-ChildItem -LiteralPath (Join-Path $repoRoot "ShareRibbon\Resources") -Filter *.html)
$jsGlobalPatterns = @(
    "function\s+([A-Za-z_$][\w$]*)\s*\(",
    "(?:window|globalThis)\.([A-Za-z_$][\w$]*)\s*=",
    "(?:const|let|var)\s+([A-Za-z_$][\w$]*)\s*=\s*(?:function|\()",
    "^\s*([A-Za-z_$][\w$]*)\s*:\s*function"
)
foreach ($jsFile in $jsScanFiles) {
    $jsText = Get-Content -LiteralPath $jsFile.FullName -Raw -Encoding UTF8
    foreach ($pattern in $jsGlobalPatterns) {
        foreach ($m in [regex]::Matches($jsText, $pattern, "Multiline")) {
            [void]$jsGlobals.Add($m.Groups[1].Value)
        }
    }
}
if ($jsGlobals.Count -lt 100) {
    $failures.Add("Bridge contract scan only saw $($jsGlobals.Count) JS globals; the resource scan is broken, not the bridge")
}

$bridgeCallSites = @{}
foreach ($bridgeFile in ($bridgeFiles + @("ShareRibbon\Controls\Services\ChatRequestOrchestrator.vb",
                                         "ShareRibbon\Controls\Services\CodeExecutionService.vb",
                                         "ExcelAi\ChatControl.vb"))) {
    $bridgeSource = Get-CodeSource $bridgeFile
    foreach ($m in [regex]::Matches($bridgeSource, '(?:_executeScript|ExecuteJavaScriptAsyncJS|ExecuteScript|executeScript)\s*\(\s*\$?"([^"{}]{0,400})')) {
        $head = [regex]::Match($m.Groups[1].Value, "^\s*([A-Za-z_$][\w$.]*)\s*\(")
        if (-not $head.Success) { continue }
        $callName = $head.Groups[1].Value
        if ($callName.Contains(".") -or $bridgeCallSkip -contains $callName) { continue }
        if (-not $bridgeCallSites.ContainsKey($callName)) {
            $bridgeCallSites[$callName] = New-Object System.Collections.Generic.List[string]
        }
        $bridgeCallSites[$callName].Add($bridgeFile)
    }
}
foreach ($callName in ($bridgeCallSites.Keys | Sort-Object)) {
    if (-not $jsGlobals.Contains($callName)) {
        $callFiles = ($bridgeCallSites[$callName] | Sort-Object -Unique) -join ", "
        $failures.Add("VB bridge calls JS '$callName' that no WebView resource defines, so the call is a silent no-op: $callFiles")
    }
}

# P5 guardrail: the JavaScriptStringEncode output is a JS literal, never message content.
# It used to be accumulated into a second buffer that history/memory read, so every persisted
# assistant reply carrying JSON lost its real line breaks and gained a spare escape round.
$streamSource = Get-CodeSource "ShareRibbon\Controls\Services\HttpStreamService.vb"
$chatSource = Get-CodeSource "ShareRibbon\Controls\BaseChatControl.vb"
if ($streamSource -match "_stateService\.MarkdownBuffer") {
    $failures.Add("The escaped JS buffer is reachable from the stream service again; history must read PlainMarkdownBuffer: ShareRibbon\Controls\Services\HttpStreamService.vb")
}
$plainAppends = [regex]::Matches($streamSource, "PlainMarkdownBuffer\.Append\(").Count
if ($plainAppends -ne 1) {
    $failures.Add("PlainMarkdownBuffer must have exactly one accumulation point (found $plainAppends); the non-streaming path used to append a second time and duplicate every reply: ShareRibbon\Controls\Services\HttpStreamService.vb")
}
$flushStart = $streamSource.IndexOf("Function FlushBufferAsync")
$flushBody = If ($flushStart -lt 0) { "" } Else { $streamSource.Substring($flushStart, [Math]::Min(2500, $streamSource.Length - $flushStart)) }
if ($flushBody -notmatch "PlainMarkdownBuffer\.Append\(plainContent\)") {
    $failures.Add("FlushBufferAsync no longer feeds the plain buffer, so streamed history would be empty: ShareRibbon\Controls\Services\HttpStreamService.vb")
}
if ($chatSource -notmatch "answerContent = _chatStateService\.PlainMarkdownBuffer\.ToString\(\)") {
    $failures.Add("Assistant history is not taken from the plain buffer again: ShareRibbon\Controls\BaseChatControl.vb")
}
if ($chatSource -match "JavaScriptStringEncode[\s\S]{0,120}?(AddMessage|SaveConversation|RecordConversationTurn|\.content =)") {
    $failures.Add("A JS-escaped string is reaching a history/memory writer: ShareRibbon\Controls\BaseChatControl.vb")
}

# P6 guardrail: a turn is persisted as a pair, or not at all.
# The real Word run left a user row with no assistant row, because Send wrote the question
# unconditionally while the answer is only written when addHistory is true.
if ($chatSource -notmatch 'If addHistory AndAlso Not String\.IsNullOrWhiteSpace\(userTurnText\)[\s\S]{0,160}?AddMessage\("user", userTurnText\)') {
    $failures.Add("The user turn is persisted again on addHistory=False paths, leaving turns with no answer: ShareRibbon\Controls\BaseChatControl.vb")
}

# P6 guardrail: a partially applied Word batch must not look like a clean failure.
# "4 succeeded, 1 failed" collapsed into one NOT_FOUND ToolResult while four writes stayed in
# the document, and recoverable=True invited the loop to replay the whole batch.
$wordBatchStart = $wordChatSource.IndexOf("Function ExecuteWordCommandsArrayWithToolResult")
$wordBatchBody = If ($wordBatchStart -lt 0) { "" } Else { $wordChatSource.Substring($wordBatchStart, [Math]::Min(4200, $wordChatSource.Length - $wordBatchStart)) }
if ($wordBatchBody -notmatch 'If successCount > 0 Then[\s\S]{0,1400}?ExceptionClassifier\.CodePartialApply') {
    $failures.Add("Word batch no longer names partial apply when some commands landed: WordAi\ChatControl.vb")
}
if ($wordBatchBody -notmatch 'If successCount > 0 Then[\s\S]{0,1400}?recoverable:=False') {
    $failures.Add("Word partial apply is replayable again, so applied inserts would run twice: WordAi\ChatControl.vb")
}
if ($wordBatchBody -notmatch '"attempted", False') {
    $failures.Add("Word batch stopped declaring that it did not roll the applied commands back: WordAi\ChatControl.vb")
}

# P6 guardrail: panel-triggered JSON execution needs its own correlation id, but must never
# steal or clear the scope an Agent run already owns.
$codeExecSource = Get-CodeSource "ShareRibbon\Controls\Services\CodeExecutionService.vb"
if ($codeExecSource -notmatch "ownsLogScope = AppLogger\.CorrelationId = ""-""") {
    $failures.Add("JSON execution no longer checks whether a correlation scope already exists: ShareRibbon\Controls\Services\CodeExecutionService.vb")
}
if ($codeExecSource -notmatch "If ownsLogScope Then AppLogger\.ClearScope\(\)") {
    $failures.Add("JSON execution would clear a correlation scope it does not own: ShareRibbon\Controls\Services\CodeExecutionService.vb")
}

# P7 guardrail: the tool registry, the host command whitelist and both prompt copies must describe
# the same Word capability set. ListParagraphs was registered and implemented but missing from
# SupportedCommands, so the host answered a planner-chosen tool with "unsupported command" and the
# loop spent three auto-repairs on a JSON format error that was never a format error.
$wordToolIds = @()
$wordToolDir = Join-Path $repoRoot 'ShareRibbon\Tools\word'
foreach ($toolFile in @(Get-ChildItem -LiteralPath $wordToolDir -Filter '*.json' -ErrorAction SilentlyContinue)) {
    $raw = Get-Content -LiteralPath $toolFile.FullName -Raw -Encoding UTF8
    $toolId = [regex]::Match($raw, '"id"\s*:\s*"([^"]+)"').Groups[1].Value
    if ($toolId) { $wordToolIds += $toolId }
}
if ($wordToolIds.Count -lt 11) {
    $failures.Add("Word tool scan only saw $($wordToolIds.Count) ids; the registry drift check itself is broken")
}
$schemaSource = Get-CodeSource "WordAi\WordJsonCommandSchema.vb"
$whitelistMatch = [regex]::Match($schemaSource, '(?s)SupportedCommands As String\(\) = \{(?<write>.*?)\}.*?SupportedReadCommands As String\(\) = \{(?<read>.*?)\}')
if (-not $whitelistMatch.Success) {
    $failures.Add("Word host whitelist no longer exposes both write and read command arrays: WordAi\WordJsonCommandSchema.vb")
} else {
    $hostAccepted = @()
    foreach ($group in @('write', 'read')) {
        foreach ($m in [regex]::Matches($whitelistMatch.Groups[$group].Value, '"([^"]+)"')) { $hostAccepted += $m.Groups[1].Value }
    }
    foreach ($toolId in $wordToolIds) {
        if ($hostAccepted -notcontains $toolId) {
            $failures.Add("Registry advertises Word tool '$toolId' that the host whitelist rejects as unsupported: WordAi\WordJsonCommandSchema.vb")
        }
    }
}
# Both prompt copies must state the same set, otherwise the model obeys the older one.
foreach ($promptFile in @('WordAi\WordJsonCommandSchema.vb', 'ShareRibbon\Config\PromptManager.vb')) {
    $promptSource = Get-CodeSource $promptFile
    foreach ($readId in @('ListParagraphs', 'GetParagraphInfo')) {
        if ($promptSource -notmatch [regex]::Escape($readId)) {
            $failures.Add("Prompt copy does not declare the read command $readId, so the model is told it does not exist: $promptFile")
        }
    }
}

# P8 guardrail: the default Word skill's allowed-tools is the gate the loop hits last, so a tool the
# host accepts and the prompt advertises must be inside it. DeleteText was not, and the run died with
# TOOL_NOT_ALLOWED after the model had been explicitly told the command exists.
$wordSkillMd = Join-Path $repoRoot 'ShareRibbon\Skills\word-document-agent\SKILL.md'
if (-not (Test-Path -LiteralPath $wordSkillMd)) {
    $failures.Add("Default Word skill disappeared, so the allowed-tools cross-check has nothing to read: ShareRibbon\Skills\word-document-agent\SKILL.md")
} else {
    $skillText = [System.IO.File]::ReadAllText($wordSkillMd, [System.Text.Encoding]::UTF8)
    $allowedMatch = [regex]::Match($skillText, '(?m)^allowed-tools:\s*(.+)$')
    if (-not $allowedMatch.Success) {
        $failures.Add("Default Word skill declares no allowed-tools; an empty list silently turns the tool gate off: ShareRibbon\Skills\word-document-agent\SKILL.md")
    } else {
        $skillAllowed = @($allowedMatch.Groups[1].Value -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        if ($skillAllowed.Count -lt 11) {
            $failures.Add("Default Word skill only lists $($skillAllowed.Count) allowed tools; the allowed-tools drift check itself is broken")
        }
        foreach ($toolId in $wordToolIds) {
            if ($skillAllowed -notcontains $toolId) {
                $failures.Add("Registry and prompts advertise Word tool '$toolId' but the default skill gate blocks it: ShareRibbon\Skills\word-document-agent\SKILL.md")
            }
        }
        foreach ($skillId in $skillAllowed) {
            if ($wordToolIds -notcontains $skillId) {
                $failures.Add("Default Word skill allows '$skillId' which no longer exists in the Word tool registry: ShareRibbon\Skills\word-document-agent\SKILL.md")
            }
        }
    }
}

if ($failures.Count -gt 0) {
    Write-Host "P0 guardrail failures:" -ForegroundColor Red
    foreach ($failure in $failures) {
        Write-Host " - $failure" -ForegroundColor Red
    }
    throw "P0 guardrail audit failed with $($failures.Count) issue(s)."
}

Write-Host "P0 guardrails passed."
Write-Host "Oversized baseline files: $($largeFiles.Count)"
Write-Host "Task.Run + Interop baseline files: $($taskRunInteropFiles.Count)"
