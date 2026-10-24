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
# scan to the next function so the guard does not silently fall out of range when the body grows
$wordBatchEnd = If ($wordBatchStart -lt 0) { -1 } Else { $wordChatSource.IndexOf("Private Function ExecuteWordCommandWithToolResult", $wordBatchStart + 1) }
$wordBatchLength = If ($wordBatchEnd -gt $wordBatchStart) { $wordBatchEnd - $wordBatchStart } ElseIf ($wordBatchStart -ge 0) { [Math]::Min(8000, $wordChatSource.Length - $wordBatchStart) } Else { 0 }
$wordBatchBody = If ($wordBatchStart -lt 0) { "" } Else { $wordChatSource.Substring($wordBatchStart, $wordBatchLength) }
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

# P9 guardrail: Excel and PowerPoint capability sets must agree across the tool registry, the
# host command whitelist, and the default skill gate. The Word checks above showed how a late
# gate turns a valid plan into TOOL_NOT_ALLOWED; the same drift stays invisible until a real run.
$hostCapabilityCases = @(
    @{ App = 'Excel'; ToolDir = 'ShareRibbon\Tools\excel'; Schema = 'ExcelAi\ExcelJsonCommandSchema.vb'; Skill = 'ShareRibbon\Skills\excel-table-agent\SKILL.md'; MinTools = 20 },
    @{ App = 'PowerPoint'; ToolDir = 'ShareRibbon\Tools\ppt'; Schema = 'PowerPointAi\PowerPointJsonCommandSchema.vb'; Skill = 'ShareRibbon\Skills\powerpoint-deck-agent\SKILL.md'; MinTools = 15 }
)
# Registered and whitelisted, but disabled by SafetyGate by default. The skill gate must not
# advertise a fallback the runtime refuses.
$policyDisabledTools = @('ExecuteVBA')
# Routed by the ToolResult entry before the command whitelist, so they are deliberately absent
# from SupportedCommands (and still must be allowed by the default skill).
$hostBypassTools = @('DiscoverOfficeCapability', 'OfficeObjectOperation')

foreach ($case in $hostCapabilityCases) {
    $toolIds = @()
    $toolDir = Join-Path $repoRoot $case.ToolDir
    foreach ($toolFile in @(Get-ChildItem -LiteralPath $toolDir -Filter '*.json' -ErrorAction SilentlyContinue)) {
        $raw = Get-Content -LiteralPath $toolFile.FullName -Raw -Encoding UTF8
        $toolId = [regex]::Match($raw, '"id"\s*:\s*"([^"]+)"').Groups[1].Value
        if ($toolId) { $toolIds += $toolId }
    }
    if ($toolIds.Count -lt $case.MinTools) {
        $failures.Add("$($case.App) tool scan only saw $($toolIds.Count) ids; the registry drift check itself is broken")
    }

    $schemaSource = Get-CodeSource $case.Schema
    $whitelistBody = [regex]::Match($schemaSource, '(?s)SupportedCommands As String\(\) = \{(?<body>.*?)\}').Groups['body'].Value
    if ([string]::IsNullOrWhiteSpace($whitelistBody)) {
        $failures.Add("$($case.App) host whitelist is missing or changed shape: $($case.Schema)")
    } else {
        $hostAccepted = @()
        foreach ($m in [regex]::Matches($whitelistBody, '"([^"]+)"')) { $hostAccepted += $m.Groups[1].Value }
        foreach ($toolId in $toolIds) {
            if ($hostBypassTools -contains $toolId) { continue }
            if ($hostAccepted -notcontains $toolId) {
                $failures.Add("Registry advertises $($case.App) tool '$toolId' that the host whitelist rejects as unsupported: $($case.Schema)")
            }
        }
    }

    $skillPath = Join-Path $repoRoot $case.Skill
    if (-not (Test-Path -LiteralPath $skillPath)) {
        $failures.Add("Default $($case.App) skill disappeared, so the allowed-tools cross-check has nothing to read: $($case.Skill)")
        continue
    }
    $skillText = [System.IO.File]::ReadAllText($skillPath, [System.Text.Encoding]::UTF8)
    $allowedMatch = [regex]::Match($skillText, '(?m)^allowed-tools:\s*(.+)$')
    if (-not $allowedMatch.Success) {
        $failures.Add("Default $($case.App) skill declares no allowed-tools; an empty list silently turns the tool gate off: $($case.Skill)")
        continue
    }
    $skillAllowed = @($allowedMatch.Groups[1].Value -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    foreach ($toolId in $toolIds) {
        if ($policyDisabledTools -contains $toolId) { continue }
        if ($skillAllowed -notcontains $toolId) {
            $failures.Add("Registry advertises $($case.App) tool '$toolId' but the default skill gate blocks it: $($case.Skill)")
        }
    }
    foreach ($skillId in $skillAllowed) {
        if ($toolIds -notcontains $skillId) {
            $failures.Add("Default $($case.App) skill allows '$skillId' which no longer exists in the $($case.App) tool registry: $($case.Skill)")
        }
    }
}

# P10 guardrail: the single-run write change log must stay wired end to end.
# A change log that silently stops recording is worse than none: the UI would tell the user
# "nothing was written" after a partial apply. Contract: openspec/changes/write-trust-contract.md.
$changeLogSource = Get-CodeSource "ShareRibbon\Agent\Execution\AgentChangeLog.vb"
if ($changeLogSource -notmatch 'String\.Equals\(kind, "write"') {
    $failures.Add("Change log no longer recognizes host write observations: ShareRibbon\Agent\Execution\AgentChangeLog.vb")
}
if ($changeLogSource -notmatch '"code_execution"') {
    $failures.Add("Change log no longer records script/VBA writes: ShareRibbon\Agent\Execution\AgentChangeLog.vb")
}
if (($changeLogSource + (Get-CodeSource "ShareRibbon\Storage\AgentChangeLogRepository.vb")) -match "Base64|DataUrl|VisualEvidence") {
    $failures.Add("Change log must not persist visual evidence payloads")
}
if ($loopSource -notmatch "session\.ChangeLog\.AddFromStep\(") {
    $failures.Add("LoopEngine no longer records write steps into the run change log")
}
$agentKernelSource = Get-CodeSource "ShareRibbon\Agent\AgentKernel.vb"
if ($agentKernelSource -notmatch "result\.ChangeLog = _session\.ChangeLog") {
    $failures.Add("AgentResult is not carrying the run change log: ShareRibbon\Agent\AgentKernel.vb")
}
$harnessSource = Get-CodeSource "ShareRibbon\Agent\Harness\OfficeHarness.vb"
if ([regex]::Matches($harnessSource, "SafePersistChangeLog\(").Count -lt 3) {
    $failures.Add("Both OfficeHarness completion paths must persist the change log (call + definition expected)")
}
$changeRepoSource = Get-CodeSource "ShareRibbon\Storage\AgentChangeLogRepository.vb"
if ($changeRepoSource -notmatch "INSERT OR REPLACE INTO agent_run_change") {
    $failures.Add("Change log repository no longer writes agent_run_change rows")
}
if ($changeRepoSource -notmatch "Function GetByRun\(") {
    $failures.Add("Change log repository lost its per-run read entry: AgentChangeLogRepository.GetByRun")
}
if ($codeExecSource -notmatch '\.kind = "code_execution"') {
    $failures.Add("Script execution observation lost its kind marker, so script writes cannot enter the change log: CodeExecutionService.vb")
}
if ($codeExecSource -notmatch "\.changed = ok") {
    $failures.Add("Script execution observation no longer declares whether the document changed: CodeExecutionService.vb")
}
if (-not (Test-Path -LiteralPath (Join-Path $repoRoot "ShareRibbon\Storage\Migrations\015_agent_run_change_log.sql"))) {
    $failures.Add("Change log migration is missing: ShareRibbon\Storage\Migrations\015_agent_run_change_log.sql")
}
if (-not (Test-Path -LiteralPath (Join-Path $repoRoot "ShareRibbon\Storage\Migrations\016_agent_run_change_undo.sql"))) {
    $failures.Add("Undo migration is missing: ShareRibbon\Storage\Migrations\016_agent_run_change_undo.sql")
}
# E1-03/E1-04 guardrail: the end-of-run message and the one-click undo must both derive from the
# persisted change log, never from an optimistic in-memory count.
if ($changeRepoSource -notmatch "Function GetUndoScopeByRun\(") {
    $failures.Add("Change log repository lost its run undo scope query: AgentChangeLogRepository.GetUndoScopeByRun")
}
if ($loopSource -notmatch "Function AppendWriteSummary\(" -or $loopSource -notmatch "WrittenCount") {
    $failures.Add("LoopEngine no longer reports how much was written on cancel/failure")
}
$kernelServiceSource = Get-CodeSource "ShareRibbon\Controls\Services\AgentKernelService.vb"
if ($kernelServiceSource -notmatch "AgentChangeLogRepository\.SaveRun\(") {
    $failures.Add("AgentKernelService does not persist the change log before rendering the terminal UI")
}
if ($kernelServiceSource -notmatch "Function BuildUndoDescriptor\(" -or $kernelServiceSource -notmatch "completeAgent\([\s\S]{0,300}?undoJson") {
    $failures.Add("AgentKernelService no longer passes the undo descriptor to completeAgent")
}
if ($baseChatSource -notmatch 'Register\("agent:undoRun"') {
    $failures.Add("BaseChatControl no longer routes agent:undoRun")
}
if ($wordChatSource -notmatch "Function PerformUndoAgentRun[\s\S]{0,2400}?GetUndoScopeByRun") {
    $failures.Add("Word run undo does not read the persisted undo scope: WordAi\ChatControl.vb")
}
if ($wordChatSource -notmatch "Function PerformUndoAgentRun[\s\S]{0,3000}?scope\.BaselineHash") {
    $failures.Add("Word run undo does not verify the restored document fingerprint: WordAi\ChatControl.vb")
}
if ($wordChatSource -notmatch "Function PerformUndoAgentRun[\s\S]{0,3000}?MarkRunRolledBack") {
    $failures.Add("Word run undo does not mark the run as rolled back: WordAi\ChatControl.vb")
}
$agentCardSource = Get-CodeSource "ShareRibbon\Resources\js\agent-card.js"
if ($agentCardSource -notmatch "undo\.canUndo" -or $agentCardSource -notmatch "function setAgentUndoResult") {
    $failures.Add("Agent card no longer renders or resolves the run undo entry: agent-card.js")
}

# E1-05 guardrail: approval must offer once / session / reject, and a rejection may only skip the
# current step instead of silently executing or killing the whole plan.
$toolContextSource = Get-CodeSource "ShareRibbon\Agent\ToolExecutionContext.vb"
$toolRegistrySource = Get-CodeSource "ShareRibbon\Agent\ToolRegistry.vb"
if ($loopSource -notmatch "ApprovalDecision\.Session" -or $loopSource -notmatch "ApproveToolForSession\(") {
    $failures.Add("LoopEngine no longer honors the session-scoped approval decision")
}
if ($loopSource -notmatch "Function IsRejectedApproval\(" -or $loopSource -notmatch "StepStatus\.Skipped") {
    $failures.Add("LoopEngine no longer limits a rejected approval to the current step")
}
if ($harnessSource -notmatch "Function DecideAsync\(" -or $harnessSource -notmatch "ApprovalDecision\.Session") {
    $failures.Add("OfficeHarness lost the three-state approval decision channel")
}
if ($toolContextSource -notmatch "Function IsToolApprovedForSession\(" -or
    $toolContextSource -notmatch "Sub ApproveToolForSession\(") {
    $failures.Add("ToolExecutionContext no longer keeps non-consuming session approvals")
}
if ($toolRegistrySource -notmatch "IsToolApprovedForSession\(") {
    $failures.Add("ToolRegistry can no longer bypass repeated prompts for a session-approved tool")
}
if ($baseChatSource -notmatch 'Register\("agent:approveToolSession"') {
    $failures.Add("BaseChatControl no longer routes agent:approveToolSession")
}
if ($agentCardSource -notmatch "function agentApproveSession" -or
    $agentCardSource -notmatch "agent-btn-approve-session") {
    $failures.Add("Agent card no longer offers the session-scoped approval button: agent-card.js")
}

# E1-05b guardrail: timeout / cancel / supersede must abort the whole round, and only Once/Session may
# unlock a write. Treating Abort as allowed would execute an unconfirmed write after the user walked away.
$harnessApprovalsSource = Get-CodeSource "ShareRibbon\Agent\Harness\OfficeHarness.Approvals.vb"
if ($loopSource -notmatch "Function IsAbortedApproval\(" -or $loopSource -notmatch '"aborted"') {
    $failures.Add("LoopEngine can no longer stop the whole round when the approval channel aborts")
}
if ($harnessSource -notmatch "Dim allowed = decision = ApprovalDecision\.Once OrElse decision = ApprovalDecision\.Session") {
    $failures.Add("OfficeHarness.DecideAsync no longer treats only Once/Session as allowed")
}
if ($harnessSource -notmatch "SettleApprovalAsync\(superseded,\s*ApprovalDecision\.Abort") {
    $failures.Add("OfficeHarness no longer aborts the stale approval when a new turn supersedes it")
}
if ($harnessSource -notmatch "DecideAsync\(runId, ApprovalDecision\.Abort") {
    $failures.Add("OfficeHarness.CancelAsync no longer aborts a pending approval")
}
if ($harnessApprovalsSource -notmatch "ApprovalDecision\.Abort") {
    $failures.Add("The approval timeout no longer aborts the round: OfficeHarness.Approvals.vb")
}

# E1-05c guardrail: the WebView2 panel runs the JS/CSS extracted into %LOCALAPPDATA%\OfficeAI\www, not the
# bytes inside the DLL. A version-stamp skip there pins the UI to an old build with zero errors, so refresh
# must be decided by content comparison.
$extractorSource = Get-CodeSource "ShareRibbon\Resources\ResourceExtractor.vb"
if ($extractorSource -notmatch "Function ByteArrayEquals\(" -or $extractorSource -notmatch "ByRef refreshed As Integer") {
    $failures.Add("ResourceExtractor no longer refreshes static files by content comparison")
}
if ($extractorSource -match '\.version') {
    $failures.Add("ResourceExtractor went back to a version-stamp skip: stale frontend fails silently")
}

# E2-01 guardrail: every chat turn is stamped with the host that wrote it, and history rows
# created before app_type existed must stay visible instead of vanishing after the upgrade.
$convSource = Get-CodeSource "ShareRibbon\Storage\ConversationRepository.vb"
$chatStateSource = Get-CodeSource "ShareRibbon\Controls\Services\ChatStateService.vb"
$historyServiceSource = Get-CodeSource "ShareRibbon\Controls\Services\HistorySessionService.vb"
if ($convSource -notmatch "INSERT INTO conversation_event \(event_id, session_id, app_type,") {
    $failures.Add("ConversationRepository no longer stamps the host app type on a chat message")
}
if ($convSource -notmatch "IFNULL\(e\.app_type, ''\) = ''") {
    $failures.Add("GetRecentSessions lost the legacy-row exemption, pre-migration history disappears from the sidebar")
}
if ($chatStateSource -notmatch "ConversationRepository\.InsertMessage\(sid, r, c, False, app,") {
    $failures.Add("ChatStateService no longer passes the host app type when persisting a turn: ChatStateService.vb")
}
if ($baseChatSource -notmatch "_chatStateService\.AppTypeProvider = AddressOf GetOfficeAppType") {
    $failures.Add("BaseChatControl no longer supplies the host app type that scopes history")
}

# E2-02 guardrail: a session belongs to the document that created it. The binding is written once
# (first turn wins, later turns must not move the session) and the document identity is read on the
# UI thread only, because persistence happens on a background thread where COM is cross-apartment.
if (-not (Test-Path -LiteralPath (Join-Path $repoRoot "ShareRibbon\Storage\Migrations\017_session_document_binding.sql"))) {
    $failures.Add("Document binding migration is missing: ShareRibbon\Storage\Migrations\017_session_document_binding.sql")
}
$dbSource = Get-CodeSource "ShareRibbon\Storage\OfficeAiDatabase.vb"
if ($dbSource -notmatch "\{\s*17,") {
    $failures.Add("OfficeAiDatabase has no versioned migration registered for schema 17")
}
if ($dbSource -notmatch "Sub EnsureSessionProfileSchema\(") {
    $failures.Add("OfficeAiDatabase lost the self-healing pass for a half-migrated session profile table")
}
if ($convSource -notmatch "Sub UpsertSessionProfile\(") {
    $failures.Add("ConversationRepository no longer maintains the one-row-per-session profile")
}
if ($convSource -notmatch "CASE WHEN IFNULL\(document_id, ''\) = '' THEN") {
    $failures.Add("Session document binding is no longer first-write-wins, so a later turn could move the session")
}
if ($convSource -notmatch "ORDER BY \(IFNULL\(MAX\(p\.document_id\), ''\) = @doc\) DESC") {
    $failures.Add("GetRecentSessions no longer ranks the current document first in all scope")
}
if ($convSource -notmatch "SessionScopeDocument") {
    $failures.Add("The document-only history scope disappeared from ConversationRepository")
}
if ($convSource -notmatch "DELETE FROM session_summary WHERE session_id = @sid") {
    $failures.Add("DeleteSessionCascade no longer removes the session profile row")
}
$sessionDocSource = Get-CodeSource "ShareRibbon\Controls\Models\SessionDocument.vb"
if ($sessionDocSource -notmatch "Function FromFullPath\(") {
    $failures.Add("SessionDocument no longer derives the document identity from the host path")
}
if ($sessionDocSource -notmatch "If String\.IsNullOrEmpty\(directory\) Then Return doc") {
    $failures.Add("SessionDocument lost the unsaved-document guard: two new documents would share one session bucket")
}
if ($sessionDocSource -notmatch "ToLowerInvariant\(\)") {
    $failures.Add("The document id must be case-insensitive or a rename alone splits the history")
}
foreach ($hostChat in @(
        @("WordAi\ChatControl.vb", $wordChatSource),
        @("ExcelAi\ChatControl.vb", $excelChatSource),
        @("PowerPointAi\ChatControl.vb", $pptChatSource)
    )) {
    if ($hostChat[1] -notmatch "Overrides Function GetSessionDocument\(\)") {
        $failures.Add("$($hostChat[0]) no longer reports the active document for session binding")
    }
}
if ($baseChatSource -notmatch "AddressOf ResolveSessionDocument") {
    $failures.Add("HistorySessionService is no longer wired to the host document resolver")
}
if ([regex]::Matches($baseChatSource, "RefreshSessionDocumentContext\(\)").Count -lt 4) {
    $failures.Add("Every UI-thread persistence entry must refresh the session document context before handing work to a background thread")
}
if ($historyServiceSource -notmatch "GetRecentSessions\(") {
    $failures.Add("HandleGetSessionList no longer asks for the host/document scoped session list")
}
$historyJsSource = Get-JsSource (Join-Path $jsResourceDir "history-manager.js")
if ($historyJsSource -notmatch "data-scope") {
    $failures.Add("The history sidebar no longer exposes the current-document / all scope switch")
}
if ($historyJsSource -match "this\.items\.sort\(") {
    $failures.Add("history-manager.js re-sorts sessions and wipes the current-document-first order from the backend")
}

# E2-03 guardrail: the sidebar title must stay readable without spending a model call, a user rename
# must win over the auto title, and delete must be confirmed and cascade through the repository.
if ($convSource -notmatch "Function AutoTitle\(") {
    $failures.Add("ConversationRepository no longer derives a readable auto title from the first user turn")
}
if ($convSource -notmatch 'Regex\.Replace' -or $convSource -notmatch '\\s\+') {
    $failures.Add("AutoTitle no longer collapses whitespace, a pasted block becomes a one-line wall")
}
if ($convSource -notmatch "If\(String\.IsNullOrEmpty\(customTitle\), derivedTitle, customTitle\)") {
    $failures.Add("A user rename no longer overrides the auto title")
}
if ($convSource -notmatch "Function RenameSession\(") {
    $failures.Add("ConversationRepository lost RenameSession, the rename box would silently do nothing")
}
if ($baseChatSource -notmatch 'Register\("renameSession"') {
    $failures.Add("BaseChatControl no longer routes renameSession")
}
if ($historyServiceSource -notmatch "Sub HandleRenameSession\(" -or $historyServiceSource -notmatch "Sub HandleDeleteSession\(") {
    $failures.Add("HistorySessionService lost the rename or delete handler")
}
if ($historyJsSource -notmatch "type: 'renameSession'") {
    $failures.Add("history-manager.js no longer submits a rename request")
}
if ($historyJsSource -notmatch "list\.dataset\.bound") {
    $failures.Add("History row actions are bound per item again: every refresh would stack duplicate listeners")
}
if ($historyJsSource -notmatch "confirm\(") {
    $failures.Add("Deleting a session cascades to messages, memory and run traces, so it must stay confirmed")
}

# Memory pipeline contract guardrail: the queue job is keyed by session id and both the history read
# path and the extractor read path filter conversation_event on event_type = 'chat_message'. When one
# side drifts the job completes while extracting nothing, which is silent data loss, not an error.
$memoryPipelineSource = Get-CodeSource "ShareRibbon\Services\Memory\AgentMemoryPipelineService.vb"
$memoryRecorderSource = Get-CodeSource "ShareRibbon\Controls\Services\MemoryTurnRecorder.vb"
$agentMemRepoSource = Get-CodeSource "ShareRibbon\Storage\AgentMemoryRepository.vb"
if ($memoryRecorderSource -notmatch "\.TargetId = capturedSessionId") {
    $failures.Add("MemoryTurnRecorder no longer keys the extract_memory job by session id")
}
if ($memoryPipelineSource -notmatch "GetUnprocessedConversationEventsBySession\(sessionId\)") {
    $failures.Add("The extract_memory job no longer resolves events through the session, so it reads an empty event set")
}
if ($agentMemRepoSource -notmatch "event_type = 'chat_message' AND processed_at IS NULL") {
    $failures.Add("Unprocessed event selection no longer matches the single chat history event type")
}

# E3-03 guardrail: an unconfigured model must be blocked inside the shared send funnel and answered with
# a clickable in-panel CTA. A status-bar string is not an actionable next step, and the Agent entry must
# not keep a weaker second check, because half-started agent runs leave an empty thinking bubble behind.
$gateSource = Get-CodeSource "ShareRibbon\Controls\Services\ModelConfigGate.vb"
$sendValidatorSource = Get-CodeSource "ShareRibbon\Controls\Services\ChatSendValidator.vb"
$chatManagerJsSource = Get-JsSource (Join-Path $jsResourceDir "chat-manager.js")
if ($sendValidatorSource -notmatch "Function ValidateConfig\(") {
    $failures.Add("ChatSendValidator no longer separates the config check from the question check")
}
if ($sendValidatorSource -notmatch "ValidateConfig\(apiUrl, apiKey\)") {
    $failures.Add("Validate() no longer reuses the config check, so the chat and agent gates can drift")
}
if ($gateSource -notmatch '"missing_api_key"' -or $gateSource -notmatch '"missing_api_url"') {
    $failures.Add("ModelConfigGate lost the reason codes, the panel can no longer tell key from url")
}
if ($gateSource -notmatch "JObject|ToJson") {
    $failures.Add("ModelConfigGate no longer serialises the CTA payload through JSON, hand-escaped JS is unsafe")
}
if ($baseChatSource -notmatch "IsBlockedByModelConfig\(ModelConfigGate\.ActionChat, validation\)") {
    $failures.Add("Send no longer routes a missing model config to the in-panel CTA")
}
if ($baseChatSource -notmatch "IsBlockedByModelConfig\(ModelConfigGate\.ActionAgent\)") {
    $failures.Add("HandleStartAgent can still start the agent chain without an API key")
}
if ($baseChatSource -notmatch "BuildHideScript\(\)") {
    $failures.Add("Saving the config form no longer dismisses the stale configuration CTA")
}
if ($chatManagerJsSource -notmatch "window\.showModelConfigRequired") {
    $failures.Add("chat-manager.js lost the configuration CTA renderer")
}
if ($chatManagerJsSource -notmatch "window\.hideModelConfigRequired") {
    $failures.Add("chat-manager.js can no longer remove the configuration CTA")
}
if ($chatManagerJsSource -notmatch "model-config-cta") {
    $failures.Add("The configuration CTA is rebuilt with an unstable id, so repeated blocked sends stack cards")
}
if ($chatManagerJsSource -notmatch "openModelConfig") {
    $failures.Add("The configuration CTA no longer reuses the existing openApiConfigForm bridge")
}

# E3-03b guardrail: ConfigSettings is the process-wide url/key snapshot every send reads. It used to be
# written only inside an "if selected" loop at load time, so clearing a url or key left the previous
# endpoint live and the gate was satisfied by a stale key. The projection must be unconditional and must
# run after every write to the store, otherwise the panel and the config file disagree for the whole session.
$configManagerSource = Get-CodeSource "ShareRibbon\Config\ConfigManager.vb"
if ($configManagerSource -notmatch "Public Shared Sub ApplyConfigSnapshot") {
    $failures.Add("ConfigManager lost ApplyConfigSnapshot, so the global config snapshot has no single writer")
}
if ($configManagerSource -notmatch 'ConfigSettings\.ApiUrl = If\(chosen\?\.url, ""\)') {
    $failures.Add("ApplyConfigSnapshot no longer assigns the url unconditionally, a cleared url can stay live")
}
if ($configManagerSource -notmatch 'ConfigSettings\.ApiKey = If\(chosen\?\.key, ""\)') {
    $failures.Add("ApplyConfigSnapshot no longer assigns the key unconditionally, a cleared key can stay live")
}
if ($configManagerSource -match "ConfigSettings\.ApiUrl = item\.url") {
    $failures.Add("LoadConfig reverted to assigning the snapshot inside a selected-only loop")
}
$loadConfigBody = [regex]::Match($configManagerSource, "Public Sub LoadConfig[\s\S]*?\r?\n    End Sub").Value
if ($loadConfigBody -notmatch "ApplyConfigSnapshot\(\)") {
    $failures.Add("LoadConfig no longer projects the snapshot, startup can keep an old endpoint")
}
$saveConfigBody = [regex]::Match($configManagerSource, "Public Shared Sub SaveConfig[\s\S]*?\r?\n    End Sub").Value
if ($saveConfigBody -notmatch "ApplyConfigSnapshot\(\)") {
    $failures.Add("SaveConfig writes the store but leaves the live snapshot stale, so a saved config change needs a restart")
}
$checksSource = Get-Content -LiteralPath (Join-Path $repoRoot "scripts\run-code-checks.ps1") -Raw -Encoding UTF8
if ($checksSource -notmatch [regex]::Escape("smoke-config-snapshot.ps1")) {
    $failures.Add("smoke-config-snapshot.ps1 is not registered in run-code-checks.ps1, the stale-snapshot defect can return")
}

# E3-01 guardrail: the first-run wizard may only carry what the config window never had - the privacy
# acknowledgement and the example tasks. Choosing a provider, typing an api key and probing the connection
# already live in ConfigApiForm; re-implementing them in the panel creates a second truth about the same key.
$onboardingSource = Get-CodeSource "ShareRibbon\Controls\Services\OnboardingService.vb"
$configApiSource = Get-CodeSource "ShareRibbon\Config\ConfigApiForm.vb"
$configManagerSource = Get-CodeSource "ShareRibbon\Config\ConfigManager.vb"
$onboardingJsSource = Get-JsSource (Join-Path $jsResourceDir "onboarding.js")

if ($baseChatSource -notmatch "New OnboardingService\(") {
    $failures.Add("BaseChatControl no longer wires the onboarding service, so the wizard has no host behind it")
}
foreach ($onboardingCommand in @("getOnboardingState", "onboardingSetPrivacy", "onboardingFinish")) {
    if ($baseChatSource -notmatch [regex]::Escape("Register(`"$onboardingCommand`"")) {
        $failures.Add("The wizard posts '$onboardingCommand' but the host router no longer registers it")
    }
}

# Anti-regrowth: the moment the wizard writes config or probes a socket itself, it is a second config layer.
foreach ($forbidden in @("ConfigManager\.SaveConfig", "ModelConnectivityTester", "SecureConfigIntegration",
                         'ConfigSettings\.(ApiUrl|ApiKey|platform|ModelName)\s*=', "ConfigManager\.ConfigData")) {
    if ($onboardingSource -match $forbidden) {
        $failures.Add("OnboardingService touches configuration itself ($forbidden) again instead of routing the user to ConfigApiForm")
    }
}
if ($onboardingSource -notmatch "ChatSendValidator\(\)\.ValidateConfig\(ConfigSettings\.ApiUrl, ConfigSettings\.ApiKey\)") {
    $failures.Add("The wizard decides on its own whether a model is configured, so it can disagree with the send gate")
}
foreach ($wizardStep in @("privacy", "example")) {
    if ($onboardingSource -notmatch [regex]::Escape("Return `"$wizardStep`"")) {
        $failures.Add("InitialStep no longer produces the '$wizardStep' step the panel renders")
    }
}
if ($onboardingSource -match 'Return "(provider|key|test)"') {
    $failures.Add("InitialStep can still hand out a provider/key/test step, but the wizard deleted the renderers for them")
}

# The removed message types must stay removed, on both sides of the bridge.
foreach ($deadPost in @("onboardingSelectProvider", "onboardingSaveKey", "onboardingTestConnection")) {
    if ($baseChatSource -match [regex]::Escape("Register(`"$deadPost`"")) {
        $failures.Add("The '$deadPost' handler is registered again: provider/key/test belongs to ConfigApiForm only")
    }
    if ($onboardingJsSource -match $deadPost) {
        $failures.Add("onboarding.js still posts '$deadPost', which recreates the second config implementation")
    }
}

if ($onboardingJsSource -notmatch "window\.renderOnboarding = function") {
    $failures.Add("onboarding.js lost the renderer the host pushes into")
}
if ($onboardingJsSource -notmatch "window\.showOnboarding = function") {
    $failures.Add("onboarding.js lost the entry point the configuration CTA hands over to")
}
if ($onboardingJsSource -notmatch "openModelConfig") {
    $failures.Add("onboarding.js no longer opens the existing config window, so the wizard has no way to configure")
}
if ($onboardingJsSource -match "onboardingStepPinned|window\.onboardingStep\s*=") {
    $failures.Add("onboarding.js keeps a front-end copy of the wizard step again, which can contradict the host state")
}
if ($chatManagerJsSource -notmatch "showOnboarding") {
    $failures.Add("The blocked-send CTA no longer offers the guided setup path (E3-03 -> E3-01 handover)")
}
if ($chatManagerJsSource -match "showOnboarding\('[a-z]") {
    $failures.Add("The CTA passes a wizard step into showOnboarding, but only the host state decides which step shows")
}

# A stale wizard card after a successful save is the same class of bug as a stale config snapshot.
$configFormHandler = [regex]::Match($baseChatSource, "Protected Sub HandleOpenApiConfigForm[\s\S]*?\n    End Sub").Value
if ($configFormHandler -notmatch "ModelConfigGate\.BuildHideScript\(\)") {
    $failures.Add("Saving the config form no longer dismisses the stale configuration CTA")
}
if ($configFormHandler -notmatch "OnboardingSvc\.HandleGetState\(\)") {
    $failures.Add("Closing the config form no longer refreshes the wizard, so it keeps claiming nothing is configured")
}

# One connectivity probe, one caller that shows the verdict: the config window must explain the failure.
if ($configApiSource -notmatch "New ModelConnectivityTester\(\)") {
    $failures.Add("ConfigApiForm no longer delegates the test button to the shared connectivity tester")
}
if ($configApiSource -notmatch "As Task\(Of ModelConnectivityResult\)") {
    $failures.Add("ConfigApiForm collapses the probe result to a boolean again, so it cannot say why the test failed")
}
if ($configApiSource -notmatch "ValidationFailureText\(validationResult\)") {
    $failures.Add("The save path no longer shows the tester reason/detail to the user")
}
if ($configApiSource -match "max_tokens") {
    $failures.Add("ConfigApiForm builds its own probe request again, so the panel and the form can disagree")
}
if ($configApiSource -match 'ConfigSettings\.(ApiUrl|ApiKey|platform|ModelName)\s*=') {
    $failures.Add("ConfigApiForm hand-writes the global config snapshot again; SaveConfig -> ApplyConfigSnapshot is the only writer")
}

# The wizard writes through ConfigManager, so at-rest protection has to live in that one place.
if ($configManagerSource -notmatch "SecureConfigManager\.SaveApiKey\(item\.key, item\.platform\)") {
    $failures.Add("SaveConfig writes api keys again without routing them into the current-user encrypted store")
}
if ($configManagerSource -notmatch '\[encrypted\]') {
    $failures.Add("SaveConfig no longer keeps the encrypted placeholder in the config file")
}
if ($onboardingSource -match "File\.WriteAllText|SqliteConnection|DbContext") {
    $failures.Add("OnboardingService persists config outside ConfigManager, creating a second source of truth")
}

# Deleted steps must not leave dead styles behind - the wizard is small, so every onboarding rule has to
# still be produced by onboarding.js.
$stylesCssSource = Get-Content -LiteralPath (Join-Path $repoRoot "ShareRibbon\Resources\css\styles.css") -Raw -Encoding UTF8
foreach ($rule in [regex]::Matches($stylesCssSource, '\.onboarding-([a-z][a-z-]*)')) {
    if ($onboardingJsSource -notmatch [regex]::Escape($rule.Groups[1].Value)) {
        $failures.Add("styles.css styles '.onboarding-$($rule.Groups[1].Value)' but onboarding.js no longer renders it")
    }
}

# Static resource registration: a new js file that is not registered in all four places never reaches the
# WebView, and the panel fails with a missing function instead of a build error.
$resxSource = Get-Content -LiteralPath (Join-Path $repoRoot "ShareRibbon\My Project\Resources.resx") -Raw -Encoding UTF8
$extractorSource = Get-CodeSource "ShareRibbon\Resources\ResourceExtractor.vb"
$templateFiles = @(Get-ChildItem -LiteralPath (Join-Path $repoRoot "ShareRibbon\Resources") -Filter *.html -File)
foreach ($js in $jsFiles) {
    if ($resxSource -notmatch [regex]::Escape($js.Name)) {
        $failures.Add("$($js.Name) is not registered in ShareRibbon\My Project\Resources.resx")
    }
    if ($extractorSource -notmatch [regex]::Escape($js.Name)) {
        $failures.Add("$($js.Name) is not in the ResourceExtractor map, so it never reaches %LOCALAPPDATA%\OfficeAI\www")
    }
}

# Bridge contract: every message type the panel posts must have a host handler, otherwise the click is a
# no-op that only shows up during manual testing.
$handledTypes = New-Object 'System.Collections.Generic.HashSet[string]'
$codeProjectDirs = @("ShareRibbon", "WordAi", "ExcelAi", "PowerPointAi")
foreach ($dir in $codeProjectDirs) {
    foreach ($vbFile in (Get-ChildItem -LiteralPath (Join-Path $repoRoot $dir) -Filter *.vb -Recurse |
        Where-Object { $_.FullName -notmatch "[\\/](bin|obj)[\\/]" })) {
        $vbText = Get-Content -LiteralPath $vbFile.FullName -Raw -Encoding UTF8
        foreach ($m in [regex]::Matches($vbText, 'Register\(\s*"([A-Za-z][A-Za-z0-9_]*)"')) {
            [void]$handledTypes.Add($m.Groups[1].Value)
        }
        foreach ($m in [regex]::Matches($vbText, 'Case\s+"([A-Za-z][A-Za-z0-9_]*)"')) {
            [void]$handledTypes.Add($m.Groups[1].Value)
        }
    }
}

$postedTypes = New-Object 'System.Collections.Generic.HashSet[string]'
$bridgeJsFiles = @($jsFiles) + @($templateFiles)
foreach ($file in $bridgeJsFiles) {
    $text = Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8
    foreach ($m in [regex]::Matches($text, "type\s*:\s*['""]([A-Za-z][A-Za-z0-9_]*)['""]")) {
        [void]$postedTypes.Add($m.Groups[1].Value)
    }
}
if ($postedTypes.Count -lt 40) {
    $failures.Add("Bridge post scan only saw $($postedTypes.Count) message types; the scan is broken, not the bridge")
}
foreach ($posted in ($postedTypes | Sort-Object)) {
    if (-not $handledTypes.Contains($posted)) {
        $failures.Add("The panel posts message type '$posted' but no host handler registers it")
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
