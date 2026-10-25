' ShareRibbon\Agent\TaskPackageCatalog.vb
' E6-02 任务包合同：一个任务包一份 JSON，只承载 plan §3 模板里目前没有任何持有者的字段
' （预览策略、验证器、成功标准、交付物、解释模板、失败恢复、黄金样例引用、已有实现）。
' 入口话术归 QuickActionCatalog、工具边界归 Skill 的 allowed-tools、capability 归宿主注册表、
' 变更事实归 AgentRunChangeLog——本文件只引用它们的 id，不复制第二份清单，
' 所以这里连 allowedTools 字段都不设。引用能否解析由 scripts/check-task-package-contract.ps1 逐级核对，
' 断裂即构建失败；运行时这里只再校验一次入口动作，防止"面板说得到、目录里根本没有"。
' 词表（字段名、枚举值、断言事实）在这里是唯一定义处，门禁脚本从本文件和 AgentChangeLog.vb 反查，不再自带名单。

Imports System.IO
Imports System.Reflection
Imports Newtonsoft.Json.Linq

Namespace Agent


    ''' <summary>
    ''' 任务包目录：加载即严格校验，不合格的包不入库并记进 ValidationErrors。
    ''' 未声明的包（前 6 包里 Word 模板排版、PPT 成稿、PPT 品牌模板）不在这里补占位文件——
    ''' E6-01/E6-07 之前它们没有可解析入口，"先建壳"正是 plan §1 明令禁止的假能力。
    ''' 严格校验规则本身在 TaskPackageCatalog.Parsing.vb，本文件只持有词表、候选目录和查询入口。
    ''' </summary>
    Public NotInheritable Partial Class TaskPackageCatalog

        Private Const AllowedSchemaVersion As Integer = 1

        ' 以下六张词表是任务包 JSON 的唯一 schema，门禁脚本按名字从这里反查，不允许自己另写一份。
        Private Shared ReadOnly AllowedTopLevelKeys As IReadOnlyCollection(Of String) = ToSet(
            "schemaVersion", "id", "title", "application", "entryActions", "capabilityIds",
            "skill", "previewPolicy", "verifier", "successCriteria", "deliverable",
            "explanation", "recovery", "goldenCases", "existingImplementation")

        Private Shared ReadOnly AllowedPreviewKeys As IReadOnlyCollection(Of String) = ToSet(
            "mode", "requireBeforeWrite", "note")

        Private Shared ReadOnly AllowedVerifierKeys As IReadOnlyCollection(Of String) = ToSet(
            "assertions", "onFail")

        Private Shared ReadOnly AllowedAssertionKeys As IReadOnlyCollection(Of String) = ToSet(
            "fact", "op", "value", "why")

        Private Shared ReadOnly AllowedExplanationKeys As IReadOnlyCollection(Of String) = ToSet(
            "whatChanged", "whyHint", "undoHint")

        Private Shared ReadOnly AllowedRecoveryKeys As IReadOnlyCollection(Of String) = ToSet(
            "strategy", "notes")

        Private Shared ReadOnly AllowedImplementationKeys As IReadOnlyCollection(Of String) = ToSet(
            "reuse", "gap")

        Private Shared ReadOnly AllowedPreviewModes As IReadOnlyCollection(Of String) = ToSet(
            "none", "summary", "inDocumentHighlight", "thumbnail")

        Private Shared ReadOnly AllowedOnFailStrategies As IReadOnlyCollection(Of String) = ToSet(
            "repair", "rollback", "report")

        Private Shared ReadOnly AllowedRecoveryStrategies As IReadOnlyCollection(Of String) = ToSet(
            "repair", "rollback", "degrade", "report")

        Private Shared ReadOnly AllowedOps As IReadOnlyCollection(Of String) = ToSet(
            "eq", "gte", "lte", "isTrue", "isFalse")

        Private Shared ReadOnly AllowedApplications As IReadOnlyCollection(Of String) = ToSet(
            "Word", "Excel", "PowerPoint")

        Private Shared ReadOnly _packages As New List(Of TaskPackage)()

        Private Shared ReadOnly _validationErrors As New List(Of String)()

        ''' <summary>每个 id 第一次见到的合同正文，用于发现「源码目录与构建输出目录内容不一致」这种漂移。</summary>
        Private Shared ReadOnly _sourceTexts As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        ''' <summary>是否已经跑过一次候选目录加载；探针脚本用 Reset 把它拨回未加载状态。</summary>
        Private Shared _loaded As Boolean

        Private Sub New()
        End Sub

        ''' <summary>
        ''' 断言可用的事实词表：直接从 AgentRunChangeLog 的公开 Integer/Boolean 只读属性推导，
        ''' 让"事实"这件事的唯一持有者仍是变更日志本身；改了属性名，合同自动跟着收紧。
        ''' 事实名的写法也取自变更日志（ToFactName），运行时读取与词表因此不可能各写一套。
        ''' </summary>
        Public Shared ReadOnly Property VerifiableFacts As IReadOnlyCollection(Of String) =
            BuildVerifiableFacts()

        Private Shared Function BuildVerifiableFacts() As IReadOnlyCollection(Of String)
            Dim names = GetType(AgentRunChangeLog).
                GetProperties(BindingFlags.Public Or BindingFlags.Instance).
                Where(Function(p) p.CanRead AndAlso Not p.GetIndexParameters().Any()).
                Select(Function(p) New With {
                    .Name = AgentRunChangeLog.ToFactName(p.Name),
                    .IsNumeric = (p.PropertyType Is GetType(Integer) OrElse p.PropertyType Is GetType(Long)),
                    .IsBoolean = (p.PropertyType Is GetType(Boolean))
                }).
                Where(Function(item) item.IsNumeric OrElse item.IsBoolean).
                Select(Function(item) item.Name).
                Distinct().
                ToList()
            Return New List(Of String)(names).AsReadOnly()
        End Function

        Private Shared Function ToSet(ParamArray values() As String) As IReadOnlyCollection(Of String)
            Return New List(Of String)(values).AsReadOnly()
        End Function

        Public Shared ReadOnly Property Packages As IReadOnlyList(Of TaskPackage)
            Get
                Return _packages.AsReadOnly()
            End Get
        End Property

        ''' <summary>加载期被拒绝的包及原因。调用方（面板/冒烟脚本）要把非空结果显式暴露出来，不能吞掉。</summary>
        Public Shared ReadOnly Property ValidationErrors As IReadOnlyList(Of String)
            Get
                Return _validationErrors.AsReadOnly()
            End Get
        End Property

        Public Shared ReadOnly Property PackageCount As Integer
            Get
                Return _packages.Count
            End Get
        End Property

        ''' <summary>
        ''' 从 VSTO 宿主输出、共享程序集输出和开发仓库候选目录加载任务包合同。
        ''' 候选目录算法与 ToolRegistry.LoadFromRuntimeDirectories 保持一致，避免两套运行时视图。
        ''' </summary>
        Public Shared Function LoadFromRuntimeDirectories(Optional preferredBaseDirectory As String = Nothing) As Integer
            Reset()
            Dim roots As New List(Of String) From {
                preferredBaseDirectory,
                Path.GetDirectoryName(GetType(TaskPackageCatalog).Assembly.Location),
                AppDomain.CurrentDomain.BaseDirectory
            }

            Dim candidates As New List(Of String)()
            For Each root In roots.Where(Function(value) Not String.IsNullOrWhiteSpace(value)).Distinct()
                Dim current = root
                While Not String.IsNullOrWhiteSpace(current)
                    AddCandidate(candidates, Path.Combine(current, "TaskPackages"))
                    AddCandidate(candidates, Path.Combine(current, "ShareRibbon", "TaskPackages"))
                    current = Path.GetDirectoryName(current)
                End While
            Next

            For Each candidate In candidates
                If Directory.Exists(candidate) Then LoadDirectoryInternal(candidate)
            Next

            If _packages.Count = 0 Then
                AddError("任务包目录未加载到任何合同",
                         "已尝试候选目录：" & String.Join(" | ", candidates.Where(Function(p) Directory.Exists(p)).DefaultIfEmpty("(无)")))
            End If

            AppLogger.Info("TaskPackageCatalog",
                           $"loaded={_packages.Count}, rejected={_validationErrors.Count}, dirs={candidates.Count}")
            If _validationErrors.Count > 0 Then
                ' 被拒绝的包等于这个任务在面板上没有归组、在验证器上没有断言；不写清楚就是一次静默降级。
                AppLogger.Warn("TaskPackageCatalog",
                               $"rejected={_validationErrors.Count}, first={_validationErrors(0)}")
            End If
            _loaded = True
            Return _packages.Count
        End Function

        ''' <summary>
        ''' 面板和冒烟脚本只问"这个宿主观得上面的任务包"，不应该各自记得先调一次加载；
        ''' 没加载过就现场加载一次，避免"清单在磁盘上但运行时视图是空的"这种看不见的空表。
        ''' </summary>
        Private Shared Sub EnsureLoaded()
            If _loaded Then Return
            LoadFromRuntimeDirectories()
        End Sub

        ''' <summary>重新加载前清空，保证重复调用不会把同一个包登记两次。</summary>
        Public Shared Sub Reset()
            _packages.Clear()
            _validationErrors.Clear()
            _sourceTexts.Clear()
            _loaded = False
        End Sub

        

        ''' <summary>候选目录去重：与 ToolRegistry 同一算法，同一物理目录只加载一次。</summary>
        Private Shared Sub AddCandidate(candidates As List(Of String), rawPath As String)
            If candidates Is Nothing OrElse String.IsNullOrWhiteSpace(rawPath) Then Return
            Dim fullPath = rawPath
            Try
                fullPath = Path.GetFullPath(rawPath)
            Catch
            End Try
            If candidates.Any(Function(existing) String.Equals(existing, fullPath, StringComparison.OrdinalIgnoreCase)) Then Return
            candidates.Add(fullPath)
        End Sub

        Public Shared Sub LoadDirectory(dir As String)
            If Not Directory.Exists(dir) Then Return
            LoadDirectoryInternal(dir)
        End Sub

        Private Shared Sub LoadDirectoryInternal(dir As String)
            For Each filePath In Directory.GetFiles(dir, "*.json", SearchOption.AllDirectories)
                LoadPackageFile(filePath)
            Next
        End Sub

        Private Shared Sub LoadPackageFile(jsonPath As String)
            Dim relative = jsonPath
            Try
                relative = Path.GetFileName(jsonPath)
            Catch
            End Try

            Dim node As JObject = Nothing
            Try
                node = CType(JToken.Parse(IO.File.ReadAllText(jsonPath)), JObject)
            Catch ex As Exception
                AddError(relative, $"JSON 解析失败：{AppLogger.Redact(ex.Message)}")
                Return
            End Try

            Dim errors As New List(Of String)()
            Dim pkg = ParsePackage(node, relative, errors)
            If errors.Count > 0 OrElse pkg Is Nothing Then
                For Each item In errors
                    AddError(relative, item)
                Next
                AppLogger.Warn("TaskPackageCatalog", $"rejected file={relative}, errors={errors.Count}")
                Return
            End If

            Dim parentDir = Path.GetDirectoryName(jsonPath)
            Dim canonical = node.ToString(Newtonsoft.Json.Formatting.None)
            Dim sameId = _packages.FirstOrDefault(Function(existing) String.Equals(existing.Id, pkg.Id, StringComparison.OrdinalIgnoreCase))
            If sameId IsNot Nothing Then
                ' 候选目录算法会同时命中源码目录和构建输出目录，同一份合同被读两次是预期的；
                ' 只有同一目录里出现两个同 id 才算真冲突，而那种情况已经由"文件名必须等于包 id"拦住。
                If String.Equals(sameId.SourceDirectory, parentDir, StringComparison.OrdinalIgnoreCase) Then
                    AddError(relative, $"任务包 id 在同一目录内重复：{pkg.Id}")
                Else
                    ' 两份内容不一致时取先加载的那一份，等于让一个没人读过的旧合同冒充当前合同：
                    ' 断言改了、黄金样例没改，收尾仍然按旧口径报成功。这类漂移必须报错，不能静默忽略。
                    Dim previous As String = Nothing
                    If _sourceTexts.TryGetValue(pkg.Id, previous) AndAlso
                       Not String.Equals(previous, canonical, StringComparison.Ordinal) Then
                        AddError(relative, $"同一任务包 id 在不同目录里内容不一致：{pkg.Id}")
                    End If
                End If
                Return
            End If
            pkg.SourceDirectory = parentDir
            _sourceTexts(pkg.Id) = canonical
            _packages.Add(pkg)
        End Sub

        Public Shared Function Find(id As String) As TaskPackage
            EnsureLoaded()
            Return _packages.FirstOrDefault(Function(item) String.Equals(item.Id, id, StringComparison.OrdinalIgnoreCase))
        End Function

        ''' <summary>该宿主的任务包；没有合同的宿主返回空表，调用方据此回落到扁平动作列表而不是猜归属。</summary>
        Public Shared Function PackagesFor(appType As String) As IReadOnlyList(Of TaskPackage)
            EnsureLoaded()
            Return _packages.
                Where(Function(item) String.Equals(item.Application, If(appType, "").Trim(), StringComparison.OrdinalIgnoreCase)).
               ToList().
                AsReadOnly()
        End Function

        ''' <summary>
        ''' 某个入口动作属于哪个任务包；未归组返回 Nothing，面板据此把它列在"其他"里。
        ''' </summary>
        Public Shared Function FindByAction(actionId As String) As TaskPackage
            If String.IsNullOrWhiteSpace(actionId) Then Return Nothing
            EnsureLoaded()
            Return _packages.FirstOrDefault(Function(item) item.EntryActions.Any(
                Function(existing) String.Equals(existing, actionId, StringComparison.OrdinalIgnoreCase)))
        End Function

        Private Shared Sub AddError(sourceFile As String, message As String)
            _validationErrors.Add($"{sourceFile}: {message}")
        End Sub
    End Class

End Namespace
