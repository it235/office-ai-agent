Imports System.IO
Imports System.Linq
Imports System.Threading
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Agent

    ''' <summary>
    ''' 统一 Agent Kernel - 替代 RalphLoopController + RalphAgentController
    ''' 整合 Prompt/Memory/Skills/Tools/Loop 六大维度
    ''' </summary>
    Public Class AgentKernel
        Private ReadOnly _promptManager As PromptManager
        Private ReadOnly _toolRegistry As ToolRegistry
        Private ReadOnly _skillRegistry As SkillRegistry
        Private ReadOnly _memory As AgentMemory
        Private ReadOnly _loopEngine As LoopEngine

        ' 当前会话
        Private _session As AgentSession

        ' 配置
        Public Property PromptsDirectory As String
        Public Property ToolsDirectory As String
        Public Property SkillsDirectory As String

        ' 外部回调（由 BaseChatControl 设置）
        Public Property SendAIRequest As Func(Of String, String, List(Of HistoryMessage), Task(Of String))
        Public Property SendAIRequestWithMessages As Func(Of JArray, Task(Of String))
        Public Property ExecuteCodeWithToolResult As Func(Of String, String, Boolean, ToolResult)

        ' MCP 客户端（由外部设置，可选）
        Public Property McpClient As StreamJsonRpcMCPClient
            Get
                Return _toolRegistry.McpClient
            End Get
            Set(value As StreamJsonRpcMCPClient)
                _toolRegistry.McpClient = value
            End Set
        End Property

        ' 状态通知
        Public Event OnStatusChanged(status As String)
        Public Event OnIterationUpdate(iteration As ReActIteration)
        Public Event OnStepCompleted(stepIndex As Integer, success As Boolean, message As String)
        Public Event OnExecutionExplained(explanation As ExecutionExplanation)
        Public Event OnRequestApproval(message As String, callback As Action(Of Boolean))
        Public Event OnPlanGenerated(plan As ExecutionPlan)
        Public Event OnCompleted(result As AgentResult)

        Public Sub New()
            Dim baseDir = ResolveRuntimeBaseDirectory()
            PromptsDirectory = Path.Combine(baseDir, "Prompts")
            ToolsDirectory = Path.Combine(baseDir, "Tools")
            SkillsDirectory = Path.Combine(baseDir, "Skills")

            ' 初始化组件
            _promptManager = New PromptManager(PromptsDirectory)
            _toolRegistry = New ToolRegistry()
            _skillRegistry = New SkillRegistry()
            _memory = New AgentMemory()
            _loopEngine = New LoopEngine(_toolRegistry, _memory, _promptManager)
        End Sub

        ''' <summary>
        ''' 初始化 - 加载所有配置（同步部分）
        ''' </summary>
        Public Sub Initialize()
            ' 加载工具定义
            LoadToolDefinitions()

            ' 加载技能定义
            If Directory.Exists(SkillsDirectory) Then
                _skillRegistry.LoadFromDirectory(SkillsDirectory)
            End If

            ' 将 Skill 脚本注册为可执行工具
            _toolRegistry.LoadSkillScriptsAsTools()

            ' 绑定回调
            _loopEngine.SendAIRequest = Function(prompt, system, history)
                                            Return SendAIRequest(prompt, system, history)
                                        End Function
            _loopEngine.SendAIRequestWithMessages = Function(messages)
                                                        If SendAIRequestWithMessages Is Nothing Then Return Task.FromResult(Of String)(Nothing)
                                                        Return SendAIRequestWithMessages(messages)
                                                    End Function

            _loopEngine.OnPlanGenerated = Sub(plan)
                                              RaiseEvent OnPlanGenerated(plan)
                                          End Sub

            _loopEngine.OnStatusChanged = Sub(status)
                                              RaiseEvent OnStatusChanged(status)
                                          End Sub

            _loopEngine.OnIterationUpdate = Sub(iteration)
                                                RaiseEvent OnIterationUpdate(iteration)
                                            End Sub

            _loopEngine.OnStepCompleted = Sub(stepIndex, success, message)
                                              RaiseEvent OnStepCompleted(stepIndex, success, message)
                                          End Sub

            _loopEngine.OnExecutionExplained = Sub(explanation)
                                                   RaiseEvent OnExecutionExplained(explanation)
                                               End Sub

            _loopEngine.OnRequestApproval = Async Function(msg)
                                                Dim tcs As New TaskCompletionSource(Of Boolean)()
                                                RaiseEvent OnRequestApproval(msg, Sub(approved) tcs.TrySetResult(approved))
                                                Return Await tcs.Task
                                            End Function

            _memory.SendAIRequest = Function(prompt, system, history)
                                        Return SendAIRequest(prompt, system, history)
                                    End Function
        End Sub

        Private Sub LoadToolDefinitions()
            Dim added = _toolRegistry.LoadFromRuntimeDirectories(ToolsDirectory)

            If added = 0 AndAlso _toolRegistry.ToolCount <= 4 Then
                AppLogger.Warn("AgentKernel", $"Tools directory not found. Primary={ToolsDirectory}")
            Else
                AppLogger.Info("AgentKernel", $"Tool registry initialized. Count={_toolRegistry.ToolCount}")
            End If
        End Sub

        Private Shared Function ResolveRuntimeBaseDirectory() As String
            For Each candidate In GetRuntimeBaseDirectoryCandidates()
                If Directory.Exists(Path.Combine(candidate, "Tools")) OrElse
                   Directory.Exists(Path.Combine(candidate, "Prompts")) OrElse
                   Directory.Exists(Path.Combine(candidate, "Skills")) Then
                    Return candidate
                End If
            Next

            Dim assemblyLocation = GetType(AgentKernel).Assembly.Location
            If Not String.IsNullOrWhiteSpace(assemblyLocation) Then
                Return Path.GetDirectoryName(assemblyLocation)
            End If

            Return AppDomain.CurrentDomain.BaseDirectory
        End Function

        Private Shared Function GetRuntimeBaseDirectoryCandidates() As List(Of String)
            Dim candidates As New List(Of String)()
            AddCandidate(candidates, GetDirectoryFromPath(GetType(AgentKernel).Assembly.Location))

            Try
                Dim codeBase = GetType(AgentKernel).Assembly.CodeBase
                If Not String.IsNullOrWhiteSpace(codeBase) Then
                    AddCandidate(candidates, GetDirectoryFromPath(New Uri(codeBase).LocalPath))
                End If
            Catch
            End Try

            AddCandidate(candidates, AppDomain.CurrentDomain.BaseDirectory)
            AddCandidate(candidates, AppDomain.CurrentDomain.RelativeSearchPath)

            Return candidates.Where(Function(c) Not String.IsNullOrWhiteSpace(c)).ToList()
        End Function

        Private Shared Function GetDirectoryFromPath(rawPath As String) As String
            If String.IsNullOrWhiteSpace(rawPath) Then Return Nothing
            Try
                If Directory.Exists(rawPath) Then Return rawPath
                Return Path.GetDirectoryName(rawPath)
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Sub AddCandidate(candidates As List(Of String), path As String)
            If String.IsNullOrWhiteSpace(path) Then Return
            If candidates.Any(Function(p) String.Equals(p, path, StringComparison.OrdinalIgnoreCase)) Then Return
            candidates.Add(path)
        End Sub

        ''' <summary>
        ''' 异步初始化 MCP 工具（在 MCP 客户端就绪后调用）
        ''' </summary>
        Public Async Function LoadMcpToolsAsync() As Task
            Try
                Await _toolRegistry.LoadMcpToolsAsync()
            Catch ex As Exception
                Debug.WriteLine($"[AgentKernel] 加载 MCP 工具失败: {ex.Message}")
            End Try
        End Function

        ''' <summary>
        ''' 执行 Agent 任务 - 统一入口（增强版 - 支持上下文自动注入）
        ''' </summary>
        ''' <param name="userRequest">用户请求</param>
        ''' <param name="appType">应用类型（Excel/Word/PowerPoint）</param>
        ''' <param name="currentContent">当前文档内容</param>
        ''' <param name="officeContext">Office 上下文（可选，未传入时自动采集）</param>
        Public Async Function ExecuteAsync(userRequest As String,
                                            appType As String,
                                            currentContent As String,
                                            Optional officeContext As Context.OfficeContext = Nothing,
                                            Optional contextPack As Context.ContextPack = Nothing,
                                            Optional taskSpec As AgentTaskSpec = Nothing,
                                            Optional selectedSkills As List(Of SkillFileDefinition) = Nothing,
                                            Optional cancellationToken As CancellationToken = Nothing) As Task(Of AgentResult)
            Dim cid = AppLogger.BeginScope()
            AppLogger.Info("AgentKernel", $"ExecuteAsync start appType={appType} cid={cid}")

            Try
                ' 创建会话
                _session = New AgentSession(userRequest, appType, currentContent)
                _session.Spec = taskSpec
                _memory.ClearWorking()
                _memory.AddSessionMessage("user", userRequest)

                ' Production path must inject CaptureOfficeContext; empty context is last-resort only.
                If officeContext Is Nothing Then
                    AppLogger.Warn("AgentKernel", "ExecuteAsync called without OfficeContext; using empty context")
                    officeContext = New Context.OfficeContext With {.AppType = appType}
                End If

                ' 将上下文保存到记忆中（供多轮对话使用）
                _memory.SetWorking("lastOfficeContext", officeContext)
                If contextPack Is Nothing Then contextPack = Context.ContextPack.FromOfficeContext(officeContext, currentContent)
                _memory.SetWorking("lastContextPack", contextPack)

                ' 绑定执行回调。Agent 原生工具必须返回 ToolResult，禁止退回 Boolean 假成功。
                _toolRegistry.ExecuteCodeWithToolResult = ExecuteCodeWithToolResult

                ' 自动选择 Skill：优先使用 filesystem Skill 索引，旧 JSON SkillRegistry 作为兜底。
                Dim matchedSkill As AgentSkill = Nothing
                If selectedSkills IsNot Nothing AndAlso selectedSkills.Count > 0 Then
                    Dim selectedDetail = SkillsDirectoryService.LoadSkillDetail(selectedSkills(0))
                    If selectedDetail Is Nothing Then selectedDetail = selectedSkills(0)
                    _session.SelectedSkill = selectedDetail
                    matchedSkill = ConvertFileSkillToAgentSkill(selectedDetail)
                    Debug.WriteLine($"[AgentKernel] 复用 AI Native 已选择 Skill: {matchedSkill.Name}")
                End If
                If matchedSkill Is Nothing Then matchedSkill = SelectSkillForRequest(userRequest, appType)
                If matchedSkill IsNot Nothing Then _session.Skill = matchedSkill
                Dim executionContext As ToolExecutionContext = ToolExecutionContext.FromSession(_session, matchedSkill)

                ' 构建系统提示词（注入上下文 + 当前 Skill 可见工具）
                Dim contextText = contextPack.ToPromptText()
                Dim systemPrompt = _promptManager.BuildSystemPrompt(
                    appType,
                    _toolRegistry.GetVisibleTools(appType, executionContext),
                    _memory,
                    contextText
                )

                ' 执行 ReAct Loop
                Dim result = Await _loopEngine.RunAsync(_session, systemPrompt, matchedSkill, cancellationToken)

                ' 保存记忆
                _memory.AddTaskRecord(result)
                _memory.AddSessionMessage("assistant", result.Message)

                If result IsNot Nothing AndAlso Not result.Success Then
                    AppLogger.Warn("AgentKernel", $"ExecuteAsync finished with failure: {result.Message}")
                Else
                    AppLogger.Info("AgentKernel", "ExecuteAsync finished success")
                End If

                ' 通知完成
                RaiseEvent OnCompleted(result)

                Return result
            Catch ex As Exception
                AppLogger.Error("AgentKernel", "ExecuteAsync unhandled exception", ex)
                Dim classified = ExceptionClassifier.Classify(ex)
                Dim failed = AgentResult.Failed(If(_session?.Id, Guid.NewGuid().ToString()),
                                                $"执行异常: [{classified.ErrorCode}] {classified.UserMessage}")
                RaiseEvent OnCompleted(failed)
                Return failed
            Finally
                AppLogger.ClearScope()
            End Try
        End Function

        ''' <summary>
        ''' 添加历史消息到记忆（启动前预加载）
        ''' </summary>
        Public Sub AddHistoryMessage(role As String, content As String)
            _memory.AddSessionMessage(role, content)
        End Sub

        ''' <summary>
        ''' 获取工具数量
        ''' </summary>
        Public ReadOnly Property ToolCount As Integer
            Get
                Return _toolRegistry.ToolCount
            End Get
        End Property

        ''' <summary>
        ''' 获取技能数量
        ''' </summary>
        Public ReadOnly Property SkillCount As Integer
            Get
                Return _skillRegistry.SkillCount
            End Get
        End Property

        Private Function SelectSkillForRequest(userRequest As String, appType As String) As AgentSkill
            Try
                Dim selected = SkillsIndexService.SelectSkillDefinitions(userRequest, Nothing, appType, 1)
                If selected IsNot Nothing AndAlso selected.Count > 0 Then
                    Dim detail = SkillsDirectoryService.LoadSkillDetail(selected(0))
                    If detail Is Nothing Then detail = selected(0)
                    _session.SelectedSkill = detail
                    Dim skill = ConvertFileSkillToAgentSkill(detail)
                    Debug.WriteLine($"[AgentKernel] 通过 SkillsIndexService 自动选择 Skill: {skill.Name}")
                    Return skill
                End If
            Catch ex As Exception
                Debug.WriteLine($"[AgentKernel] filesystem Skill 自动选择失败: {ex.Message}")
            End Try

            Dim fallback = _skillRegistry.MatchSkill(userRequest)
            If fallback IsNot Nothing Then
                Debug.WriteLine($"[AgentKernel] 使用旧 SkillRegistry 兜底 Skill: {fallback.Name}")
            End If
            Return fallback
        End Function

        Private Function ConvertFileSkillToAgentSkill(skill As SkillFileDefinition) As AgentSkill
            Dim agentSkill As New AgentSkill With {
                .Id = $"filesystem.{skill.Name}",
                .Name = skill.Name,
                .Description = If(skill.Description, ""),
                .PromptTemplate = SkillsService.BuildSkillDetailMessage(skill),
                .MaxSteps = 8,
                .AutoApprove = True
            }

            If skill.Tags IsNot Nothing Then
                For Each tag In skill.Tags
                    If Not String.IsNullOrWhiteSpace(tag) Then agentSkill.TriggerPatterns.Add(tag)
                Next
            End If

            If skill.AllowedTools IsNot Nothing Then
                For Each toolId In skill.AllowedTools
                    If Not String.IsNullOrWhiteSpace(toolId) Then agentSkill.RequiredTools.Add(toolId.Trim())
                Next
            End If

            If skill.Scripts IsNot Nothing Then
                For Each script In skill.Scripts
                    agentSkill.RequiredTools.Add($"skill_script.{skill.Name}.{script.FileName}")
                Next
            End If

            Return agentSkill
        End Function

    End Class

End Namespace
