Imports System.IO
Imports System.Threading.Tasks
Imports Newtonsoft.Json

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
        Public Property ExecuteCode As Action(Of String, String, Boolean)

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
        Public Event OnRequestApproval(message As String, callback As Action(Of Boolean))
        Public Event OnPlanGenerated(plan As ExecutionPlan)
        Public Event OnCompleted(result As AgentResult)

        Public Sub New()
            Dim baseDir = Path.GetDirectoryName(GetType(AgentKernel).Assembly.Location)
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
            If Directory.Exists(ToolsDirectory) Then
                _toolRegistry.LoadFromDirectory(ToolsDirectory)
            End If

            ' 加载技能定义
            If Directory.Exists(SkillsDirectory) Then
                _skillRegistry.LoadFromDirectory(SkillsDirectory)
            End If

            ' 绑定回调
            _loopEngine.SendAIRequest = Function(prompt, system, history)
                                            Return SendAIRequest(prompt, system, history)
                                        End Function

            _loopEngine.OnPlanGenerated = Sub(plan)
                                              RaiseEvent OnPlanGenerated(plan)
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
        ''' 执行 Agent 任务 - 统一入口
        ''' </summary>
        Public Async Function ExecuteAsync(userRequest As String,
                                            appType As String,
                                            currentContent As String) As Task(Of AgentResult)
            ' 创建会话
            _session = New AgentSession(userRequest, appType, currentContent)
            _memory.ClearWorking()
            _memory.AddSessionMessage("user", userRequest)

            ' 绑定执行回调
            _toolRegistry.ExecuteCode = ExecuteCode

            ' 构建系统提示词
            Dim systemPrompt = _promptManager.BuildSystemPrompt(
                appType,
                _toolRegistry.GetAvailableTools(appType),
                _memory
            )

            ' 匹配技能
            Dim matchedSkill = _skillRegistry.MatchSkill(userRequest)
            If matchedSkill IsNot Nothing Then
                _session.Skill = matchedSkill
            End If

            ' 执行 ReAct Loop
            Dim result = Await _loopEngine.RunAsync(_session, systemPrompt, matchedSkill)

            ' 保存记忆
            _memory.AddTaskRecord(result)
            _memory.AddSessionMessage("assistant", result.Message)

            ' 通知完成
            RaiseEvent OnCompleted(result)

            Return result
        End Function

        ''' <summary>
        ''' 获取当前会话
        ''' </summary>
        Public Function GetCurrentSession() As AgentSession
            Return _session
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

        ''' <summary>
        ''' 重新加载提示词（热加载）
        ''' </summary>
        Public Sub ReloadPrompts()
            _promptManager.Reload()
        End Sub

        ''' <summary>
        ''' 重新加载工具
        ''' </summary>
        Public Sub ReloadTools()
            _toolRegistry.Clear()
            If Directory.Exists(ToolsDirectory) Then
                _toolRegistry.LoadFromDirectory(ToolsDirectory)
            End If
        End Sub

    End Class

End Namespace
