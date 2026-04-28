Imports System.IO
Imports System.Text
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Agent

    ''' <summary>
    ''' 提示词管理器 - 从 JSON 文件加载并分层组装提示词
    ''' </summary>
    Public Class PromptManager
        Private ReadOnly _promptDir As String
        Private ReadOnly _promptCache As New Dictionary(Of String, JObject)(StringComparer.OrdinalIgnoreCase)

        Public Sub New(promptDir As String)
            _promptDir = promptDir
            LoadAllPrompts()
        End Sub

        ''' <summary>
        ''' 加载所有提示词 JSON 文件
        ''' </summary>
        Private Sub LoadAllPrompts()
            If Not Directory.Exists(_promptDir) Then Return
            For Each file In Directory.GetFiles(_promptDir, "*.json")
                Try
                    Dim name = Path.GetFileNameWithoutExtension(file)
                    _promptCache(name) = JObject.Parse(System.IO.File.ReadAllText(file))
                Catch ex As Exception
                    Debug.WriteLine($"[PromptManager] 加载提示词失败 {file}: {ex.Message}")
                End Try
            Next
        End Sub

        ''' <summary>
        ''' 构建系统提示词（5 层架构）
        ''' Layer 1: System Base
        ''' Layer 2: App Context
        ''' Layer 3: Tool Schema
        ''' Layer 4: Memory Context
        ''' Layer 5: User Request
        ''' </summary>
        Public Function BuildSystemPrompt(appType As String,
                                          tools As List(Of ToolDescriptor),
                                          Optional memory As AgentMemory = Nothing) As String
            Dim sb As New StringBuilder()

            ' Layer 1: System Base
            Dim basePrompt = GetPrompt("system-base")
            If basePrompt IsNot Nothing Then
                sb.AppendLine(basePrompt("role")?.ToString())
                sb.AppendLine()
                Dim constraints = TryCast(basePrompt("constraints"), JArray)
                If constraints IsNot Nothing Then
                    sb.AppendLine("【通用约束】")
                    For Each c In constraints
                        sb.AppendLine($"- {c}")
                    Next
                End If
            End If

            ' Layer 2: App Context
            Dim appContext = GetPrompt($"{appType}-context")
            If appContext IsNot Nothing Then
                sb.AppendLine()
                sb.AppendLine(appContext("role")?.ToString())
                sb.AppendLine()
                Dim appConstraints = TryCast(appContext("constraints"), JArray)
                If appConstraints IsNot Nothing Then
                    sb.AppendLine("【应用约束】")
                    For Each c In appConstraints
                        sb.AppendLine($"- {c}")
                    Next
                End If
                Dim dynamicRanges = TryCast(appContext("dynamicRanges"), JArray)
                If dynamicRanges IsNot Nothing Then
                    sb.AppendLine($"【动态范围占位符】{String.Join(", ", dynamicRanges)}")
                End If
            End If

            ' Layer 3: Tool Schema
            sb.AppendLine()
            sb.AppendLine("【已注册工具】")
            For Each tool In tools.OrderBy(Function(t) t.Category).ThenBy(Function(t) t.Id)
                sb.AppendLine($"{tool.Id}: {tool.Name} - {tool.Description}")
                For Each p In tool.Parameters
                    Dim req = If(p.Required, "必需", "可选")
                    sb.AppendLine($"  - {p.Name} ({p.Type}, {req}): {p.Description}")
                Next
            Next

            ' Layer 4: Memory Context
            If memory IsNot Nothing Then
                Dim relevantMemories = memory.Search("", 5)
                If relevantMemories.Count > 0 Then
                    sb.AppendLine()
                    sb.AppendLine("【相关记忆】")
                    For Each m In relevantMemories.Take(5)
                        sb.AppendLine($"- {m}")
                    Next
                End If
            End If

            sb.AppendLine()
            sb.AppendLine("【输出格式】")
            sb.AppendLine("你必须以 JSON 对象返回结果，使用 ```json 代码块包裹。")
            sb.AppendLine("JSON 必须包含以下字段：")
            sb.AppendLine("- thought: 你的思考过程（中文）")
            sb.AppendLine("- action: { tool: 工具ID, params: { 参数... } }")

            Return sb.ToString()
        End Function

        ''' <summary>
        ''' 构建规划阶段提示词
        ''' </summary>
        Public Function BuildPlanningPrompt(session As AgentSession,
                                             systemPrompt As String,
                                             Optional skill As AgentSkill = Nothing) As String
            Dim sb As New StringBuilder()
            sb.AppendLine(systemPrompt)
            sb.AppendLine()
            sb.AppendLine("【用户请求】")
            sb.AppendLine(session.UserRequest)

            If Not String.IsNullOrWhiteSpace(session.CurrentContent) Then
                sb.AppendLine()
                sb.AppendLine("【当前文档内容摘要】")
                Dim content = session.CurrentContent
                If content.Length > 500 Then
                    content = content.Substring(0, 500) & "..."
                End If
                sb.AppendLine(content)
            End If

            If skill IsNot Nothing Then
                sb.AppendLine()
                sb.AppendLine($"【匹配技能】{skill.Name}: {skill.Description}")
            End If

            sb.AppendLine()
            sb.AppendLine("请分析用户需求，制定详细的执行计划。")
            sb.AppendLine("返回 JSON 格式：")
            sb.AppendLine("```json")
            sb.AppendLine("{")
            sb.AppendLine("  ""understanding"": ""对用户需求的理解""")
            sb.AppendLine("  ""steps"": [")
            sb.AppendLine("    {")
            sb.AppendLine("      ""step"": 1,")
            sb.AppendLine("      ""description"": ""步骤描述""")
            sb.AppendLine("      ""code"": ""{\\""command\\"":\\""工具ID\\"",\\""params\\"":{...}}""")
            sb.AppendLine("    }")
            sb.AppendLine("  ],")
            sb.AppendLine("  ""summary"": ""预期结果""")
            sb.AppendLine("}")
            sb.AppendLine("```")

            Return sb.ToString()
        End Function

        ''' <summary>
        ''' 构建 ReAct 步骤提示词
        ''' </summary>
        Public Function BuildReactPrompt(planStep As PlanStep,
                                          memory As AgentMemory,
                                          Optional previousObservation As String = "") As String
            Dim sb As New StringBuilder()
            sb.AppendLine("请完成以下步骤：")
            sb.AppendLine($"步骤 {planStep.StepNumber}: {planStep.Description}")
            sb.AppendLine()

            If Not String.IsNullOrWhiteSpace(previousObservation) Then
                sb.AppendLine("【上一步的观察结果】")
                sb.AppendLine(previousObservation)
                sb.AppendLine()
            End If

            Dim lastObservation = memory.GetWorking("lastObservation")
            If lastObservation IsNot Nothing Then
                sb.AppendLine("【最新观察】")
                sb.AppendLine(lastObservation.ToString())
                sb.AppendLine()
            End If

            sb.AppendLine("请输出思考过程 + 工具调用：")
            sb.AppendLine("```json")
            sb.AppendLine("{")
            sb.AppendLine("  ""thought"": ""你的思考过程""")
            sb.AppendLine("  ""action"": { ""tool"": ""工具ID"", ""params"": { ... } }")
            sb.AppendLine("}")
            sb.AppendLine("```")

            Return sb.ToString()
        End Function

        ''' <summary>
        ''' 构建反思/修复提示词
        ''' </summary>
        Public Function BuildReflectionPrompt(session As AgentSession,
                                               failedObservation As String) As String
            Dim sb As New StringBuilder()
            sb.AppendLine("上一步执行失败，请分析原因并决定下一步行动。")
            sb.AppendLine()
            sb.AppendLine($"【失败原因】{failedObservation}")
            sb.AppendLine()

            If session.Iterations.Count > 0 Then
                sb.AppendLine("【执行历史】")
                Dim startIdx = Math.Max(0, session.Iterations.Count - 3)
                For i = startIdx To session.Iterations.Count - 1
                    Dim it = session.Iterations(i)
                    sb.AppendLine($"步骤 {it.Index}: {it.Action.ToolId} - {If(it.Observation, "成功", "失败")}")
                Next
            End If

            sb.AppendLine()
            sb.AppendLine("请返回决策（JSON）：")
            sb.AppendLine("```json")
            sb.AppendLine("{")
            sb.AppendLine("  ""analysis"": ""失败原因分析""")
            sb.AppendLine("  ""strategy"": ""retry|skip|replan""")
            sb.AppendLine("  ""reason"": ""选择该策略的理由""")
            sb.AppendLine("}")
            sb.AppendLine("```")

            Return sb.ToString()
        End Function

        ''' <summary>
        ''' 获取已加载的提示词
        ''' </summary>
        Private Function GetPrompt(name As String) As JObject
            If _promptCache.ContainsKey(name) Then
                Return _promptCache(name)
            End If
            Return Nothing
        End Function

        ''' <summary>
        ''' 重新加载所有提示词（支持热加载）
        ''' </summary>
        Public Sub Reload()
            _promptCache.Clear()
            LoadAllPrompts()
        End Sub
    End Class

End Namespace
