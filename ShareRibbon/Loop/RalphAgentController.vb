Imports System.Text
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

''' <summary>
''' Ralph Agent 控制器 - 类似Cursor的自动化Agent
''' 流程：获取内容 -> 规划 -> 用户确认 -> 自动逐步执行 -> 应用到Office
''' </summary>
Public Class RalphAgentController
    Private ReadOnly _memory As RalphLoopMemory
    
    ' Agent状态
    Private _currentSession As RalphAgentSession
    Private _isRunning As Boolean = False
    
    ' 回调委托
    Public Property OnStatusChanged As Action(Of String)
    Public Property OnStepStarted As Action(Of Integer, String)
    Public Property OnStepCompleted As Action(Of Integer, Boolean, String)
    Public Property OnAgentCompleted As Action(Of Boolean)
    Public Property OnRequestUserConfirm As Action(Of String, Action, Action) ' 消息、确认回调、取消回调
    
    ' AI请求委托（由BaseChatControl设置）
    Public Property SendAIRequest As Func(Of String, String, Task(Of String))
    
    ' 代码执行委托
    Public Property ExecuteCode As Action(Of String, String, Boolean)

    ' 规划提示词 - 一次性返回包含可执行代码的完整计划
    Private Const PLANNING_PROMPT As String = "你是一个Office自动化专家。用户有一个任务需要在Office中完成。

当前Office应用: {0}
当前选中/活动的内容:
```
{1}
```

用户需求: {2}

请分析这个需求，并制定一个详细的执行计划。每个步骤必须包含可直接执行的代码。

直接返回JSON对象，格式如下:
{{
  ""understanding"": ""对用户需求的理解"",
  ""steps"": [
    {{
      ""step"": 1,
      ""description"": ""步骤描述（用户可读）"",
      ""code"": ""可执行的JSON命令"",
      ""language"": ""json或vba""
    }}
  ],
  ""summary"": ""执行完成后的预期结果""
}}

【代码格式要求】
对于Word应用，code字段必须是JSON命令格式:
- 单命令: {{""command"":""InsertText"",""params"":{{""content"":""文本""}}}}
- 多命令: {{""commands"":[{{""command"":""InsertText"",""params"":{{""content"":""文本""}}}},...]}}
- Word支持的command: InsertText, FormatText, ReplaceText, InsertTable, ApplyStyle, GenerateTOC, BeautifyDocument

【重要】Word文本格式要求:
1. 段落之间必须使用 \\n\\n (两个换行符) 分隔
2. 标题后必须加 \\n\\n
3. 需要缩进时使用全角空格或多个半角空格
4. 示例: {{""command"":""InsertText"",""params"":{{""content"":""请假申请单\\n\\n    申请人：张三\\n    日期：2024年1月1日\\n\\n事由：...""}}}}

对于Excel应用，code字段必须使用JSON命令格式:
- 单命令: {{""command"":""ApplyFormula"",""params"":{{""targetRange"":""C1:C100"",""formula"":""=A1+B1""}}}}
- 多命令: {{""commands"":[{{""command"":""ApplyFormula"",""params"":{{...}}}},...]}}
- Excel只支持这些command: ApplyFormula, WriteData, FormatRange, CreateChart, CleanData
- ApplyFormula参数: targetRange(必需), formula(必需), fillDown(可选)
- WriteData参数: targetRange(必需), data(必需)
- FormatRange参数: range(必需), style(可选)
- CreateChart参数: dataRange(必需), chartType(可选)
- CleanData参数: range(必需), operation(必需)
- 动态范围使用{{lastRow}}占位符，不要用JS表达式

对于PowerPoint，code字段可以是VBA代码。

注意：
1. 每个步骤的code字段必须是可直接执行的完整代码
2. 不要对JSON引号进行转义
3. 只返回一个JSON对象，不要其他内容"

    ' 步骤执行提示词
    Private Const STEP_EXECUTION_PROMPT As String = "你是一个Office自动化专家。现在需要执行以下操作：

当前Office应用: {0}
当前文档内容:
```
{1}
```

要执行的步骤: {2}
步骤详情: {3}

请生成可以直接执行的代码。根据Office类型选择合适的代码格式：

【Excel】必须使用JSON命令格式，格式如下：
单个命令：
```json
{{
  ""command"": ""ApplyFormula"",
  ""params"": {{
    ""targetRange"": ""C1:C100"",
    ""formula"": ""=A1+B1""
  }}
}}
```
多个命令：
```json
{{
  ""commands"": [
    {{
      ""command"": ""ApplyFormula"",
      ""params"": {{ ""targetRange"": ""C1"", ""formula"": ""=A1+B1"" }}
    }},
    {{
      ""command"": ""WriteData"",
      ""params"": {{ ""targetRange"": ""D1"", ""data"": [[""标题""]] }}
    }}
  ]
}}
```
Excel只支持这些command: ApplyFormula, WriteData, FormatRange, CreateChart, CleanData
- ApplyFormula参数: targetRange(必需), formula(必需), fillDown(可选)
- WriteData参数: targetRange(必需), data(必需)
- FormatRange参数: range(必需), style(可选)
- CreateChart参数: dataRange(必需), chartType(可选)
- CleanData参数: range(必需), operation(必需)
动态范围使用{{lastRow}}占位符

【Word】必须使用JSON命令格式，格式如下：
单个命令：
```json
{{
  ""command"": ""InsertText"",
  ""params"": {{
    ""position"": ""cursor"",
    ""content"": ""要插入的文本""
  }}
}}
```
多个命令：
```json
{{
  ""commands"": [
    {{
      ""command"": ""InsertText"",
      ""params"": {{ ""content"": ""文本内容"" }}
    }},
    {{
      ""command"": ""FormatText"",
      ""params"": {{ ""range"": ""selection"", ""bold"": true }}
    }}
  ]
}}
```
Word支持的command类型：InsertText, FormatText, ReplaceText, InsertTable, ApplyStyle, GenerateTOC, BeautifyDocument

【PowerPoint】使用VBA代码

只返回可执行的代码，用```vba或```json包裹。"

    Public Sub New()
        _memory = RalphLoopMemory.Instance
    End Sub

    ''' <summary>
    ''' 启动Agent - 分析需求并规划
    ''' </summary>
    Public Async Function StartAgent(userRequest As String, appType As String, currentContent As String) As Task(Of Boolean)
        If _isRunning Then
            Return False
        End If

        _isRunning = True
        _currentSession = New RalphAgentSession() With {
            .UserRequest = userRequest,
            .ApplicationType = appType,
            .CurrentContent = currentContent,
            .Status = AgentStatus.Planning
        }

        OnStatusChanged?.Invoke("正在分析您的需求...")

        Try
            ' 调用AI进行规划
            Dim prompt = String.Format(PLANNING_PROMPT, appType, currentContent, userRequest)
            Dim response = Await SendAIRequest(prompt, "")

            ' 解析规划结果
            If ParsePlanningResult(response) Then
                _currentSession.Status = AgentStatus.WaitingConfirm
                OnStatusChanged?.Invoke($"规划完成，共 {_currentSession.Steps.Count} 个步骤")
                Return True
            Else
                _currentSession.Status = AgentStatus.Failed
                OnStatusChanged?.Invoke("规划失败，请重试")
                _isRunning = False
                Return False
            End If
        Catch ex As Exception
            _currentSession.Status = AgentStatus.Failed
            OnStatusChanged?.Invoke($"规划出错: {ex.Message}")
            _isRunning = False
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 解析规划结果
    ''' </summary>
    Private Function ParsePlanningResult(response As String) As Boolean
        Try
            ' 提取JSON
            Dim jsonStart = response.IndexOf("{")
            Dim jsonEnd = response.LastIndexOf("}")
            If jsonStart < 0 OrElse jsonEnd <= jsonStart Then Return False

            Dim jsonStr = response.Substring(jsonStart, jsonEnd - jsonStart + 1)

            ' 处理可能的转义符问题（LLM有时会返回带转义的JSON字符串）
            If jsonStr.Contains("\""") AndAlso Not jsonStr.Contains("\\""") Then
                jsonStr = jsonStr.Replace("\""", """")
                Debug.WriteLine("[RalphAgent] 检测到转义JSON，已还原")
            End If

            Dim planObj = JObject.Parse(jsonStr)

            _currentSession.Understanding = planObj("understanding")?.ToString()
            _currentSession.Summary = planObj("summary")?.ToString()

            Dim stepsArray = planObj("steps")
            If stepsArray Is Nothing Then Return False

            _currentSession.Steps.Clear()
            For Each stepItem In stepsArray
                Dim stepObj As New RalphAgentStep() With {
                    .StepNumber = CInt(stepItem("step")),
                    .Description = stepItem("description")?.ToString(),
                    .Status = StepStatus.Pending
                }

                ' 解析代码 - 新格式直接包含code字段
                Dim codeValue = stepItem("code")
                If codeValue IsNot Nothing Then
                    If codeValue.Type = JTokenType.Object Then
                        ' code是JSON对象，转为字符串
                        stepObj.GeneratedCode = codeValue.ToString(Newtonsoft.Json.Formatting.None)
                    Else
                        ' code是字符串
                        stepObj.GeneratedCode = codeValue.ToString()
                    End If
                    stepObj.CodeLanguage = stepItem("language")?.ToString()
                    If String.IsNullOrEmpty(stepObj.CodeLanguage) Then
                        ' 自动检测语言
                        stepObj.CodeLanguage = If(stepObj.GeneratedCode.TrimStart().StartsWith("{"), "json", "vba")
                    End If
                End If

                ' 兼容旧格式
                If String.IsNullOrEmpty(stepObj.GeneratedCode) Then
                    stepObj.ActionType = stepItem("action_type")?.ToString()
                    stepObj.Detail = stepItem("detail")?.ToString()
                End If

                _currentSession.Steps.Add(stepObj)
            Next

            Return _currentSession.Steps.Count > 0
        Catch ex As Exception
            Debug.WriteLine($"[RalphAgent] 解析规划失败: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 用户确认后开始执行
    ''' </summary>
    Public Async Function StartExecution() As Task
        If _currentSession Is Nothing OrElse _currentSession.Status <> AgentStatus.WaitingConfirm Then
            Return
        End If

        _currentSession.Status = AgentStatus.Executing
        _currentSession.CurrentStepIndex = 0

        ' 开始执行第一步
        Await ExecuteNextStep()
    End Function

    ''' <summary>
    ''' 执行下一步
    ''' </summary>
    Public Async Function ExecuteNextStep() As Task
        If _currentSession Is Nothing OrElse _currentSession.Status <> AgentStatus.Executing Then
            Return
        End If

        Dim stepIndex = _currentSession.CurrentStepIndex
        If stepIndex >= _currentSession.Steps.Count Then
            ' 所有步骤完成
            CompleteAgent(True)
            Return
        End If

        Dim currentStep = _currentSession.Steps(stepIndex)
        currentStep.Status = StepStatus.Running

        OnStepStarted?.Invoke(stepIndex, currentStep.Description)
        OnStatusChanged?.Invoke($"正在执行步骤 {stepIndex + 1}: {currentStep.Description}")

        Try
            Dim code As String = currentStep.GeneratedCode
            Dim language As String = currentStep.CodeLanguage

            ' 如果规划时没有生成代码（兼容旧格式），则调用LLM生成
            If String.IsNullOrEmpty(code) Then
                Debug.WriteLine($"[RalphAgent] 步骤 {stepIndex + 1} 没有预生成代码，调用LLM生成")
                Dim prompt = String.Format(STEP_EXECUTION_PROMPT,
                    _currentSession.ApplicationType,
                    _currentSession.CurrentContent,
                    currentStep.Description,
                    currentStep.Detail)

                Dim codeResponse = Await SendAIRequest(prompt, "")
                code = ExtractCode(codeResponse)
                language = DetectLanguage(codeResponse)
                currentStep.GeneratedCode = code
                currentStep.CodeLanguage = language
            End If

            If Not String.IsNullOrEmpty(code) Then
                Debug.WriteLine($"[RalphAgent] 执行步骤 {stepIndex + 1} 代码: {code.Substring(0, Math.Min(100, code.Length))}...")

                ' 执行代码 - 不使用预览模式
                ExecuteCode?.Invoke(code, language, False)

                currentStep.Status = StepStatus.Completed
                OnStepCompleted?.Invoke(stepIndex, True, "执行成功")
            Else
                currentStep.Status = StepStatus.Failed
                OnStepCompleted?.Invoke(stepIndex, False, "无法生成执行代码")
            End If
            
            ' 准备下一步
            _currentSession.CurrentStepIndex += 1
            
            ' 自动执行下一步（类似Cursor）
            If _currentSession.CurrentStepIndex < _currentSession.Steps.Count Then
                ' 短暂延迟后执行下一步，让用户看到进度
                Await Task.Delay(500)
                Await ExecuteNextStep()
            Else
                CompleteAgent(True)
            End If
            
        Catch ex As Exception
            currentStep.Status = StepStatus.Failed
            currentStep.ErrorMessage = ex.Message
            OnStepCompleted?.Invoke(stepIndex, False, ex.Message)
            
            ' 询问是否继续
            OnStatusChanged?.Invoke($"步骤 {stepIndex + 1} 执行失败: {ex.Message}")
        End Try
    End Function

    ''' <summary>
    ''' 提取代码块
    ''' </summary>
    Private Function ExtractCode(response As String) As String
        ' 查找```包裹的代码块
        Dim codeStart = response.IndexOf("```")
        If codeStart < 0 Then Return response.Trim()
        
        ' 跳过```和语言标识
        Dim lineEnd = response.IndexOf(vbLf, codeStart)
        If lineEnd < 0 Then lineEnd = response.IndexOf(vbCr, codeStart)
        If lineEnd < 0 Then Return ""
        
        Dim codeEnd = response.IndexOf("```", lineEnd)
        If codeEnd < 0 Then Return response.Substring(lineEnd + 1).Trim()
        
        Return response.Substring(lineEnd + 1, codeEnd - lineEnd - 1).Trim()
    End Function

    ''' <summary>
    ''' 检测代码语言
    ''' </summary>
    Private Function DetectLanguage(response As String) As String
        If response.Contains("```vba") OrElse response.Contains("```vbscript") Then
            Return "vba"
        ElseIf response.Contains("```json") Then
            Return "json"
        ElseIf response.Contains("```javascript") OrElse response.Contains("```js") Then
            Return "javascript"
        Else
            Return "vba" ' 默认VBA
        End If
    End Function

    ''' <summary>
    ''' 完成Agent
    ''' </summary>
    Private Sub CompleteAgent(success As Boolean)
        _currentSession.Status = If(success, AgentStatus.Completed, AgentStatus.Failed)
        _isRunning = False
        
        ' 保存到记忆
        _memory.AddTaskRecord(New RalphTaskRecord() With {
            .UserInput = _currentSession.UserRequest,
            .Intent = "agent_task",
            .Plan = String.Join(" -> ", _currentSession.Steps.Select(Function(s) s.Description)),
            .Result = If(success, "成功完成", "执行失败"),
            .Success = success,
            .ApplicationType = _currentSession.ApplicationType
        })
        _memory.Save()
        
        OnAgentCompleted?.Invoke(success)
        OnStatusChanged?.Invoke(If(success, "所有步骤执行完成！", "执行未完成"))
    End Sub

    ''' <summary>
    ''' 终止Agent
    ''' </summary>
    Public Sub AbortAgent()
        If _currentSession IsNot Nothing Then
            _currentSession.Status = AgentStatus.Aborted
        End If
        _isRunning = False
        OnStatusChanged?.Invoke("已终止")
        OnAgentCompleted?.Invoke(False)
    End Sub

    ''' <summary>
    ''' 获取当前会话
    ''' </summary>
    Public Function GetCurrentSession() As RalphAgentSession
        Return _currentSession
    End Function

    ''' <summary>
    ''' 是否正在运行
    ''' </summary>
    Public Function IsRunning() As Boolean
        Return _isRunning
    End Function
End Class

''' <summary>
''' Agent会话
''' </summary>
Public Class RalphAgentSession
    Public Property Id As String = Guid.NewGuid().ToString()
    Public Property UserRequest As String
    Public Property ApplicationType As String
    Public Property CurrentContent As String
    Public Property Understanding As String
    Public Property Summary As String
    Public Property Steps As New List(Of RalphAgentStep)
    Public Property CurrentStepIndex As Integer = 0
    Public Property Status As AgentStatus = AgentStatus.Idle
End Class

''' <summary>
''' Agent步骤
''' </summary>
Public Class RalphAgentStep
    Public Property StepNumber As Integer
    Public Property Description As String
    Public Property ActionType As String
    Public Property Detail As String
    Public Property GeneratedCode As String
    Public Property CodeLanguage As String
    Public Property Status As StepStatus = StepStatus.Pending
    Public Property ErrorMessage As String
End Class

''' <summary>
''' Agent状态
''' </summary>
Public Enum AgentStatus
    Idle
    Planning
    WaitingConfirm
    Executing
    Completed
    Failed
    Aborted
End Enum

''' <summary>
''' 步骤状态
''' </summary>
Public Enum StepStatus
    Pending
    Running
    Completed
    Failed
    Skipped
End Enum
