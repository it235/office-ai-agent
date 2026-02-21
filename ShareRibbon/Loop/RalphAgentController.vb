Imports System.Text
Imports System.Text.RegularExpressions
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

''' <summary>
''' Ralph Agent 控制器 - 类似Cursor的自动化Agent
''' 流程：意图识别 -> 记忆RAG检索 -> 规划 -> 用户确认 -> 自动逐步执行 -> 应用到Office
''' </summary>
Public Class RalphAgentController
    Private ReadOnly _memory As RalphLoopMemory

    ' 意图识别结果
    Private _intentResult As IntentResult
    ' RAG检索到的相关记忆
    Private _ragMemories As List(Of AtomicMemoryRecord)

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

    ' Excel专用规划提示词
    Private Const PLANNING_PROMPT_EXCEL As String = "你是一个智能Excel自动化专家。用户有一个任务需要在Excel中完成。

当前选中/活动的内容:
```
{1}
```

【历史对话上下文】
{5}

用户当前需求: {2}

【意图识别结果】
{3}

【相关记忆（RAG检索）】
{4}

请深度分析这个需求，结合识别到的意图、相关记忆和历史对话，制定一个详细、可执行的计划。每个步骤必须包含可直接执行的代码。

直接返回JSON对象，格式如下:
{{
  ""understanding"": ""对用户需求的理解（结合意图和记忆）"",
  ""steps"": [
    {{
      ""step"": 1,
      ""description"": ""步骤描述（用户可读）"",
      ""code"": ""可执行的JSON命令"",
      ""language"": ""json""
    }}
  ],
  ""summary"": ""执行完成后的预期结果""
}}

【Excel代码格式要求】
code字段必须使用JSON命令格式:
- 单命令: {{""command"":""ApplyFormula"",""params"":{{""targetRange"":""C1:C100"",""formula"":""=A1+B1""}}}}
- 多命令: {{""commands"":[{{""command"":""ApplyFormula"",""params"":{{...}}}},...]}}
- Excel只支持这些command: ApplyFormula, WriteData, FormatRange, CreateChart, CleanData
- ApplyFormula参数: targetRange(必需), formula(必需), fillDown(可选)
- WriteData参数: targetRange(必需), data(必需)
- FormatRange参数: range(必需), style(可选)
- CreateChart参数: dataRange(必需), chartType(可选)
- CleanData参数: range(必需), operation(必需)
- 动态范围使用{{lastRow}}占位符

注意：
1. 每个步骤的code字段必须是可直接执行的完整代码
2. 不要对JSON引号进行转义
3. 只返回一个JSON对象，不要其他内容"

    ' Word专用规划提示词
    Private Const PLANNING_PROMPT_WORD As String = "你是一个智能Word自动化专家。用户有一个任务需要在Word中完成。

当前选中/活动的内容:
```
{1}
```

【历史对话上下文】
{5}

用户当前需求: {2}

【意图识别结果】
{3}

【相关记忆（RAG检索）】
{4}

请深度分析这个需求，结合识别到的意图、相关记忆和历史对话，制定一个详细、可执行的计划。每个步骤必须包含可直接执行的代码。

直接返回JSON对象，格式如下:
{{
  ""understanding"": ""对用户需求的理解（结合意图和记忆）"",
  ""steps"": [
    {{
      ""step"": 1,
      ""description"": ""步骤描述（用户可读）"",
      ""code"": ""可执行的JSON命令"",
      ""language"": ""json""
    }}
  ],
  ""summary"": ""执行完成后的预期结果""
}}

【Word代码格式要求】
code字段必须是JSON命令格式:
- 单命令: {{""command"":""InsertText"",""params"":{{""content"":""文本""}}}}
- 多命令: {{""commands"":[{{""command"":""InsertText"",""params"":{{""content"":""文本""}}}},...]}}
- Word支持的command: InsertText, FormatText, ReplaceText, InsertTable, ApplyStyle, GenerateTOC, BeautifyDocument

【重要】Word文本格式要求:
1. 段落之间必须使用 \n\n (两个换行符) 分隔
2. 标题后必须加 \n\n
3. 需要缩进时使用全角空格或多个半角空格
4. 示例: {{""command"":""InsertText"",""params"":{{""content"":""请假申请单\n\n    申请人：张三\n    日期：2024年1月1日\n\n事由：...""}}}}

注意：
1. 每个步骤的code字段必须是可直接执行的完整代码
2. 不要对JSON引号进行转义
3. 只返回一个JSON对象，不要其他内容"

    ' PowerPoint专用规划提示词
    Private Const PLANNING_PROMPT_POWERPOINT As String = "你是一个智能PowerPoint自动化专家。用户有一个任务需要在PowerPoint中完成。

当前选中/活动的内容:
```
{1}
```

【历史对话上下文】
{5}

用户当前需求: {2}

【意图识别结果】
{3}

【相关记忆（RAG检索）】
{4}

请深度分析这个需求，结合识别到的意图、相关记忆和历史对话，制定一个详细、可执行的计划。每个步骤必须包含可直接执行的代码。

直接返回JSON对象，格式如下:
{{
  ""understanding"": ""对用户需求的理解（结合意图和记忆）"",
  ""steps"": [
    {{
      ""step"": 1,
      ""description"": ""步骤描述（用户可读）"",
      ""code"": ""可执行的VBA代码"",
      ""language"": ""vba""
    }}
  ],
  ""summary"": ""执行完成后的预期结果""
}}

【PowerPoint代码格式要求】
code字段可以是VBA代码。

注意：
1. 每个步骤的code字段必须是可直接执行的完整代码
2. 不要对JSON引号进行转义
3. 只返回一个JSON对象，不要其他内容"

    ' 获取对应应用的规划提示词
    Private Function GetPlanningPrompt(appType As String) As String
        Select Case appType.ToLower()
            Case "excel"
                Return PLANNING_PROMPT_EXCEL
            Case "word"
                Return PLANNING_PROMPT_WORD
            Case "powerpoint", "ppt"
                Return PLANNING_PROMPT_POWERPOINT
            Case Else
                Return PLANNING_PROMPT_EXCEL
        End Select
    End Function

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
    ''' 启动Agent - 意图识别 -> RAG检索 -> 规划
    ''' </summary>
    Public Async Function StartAgent(userRequest As String, appType As String, currentContent As String, historyMessages As List(Of Tuple(Of String, String))) As Task(Of Boolean)
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

        Try
            ' 步骤1: 意图识别（使用LLM进行更智能的识别）
            OnStatusChanged?.Invoke("正在识别您的意图...")
            _intentResult = Await RecognizeIntentAsync(userRequest, appType, currentContent, historyMessages)
            Debug.WriteLine($"[RalphAgent] 识别到意图: {_intentResult.OfficeIntent}，置信度: {_intentResult.Confidence}")

            ' 步骤2: 记忆RAG检索
            OnStatusChanged?.Invoke("正在检索相关记忆...")
            _ragMemories = MemoryService.GetRelevantMemories(userRequest, 5)
            Debug.WriteLine($"[RalphAgent] 检索到 {_ragMemories.Count} 条相关记忆")

            ' 步骤3: 格式化历史对话
            Dim historyInfo = FormatHistoryMessages(historyMessages)
            Dim historyCount = If(historyMessages IsNot Nothing, historyMessages.Count, 0)
            Debug.WriteLine($"[RalphAgent] 历史对话包含 {historyCount} 条消息")
            Debug.WriteLine(historyMessages)

            ' 步骤4: 制定执行计划
            OnStatusChanged?.Invoke("正在制定执行计划...")
            Dim intentInfo = FormatIntentInfo(_intentResult)
            Dim ragInfo = FormatRagMemories(_ragMemories)

            ' 调用AI进行规划
            Dim planningPrompt = GetPlanningPrompt(appType)
            Dim prompt = String.Format(planningPrompt, appType, currentContent, userRequest, intentInfo, ragInfo, historyInfo)
            Debug.WriteLine($"[RalphAgent] 规划提示词:\n{prompt}")
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
    ''' 智能修复 JSON 格式
    ''' </summary>
    Private Function FixJsonFormat(jsonStr As String) As String
        Try
            Dim fixedJson As String = jsonStr

            ' 1. 修复未加引号的属性名（在 { 或 , 后面）
            fixedJson = Regex.Replace(fixedJson, "([{,])\s*(\w+)\s*:", "$1""$2"":")

            ' 2. 修复在 " 前面没有逗号的问题（在 } 前面）
            fixedJson = Regex.Replace(fixedJson, "([^\s,])\s*}", "$1}")

            ' 3. 修复转义引号问题
            fixedJson = fixedJson.Replace("\""", """")

            ' 4. 修复常见的格式问题
            ' 修复多行字符串中缺少的引号
            fixedJson = FixMissingQuotes(fixedJson)

            ' 5. 尝试平衡大括号
            fixedJson = BalanceBraces(fixedJson)

            Return fixedJson
        Catch ex As Exception
            Debug.WriteLine($"FixJsonFormat 出错: {ex.Message}")
            Return jsonStr
        End Try
    End Function

    ''' <summary>
    ''' 修复缺失的引号
    ''' </summary>
    Private Function FixMissingQuotes(jsonStr As String) As String
        Dim result As New StringBuilder()
        Dim inString As Boolean = False
        Dim escapeNext As Boolean = False
        Dim i As Integer = 0

        While i < jsonStr.Length
            Dim c As Char = jsonStr(i)

            If escapeNext Then
                result.Append(c)
                escapeNext = False
            ElseIf c = "\""" Then
                inString = Not inString
                result.Append(c)
            ElseIf c = "\" AndAlso inString Then
                result.Append(c)
                escapeNext = True
            ElseIf (c = "," OrElse c = "}") AndAlso Not inString Then
                ' 检查前面是否有未闭合的字符串
                Dim lastQuoteIndex As Integer = result.ToString().LastIndexOf(""""c)
                If lastQuoteIndex >= 0 Then
                    Dim betweenQuotesAndCurrent As String = result.ToString().Substring(lastQuoteIndex + 1)
                    If Not betweenQuotesAndCurrent.Contains(""""c) AndAlso betweenQuotesAndCurrent.Contains(":") Then
                        ' 可能缺少闭合引号，尝试添加
                        result.Append(""""c)
                    End If
                End If
                result.Append(c)
            Else
                result.Append(c)
            End If

            i += 1
        End While

        Return result.ToString()
    End Function

    ''' <summary>
    ''' 平衡大括号
    ''' </summary>
    Private Function BalanceBraces(jsonStr As String) As String
        Dim openBraces As Integer = 0
        Dim closeBraces As Integer = 0

        For Each c As Char In jsonStr
            If c = "{" Then openBraces += 1
            If c = "}" Then closeBraces += 1
        Next

        Dim result As String = jsonStr.Trim()

        ' 如果缺少闭合大括号，添加
        While openBraces > closeBraces
            result &= "}"
            closeBraces += 1
        End While

        Return result
    End Function

    ''' <summary>
    ''' 解析规划结果
    ''' </summary>
    Private Function ParsePlanningResult(response As String) As Boolean
        Try
            ' 提取第一个完整的JSON对象
            Dim jsonStr As String = ExtractFirstCompleteJson(response)
            If String.IsNullOrEmpty(jsonStr) Then
                Debug.WriteLine("[RalphAgent] 无法从响应中提取完整JSON")
                Return False
            End If

            Debug.WriteLine($"[RalphAgent] 提取到JSON，长度: {jsonStr.Length}")

            ' 处理可能的转义符问题
            Dim planObj As JObject = Nothing
            Dim parseSuccess As Boolean = False
            Dim currentJson As String = jsonStr

            ' 首先尝试直接解析
            Try
                planObj = JObject.Parse(currentJson)
                parseSuccess = True
            Catch
                parseSuccess = False
            End Try

            ' 如果直接解析失败，尝试修复转义问题
            If Not parseSuccess Then
                Try
                    planObj = JsonConvert.DeserializeObject(Of JObject)(currentJson)
                    parseSuccess = True
                    Debug.WriteLine("[RalphAgent] 使用JsonConvert.DeserializeObject解析成功")
                Catch
                    parseSuccess = False
                End Try
            End If

            ' 如果还失败，尝试使用JToken.Parse
            If Not parseSuccess Then
                Try
                    Dim token = JToken.Parse(currentJson)
                    If token.Type = JTokenType.Object Then
                        planObj = DirectCast(token, JObject)
                        parseSuccess = True
                        Debug.WriteLine("[RalphAgent] 使用JToken.Parse解析成功")
                    End If
                Catch
                    parseSuccess = False
                End Try
            End If

            ' 如果所有方法都失败，尝试智能修复
            If Not parseSuccess Then
                Debug.WriteLine("[RalphAgent] 尝试智能修复JSON格式")
                currentJson = FixJsonFormat(jsonStr)
                Debug.WriteLine($"[RalphAgent] 格式修正提示已生成，长度: {currentJson.Length}")

                ' 再次尝试解析修复后的JSON
                Try
                    planObj = JObject.Parse(currentJson)
                    parseSuccess = True
                    Debug.WriteLine("[RalphAgent] 智能修复后解析成功")
                Catch
                    parseSuccess = False
                End Try

                If Not parseSuccess Then
                    Try
                        planObj = JsonConvert.DeserializeObject(Of JObject)(currentJson)
                        parseSuccess = True
                        Debug.WriteLine("[RalphAgent] 智能修复后使用JsonConvert.DeserializeObject解析成功")
                    Catch
                        parseSuccess = False
                    End Try
                End If
            End If

            If Not parseSuccess Then
                Debug.WriteLine("[RalphAgent] 所有解析方法都失败")
                Return False
            End If

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
                        ' code是字符串，尝试解析它（可能是转义的JSON）
                        Dim codeStr = codeValue.ToString()
                        Dim innerJson As JObject = Nothing
                        Dim parseCodeSuccess As Boolean = False
                        Dim currentCodeJson As String = codeStr

                        ' 尝试解析
                        Try
                            innerJson = JObject.Parse(currentCodeJson)
                            parseCodeSuccess = True
                        Catch
                        End Try

                        ' 尝试智能修复
                        If Not parseCodeSuccess Then
                            currentCodeJson = FixJsonFormat(codeStr)
                            Try
                                innerJson = JObject.Parse(currentCodeJson)
                                parseCodeSuccess = True
                                Debug.WriteLine("[RalphAgent] 步骤代码智能修复后解析成功")
                            Catch
                            End Try
                        End If

                        If parseCodeSuccess Then
                            stepObj.GeneratedCode = innerJson.ToString(Newtonsoft.Json.Formatting.None)
                        Else
                            ' 如果解析失败，直接使用原字符串
                            stepObj.GeneratedCode = codeStr
                        End If
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

            Debug.WriteLine($"[RalphAgent] 解析成功，共 {_currentSession.Steps.Count} 个步骤")
            Return _currentSession.Steps.Count > 0
        Catch ex As Exception
            Debug.WriteLine($"[RalphAgent] 解析规划失败: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 从响应中提取第一个完整的JSON对象（处理嵌套结构）
    ''' </summary>
    Private Function ExtractFirstCompleteJson(response As String) As String
        Dim jsonStart = response.IndexOf("{")
        If jsonStart < 0 Then Return Nothing

        Dim braceCount = 0
        Dim jsonEnd = -1

        For i As Integer = jsonStart To response.Length - 1
            If response(i) = "{"c Then
                braceCount += 1
            ElseIf response(i) = "}"c Then
                braceCount -= 1
                If braceCount = 0 Then
                    jsonEnd = i
                    Exit For
                End If
            End If
        Next

        If jsonEnd < 0 Then Return Nothing

        Return response.Substring(jsonStart, jsonEnd - jsonStart + 1)
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

    ''' <summary>
    ''' 识别用户意图（包含历史上下文）- 使用LLM进行更智能的识别
    ''' </summary>
    Private Async Function RecognizeIntentAsync(userRequest As String, appType As String, currentContent As String, historyMessages As List(Of Tuple(Of String, String))) As Task(Of IntentResult)
        ' 创建临时的IntentRecognitionService进行意图识别
        Dim intentAppType As OfficeApplicationType
        Select Case appType.ToLower()
            Case "excel"
                intentAppType = OfficeApplicationType.Excel
            Case "word"
                intentAppType = OfficeApplicationType.Word
            Case "powerpoint", "ppt"
                intentAppType = OfficeApplicationType.PowerPoint
            Case Else
                intentAppType = OfficeApplicationType.Excel
        End Select

        Dim intentService As New IntentRecognitionService(intentAppType)
        Dim context As New JObject()
        context("currentContent") = currentContent

        ' 添加历史对话到上下文
        If historyMessages IsNot Nothing AndAlso historyMessages.Count > 0 Then
            Dim historyContext As New StringBuilder()
            For Each msg In historyMessages
                historyContext.AppendLine($"{msg.Item1}: {msg.Item2}")
            Next
            context("history") = historyContext.ToString()
        End If

        ' 优先使用LLM进行意图识别（更智能）
        Dim result = Await intentService.IdentifyIntentAsync(userRequest, context)
        result.OriginalInput = userRequest

        Return result
    End Function

    ''' <summary>
    ''' 识别用户意图（包含历史上下文）- 基于关键词匹配（备用方法）
    ''' </summary>
    Private Function RecognizeIntent(userRequest As String, appType As String, currentContent As String, historyMessages As List(Of Tuple(Of String, String))) As IntentResult
        ' 创建临时的IntentRecognitionService进行意图识别
        Dim intentAppType As OfficeApplicationType
        Select Case appType.ToLower()
            Case "excel"
                intentAppType = OfficeApplicationType.Excel
            Case "word"
                intentAppType = OfficeApplicationType.Word
            Case "powerpoint", "ppt"
                intentAppType = OfficeApplicationType.PowerPoint
            Case Else
                intentAppType = OfficeApplicationType.Excel
        End Select

        Dim intentService As New IntentRecognitionService(intentAppType)
        Dim context As New JObject()
        context("currentContent") = currentContent

        ' 添加历史对话到上下文
        If historyMessages IsNot Nothing AndAlso historyMessages.Count > 0 Then
            Dim historyContext As New StringBuilder()
            For Each msg In historyMessages
                historyContext.AppendLine($"{msg.Item1}: {msg.Item2}")
            Next
            context("history") = historyContext.ToString()
        End If

        Dim result = intentService.IdentifyIntent(userRequest, context)
        result.OriginalInput = userRequest

        Return result
    End Function

    ''' <summary>
    ''' 格式化历史对话消息
    ''' </summary>
    Private Function FormatHistoryMessages(historyMessages As List(Of Tuple(Of String, String))) As String
        If historyMessages Is Nothing OrElse historyMessages.Count = 0 Then
            Return "无历史对话"
        End If

        Dim sb As New StringBuilder()
        ' 使用配置中的会话摘要条数限制
        Dim limit = MemoryConfig.SessionSummaryLimit
        Dim startIndex = Math.Max(0, historyMessages.Count - limit)
        For i = startIndex To historyMessages.Count - 1
            Dim msg = historyMessages(i)
            Dim role = If(msg.Item1 = "user", "用户", "AI助手")
            sb.AppendLine($"[{role}] {msg.Item2}")
        Next

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 格式化意图信息用于提示词
    ''' </summary>
    Private Function FormatIntentInfo(intentResult As IntentResult) As String
        If intentResult Is Nothing Then
            Return "无意图信息"
        End If

        Dim sb As New StringBuilder()
        sb.AppendLine($"识别到的主要意图: {intentResult.OfficeIntent}")
        sb.AppendLine($"置信度: {intentResult.Confidence:P0}")
        If Not String.IsNullOrEmpty(intentResult.ResponseMode) Then
            sb.AppendLine($"响应模式: {intentResult.ResponseMode}")
        End If
        If intentResult.SecondaryIntents IsNot Nothing AndAlso intentResult.SecondaryIntents.Count > 0 Then
            sb.AppendLine($"次要意图: {String.Join(", ", intentResult.SecondaryIntents)}")
        End If
        If Not String.IsNullOrEmpty(intentResult.UserFriendlyDescription) Then
            sb.AppendLine($"意图描述: {intentResult.UserFriendlyDescription}")
        End If

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 格式化RAG记忆用于提示词
    ''' </summary>
    Private Function FormatRagMemories(memories As List(Of AtomicMemoryRecord)) As String
        If memories Is Nothing OrElse memories.Count = 0 Then
            Return "无相关记忆"
        End If

        Dim sb As New StringBuilder()
        ' 使用配置中的RAG检索条数限制
        Dim limit = MemoryConfig.RagTopN
        For i = 0 To Math.Min(memories.Count - 1, limit - 1)
            Dim mem = memories(i)
            sb.AppendLine($"[记忆{i + 1}]")
            sb.AppendLine($"内容: {mem.Content}")
            If Not String.IsNullOrEmpty(mem.CreateTime) Then
                sb.AppendLine($"时间: {mem.CreateTime}")
            End If
            sb.AppendLine()
        Next

        Return sb.ToString()
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
