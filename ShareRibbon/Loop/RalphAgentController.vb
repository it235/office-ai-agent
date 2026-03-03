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

    ' AI请求委托（由BaseChatControl设置）- 第3参数为历史对话消息列表（可选）
    Public Property SendAIRequest As Func(Of String, String, List(Of HistoryMessage), Task(Of String))

    ' 代码执行委托
    Public Property ExecuteCode As Action(Of String, String, Boolean)

    ' === Excel 规划提示词 ===
    Private Const PLANNING_SYSTEM_EXCEL As String = "你是一个智能Excel自动化专家。请深度分析用户需求，结合识别到的意图、相关记忆和历史对话，制定一个详细、可执行的计划。每个步骤必须包含可直接执行的代码。

直接返回markdown的JSON代码块对象，格式如下:
```{
  ""understanding"": ""对用户需求的理解（结合意图和记忆）"",
  ""steps"": [
    {{
      ""step"": 1,
      ""description"": ""步骤描述（用户可读）"",
      ""code"": ""可执行的Excel JSON命令，不能返回\n\r等转义字符"",
      ""language"": ""json""
    }}
  ],
  ""summary"": ""执行完成后的预期结果""
}```

【Excel代码格式要求】
code字段必须使用JSON命令格式，不能包含\n、\r等换行转义字符，这不符合json规范:
- 单命令: {{""command"":""ApplyFormula"",""params"":{{""targetRange"":""C1:C100"",""formula"":""=A1+B1""}}}}
- 多命令: {{""commands"":[{{""command"":""ApplyFormula"",""params"":{{...}}}},...]}}

【Excel支持的22个命令】

=== 基础操作 (5个) ===
1. ApplyFormula - 应用公式 {targetRange:必需, formula:必需, fillDown:可选}
2. WriteData - 写入数据 {targetRange:必需, data:必需(单值或二维数组)}
3. FormatRange - 格式化 {range:必需, style:header/total/data, bold/italic/fontSize/backgroundColor/fontColor, borders:true/""all""/""outline""/""none""}
4. CreateChart - 创建图表 {dataRange:必需, type:column/line/pie/bar/scatter/area, title:可选, position:可选, seriesNames:系列名称数组, categoryAxis:分类轴范围, legendPosition:right/left/top/bottom}
5. CleanData - 数据清洗 {range:必需, operation:removeduplicates/fillempty/trim/replace}

=== 数据操作 (8个) ===
6. SortData - 排序 {range:必需, sortColumn:列号从1开始, order:asc/desc, hasHeader:默认true}
7. FilterData - 筛选 {range:必需, column:列号, criteria:筛选条件如"">100"", clearFilter:true则清除}
8. RemoveDuplicates - 删除重复 {range:必需, columns:列号数组(可选), hasHeader:可选}
9. ConditionalFormat - 条件格式 {range:必需, rule:highlight/databar/colorscale/iconset, condition:可选, color:可选}
10. MergeCells - 合并单元格 {range:必需, unmerge:true则取消合并}
11. AutoFit - 自动调整 {range:必需, type:columns/rows/both}
12. FindReplace - 查找替换 {range:""all""或指定范围, find:必需, replace:必需, matchCase:可选, matchEntireCell:可选}
13. CreatePivotTable - 透视表 {sourceRange:必需, targetCell:必需, rowFields:数组, valueFields:数组, columnFields:可选}

=== 工作表操作 (4个) ===
14. CreateSheet - 创建工作表 {name:必需, position:before/after, referenceSheet:可选}
15. DeleteSheet - 删除工作表 {name:必需}
16. RenameSheet - 重命名 {oldName:必需, newName:必需}
17. CopySheet - 复制工作表 {sourceName:必需, newName:必需}

=== 高级功能 (4个) ===
18. InsertRowCol - 插入行列 {type:row/column, position:行号或列字母, count:默认1}
19. DeleteRowCol - 删除行列 {type:row/column, position:必需, count:默认1}
20. HideRowCol - 隐藏行列 {type:row/column, position:必需, unhide:true则取消隐藏}
21. ProtectSheet - 保护工作表 {sheetName:可选, password:可选, unprotect:true则取消保护}

=== VBA回退 (1个) ===
22. ExecuteVBA - 执行VBA代码 {code:必需,完整的Sub或Function代码}
    当以上命令无法满足需求时,生成VBA代码作为回退方案

【动态范围占位符】
使用 {lastRow} 表示最后一行，{lastCol} 表示最后一列，{selection} 表示当前选择

【绝对禁止】
- 禁止使用 actions/operations 数组
- 禁止省略 params 包装
- 禁止自创其他命令（如translateText等）
- 禁止使用Word/PowerPoint专属命令
- 禁止返回不带代码块的裸JSON

【不支持的功能 - 请告知用户使用工具栏按钮】
- 翻译功能：请告知用户点击工具栏上的「AI翻译」按钮
- 校对功能：请告知用户点击工具栏上的「AI校对」按钮

【决策优先级】
1. 优先使用上述22个命令处理需求
2. 需求不明确时，用中文询问用户
3. 每个步骤的code字段必须是可直接执行的完整代码
4. 不要对JSON引号进行转义，不能包含\n、\r等换行转义字符，这不符合json规范
5. 只返回一个JSON对象，不要其他内容"

    ' === Word 规划提示词 ===
    Private Const PLANNING_SYSTEM_WORD As String = "你是一个智能Word自动化专家。请深度分析用户需求，结合识别到的意图、相关记忆和历史对话，制定一个详细、可执行的计划。每个步骤必须包含可直接执行的代码。

直接返回markdown的JSON代码块对象，格式如下:
```{
  ""understanding"": ""对用户需求的理解（结合意图和记忆）"",
  ""steps"": [
    {{
      ""step"": 1,
      ""description"": ""步骤描述（用户可读）"",
      ""code"": ""可执行的Excel JSON命令，不能返回\n\r等转义字符"",
      ""language"": ""json""
    }}
  ],
  ""summary"": ""执行完成后的预期结果""
}```

【Word代码格式要求】
code字段必须是JSON命令格式:
- 单命令: {{""command"":""InsertText"",""params"":{{""content"":""文本""}}}}
- 多命令: {{""commands"":[{{""command"":""InsertText"",""params"":{{""content"":""文本""}}}},...]}}

【Word支持的22个命令】

=== 基础文本操作 (5个) ===
1. InsertText - 插入文本 {content:必需, position:cursor/start/end}
2. FormatText - 格式化 {range:selection/all, bold/italic/fontSize/fontName/underline/color}
3. ReplaceText - 查找替换 {find:必需, replace:必需, matchCase:可选}
4. DeleteText - 删除文本 {range:selection/all}
5. CopyPasteText - 复制粘贴 {sourceRange:必需, targetPosition:可选}

=== 段落和样式 (5个) ===
6. ApplyStyle - 应用样式 {styleName:必需如""标题 1"", range:selection/paragraph}
7. SetParagraphFormat - 段落格式 {alignment:left/center/right/justify, firstLineIndent/beforeSpacing/afterSpacing}
8. InsertParagraph - 插入段落 {count:默认1, pageBreak:true则分页}
9. SetLineSpacing - 行距 {spacing:1/1.5/2, range:selection/all}
10. SetIndent - 缩进 {left/right/firstLine:cm值}

=== 表格操作 (4个) ===
11. InsertTable - 插入表格 {rows:必需, cols:必需, data:可选}
12. FormatTable - 格式化表格 {tableIndex:从1开始, style/borders/headerRow}
13. InsertTableRow - 插入行 {tableIndex:必需, position:after/before}
14. DeleteTableRow - 删除行 {tableIndex:必需, rowIndex:必需}

=== 文档结构 (4个) ===
15. GenerateTOC - 生成目录 {position:start/cursor, levels:1-9}
16. InsertHeader - 页眉 {content:必需, alignment:left/center/right}
17. InsertFooter - 页脚 {content:必需, alignment:left/center/right}
18. InsertPageNumber - 页码 {position:header/footer, alignment}

=== 文档美化 (2个) ===
19. BeautifyDocument - 美化 {theme:{h1/h2/body设置}, margins:{top/bottom/left/right}}
20. SetPageMargins - 页边距 {top/bottom/left/right:cm值}

=== 高级功能 (1个) ===
21. InsertImage - 插入图片 {imagePath:必需, width/height:可选}

=== VBA回退 (1个) ===
22. ExecuteVBA - VBA代码 {code:必需,完整Sub/Function}

【绝对禁止】
- 禁止使用 actions/operations 数组
- 禁止省略 params 包装
- 禁止使用Excel/PowerPoint专属命令

【决策优先级】
1. 优先使用上述22个命令
2. 复杂需求用ExecuteVBA
3. 翻译用工具栏按钮
4. 需求不明确时中文询问

【重要】Word文本格式要求:
1. 段落之间必须使用 \n\n (两个换行符) 分隔
2. 标题后必须加 \n\n
3. 需要缩进时使用全角空格或多个半角空格
4. 示例: {{""command"":""InsertText"",""params"":{{""content"":""请假申请单\n\n    申请人：张三\n    日期：2024年1月1日\n\n事由：...""}}}}

注意：
1. 每个步骤的code字段必须是可直接执行的完整代码
2. 不要对JSON引号进行转义
3. 只返回一个JSON对象，不要其他内容"

    ' === PowerPoint 规划提示词 ===
    Private Const PLANNING_SYSTEM_PPT As String = "你是一个智能PowerPoint自动化专家。请深度分析用户需求，结合识别到的意图、相关记忆和历史对话，制定一个详细、可执行的计划。每个步骤必须包含可直接执行的代码。

直接返回markdown的JSON代码块对象，格式如下:
```{
  ""understanding"": ""对用户需求的理解（结合意图和记忆）"",
  ""steps"": [
    {{
      ""step"": 1,
      ""description"": ""步骤描述（用户可读）"",
      ""code"": ""可执行的PowerPoint VBA或JSON命令"",
      ""language"": ""json""
    }}
  ],
  ""summary"": ""执行完成后的预期结果""
}```

【PowerPoint代码格式要求】
- JSON命令: {{""command"":""InsertSlide"",""params"":{{""title"":""标题""}}}}
- VBA代码通过ExecuteVBA执行，code字段内换行必须用\n转义，不能有真实换行

【PowerPoint支持的22个命令】

=== 幻灯片操作 (5个) ===
1. InsertSlide - 插入幻灯片 {position:current/end, layout, title, content}
2. DeleteSlide - 删除幻灯片 {slideIndex:必需,-1当前}
3. DuplicateSlide - 复制幻灯片 {slideIndex:必需}
4. MoveSlide - 移动幻灯片 {fromIndex:必需, toIndex:必需}
5. CreateSlides - 批量创建 {slides:数组含title/content/layout}

=== 内容操作 (5个) ===
6. InsertText - 插入文本 {content:必需, slideIndex:-1当前, x/y:可选}
7. FormatText - 格式化文本 {bold/italic/fontSize/fontName/color}
8. InsertShape - 插入形状 {shapeType:必需, x:必需, y:必需}
9. InsertImage - 插入图片 {imagePath:必需, x/y/width/height:可选}
10. InsertTable - 插入表格 {rows:必需, cols:必需, data:可选}

=== 样式和动画 (5个) ===
11. FormatSlide - 格式化幻灯片 {background, layout}
12. AddAnimation - 添加动画 {effect:fadeIn/flyIn/zoom/wipe, targetShapes:all/title}
13. ApplyTransition - 切换效果 {transitionType:fade/push/wipe, scope:all/current}
14. BeautifySlides - 美化 {scope:all/current, theme:{background/titleFont/bodyFont}}
15. SetSlideLayout - 设置布局 {layout:title/titleAndContent/blank}

=== 高级功能 (4个) ===
16. InsertChart - 插入图表 {chartType:column/line/pie, data:二维数组}
17. InsertVideo - 插入视频 {videoPath:必需, autoPlay:可选}
18. AddSpeakerNotes - 演讲备注 {notes:必需, slideIndex:可选}
19. SetSlideShow - 放映设置 {loopUntilEsc/advanceMode等}

=== 母版和主题 (2个) ===
20. ApplyTheme - 应用主题 {themeName或themeFile}
21. EditSlideMaster - 编辑母版 {background/titleFont/bodyFont}

=== VBA回退 (1个) ===
22. ExecuteVBA - VBA代码 {code:必需,完整Sub/Function}

【绝对禁止】
- 禁止使用 actions/operations 数组
- 禁止省略 params 包装
- 禁止使用Excel/Word专属命令

【VBA代码规范】
- 嵌套For Each循环必须使用不同变量名，如外层shp内层用innerShp
- 变量声明必须带As关键字: Dim p As Paragraph，不能写Dim p Paragraph
- PowerPoint中检查文本框用shp.HasTextFrame，不能用TypeOf比较形状类型
- PowerPoint VBA不能用Excel的Range类型，文本范围用TextRange
- 所有变量在Sub开头用Dim声明
- VBA换行用\n转义，不能有真实换行
- 输出前自检：每个Dim后必须有As，每个For有Next，每个If有End If

【决策优先级】
1. 优先使用上述22个命令
2. 复杂需求用ExecuteVBA
3. 翻译用工具栏按钮
4. 需求不明确时中文询问

注意：
1. 每个步骤的code字段必须是可直接执行的完整代码
2. VBA代码的换行必须使用\n转义，以保证JSON格式正确
3. 只返回一个JSON对象，不要其他内容"

    ' === 规划 user 模板（所有应用共用，历史对话通过 messages 数组传递，不再嵌入文本） ===
    Private Const PLANNING_USER_TEMPLATE As String = "当前选中/活动的内容:
```
{0}
```

用户当前需求: {1}

【意图识别结果】
{2}

【相关记忆（RAG检索）】
{3}"

    Private Function GetPlanningPrompt(appType As String) As Tuple(Of String, String)
        Dim sys As String
        Select Case appType.ToLower()
            Case "excel"
                sys = PLANNING_SYSTEM_EXCEL
            Case "word"
                sys = PLANNING_SYSTEM_WORD
            Case "powerpoint", "ppt"
                sys = PLANNING_SYSTEM_PPT
            Case Else
                sys = PLANNING_SYSTEM_EXCEL
        End Select
        ' 提示词常量按 String.Format 惯例用 {{ }} 转义花括号，此处还原为单括号再发给 AI
        sys = sys.Replace("{{", "{").Replace("}}", "}")
        Return Tuple.Create(sys, PLANNING_USER_TEMPLATE)
    End Function

    Private Const STEP_EXECUTION_SYSTEM As String = "你是一个Office自动化专家。请根据Office类型生成可以直接执行的代码。

【Excel】必须使用JSON命令格式:
- 单命令: {{""command"":""ApplyFormula"",""params"":{{""targetRange"":""C1:C100"",""formula"":""=A1+B1""}}}}
- 多命令: {{""commands"":[{{""command"":""ApplyFormula"",""params"":{{...}}}},...]}}
- Excel只支持这些command: ApplyFormula, WriteData, FormatRange, CreateChart, CleanData
- ApplyFormula参数: targetRange(必需), formula(必需), fillDown(可选)
- WriteData参数: targetRange(必需), data(必需)
- FormatRange参数: range(必需), style(可选)
- CreateChart参数: dataRange(必需), chartType(可选)
- CleanData参数: range(必需), operation(必需)
- 动态范围使用{{lastRow}}占位符

【Word】必须使用JSON命令格式:
- 单命令: {{""command"":""InsertText"",""params"":{{""content"":""文本""}}}}
- 多命令: {{""commands"":[{{""command"":""InsertText"",""params"":{{""content"":""文本""}}}},...]}}
- Word支持的command: InsertText, FormatText, ReplaceText, InsertTable, ApplyStyle, GenerateTOC, BeautifyDocument

【PowerPoint】使用VBA代码

只返回可执行的代码，用```vba或```json包裹。"

    Private Const STEP_EXECUTION_USER As String = "当前Office应用: {0}
当前文档内容:
```
{1}
```

要执行的步骤: {2}
步骤详情: {3}"

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
            _ragMemories = MemoryService.GetRelevantMemories(userRequest, 5, Nothing, Nothing, appType)
            Debug.WriteLine($"[RalphAgent] 检索到 {_ragMemories.Count} 条相关记忆")

            ' 步骤3: 转换历史对话为 HistoryMessage 列表（符合 OpenAI API messages 规范）
            Dim historyMsgs As New List(Of HistoryMessage)()
            If historyMessages IsNot Nothing Then
                Dim limit = MemoryConfig.SessionSummaryLimit
                Dim startIdx = Math.Max(0, historyMessages.Count - limit)
                For i = startIdx To historyMessages.Count - 1
                    Dim msg = historyMessages(i)
                    Dim role = If(msg.Item1 = "user", "user", "assistant")
                    If Not String.IsNullOrEmpty(msg.Item2) Then
                        historyMsgs.Add(New HistoryMessage With {.role = role, .content = msg.Item2})
                    End If
                Next
            End If
            Debug.WriteLine($"[RalphAgent] 历史对话包含 {historyMsgs.Count} 条消息，将作为 messages 数组发送")

            ' 步骤4: 制定执行计划
            OnStatusChanged?.Invoke("正在制定执行计划...")
            Dim intentInfo = FormatIntentInfo(_intentResult)
            Dim ragInfo = FormatRagMemories(_ragMemories)

            Dim planningParts = GetPlanningPrompt(appType)
            Dim systemPrompt = planningParts.Item1
            Dim userPrompt = String.Format(planningParts.Item2, currentContent, userRequest, intentInfo, ragInfo)
            Debug.WriteLine($"[RalphAgent] system 提示词长度: {systemPrompt.Length}, user 提示词长度: {userPrompt.Length}")
            Dim response = Await SendAIRequest(userPrompt, systemPrompt, historyMsgs)

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
            fixedJson = fixedJson.Replace("\n", "")
            fixedJson = fixedJson.Replace("\r", "")

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

            Debug.WriteLine($"[RalphAgent] 提取到JSON: {jsonStr}")

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
                Dim stepUserPrompt = String.Format(STEP_EXECUTION_USER,
                    _currentSession.ApplicationType,
                    _currentSession.CurrentContent,
                    currentStep.Description,
                    currentStep.Detail)

                Dim codeResponse = Await SendAIRequest(stepUserPrompt, STEP_EXECUTION_SYSTEM.Replace("{{", "{").Replace("}}", "}"), Nothing)
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
