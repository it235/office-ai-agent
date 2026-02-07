' ShareRibbon\Config\PromptManager.vb
' 统一提示词管理中心

Imports System.IO
Imports System.Text
Imports Newtonsoft.Json

''' <summary>
''' 统一提示词管理中心 - 单例模式
''' 负责管理所有提示词的加载、组合和获取
''' </summary>
Public Class PromptManager
    Private Shared _instance As PromptManager
    Private _promptConfig As PromptConfiguration

    ''' <summary>
    ''' 获取单例实例
    ''' </summary>
    Public Shared ReadOnly Property Instance As PromptManager
        Get
            If _instance Is Nothing Then
                _instance = New PromptManager()
            End If
            Return _instance
        End Get
    End Property

    ''' <summary>
    ''' 私有构造函数
    ''' </summary>
    Private Sub New()
        LoadPromptConfiguration()
    End Sub

    ''' <summary>
    ''' 重新加载配置
    ''' </summary>
    Public Sub ReloadConfiguration()
        LoadPromptConfiguration()
    End Sub

    ''' <summary>
    ''' 加载提示词配置
    ''' </summary>
    Private Sub LoadPromptConfiguration()
        Dim configPath = GetPromptConfigPath()

        Try
            If File.Exists(configPath) Then
                Dim json = File.ReadAllText(configPath, Encoding.UTF8)
                _promptConfig = JsonConvert.DeserializeObject(Of PromptConfiguration)(json)
            Else
                ' 使用默认配置
                _promptConfig = CreateDefaultConfiguration()
                SavePromptConfiguration()
            End If
        Catch ex As Exception
            Debug.WriteLine($"加载提示词配置失败: {ex.Message}")
            _promptConfig = CreateDefaultConfiguration()
        End Try
    End Sub

    ''' <summary>
    ''' 保存提示词配置
    ''' </summary>
    Public Sub SavePromptConfiguration()
        Try
            Dim configPath = GetPromptConfigPath()
            Dim dir = Path.GetDirectoryName(configPath)

            If Not Directory.Exists(dir) Then
                Directory.CreateDirectory(dir)
            End If

            Dim json = JsonConvert.SerializeObject(_promptConfig, Formatting.Indented)
            File.WriteAllText(configPath, json, Encoding.UTF8)
        Catch ex As Exception
            Debug.WriteLine($"保存提示词配置失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 更新指定应用的JSON Schema约束
    ''' </summary>
    Public Sub UpdateJsonSchemaConstraint(appType As String, constraint As String)
        Dim appConfig = _promptConfig.Applications.FirstOrDefault(Function(a) a.Type.Equals(appType, StringComparison.OrdinalIgnoreCase))
        If appConfig IsNot Nothing Then
            appConfig.JsonSchemaConstraint = constraint
        End If
    End Sub

    ''' <summary>
    ''' 重置指定应用的JSON Schema约束为默认值
    ''' </summary>
    Public Sub ResetJsonSchemaConstraint(appType As String)
        Dim defaultConfig = CreateDefaultConfiguration()
        Dim defaultAppConfig = defaultConfig.Applications.FirstOrDefault(Function(a) a.Type.Equals(appType, StringComparison.OrdinalIgnoreCase))

        If defaultAppConfig IsNot Nothing Then
            Dim currentAppConfig = _promptConfig.Applications.FirstOrDefault(Function(a) a.Type.Equals(appType, StringComparison.OrdinalIgnoreCase))
            If currentAppConfig IsNot Nothing Then
                currentAppConfig.JsonSchemaConstraint = defaultAppConfig.JsonSchemaConstraint
            End If
        End If
    End Sub

    ''' <summary>
    ''' 获取组合后的提示词（融合模式）
    ''' </summary>
    ''' <param name="context">提示词上下文</param>
    ''' <returns>组合后的完整提示词</returns>
    Public Function GetCombinedPrompt(context As PromptContext) As String
        Dim sb As New StringBuilder()

        ' 判断是否为功能性模式（校对/排版/续写/模板渲染等）
        ' 功能性模式不使用用户配置的提示词，避免污染
        Dim isInFunctionalMode As Boolean = CheckIsFunctionalMode(context.FunctionMode)

        ' 1. 用户配置提示词（仅在非功能性模式下使用）
        If Not isInFunctionalMode AndAlso Not String.IsNullOrEmpty(ConfigSettings.propmtContent) Then
            sb.AppendLine(ConfigSettings.propmtContent)
            sb.AppendLine()
        End If

        ' 2. 意图专用提示词（仅在非功能性模式下使用，置信度>0.2时）
        If Not isInFunctionalMode AndAlso context.IntentResult IsNot Nothing AndAlso context.IntentResult.Confidence > 0.2 Then
            Dim intentPrompt = GetIntentSpecificPrompt(context)
            If Not String.IsNullOrEmpty(intentPrompt) Then
                sb.AppendLine(intentPrompt)
                sb.AppendLine()
            End If
        End If

        ' 3. 功能模式提示词（校对/排版/续写/模板渲染）
        If Not String.IsNullOrEmpty(context.FunctionMode) Then
            Dim modePrompt = GetFunctionModePrompt(context)
            If Not String.IsNullOrEmpty(modePrompt) Then
                sb.AppendLine(modePrompt)
                sb.AppendLine()
            End If
        End If

        ' 4. 输出格式约束（JSON Schema或纯文本）- 仅在非功能性模式下添加
        If Not isInFunctionalMode Then
            Dim formatConstraint = GetOutputFormatConstraint(context)
            If Not String.IsNullOrEmpty(formatConstraint) Then
                sb.AppendLine(formatConstraint)
            End If
        End If

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 判断是否为功能性模式（这些模式使用专用提示词，不受用户配置和历史记录影响）
    ''' </summary>
    Private Function CheckIsFunctionalMode(functionMode As String) As Boolean
        If String.IsNullOrEmpty(functionMode) Then Return False

        Dim functionalModes As String() = {"proofread", "reformat", "continuation", "template_render"}
        Return functionalModes.Contains(functionMode.ToLower())
    End Function

    ''' <summary>
    ''' 获取意图专用提示词
    ''' </summary>
    Private Function GetIntentSpecificPrompt(context As PromptContext) As String
        Dim appType = context.ApplicationType
        Dim intentType = If(context.IntentResult?.OfficeIntent.ToString(), "GENERAL_QUERY")

        ' 从配置中查找对应的意图提示词
        Dim appConfig = _promptConfig.Applications.FirstOrDefault(Function(a) a.Type.Equals(appType, StringComparison.OrdinalIgnoreCase))
        If appConfig Is Nothing Then Return String.Empty

        Dim intentPrompt = appConfig.IntentPrompts.FirstOrDefault(Function(p) p.IntentType.Equals(intentType, StringComparison.OrdinalIgnoreCase))
        If intentPrompt Is Nothing Then
            ' 如果没有找到特定意图，使用通用提示词
            intentPrompt = appConfig.IntentPrompts.FirstOrDefault(Function(p) p.IntentType.Equals("GENERAL_QUERY", StringComparison.OrdinalIgnoreCase))
        End If

        Return If(intentPrompt?.Content, String.Empty)
    End Function

    ''' <summary>
    ''' 获取功能模式专用提示词
    ''' </summary>
    Private Function GetFunctionModePrompt(context As PromptContext) As String
        Dim appType = context.ApplicationType

        ' 从配置中查找对应的功能模式提示词
        Dim appConfig = _promptConfig.Applications.FirstOrDefault(Function(a) a.Type.Equals(appType, StringComparison.OrdinalIgnoreCase))
        If appConfig Is Nothing Then Return String.Empty

        Dim modePrompt = appConfig.FunctionModePrompts.FirstOrDefault(Function(p) p.Mode.Equals(context.FunctionMode, StringComparison.OrdinalIgnoreCase))
        Return If(modePrompt?.Content, String.Empty)
    End Function

    ''' <summary>
    ''' 获取输出格式约束
    ''' </summary>
    Private Function GetOutputFormatConstraint(context As PromptContext) As String
        ' 判断是否需要JSON输出
        Dim needsJsonOutput = DetermineIfNeedsJsonOutput(context)

        If needsJsonOutput Then
            ' 返回JSON Schema约束
            Return GetJsonSchemaConstraint(context.ApplicationType)
        Else
            ' 返回纯文本输出约束
            Return GetPlainTextConstraint(context.FunctionMode)
        End If
    End Function

    ''' <summary>
    ''' 判断是否需要JSON输出
    ''' </summary>
    Private Function DetermineIfNeedsJsonOutput(context As PromptContext) As Boolean
        ' 特殊功能模式使用特定格式
        Select Case context.FunctionMode?.ToLower()
            Case "continuation", "template_render"
                ' 续写和模板渲染返回纯文本
                Return False
            Case "proofread", "reformat"
                ' 校对和排版返回JSON数组
                Return True
        End Select

        ' Office应用始终需要JSON Schema约束
        ' 因为用户请求可能涉及命令操作（如翻译、公式、图表等）
        ' JSON Schema中已说明"对于简单问候或问答，直接用中文回复"
        Dim appType = context.ApplicationType?.ToLower()
        If appType = "excel" OrElse appType = "word" OrElse appType = "powerpoint" Then
            Return True
        End If

        ' 意图判断
        If context.IntentResult IsNot Nothing Then
            ' 高置信度（>0.7）时，即使是GENERAL_QUERY也返回JSON（用户需求明确）
            If context.IntentResult.Confidence > 0.7 Then
                Return True
            End If

            ' 中等置信度（>0.2）且不是一般查询，需要JSON
            If context.IntentResult.Confidence > 0.2 AndAlso
               context.IntentResult.OfficeIntent <> OfficeIntentType.GENERAL_QUERY Then
                Return True
            End If
        End If

        Return False
    End Function

    ''' <summary>
    ''' 获取JSON Schema约束（根据Office应用类型）- 供外部调用
    ''' </summary>
    Public Function GetJsonSchemaConstraint(appType As String) As String
        Dim appConfig = _promptConfig.Applications.FirstOrDefault(Function(a) a.Type.Equals(appType, StringComparison.OrdinalIgnoreCase))
        Dim userConstraint = appConfig?.JsonSchemaConstraint
        
        ' 如果用户配置为空或明显不完整，则使用内置默认值
        If String.IsNullOrEmpty(userConstraint) OrElse Not IsValidJsonSchemaConstraint(userConstraint) Then
            Return GetDefaultJsonSchemaConstraint(appType)
        End If
        
        Return userConstraint
    End Function
    
    ''' <summary>
    ''' 获取默认的JSON Schema约束（内置硬编码）
    ''' </summary>
    Private Function GetDefaultJsonSchemaConstraint(appType As String) As String
        Select Case appType?.ToLower()
            Case "excel"
                Return GetExcelJsonSchemaConstraintDefault()
            Case "word"
                Return GetWordJsonSchemaConstraintDefault()
            Case "powerpoint"
                Return GetPowerPointJsonSchemaConstraintDefault()
            Case Else
                Return GetExcelJsonSchemaConstraintDefault()
        End Select
    End Function
    
    ''' <summary>
    ''' Excel专用JSON Schema约束（默认值）
    ''' </summary>
    Private Function GetExcelJsonSchemaConstraintDefault() As String
        Return "
【Excel JSON输出格式规范 - 必须严格遵守】

【重要】JSON必须使用Markdown代码块格式返回，例如：
```json
{""command"": ""ApplyFormula"", ""params"": {...}}
```
禁止直接返回裸JSON文本！

你必须且只能返回以下两种格式之一：

单命令格式（必须包含command字段）：
```json
{""command"": ""ApplyFormula"", ""params"": {""targetRange"": ""A1:B10"", ""formula"": ""=SUM(A1:A10)""}}
```

多命令格式（必须包含commands数组）：
```json
{""commands"": [{""command"": ""WriteData"", ""params"": {""data"": [[""姓名"", ""年龄""]], ""targetRange"": ""A1""}}, {""command"": ""FormatRange"", ""params"": {""range"": ""A1:B1"", ""style"": ""header""}}]}
```

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
2. 复杂需求无法用命令实现时，使用ExecuteVBA生成VBA代码
3. 需求不明确时，用中文询问用户"
    End Function
    
    ''' <summary>
    ''' Word专用JSON Schema约束（默认值）
    ''' </summary>
    Private Function GetWordJsonSchemaConstraintDefault() As String
        Return "
【Word JSON输出格式规范 - 必须严格遵守】

【重要】JSON必须使用Markdown代码块格式返回，例如：
```json
{""command"": ""InsertText"", ""params"": {...}}
```
禁止直接返回裸JSON文本！

你必须且只能返回以下两种格式之一：

单命令格式：
```json
{""command"": ""InsertText"", ""params"": {""content"": ""插入的内容"", ""position"": ""cursor""}}
```

多命令格式：
```json
{""commands"": [{""command"": ""InsertText"", ""params"": {""content"": ""标题""}}, {""command"": ""FormatText"", ""params"": {""range"": ""selection"", ""bold"": true}}]}
```

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
4. 需求不明确时中文询问"
    End Function
    
    ''' <summary>
    ''' PowerPoint专用JSON Schema约束（默认值）
    ''' </summary>
    Private Function GetPowerPointJsonSchemaConstraintDefault() As String
        Return "
【PowerPoint JSON输出格式规范 - 必须严格遵守】

【重要】JSON必须使用Markdown代码块格式返回，例如：
```json
{""command"": ""InsertSlide"", ""params"": {...}}
```
禁止直接返回裸JSON文本！

单命令格式：
```json
{""command"": ""InsertSlide"", ""params"": {""title"": ""标题"", ""layout"": ""titleAndContent""}}
```

多命令格式：
```json
{""commands"": [{""command"": ""CreateSlides"", ""params"": {""slides"": [{""title"": ""第一页""}]}}, {""command"": ""AddAnimation"", ""params"": {""effect"": ""fadeIn"", ""scope"": ""all""}}]}
```

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

【决策优先级】
1. 优先使用上述22个命令
2. 复杂需求用ExecuteVBA
3. 翻译用工具栏按钮
4. 需求不明确时中文询问"
    End Function
    
    ''' <summary>
    ''' 验证JSON Schema约束是否有效（包含必要的格式要求）
    ''' </summary>
    Private Function IsValidJsonSchemaConstraint(constraint As String) As Boolean
        If String.IsNullOrEmpty(constraint) Then Return False
        
        ' 检查是否包含关键约束词汇
        Dim requiredKeywords() As String = {
            "JSON必须使用Markdown代码块格式",
            "禁止直接返回裸JSON文本",
            "command",
            "commands",
            "params"
        }
        
        For Each keyword In requiredKeywords
            If Not constraint.Contains(keyword) Then
                Return False
            End If
        Next
        
        Return True
    End Function
    
    ''' <summary>
    ''' 根据字符串获取应用类型枚举
    ''' </summary>
    Private Function GetApplicationTypeFromString(appType As String) As OfficeApplicationType
        Select Case appType?.ToLower()
            Case "excel"
                Return OfficeApplicationType.Excel
            Case "word" 
                Return OfficeApplicationType.Word
            Case "powerpoint"
                Return OfficeApplicationType.PowerPoint
            Case Else
                Return OfficeApplicationType.Excel ' 默认值
        End Select
    End Function

    ''' <summary>
    ''' 获取纯文本输出约束
    ''' </summary>
    Private Function GetPlainTextConstraint(functionMode As String) As String
        Select Case functionMode?.ToLower()
            Case "continuation"
                Return "
【重要输出要求】
- 只输出续写内容，不要添加任何前缀、后缀或说明
- 保持与原文一致的语言风格和术语
- 内容要连贯自然，不要重复上文已有内容"

            Case "template_render"
                Return "
【重要格式要求】
- 严禁使用Markdown代码块格式（禁止使用```符号）
- 严禁使用任何Markdown格式标记（如#、**、-、>等）
- 直接输出纯文本内容，不要包装在任何代码块中
- 不要添加任何前缀、后缀、解释或说明文字"

            Case Else
                Return String.Empty
        End Select
    End Function

    ''' <summary>
    ''' 获取配置文件路径
    ''' </summary>
    Private Function GetPromptConfigPath() As String
        Return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            ConfigSettings.OfficeAiAppDataFolder,
            "prompt_templates.json")
    End Function

    ''' <summary>
    ''' 创建默认配置
    ''' </summary>
    Private Function CreateDefaultConfiguration() As PromptConfiguration
        Dim config As New PromptConfiguration()

        ' Excel应用配置
        config.Applications.Add(CreateExcelConfig())

        ' Word应用配置
        config.Applications.Add(CreateWordConfig())

        ' PowerPoint应用配置
        config.Applications.Add(CreatePowerPointConfig())

        Return config
    End Function

    ''' <summary>
    ''' 创建Excel默认配置
    ''' </summary>
    Private Function CreateExcelConfig() As ApplicationPromptConfig
        Dim excelApp As New ApplicationPromptConfig With {
            .Type = "Excel"
        }

        ' 意图提示词
        excelApp.IntentPrompts.Add(New IntentPromptTemplate With {
            .IntentType = "DATA_ANALYSIS",
            .Content = "你是Excel数据分析助手。如果用户需求明确，返回JSON代码片段执行数据分析。如果用户需求不明确，请先询问用户想要什么样的分析结果。"
        })

        excelApp.IntentPrompts.Add(New IntentPromptTemplate With {
            .IntentType = "FORMULA_CALC",
            .Content = "你是Excel公式助手。如果用户需求明确，返回JSON代码片段应用公式。如果用户需求不明确，请先询问用户具体想计算什么。"
        })

        excelApp.IntentPrompts.Add(New IntentPromptTemplate With {
            .IntentType = "CHART_GEN",
            .Content = "你是Excel图表助手。如果用户需求明确，返回JSON代码片段创建图表。请根据数据特点推荐合适的图表类型（柱状图、折线图、饼图等）。"
        })

        excelApp.IntentPrompts.Add(New IntentPromptTemplate With {
            .IntentType = "DATA_CLEANING",
            .Content = "你是Excel数据清洗助手。如果用户需求明确，返回JSON代码片段清洗数据。支持去重、填充空值、去除空格等操作。"
        })

        excelApp.IntentPrompts.Add(New IntentPromptTemplate With {
            .IntentType = "GENERAL_QUERY",
            .Content = "你是Excel助手。如果用户需求明确且可以执行，返回JSON代码片段；如果用户需求不明确，请先询问用户澄清；对于简单问候或问答，直接用中文回复。"
        })

        ' JSON Schema约束
        excelApp.JsonSchemaConstraint = "
【Excel JSON输出格式规范 - 必须严格遵守】

【重要】JSON必须使用Markdown代码块格式返回，例如：
```json
{""command"": ""ApplyFormula"", ""params"": {...}}
```
禁止直接返回裸JSON文本！

你必须且只能返回以下两种格式之一：

单JSON代码格式：
```json
{""command"": ""ApplyFormula"", ""params"": {""targetRange"": ""C1:C{lastRow}"", ""formula"": ""=A1+B1"", ""fillDown"": true}}
```

多JSON代码格式：
```json
{""commands"": [{""command"": ""ApplyFormula"", ""params"": {...}}, {...}]}
```

【绝对禁止】
- 禁止使用 actions 数组
- 禁止使用 operations 数组
- 禁止省略 params 包装
- 禁止返回下方未指定的command类型
- 禁止返回不带代码块的裸JSON

【Excel command类型】
1. ApplyFormula - 应用公式 (targetRange, formula, fillDown)
2. WriteData - 写入数据 (targetRange, data)
3. FormatRange - 格式化范围 (range, style, bold, fontSize, fontColor, bgColor)
4. CreateChart - 创建图表 (dataRange, chartType, title)
5. CleanData - 清洗数据 (range, operation: removeDuplicates/fillEmpty/trim)

【动态范围占位符】
使用 {lastRow} 表示最后一行，系统会自动替换为实际行号"

        Return excelApp
    End Function

    ''' <summary>
    ''' 创建Word默认配置
    ''' </summary>
    Private Function CreateWordConfig() As ApplicationPromptConfig
        Dim wordApp As New ApplicationPromptConfig With {
            .Type = "Word"
        }

        ' 意图提示词
        wordApp.IntentPrompts.Add(New IntentPromptTemplate With {
            .IntentType = "DOCUMENT_EDIT",
            .Content = "你是Word文档编辑助手。如果用户需求明确，返回JSON代码片段执行文档编辑操作。支持插入、删除、替换文本等操作。"
        })

        wordApp.IntentPrompts.Add(New IntentPromptTemplate With {
            .IntentType = "TOC_GENERATION",
            .Content = "你是Word目录生成助手。如果用户说'生成目录'或'添加目录'，直接返回GenerateTOC命令。如果需要澄清，询问：目录放在开头还是当前位置？显示几级标题？"
        })

        wordApp.IntentPrompts.Add(New IntentPromptTemplate With {
            .IntentType = "FORMAT_STYLE",
            .Content = "你是Word格式样式助手。如果用户需要美化文档，返回BeautifyDocument命令。支持应用主题、设置字体、调整间距等。"
        })

        wordApp.IntentPrompts.Add(New IntentPromptTemplate With {
            .IntentType = "GENERAL_QUERY",
            .Content = "你是Word助手。如果用户需求明确且可以执行，返回JSON代码片段；如果用户需求不明确，请先询问用户澄清；对于简单问候或问答，直接用中文回复。"
        })

        ' 功能模式提示词
        wordApp.FunctionModePrompts.Add(New FunctionModePromptTemplate With {
            .Mode = "proofread",
            .Content = "你是Word内容校对助手。请检查以下内容中的错字、错标点或不当换行，并给出修正建议。

【输出格式】
必须返回JSON数组，每个元素包含：
[{""paraIndex"": 0, ""original"": ""原文片段"", ""corrected"": ""修正后的文字"", ""reason"": ""简短说明修正原因""}]

如果没有需要修正的内容，返回空数组 []"
        })

        wordApp.FunctionModePrompts.Add(New FunctionModePromptTemplate With {
            .Mode = "reformat",
            .Content = "你是Word排版助手。我提供文档段落，请帮我优化排版。

【排版规则】
1. 中文字体使用宋体，英文使用Times New Roman
2. 正文字号12pt（小四），标题根据级别设置（如16pt/14pt）
3. 段落首行缩进2字符
4. 行距1.5倍

【输出格式】
必须返回JSON数组，格式如下：
[{""paraIndex"": 0, ""formatting"": {""fontNameCN"": ""宋体"", ""fontNameEN"": ""Times New Roman"", ""fontSize"": 12, ""bold"": false, ""alignment"": ""left"", ""firstLineIndent"": 2, ""lineSpacing"": 1.5}}]"
        })

        wordApp.FunctionModePrompts.Add(New FunctionModePromptTemplate With {
            .Mode = "continuation",
            .Content = "你是一个专业的写作助手。根据提供的上下文，自然地续写内容。

要求：
1. 保持与原文一致的语言风格、语气和术语
2. 内容要连贯自然，不要重复上文已有内容
3. 只输出续写内容，不要添加任何解释、前缀或标记
4. 如果上下文不足，可以合理推断但保持谨慎
5. 续写长度适中，约100-300字，除非用户另有要求"
        })

        wordApp.FunctionModePrompts.Add(New FunctionModePromptTemplate With {
            .Mode = "template_render",
            .Content = "你是一个专业的文档内容生成助手。你需要根据用户提供的模板结构（JSON格式）和风格来生成新的内容。

【模板JSON结构说明】
- elements: 文档元素数组，每个元素包含type(类型)、text(文本)、styleName(样式名)、formatting(格式详情)
- formatting包含: fontName(字体)、fontSize(字号)、bold(加粗)、italic(斜体)、alignment(对齐)等

【内容生成要求】
1. 严格遵循模板的层级结构（如：标题、副标题、正文的层次关系）
2. 保持与模板一致的语气、术语规范和风格
3. 参考模板中的字号来判断内容的重要程度（大字号=标题，小字号=正文）
4. 内容要专业、连贯、符合实际使用场景
5. 按照模板中元素的顺序来组织输出内容"
        })

        ' JSON Schema约束
        wordApp.JsonSchemaConstraint = "
【Word JSON输出格式规范 - 必须严格遵守】

【重要】JSON必须使用Markdown代码块格式返回，例如：
```json
{""command"": ""InsertText"", ""params"": {...}}
```
禁止直接返回裸JSON文本！

你必须且只能返回以下两种格式之一：

单JSON代码格式：
```json
{""command"": ""InsertText"", ""params"": {""content"": ""内容"", ""position"": ""cursor""}}
```

多JSON代码格式：
```json
{""commands"": [{""command"": ""InsertText"", ""params"": {...}}, {...}]}
```

【Word command类型】
1. InsertText - 插入文本 (content, position: cursor/start/end)
2. FormatText - 格式化文本 (range: selection/all, bold, italic, fontSize, fontName)
3. ReplaceText - 替换文本 (find, replace, matchCase, matchWholeWord)
4. InsertTable - 插入表格 (rows, cols, data)
5. ApplyStyle - 应用样式 (styleName, range)
6. GenerateTOC - 生成目录 (position: start/cursor, levels: 1-9)
7. BeautifyDocument - 美化文档 (theme, margins)

【绝对禁止】
- 禁止使用Excel命令(WriteData, ApplyFormula等)
- 禁止使用PPT命令(InsertSlide, CreateSlides等)
- 禁止返回上方未指定的command类型
- 禁止返回不带代码块的裸JSON"

        Return wordApp
    End Function

    ''' <summary>
    ''' 创建PowerPoint默认配置
    ''' </summary>
    Private Function CreatePowerPointConfig() As ApplicationPromptConfig
        Dim pptApp As New ApplicationPromptConfig With {
            .Type = "PowerPoint"
        }

        ' 意图提示词
        pptApp.IntentPrompts.Add(New IntentPromptTemplate With {
            .IntentType = "SLIDE_CREATE",
            .Content = "你是PowerPoint幻灯片创建助手。当用户说'生成N页PPT'时，使用CreateSlides命令批量创建；当用户说'添加一页'时，使用InsertSlide命令创建单页。"
        })

        pptApp.IntentPrompts.Add(New IntentPromptTemplate With {
            .IntentType = "ANIMATION_EFFECT",
            .Content = "你是PowerPoint动画效果助手。支持添加进入动画（fadeIn、flyIn、zoom、wipe等）和退出动画。可以为所有形状或仅标题添加动画。"
        })

        pptApp.IntentPrompts.Add(New IntentPromptTemplate With {
            .IntentType = "TRANSITION_EFFECT",
            .Content = "你是PowerPoint切换效果助手。支持应用切换效果（fade、push、wipe、split等）到当前幻灯片或所有幻灯片。"
        })

        pptApp.IntentPrompts.Add(New IntentPromptTemplate With {
            .IntentType = "GENERAL_QUERY",
            .Content = "你是PowerPoint助手。如果用户需求明确且可以执行，返回JSON代码片段；如果用户需求不明确，请先询问用户澄清；对于简单问候或问答，直接用中文回复。"
        })

        ' 功能模式提示词
        pptApp.FunctionModePrompts.Add(New FunctionModePromptTemplate With {
            .Mode = "continuation",
            .Content = "你是一个专业的演示文稿写作助手。根据提供的幻灯片上下文，自然地续写内容。

要求：
1. 保持与原有幻灯片一致的风格和术语
2. 内容要简洁有力，适合演示展示
3. 只输出续写内容，不要添加任何解释
4. 每页内容控制在合理的篇幅内"
        })

        pptApp.FunctionModePrompts.Add(New FunctionModePromptTemplate With {
            .Mode = "template_render",
            .Content = "你是一个专业的演示文稿内容生成助手。根据提供的PPT模板结构生成新的内容。

【PPT模板结构说明】
- slides: 幻灯片数组，每个包含layout(布局)和elements(元素列表)
- elements包含: type(类型)、text(文本)、formatting(格式)

【内容生成要求】
1. 按照模板的幻灯片数量和布局生成内容
2. 标题要简洁有力，正文要点要清晰
3. 内容适合演示场景，避免过长的段落"
        })

        ' JSON Schema约束
        pptApp.JsonSchemaConstraint = "
【PowerPoint JSON输出格式规范 - 必须严格遵守】

【重要】JSON必须使用Markdown代码块格式返回，例如：
```json
{""command"": ""InsertSlide"", ""params"": {...}}
```
禁止直接返回裸JSON文本！

你必须且只能返回以下两种格式之一：

单JSON代码格式：
```json
{""command"": ""InsertSlide"", ""params"": {""title"": ""标题"", ""content"": ""内容""}}
```

多JSON代码格式：
```json
{""commands"": [{""command"": ""CreateSlides"", ""params"": {...}}, {...}]}
```

【PowerPoint command类型】
1. InsertSlide - 插入单页幻灯片 (title, content, layout)
2. CreateSlides - 批量创建多页幻灯片 (slides数组，每项含title/content/layout) 【推荐用于多页】
3. InsertText - 插入文本 (content, slideIndex)
4. InsertShape - 插入形状 (shapeType, text)
5. FormatSlide - 格式化幻灯片 (slideIndex, background, theme)
6. InsertTable - 插入表格 (rows, cols, data)
7. AddAnimation - 添加动画 (effect: fadeIn/flyIn/zoom/wipe, targetShapes: all/title)
8. ApplyTransition - 应用切换效果 (transitionType: fade/push/wipe/split, scope: all/current)
9. BeautifySlides - 美化幻灯片 (theme, colorScheme)

【slideIndex说明】
- -1 或不填表示当前幻灯片
- 0 表示第一张幻灯片

【绝对禁止】
- 禁止使用Excel命令(WriteData, ApplyFormula等)
- 禁止使用Word命令(InsertText的Word版本、GenerateTOC等)
- 禁止返回上方未指定的command类型
- 禁止返回不带代码块的裸JSON"

        Return pptApp
    End Function
End Class
