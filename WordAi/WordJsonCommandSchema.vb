' WordAi\WordJsonCommandSchema.vb
' Word JSON命令Schema定义和校验

Imports System.Diagnostics
Imports Newtonsoft.Json.Linq

''' <summary>
''' Word JSON命令Schema和校验器
''' </summary>
Public Class WordJsonCommandSchema

    ''' <summary>
    ''' 仅列出 ExecuteWordCommand 已实现的命令，避免向 Agent 暴露无执行器的假能力。
    ''' </summary>
    Public Shared ReadOnly SupportedCommands As String() = {
        "InsertText",
        "FormatText",
        "ReplaceText",
        "DeleteText",
        "ApplyStyle",
        "SetParagraphFormat",
        "InsertTable",
        "GenerateTOC",
        "BeautifyDocument"
    }

    ''' <summary>
    ''' 只读命令：注册表（ShareRibbon/Tools/word/*.json）已经把它们发给 Planner，执行器也确实实现了。
    ''' 之前它们不在白名单里，于是 Agent 规划出 ListParagraphs 后被宿主判成"不支持的命令"，
    ''' 还按 JSON 格式错误自动修复了 3 次——格式从来不是问题，两套能力清单不一致才是。
    ''' </summary>
    Public Shared ReadOnly SupportedReadCommands As String() = {
        "ListParagraphs",
        "GetParagraphInfo"
    }

    Public Shared Function IsReadCommand(command As String) As Boolean
        Return SupportedReadCommands.Any(Function(c) c.Equals(command, StringComparison.OrdinalIgnoreCase))
    End Function

    Public Shared Function IsSupportedCommand(command As String) As Boolean
        If String.IsNullOrEmpty(command) Then Return False
        If SupportedCommands.Any(Function(c) c.Equals(command, StringComparison.OrdinalIgnoreCase)) Then Return True
        Return IsReadCommand(command)
    End Function

    ''' <summary>
    ''' 获取严格的JSON Schema定义（用于约束大模型输出）
    ''' </summary>
    Public Shared Function GetStrictJsonSchemaPrompt() As String
        Return "
【重要】你必须且只能返回以下两种JSON格式之一：

格式1 - 单个命令：
```json
{
  ""command"": ""InsertText"",
  ""params"": {
    ""position"": ""cursor"",
    ""content"": ""要插入的文本内容""
  }
}
```

格式2 - 多个命令（批量操作）：
```json
{
  ""commands"": [
    {""command"": ""InsertText"", ""params"": {""position"": ""cursor"", ""content"": ""标题""}},
    {""command"": ""FormatText"", ""params"": {""range"": ""selection"", ""bold"": true, ""fontSize"": 16}}
  ]
}
```

【绝对禁止的格式】
- 禁止 {""command"": ""xxx"", ""actions"": [...]}
- 禁止 {""command"": ""xxx"", ""content"": ""...""} (缺少params包装)
- 禁止 {""operations"": [...]}

【Word支持的11个命令】

=== 基础文本操作 (4个) ===
1. InsertText - 插入文本 {content:必需, position:cursor/start/end}
2. FormatText - 格式化 {range:selection/all, bold/italic/fontSize/fontName/underline/color}
3. ReplaceText - 查找替换 {find:必需, replace:必需, matchCase:可选, matchWholeWord:可选}
4. DeleteText - 删除文本 {range:selection/all}

=== 段落和样式 (2个) ===
5. ApplyStyle - 应用样式 {styleName:必需如""标题 1""/""正文"", range:selection/paragraph}
6. SetParagraphFormat - 段落格式 {paragraphIndex:必需从1开始, alignment:left/center/right/justify, firstLineIndent:缩进字符数, beforeSpacing/afterSpacing:磅}

=== 表格操作 (1个) ===
7. InsertTable - 插入表格 {rows:必需, cols:必需, data:可选二维数组, style:可选}

=== 文档结构 (1个) ===
8. GenerateTOC - 生成目录 {position:start/cursor, levels:1-9, includePageNumbers:默认true}

=== 文档美化 (1个) ===
9. BeautifyDocument - 美化文档 {theme:{h1/h2/body字体设置}, margins:{top/bottom/left/right}}

=== 只读命令 (2个，不改文档) ===
10. ListParagraphs - 列出段落 {maxCount:可选，默认50}，返回段落序号、文本摘要、样式和格式
11. GetParagraphInfo - 读取单段详情 {paragraphIndex:必需从1开始}

【重要决策规则】
1. 只能使用上述11个已实现命令处理用户需求；10、11只读，用来先看清文档再规划，不能作为最终交付
2. 翻译需求请告知用户使用工具栏的""翻译""按钮
3. 如果用户需求不明确，直接用中文询问
4. InsertText 的 position:end 只把文本追加到当前最后一段，不会新建段落；要得到多段结构必须在 content 里自带换行。SetParagraphFormat 的 paragraphIndex 指向文档当时的真实段落序号，同批次里靠前的命令会改变段落数，不要按""插了几段就是第几段""预先编号
5. paragraphIndex 必须是文档当前真实存在的序号：上下文里的文档结构会给出「共 N 段」，只能取 1..N。不确定就先用 ListParagraphs 读一次再规划。禁止 999、0、-1 这类占位或""指向最后一段""式的猜法；越界会直接失败并把实际段落数回给你"
    End Function

    ''' <summary>
    ''' 获取格式校验失败的重试提示（Self-check机制）
    ''' </summary>
    Public Shared Function GetFormatCorrectionPrompt(originalJson As String, errorMessage As String) As String
        Return $"你之前返回的JSON格式不符合规范:

【错误原因】{errorMessage}

【你返回的内容】
{originalJson}

【正确格式示例】
单命令:
{{""command"": ""InsertText"", ""params"": {{""position"": ""cursor"", ""content"": ""文本内容""}}}}

多命令:
{{""commands"": [{{""command"": ""InsertText"", ""params"": {{""content"": ""内容1""}}}}, {{""command"": ""FormatText"", ""params"": {{""range"": ""selection"", ""bold"": true}}}}]}}

请严格按照上述格式重新返回JSON命令。"
    End Function

    ''' <summary>
    ''' 验证整个JSON响应结构是否符合规范
    ''' </summary>
    Public Shared Function ValidateJsonStructure(jsonText As String, ByRef errorMessage As String, ByRef normalizedJson As JToken) As Boolean
        Try
            errorMessage = ""
            normalizedJson = Nothing

            Dim token = JToken.Parse(jsonText)
            If token.Type <> JTokenType.Object Then
                errorMessage = "响应必须是JSON对象"
                Return False
            End If

            Dim jsonObj = CType(token, JObject)

            ' 检查是否是 commands 数组格式
            If jsonObj("commands") IsNot Nothing Then
                If jsonObj("commands").Type <> JTokenType.Array Then
                    errorMessage = "commands必须是数组"
                    Return False
                End If

                ' 验证数组中的每个命令
                Dim commands = CType(jsonObj("commands"), JArray)
                For i As Integer = 0 To commands.Count - 1
                    Dim cmd = commands(i)
                    If cmd.Type <> JTokenType.Object Then
                        errorMessage = $"commands[{i}]必须是对象"
                        Return False
                    End If
                    
                    ' 标准化并验证每个命令
                    Dim cmdObj = CType(cmd, JObject)
                    cmdObj = NormalizeCommandStructure(cmdObj)
                    commands(i) = cmdObj
                    
                    Dim cmdError As String = ""
                    If Not ValidateCommand(cmdObj, cmdError) Then
                        errorMessage = $"commands[{i}]: {cmdError}"
                        Return False
                    End If
                Next
                
                normalizedJson = jsonObj
                Return True
            End If

            ' 检查是否有禁止的格式
            If jsonObj("actions") IsNot Nothing Then
                errorMessage = "禁止使用actions格式，请使用commands数组"
                Return False
            End If

            If jsonObj("operations") IsNot Nothing Then
                errorMessage = "禁止使用operations格式，请使用commands数组"
                Return False
            End If

            ' 单命令格式
            If jsonObj("command") IsNot Nothing Then
                jsonObj = NormalizeCommandStructure(jsonObj)
                Dim cmdError As String = ""
                If Not ValidateCommand(jsonObj, cmdError) Then
                    errorMessage = cmdError
                    Return False
                End If
                normalizedJson = jsonObj
                Return True
            End If

            errorMessage = "缺少command或commands字段"
            Return False

        Catch ex As Newtonsoft.Json.JsonReaderException
            errorMessage = $"JSON解析失败: {ex.Message}"
            Return False
        Catch ex As Exception
            errorMessage = $"验证异常: {ex.Message}"
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 标准化JSON命令结构 - 将扁平结构自动包装到params中
    ''' </summary>
    Public Shared Function NormalizeCommandStructure(json As JObject) As JObject
        Try
            ' 检查是否已有params字段
            If json("params") IsNot Nothing Then
                Return json
            End If

            ' 检查是否有command字段
            Dim command = json("command")?.ToString()
            If String.IsNullOrEmpty(command) Then
                Return json
            End If

            ' 需要移到params中的字段
            Dim topLevelFields As String() = {"command"}
            Dim paramsFields As New JObject()

            For Each prop In json.Properties().ToList()
                If Not topLevelFields.Contains(prop.Name, StringComparer.OrdinalIgnoreCase) Then
                    paramsFields(prop.Name) = prop.Value
                    json.Remove(prop.Name)
                End If
            Next

            If paramsFields.Count > 0 Then
                json("params") = paramsFields
            End If

            Return json
        Catch ex As Exception
            Debug.WriteLine($"NormalizeCommandStructure 出错: {ex.Message}")
            Return json
        End Try
    End Function

    ''' <summary>
    ''' 校验JSON命令是否有效
    ''' </summary>
    Public Shared Function ValidateCommand(json As JObject, ByRef errorMessage As String) As Boolean
        Try
            errorMessage = ""
            
            ' 首先进行结构标准化
            json = NormalizeCommandStructure(json)
            
            Dim command = json("command")?.ToString()
            If String.IsNullOrEmpty(command) Then
                errorMessage = "缺少command字段"
                Return False
            End If
            
            If Not IsSupportedCommand(command) Then
                errorMessage = $"不支持的命令: {command}。支持的命令: {String.Join(", ", SupportedCommands.Concat(SupportedReadCommands))}"
                Return False
            End If

            Dim params = json("params")
            If params Is Nothing Then
                ' 只读命令允许不带 params（ListParagraphs 走默认值），写命令仍然必须有参包装
                If IsReadCommand(command) Then
                    json("params") = New JObject()
                Else
                    errorMessage = "缺少params字段"
                    Return False
                End If
            End If
            
            ' 根据命令类型校验参数
            Select Case command.ToLower()
                ' === 基础文本操作 ===
                Case "inserttext"
                    Return ValidateInsertText(params, errorMessage)
                Case "formattext"
                    Return ValidateFormatText(params, errorMessage)
                Case "replacetext"
                    Return ValidateReplaceText(params, errorMessage)
                Case "deletetext"
                    Return ValidateDeleteText(params, errorMessage)
                ' === 段落和样式 ===
                Case "applystyle"
                    Return ValidateApplyStyle(params, errorMessage)
                Case "setparagraphformat"
                    Return ValidateSetParagraphFormat(params, errorMessage)
                ' === 表格操作 ===
                Case "inserttable"
                    Return ValidateInsertTable(params, errorMessage)
                ' === 文档结构 ===
                Case "generatetoc"
                    Return ValidateGenerateTOC(params, errorMessage)
                ' === 文档美化 ===
                Case "beautifydocument"
                    Return ValidateBeautifyDocument(params, errorMessage)
                Case Else
                    Return True
            End Select
            
        Catch ex As Exception
            errorMessage = $"JSON校验异常: {ex.Message}"
            Return False
        End Try
    End Function

    Private Shared Function ValidateInsertText(params As JToken, ByRef errorMessage As String) As Boolean
        Dim content = params("content")?.ToString()
        ' 兼容处理：大模型在GENERAL_QUERY模式下可能返回text而不是content
        If String.IsNullOrEmpty(content) Then
            content = params("text")?.ToString()
            ' 如果text存在，将其复制到content字段以便后续执行
            If Not String.IsNullOrEmpty(content) AndAlso params.Type = JTokenType.Object Then
                CType(params, JObject)("content") = content
            End If
        End If
        If String.IsNullOrEmpty(content) Then
            errorMessage = "InsertText缺少content参数"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateFormatText(params As JToken, ByRef errorMessage As String) As Boolean
        ' FormatText至少需要一个格式化属性
        If params("bold") Is Nothing AndAlso params("italic") Is Nothing AndAlso 
           params("fontSize") Is Nothing AndAlso params("fontName") Is Nothing AndAlso
           params("underline") Is Nothing Then
            errorMessage = "FormatText至少需要一个格式化属性(bold/italic/fontSize/fontName/underline)"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateReplaceText(params As JToken, ByRef errorMessage As String) As Boolean
        Dim find = params("find")?.ToString()
        If String.IsNullOrEmpty(find) Then
            errorMessage = "ReplaceText缺少find参数"
            Return False
        End If
        
        If params("replace") Is Nothing Then
            errorMessage = "ReplaceText缺少replace参数"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateInsertTable(params As JToken, ByRef errorMessage As String) As Boolean
        Dim rows = params("rows")
        Dim cols = params("cols")
        
        If rows Is Nothing OrElse cols Is Nothing Then
            errorMessage = "InsertTable缺少rows或cols参数"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateApplyStyle(params As JToken, ByRef errorMessage As String) As Boolean
        Dim styleName = params("styleName")?.ToString()
        If String.IsNullOrEmpty(styleName) Then
            errorMessage = "ApplyStyle缺少styleName参数"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateGenerateTOC(params As JToken, ByRef errorMessage As String) As Boolean
        ' GenerateTOC参数都是可选的
        Dim levels = params("levels")
        If levels IsNot Nothing Then
            Dim levelValue = levels.Value(Of Integer)()
            If levelValue < 1 OrElse levelValue > 9 Then
                errorMessage = "GenerateTOC的levels参数必须在1-9之间"
                Return False
            End If
        End If
        Return True
    End Function

    Private Shared Function ValidateBeautifyDocument(params As JToken, ByRef errorMessage As String) As Boolean
        ' BeautifyDocument至少需要theme或margins之一
        If params("theme") Is Nothing AndAlso params("margins") Is Nothing Then
            errorMessage = "BeautifyDocument至少需要theme或margins参数"
            Return False
        End If
        Return True
    End Function

#Region "新增命令验证方法"

    Private Shared Function ValidateDeleteText(params As JToken, ByRef errorMessage As String) As Boolean
        ' DeleteText可以用range指定范围,默认删除选中内容
        Return True
    End Function

    Private Shared Function ValidateSetParagraphFormat(params As JToken, ByRef errorMessage As String) As Boolean
        ' 执行器按段落序号定位，缺 paragraphIndex 只会退化成布尔失败，模型拿不到可修的原因
        Dim paragraphIndex = params("paragraphIndex")
        If paragraphIndex Is Nothing OrElse paragraphIndex.Type = JTokenType.Null OrElse
           paragraphIndex.Value(Of Integer)() < 1 Then
            errorMessage = "SetParagraphFormat需要paragraphIndex(整数,从1开始,指向文档当前真实段落序号)"
            Return False
        End If

        ' 至少需要一个段落格式属性
        If params("alignment") Is Nothing AndAlso params("firstLineIndent") Is Nothing AndAlso
           params("beforeSpacing") Is Nothing AndAlso params("afterSpacing") Is Nothing AndAlso
           params("fontSize") Is Nothing AndAlso params("fontName") Is Nothing AndAlso
           params("bold") Is Nothing AndAlso params("italic") Is Nothing AndAlso
           params("color") Is Nothing Then
            errorMessage = "SetParagraphFormat至少需要一个格式属性"
            Return False
        End If
        Return True
    End Function

#End Region

End Class
