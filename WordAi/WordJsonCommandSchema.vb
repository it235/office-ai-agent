' WordAi\WordJsonCommandSchema.vb
' Word JSON命令Schema定义和校验

Imports System.Diagnostics
Imports System.Text.RegularExpressions
Imports Newtonsoft.Json.Linq
Imports Newtonsoft.Json.Schema

''' <summary>
''' Word JSON命令Schema和校验器
''' </summary>
Public Class WordJsonCommandSchema

    ''' <summary>
    ''' 支持的命令类型 (22个命令覆盖主流Word操作场景)
    ''' 基础文本操作: InsertText, FormatText, ReplaceText, DeleteText, CopyPasteText
    ''' 段落和样式: ApplyStyle, SetParagraphFormat, InsertParagraph, SetLineSpacing, SetIndent
    ''' 表格操作: InsertTable, FormatTable, InsertTableRow, DeleteTableRow
    ''' 文档结构: GenerateTOC, InsertHeader, InsertFooter, InsertPageNumber
    ''' 文档美化: BeautifyDocument, SetPageMargins
    ''' 高级功能: InsertImage
    ''' VBA回退: ExecuteVBA
    ''' </summary>
    Public Shared ReadOnly SupportedCommands As String() = {
        "InsertText",
        "FormatText",
        "ReplaceText",
        "DeleteText",
        "CopyPasteText",
        "ApplyStyle",
        "SetParagraphFormat",
        "InsertParagraph",
        "SetLineSpacing",
        "SetIndent",
        "InsertTable",
        "FormatTable",
        "InsertTableRow",
        "DeleteTableRow",
        "GenerateTOC",
        "InsertHeader",
        "InsertFooter",
        "InsertPageNumber",
        "BeautifyDocument",
        "SetPageMargins",
        "InsertImage",
        "ExecuteVBA"
    }

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

【Word支持的22个命令】

=== 基础文本操作 (5个) ===
1. InsertText - 插入文本 {content:必需, position:cursor/start/end}
2. FormatText - 格式化 {range:selection/all, bold/italic/fontSize/fontName/underline/color}
3. ReplaceText - 查找替换 {find:必需, replace:必需, matchCase:可选, matchWholeWord:可选}
4. DeleteText - 删除文本 {range:selection/all}
5. CopyPasteText - 复制粘贴 {sourceRange:必需, targetPosition:cursor/start/end}

=== 段落和样式 (5个) ===
6. ApplyStyle - 应用样式 {styleName:必需如""标题 1""/""正文"", range:selection/paragraph}
7. SetParagraphFormat - 段落格式 {alignment:left/center/right/justify, firstLineIndent:可选, beforeSpacing/afterSpacing:可选}
8. InsertParagraph - 插入段落 {count:默认1, pageBreak:true则插入分页符}
9. SetLineSpacing - 行距 {spacing:1/1.5/2或具体值, range:selection/all}
10. SetIndent - 缩进 {left:左缩进cm, right:右缩进cm, firstLine:首行缩进cm, range:selection/paragraph}

=== 表格操作 (4个) ===
11. InsertTable - 插入表格 {rows:必需, cols:必需, data:可选二维数组, style:可选}
12. FormatTable - 格式化表格 {tableIndex:表格索引从1开始, style:可选, borders:可选, headerRow:可选}
13. InsertTableRow - 插入行 {tableIndex:必需, position:after/before, rowIndex:可选}
14. DeleteTableRow - 删除行 {tableIndex:必需, rowIndex:必需}

=== 文档结构 (4个) ===
15. GenerateTOC - 生成目录 {position:start/cursor, levels:1-9, includePageNumbers:默认true}
16. InsertHeader - 页眉 {content:必需, alignment:left/center/right}
17. InsertFooter - 页脚 {content:必需, alignment:left/center/right}
18. InsertPageNumber - 页码 {position:header/footer, alignment:left/center/right, format:可选}

=== 文档美化 (2个) ===
19. BeautifyDocument - 美化文档 {theme:{h1/h2/body字体设置}, margins:{top/bottom/left/right}}
20. SetPageMargins - 页边距 {top:cm, bottom:cm, left:cm, right:cm}

=== 高级功能 (1个) ===
21. InsertImage - 插入图片 {imagePath:必需, width:可选, height:可选, position:cursor/start/end}

=== VBA回退 (1个) ===
22. ExecuteVBA - 执行VBA代码 {code:必需,完整的Sub或Function代码}
    当以上命令无法满足需求时,生成VBA代码作为回退方案

【重要决策规则】
1. 优先使用上述22个命令处理用户需求
2. 复杂需求无法用命令实现时，使用ExecuteVBA生成VBA代码
3. 翻译需求请告知用户使用工具栏的""翻译""按钮
4. 如果用户需求不明确，直接用中文询问"
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
            
            If Not SupportedCommands.Any(Function(c) c.Equals(command, StringComparison.OrdinalIgnoreCase)) Then
                errorMessage = $"不支持的命令: {command}。支持的命令: {String.Join(", ", SupportedCommands)}"
                Return False
            End If
            
            Dim params = json("params")
            If params Is Nothing Then
                errorMessage = "缺少params字段"
                Return False
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
                Case "copypastetext"
                    Return ValidateCopyPasteText(params, errorMessage)
                ' === 段落和样式 ===
                Case "applystyle"
                    Return ValidateApplyStyle(params, errorMessage)
                Case "setparagraphformat"
                    Return ValidateSetParagraphFormat(params, errorMessage)
                Case "insertparagraph"
                    Return ValidateInsertParagraph(params, errorMessage)
                Case "setlinespacing"
                    Return ValidateSetLineSpacing(params, errorMessage)
                Case "setindent"
                    Return ValidateSetIndent(params, errorMessage)
                ' === 表格操作 ===
                Case "inserttable"
                    Return ValidateInsertTable(params, errorMessage)
                Case "formattable"
                    Return ValidateFormatTable(params, errorMessage)
                Case "inserttablerow"
                    Return ValidateInsertTableRow(params, errorMessage)
                Case "deletetablerow"
                    Return ValidateDeleteTableRow(params, errorMessage)
                ' === 文档结构 ===
                Case "generatetoc"
                    Return ValidateGenerateTOC(params, errorMessage)
                Case "insertheader"
                    Return ValidateInsertHeader(params, errorMessage)
                Case "insertfooter"
                    Return ValidateInsertFooter(params, errorMessage)
                Case "insertpagenumber"
                    Return ValidateInsertPageNumber(params, errorMessage)
                ' === 文档美化 ===
                Case "beautifydocument"
                    Return ValidateBeautifyDocument(params, errorMessage)
                Case "setpagemargins"
                    Return ValidateSetPageMargins(params, errorMessage)
                ' === 高级功能 ===
                Case "insertimage"
                    Return ValidateInsertImage(params, errorMessage)
                ' === VBA回退 ===
                Case "executevba"
                    Return ValidateExecuteVBA(params, errorMessage)
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

    Private Shared Function ValidateCopyPasteText(params As JToken, ByRef errorMessage As String) As Boolean
        Dim sourceRange = params("sourceRange")?.ToString()
        If String.IsNullOrEmpty(sourceRange) Then
            errorMessage = "CopyPasteText缺少sourceRange参数"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateSetParagraphFormat(params As JToken, ByRef errorMessage As String) As Boolean
        ' 至少需要一个段落格式属性
        If params("alignment") Is Nothing AndAlso params("firstLineIndent") Is Nothing AndAlso
           params("beforeSpacing") Is Nothing AndAlso params("afterSpacing") Is Nothing Then
            errorMessage = "SetParagraphFormat至少需要一个格式属性"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateInsertParagraph(params As JToken, ByRef errorMessage As String) As Boolean
        ' 所有参数都是可选的
        Return True
    End Function

    Private Shared Function ValidateSetLineSpacing(params As JToken, ByRef errorMessage As String) As Boolean
        If params("spacing") Is Nothing Then
            errorMessage = "SetLineSpacing缺少spacing参数"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateSetIndent(params As JToken, ByRef errorMessage As String) As Boolean
        ' 至少需要一个缩进属性
        If params("left") Is Nothing AndAlso params("right") Is Nothing AndAlso params("firstLine") Is Nothing Then
            errorMessage = "SetIndent至少需要一个缩进属性(left/right/firstLine)"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateFormatTable(params As JToken, ByRef errorMessage As String) As Boolean
        Dim tableIndex = params("tableIndex")
        If tableIndex Is Nothing Then
            errorMessage = "FormatTable缺少tableIndex参数"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateInsertTableRow(params As JToken, ByRef errorMessage As String) As Boolean
        Dim tableIndex = params("tableIndex")
        If tableIndex Is Nothing Then
            errorMessage = "InsertTableRow缺少tableIndex参数"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateDeleteTableRow(params As JToken, ByRef errorMessage As String) As Boolean
        If params("tableIndex") Is Nothing Then
            errorMessage = "DeleteTableRow缺少tableIndex参数"
            Return False
        End If
        If params("rowIndex") Is Nothing Then
            errorMessage = "DeleteTableRow缺少rowIndex参数"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateInsertHeader(params As JToken, ByRef errorMessage As String) As Boolean
        Dim content = params("content")?.ToString()
        If String.IsNullOrEmpty(content) Then
            errorMessage = "InsertHeader缺少content参数"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateInsertFooter(params As JToken, ByRef errorMessage As String) As Boolean
        Dim content = params("content")?.ToString()
        If String.IsNullOrEmpty(content) Then
            errorMessage = "InsertFooter缺少content参数"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateInsertPageNumber(params As JToken, ByRef errorMessage As String) As Boolean
        ' 所有参数都是可选的
        Return True
    End Function

    Private Shared Function ValidateSetPageMargins(params As JToken, ByRef errorMessage As String) As Boolean
        ' 至少需要一个边距属性
        If params("top") Is Nothing AndAlso params("bottom") Is Nothing AndAlso
           params("left") Is Nothing AndAlso params("right") Is Nothing Then
            errorMessage = "SetPageMargins至少需要一个边距属性"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateInsertImage(params As JToken, ByRef errorMessage As String) As Boolean
        Dim imagePath = params("imagePath")?.ToString()
        If String.IsNullOrEmpty(imagePath) Then
            errorMessage = "InsertImage缺少imagePath参数"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateExecuteVBA(params As JToken, ByRef errorMessage As String) As Boolean
        Dim code = params("code")?.ToString()
        If String.IsNullOrEmpty(code) Then
            errorMessage = "ExecuteVBA缺少code参数"
            Return False
        End If
        
        ' 基本的VBA代码验证
        If Not code.ToLower().Contains("sub") AndAlso Not code.ToLower().Contains("function") Then
            errorMessage = "ExecuteVBA的code必须包含Sub或Function定义"
            Return False
        End If
        
        Return True
    End Function

#End Region

End Class
