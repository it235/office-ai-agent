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
    ''' 支持的命令类型
    ''' </summary>
    Public Shared ReadOnly SupportedCommands As String() = {
        "InsertText",
        "FormatText",
        "ReplaceText",
        "InsertTable",
        "ApplyStyle"
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
    {
      ""command"": ""InsertText"",
      ""params"": {
        ""position"": ""cursor"",
        ""content"": ""第一段内容""
      }
    },
    {
      ""command"": ""FormatText"",
      ""params"": {
        ""range"": ""selection"",
        ""bold"": true,
        ""fontSize"": 14
      }
    }
  ]
}
```

【绝对禁止的格式】
- 禁止 {""command"": ""xxx"", ""actions"": [...]}
- 禁止 {""command"": ""xxx"", ""content"": ""...""} (缺少params包装)
- 禁止 {""operations"": [...]}
- 禁止任何其他自创格式

【command类型限制】
只能使用: InsertText, FormatText, ReplaceText, InsertTable, ApplyStyle

【params必须包含的字段】
- InsertText: content(必需), position(可选: cursor/start/end)
- FormatText: range(必需: selection/all), bold/italic/fontSize/fontName(可选)
- ReplaceText: find(必需), replace(必需), matchCase(可选)
- InsertTable: rows(必需), cols(必需), data(可选)
- ApplyStyle: styleName(必需), range(可选: selection/paragraph)

如果用户需求不明确，请直接用中文询问用户，不要返回JSON。"
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
                Case "inserttext"
                    Return ValidateInsertText(params, errorMessage)
                Case "formattext"
                    Return ValidateFormatText(params, errorMessage)
                Case "replacetext"
                    Return ValidateReplaceText(params, errorMessage)
                Case "inserttable"
                    Return ValidateInsertTable(params, errorMessage)
                Case "applystyle"
                    Return ValidateApplyStyle(params, errorMessage)
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

End Class
