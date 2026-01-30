' PowerPointAi\PowerPointJsonCommandSchema.vb
' PowerPoint JSON命令Schema定义和校验

Imports System.Diagnostics
Imports System.Text.RegularExpressions
Imports Newtonsoft.Json.Linq
Imports Newtonsoft.Json.Schema

''' <summary>
''' PowerPoint JSON命令Schema和校验器
''' </summary>
Public Class PowerPointJsonCommandSchema

    ''' <summary>
    ''' 支持的命令类型
    ''' </summary>
    Public Shared ReadOnly SupportedCommands As String() = {
        "InsertSlide",
        "InsertText",
        "InsertShape",
        "FormatSlide",
        "InsertTable",
        "CreateSlides",
        "AddAnimation",
        "ApplyTransition",
        "BeautifySlides"
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
  ""command"": ""InsertSlide"",
  ""params"": {
    ""position"": ""end"",
    ""layout"": ""title"",
    ""title"": ""幻灯片标题""
  }
}
```

格式2 - 多个命令（批量操作）：
```json
{
  ""commands"": [
    {
      ""command"": ""InsertSlide"",
      ""params"": {
        ""position"": ""end"",
        ""title"": ""第一页标题""
      }
    },
    {
      ""command"": ""InsertText"",
      ""params"": {
        ""slideIndex"": -1,
        ""content"": ""文本内容"",
        ""x"": 100,
        ""y"": 200
      }
    }
  ]
}
```

【绝对禁止的格式】
- 禁止 {""command"": ""xxx"", ""actions"": [...]}
- 禁止 {""command"": ""xxx"", ""title"": ""...""} (缺少params包装)
- 禁止 {""operations"": [...]}
- 禁止任何其他自创格式

【command类型限制】
只能使用: InsertSlide, InsertText, InsertShape, FormatSlide, InsertTable, CreateSlides, AddAnimation, ApplyTransition, BeautifySlides

【params必须包含的字段】
- InsertSlide: position(可选: current/end), layout(可选), title(可选), content(可选)
- InsertText: content(必需), slideIndex(可选: -1表示当前), x/y(可选)
- InsertShape: shapeType(必需), x(必需), y(必需), width(可选), height(可选)
- FormatSlide: slideIndex(可选), background(可选), transition(可选)
- InsertTable: rows(必需), cols(必需), data(可选), slideIndex(可选)
- CreateSlides: slides(必需，数组，每个元素含title/content/layout)
- AddAnimation: effect(必需: fadeIn/flyIn/zoom/wipe), slideIndex(可选), targetShapes(可选: all/title)
- ApplyTransition: transitionType(必需: fade/push/wipe/split), scope(可选: all/current), duration(可选)
- BeautifySlides: scope(可选: all/current), theme(可选，含background/titleFont/bodyFont)

【slideIndex说明】
- -1 或不填表示当前幻灯片
- 0 表示第一张幻灯片
- 正数表示具体幻灯片索引

【批量生成幻灯片示例】
```json
{
  ""command"": ""CreateSlides"",
  ""params"": {
    ""slides"": [
      {""title"": ""第一章 概述"", ""content"": ""这是第一页内容"", ""layout"": ""titleAndContent""},
      {""title"": ""第二章 详情"", ""content"": ""这是第二页内容"", ""layout"": ""titleAndContent""}
    ]
  }
}
```

【添加动画示例】
```json
{
  ""command"": ""AddAnimation"",
  ""params"": {
    ""slideIndex"": -1,
    ""effect"": ""fadeIn"",
    ""targetShapes"": ""all""
  }
}
```

【幻灯片切换效果示例】
```json
{
  ""command"": ""ApplyTransition"",
  ""params"": {
    ""scope"": ""all"",
    ""transitionType"": ""fade"",
    ""duration"": 1.0
  }
}
```

【幻灯片美化示例】
```json
{
  ""command"": ""BeautifySlides"",
  ""params"": {
    ""scope"": ""all"",
    ""theme"": {
      ""background"": ""#F5F5F5"",
      ""titleFont"": {""name"": ""微软雅黑"", ""size"": 28, ""color"": ""#333333""},
      ""bodyFont"": {""name"": ""微软雅黑"", ""size"": 18, ""color"": ""#666666""}
    }
  }
}
```

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
{{""command"": ""InsertSlide"", ""params"": {{""position"": ""end"", ""title"": ""标题""}}}}

多命令:
{{""commands"": [{{""command"": ""InsertSlide"", ""params"": {{""title"": ""第一页""}}}}, {{""command"": ""InsertText"", ""params"": {{""content"": ""内容""}}}}]}}

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
                Case "insertslide"
                    Return ValidateInsertSlide(params, errorMessage)
                Case "inserttext"
                    Return ValidateInsertText(params, errorMessage)
                Case "insertshape"
                    Return ValidateInsertShape(params, errorMessage)
                Case "formatslide"
                    Return ValidateFormatSlide(params, errorMessage)
                Case "inserttable"
                    Return ValidateInsertTable(params, errorMessage)
                Case "createslides"
                    Return ValidateCreateSlides(params, errorMessage)
                Case "addanimation"
                    Return ValidateAddAnimation(params, errorMessage)
                Case "applytransition"
                    Return ValidateApplyTransition(params, errorMessage)
                Case "beautifyslides"
                    Return ValidateBeautifySlides(params, errorMessage)
                Case Else
                    Return True
            End Select
            
        Catch ex As Exception
            errorMessage = $"JSON校验异常: {ex.Message}"
            Return False
        End Try
    End Function

    Private Shared Function ValidateInsertSlide(params As JToken, ByRef errorMessage As String) As Boolean
        ' InsertSlide参数都是可选的，基本验证通过
        Return True
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

    Private Shared Function ValidateInsertShape(params As JToken, ByRef errorMessage As String) As Boolean
        Dim shapeType = params("shapeType")?.ToString()
        If String.IsNullOrEmpty(shapeType) Then
            errorMessage = "InsertShape缺少shapeType参数"
            Return False
        End If
        
        If params("x") Is Nothing OrElse params("y") Is Nothing Then
            errorMessage = "InsertShape缺少x或y参数"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateFormatSlide(params As JToken, ByRef errorMessage As String) As Boolean
        ' FormatSlide至少需要一个格式化属性
        If params("background") Is Nothing AndAlso params("transition") Is Nothing AndAlso
           params("layout") Is Nothing Then
            errorMessage = "FormatSlide至少需要一个格式化属性(background/transition/layout)"
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

    Private Shared Function ValidateCreateSlides(params As JToken, ByRef errorMessage As String) As Boolean
        Dim slides = params("slides")
        If slides Is Nothing OrElse slides.Type <> JTokenType.Array Then
            errorMessage = "CreateSlides缺少slides数组参数"
            Return False
        End If
        
        Dim slidesArray = CType(slides, JArray)
        If slidesArray.Count = 0 Then
            errorMessage = "CreateSlides的slides数组不能为空"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateAddAnimation(params As JToken, ByRef errorMessage As String) As Boolean
        Dim effect = params("effect")?.ToString()
        If String.IsNullOrEmpty(effect) Then
            errorMessage = "AddAnimation缺少effect参数"
            Return False
        End If
        
        Dim validEffects = {"fadein", "flyin", "zoom", "wipe", "appear", "float"}
        If Not validEffects.Contains(effect.ToLower()) Then
            errorMessage = $"AddAnimation的effect参数无效: {effect}。有效值: fadeIn, flyIn, zoom, wipe, appear, float"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateApplyTransition(params As JToken, ByRef errorMessage As String) As Boolean
        Dim transType = params("transitionType")?.ToString()
        If String.IsNullOrEmpty(transType) Then
            errorMessage = "ApplyTransition缺少transitionType参数"
            Return False
        End If
        
        Dim validTypes = {"fade", "push", "wipe", "split", "reveal", "random"}
        If Not validTypes.Contains(transType.ToLower()) Then
            errorMessage = $"ApplyTransition的transitionType参数无效: {transType}。有效值: fade, push, wipe, split, reveal, random"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateBeautifySlides(params As JToken, ByRef errorMessage As String) As Boolean
        ' BeautifySlides参数都是可选的，但至少需要一个
        If params("scope") Is Nothing AndAlso params("theme") Is Nothing Then
            errorMessage = "BeautifySlides至少需要scope或theme参数"
            Return False
        End If
        Return True
    End Function

End Class
