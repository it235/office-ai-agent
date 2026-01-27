' ExcelAi\ExcelJsonCommandSchema.vb
' Excel JSON命令Schema定义和校验

Imports System.Diagnostics
Imports System.Text.RegularExpressions
Imports Newtonsoft.Json.Linq
Imports Newtonsoft.Json.Schema

''' <summary>
''' Excel JSON命令Schema和校验器
''' </summary>
Public Class ExcelJsonCommandSchema

    ''' <summary>
    ''' 支持的命令类型
    ''' </summary>
    Public Shared ReadOnly SupportedCommands As String() = {
        "ApplyFormula",
        "WriteData",
        "FormatRange",
        "CreateChart",
        "CleanData"
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
  ""command"": ""ApplyFormula"",
  ""params"": {
    ""targetRange"": ""C1:C{lastRow}"",
    ""formula"": ""=A1+B1"",
    ""fillDown"": true
  }
}
```

格式2 - 多个命令（批量操作）：
```json
{
  ""commands"": [
    {
      ""command"": ""ApplyFormula"",
      ""params"": {
        ""targetRange"": ""C1:C{lastRow}"",
        ""formula"": ""=A1+B1""
      }
    },
    {
      ""command"": ""ApplyFormula"",
      ""params"": {
        ""targetRange"": ""E1:E{lastRow}"",
        ""formula"": ""=C1*D1""
      }
    }
  ]
}
```

【绝对禁止的格式】
- ❌ {""command"": ""xxx"", ""actions"": [...]}
- ❌ {""command"": ""xxx"", ""formula"": ""..."", ""range"": ""...""} (缺少params包装)
- ❌ {""operations"": [...]}
- ❌ 任何其他自创格式

【command类型限制】
只能使用: ApplyFormula, WriteData, FormatRange, CreateChart, CleanData

【params必须包含的字段】
- ApplyFormula: targetRange(必需), formula(必需), fillDown(可选)
- WriteData: targetRange(必需), data(必需)
- FormatRange: range(必需), style(可选)
- CreateChart: dataRange(必需), chartType(可选)
- CleanData: range(必需), operation(必需)

【动态范围占位符】
使用 {lastRow} 表示最后一行，不要用JS表达式

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
{{""command"": ""ApplyFormula"", ""params"": {{""targetRange"": ""C1:C{{lastRow}}"", ""formula"": ""=A1+B1""}}}}

多命令:
{{""commands"": [{{""command"": ""ApplyFormula"", ""params"": {{""targetRange"": ""C1"", ""formula"": ""=A1+B1""}}}}, {{""command"": ""ApplyFormula"", ""params"": {{""targetRange"": ""E1"", ""formula"": ""=C1*D1""}}}}]}}

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
    ''' 标准化JSON命令结构 - 将扁平结构自动包装到params中，并统一参数命名
    ''' 例如: {"command": "WriteData", "data": [...], "startCell": "A1"} 
    ''' 转换为: {"command": "WriteData", "params": {"data": [...], "startCell": "A1"}}
    ''' </summary>
    Public Shared Function NormalizeCommandStructure(json As JObject) As JObject
        Try
            ' 检查是否已有params字段
            If json("params") IsNot Nothing Then
                ' 即使已有params，也要标准化参数名
                NormalizeParamNames(json("params"), json("command")?.ToString())
                Return json
            End If

            ' 检查是否有command字段
            Dim command = json("command")?.ToString()
            If String.IsNullOrEmpty(command) Then
                Return json ' 无效结构，直接返回
            End If

            ' 需要移到params中的字段（排除command和workbook等顶级字段）
            Dim topLevelFields As String() = {"command", "workbook"}
            Dim paramsFields As New JObject()

            ' 遍历所有字段，将非顶级字段移动到params中
            For Each prop In json.Properties().ToList()
                If Not topLevelFields.Contains(prop.Name, StringComparer.OrdinalIgnoreCase) Then
                    paramsFields(prop.Name) = prop.Value
                    json.Remove(prop.Name)
                End If
            Next

            ' 只有当确实有字段需要移动时，才创建params
            If paramsFields.Count > 0 Then
                ' 标准化参数名
                NormalizeParamNames(paramsFields, command)
                json("params") = paramsFields
            End If

            Return json
        Catch ex As Exception
            Debug.WriteLine($"NormalizeCommandStructure 出错: {ex.Message}")
            Return json
        End Try
    End Function

    ''' <summary>
    ''' 标准化参数名 - 将各种别名统一为标准参数名
    ''' </summary>
    Private Shared Sub NormalizeParamNames(params As JToken, command As String)
        If params Is Nothing OrElse params.Type <> JTokenType.Object Then Return
        
        Dim paramsObj = CType(params, JObject)
        
        Select Case command?.ToLower()
            Case "applyformula"
                ' range -> targetRange
                If paramsObj("targetRange") Is Nothing AndAlso paramsObj("range") IsNot Nothing Then
                    paramsObj("targetRange") = paramsObj("range")
                    paramsObj.Remove("range")
                End If
                
            Case "writedata"
                ' startCell/range -> targetRange (如果targetRange不存在)
                If paramsObj("targetRange") Is Nothing Then
                    If paramsObj("startCell") IsNot Nothing Then
                        paramsObj("targetRange") = paramsObj("startCell")
                        paramsObj.Remove("startCell")
                    ElseIf paramsObj("range") IsNot Nothing Then
                        paramsObj("targetRange") = paramsObj("range")
                        paramsObj.Remove("range")
                    End If
                End If
                ' targetData -> data
                If paramsObj("data") Is Nothing AndAlso paramsObj("targetData") IsNot Nothing Then
                    paramsObj("data") = paramsObj("targetData")
                    paramsObj.Remove("targetData")
                End If
                
            Case "formatrange", "cleandata"
                ' targetRange -> range (这两个命令使用range)
                If paramsObj("range") Is Nothing AndAlso paramsObj("targetRange") IsNot Nothing Then
                    paramsObj("range") = paramsObj("targetRange")
                    paramsObj.Remove("targetRange")
                End If
        End Select
        
        ' 处理 targetSheet + targetRange 组合
        Dim targetSheet = paramsObj("targetSheet")?.ToString()
        Dim targetRange = paramsObj("targetRange")?.ToString()
        If Not String.IsNullOrEmpty(targetSheet) AndAlso Not String.IsNullOrEmpty(targetRange) Then
            If Not targetRange.Contains("!") Then
                paramsObj("targetRange") = $"{targetSheet}!{targetRange}"
            End If
        End If
    End Sub

    ''' <summary>
    ''' 校验JSON命令是否有效（自动进行结构标准化）
    ''' </summary>
    Public Shared Function ValidateCommand(json As JObject, ByRef errorMessage As String) As Boolean
        Try
            errorMessage = ""
            
            ' 首先进行结构标准化
            json = NormalizeCommandStructure(json)
            
            ' 检查command字段
            Dim command = json("command")?.ToString()
            If String.IsNullOrEmpty(command) Then
                errorMessage = "缺少command字段"
                Return False
            End If
            
            ' 检查是否是支持的命令
            If Not SupportedCommands.Any(Function(c) c.Equals(command, StringComparison.OrdinalIgnoreCase)) Then
                errorMessage = $"不支持的命令: {command}。支持的命令: {String.Join(", ", SupportedCommands)}"
                Return False
            End If
            
            ' 检查params字段
            Dim params = json("params")
            If params Is Nothing Then
                errorMessage = "缺少params字段"
                Return False
            End If
            
            ' 根据命令类型校验参数
            Select Case command.ToLower()
                Case "applyformula"
                    Return ValidateApplyFormula(params, errorMessage)
                Case "writedata"
                    Return ValidateWriteData(params, errorMessage)
                Case "formatrange"
                    Return ValidateFormatRange(params, errorMessage)
                Case "createchart"
                    Return ValidateCreateChart(params, errorMessage)
                Case "cleandata"
                    Return ValidateCleanData(params, errorMessage)
                Case Else
                    Return True
            End Select
            
        Catch ex As Exception
            errorMessage = $"JSON校验异常: {ex.Message}"
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 校验ApplyFormula命令参数
    ''' </summary>
    Private Shared Function ValidateApplyFormula(params As JToken, ByRef errorMessage As String) As Boolean
        ' 支持多种参数名：targetRange, range
        Dim targetRange = params("targetRange")?.ToString()
        If String.IsNullOrEmpty(targetRange) Then
            targetRange = params("range")?.ToString()
        End If
        
        Dim formula = params("formula")?.ToString()
        
        If String.IsNullOrEmpty(targetRange) Then
            errorMessage = "ApplyFormula缺少targetRange或range参数"
            Return False
        End If
        
        If String.IsNullOrEmpty(formula) Then
            errorMessage = "ApplyFormula缺少formula参数"
            Return False
        End If
        
        ' 校验范围格式 (支持占位符和Sheet!Range格式)
        If Not IsValidRangeFormat(targetRange) Then
            errorMessage = $"无效的范围格式: {targetRange}"
            Return False
        End If
        
        Return True
    End Function

    ''' <summary>
    ''' 校验WriteData命令参数
    ''' </summary>
    Private Shared Function ValidateWriteData(params As JToken, ByRef errorMessage As String) As Boolean
        ' 支持多种参数名：targetRange, startCell, range
        Dim targetRange = params("targetRange")?.ToString()
        If String.IsNullOrEmpty(targetRange) Then
            targetRange = params("startCell")?.ToString()
        End If
        If String.IsNullOrEmpty(targetRange) Then
            targetRange = params("range")?.ToString()
        End If
        
        ' 如果有targetSheet，组合成完整地址
        Dim targetSheet = params("targetSheet")?.ToString()
        If Not String.IsNullOrEmpty(targetSheet) AndAlso Not String.IsNullOrEmpty(targetRange) Then
            ' 如果targetRange不包含!，则添加工作表名
            If Not targetRange.Contains("!") Then
                targetRange = $"{targetSheet}!{targetRange}"
            End If
        End If
        
        Dim data = params("data")
        ' 支持data或targetData
        If data Is Nothing Then
            data = params("targetData")
        End If
        
        If String.IsNullOrEmpty(targetRange) Then
            errorMessage = "WriteData缺少targetRange或startCell参数"
            Return False
        End If
        
        If data Is Nothing Then
            errorMessage = "WriteData缺少data参数"
            Return False
        End If
        
        Return True
    End Function

    ''' <summary>
    ''' 校验FormatRange命令参数
    ''' </summary>
    Private Shared Function ValidateFormatRange(params As JToken, ByRef errorMessage As String) As Boolean
        Dim range = params("range")?.ToString()
        If String.IsNullOrEmpty(range) Then
            range = params("targetRange")?.ToString()
        End If
        
        If String.IsNullOrEmpty(range) Then
            errorMessage = "FormatRange缺少range参数"
            Return False
        End If
        
        Return True
    End Function

    ''' <summary>
    ''' 校验CreateChart命令参数
    ''' </summary>
    Private Shared Function ValidateCreateChart(params As JToken, ByRef errorMessage As String) As Boolean
        Dim dataRange = params("dataRange")?.ToString()
        
        If String.IsNullOrEmpty(dataRange) Then
            errorMessage = "CreateChart缺少dataRange参数"
            Return False
        End If
        
        Return True
    End Function

    ''' <summary>
    ''' 校验CleanData命令参数
    ''' </summary>
    Private Shared Function ValidateCleanData(params As JToken, ByRef errorMessage As String) As Boolean
        Dim range = params("range")?.ToString()
        Dim operation = params("operation")?.ToString()
        
        If String.IsNullOrEmpty(range) Then
            errorMessage = "CleanData缺少range参数"
            Return False
        End If
        
        Return True
    End Function

    ''' <summary>
    ''' 校验范围格式是否有效
    ''' </summary>
    Private Shared Function IsValidRangeFormat(range As String) As Boolean
        If String.IsNullOrEmpty(range) Then Return False
        
        ' 支持格式: 
        ' - 简单格式: A1, A1:B10, A:A, 1:1
        ' - 带工作表: Sheet1!A1, Sheet1!A1:B10
        ' - 占位符: A1:{lastRow}, {selection}
        ' - 中文工作表名: 汇总结果!A1
        
        ' 移除工作表前缀进行校验
        Dim rangeOnly = range
        If range.Contains("!") Then
            Dim parts = range.Split("!"c)
            If parts.Length = 2 Then
                rangeOnly = parts(1)
            End If
        End If
        
        ' 如果包含占位符，认为有效
        If rangeOnly.Contains("{") Then Return True
        
        ' 校验范围格式
        Dim pattern = "^([A-Za-z]+\d+|[A-Za-z]+|[0-9]+)(:[A-Za-z]*\d*)?$"
        Return Regex.IsMatch(rangeOnly, pattern)
    End Function

    ''' <summary>
    ''' 替换JSON中的占位符
    ''' </summary>
    Public Shared Function ReplacePlaceholders(json As JObject, context As Dictionary(Of String, String)) As JObject
        Dim jsonStr = json.ToString()
        
        For Each kvp In context
            jsonStr = jsonStr.Replace("{" & kvp.Key & "}", kvp.Value)
        Next
        
        Return JObject.Parse(jsonStr)
    End Function

    ''' <summary>
    ''' 获取当前Excel上下文用于占位符替换
    ''' </summary>
    Public Shared Function GetExcelContext(excelApp As Object) As Dictionary(Of String, String)
        Dim context As New Dictionary(Of String, String)
        
        Try
            Dim ws = excelApp.ActiveSheet
            Dim usedRange = ws.UsedRange
            
            ' 最后一行
            Dim lastRow = usedRange.Row + usedRange.Rows.Count - 1
            context("lastRow") = lastRow.ToString()
            
            ' 最后一列
            Dim lastCol = usedRange.Column + usedRange.Columns.Count - 1
            context("lastCol") = GetColumnLetter(lastCol)
            
            ' 当前选择
            Dim selection = excelApp.Selection
            If selection IsNot Nothing Then
                context("selection") = selection.Address(False, False)
            End If
            
        Catch ex As Exception
            ' 默认值
            context("lastRow") = "100"
            context("lastCol") = "Z"
            context("selection") = "A1"
        End Try
        
        Return context
    End Function

    ''' <summary>
    ''' 数字转列字母
    ''' </summary>
    Private Shared Function GetColumnLetter(colNum As Integer) As String
        Dim result = ""
        While colNum > 0
            colNum -= 1
            result = Chr(65 + (colNum Mod 26)) & result
            colNum \= 26
        End While
        Return result
    End Function

    ''' <summary>
    ''' 生成校验失败后的重试提示
    ''' </summary>
    Public Shared Function GetRetryPrompt(originalRequest As String, errorMessage As String) As String
        Return $"你之前返回的JSON命令格式有误: {errorMessage}

请严格按照以下格式重新返回:
```json
{{
  ""command"": ""ApplyFormula"",
  ""params"": {{
    ""targetRange"": ""C1:C{{lastRow}}"",
    ""formula"": ""=A1+B1"",
    ""fillDown"": true
  }}
}}
```

注意:
1. 必须是有效的JSON格式
2. 动态范围使用占位符 {{lastRow}} 而不是JS表达式
3. command必须是: {String.Join(", ", SupportedCommands)}

原始请求: {originalRequest}"
    End Function

End Class
