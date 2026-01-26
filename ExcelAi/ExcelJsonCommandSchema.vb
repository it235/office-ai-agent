' ExcelAi\ExcelJsonCommandSchema.vb
' Excel JSON命令Schema定义和校验

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
    ''' 校验JSON命令是否有效
    ''' </summary>
    Public Shared Function ValidateCommand(json As JObject, ByRef errorMessage As String) As Boolean
        Try
            errorMessage = ""
            
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
        Dim targetRange = params("targetRange")?.ToString()
        Dim formula = params("formula")?.ToString()
        
        If String.IsNullOrEmpty(targetRange) Then
            errorMessage = "ApplyFormula缺少targetRange参数"
            Return False
        End If
        
        If String.IsNullOrEmpty(formula) Then
            errorMessage = "ApplyFormula缺少formula参数"
            Return False
        End If
        
        ' 校验范围格式 (支持占位符)
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
        Dim targetRange = params("targetRange")?.ToString()
        Dim data = params("data")
        
        If String.IsNullOrEmpty(targetRange) Then
            errorMessage = "WriteData缺少targetRange参数"
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
        ' 支持格式: A1, A1:B10, A:A, 1:1, A1:{lastRow}, {selection}
        Dim pattern = "^([A-Za-z]+\d+|[A-Za-z]+|[0-9]+)(:[A-Za-z]*\d*|\{[a-zA-Z]+\})?$"
        Return Regex.IsMatch(range, pattern) OrElse range.Contains("{")
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
