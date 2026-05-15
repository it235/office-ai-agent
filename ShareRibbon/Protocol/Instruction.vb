' ShareRibbon\Protocol\Instruction.vb
' 指令实例模型

Imports Newtonsoft.Json.Linq

''' <summary>
''' 指令实例
''' </summary>
Public Class Instruction

    ''' <summary>指令唯一ID</summary>
    Public Property Id As String

    ''' <summary>操作类型（如setParagraphStyle）</summary>
    Public Property Operation As String

    ''' <summary>目标对象描述</summary>
    Public Property Target As JObject

    ''' <summary>操作参数</summary>
    Public Property Params As JObject

    ''' <summary>预期结果描述</summary>
    Public Property Expected As JObject

    ''' <summary>回滚信息</summary>
    Public Property Rollback As JObject

    ''' <summary>元数据</summary>
    Public Property Metadata As JObject

    Public Sub New(operation As String, params As JObject, id As String)
        Me.Id = If(String.IsNullOrEmpty(id), Guid.NewGuid().ToString("N"), id)
        Me.Operation = operation
        Me.Params = If(params, New JObject())
        Me.Target = New JObject()
        Me.Expected = New JObject()
        Me.Rollback = New JObject()
        Me.Metadata = New JObject()
    End Sub

    Public Sub New(operation As String, params As JObject)
        Me.New(operation, params, Nothing)
    End Sub

    ''' <summary>获取参数值（安全）</summary>
    Public Function GetParam(path As String, defaultValue As Object) As Object
        Try
            Dim token = Params.SelectToken(path)
            If token IsNot Nothing Then
                Return token.ToObject(Of Object)()
            End If
        Catch
        End Try
        Return defaultValue
    End Function

    ''' <summary>获取目标选择器</summary>
    Public Function GetTargetSelector() As String
        If Target IsNot Nothing AndAlso Target("selector") IsNot Nothing Then
            Return Target("selector").ToString()
        End If
        Return String.Empty
    End Function

    ''' <summary>获取目标类型</summary>
    Public Function GetTargetType() As String
        If Target IsNot Nothing AndAlso Target("type") IsNot Nothing Then
            Return Target("type").ToString()
        End If
        Return String.Empty
    End Function

    ''' <summary>获取操作描述（用于日志和展示）</summary>
    Public Function GetDescription() As String
        If Expected IsNot Nothing AndAlso Expected("description") IsNot Nothing Then
            Return Expected("description").ToString()
        End If
        Return $"{Operation}"
    End Function

    ''' <summary>
    ''' 从DSL JSON字符串解析指令列表
    ''' </summary>
    Public Shared Function ParseInstructions(jsonText As String) As List(Of Instruction)
        Dim instructions As New List(Of Instruction)()
        If String.IsNullOrWhiteSpace(jsonText) Then Return instructions

        Try
            ' 尝试提取代码块中的JSON
            Dim content = jsonText.Trim()
            Dim codeBlockMatch = System.Text.RegularExpressions.Regex.Match(content, "```(?:json)?\s*([\s\S]*?)\s*```", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            If codeBlockMatch.Success Then
                content = codeBlockMatch.Groups(1).Value.Trim()
            End If

            ' 尝试直接解析JSON对象
            Dim json = JObject.Parse(content)

            If json("instructions") Is Nothing OrElse json("instructions").Type <> JTokenType.Array Then
                Return instructions
            End If

            Dim instructionArray = CType(json("instructions"), JArray)
            For Each item In instructionArray
                If item.Type <> JTokenType.Object Then Continue For
                Dim itemObj = CType(item, JObject)
                If itemObj("op") Is Nothing Then Continue For

                Dim op = itemObj("op").ToString()
                Dim params = If(itemObj("params"), New JObject())
                Dim id = itemObj("id")?.ToString()

                Dim instruction = New Instruction(op, CType(params, JObject), id)
                If itemObj("target") IsNot Nothing Then
                    instruction.Target = CType(itemObj("target"), JObject)
                End If
                If itemObj("expected") IsNot Nothing Then
                    instruction.Expected = CType(itemObj("expected"), JObject)
                End If
                If itemObj("rollback") IsNot Nothing Then
                    instruction.Rollback = CType(itemObj("rollback"), JObject)
                End If
                If itemObj("metadata") IsNot Nothing Then
                    instruction.Metadata = CType(itemObj("metadata"), JObject)
                End If

                instructions.Add(instruction)
            Next
        Catch
        End Try

        Return instructions
    End Function

End Class
