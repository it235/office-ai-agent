' ShareRibbon\Loop\InstructionError.vb
' 指令错误模型

''' <summary>
''' 错误级别
''' </summary>
Public Enum ErrorLevel
    Warning
    [Error]
    Critical
End Enum

''' <summary>
''' 指令错误信息
''' </summary>
Public Class InstructionError

    ''' <summary>错误级别</summary>
    Public Property Level As ErrorLevel

    ''' <summary>错误消息</summary>
    Public Property Message As String

    ''' <summary>错误发生的指令ID</summary>
    Public Property InstructionId As String

    ''' <summary>是否可自动修正</summary>
    Public Property IsAutoCorrectable As Boolean

    ''' <summary>修正建议</summary>
    Public Property Suggestion As String

    Public Sub New(level As ErrorLevel, message As String)
        Me.Level = level
        Me.Message = message
        Me.IsAutoCorrectable = False
        Me.Suggestion = Nothing
        Me.InstructionId = String.Empty
    End Sub

    Public Sub New(level As ErrorLevel, message As String, instructionId As String)
        Me.New(level, message)
        Me.InstructionId = instructionId
    End Sub

    Public Sub New(level As ErrorLevel, message As String, instructionId As String, isAutoCorrectable As Boolean, suggestion As String)
        Me.New(level, message, instructionId)
        Me.IsAutoCorrectable = isAutoCorrectable
        Me.Suggestion = suggestion
    End Sub

End Class
