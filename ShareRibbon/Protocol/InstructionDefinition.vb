' ShareRibbon\Protocol\InstructionDefinition.vb
' 指令定义模型

Imports System.Collections.Generic

''' <summary>
''' 指令定义 - 描述一个指令类型的元数据和参数Schema
''' </summary>
Public Class InstructionDefinition

    ''' <summary>操作名称</summary>
    Public Property Operation As String = String.Empty

    ''' <summary>操作分类（reformat/proofread/general）</summary>
    Public Property Category As String = String.Empty

    ''' <summary>中文描述</summary>
    Public Property DisplayName As String = String.Empty

    ''' <summary>详细说明</summary>
    Public Property Description As String = String.Empty

    ''' <summary>必需参数列表</summary>
    Public Property RequiredParams As List(Of String)

    ''' <summary>可选参数列表</summary>
    Public Property OptionalParams As List(Of String)

    ''' <summary>参数Schema定义</summary>
    Public Property ParamSchema As Dictionary(Of String, ParamType)

    ''' <summary>是否会产生破坏性修改</summary>
    Public Property IsDestructive As Boolean = False

    ''' <summary>是否需要用户确认</summary>
    Public Property RequiresConfirmation As Boolean = False

    ''' <summary>支持的目标类型</summary>
    Public Property SupportedTargetTypes As List(Of String)

    Public Sub New()
        RequiredParams = New List(Of String)()
        OptionalParams = New List(Of String)()
        ParamSchema = New Dictionary(Of String, ParamType)()
        SupportedTargetTypes = New List(Of String)()
    End Sub

End Class

''' <summary>
''' 参数类型定义
''' </summary>
Public Class ParamType

    ''' <summary>基础类型（string/number/boolean/array/object/enum）</summary>
    Public Property BaseType As String = "string"

    ''' <summary>枚举允许值（当BaseType=enum时）</summary>
    Public Property EnumValues As List(Of String)

    ''' <summary>参数描述</summary>
    Public Property Description As String = String.Empty

    ''' <summary>默认值</summary>
    Public Property DefaultValue As Object = Nothing

    ''' <summary>是否可为空</summary>
    Public Property IsNullable As Boolean = True

    Public Sub New(baseType As String)
        Me.BaseType = baseType
        EnumValues = New List(Of String)()
    End Sub

    Public Sub New(baseType As String, enumValues As String())
        Me.New(baseType)
        If enumValues IsNot Nothing Then
            EnumValues = enumValues
        End If
    End Sub

    Public Shared Function StringType() As ParamType
        Return New ParamType("string")
    End Function

    Public Shared Function NumberType() As ParamType
        Return New ParamType("number")
    End Function

    Public Shared Function BooleanType() As ParamType
        Return New ParamType("boolean")
    End Function

    Public Shared Function EnumType(values As String()) As ParamType
        Return New ParamType("enum", values)
    End Function

    Public Shared Function ArrayType() As ParamType
        Return New ParamType("array")
    End Function

    Public Shared Function ObjectType() As ParamType
        Return New ParamType("object")
    End Function

End Class
