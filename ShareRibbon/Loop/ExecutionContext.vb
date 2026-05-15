' ShareRibbon\Loop\ExecutionContext.vb
' 执行上下文模型 - 贯穿整个自检Loop的数据载体

Imports System.Collections.Generic
Imports Newtonsoft.Json.Linq

''' <summary>
''' 执行上下文 - 贯穿整个自检Loop流程的数据载体
''' </summary>
Public Class ExecutionContext

    ''' <summary>用户原始输入消息</summary>
    Public Property UserMessage As String = String.Empty

    ''' <summary>原始问题（未经处理的）</summary>
    Public Property OriginalQuestion As String = String.Empty

    ''' <summary>用户请求的操作类型</summary>
    Public Property RequestedOperation As RequestedOperation = RequestedOperation.GeneralQuery

    ''' <summary>是否需要选区</summary>
    Public Property RequiresSelection As Boolean = False

    ''' <summary>当前选区信息</summary>
    Public Property SelectionInfo As SelectionInfo = Nothing

    ''' <summary>当前Office文档内容摘要</summary>
    Public Property OfficeContent As OfficeContentInfo = Nothing

    ''' <summary>是否是追问模式</summary>
    Public Property IsFollowUp As Boolean = False

    ''' <summary>会话历史</summary>
    Public Property ConversationHistory As List(Of HistoryMessage)

    ''' <summary>期望的指令格式</summary>
    Public Property ExpectedFormat As InstructionFormat = InstructionFormat.None

    ''' <summary>意图识别结果</summary>
    Public Property IntentResult As IntentResult = Nothing

    ''' <summary>附加文件路径</summary>
    Public Property FilePaths As List(Of String)

    ''' <summary>附加文件内容</summary>
    Public Property FileContent As String = String.Empty

    ''' <summary>当前Office应用类型</summary>
    Public Property OfficeAppType As OfficeAppType = OfficeAppType.Unknown

    ''' <summary>段落文本列表（排版/校对使用）</summary>
    Public Property Paragraphs As List(Of String)

    ''' <summary>Word段落对象列表（排版使用）</summary>
    Public Property WordParagraphs As List(Of Object)

    Public Sub New()
        ConversationHistory = New List(Of HistoryMessage)()
        FilePaths = New List(Of String)()
        Paragraphs = New List(Of String)()
        WordParagraphs = New List(Of Object)()
    End Sub

End Class

''' <summary>
''' 用户请求的操作类型
''' </summary>
Public Enum RequestedOperation
    GeneralQuery
    Reformat
    Proofread
    Translation
    Continuation
    DataAnalysis
    ChartGeneration
End Enum

''' <summary>
''' 指令格式类型
''' </summary>
Public Enum InstructionFormat
    None
    DslJson
    ProofreadJson
    OpenXmlFragment
    LegacyJsonCommand
End Enum

''' <summary>
''' Office内容信息
''' </summary>
Public Class OfficeContentInfo

    ''' <summary>内容是否为空</summary>
    Public Property IsEmpty As Boolean = True

    ''' <summary>是否为只读</summary>
    Public Property IsReadOnly As Boolean = False

    ''' <summary>文档类型</summary>
    Public Property DocumentType As String = String.Empty

    ''' <summary>段落数</summary>
    Public Property ParagraphCount As Integer = 0

    ''' <summary>选中的段落文本</summary>
    Public Property SelectedParagraphs As List(Of String)

    ''' <summary>文档标题</summary>
    Public Property Title As String = String.Empty

    Public Sub New()
        SelectedParagraphs = New List(Of String)()
    End Sub

End Class

''' <summary>
''' Office应用类型
''' </summary>
Public Enum OfficeAppType
    Unknown
    Word
    Excel
    PowerPoint
End Enum
