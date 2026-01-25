Imports System.Text

''' <summary>
''' 光标位置类型
''' </summary>
Public Enum CursorPositionType
    DocumentStart   ' 文档开头
    DocumentMiddle  ' 文档中间
    DocumentEnd     ' 文档末尾
End Enum

''' <summary>
''' 插入位置
''' </summary>
Public Enum InsertPosition
    AtCursor        ' 在光标位置插入（Word: 光标位置, PPT: 当前页）
    AfterParagraph  ' 在段落后插入
    NewParagraph    ' 新建段落插入
    DocumentStart   ' 文档开头插入（Word: 文档开头, PPT: 首页）
    DocumentEnd     ' 文档结尾插入（Word: 文档结尾, PPT: 末页）
End Enum

''' <summary>
''' 续写上下文信息
''' </summary>
Public Class ContinuationContext
    ''' <summary>光标前的段落文本</summary>
    Public Property ContextBefore As String

    ''' <summary>光标后的段落文本</summary>
    Public Property ContextAfter As String

    ''' <summary>光标在文档中的位置</summary>
    Public Property CursorPosition As Integer

    ''' <summary>文档路径</summary>
    Public Property DocumentPath As String

    ''' <summary>位置类型（开头/中间/末尾）</summary>
    Public Property PositionType As CursorPositionType

    ''' <summary>当前段落的文本（光标所在段落）</summary>
    Public Property CurrentParagraphText As String

    ''' <summary>光标在当前段落中的偏移</summary>
    Public Property CursorOffsetInParagraph As Integer

    ''' <summary>
    ''' 构建用于AI的上下文提示
    ''' </summary>
    Public Function BuildPrompt() As String
        Dim sb As New StringBuilder()

        If Not String.IsNullOrWhiteSpace(ContextBefore) Then
            sb.AppendLine("【前文内容】")
            sb.AppendLine(ContextBefore)
            sb.AppendLine()
        End If

        If Not String.IsNullOrWhiteSpace(CurrentParagraphText) Then
            sb.AppendLine("【当前段落】")
            sb.AppendLine(CurrentParagraphText)
            sb.AppendLine()
        End If

        If Not String.IsNullOrWhiteSpace(ContextAfter) Then
            sb.AppendLine("【后文内容】")
            sb.AppendLine(ContextAfter)
            sb.AppendLine()
        End If

        ' 添加位置说明
        Select Case PositionType
            Case CursorPositionType.DocumentStart
                sb.AppendLine("【位置说明】光标位于文档开头")
            Case CursorPositionType.DocumentEnd
                sb.AppendLine("【位置说明】光标位于文档末尾")
            Case Else
                sb.AppendLine("【位置说明】光标位于文档中间")
        End Select

        Return sb.ToString()
    End Function
End Class

''' <summary>
''' 续写结果
''' </summary>
Public Class ContinuationResult
    ''' <summary>续写内容</summary>
    Public Property Content As String

    ''' <summary>是否成功</summary>
    Public Property Success As Boolean = True

    ''' <summary>错误消息</summary>
    Public Property ErrorMessage As String = ""
End Class

''' <summary>
''' 续写服务基类 - 定义续写功能的核心接口
''' </summary>
Public MustInherit Class ContinuationService

    ''' <summary>续写完成事件</summary>
    Public Event ContinuationCompleted As EventHandler(Of ContinuationResult)

    ''' <summary>
    ''' 获取光标位置的上下文
    ''' </summary>
    ''' <param name="paragraphsBefore">光标前要获取的段落数</param>
    ''' <param name="paragraphsAfter">光标后要获取的段落数</param>
    ''' <returns>续写上下文信息</returns>
    Public MustOverride Function GetCursorContext(paragraphsBefore As Integer, paragraphsAfter As Integer) As ContinuationContext

    ''' <summary>
    ''' 插入续写内容到文档
    ''' </summary>
    ''' <param name="content">要插入的内容</param>
    ''' <param name="insertPosition">插入位置</param>
    Public MustOverride Sub InsertContinuation(content As String, insertPosition As InsertPosition)

    ''' <summary>
    ''' 检查当前是否可以进行续写
    ''' </summary>
    ''' <returns>是否可以续写</returns>
    Public MustOverride Function CanContinue() As Boolean

    ''' <summary>
    ''' 获取续写的系统提示词
    ''' </summary>
    Public Overridable Function GetSystemPrompt() As String
        Return "你是一个专业的写作助手。根据提供的上下文，自然地续写内容。要求：
1. 保持与原文一致的语言风格、语气和术语
2. 内容要连贯自然，不要重复上文已有内容
3. 只输出续写内容，不要添加任何解释、前缀或标记
4. 如果上下文不足，可以合理推断但保持谨慎
5. 续写长度适中，约100-300字，除非用户另有要求"
    End Function

    ''' <summary>
    ''' 构建续写请求的用户提示
    ''' </summary>
    ''' <param name="context">上下文信息</param>
    ''' <param name="style">可选的风格要求</param>
    Public Overridable Function BuildUserPrompt(context As ContinuationContext, Optional style As String = "") As String
        Dim sb As New StringBuilder()

        sb.AppendLine("请根据以下上下文续写内容：")
        sb.AppendLine()
        sb.Append(context.BuildPrompt())

        If Not String.IsNullOrWhiteSpace(style) Then
            sb.AppendLine()
            sb.AppendLine($"【风格要求】{style}")
        End If

        sb.AppendLine()
        sb.AppendLine("请直接输出续写内容，不要添加任何前缀或说明：")

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 触发续写完成事件
    ''' </summary>
    Protected Sub OnContinuationCompleted(result As ContinuationResult)
        RaiseEvent ContinuationCompleted(Me, result)
    End Sub
End Class
