' ShareRibbon\Controls\Models\HistoryMessage.vb
' 历史消息实体类，用于存储聊天历史记录

''' <summary>
''' 聊天历史消息实体
''' </summary>
Public Class HistoryMessage
    ''' <summary>
    ''' 消息角色：system, user, assistant
    ''' </summary>
    Public Property role As String

    ''' <summary>
    ''' 消息内容
    ''' </summary>
    Public Property content As String

    ''' <summary>
    ''' 消息时间戳
    ''' </summary>
    Public Property Timestamp As DateTime = DateTime.Now

    ''' <summary>
    ''' 消息UUID（可选）
    ''' </summary>
    Public Property Uuid As String
End Class

''' <summary>
''' 文件内容解析结果
''' </summary>
Public Class FileContentResult
    ''' <summary>
    ''' 文件名
    ''' </summary>
    Public Property FileName As String

    ''' <summary>
    ''' 文件类型：Excel, Word, Text, CSV 等
    ''' </summary>
    Public Property FileType As String

    ''' <summary>
    ''' 格式化的内容字符串
    ''' </summary>
    Public Property ParsedContent As String

    ''' <summary>
    ''' 原始数据，可用于进一步处理
    ''' </summary>
    Public Property RawData As Object
End Class

''' <summary>
''' 发送消息时的引用内容项
''' </summary>
Public Class SendMessageReferenceContentItem
    Public Property Id As String
    Public Property SheetName As String
    Public Property Address As String
End Class

''' <summary>
''' Token使用信息
''' </summary>
Public Structure TokenInfo
    Public PromptTokens As Integer
    Public CompletionTokens As Integer
    Public TotalTokens As Integer
End Structure
