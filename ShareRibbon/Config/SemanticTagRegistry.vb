' ShareRibbon\Config\SemanticTagRegistry.vb
' 两层语义标签注册表

''' <summary>
''' 语义标签注册表 - 管理两层标签体系
''' Layer 1: 固定语义标签（AI可见的粗粒度分类）
''' Layer 2: 模板级细分标签（从模板动态生成的细粒度标签）
''' </summary>
Public Class SemanticTagRegistry

#Region "Layer 1 固定语义标签"
    Public Const TAG_TITLE As String = "title"
    Public Const TAG_BODY As String = "body"
    Public Const TAG_LIST As String = "list"
    Public Const TAG_QUOTE As String = "quote"
    Public Const TAG_CAPTION As String = "caption"
    Public Const TAG_HEADER As String = "header"
    Public Const TAG_FOOTER As String = "footer"
    Public Const TAG_CODE As String = "code"
    Public Const TAG_TABLE_HEADER As String = "table_header"
#End Region

#Region "Layer 2 常用细分标签"
    Public Const TAG_TITLE_1 As String = "title.1"
    Public Const TAG_TITLE_2 As String = "title.2"
    Public Const TAG_TITLE_3 As String = "title.3"
    Public Const TAG_BODY_NORMAL As String = "body.normal"
    Public Const TAG_BODY_EMPHASIS As String = "body.emphasis"
    Public Const TAG_LIST_ORDERED As String = "list.ordered"
    Public Const TAG_LIST_UNORDERED As String = "list.unordered"
#End Region

    ''' <summary>所有Layer 1标签</summary>
    Private Shared ReadOnly Layer1Tags As String() = {
        TAG_TITLE, TAG_BODY, TAG_LIST, TAG_QUOTE,
        TAG_CAPTION, TAG_HEADER, TAG_FOOTER, TAG_CODE, TAG_TABLE_HEADER
    }

    ''' <summary>
    ''' 获取标签层级（1=语义层, 2=细分层）
    ''' </summary>
    Public Shared Function GetTagLevel(tagId As String) As Integer
        If String.IsNullOrEmpty(tagId) Then Return 0
        If Layer1Tags.Contains(tagId) Then Return 1
        If tagId.Contains(".") Then Return 2
        Return 1
    End Function

    ''' <summary>
    ''' 获取父级标签ID（"title.1" → "title"）
    ''' </summary>
    Public Shared Function GetParentTag(tagId As String) As String
        If String.IsNullOrEmpty(tagId) Then Return ""
        Dim dotIndex = tagId.IndexOf("."c)
        If dotIndex > 0 Then
            Return tagId.Substring(0, dotIndex)
        End If
        Return ""
    End Function

    ''' <summary>
    ''' 验证标签ID是否合法（Layer1已知标签 或 parent.sub格式）
    ''' </summary>
    Public Shared Function IsValidTag(tagId As String) As Boolean
        If String.IsNullOrEmpty(tagId) Then Return False
        If Layer1Tags.Contains(tagId) Then Return True
        Dim parent = GetParentTag(tagId)
        Return Layer1Tags.Contains(parent)
    End Function

    ''' <summary>
    ''' 获取所有Layer 1标签及其默认显示名称
    ''' </summary>
    Public Shared Function GetLayer1TagDescriptions() As Dictionary(Of String, String)
        Return New Dictionary(Of String, String) From {
            {TAG_TITLE, "标题"},
            {TAG_BODY, "正文"},
            {TAG_LIST, "列表"},
            {TAG_QUOTE, "引用"},
            {TAG_CAPTION, "图表题注"},
            {TAG_HEADER, "页眉"},
            {TAG_FOOTER, "页脚"},
            {TAG_CODE, "代码"},
            {TAG_TABLE_HEADER, "表头"}
        }
    End Function

    ''' <summary>
    ''' 获取常用Layer 2标签及其默认显示名称
    ''' </summary>
    Public Shared Function GetCommonLayer2Tags() As Dictionary(Of String, String)
        Return New Dictionary(Of String, String) From {
            {TAG_TITLE_1, "一级标题"},
            {TAG_TITLE_2, "二级标题"},
            {TAG_TITLE_3, "三级标题"},
            {TAG_BODY_NORMAL, "正文"},
            {TAG_BODY_EMPHASIS, "强调段落"},
            {TAG_LIST_ORDERED, "有序列表"},
            {TAG_LIST_UNORDERED, "无序列表"}
        }
    End Function
End Class
