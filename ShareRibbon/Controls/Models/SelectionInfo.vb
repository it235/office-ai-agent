' ShareRibbon\Controls\Models\SelectionInfo.vb
' 选区信息实体类，用于存储文档选区的详细信息

''' <summary>
''' 存储文档选区的信息，用于校对/排版等功能的写回操作
''' </summary>
Public Class SelectionInfo
    ''' <summary>
    ''' 文档路径
    ''' </summary>
    Public Property DocumentPath As String

    ''' <summary>
    ''' 选区起始位置
    ''' </summary>
    Public Property StartPos As Integer

    ''' <summary>
    ''' 选区结束位置
    ''' </summary>
    Public Property EndPos As Integer

    ''' <summary>
    ''' 选中的文本内容
    ''' </summary>
    Public Property SelectedText As String

    ''' <summary>
    ''' 工作表名称（Excel专用）
    ''' </summary>
    Public Property SheetName As String

    ''' <summary>
    ''' 单元格地址（Excel专用）
    ''' </summary>
    Public Property CellAddress As String
End Class
