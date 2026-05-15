' ShareRibbon\Protocol\ReformatInstructionSet.vb
' 排版指令集 - 定义所有文档排版相关的DSL指令

Imports System.Collections.Generic
Imports Newtonsoft.Json.Linq

''' <summary>
''' 排版指令集 - 包含所有支持的文档排版操作指令
''' </summary>
Public Class ReformatInstructionSet

    ''' <summary>
    ''' 段落样式指令 - 设置段落样式
    ''' </summary>
    Public Class SetParagraphStyleInstruction
        Inherits DslInstruction

        ''' <summary>目标段落选择器</summary>
        Public Property Target As ParagraphTarget
        ''' <summary>样式参数</summary>
        Public Property Params As ParagraphStyleParams

        Public Sub New()
            MyBase.New("setParagraphStyle")
        End Sub
    End Class

    ''' <summary>
    ''' 字符格式指令 - 设置文本字符格式
    ''' </summary>
    Public Class SetCharacterFormatInstruction
        Inherits DslInstruction

        ''' <summary>目标文本选择器</summary>
        Public Property Target As TextTarget
        ''' <summary>样式参数</summary>
        Public Property Params As CharacterStyleParams

        Public Sub New()
            MyBase.New("setCharacterFormat")
        End Sub
    End Class

    ''' <summary>
    ''' 插入表格指令
    ''' </summary>
    Public Class InsertTableInstruction
        Inherits DslInstruction

        ''' <summary>目标位置</summary>
        Public Property Target As PositionTarget
        ''' <summary>表格参数</summary>
        Public Property Params As TableParams

        Public Sub New()
            MyBase.New("insertTable")
        End Sub
    End Class

    ''' <summary>
    ''' 格式化表格指令
    ''' </summary>
    Public Class FormatTableInstruction
        Inherits DslInstruction

        ''' <summary>目标表格选择器</summary>
        Public Property Target As TableTarget
        ''' <summary>格式化参数</summary>
        Public Property Params As TableFormatParams

        Public Sub New()
            MyBase.New("formatTable")
        End Sub
    End Class

    ''' <summary>
    ''' 页面设置指令
    ''' </summary>
    Public Class SetPageSetupInstruction
        Inherits DslInstruction

        ''' <summary>页面设置参数</summary>
        Public Property Params As PageSetupParams

        Public Sub New()
            MyBase.New("setPageSetup")
        End Sub
    End Class

    ''' <summary>
    ''' 插入分隔符指令
    ''' </summary>
    Public Class InsertBreakInstruction
        Inherits DslInstruction

        ''' <summary>目标位置</summary>
        Public Property Target As PositionTarget
        ''' <summary>分隔符参数</summary>
        Public Property Params As BreakParams

        Public Sub New()
            MyBase.New("insertBreak")
        End Sub
    End Class

    ''' <summary>
    ''' 应用列表格式指令
    ''' </summary>
    Public Class ApplyListFormatInstruction
        Inherits DslInstruction

        ''' <summary>目标段落选择器</summary>
        Public Property Target As ParagraphTarget
        ''' <summary>列表格式参数</summary>
        Public Property Params As ListFormatParams

        Public Sub New()
            MyBase.New("applyListFormat")
        End Sub
    End Class

    ''' <summary>
    ''' 分栏设置指令
    ''' </summary>
    Public Class SetColumnFormatInstruction
        Inherits DslInstruction

        ''' <summary>目标区域</summary>
        Public Property Target As SectionTarget
        ''' <summary>分栏参数</summary>
        Public Property Params As ColumnFormatParams

        Public Sub New()
            MyBase.New("setColumnFormat")
        End Sub
    End Class

    ''' <summary>
    ''' 插入页眉页脚指令
    ''' </summary>
    Public Class InsertHeaderFooterInstruction
        Inherits DslInstruction

        ''' <summary>目标类型</summary>
        Public Property Params As HeaderFooterParams

        Public Sub New()
            MyBase.New("insertHeaderFooter")
        End Sub
    End Class

    ''' <summary>
    ''' 生成目录指令
    ''' </summary>
    Public Class GenerateTocInstruction
        Inherits DslInstruction

        ''' <summary>目标位置</summary>
        Public Property Target As PositionTarget
        ''' <summary>目录参数</summary>
        Public Property Params As TocParams

        Public Sub New()
            MyBase.New("generateToc")
        End Sub
    End Class

    ''' <summary>
    ''' 段落样式参数
    ''' </summary>
    Public Class ParagraphStyleParams
        ''' <summary>样式名称</summary>
        Public Property StyleName As String
        ''' <summary>字体设置</summary>
        Public Property Font As FontConfig
        ''' <summary>对齐方式</summary>
        Public Property Alignment As String
        ''' <summary>行距设置</summary>
        Public Property Spacing As SpacingConfig
        ''' <summary>缩进设置</summary>
        Public Property Indent As IndentConfig
    End Class

    ''' <summary>
    ''' 字符样式参数
    ''' </summary>
    Public Class CharacterStyleParams
        ''' <summary>字体设置</summary>
        Public Property Font As FontConfig
        ''' <summary>颜色</summary>
        Public Property Color As String
        ''' <summary>是否加粗</summary>
        Public Property Bold As Boolean
        ''' <summary>是否斜体</summary>
        Public Property Italic As Boolean
        ''' <summary>下划线类型</summary>
        Public Property Underline As String
    End Class

    ''' <summary>
    ''' 表格参数
    ''' </summary>
    Public Class TableParams
        ''' <summary>行数</summary>
        Public Property Rows As Integer
        ''' <summary>列数</summary>
        Public Property Cols As Integer
        ''' <summary>样式名称</summary>
        Public Property Style As String
        ''' <summary>是否包含表头行</summary>
        Public Property HeaderRow As Boolean
        ''' <summary>表格数据</summary>
        Public Property Data As List(Of List(Of Object))
    End Class

    ''' <summary>
    ''' 表格格式化参数
    ''' </summary>
    Public Class TableFormatParams
        ''' <summary>样式名称</summary>
        Public Property Style As String
        ''' <summary>是否显示边框</summary>
        Public Property Borders As Boolean
        ''' <summary>是否包含表头行</summary>
        Public Property HeaderRow As Boolean
    End Class

    ''' <summary>
    ''' 页面设置参数
    ''' </summary>
    Public Class PageSetupParams
        ''' <summary>边距设置</summary>
        Public Property Margins As MarginConfig
        ''' <summary>页面方向</summary>
        Public Property Orientation As String
        ''' <summary>纸张大小</summary>
        Public Property PaperSize As String
    End Class

    ''' <summary>
    ''' 分隔符参数
    ''' </summary>
    Public Class BreakParams
        ''' <summary>分隔符类型</summary>
        Public Property Type As String
    End Class

    ''' <summary>
    ''' 列表格式参数
    ''' </summary>
    Public Class ListFormatParams
        ''' <summary>列表类型</summary>
        Public Property ListType As String
        ''' <summary>编号格式</summary>
        Public Property NumberFormat As String
    End Class

    ''' <summary>
    ''' 分栏参数
    ''' </summary>
    Public Class ColumnFormatParams
        ''' <summary>栏数</summary>
        Public Property ColumnCount As Integer
        ''' <summary>栏宽</summary>
        Public Property ColumnWidth As Double
        ''' <summary>是否显示分隔线</summary>
        Public Property Separator As Boolean
    End Class

    ''' <summary>
    ''' 页眉页脚参数
    ''' </summary>
    Public Class HeaderFooterParams
        ''' <summary>类型（header/footer）</summary>
        Public Property Type As String
        ''' <summary>内容</summary>
        Public Property Content As String
        ''' <summary>对齐方式</summary>
        Public Property Alignment As String
    End Class

    ''' <summary>
    ''' 目录参数
    ''' </summary>
    Public Class TocParams
        ''' <summary>显示级别</summary>
        Public Property Levels As Integer
        ''' <summary>是否包含页码</summary>
        Public Property IncludePageNumbers As Boolean
    End Class

    ''' <summary>
    ''' 字体配置
    ''' </summary>
    Public Class FontConfig
        ''' <summary>字体名称</summary>
        Public Property Name As String
        ''' <summary>字体大小</summary>
        Public Property Size As Double
        ''' <summary>是否加粗</summary>
        Public Property Bold As Boolean
        ''' <summary>是否斜体</summary>
        Public Property Italic As Boolean
        ''' <summary>字体颜色</summary>
        Public Property Color As String
    End Class

    ''' <summary>
    ''' 间距配置
    ''' </summary>
    Public Class SpacingConfig
        ''' <summary>段前间距（磅）</summary>
        Public Property Before As Double
        ''' <summary>段后间距（磅）</summary>
        Public Property After As Double
        ''' <summary>行距（倍数）</summary>
        Public Property Line As Double
    End Class

    ''' <summary>
    ''' 缩进配置
    ''' </summary>
    Public Class IndentConfig
        ''' <summary>首行缩进（字符）</summary>
        Public Property FirstLine As Double
        ''' <summary>左缩进（字符）</summary>
        Public Property Left As Double
        ''' <summary>右缩进（字符）</summary>
        Public Property Right As Double
    End Class

    ''' <summary>
    ''' 段落目标选择器
    ''' </summary>
    Public Class ParagraphTarget
        ''' <summary>选择器类型</summary>
        Public Property Type As String
        ''' <summary>语义选择器</summary>
        Public Property Selector As String
        ''' <summary>索引位置</summary>
        Public Property Index As Integer?
    End Class

    ''' <summary>
    ''' 文本目标选择器
    ''' </summary>
    Public Class TextTarget
        ''' <summary>选择器类型</summary>
        Public Property Type As String
        ''' <summary>文本匹配模式</summary>
        Public Property Match As String
        ''' <summary>索引位置</summary>
        Public Property Index As Integer?
    End Class

    ''' <summary>
    ''' 表格目标选择器
    ''' </summary>
    Public Class TableTarget
        ''' <summary>选择器类型</summary>
        Public Property Type As String
        ''' <summary>语义选择器</summary>
        Public Property Selector As String
        ''' <summary>索引位置</summary>
        Public Property Index As Integer?
    End Class

    ''' <summary>
    ''' 位置目标选择器
    ''' </summary>
    Public Class PositionTarget
        ''' <summary>选择器类型</summary>
        Public Property Type As String
        ''' <summary>位置</summary>
        Public Property Position As String
    End Class

    ''' <summary>
    ''' 区域目标选择器
    ''' </summary>
    Public Class SectionTarget
        ''' <summary>选择器类型</summary>
        Public Property Type As String
    End Class

    ''' <summary>
    ''' 边距配置
    ''' </summary>
    Public Class MarginConfig
        ''' <summary>上边距</summary>
        Public Property Top As Double
        ''' <summary>下边距</summary>
        Public Property Bottom As Double
        ''' <summary>左边距</summary>
        Public Property Left As Double
        ''' <summary>右边距</summary>
        Public Property Right As Double
    End Class

    ''' <summary>
    ''' 从DSL JSON创建排版指令
    ''' </summary>
    Public Shared Function FromJson(json As JObject) As List(Of DslInstruction)
        Dim instructions As New List(Of DslInstruction)()

        If json("instructions") Is Nothing Then
            Return instructions
        End If

        For Each item In json("instructions")
            Dim instruction = CreateSingleInstruction(CType(item, JObject))
            If instruction IsNot Nothing Then
                instructions.Add(instruction)
            End If
        Next

        Return instructions
    End Function

    ''' <summary>
    ''' 创建单个指令
    ''' </summary>
    Private Shared Function CreateSingleInstruction(json As JObject) As DslInstruction
        Dim op = json("op")?.ToString()
        If String.IsNullOrEmpty(op) Then
            Return Nothing
        End If

        Select Case op.ToLower()
            Case "setparagraphstyle"
                Return CreateInstruction(Of SetParagraphStyleInstruction)(json)
            Case "setcharacterformat"
                Return CreateInstruction(Of SetCharacterFormatInstruction)(json)
            Case "inserttable"
                Return CreateInstruction(Of InsertTableInstruction)(json)
            Case "formattable"
                Return CreateInstruction(Of FormatTableInstruction)(json)
            Case "setpagesetup"
                Return CreateInstruction(Of SetPageSetupInstruction)(json)
            Case "insertbreak"
                Return CreateInstruction(Of InsertBreakInstruction)(json)
            Case "applylistformat"
                Return CreateInstruction(Of ApplyListFormatInstruction)(json)
            Case "setcolumnformat"
                Return CreateInstruction(Of SetColumnFormatInstruction)(json)
            Case "insertheaderfooter"
                Return CreateInstruction(Of InsertHeaderFooterInstruction)(json)
            Case "generatetoc"
                Return CreateInstruction(Of GenerateTocInstruction)(json)
            Case Else
                Return New DslInstruction(op)
        End Select
    End Function

    ''' <summary>
    ''' 创建具体指令
    ''' </summary>
    Private Shared Function CreateInstruction(Of T As {DslInstruction, New})(json As JObject) As T
        Dim instruction = New T()

        If json("target") IsNot Nothing Then
            instruction.Target = json("target").ToObject(Of JObject)()
        End If

        If json("params") IsNot Nothing Then
            instruction.Params = json("params").ToObject(Of JObject)()
        End If

        If json("expected") IsNot Nothing Then
            instruction.Expected = json("expected").ToObject(Of JObject)()
        End If

        If json("rollback") IsNot Nothing Then
            instruction.Rollback = json("rollback").ToObject(Of JObject)()
        End If

        Return instruction
    End Function
End Class

''' <summary>
''' 基础DSL指令
''' </summary>
Public Class DslInstruction
    ''' <summary>指令ID</summary>
    Public Property Id As String
    ''' <summary>操作类型</summary>
    Public Property Operation As String
    ''' <summary>目标对象</summary>
    Public Property Target As JObject
    ''' <summary>操作参数</summary>
    Public Property Params As JObject
    ''' <summary>预期结果</summary>
    Public Property Expected As JObject
    ''' <summary>回滚信息</summary>
    Public Property Rollback As JObject
    ''' <summary>元数据</summary>
    Public Property Metadata As JObject

    Public Sub New()
        Id = Guid.NewGuid().ToString("N")
    End Sub

    Public Sub New(operation As String)
        Me.New()
        Me.Operation = operation
    End Sub

    Public Sub New(operation As String, id As String)
        Me.New(operation)
        Me.Id = id
    End Sub
End Class
