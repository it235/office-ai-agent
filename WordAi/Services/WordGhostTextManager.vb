Imports System.Drawing
Imports System.Threading
Imports System.Diagnostics
Imports Microsoft.Office.Interop.Word

''' <summary>
''' Word 内联灰色补全文本管理器 (Ghost Text Manager)
''' 负责在文档中显示、接受和清除灰色预览文本
''' </summary>
Public Class WordGhostTextManager
    Private _wordApp As Application
    Private _ghostRange As Range  ' 跟踪灰色文本的 Range
    Private _originalCursorPos As Integer  ' 记录原始光标位置
    Private _ghostTextContent As String = ""  ' 记录ghost text原始内容（用于安全删除）
    Private _uiSyncContext As SynchronizationContext  ' UI线程同步上下文
    
    ' 灰色颜色值
    Private Const GHOST_TEXT_COLOR As Integer = &H999999  ' RGB(153, 153, 153)
    
    Public Sub New(wordApp As Application)
        _wordApp = wordApp
        _uiSyncContext = SynchronizationContext.Current
        If _uiSyncContext Is Nothing Then
            _uiSyncContext = New System.Windows.Forms.WindowsFormsSynchronizationContext()
        End If
    End Sub
    
    ''' <summary>
    ''' 检查是否有活动的 ghost text
    ''' </summary>
    Public ReadOnly Property HasGhostText As Boolean
        Get
            Return _ghostRange IsNot Nothing
        End Get
    End Property
    
    ''' <summary>
    ''' 获取当前 ghost text 内容
    ''' </summary>
    Public ReadOnly Property CurrentGhostText As String
        Get
            If _ghostRange IsNot Nothing Then
                Try
                    Return _ghostRange.Text
                Catch
                    Return ""
                End Try
            End If
            Return ""
        End Get
    End Property
    
    ''' <summary>
    ''' 显示灰色补全文本
    ''' </summary>
    Public Sub ShowGhostText(suggestion As String)
        ' 确保在 UI 线程执行
        If _uiSyncContext IsNot Nothing Then
            _uiSyncContext.Post(Sub(state) ShowGhostTextInternal(suggestion), Nothing)
        Else
            ShowGhostTextInternal(suggestion)
        End If
    End Sub
    
    Private Sub ShowGhostTextInternal(suggestion As String)
        Try
            ' 先清除旧的 ghost text
            ClearGhostTextInternal()
            
            If String.IsNullOrEmpty(suggestion) Then Return
            If _wordApp Is Nothing OrElse _wordApp.ActiveDocument Is Nothing Then Return
            
            Dim sel = _wordApp.Selection
            If sel Is Nothing Then Return
            
            ' 记录原始光标位置
            _originalCursorPos = sel.Range.Start
            
            ' 在光标位置后插入灰色文本
            _ghostRange = sel.Range.Duplicate
            _ghostRange.Collapse(WdCollapseDirection.wdCollapseEnd)
            
            ' 插入文本
            _ghostRange.Text = suggestion
            _ghostTextContent = suggestion
            
            ' 设置灰色字体
            _ghostRange.Font.Color = CType(GHOST_TEXT_COLOR, WdColor)
            _ghostRange.Font.Italic = CInt(True)  ' 斜体以区分
            
            ' 将光标移回原位（不选中灰色文本）
            sel.SetRange(_originalCursorPos, _originalCursorPos)
            
            Debug.WriteLine($"[GhostText] 显示补全: '{suggestion}'")
            
        Catch ex As Exception
            Debug.WriteLine($"[GhostText] ShowGhostText 出错: {ex.Message}")
            _ghostRange = Nothing
        End Try
    End Sub
    
    ''' <summary>
    ''' 接受补全 - 将灰色文本变为正常颜色
    ''' </summary>
    Public Sub AcceptGhostText()
        If _uiSyncContext IsNot Nothing Then
            _uiSyncContext.Post(Sub(state) AcceptGhostTextInternal(), Nothing)
        Else
            AcceptGhostTextInternal()
        End If
    End Sub
    
    Private Sub AcceptGhostTextInternal()
        Try
            If _ghostRange Is Nothing Then Return
            If _wordApp Is Nothing Then Return
            
            ' 将文本颜色改为自动（正常颜色）
            _ghostRange.Font.ColorIndex = WdColorIndex.wdAuto
            _ghostRange.Font.Italic = CInt(False)  ' 取消斜体
            
            ' 将光标移动到补全文本末尾
            _wordApp.Selection.SetRange(_ghostRange.End, _ghostRange.End)
            
            Debug.WriteLine($"[GhostText] 已接受补全")
            
            _ghostRange = Nothing
            
        Catch ex As Exception
            Debug.WriteLine($"[GhostText] AcceptGhostText 出错: {ex.Message}")
            _ghostRange = Nothing
        End Try
    End Sub
    
    ''' <summary>
    ''' 清除灰色文本（异步Post到UI线程）
    ''' </summary>
    Public Sub ClearGhostText()
        If _uiSyncContext IsNot Nothing Then
            _uiSyncContext.Post(Sub(state) ClearGhostTextInternal(), Nothing)
        Else
            ClearGhostTextInternal()
        End If
    End Sub

    ''' <summary>
    ''' 同步清除灰色文本（调用方已在UI线程时使用，避免Post延迟导致的竞态）
    ''' </summary>
    Public Sub ClearGhostTextDirect()
        ClearGhostTextInternal()
    End Sub
    
    Private Sub ClearGhostTextInternal()
        Try
            If _ghostRange Is Nothing Then Return
            
            ' 安全删除：ghost range 可能因用户在起始位置输入而扩张
            ' 只删除末尾的 ghost text 部分，不删除用户新输入的内容
            If Not String.IsNullOrEmpty(_ghostTextContent) Then
                Dim rangeText As String = Nothing
                Try
                    rangeText = _ghostRange.Text
                Catch
                End Try

                If rangeText IsNot Nothing AndAlso rangeText.Length > _ghostTextContent.Length Then
                    ' Range 扩张了（用户在 ghost 起始位置输入了文字）
                    ' 仅删除末尾的 ghost text 部分
                    Dim ghostLen = _ghostTextContent.Length
                    Dim safeStart = _ghostRange.End - ghostLen
                    If safeStart >= _ghostRange.Start Then
                        Dim safeRange = _wordApp.ActiveDocument.Range(safeStart, _ghostRange.End)
                        safeRange.Delete()
                        Debug.WriteLine($"[GhostText] 安全清除（跳过 {rangeText.Length - ghostLen} 个用户字符）")
                        _ghostTextContent = ""
                        _ghostRange = Nothing
                        Return
                    End If
                End If
            End If

            ' 常规删除
            _ghostRange.Delete()
            Debug.WriteLine($"[GhostText] 已清除补全")
            
        Catch ex As Exception
            Debug.WriteLine($"[GhostText] ClearGhostText 出错: {ex.Message}")
        Finally
            _ghostRange = Nothing
            _ghostTextContent = ""
        End Try
    End Sub
    
    ''' <summary>
    ''' 检查光标是否仍在 ghost text 起始位置（用于判断是否应该清除）
    ''' </summary>
    Public Function IsCursorAtGhostTextStart() As Boolean
        Try
            If _ghostRange Is Nothing Then Return False
            If _wordApp Is Nothing OrElse _wordApp.Selection Is Nothing Then Return False
            
            Dim currentPos = _wordApp.Selection.Range.Start
            Return currentPos = _originalCursorPos
            
        Catch ex As Exception
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' 清理资源
    ''' </summary>
    Public Sub Dispose()
        ClearGhostTextInternal()
    End Sub
End Class
