Imports System.Diagnostics
Imports System.Drawing
Imports System.Threading
Imports System.Windows.Forms
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.PowerPoint

''' <summary>
''' PowerPoint 内联灰色补全文本管理器 (Ghost Text Manager)
''' 负责在演示文稿中显示、接受和清除灰色预览文本
''' </summary>
Public Class PowerPointGhostTextManager
    Private _pptApp As Microsoft.Office.Interop.PowerPoint.Application
    Private _ghostTextRange As TextRange  ' 跟踪灰色文本的 TextRange
    Private _originalSelStart As Integer  ' 记录原始选区起始位置
    Private _uiSyncContext As SynchronizationContext  ' UI线程同步上下文
    
    ' 灰色颜色值
    Private Const GHOST_TEXT_COLOR As Integer = &H999999  ' RGB(153, 153, 153)
    
    Public Sub New(pptApp As Microsoft.Office.Interop.PowerPoint.Application)
        _pptApp = pptApp
        _uiSyncContext = SynchronizationContext.Current
        If _uiSyncContext Is Nothing Then
            _uiSyncContext = New WindowsFormsSynchronizationContext()
        End If
    End Sub
    
    ''' <summary>
    ''' 检查是否有活动的 ghost text
    ''' </summary>
    Public ReadOnly Property HasGhostText As Boolean
        Get
            Return _ghostTextRange IsNot Nothing
        End Get
    End Property
    
    ''' <summary>
    ''' 获取当前 ghost text 内容
    ''' </summary>
    Public ReadOnly Property CurrentGhostText As String
        Get
            If _ghostTextRange IsNot Nothing Then
                Try
                    Return _ghostTextRange.Text
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
            If _pptApp Is Nothing OrElse _pptApp.ActiveWindow Is Nothing Then Return
            
            Dim sel = _pptApp.ActiveWindow.Selection
            If sel Is Nothing OrElse sel.Type <> PpSelectionType.ppSelectionText Then Return
            
            Dim textRange = sel.TextRange
            If textRange Is Nothing Then Return
            
            ' 记录原始位置
            _originalSelStart = textRange.Start
            
            ' 在选区末尾插入灰色文本
            _ghostTextRange = textRange.InsertAfter(suggestion)
            
            ' 设置灰色字体
            _ghostTextRange.Font.Color.RGB = GHOST_TEXT_COLOR
            _ghostTextRange.Font.Italic = MsoTriState.msoTrue  ' 斜体以区分
            
            ' 将光标移回原位
            textRange.Select()
            
            Debug.WriteLine($"[PPT GhostText] 显示补全: '{suggestion}'")
            
        Catch ex As Exception
            Debug.WriteLine($"[PPT GhostText] ShowGhostText 出错: {ex.Message}")
            _ghostTextRange = Nothing
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
            If _ghostTextRange Is Nothing Then Return
            If _pptApp Is Nothing Then Return
            
            ' 将文本颜色改为自动（黑色）
            _ghostTextRange.Font.Color.RGB = &H0  ' 黑色
            _ghostTextRange.Font.Italic = MsoTriState.msoFalse  ' 取消斜体
            
            ' 将光标移动到补全文本末尾
            Dim endPos = _ghostTextRange.Start + _ghostTextRange.Length
            _ghostTextRange.Parent.Characters(endPos, 0).Select()
            
            Debug.WriteLine($"[PPT GhostText] 已接受补全")
            
            _ghostTextRange = Nothing
            
        Catch ex As Exception
            Debug.WriteLine($"[PPT GhostText] AcceptGhostText 出错: {ex.Message}")
            _ghostTextRange = Nothing
        End Try
    End Sub
    
    ''' <summary>
    ''' 清除灰色文本
    ''' </summary>
    Public Sub ClearGhostText()
        If _uiSyncContext IsNot Nothing Then
            _uiSyncContext.Post(Sub(state) ClearGhostTextInternal(), Nothing)
        Else
            ClearGhostTextInternal()
        End If
    End Sub
    
    Private Sub ClearGhostTextInternal()
        Try
            If _ghostTextRange Is Nothing Then Return
            
            ' 删除灰色文本
            _ghostTextRange.Delete()
            
            Debug.WriteLine($"[PPT GhostText] 已清除补全")
            
        Catch ex As Exception
            ' TextRange 可能已失效，忽略错误
            Debug.WriteLine($"[PPT GhostText] ClearGhostText 出错: {ex.Message}")
        Finally
            _ghostTextRange = Nothing
        End Try
    End Sub
    
    ''' <summary>
    ''' 检查光标是否仍在 ghost text 起始位置
    ''' </summary>
    Public Function IsCursorAtGhostTextStart() As Boolean
        Try
            If _ghostTextRange Is Nothing Then Return False
            If _pptApp Is Nothing OrElse _pptApp.ActiveWindow Is Nothing Then Return False
            
            Dim sel = _pptApp.ActiveWindow.Selection
            If sel Is Nothing OrElse sel.Type <> PpSelectionType.ppSelectionText Then Return False
            
            Dim currentPos = sel.TextRange.Start
            Return currentPos = _originalSelStart
            
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
