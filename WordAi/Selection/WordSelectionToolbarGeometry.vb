' WordAi\Selection\WordSelectionToolbarGeometry.vb
' E4-01：选区工具条的落点计算。纯几何 + Win32 取点，不碰 Word COM、不碰窗体，便于隔离验证。

Imports System.Drawing
Imports System.Runtime.InteropServices

Namespace Selection

    ''' <summary>
    ''' 工具条位置与尺寸的唯一计算处。Word 的 Selection 没有可靠的屏幕坐标，
    ''' 所以以鼠标锚点为主（框选结束鼠标就停在选区末端），取不到时贴着 Word 窗口顶部居中，
    ''' 而不是各调用点各自换算 Range 坐标。
    ''' </summary>
    Public NotInheritable Class WordSelectionToolbarGeometry

        Public Const StripHeight As Integer = 28
        Public Const ItemWidth As Integer = 76
        Public Const ItemGap As Integer = 2
        Public Const EdgePadding As Integer = 6
        Public Const OffsetBelowAnchor As Integer = 6
        Public Const FallbackTopOffset As Integer = 44

        <StructLayout(LayoutKind.Sequential)>
        Private Structure NativePoint
            Public X As Integer
            Public Y As Integer
        End Structure

        <StructLayout(LayoutKind.Sequential)>
        Private Structure NativeRect
            Public Left As Integer
            Public Top As Integer
            Public Right As Integer
            Public Bottom As Integer
        End Structure

        <DllImport("user32.dll")>
        Private Shared Function GetCursorPos(ByRef point As NativePoint) As Boolean
        End Function

        <DllImport("user32.dll")>
        Private Shared Function GetWindowRect(hwnd As IntPtr, ByRef rect As NativeRect) As Boolean
        End Function

        <DllImport("user32.dll")>
        Private Shared Function GetForegroundWindow() As IntPtr
        End Function

        Private Sub New()
        End Sub

        ''' <summary>鼠标当前屏幕坐标；失败返回 False，调用方必须走降级而不是猜 (0,0)。</summary>
        Public Shared Function TryGetCursorPoint(ByRef point As Point) As Boolean
            Dim native As NativePoint
            If Not GetCursorPos(native) Then Return False
            point = New Point(native.X, native.Y)
            Return True
        End Function

        ''' <summary>
        ''' 当前前台窗口矩形。选区工具条只在用户刚在 Word 里完成框选时出现，此刻前台窗口就是 Word；
        ''' 取不到（焦点在面板或弹窗）时返回 False，由 ResolveLocation 的降级分支处理。
        ''' </summary>
        Public Shared Function TryGetForegroundWindowRect(ByRef rect As Rectangle) As Boolean
            Dim hwnd = GetForegroundWindow()
            If hwnd = IntPtr.Zero Then Return False
            Dim native As NativeRect
            If Not GetWindowRect(hwnd, native) Then Return False
            rect = New Rectangle(native.Left, native.Top, native.Right - native.Left, native.Bottom - native.Top)
            Return rect.Width > 0 AndAlso rect.Height > 0
        End Function

        Public Shared Function ComputeStripSize(itemCount As Integer) As Size
            Dim count = Math.Max(0, itemCount)
            Return New Size(EdgePadding * 2 + count * (ItemWidth + ItemGap), StripHeight)
        End Function

        ''' <summary>
        ''' 锚点在前台窗口内 → 出现在锚点下方居中，越界则翻到上方；
        ''' 锚点不可信（键盘选择、鼠标已移开）→ 贴着窗口顶部居中，位置可预期。
        ''' 结果一定落在锚点所在显示器的工作区内，多屏与任务栏都不会把工具条切掉。
        ''' </summary>
        Public Shared Function ResolveLocation(anchor As Point, ownerRect As Rectangle, stripSize As Size) As Rectangle
            Dim workArea = System.Windows.Forms.Screen.FromPoint(anchor).WorkingArea
            Return ResolveLocation(anchor, ownerRect, workArea, stripSize)
        End Function

        ''' <summary>工作区显式传入的重载：隔离验证用同一份逻辑，不依赖真实显示器。</summary>
        Public Shared Function ResolveLocation(anchor As Point, ownerRect As Rectangle, workArea As Rectangle, stripSize As Size) As Rectangle
            Dim left As Integer
            Dim top As Integer

            Dim anchorUsable = ownerRect.Width > 0 AndAlso ownerRect.Height > 0 AndAlso ownerRect.Contains(anchor)
            If anchorUsable Then
                left = anchor.X - stripSize.Width \ 2
                top = anchor.Y + OffsetBelowAnchor
                If top + stripSize.Height > workArea.Bottom Then
                    top = anchor.Y - OffsetBelowAnchor - stripSize.Height
                End If
            Else
                Dim centerX = If(ownerRect.Width > 0, ownerRect.Left + ownerRect.Width \ 2,
                                 workArea.Left + workArea.Width \ 2)
                Dim ownerTop = If(ownerRect.Height > 0, ownerRect.Top, workArea.Top)
                left = centerX - stripSize.Width \ 2
                top = ownerTop + FallbackTopOffset
            End If

            Return ClampToWorkArea(New Rectangle(left, top, stripSize.Width, stripSize.Height), workArea)
        End Function

        Private Shared Function ClampToWorkArea(rect As Rectangle, workArea As Rectangle) As Rectangle
            Dim left = rect.Left
            If rect.Width >= workArea.Width Then
                left = workArea.Left
            Else
                If left + rect.Width > workArea.Right Then left = workArea.Right - rect.Width
                If left < workArea.Left Then left = workArea.Left
            End If

            Dim top = rect.Top
            If top + rect.Height > workArea.Bottom Then top = workArea.Bottom - rect.Height
            If top < workArea.Top Then top = workArea.Top

            Return New Rectangle(left, top, rect.Width, rect.Height)
        End Function
    End Class

End Namespace
