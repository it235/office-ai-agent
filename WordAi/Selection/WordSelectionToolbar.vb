' WordAi\Selection\WordSelectionToolbar.vb
' E4-01：选中文字后贴在选区下方的一条不抢焦点的快捷动作条。
' 只做"显示动作 + 把动作交出去"，不读文档、不执行任何操作，执行由控制器交回 Agent 主链。

Imports System.Drawing
Imports System.Windows.Forms
Imports ShareRibbon.Agent

Namespace Selection

    ''' <summary>
    ''' 无边框、不激活的顶层动作条。用 ShowWithoutActivation + WS_EX_NOACTIVATE 是因为
    ''' 一旦这个窗口抢了焦点，Word 的选区就会被销毁，动作也就失去了作用对象。
    ''' </summary>
    Public Class WordSelectionToolbar
        Inherits Form

        Public Event ActionChosen(action As QuickAction)

        Private Const WS_EX_NOACTIVATE As Integer = &H8000
        Private Const WS_EX_TOOLWINDOW As Integer = &H80

        Private ReadOnly _backColorNormal As Color = Color.FromArgb(33, 37, 45)
        Private ReadOnly _backColorHover As Color = Color.FromArgb(51, 65, 92)

        Private ReadOnly _buttons As New List(Of Button)
        Private ReadOnly _tips As ToolTip
        Private _builtActionIds As String = ""

        Public Sub New()
            FormBorderStyle = FormBorderStyle.None
            ShowInTaskbar = False
            StartPosition = FormStartPosition.Manual
            TopMost = True
            AutoScaleMode = AutoScaleMode.None
            BackColor = _backColorNormal
            Font = New Font("Microsoft YaHei UI", 9.0F, FontStyle.Regular)
            _tips = New ToolTip With {.AutomaticDelay = 500}
        End Sub

        Protected Overrides ReadOnly Property ShowWithoutActivation As Boolean
            Get
                Return True
            End Get
        End Property

        Protected Overrides ReadOnly Property CreateParams As CreateParams
            Get
                Dim cp = MyBase.CreateParams
                cp.ExStyle = cp.ExStyle Or WS_EX_NOACTIVATE Or WS_EX_TOOLWINDOW
                Return cp
            End Get
        End Property

        ''' <summary>
        ''' 就地更新动作条：动作集没变时不重建按钮，避免每次防抖都闪一下、也避免重复绑定 Click。
        ''' </summary>
        Public Sub ShowAt(actions As IReadOnlyList(Of QuickAction), bounds As Rectangle)
            Dim signature = BuildSignature(actions)
            If Not String.Equals(signature, _builtActionIds, StringComparison.Ordinal) Then
                RebuildButtons(actions)
                _builtActionIds = signature
            End If

            ClientSize = bounds.Size
            Location = bounds.Location
            If Not Visible Then Show()
        End Sub

        Public Sub HideStrip()
            If Visible Then Hide()
        End Sub

        Protected Overrides Sub OnPaint(e As PaintEventArgs)
            MyBase.OnPaint(e)
            Using pen As New Pen(Color.FromArgb(&HA0, &HA6, &HB2), 1)
                e.Graphics.DrawRectangle(pen, 0, 0, Width - 1, Height - 1)
            End Using
        End Sub

        Protected Overrides Sub Dispose(disposing As Boolean)
            If disposing Then
                _tips.Dispose()
                For Each button In _buttons
                    button.Dispose()
                Next
                _buttons.Clear()
            End If
            MyBase.Dispose(disposing)
        End Sub

        Private Shared Function BuildSignature(actions As IReadOnlyList(Of QuickAction)) As String
            Dim parts = New List(Of String)
            For Each action In actions
                parts.Add(action.Id)
            Next
            Return String.Join("|", parts)
        End Function

        Private Sub RebuildButtons(actions As IReadOnlyList(Of QuickAction))
            For Each button In _buttons
                RemoveHandler button.Click, AddressOf OnActionButtonClick
                Controls.Remove(button)
                button.Dispose()
            Next
            _buttons.Clear()

            Dim left = WordSelectionToolbarGeometry.EdgePadding
            For Each action In actions
                Dim button As New Button With {
                    .Text = action.Label,
                    .Tag = action,
                    .FlatStyle = FlatStyle.Flat,
                    .ForeColor = Color.White,
                    .BackColor = _backColorNormal,
                    .Size = New Size(WordSelectionToolbarGeometry.ItemWidth,
                                     WordSelectionToolbarGeometry.StripHeight - 6),
                    .Location = New Point(left, 3),
                    .TabStop = False,
                    .UseCompatibleTextRendering = False
                }
                button.FlatAppearance.BorderSize = 0
                button.FlatAppearance.MouseOverBackColor = _backColorHover
                AddHandler button.Click, AddressOf OnActionButtonClick
                _tips.SetToolTip(button, action.RequestText)
                Controls.Add(button)
                _buttons.Add(button)
                left += WordSelectionToolbarGeometry.ItemWidth + WordSelectionToolbarGeometry.ItemGap
            Next
        End Sub

        Private Sub OnActionButtonClick(sender As Object, e As EventArgs)
            Dim button = TryCast(sender, Button)
            Dim action = TryCast(button?.Tag, QuickAction)
            If action Is Nothing Then Return
            HideStrip()
            RaiseEvent ActionChosen(action)
        End Sub
    End Class

End Namespace
