' ShareRibbon\Config\MemoryManageForm.vb
' 原子记忆、用户画像的 CRUD 管理

Imports System.Drawing
Imports System.Windows.Forms

''' <summary>
''' 记忆管理窗口：查看/编辑/删除原子记忆与用户画像
''' </summary>
Public Class MemoryManageForm
    Inherits Form

    Private tabControl As TabControl
    Private tabAtomic As TabPage
    Private tabProfile As TabPage
    Private splitAtomic As SplitContainer
    Private listAtomic As ListBox
    Private txtAtomicContent As TextBox
    Private txtUserProfile As TextBox
    Private _atomicRecords As New List(Of AtomicMemoryRecord)()

    Public Sub New()
        Me.Text = "记忆管理"
        Me.Size = New Size(600, 450)
        Me.MinimumSize = New Size(450, 380)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.Font = New Font("Microsoft YaHei UI", 9)
        AddHandler Me.FormClosing, AddressOf OnFormClosing
        AddHandler Me.Load, AddressOf OnFormLoad
        InitializeUI()
        LoadAtomicMemories()
        LoadUserProfile()
    End Sub

    Private Sub OnFormLoad(sender As Object, e As EventArgs)
        If splitAtomic Is Nothing Then Return
        Try
            splitAtomic.Panel1MinSize = 80
            splitAtomic.Panel2MinSize = 150
            Dim maxDist = splitAtomic.Width - splitAtomic.Panel2MinSize - splitAtomic.SplitterWidth
            If maxDist >= splitAtomic.Panel1MinSize Then
                splitAtomic.SplitterDistance = Math.Max(splitAtomic.Panel1MinSize, Math.Min(220, maxDist))
            End If
        Catch
        End Try
    End Sub

    Private Sub OnFormClosing(sender As Object, e As FormClosingEventArgs)
        If Me.Controls.Contains(GlobalStatusStrip.StatusStrip) Then
            Me.Controls.Remove(GlobalStatusStrip.StatusStrip)
        End If
    End Sub

    Private Sub InitializeUI()
        tabControl = New TabControl() With {
            .Location = New Point(10, 10),
            .Size = New Size(570, 390),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        }

        ' Tab 1: 原子记忆（可拖拽分隔条调整列表宽度）
        tabAtomic = New TabPage("原子记忆")
        splitAtomic = New SplitContainer() With {
            .Dock = DockStyle.Fill,
            .Panel1MinSize = 0,
            .Panel2MinSize = 0,
            .FixedPanel = FixedPanel.None
        }
        Dim lblList As New Label() With {.Text = "原子记忆列表（最近 100 条）", .Location = New Point(0, 0), .Size = New Size(200, 20), .Anchor = AnchorStyles.Top Or AnchorStyles.Left}
        splitAtomic.Panel1.Controls.Add(lblList)
        listAtomic = New ListBox() With {
            .Location = New Point(0, 22),
            .Size = New Size(216, 350),
            .DisplayMember = "DisplayText",
            .Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right,
            .HorizontalScrollbar = True,
            .HorizontalExtent = 3000
        }
        AddHandler listAtomic.SelectedIndexChanged, AddressOf AtomicSelectionChanged
        splitAtomic.Panel1.Controls.Add(listAtomic)
        txtAtomicContent = New TextBox() With {
            .Location = New Point(5, 5),
            .Size = New Size(340, 300),
            .Multiline = True,
            .ScrollBars = ScrollBars.Both,
            .ReadOnly = True,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        }
        splitAtomic.Panel2.Controls.Add(txtAtomicContent)
        Dim btnRefreshAtomic As New Button() With {.Text = "刷新", .Location = New Point(5, 330), .Size = New Size(70, 28), .Anchor = AnchorStyles.Bottom Or AnchorStyles.Left}
        AddHandler btnRefreshAtomic.Click, AddressOf LoadAtomicMemories
        splitAtomic.Panel2.Controls.Add(btnRefreshAtomic)
        Dim btnDeleteAtomic As New Button() With {.Text = "删除选中", .Location = New Point(85, 330), .Size = New Size(80, 28), .Anchor = AnchorStyles.Bottom Or AnchorStyles.Left}
        AddHandler btnDeleteAtomic.Click, AddressOf BtnDeleteAtomicClick
        splitAtomic.Panel2.Controls.Add(btnDeleteAtomic)
        Dim btnCopyAtomic As New Button() With {.Text = "复制选中", .Location = New Point(175, 330), .Size = New Size(70, 28), .Anchor = AnchorStyles.Bottom Or AnchorStyles.Left}
        AddHandler btnCopyAtomic.Click, Sub(s, ev)
                                           If listAtomic.SelectedItem IsNot Nothing Then
                                               Try
                                                   Clipboard.SetText(listAtomic.SelectedItem.ToString())
                                                   GlobalStatusStrip.ShowInfo("已复制")
                                               Catch ex As Exception
                                                   GlobalStatusStrip.ShowWarning("复制失败: " & ex.Message)
                                               End Try
                                           Else
                                               GlobalStatusStrip.ShowWarning("请先选择一项")
                                           End If
                                       End Sub
        splitAtomic.Panel2.Controls.Add(btnCopyAtomic)
        tabAtomic.Controls.Add(splitAtomic)
        tabControl.TabPages.Add(tabAtomic)

        ' Tab 2: 用户画像
        tabProfile = New TabPage("用户画像")
        Dim lblProfile As New Label() With {.Text = "用户画像内容（用于 [1] 层注入）", .Location = New Point(10, 10), .Size = New Size(250, 20), .Anchor = AnchorStyles.Top Or AnchorStyles.Left}
        tabProfile.Controls.Add(lblProfile)
        txtUserProfile = New TextBox() With {
            .Location = New Point(10, 35),
            .Size = New Size(540, 250),
            .Multiline = True,
            .ScrollBars = ScrollBars.Vertical,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        }
        tabProfile.Controls.Add(txtUserProfile)
        Dim btnSaveProfile As New Button() With {.Text = "保存", .Location = New Point(10, 295), .Size = New Size(80, 28), .Anchor = AnchorStyles.Bottom Or AnchorStyles.Left}
        AddHandler btnSaveProfile.Click, AddressOf BtnSaveProfileClick
        tabProfile.Controls.Add(btnSaveProfile)
        Dim btnClearProfile As New Button() With {.Text = "清空", .Location = New Point(100, 295), .Size = New Size(70, 28), .Anchor = AnchorStyles.Bottom Or AnchorStyles.Left}
        AddHandler btnClearProfile.Click, AddressOf BtnClearProfileClick
        tabProfile.Controls.Add(btnClearProfile)
        tabControl.TabPages.Add(tabProfile)

        Me.Controls.Add(tabControl)

        Dim btnClose As New Button() With {.Text = "关闭", .Location = New Point(500, 410), .Size = New Size(80, 28), .Anchor = AnchorStyles.Bottom Or AnchorStyles.Right}
        AddHandler btnClose.Click, Sub(s, e) Me.Close()
        Me.Controls.Add(btnClose)

        Me.Controls.Add(GlobalStatusStrip.StatusStrip)
    End Sub

    Private Sub LoadAtomicMemories(sender As Object, e As EventArgs)
        LoadAtomicMemories()
    End Sub

    Private Sub LoadAtomicMemories()
        Try
            OfficeAiDatabase.EnsureInitialized()
            _atomicRecords = MemoryRepository.ListAtomicMemories(100, 0)
            listAtomic.DataSource = Nothing
            listAtomic.Items.Clear()
            For Each r In _atomicRecords
                Dim preview = If(r.Content?.Length > 40, r.Content.Substring(0, 40) & "...", r.Content)
                listAtomic.Items.Add(New AtomicItem With {.Record = r, .DisplayText = $"[{r.CreateTime}] {preview}"})
            Next
        Catch ex As Exception
            listAtomic.Items.Clear()
            listAtomic.Items.Add("(加载失败: " & ex.Message & ")")
            GlobalStatusStrip.ShowWarning("加载失败")
        End Try
    End Sub

    Private Sub AtomicSelectionChanged(sender As Object, e As EventArgs)
        Dim item = TryCast(listAtomic.SelectedItem, AtomicItem)
        If item Is Nothing Then
            txtAtomicContent.Text = ""
            Return
        End If
        txtAtomicContent.Text = item.Record.Content
    End Sub

    Private Sub BtnDeleteAtomicClick(sender As Object, e As EventArgs)
        Dim item = TryCast(listAtomic.SelectedItem, AtomicItem)
        If item Is Nothing Then
            GlobalStatusStrip.ShowWarning("请先选择一条记录")
            Return
        End If
        If MessageBox.Show("确定删除此条原子记忆？", "确认", MessageBoxButtons.YesNo) <> DialogResult.Yes Then Return
        Try
            MemoryRepository.DeleteAtomicMemory(item.Record.Id)
            GlobalStatusStrip.ShowInfo("已删除")
            LoadAtomicMemories()
        Catch ex As Exception
            GlobalStatusStrip.ShowWarning("删除失败: " & ex.Message)
        End Try
    End Sub

    Private Sub LoadUserProfile()
        Try
            txtUserProfile.Text = MemoryRepository.GetUserProfile()
        Catch
            txtUserProfile.Text = ""
        End Try
    End Sub

    Private Sub BtnSaveProfileClick(sender As Object, e As EventArgs)
        Try
            MemoryRepository.UpdateUserProfile(txtUserProfile.Text)
            GlobalStatusStrip.ShowInfo("用户画像已保存")
        Catch ex As Exception
            GlobalStatusStrip.ShowWarning("保存失败: " & ex.Message)
        End Try
    End Sub

    Private Sub BtnClearProfileClick(sender As Object, e As EventArgs)
        If MessageBox.Show("确定清空用户画像？", "确认", MessageBoxButtons.YesNo) <> DialogResult.Yes Then Return
        Try
            MemoryRepository.UpdateUserProfile("")
            txtUserProfile.Text = ""
            GlobalStatusStrip.ShowInfo("已清空")
        Catch ex As Exception
            GlobalStatusStrip.ShowWarning("清空失败: " & ex.Message)
        End Try
    End Sub

    Private Class AtomicItem
        Public Property Record As AtomicMemoryRecord
        Public Property DisplayText As String
        Public Overrides Function ToString() As String
            Return DisplayText
        End Function
    End Class
End Class
