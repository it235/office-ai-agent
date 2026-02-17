' ShareRibbon\Config\MemoryConfigForm.vb
' 记忆配置：rag_top_n、enable_user_profile 等

Imports System.Drawing
Imports System.Windows.Forms

''' <summary>
''' 记忆配置窗口：RAG 参数、用户画像、ContextBuilder 开关等
''' </summary>
Public Class MemoryConfigForm
    Inherits Form

    Private numRagTopN As NumericUpDown
    Private chkEnableUserProfile As CheckBox
    Private chkEnableAgenticSearch As CheckBox
    Private numAtomicMaxLen As NumericUpDown
    Private numSessionSummaryLimit As NumericUpDown
    Private chkUseContextBuilder As CheckBox

    Public Sub New()
        Me.Text = "记忆配置"
        Me.Size = New Size(420, 320)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.Font = New Font("Microsoft YaHei UI", 9)
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        AddHandler Me.FormClosing, AddressOf OnFormClosing
        InitializeUI()
        LoadValues()
    End Sub

    Private Sub OnFormClosing(sender As Object, e As FormClosingEventArgs)
        ' 关闭前移出共享 StatusStrip，避免随窗体被释放
        If Me.Controls.Contains(GlobalStatusStrip.StatusStrip) Then
            Me.Controls.Remove(GlobalStatusStrip.StatusStrip)
        End If
    End Sub

    Private Sub InitializeUI()
        Dim y As Integer = 20

        chkUseContextBuilder = New CheckBox() With {
            .Text = "使用 ContextBuilder（分层组装 Memory/Skills）",
            .Location = New Point(20, y),
            .Size = New Size(360, 24),
            .Checked = MemoryConfig.UseContextBuilder
        }
        Me.Controls.Add(chkUseContextBuilder)
        y += 32

        chkEnableUserProfile = New CheckBox() With {
            .Text = "启用用户画像",
            .Location = New Point(20, y),
            .Size = New Size(200, 24),
            .Checked = MemoryConfig.EnableUserProfile
        }
        Me.Controls.Add(chkEnableUserProfile)
        y += 28

        Dim lblRag As New Label() With {.Text = "RAG 检索条数 (1-20)：", .Location = New Point(20, y + 2), .Size = New Size(160, 20)}
        Me.Controls.Add(lblRag)
        numRagTopN = New NumericUpDown() With {
            .Location = New Point(185, y),
            .Size = New Size(60, 24),
            .Minimum = 1,
            .Maximum = 20,
            .Value = MemoryConfig.RagTopN
        }
        Me.Controls.Add(numRagTopN)
        y += 32

        Dim lblAtomic As New Label() With {.Text = "原子记忆最大长度 (50-500)：", .Location = New Point(20, y + 2), .Size = New Size(180, 20)}
        Me.Controls.Add(lblAtomic)
        numAtomicMaxLen = New NumericUpDown() With {
            .Location = New Point(205, y),
            .Size = New Size(60, 24),
            .Minimum = 50,
            .Maximum = 500,
            .Value = MemoryConfig.AtomicContentMaxLength
        }
        Me.Controls.Add(numAtomicMaxLen)
        y += 32

        Dim lblSummary As New Label() With {.Text = "近期会话摘要条数 (1-15)：", .Location = New Point(20, y + 2), .Size = New Size(180, 20)}
        Me.Controls.Add(lblSummary)
        numSessionSummaryLimit = New NumericUpDown() With {
            .Location = New Point(205, y),
            .Size = New Size(60, 24),
            .Minimum = 1,
            .Maximum = 15,
            .Value = MemoryConfig.SessionSummaryLimit
        }
        Me.Controls.Add(numSessionSummaryLimit)
        y += 32

        chkEnableAgenticSearch = New CheckBox() With {
            .Text = "启用 MCP 记忆搜索（Agentic Search）",
            .Location = New Point(20, y),
            .Size = New Size(280, 24),
            .Checked = MemoryConfig.EnableAgenticSearch
        }
        Me.Controls.Add(chkEnableAgenticSearch)
        y += 40

        Dim btnSave As New Button() With {.Text = "保存", .Location = New Point(20, y), .Size = New Size(80, 28)}
        AddHandler btnSave.Click, AddressOf BtnSaveClick
        Me.Controls.Add(btnSave)

        Dim btnClose As New Button() With {.Text = "关闭", .Location = New Point(320, y), .Size = New Size(80, 28)}
        AddHandler btnClose.Click, Sub(s, e) Me.Close()
        Me.Controls.Add(btnClose)

        Me.Controls.Add(GlobalStatusStrip.StatusStrip)
    End Sub

    Private Sub LoadValues()
        chkUseContextBuilder.Checked = MemoryConfig.UseContextBuilder
        chkEnableUserProfile.Checked = MemoryConfig.EnableUserProfile
        numRagTopN.Value = MemoryConfig.RagTopN
        numAtomicMaxLen.Value = MemoryConfig.AtomicContentMaxLength
        numSessionSummaryLimit.Value = MemoryConfig.SessionSummaryLimit
        chkEnableAgenticSearch.Checked = MemoryConfig.EnableAgenticSearch
    End Sub

    Private Sub BtnSaveClick(sender As Object, e As EventArgs)
        Try
            MemoryConfig.UseContextBuilder = chkUseContextBuilder.Checked
            MemoryConfig.EnableUserProfile = chkEnableUserProfile.Checked
            MemoryConfig.RagTopN = CInt(numRagTopN.Value)
            MemoryConfig.AtomicContentMaxLength = CInt(numAtomicMaxLen.Value)
            MemoryConfig.SessionSummaryLimit = CInt(numSessionSummaryLimit.Value)
            MemoryConfig.EnableAgenticSearch = chkEnableAgenticSearch.Checked
            GlobalStatusStrip.ShowInfo("记忆配置已保存")
        Catch ex As Exception
            GlobalStatusStrip.ShowWarning("保存失败: " & ex.Message)
        End Try
    End Sub
End Class
