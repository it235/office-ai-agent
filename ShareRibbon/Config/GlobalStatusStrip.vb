Imports System.Windows.Forms
Imports System.Drawing

Public Module GlobalStatusStrip
    Public StatusStrip As New StatusStrip()
    Public ToolStripStatusLabel As New ToolStripStatusLabel()
    Private Timer As New Timer()

    Sub New()
        StatusStrip.Items.Add(ToolStripStatusLabel)
        StatusStrip.Visible = False
        Timer.Interval = 5000 ' ������ʾ��ʾʱ��Ϊ5��
        AddHandler Timer.Tick, AddressOf Timer_Tick
    End Sub

    Public Sub ShowWarning(message As String)
        ToolStripStatusLabel.Text = "���棺" & message
        ToolStripStatusLabel.ForeColor = Color.Red
        StatusStrip.Visible = True
        Timer.Start()
    End Sub
    Public Sub ShowInfo(message As String)
        ToolStripStatusLabel.Text = "��ʾ��" & message
        ToolStripStatusLabel.ForeColor = Color.Black
        StatusStrip.Visible = True
        Timer.Start()
    End Sub

    Private Sub Timer_Tick(sender As Object, e As EventArgs)
        ToolStripStatusLabel.Text = ""
        StatusStrip.Visible = False
        Timer.Stop()
    End Sub
End Module
