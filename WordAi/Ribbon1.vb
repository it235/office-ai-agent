' WordAi\Ribbon1.vb
Imports System.Diagnostics
Imports System.Windows.Forms
Imports Microsoft.Office.Tools.Ribbon
Imports ShareRibbon  ' ��Ӵ�����

Public Class Ribbon1
    Inherits BaseOfficeRibbon

    Protected Overrides Async Sub ChatButton_Click(sender As Object, e As RibbonControlEventArgs)
        Globals.ThisAddIn.ShowChatTaskPane()
    End Sub

    Protected Overrides Sub DataAnalysisButton_Click(sender As Object, e As RibbonControlEventArgs)
        ' Word �ض������ݷ����߼�
        MessageBox.Show("Word���ݷ����������ڿ�����...")
    End Sub

    Protected Overrides Function GetApplication() As Object
        Return Globals.ThisAddIn.Application
    End Function
End Class