' ShareRibbon\Ribbon\BaseOfficeRibbon.vb
Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Microsoft.Office.Tools.Excel
Imports Microsoft.Office.Tools.Ribbon
Imports Newtonsoft.Json.Linq

Public MustInherit Class BaseOfficeRibbon
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    'Public Sub New(ByVal factory As Microsoft.Office.Tools.Ribbon.RibbonFactory)
    '    MyBase.New(factory)
    '    InitializeComponent()  ' Designer �ж���ĳ�ʼ��
    '    InitializeBaseRibbon()  ' �����е�ͨ�ó�ʼ��
    'End Sub

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        Dim apiConfig As New ConfigManager()
        apiConfig.LoadConfig()
        Dim promptConfig As New ConfigPromptForm()
        promptConfig.LoadConfig()
        InitializeBaseRibbon()
    End Sub

    Protected Overridable Sub InitializeBaseRibbon()
        ' ���û������¼��������
        'AddHandler ChatButton.Click, AddressOf ChatButton_Click
        'AddHandler PromptConfigButton.Click, AddressOf PromptConfigButton_Click
        'AddHandler ClearCacheButton.Click, AddressOf ClearCacheButton_Click
        'AddHandler AboutButton.Click, AddressOf AboutButton_Click
        'AddHandler DataAnalysisButton.Click, AddressOf DataAnalysisButton_Click
    End Sub

    ' �����Ұ�ť����¼�
    Private Sub AboutButton_Click_1(sender As Object, e As RibbonControlEventArgs) Handles AboutButton.Click
        MsgBox("��Һã�����Bվ�ľ��磬�˺� �����ı�� ���ò���������������һλBվ�ķ�˿���������������صĹ�������������򽻵����ܶ�ʱ�����е������޷�ͨ���̶��Ĺ�ʽ�����㣬����������������־�����ͬ�����壬����Excel AI�����ˡ�
����ڳ����Ż��У��ұ�����Excel�򽻵��Ƚ��٣�������и���õ�idea���Թ����������Ի����ۣ��������Ƹò����ExcelAi���ݵ�Ĭ�ϴ��Ŀ¼�ڵ�ǰ�û�/�ĵ�/" + ConfigSettings.OfficeAiAppDataFolder + "�¡�")
    End Sub

    ' ���������ð�ť����¼�
    Private Sub ClearCacheConfig_Click_1(sender As Object, e As RibbonControlEventArgs) Handles ClearCacheButton.Click
        ' ����ȷ�Ͽ�
        Dim result = MessageBox.Show("��ɾ���ĵ�\" & ConfigSettings.OfficeAiAppDataFolder & "Ŀ¼�����е����ã������¼��Ϣ����ȷ��Ҫ������", "ȷ�ϲ���", MessageBoxButtons.OKCancel, MessageBoxIcon.Question)
        If result <> DialogResult.OK Then
            Return
        End If

        Dim appDataPath As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & ConfigSettings.OfficeAiAppDataFolder
        If System.IO.Directory.Exists(appDataPath) Then
            Try
                Dim files As String() = System.IO.Directory.GetFiles(appDataPath)
                For Each file In files
                    System.IO.File.Delete(file)
                Next
                MsgBox("��������������")
            Catch ex As Exception
                MsgBox("����������ʱ����" & ex.Message, vbCritical)
            End Try
        Else
            MsgBox("����Ŀ¼�����ڣ�")
        End If
    End Sub


    'Private Async Sub DataAnalysisButton_Click_1(sender As Object, e As RibbonControlEventArgs) Handles DataAnalysisButton.Click
    '    If String.IsNullOrWhiteSpace(ConfigSettings.ApiKey) Then
    '        MsgBox("������ApiKey��")
    '        Return
    '    End If

    '    If String.IsNullOrWhiteSpace(ConfigSettings.ApiUrl) Then
    '        MsgBox("��ѡ���ģ�ͣ�")
    '        Return
    '    End If

    '    ' ��ȡѡ�еĵ�Ԫ������
    '    Dim selection As Excel.Range = TryCast(Globals.ThisAddIn.Application.Selection, Excel.Range)
    '    If selection IsNot Nothing Then
    '        Dim cellValues As New StringBuilder()

    '        Dim cellIndices As New StringBuilder()
    '        Dim cellList As New List(Of String)

    '        ' ���б�����ÿ���þֲ�������¼����������
    '        For col As Integer = selection.Column To selection.Column + selection.Columns.Count
    '            Dim emptyCount As Integer = 0
    '            For row As Integer = selection.Row To selection.Row + selection.Rows.Count - 1
    '                Dim cell As Excel.Range = selection.Worksheet.Cells(row, col)
    '                ' ������ڷǿ����ݣ����������ÿռ���
    '                If cell.Value IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(cell.Value.ToString()) Then
    '                    cellValues.AppendLine(cell.Value.ToString())
    '                    cellList.Add(cell.Address(False, False))
    '                    emptyCount = 0
    '                Else
    '                    emptyCount += 1
    '                    If emptyCount >= 50 Then
    '                        Exit For  ' ��������50��Ϊ�գ��˳���ǰ��ѭ��
    '                    End If
    '                End If
    '            Next
    '        Next


    '        ' ���վ���չ����ʽ��ʾ��Ԫ������
    '        Dim groupedCells = cellList.GroupBy(Function(c) Regex.Replace(c, "\d", ""))
    '        For Each group In groupedCells
    '            cellIndices.AppendLine(String.Join(",", group))
    '        Next

    '        ' ��ʾ���е�Ԫ���ֵ
    '        If cellValues.Length > 0 Then
    '            Dim previewForm As New TextPreviewForm(cellIndices.ToString())
    '            previewForm.ShowDialog()

    '            If previewForm.IsConfirmed Then
    '                ' ��ȡ��ѯ���ݺ�����
    '                Dim question As String = cellValues.ToString
    '                question = previewForm.InputText & ������ֻ��Ҫ����markdown��ʽ�ı�񼴿ɣ����ʲô����Ҫ˵����Ҫ�κ�������������֡�ԭʼ�������£��� & question

    '                Dim requestBody As String = CreateRequestBody(question)

    '                ' ���� HTTP ���󲢻�ȡ��Ӧ
    '                Dim response As String = Await SendHttpRequest(ConfigSettings.ApiUrl, ConfigSettings.ApiKey, requestBody)

    '                ' �����ӦΪ�գ�����ִֹ��
    '                If String.IsNullOrEmpty(response) Then
    '                    Return
    '                End If

    '                ' ������д����Ӧ����
    '                WriteResponseToSheet(response)
    '            End If
    '        Else
    '            MsgBox("ѡ�еĵ�Ԫ�����ı����ݣ�")
    '        End If
    '    Else
    '        MsgBox("��ѡ��һ����Ԫ������")

    '    End If

    'End Sub

    ' ����������
    Protected Function CreateRequestBody(question As String) As String
        Dim result As String = question.Replace("\", "\\").Replace("""", "\""").
                                  Replace(vbCr, "\r").Replace(vbLf, "\n").
                                  Replace(vbTab, "\t").Replace(vbBack, "\b").
                                  Replace(Chr(12), "\f")
        ' ʹ�ô� ConfigSettings �л�ȡ��ģ������
        Return "{""model"": """ & ConfigSettings.ModelName & """, ""messages"": [{""role"": ""user"", ""content"": """ & result & """}]}"
    End Function


    ' ���� HTTP ����
    Protected Async Function SendHttpRequest(apiUrl As String, apiKey As String, requestBody As String) As Task(Of String)
        Try
            ' ǿ��ʹ�� TLS 1.2
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
            Dim handler As New HttpClientHandler()
            Using client As New HttpClient(handler)
                client.Timeout = TimeSpan.FromSeconds(120) ' ���ó�ʱʱ��Ϊ 120 ��
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " & apiKey)
                Dim content As New StringContent(requestBody, Encoding.UTF8, "application/json")
                Dim response As HttpResponseMessage = Await client.PostAsync(apiUrl, content)
                response.EnsureSuccessStatusCode()
                Return Await response.Content.ReadAsStringAsync()
            End Using
        Catch ex As HttpRequestException
            MessageBox.Show("����ʧ��: " & ex.Message, "����", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return String.Empty
        End Try
    End Function


    ' ���Ribbon��������API��ť�󴥷�
    Private Sub ConfigApiButton_Click(sender As Object, e As RibbonControlEventArgs) Handles ConfigApiButton.Click
        ' ��������ʾ���� API �ĶԻ���
        Dim configForm As New ConfigApiForm()
        If configForm.ShowDialog() = DialogResult.OK Then
        End If
    End Sub
    Private Sub PromptConfigButton_Click(sender As Object, e As RibbonControlEventArgs) Handles PromptConfigButton.Click
        ' ��������ʾ���� API �ĶԻ���
        Dim configForm As New ConfigPromptForm()
        If configForm.ShowDialog() = DialogResult.OK Then
        End If
    End Sub


    ' ���� ComboBoxItem ��
    Private Class ComboBoxItem
        Public Property Text As String
        Public Property Value As String

        Public Sub New(text As String, value As String)
            Me.Text = text
            Me.Value = value
        End Sub

        Public Overrides Function ToString() As String
            Return Text
        End Function
    End Class


    ' ���õ��¼�������
    'Protected Sub ConfigApiButton_Click(sender As Object, e As RibbonControlEventArgs)
    '    Using configForm As New ConfigApiForm()
    '        configForm.ShowDialog()
    '    End Using
    'End Sub

    'Protected Sub PromptConfigButton_Click(sender As Object, e As RibbonControlEventArgs)
    '    Using configForm As New ConfigPromptForm()
    '        configForm.ShowDialog()
    '    End Using
    'End Sub

    Protected Sub ClearCacheButton_Click(sender As Object, e As RibbonControlEventArgs)
        If MessageBox.Show(
            $"��ɾ���ĵ�\{ConfigSettings.OfficeAiAppDataFolder}Ŀ¼�����е����ã������¼��Ϣ����ȷ��Ҫ������",
            "ȷ�ϲ���",
            MessageBoxButtons.OKCancel,
            MessageBoxIcon.Question) <> DialogResult.OK Then
            Return
        End If

        Dim appDataPath As String = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            ConfigSettings.OfficeAiAppDataFolder)

        If Directory.Exists(appDataPath) Then
            Try
                For Each file In Directory.GetFiles(appDataPath)
                    'file.Delete(file)
                Next
                MessageBox.Show("��������������")
            Catch ex As Exception
                MessageBox.Show($"����������ʱ����{ex.Message}", "����", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Protected Sub AboutButton_Click(sender As Object, e As RibbonControlEventArgs)
        MessageBox.Show(
            $"��Һã�����Bվ�ľ��磬�˺� �����ı�̡��ò���������������һλBվ�ķ�˿���������������صĹ�������������򽻵����ܶ�ʱ�����е������޷�ͨ���̶��Ĺ�ʽ�����㣬����������������־�����ͬ�����壬����Excel AI�����ˡ�{vbCrLf}����ڳ����Ż��У��ұ�����Excel�򽻵��Ƚ��٣�������и���õ�idea���Թ����������Ի����ۣ��������Ƹò����ExcelAi���ݵ�Ĭ�ϴ��Ŀ¼�ڵ�ǰ�û�/�ĵ�/{ConfigSettings.OfficeAiAppDataFolder}�¡�"
        )
    End Sub

    Protected MustOverride Sub ChatButton_Click(sender As Object, e As RibbonControlEventArgs) Handles ChatButton.Click
    Protected MustOverride Sub DataAnalysisButton_Click(sender As Object, e As RibbonControlEventArgs) Handles DataAnalysisButton.Click
    Protected MustOverride Function GetApplication() As Object
End Class