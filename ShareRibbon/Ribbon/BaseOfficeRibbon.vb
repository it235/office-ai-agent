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
    '    InitializeComponent()  ' Designer 中定义的初始化
    '    InitializeBaseRibbon()  ' 基类中的通用初始化
    'End Sub

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        Dim apiConfig As New ConfigManager()
        apiConfig.LoadConfig()
        Dim promptConfig As New ConfigPromptForm()
        promptConfig.LoadConfig()
        InitializeBaseRibbon()
    End Sub

    Protected Overridable Sub InitializeBaseRibbon()
        ' 设置基础的事件处理程序
        'AddHandler ChatButton.Click, AddressOf ChatButton_Click
        'AddHandler PromptConfigButton.Click, AddressOf PromptConfigButton_Click
        'AddHandler ClearCacheButton.Click, AddressOf ClearCacheButton_Click
        'AddHandler AboutButton.Click, AddressOf AboutButton_Click
        'AddHandler DataAnalysisButton.Click, AddressOf DataAnalysisButton_Click
    End Sub

    ' 关于我按钮点击事件
    Private Sub AboutButton_Click_1(sender As Object, e As RibbonControlEventArgs) Handles AboutButton.Click
        MsgBox("大家好，我是B站的君哥，账号 君哥聊编程 。该插件的灵感是来自于一位B站的粉丝，他是银行审计相关的工作，经常与表格打交道，很多时候表格中的数据无法通过固定的公式来计算，但是在人类理解上又具有相同的意义，所以Excel AI诞生了。
插件在持续优化中，我本身与Excel打交道比较少，如果你有更多好的idea可以过来给我留言或评论，不断完善该插件。ExcelAi数据的默认存放目录在当前用户/文档/" + ConfigSettings.OfficeAiAppDataFolder + "下。")
    End Sub

    ' 清理缓存配置按钮点击事件
    Private Sub ClearCacheConfig_Click_1(sender As Object, e As RibbonControlEventArgs) Handles ClearCacheButton.Click
        ' 弹出确认框
        Dim result = MessageBox.Show("将删除文档\" & ConfigSettings.OfficeAiAppDataFolder & "目录下所有的配置，聊天记录信息，您确定要清理吗？", "确认操作", MessageBoxButtons.OKCancel, MessageBoxIcon.Question)
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
                MsgBox("缓存配置已清理！")
            Catch ex As Exception
                MsgBox("清理缓存配置时出错：" & ex.Message, vbCritical)
            End Try
        Else
            MsgBox("缓存目录不存在！")
        End If
    End Sub


    'Private Async Sub DataAnalysisButton_Click_1(sender As Object, e As RibbonControlEventArgs) Handles DataAnalysisButton.Click
    '    If String.IsNullOrWhiteSpace(ConfigSettings.ApiKey) Then
    '        MsgBox("请输入ApiKey！")
    '        Return
    '    End If

    '    If String.IsNullOrWhiteSpace(ConfigSettings.ApiUrl) Then
    '        MsgBox("请选择大模型！")
    '        Return
    '    End If

    '    ' 获取选中的单元格区域
    '    Dim selection As Excel.Range = TryCast(Globals.ThisAddIn.Application.Selection, Excel.Range)
    '    If selection IsNot Nothing Then
    '        Dim cellValues As New StringBuilder()

    '        Dim cellIndices As New StringBuilder()
    '        Dim cellList As New List(Of String)

    '        ' 按列遍历，每列用局部变量记录连续空行数
    '        For col As Integer = selection.Column To selection.Column + selection.Columns.Count
    '            Dim emptyCount As Integer = 0
    '            For row As Integer = selection.Row To selection.Row + selection.Rows.Count - 1
    '                Dim cell As Excel.Range = selection.Worksheet.Cells(row, col)
    '                ' 如果存在非空内容，则处理，并重置空计数
    '                If cell.Value IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(cell.Value.ToString()) Then
    '                    cellValues.AppendLine(cell.Value.ToString())
    '                    cellList.Add(cell.Address(False, False))
    '                    emptyCount = 0
    '                Else
    '                    emptyCount += 1
    '                    If emptyCount >= 50 Then
    '                        Exit For  ' 本列连续50行为空，退出当前列循环
    '                    End If
    '                End If
    '            Next
    '        Next


    '        ' 按照矩阵展开方式显示单元格索引
    '        Dim groupedCells = cellList.GroupBy(Function(c) Regex.Replace(c, "\d", ""))
    '        For Each group In groupedCells
    '            cellIndices.AppendLine(String.Join(",", group))
    '        Next

    '        ' 显示所有单元格的值
    '        If cellValues.Length > 0 Then
    '            Dim previewForm As New TextPreviewForm(cellIndices.ToString())
    '            previewForm.ShowDialog()

    '            If previewForm.IsConfirmed Then
    '                ' 获取查询内容和数据
    '                Dim question As String = cellValues.ToString
    '                question = previewForm.InputText & “。你只需要返回markdown格式的表格即可，别的什么都不要说，不要任何其他多余的文字。原始数据如下：“ & question

    '                Dim requestBody As String = CreateRequestBody(question)

    '                ' 发送 HTTP 请求并获取响应
    '                Dim response As String = Await SendHttpRequest(ConfigSettings.ApiUrl, ConfigSettings.ApiKey, requestBody)

    '                ' 如果响应为空，则终止执行
    '                If String.IsNullOrEmpty(response) Then
    '                    Return
    '                End If

    '                ' 解析并写入响应数据
    '                WriteResponseToSheet(response)
    '            End If
    '        Else
    '            MsgBox("选中的单元格无文本内容！")
    '        End If
    '    Else
    '        MsgBox("请选择一个单元格区域！")

    '    End If

    'End Sub

    ' 创建请求体
    Protected Function CreateRequestBody(question As String) As String
        Dim result As String = question.Replace("\", "\\").Replace("""", "\""").
                                  Replace(vbCr, "\r").Replace(vbLf, "\n").
                                  Replace(vbTab, "\t").Replace(vbBack, "\b").
                                  Replace(Chr(12), "\f")
        ' 使用从 ConfigSettings 中获取的模型名称
        Return "{""model"": """ & ConfigSettings.ModelName & """, ""messages"": [{""role"": ""user"", ""content"": """ & result & """}]}"
    End Function


    ' 发送 HTTP 请求
    Protected Async Function SendHttpRequest(apiUrl As String, apiKey As String, requestBody As String) As Task(Of String)
        Try
            ' 强制使用 TLS 1.2
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
            Dim handler As New HttpClientHandler()
            Using client As New HttpClient(handler)
                client.Timeout = TimeSpan.FromSeconds(120) ' 设置超时时间为 120 秒
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " & apiKey)
                Dim content As New StringContent(requestBody, Encoding.UTF8, "application/json")
                Dim response As HttpResponseMessage = Await client.PostAsync(apiUrl, content)
                response.EnsureSuccessStatusCode()
                Return Await response.Content.ReadAsStringAsync()
            End Using
        Catch ex As HttpRequestException
            MessageBox.Show("请求失败: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return String.Empty
        End Try
    End Function


    ' 点击Ribbon区的配置API按钮后触发
    Private Sub ConfigApiButton_Click(sender As Object, e As RibbonControlEventArgs) Handles ConfigApiButton.Click
        ' 创建并显示配置 API 的对话框
        Dim configForm As New ConfigApiForm()
        If configForm.ShowDialog() = DialogResult.OK Then
        End If
    End Sub
    Private Sub PromptConfigButton_Click(sender As Object, e As RibbonControlEventArgs) Handles PromptConfigButton.Click
        ' 创建并显示配置 API 的对话框
        Dim configForm As New ConfigPromptForm()
        If configForm.ShowDialog() = DialogResult.OK Then
        End If
    End Sub


    ' 定义 ComboBoxItem 类
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


    ' 共用的事件处理方法
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
            $"将删除文档\{ConfigSettings.OfficeAiAppDataFolder}目录下所有的配置，聊天记录信息，您确定要清理吗？",
            "确认操作",
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
                MessageBox.Show("缓存配置已清理！")
            Catch ex As Exception
                MessageBox.Show($"清理缓存配置时出错：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Protected Sub AboutButton_Click(sender As Object, e As RibbonControlEventArgs)
        MessageBox.Show(
            $"大家好，我是B站的君哥，账号 君哥聊编程。该插件的灵感是来自于一位B站的粉丝，他是银行审计相关的工作，经常与表格打交道，很多时候表格中的数据无法通过固定的公式来计算，但是在人类理解上又具有相同的意义，所以Excel AI诞生了。{vbCrLf}插件在持续优化中，我本身与Excel打交道比较少，如果你有更多好的idea可以过来给我留言或评论，不断完善该插件。ExcelAi数据的默认存放目录在当前用户/文档/{ConfigSettings.OfficeAiAppDataFolder}下。"
        )
    End Sub

    Protected MustOverride Sub ChatButton_Click(sender As Object, e As RibbonControlEventArgs) Handles ChatButton.Click
    Protected MustOverride Sub DataAnalysisButton_Click(sender As Object, e As RibbonControlEventArgs) Handles DataAnalysisButton.Click
    Protected MustOverride Function GetApplication() As Object
End Class