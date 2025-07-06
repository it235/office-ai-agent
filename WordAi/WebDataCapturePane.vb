Imports System.Diagnostics
Imports System.IO
Imports System.Net.Http
Imports System.Threading
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Word
Imports Microsoft.Vbe.Interop
Imports Newtonsoft.Json.Linq
Imports ShareRibbon
Public Class WebDataCapturePane
    Inherits BaseDataCapturePane

    Private isViewInitialized As Boolean = False
    Public Sub New()
        MyBase.New()
        ' 创建 ChatControl 实例
        ' 订阅AI聊天请求事件
        AddHandler AiChatRequested, AddressOf HandleAiChatRequest
        ' 直接调用异步初始化方法
        InitializeWebViewAsync()
    End Sub

    ' 新增：异步初始化方法
    ' 异步初始化方法
    Private Async Sub InitializeWebViewAsync()
        Try
            Debug.WriteLine("Starting WebView initialization from WebDataCapturePane")
            ' 调用基类的初始化方法
            Await InitializeWebView2()
        Catch ex As Exception
            MessageBox.Show($"初始化网页视图失败: {ex.Message}", "错误",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub HandleAiChatRequest(sender As Object, content As String)
        ' 显示聊天窗口
        Globals.ThisAddIn.ShowChatTaskPane()
        ' 添加选中的内容到引用区
        Globals.ThisAddIn.chatControl.AddSelectedContentItem(
                "来自网页",  ' 使用文档名称作为标识
                   content.Substring(0, Math.Min(content.Length, 50)) & If(content.Length > 50, "...", ""))
    End Sub

    ' 处理表格创建
    Protected Overrides Function CreateTable(tableData As TableData) As String
        Try
            ' 获取当前文档和选定范围
            Dim doc = Globals.ThisAddIn.Application.ActiveDocument
            Dim selection = doc.Application.Selection

            ' 创建表格
            Dim table = doc.Tables.Add(
                Range:=selection.Range,
                NumRows:=tableData.Rows,
                NumColumns:=tableData.Columns)

            ' 填充数据
            For i = 0 To tableData.Data.Count - 1
                For j = 0 To tableData.Data(i).Count - 1
                    table.Cell(i + 1, j + 1).Range.Text = tableData.Data(i)(j)
                Next
            Next

            ' 如果有表头，设置表头样式
            If tableData.Headers.Count > 0 Then
                table.Rows(1).HeadingFormat = True
                table.Rows(1).Range.Bold = True
            End If

            ' 设置表格样式
            table.Style = "网格型"
            table.AllowAutoFit = True
            table.AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)

            Return "[表格已插入]" & vbCrLf
        Catch ex As Exception
            MessageBox.Show($"创建表格时出错: {ex.Message}", "错误")
            Return String.Empty
        End Try
    End Function

    Protected Overrides Sub HandleExtractedContent(content As String)
        Try
            ' 获取活动文档
            Dim doc = Globals.ThisAddIn.Application.ActiveDocument
            If doc IsNot Nothing Then
                ' 在当前光标位置插入内容
                Dim selection = doc.Application.Selection
                If selection IsNot Nothing Then

                    ' 检查内容长度，决定插入策略
                    If content.Length > 32000 Then ' Word的TypeText方法通常限制在32K字符左右
                        ' 分块插入大内容
                        InsertLargeContent(selection, content)
                    Else
                        ' 直接插入小内容
                        selection.TypeText(content)
                    End If

                    ' 确保内容完全插入
                    Debug.WriteLine($"实际插入内容长度: {content.Length}")

                End If
            End If
        Catch ex As Exception
            MessageBox.Show($"处理提取内容时出错: {ex.Message}", "错误",
                      MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' 新增：分块插入大内容的方法
    Private Sub InsertLargeContent(selection As Microsoft.Office.Interop.Word.Selection, content As String)
        Try
            Const chunkSize As Integer = 30000 ' 每块30K字符
            Dim totalLength = content.Length
            Dim insertedLength = 0

            ' 显示进度
            Debug.WriteLine($"开始分块插入，总长度: {totalLength} 字符")

            For i = 0 To content.Length - 1 Step chunkSize
                Dim chunk = content.Substring(i, Math.Min(chunkSize, content.Length - i))

                Try
                    selection.TypeText(chunk)
                    insertedLength += chunk.Length

                    ' 刷新应用程序，防止界面卡死
                    System.Windows.Forms.Application.DoEvents()

                    Debug.WriteLine($"已插入: {insertedLength}/{totalLength} 字符")

                Catch ex As Exception
                    Debug.WriteLine($"插入第 {i / chunkSize + 1} 块时出错: {ex.Message}")

                    ' 如果TypeText失败，尝试使用Range.Text
                    Try
                        Dim range = selection.Range
                        range.Text = chunk
                        selection.Start = range.End
                        insertedLength += chunk.Length
                    Catch rangeEx As Exception
                        Debug.WriteLine($"Range.Text也失败: {rangeEx.Message}")

                        ' 最后尝试使用Insert方法
                        Try
                            selection.Range.InsertAfter(chunk)
                            selection.Start = selection.Range.End
                            insertedLength += chunk.Length
                        Catch insertEx As Exception
                            Debug.WriteLine($"InsertAfter也失败: {insertEx.Message}")
                            MessageBox.Show($"在第 {i / chunkSize + 1} 块时插入失败: {insertEx.Message}", "警告")
                            Exit For
                        End Try
                    End Try
                End Try
            Next

            Debug.WriteLine($"分块插入完成，实际插入: {insertedLength} 字符")

            If insertedLength < totalLength Then
                MessageBox.Show($"警告：只插入了 {insertedLength}/{totalLength} 字符", "部分插入")
            End If

        Catch ex As Exception
            MessageBox.Show($"分块插入失败: {ex.Message}", "错误")
        End Try
    End Sub

    ' 下载并插入图片
    Protected Overrides Sub DownloadAndInsertImage(src As String, alt As String)
        Try
            ' 获取活动文档
            Dim doc = Globals.ThisAddIn.Application.ActiveDocument
            If doc Is Nothing Then
                MessageBox.Show("请先打开一个Word文档", "提示")
                Return
            End If

            Dim selection = doc.Application.Selection
            If selection Is Nothing Then
                MessageBox.Show("无法获取当前选择位置", "错误")
                Return
            End If

            ' 如果是相对路径，转换为绝对路径
            Dim imageUrl As String = src
            If Not src.StartsWith("http") Then
                Dim baseUri As New Uri(ChatBrowser.CoreWebView2.Source)
                imageUrl = New Uri(baseUri, src).ToString()
            End If

            ' 使用ThreadPool进行异步下载
            ThreadPool.QueueUserWorkItem(
                Sub(state)
                    Try
                        Using client As New HttpClient()
                            client.Timeout = TimeSpan.FromSeconds(30)
                            Dim imageData = client.GetByteArrayAsync(imageUrl).Result

                            ' 创建临时文件
                            Dim tempPath = Path.GetTempFileName()
                            Dim extension = Path.GetExtension(New Uri(imageUrl).LocalPath)
                            If String.IsNullOrEmpty(extension) Then
                                extension = ".jpg" ' 默认扩展名
                            End If

                            Dim imagePath = Path.ChangeExtension(tempPath, extension)
                            File.WriteAllBytes(imagePath, imageData)

                            ' 在UI线程中插入图片
                            Me.Invoke(Sub()
                                          Try
                                              ' 插入图片
                                              Dim shape = selection.InlineShapes.AddPicture(
                                                  FileName:=imagePath,
                                                  LinkToFile:=False,
                                                  SaveWithDocument:=True)

                                              ' 设置图片属性
                                              With shape
                                                  .AlternativeText = alt
                                                  ' 限制最大宽度为400px
                                                  If .Width > 400 Then
                                                      .Width = 400
                                                  End If
                                              End With

                                              ' 添加图片说明
                                              selection.MoveRight()
                                              selection.TypeText(vbCrLf)
                                              selection.Font.Italic = True
                                              selection.Font.Size = 9
                                              selection.TypeText($"图片说明: {alt}")
                                              selection.Font.Italic = False
                                              selection.Font.Size = 11
                                              selection.TypeText(vbCrLf & vbCrLf)

                                              'MessageBox.Show("图片插入成功", "成功")

                                              ' 清理临时文件
                                              If File.Exists(imagePath) Then
                                                  File.Delete(imagePath)
                                              End If

                                          Catch ex As Exception
                                              MessageBox.Show($"插入图片失败: {ex.Message}", "错误")
                                          End Try
                                      End Sub)
                        End Using
                    Catch ex As Exception
                        Me.Invoke(Sub()
                                      MessageBox.Show($"下载图片失败: {ex.Message}", "错误")
                                  End Sub)
                    End Try
                End Sub)


        Catch ex As Exception
            MessageBox.Show($"处理图片时出错: {ex.Message}", "错误")
        End Try
    End Sub

    ' 处理视频内容
    Protected Overrides Sub HandleVideoContent(src As String, poster As String, duration As String, width As String, height As String)
        Try
            ' 获取活动文档
            Dim doc = Globals.ThisAddIn.Application.ActiveDocument
            If doc Is Nothing Then
                MessageBox.Show("请先打开一个Word文档", "提示")
                Return
            End If

            Dim selection = doc.Application.Selection
            If selection Is Nothing Then
                MessageBox.Show("无法获取当前选择位置", "错误")
                Return
            End If

            ' 如果是相对路径，转换为绝对路径
            Dim videoUrl As String = src
            If Not src.StartsWith("http") Then
                Dim baseUri As New Uri(ChatBrowser.CoreWebView2.Source)
                videoUrl = New Uri(baseUri, src).ToString()
            End If

            ' 插入视频信息文本
            selection.Font.Bold = True
            selection.Font.Color = RGB(0, 100, 200)
            selection.TypeText("🎬 视频内容")
            selection.Font.Bold = False
            selection.Font.Color = RGB(0, 0, 0)
            selection.TypeText(vbCrLf)

            ' 创建表格来展示视频信息
            Dim table = doc.Tables.Add(
            Range:=selection.Range,
            NumRows:=5,
            NumColumns:=2)

            With table
                .Style = "网格型"
                .AllowAutoFit = True

                ' 设置表头
                .Cell(1, 1).Range.Text = "属性"
                .Cell(1, 2).Range.Text = "值"
                .Rows(1).Range.Bold = True

                ' 填充数据
                .Cell(2, 1).Range.Text = "视频链接"
                .Cell(2, 2).Range.Text = videoUrl

                .Cell(3, 1).Range.Text = "时长"
                .Cell(3, 2).Range.Text = $"{duration} 秒"

                .Cell(4, 1).Range.Text = "尺寸"
                .Cell(4, 2).Range.Text = $"{width} × {height}"

                .Cell(5, 1).Range.Text = "预览图"
                .Cell(5, 2).Range.Text = If(String.IsNullOrEmpty(poster), "无", poster)

                ' 设置视频链接为超链接
                If Not String.IsNullOrEmpty(videoUrl) Then
                    doc.Hyperlinks.Add(
                    Anchor:= .Cell(2, 2).Range,
                    Address:=videoUrl,
                    TextToDisplay:="点击观看视频")
                End If
            End With

            ' 移动到表格后面
            selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 6)
            selection.TypeText(vbCrLf)

            ' 如果有预览图且不为空，尝试下载并插入
            If Not String.IsNullOrEmpty(poster) Then
                Dim posterUrl = poster
                If Not poster.StartsWith("http") Then
                    Dim baseUri As New Uri(ChatBrowser.CoreWebView2.Source)
                    posterUrl = New Uri(baseUri, poster).ToString()
                End If

                ' 异步下载预览图
                ThreadPool.QueueUserWorkItem(
                Sub(state)
                    Try
                        Using client As New HttpClient()
                            client.Timeout = TimeSpan.FromSeconds(30)
                            Dim imageData = client.GetByteArrayAsync(posterUrl).Result

                            Dim tempPath = Path.GetTempFileName()
                            Dim extension = Path.GetExtension(New Uri(posterUrl).LocalPath)
                            If String.IsNullOrEmpty(extension) Then
                                extension = ".jpg"
                            End If

                            Dim imagePath = Path.ChangeExtension(tempPath, extension)
                            File.WriteAllBytes(imagePath, imageData)

                            ' 在UI线程中插入预览图
                            ' 修复视频预览图插入
                            Me.Invoke(Sub()
                                          Try
                                              If Not File.Exists(imagePath) OrElse New FileInfo(imagePath).Length = 0 Then
                                                  Debug.WriteLine($"预览图文件无效: {imagePath}")
                                                  Return
                                              End If

                                              Dim shape As InlineShape = Nothing
                                              Try
                                                  shape = selection.InlineShapes.AddPicture(
                          FileName:=imagePath,
                          LinkToFile:=False,
                          SaveWithDocument:=True)
                                              Catch pictureEx As Exception
                                                  Debug.WriteLine($"插入预览图失败: {pictureEx.Message}")
                                                  Return
                                              End Try

                                              If shape IsNot Nothing Then
                                                  With shape
                                                      .AlternativeText = "视频预览图"
                                                      If .Width > 300 Then
                                                          .Width = 300
                                                      End If
                                                  End With

                                                  selection.MoveRight()
                                                  selection.TypeText(vbCrLf)
                                                  selection.Font.Italic = True
                                                  selection.Font.Size = 9
                                                  selection.TypeText("视频预览图")
                                                  selection.Font.Italic = False
                                                  selection.Font.Size = 11
                                                  selection.TypeText(vbCrLf & vbCrLf)
                                              End If

                                              If File.Exists(imagePath) Then
                                                  File.Delete(imagePath)
                                              End If

                                          Catch ex As Exception
                                              Debug.WriteLine($"插入预览图失败: {ex.Message}")
                                          End Try
                                      End Sub)
                        End Using
                    Catch ex As Exception
                        Debug.WriteLine($"下载预览图失败: {ex.Message}")
                    End Try
                End Sub)
            End If

            'MessageBox.Show("视频信息已插入", "成功")

        Catch ex As Exception
            MessageBox.Show($"处理视频内容时出错: {ex.Message}", "错误")
        End Try
    End Sub

    ' 处理音频内容
    Protected Overrides Sub HandleAudioContent(src As String, duration As String)
        Try
            ' 获取活动文档
            Dim doc = Globals.ThisAddIn.Application.ActiveDocument
            If doc Is Nothing Then
                MessageBox.Show("请先打开一个Word文档", "提示")
                Return
            End If

            Dim selection = doc.Application.Selection
            If selection Is Nothing Then
                MessageBox.Show("无法获取当前选择位置", "错误")
                Return
            End If

            ' 如果是相对路径，转换为绝对路径
            Dim audioUrl As String = src
            If Not src.StartsWith("http") Then
                Dim baseUri As New Uri(ChatBrowser.CoreWebView2.Source)
                audioUrl = New Uri(baseUri, src).ToString()
            End If

            ' 插入音频信息
            selection.Font.Bold = True
            selection.Font.Color = RGB(255, 140, 0)
            selection.TypeText("🎵 音频内容")
            selection.Font.Bold = False
            selection.Font.Color = RGB(0, 0, 0)
            selection.TypeText(vbCrLf)

            ' 创建简单的音频信息表格
            Dim table = doc.Tables.Add(
            Range:=selection.Range,
            NumRows:=3,
            NumColumns:=2)

            With table
                .Style = "网格型"
                .AllowAutoFit = True

                .Cell(1, 1).Range.Text = "属性"
                .Cell(1, 2).Range.Text = "值"
                .Rows(1).Range.Bold = True

                .Cell(2, 1).Range.Text = "音频链接"
                .Cell(2, 2).Range.Text = audioUrl

                .Cell(3, 1).Range.Text = "时长"
                .Cell(3, 2).Range.Text = $"{duration} 秒"

                ' 设置音频链接为超链接
                If Not String.IsNullOrEmpty(audioUrl) Then
                    doc.Hyperlinks.Add(
                    Anchor:= .Cell(2, 2).Range,
                    Address:=audioUrl,
                    TextToDisplay:="点击播放音频")
                End If
            End With

            ' 移动到表格后面
            selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, 4)
            selection.TypeText(vbCrLf)

            'MessageBox.Show("音频信息已插入", "成功")

        Catch ex As Exception
            MessageBox.Show($"处理音频内容时出错: {ex.Message}", "错误")
        End Try
    End Sub

    ' 处理媒体容器内容
    Protected Overrides Sub HandleMediaContainerContent(containedMedia As JArray, text As String)
        Try
            ' 获取活动文档
            Dim doc = Globals.ThisAddIn.Application.ActiveDocument
            If doc Is Nothing Then
                MessageBox.Show("请先打开一个Word文档", "提示")
                Return
            End If

            Dim selection = doc.Application.Selection
            If selection Is Nothing Then
                MessageBox.Show("无法获取当前选择位置", "错误")
                Return
            End If

            ' 插入容器标题
            selection.Font.Bold = True
            selection.Font.Color = RGB(233, 30, 99)
            selection.TypeText("📦 媒体容器内容")
            selection.Font.Bold = False
            selection.Font.Color = RGB(0, 0, 0)
            selection.TypeText(vbCrLf & vbCrLf)

            ' 如果有文本内容，先插入文本
            If Not String.IsNullOrWhiteSpace(text) Then
                selection.Font.Bold = True
                selection.TypeText("文本内容:")
                selection.Font.Bold = False
                selection.TypeText(vbCrLf)
                selection.TypeText(text.Trim())
                selection.TypeText(vbCrLf & vbCrLf)
            End If

            ' 处理包含的媒体元素
            If containedMedia IsNot Nothing AndAlso containedMedia.Count > 0 Then
                selection.Font.Bold = True
                selection.TypeText($"包含的媒体元素 ({containedMedia.Count} 个):")
                selection.Font.Bold = False
                selection.TypeText(vbCrLf & vbCrLf)

                ' 创建媒体信息表格
                Dim table = doc.Tables.Add(
                Range:=selection.Range,
                NumRows:=containedMedia.Count + 1,
                NumColumns:=5)

                With table
                    .Style = "网格型"
                    .AllowAutoFit = True

                    ' 设置表头
                    .Cell(1, 1).Range.Text = "类型"
                    .Cell(1, 2).Range.Text = "链接"
                    .Cell(1, 3).Range.Text = "描述"
                    .Cell(1, 4).Range.Text = "宽度"
                    .Cell(1, 5).Range.Text = "高度"
                    .Rows(1).Range.Bold = True

                    ' 填充媒体数据
                    For i = 0 To containedMedia.Count - 1
                        Dim media = DirectCast(containedMedia(i), JObject)
                        Dim mediaType = If(media("tag")?.ToString(), "")
                        Dim mediaSrc = If(media("src")?.ToString(), "")
                        Dim mediaAlt = If(media("alt")?.ToString(), "")
                        Dim mediaWidth = If(media("width")?.ToString(), "0")
                        Dim mediaHeight = If(media("height")?.ToString(), "0")


                        ' 如果是相对路径，转换为绝对路径
                        If Not String.IsNullOrEmpty(mediaSrc) AndAlso Not mediaSrc.StartsWith("http") Then
                            Try
                                Dim baseUri As New Uri(ChatBrowser.CoreWebView2.Source)
                                mediaSrc = New Uri(baseUri, mediaSrc).ToString()
                            Catch
                                ' 如果转换失败，保持原路径
                            End Try
                        End If

                        Dim rowIndex = i + 2
                        .Cell(rowIndex, 1).Range.Text = GetMediaTypeIcon(mediaType)
                        .Cell(rowIndex, 2).Range.Text = mediaSrc
                        .Cell(rowIndex, 3).Range.Text = mediaAlt
                        .Cell(rowIndex, 4).Range.Text = mediaWidth
                        .Cell(rowIndex, 5).Range.Text = mediaHeight

                        ' 为媒体链接添加超链接
                        If Not String.IsNullOrEmpty(mediaSrc) Then
                            Try
                                doc.Hyperlinks.Add(
                                Anchor:= .Cell(rowIndex, 2).Range,
                                Address:=mediaSrc,
                                TextToDisplay:="查看媒体")
                            Catch
                                ' 如果添加超链接失败，忽略错误
                            End Try
                        End If
                    Next
                End With

                ' 移动到表格后面
                selection.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine, containedMedia.Count + 2)
                selection.TypeText(vbCrLf)

                ' 询问是否下载图片
                Dim imageCount = containedMedia.Where(Function(m) m("tag")?.ToString() = "img").Count()
                If imageCount > 0 Then
                    Dim result = MessageBox.Show($"发现 {imageCount} 张图片，是否下载并插入到文档中？",
                                           "下载图片",
                                           MessageBoxButtons.YesNo,
                                           MessageBoxIcon.Question)

                    If result = DialogResult.Yes Then
                        DownloadContainerImages(containedMedia, selection)
                    End If
                End If
            End If

            'MessageBox.Show("媒体容器内容已插入", "成功")

        Catch ex As Exception
            MessageBox.Show($"处理媒体容器时出错: {ex.Message}", "错误")
        End Try
    End Sub

    ' 获取媒体类型图标
    Private Function GetMediaTypeIcon(mediaType As String) As String
        Select Case mediaType.ToLower()
            Case "img"
                Return "📷 图片"
            Case "video"
                Return "🎬 视频"
            Case "audio"
                Return "🎵 音频"
            Case Else
                Return "📄 媒体"
        End Select
    End Function

    ' 下载容器中的图片
    Private Sub DownloadContainerImages(containedMedia As JArray, selection As Microsoft.Office.Interop.Word.Selection)
        ThreadPool.QueueUserWorkItem(
        Sub(state)
            Dim imageCount = 0
            For Each mediaObj In containedMedia
                Dim media = DirectCast(mediaObj, JObject)
                If If(media("tag")?.ToString(), "") = "img" Then
                    Try
                        Dim src = If(media("src")?.ToString(), "")
                        Dim alt = If(media("alt")?.ToString(), "")

                        If Not String.IsNullOrEmpty(src) Then
                            ' 转换为绝对路径
                            Dim imageUrl = src
                            If Not src.StartsWith("http") Then
                                Dim baseUri As New Uri(ChatBrowser.CoreWebView2.Source)
                                imageUrl = New Uri(baseUri, src).ToString()
                            End If

                            Using client As New HttpClient()
                                client.Timeout = TimeSpan.FromSeconds(30)
                                Dim imageData = client.GetByteArrayAsync(imageUrl).Result

                                Dim tempPath = Path.GetTempFileName()
                                Dim extension = Path.GetExtension(New Uri(imageUrl).LocalPath)
                                If String.IsNullOrEmpty(extension) Then
                                    extension = ".jpg"
                                End If

                                Dim imagePath = Path.ChangeExtension(tempPath, extension)
                                File.WriteAllBytes(imagePath, imageData)

                                ' 在UI线程中插入图片 - 修复空引用异常
                                Me.Invoke(Sub()
                                              Try
                                                  ' 验证文件是否存在且有效
                                                  If Not File.Exists(imagePath) Then
                                                      Debug.WriteLine($"图片文件不存在: {imagePath}")
                                                      Return
                                                  End If

                                                  ' 验证文件大小
                                                  Dim fileInfo As New FileInfo(imagePath)
                                                  If fileInfo.Length = 0 Then
                                                      Debug.WriteLine($"图片文件为空: {imagePath}")
                                                      Return
                                                  End If

                                                  ' 插入图片并检查是否成功
                                                  Dim shape As InlineShape = Nothing
                                                  Try
                                                      shape = selection.InlineShapes.AddPicture(
                          FileName:=imagePath,
                          LinkToFile:=False,
                          SaveWithDocument:=True)
                                                  Catch pictureEx As Exception
                                                      Debug.WriteLine($"AddPicture失败: {pictureEx.Message}")
                                                      MessageBox.Show($"无法插入图片: {pictureEx.Message}", "错误")
                                                      Return
                                                  End Try

                                                  ' 检查shape是否创建成功
                                                  If shape IsNot Nothing Then
                                                      Try
                                                          With shape
                                                              .AlternativeText = alt
                                                              ' 限制最大宽度为400px
                                                              If .Width > 400 Then
                                                                  .Width = 400
                                                              End If
                                                          End With

                                                          ' 添加图片说明
                                                          selection.MoveRight()
                                                          selection.TypeText(vbCrLf)
                                                          selection.Font.Italic = True
                                                          selection.Font.Size = 9
                                                          selection.TypeText($"图片说明: {alt}")
                                                          selection.Font.Italic = False
                                                          selection.Font.Size = 11
                                                          selection.TypeText(vbCrLf & vbCrLf)

                                                          Debug.WriteLine("图片插入成功")
                                                      Catch shapeEx As Exception
                                                          Debug.WriteLine($"设置图片属性失败: {shapeEx.Message}")
                                                      End Try
                                                  Else
                                                      Debug.WriteLine("AddPicture返回了Nothing")
                                                      MessageBox.Show("图片插入失败：返回对象为空", "错误")
                                                  End If

                                                  ' 清理临时文件
                                                  If File.Exists(imagePath) Then
                                                      File.Delete(imagePath)
                                                  End If

                                              Catch ex As Exception
                                                  Debug.WriteLine($"插入图片失败: {ex.Message}")
                                                  MessageBox.Show($"插入图片失败: {ex.Message}", "错误")
                                              End Try
                                          End Sub)
                            End Using
                        End If
                    Catch ex As Exception
                        Debug.WriteLine($"下载图片失败: {ex.Message}")
                    End Try
                End If
            Next

            ' 显示完成消息
            Me.Invoke(Sub()
                          If imageCount > 0 Then
                              MessageBox.Show($"成功下载并插入 {imageCount} 张图片", "完成")
                          End If
                      End Sub)
        End Sub)
    End Sub

    ' 添加视图销毁处理
    Protected Overrides Sub OnHandleDestroyed(e As EventArgs)
        isViewInitialized = False
        MyBase.OnHandleDestroyed(e)
    End Sub

End Class