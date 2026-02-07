' ShareRibbon\Controls\Services\HistoryService.vb
' 历史文件服务：处理聊天历史文件的获取、打开和管理

Imports System.IO
Imports System.Web
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

''' <summary>
''' 历史文件服务，负责聊天历史文件的管理操作
''' </summary>
Public Class HistoryService

    Private ReadOnly _executeScript As Func(Of String, Threading.Tasks.Task)

    ''' <summary>
    ''' 构造函数
    ''' </summary>
    ''' <param name="executeScript">执行 JavaScript 的委托</param>
    Public Sub New(executeScript As Func(Of String, Threading.Tasks.Task))
        _executeScript = executeScript
    End Sub

    ''' <summary>
    ''' 获取历史文件目录路径
    ''' </summary>
    Public Shared Function GetHistoryDirectory() As String
        Return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            ConfigSettings.OfficeAiAppDataFolder
        )
    End Function

    ''' <summary>
    ''' 获取历史文件列表并发送到前端
    ''' </summary>
    Public Sub GetHistoryFiles()
        Try
            Dim historyDir As String = GetHistoryDirectory()
            Dim historyFiles As New List(Of Object)()

            If Directory.Exists(historyDir) Then
                Dim files As String() = Directory.GetFiles(historyDir, "saved_chat_*.html")

                For Each filePath As String In files
                    Try
                        Dim fileInfo As New FileInfo(filePath)
                        historyFiles.Add(New With {
                            .fileName = fileInfo.Name,
                            .fullPath = fileInfo.FullName,
                            .size = fileInfo.Length,
                            .lastModified = fileInfo.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
                        })
                    Catch ex As Exception
                        Debug.WriteLine($"处理文件信息时出错: {filePath} - {ex.Message}")
                    End Try
                Next
            End If

            Dim jsonResult As String = JsonConvert.SerializeObject(historyFiles)
            Dim js As String = $"setHistoryFilesList({jsonResult});"
            _executeScript(js)

        Catch ex As Exception
            Debug.WriteLine($"获取历史文件列表时出错: {ex.Message}")
            _executeScript("setHistoryFilesList([]);")
        End Try
    End Sub

    ''' <summary>
    ''' 打开历史文件
    ''' </summary>
    Public Sub OpenHistoryFile(jsonDoc As JObject)
        Try
            Dim filePath As String = jsonDoc("filePath").ToString()

            If File.Exists(filePath) Then
                Process.Start(New ProcessStartInfo() With {
                    .FileName = filePath,
                    .UseShellExecute = True
                })
                GlobalStatusStrip.ShowInfo("已在浏览器中打开历史记录")
            Else
                GlobalStatusStrip.ShowWarning("历史记录文件不存在")
            End If

        Catch ex As Exception
            Debug.WriteLine($"打开历史文件时出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning("打开历史记录失败: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' 删除历史文件
    ''' </summary>
    Public Shared Function DeleteHistoryFile(filePath As String) As Boolean
        Try
            If File.Exists(filePath) Then
                File.Delete(filePath)
                Return True
            End If
            Return False
        Catch ex As Exception
            Debug.WriteLine($"删除历史文件时出错: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 清理所有历史文件
    ''' </summary>
    Public Shared Function ClearAllHistoryFiles() As Integer
        Try
            Dim historyDir As String = GetHistoryDirectory()
            Dim deletedCount As Integer = 0

            If Directory.Exists(historyDir) Then
                Dim files As String() = Directory.GetFiles(historyDir, "saved_chat_*.html")
                For Each filePath As String In files
                    Try
                        File.Delete(filePath)
                        deletedCount += 1
                    Catch
                    End Try
                Next
            End If

            Return deletedCount
        Catch ex As Exception
            Debug.WriteLine($"清理历史文件时出错: {ex.Message}")
            Return 0
        End Try
    End Function

End Class
