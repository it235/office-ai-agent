Imports System.IO

Public Module SimpleLogger
    Private ReadOnly LogFile As String = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "ExcelAi\logs\app.log"
    )

    Public Sub LogInfo(msg As String)
        WriteLog("INFO", msg)
    End Sub

    Public Sub LogError(msg As String, Optional ex As Exception = Nothing)
        WriteLog("ERROR", msg & If(ex IsNot Nothing, " | " & ex.ToString(), ""))
    End Sub

    Private Sub WriteLog(level As String, msg As String)
        Try
            Dim dir = Path.GetDirectoryName(LogFile)
            If Not Directory.Exists(dir) Then Directory.CreateDirectory(dir)
            File.AppendAllText(LogFile, $"{Now:yyyy-MM-dd HH:mm:ss} [{level}] {msg}{vbCrLf}")
        Catch
            ' 忽略日志写入错误
        End Try
    End Sub
End Module