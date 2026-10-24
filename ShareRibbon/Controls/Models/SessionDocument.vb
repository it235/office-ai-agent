' ShareRibbon\Controls\Models\SessionDocument.vb
' 会话所属文档的标识。共享层不接触任何宿主 COM 类型，宿主只把文档完整路径交给这里。

Imports System.IO
Imports System.Security.Cryptography
Imports System.Text

''' <summary>
''' 一段聊天历史绑定的文档。Id 为空表示未绑定（未保存文档或读取失败），界面据此显示"通用会话"。
''' </summary>
Public Class SessionDocument

    ''' <summary>稳定标识：文档绝对路径的小写哈希，只用于分组，不含路径本身。</summary>
    Public Property Id As String = ""

    ''' <summary>展示用文件名（不含目录），避免把用户本地目录结构显示在历史列表里。</summary>
    Public Property Name As String = ""

    Public ReadOnly Property IsBound As Boolean
        Get
            Return Not String.IsNullOrWhiteSpace(Id)
        End Get
    End Property

    ''' <summary>
    ''' 由宿主传入文档完整路径。路径只参与哈希、不入库，因此改名前后的同一文档仍按路径归组，
    ''' 而未保存文档（无目录、"文档1"这类可重复名字）不绑定，免得两篇新文档的会话并成一条。
    ''' </summary>
    Public Shared Function FromFullPath(fullPath As String) As SessionDocument
        Dim doc As New SessionDocument()
        If String.IsNullOrWhiteSpace(fullPath) Then Return doc
        Try
            Dim normalized = Path.GetFullPath(fullPath.Trim())
            Dim directory = Path.GetDirectoryName(normalized)
            If String.IsNullOrEmpty(directory) Then Return doc

            doc.Name = Path.GetFileName(normalized)
            Using sha = SHA256.Create()
                Dim hash = sha.ComputeHash(Encoding.UTF8.GetBytes(normalized.ToLowerInvariant()))
                doc.Id = BitConverter.ToString(hash, 0, 8).Replace("-", "").ToLowerInvariant()
            End Using
        Catch ex As Exception
            AppLogger.Debug("SessionDocument", "文档标识解析失败，按通用会话处理", ex)
            Return New SessionDocument()
        End Try
        Return doc
    End Function

    ''' <summary>用于状态提示与空态文案的文档名，未绑定时返回空串。</summary>
    Public Shared Function DisplayName(doc As SessionDocument) As String
        If doc Is Nothing OrElse Not doc.IsBound Then Return ""
        Return If(String.IsNullOrWhiteSpace(doc.Name), "未命名文档", doc.Name)
    End Function
End Class
