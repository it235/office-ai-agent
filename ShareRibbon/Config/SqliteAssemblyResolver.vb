' ShareRibbon\Config\SqliteAssemblyResolver.vb
' 部署时仅 WordAi 目录含 System.Data.SQLite.dll，ExcelAi/PowerPointAi 需从此加载

Imports System.Collections.Generic
Imports System.IO
Imports System.Reflection
Imports System.Linq

    ''' <summary>
    ''' AssemblyResolve：从 WordAi、根目录等目录加载 System.Data.SQLite 或 Markdig
    ''' </summary>
    Public Class SqliteAssemblyResolver

    Private Shared _registered As Boolean = False
    Private Shared ReadOnly _lockObj As New Object()

    Public Shared Sub EnsureRegistered()
        If _registered Then Return
        SyncLock _lockObj
            If _registered Then Return
            AddHandler AppDomain.CurrentDomain.AssemblyResolve, AddressOf OnAssemblyResolve
            _registered = True
            TryPreloadSqlite()
            TryPreloadMarkdig()
        End SyncLock
    End Sub

    Private Shared Function GetProbeDirs() As IEnumerable(Of String)
        Dim our = GetType(SqliteAssemblyResolver).Assembly
        Dim locDir = GetDir(our.Location)
        Dim cbDir = GetDirFromCodeBase(our)
        
        Dim list As New List(Of String) From {locDir, cbDir}
        
        ' 尝试从 locDir 和 cbDir 的父目录查找（安装根目录）
        Dim p1 = GetParent(locDir)
        If Not String.IsNullOrEmpty(p1) Then list.Add(p1)
        
        Dim p2 = GetParent(cbDir)
        If Not String.IsNullOrEmpty(p2) Then list.Add(p2)

        ' 显式添加常见的子目录路径
        Dim initialPaths = list.ToArray()
        For Each p In initialPaths
            If String.IsNullOrEmpty(p) Then Continue For
            list.Add(Path.Combine(p, "WordAi"))
            list.Add(Path.Combine(p, "ExcelAi"))
            list.Add(Path.Combine(p, "PowerPointAi"))
        Next

        ' 兜底：尝试从 AppDomain 基目录及其上级查找
        Dim baseDir = AppDomain.CurrentDomain.BaseDirectory
        list.Add(baseDir)
        Dim baseParent = GetParent(baseDir)
        If Not String.IsNullOrEmpty(baseParent) Then list.Add(baseParent)

        ' 去重并排除不存在的目录
        Dim result As New List(Of String)
        For Each d In list
            If Not String.IsNullOrEmpty(d) AndAlso Directory.Exists(d) Then
                If Not result.Contains(d) Then result.Add(d)
            End If
        Next
        Return result
    End Function

    Private Shared Function GetDir(filePath As String) As String
        If String.IsNullOrEmpty(filePath) Then Return Nothing
        Try
            Return Path.GetDirectoryName(filePath)
        Catch
            Return Nothing
        End Try
    End Function

    Private Shared Function GetParent(dir As String) As String
        If String.IsNullOrEmpty(dir) OrElse Not Directory.Exists(dir) Then Return Nothing
        Try
            Return Path.GetDirectoryName(dir.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar))
        Catch
            Return Nothing
        End Try
    End Function

    Private Shared Function GetDirFromCodeBase(asm As Assembly) As String
        Try
            Dim cb = asm.CodeBase
            If String.IsNullOrEmpty(cb) Then Return Nothing
            Dim uri As New Uri(cb)
            Dim localPath = uri.LocalPath
            If String.IsNullOrEmpty(localPath) Then Return Nothing
            If uri.Host?.Length > 0 Then localPath = "\\" & uri.Host & localPath
            Return Path.GetDirectoryName(localPath)
        Catch
            Return Nothing
        End Try
    End Function

    Private Shared Sub TryPreloadSqlite()
        For Each d As String In GetProbeDirs()
            Dim p = Path.Combine(d, "System.Data.SQLite.dll")
            If File.Exists(p) Then
                Try
                    Assembly.LoadFrom(p)
                Catch
                End Try
                Return
            End If
        Next
    End Sub

    Private Shared Sub TryPreloadMarkdig()
        For Each d As String In GetProbeDirs()
            Dim p = Path.Combine(d, "Markdig.dll")
            If File.Exists(p) Then
                Try
                    Assembly.LoadFrom(p)
                Catch
                End Try
                Return
            End If
        Next
    End Sub

    Private Shared Function OnAssemblyResolve(sender As Object, args As ResolveEventArgs) As Assembly
        Try
            Dim name As New AssemblyName(args.Name)
            Dim simpleName = name.Name
            If simpleName <> "System.Data.SQLite" AndAlso simpleName <> "Markdig" Then Return Nothing

            For Each d As String In GetProbeDirs()
                Dim p = Path.Combine(d, simpleName & ".dll")
                If File.Exists(p) Then
                    Try
                        Return Assembly.LoadFrom(p)
                    Catch
                    End Try
                End If
            Next
        Catch
        End Try
        Return Nothing
    End Function
End Class
