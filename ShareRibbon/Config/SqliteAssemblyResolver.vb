' ShareRibbon\Config\SqliteAssemblyResolver.vb
' 部署时仅 WordAi 目录含 System.Data.SQLite.dll，ExcelAi/PowerPointAi 需从此加载

Imports System.Collections.Generic
Imports System.IO
Imports System.Reflection

''' <summary>
''' AssemblyResolve：从 WordAi 等目录加载 System.Data.SQLite
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
        End SyncLock
    End Sub

    Private Shared Function GetProbeDirs() As IEnumerable(Of String)
        Dim our = GetType(SqliteAssemblyResolver).Assembly
        Dim locDir = GetDir(our.Location)
        Dim cbDir = GetDirFromCodeBase(our)
        Dim parent = GetParent(locDir)
        If String.IsNullOrEmpty(parent) Then parent = GetParent(cbDir)

        Dim list As New List(Of String) From {locDir, cbDir, parent}
        If Not String.IsNullOrEmpty(parent) Then
            list.Add(Path.Combine(parent, "WordAi"))
        End If
        list.Add(AppDomain.CurrentDomain.BaseDirectory)
        Return list
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
            If String.IsNullOrEmpty(d) OrElse Not Directory.Exists(d) Then Continue For
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

    Private Shared Function OnAssemblyResolve(sender As Object, args As ResolveEventArgs) As Assembly
        Dim name As New AssemblyName(args.Name)
        If name.Name <> "System.Data.SQLite" Then Return Nothing

        For Each d As String In GetProbeDirs()
            If String.IsNullOrEmpty(d) OrElse Not Directory.Exists(d) Then Continue For
            Dim p = Path.Combine(d, "System.Data.SQLite.dll")
            If File.Exists(p) Then
                Try
                    Return Assembly.LoadFrom(p)
                Catch
                End Try
                Return Nothing
            End If
        Next
        Return Nothing
    End Function
End Class
