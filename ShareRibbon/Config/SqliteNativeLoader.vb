' ShareRibbon\Config\SqliteNativeLoader.vb
' e_sqlite3.dll 加载器，参照 WebView2Loader 模式：确保原生库就绪后供 System.Data.SQLite 2.0 使用

Imports System.Collections.Generic
Imports System.IO
Imports System.Runtime.InteropServices

''' <summary>
''' 确保 e_sqlite3.dll 已就绪并可加载。若输出目录缺失则尝试从 packages 拷贝。
''' 支持多路径探测，以兼容 WPS 等宿主下 BaseDirectory 不同的情况。
''' </summary>
Public Class SqliteNativeLoader

    Private Shared _loaded As Boolean = False
    Private Shared ReadOnly _lockObj As New Object()

    ''' <summary>
    ''' 在首次使用 SQLite 前调用，确保原生库可用
    ''' </summary>
    Public Shared Sub EnsureLoaded()
        If _loaded Then Return

        SyncLock _lockObj
            If _loaded Then Return

            Dim arch As String = GetRuntimeArchitecture()
            Dim dllPath As String = Nothing

            ' 候选基目录：BaseDirectory、本程序集目录（WPS 等可能不同）
            Dim candidates As New List(Of String) From {AppDomain.CurrentDomain.BaseDirectory}
            Try
                Dim asmDir = Path.GetDirectoryName(GetType(SqliteNativeLoader).Assembly.Location)
                If Not String.IsNullOrEmpty(asmDir) AndAlso Not candidates.Contains(asmDir) Then
                    candidates.Add(asmDir)
                End If
            Catch
            End Try

            For Each baseDir In candidates
                If String.IsNullOrEmpty(baseDir) Then Continue For
                dllPath = Path.Combine(baseDir, "runtimes", arch, "native", "e_sqlite3.dll")
                If File.Exists(dllPath) Then Exit For
                ' 若不存在则尝试从 packages 拷贝（仅对首个候选，通常是调试目录）
                TryCopyFromPackages(baseDir, arch)
                If File.Exists(dllPath) Then Exit For
                dllPath = Nothing
            Next

            If String.IsNullOrEmpty(dllPath) OrElse Not File.Exists(dllPath) Then
                Throw New FileNotFoundException(
                    $"e_sqlite3.dll 未找到。请将 e_sqlite3.dll 放入 runtimes\{arch}\native\ 目录。可从 NuGet SourceGear.sqlite3 包获取。",
                    "runtimes\" & arch & "\native\e_sqlite3.dll")
            End If

            Dim handle As IntPtr = NativeMethods.LoadLibrary(dllPath)
            If handle = IntPtr.Zero Then
                Dim err As Integer = Marshal.GetLastWin32Error()
                Throw New Exception($"加载 e_sqlite3.dll 失败 (路径: {dllPath})，错误码: {err}")
            End If

            _loaded = True
        End SyncLock
    End Sub

    Private Shared Function GetRuntimeArchitecture() As String
        Select Case RuntimeInformation.ProcessArchitecture
            Case Architecture.X86
                Return "win-x86"
            Case Architecture.X64
                Return "win-x64"
            Case Architecture.Arm64
                Return "win-arm64"
            Case Else
                Return "win-x64"
        End Select
    End Function

    ''' <summary>
    ''' 若 packages 中存在 SourceGear.sqlite3，则拷贝到 runtimes\arch\native\（与 WebView2 一致）
    ''' </summary>
    Private Shared Sub TryCopyFromPackages(baseDir As String, arch As String)
        ' 从 bin\Debug 向上找到解决方案根目录（含 packages 的目录）
        Dim current As String = Path.GetFullPath(baseDir)
        Dim packagesDir As String = Nothing
        For i As Integer = 1 To 6
            Dim pkg As String = Path.Combine(current, "packages")
            If Directory.Exists(pkg) Then
                packagesDir = pkg
                Exit For
            End If
            Dim parent As String = Path.GetDirectoryName(current)
            If String.IsNullOrEmpty(parent) OrElse parent = current Then Exit Sub
            current = parent
        Next
        If String.IsNullOrEmpty(packagesDir) Then Return

        Dim sourcePath As String = Path.Combine(packagesDir, "SourceGear.sqlite3.3.50.4.5", "runtimes", arch, "native", "e_sqlite3.dll")
        If Not File.Exists(sourcePath) Then
            sourcePath = Path.Combine(packagesDir, "SourceGear.sqlite3.3.50.4.2", "runtimes", arch, "native", "e_sqlite3.dll")
        End If
        If Not File.Exists(sourcePath) Then Return

        Dim destDir As String = Path.Combine(baseDir, "runtimes", arch, "native")
        Dim destPath As String = Path.Combine(destDir, "e_sqlite3.dll")
        Try
            Directory.CreateDirectory(destDir)
            File.Copy(sourcePath, destPath, overwrite:=True)
            Debug.WriteLine($"SqliteNativeLoader: 已从 packages 拷贝 e_sqlite3.dll 到 {destPath}")
        Catch ex As Exception
            Debug.WriteLine($"SqliteNativeLoader: 从 packages 拷贝失败: {ex.Message}")
        End Try
    End Sub

    Private Class NativeMethods
        <DllImport("kernel32", CharSet:=CharSet.Unicode, SetLastError:=True)>
        Public Shared Function LoadLibrary(lpFileName As String) As IntPtr
        End Function
    End Class
End Class
