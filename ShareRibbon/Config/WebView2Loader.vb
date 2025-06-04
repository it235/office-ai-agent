Imports System.IO
Imports System.Runtime.InteropServices

Public Class WebView2Loader

    Public Shared Sub EnsureWebView2Loader()
        Try
            Dim baseDir As String = AppDomain.CurrentDomain.BaseDirectory
            Dim architecture As String = GetRuntimeArchitecture()
            Dim loaderPath As String = Path.Combine(
                baseDir,
                "runtimes",
                architecture,
                "native",
                "WebView2Loader.dll"
            )

            If Not File.Exists(loaderPath) Then
                Throw New FileNotFoundException($"WebView2Loader.dll ·����Ч: {loaderPath}")
            End If

            Dim handle As IntPtr = NativeMethods.LoadLibrary(loaderPath)
            If handle = IntPtr.Zero Then
                Dim errorCode As Integer = Marshal.GetLastWin32Error()
                Throw New Exception($"����ʧ�ܣ��������: {errorCode}")
            End If

        Catch ex As Exception
            ' �׳��쳣���¼��־
            Throw New Exception("WebView2Loader ��ʼ��ʧ��", ex)
        End Try
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
                Throw New PlatformNotSupportedException("��֧�ֵĴ������ܹ�")
        End Select
    End Function

    Private Class NativeMethods
        <DllImport("kernel32", CharSet:=CharSet.Unicode, SetLastError:=True)>
        Public Shared Function LoadLibrary(lpFileName As String) As IntPtr
        End Function
    End Class

End Class
