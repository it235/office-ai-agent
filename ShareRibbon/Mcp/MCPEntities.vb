Imports System.Collections.Generic
Imports System.Text
Imports Newtonsoft.Json.Linq

' MCP 初始化响应
Public Class MCPInitResponse
    Public Property Success As Boolean
    Public Property ErrorMessage As String
    Public Property ProtocolVersion As String
    Public Property ServerInfo As MCPServerInfo
    Public Property Capabilities As MCPServerCapabilities
End Class

' MCP 服务器信息
Public Class MCPServerInfo
    Public Property Name As String
    Public Property Version As String
End Class

' MCP 服务器能力
Public Class MCPServerCapabilities
    Public Property Tools As Boolean
    Public Property Resources As Boolean
    Public Property Prompts As Boolean
    Public Property Sampling As Boolean
    Public Property Roots As Boolean
End Class

' MCP 工具信息
Public Class MCPToolInfo
    Public Property Name As String
    Public Property Description As String
    Public Property InputSchema As Object
End Class

' MCP 工具调用结果
Public Class MCPToolResult
    Public Property IsError As Boolean
    Public Property ErrorMessage As String
    Public Property Content As List(Of MCPContent)
End Class

' MCP 内容项
Public Class MCPContent
    Public Property Type As String
    Public Property Text As String
    Public Property Data As String
    Public Property MimeType As String
    Public Property Uri As String
    Public Property Blob As String
End Class

' MCP 资源信息
Public Class MCPResourceInfo
    Public Property Uri As String
    Public Property Name As String
    Public Property Description As String
    Public Property MimeType As String
End Class

' MCP 资源读取结果
Public Class MCPResourceResult
    Public Property Contents As List(Of MCPContent)
End Class

' MCP 提示信息
Public Class MCPPromptInfo
    Public Property Name As String
    Public Property Description As String
    Public Property Arguments As List(Of MCPPromptArgument)
End Class

' MCP 提示参数
Public Class MCPPromptArgument
    Public Property Name As String
    Public Property Description As String
    Public Property Required As Boolean
End Class

' Stdio 选项
' Stdio 选项
Public Class StdioOptions
    Public Property Command As String
    Public Property Arguments As String
    Public Property WorkingDirectory As String
    Public Property EnvironmentVariables As Dictionary(Of String, String)

    Public Sub New()
        EnvironmentVariables = New Dictionary(Of String, String)()
    End Sub

    ' 从 stdio:// URL 解析 - 改进版本
    Public Shared Function Parse(stdioUrl As String) As StdioOptions
        ' 基本格式: stdio://command?args=arguments&workdir=dir&env.VAR1=value1&env.VAR2=value2
        Dim options = New StdioOptions()

        If stdioUrl.StartsWith("stdio://") Then
            Dim parts = stdioUrl.Substring("stdio://".Length).Split("?"c)
            options.Command = parts(0)

            If parts.Length > 1 Then
                Dim queryParams = parts(1).Split("&"c)
                For Each param In queryParams
                    If param.Contains("=") Then
                        ' 修正这里的分割方法，使用 VB.NET 兼容的语法
                        Dim kvp = param.Split(New Char() {"="c}, 2, StringSplitOptions.None)
                        If kvp.Length = 2 Then
                            Dim key = kvp(0)
                            Dim value = Uri.UnescapeDataString(kvp(1))

                            If key = "args" Then
                                options.Arguments = value
                            ElseIf key = "workdir" Then
                                options.WorkingDirectory = value
                            ElseIf key.StartsWith("env.") Then
                                Dim envName = key.Substring("env.".Length)
                                options.EnvironmentVariables(envName) = value
                            End If
                        End If
                    End If
                Next
            End If
        End If

        Return options
    End Function

    ' 转换回 stdio:// URL - 改进版本
    Public Function ToUrl() As String
        Dim sb = New StringBuilder("stdio://")
        sb.Append(Command)

        Dim hasQuery = False

        If Not String.IsNullOrEmpty(Arguments) Then
            sb.Append("?args=").Append(Uri.EscapeDataString(Arguments))
            hasQuery = True
        End If

        If Not String.IsNullOrEmpty(WorkingDirectory) Then
            sb.Append(If(hasQuery, "&", "?")).Append("workdir=").Append(Uri.EscapeDataString(WorkingDirectory))
            hasQuery = True
        End If

        ' 环境变量使用单独的参数
        For Each kvp In EnvironmentVariables
            sb.Append(If(hasQuery, "&", "?")).Append("env.").Append(Uri.EscapeDataString(kvp.Key)).Append("=").Append(Uri.EscapeDataString(kvp.Value))
            hasQuery = True
        Next

        Return sb.ToString()
    End Function
End Class