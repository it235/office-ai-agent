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
    Public Property Command As String = "node"
    Public Property Arguments As String = ""
    Public Property WorkingDirectory As String = ""
    Public Property EnvironmentVariables As New Dictionary(Of String, String)()

    ' 将选项转换为URL - 修复反斜杠处理
    Public Function ToUrl() As String
        Dim sb As New StringBuilder("stdio://")
        sb.Append(Command.Replace("\", "\\"))  ' 转义路径中的反斜杠

        Dim queryParams As New List(Of String)()

        ' 添加参数 - 确保转义路径中的反斜杠
        If Not String.IsNullOrEmpty(Arguments) Then
            queryParams.Add($"args={Uri.EscapeDataString(Arguments)}")
        End If

        ' 添加工作目录 - 确保转义路径中的反斜杠
        If Not String.IsNullOrEmpty(WorkingDirectory) Then
            queryParams.Add($"workdir={Uri.EscapeDataString(WorkingDirectory.Replace("\", "\\"))}")
        End If

        ' 添加环境变量
        For Each kvp In EnvironmentVariables
            ' 确保值中的反斜杠也被正确转义
            Dim escapedValue = kvp.Value.Replace("\", "\\")
            queryParams.Add($"{Uri.EscapeDataString(kvp.Key)}={Uri.EscapeDataString(escapedValue)}")
        Next

        If queryParams.Count > 0 Then
            sb.Append("?").Append(String.Join("&", queryParams))
        End If

        Return sb.ToString()
    End Function

    ' 从URL解析选项 - 修复反斜杠处理
    Public Shared Function Parse(stdioUrl As String) As StdioOptions
        Dim options As New StdioOptions()

        If Not stdioUrl.StartsWith("stdio://") Then
            Return options
        End If

        ' 解析URL
        Dim urlWithoutScheme = stdioUrl.Substring(8)
        Dim commandParts = urlWithoutScheme.Split(New Char() {"?"c}, 2)

        ' 恢复命令中的反斜杠
        options.Command = Uri.UnescapeDataString(commandParts(0)).Replace("\\", "\")

        ' 解析查询参数
        If commandParts.Length > 1 Then
            Dim queryString = commandParts(1)
            Dim pairs = queryString.Split("&"c)

            For Each pair In pairs
                Dim keyValue = pair.Split(New Char() {"="c}, 2)
                If keyValue.Length = 2 Then
                    Dim key = Uri.UnescapeDataString(keyValue(0))
                    ' 恢复值中的反斜杠
                    Dim value = Uri.UnescapeDataString(keyValue(1)).Replace("\\", "\")

                    If key = "args" Then
                        options.Arguments = value
                    ElseIf key = "workdir" Then
                        options.WorkingDirectory = value
                    Else
                        ' 所有其他参数视为环境变量
                        options.EnvironmentVariables(key) = value
                    End If
                End If
            Next
        End If

        Return options
    End Function


    ' 转换成官方格式的args数组
    ' 将参数字符串转换为数组
    Public Function GetArgsArray() As String()
        If String.IsNullOrEmpty(Arguments) Then
            Return New String() {}
        End If

        ' 处理引号内的空格
        Dim result As New List(Of String)()
        Dim currentArg As New StringBuilder()
        Dim inQuotes As Boolean = False
        Dim escaping As Boolean = False

        For i As Integer = 0 To Arguments.Length - 1
            Dim c As Char = Arguments(i)

            ' 处理转义字符
            If escaping Then
                currentArg.Append(c)
                escaping = False
                Continue For
            End If

            ' 检查是否是转义字符
            If c = "\"c Then
                escaping = True
                Continue For
            End If

            ' 处理引号
            If c = """"c Then
                inQuotes = Not inQuotes
                ' 保留引号，因为某些命令行需要引号
                currentArg.Append(c)
                Continue For
            End If

            ' 处理空格
            If c = " "c AndAlso Not inQuotes Then
                If currentArg.Length > 0 Then
                    result.Add(currentArg.ToString())
                    currentArg.Clear()
                End If
                Continue For
            End If

            ' 添加普通字符
            currentArg.Append(c)
        Next

        ' 添加最后一个参数
        If currentArg.Length > 0 Then
            result.Add(currentArg.ToString())
        End If

        Return result.ToArray()
    End Function
End Class